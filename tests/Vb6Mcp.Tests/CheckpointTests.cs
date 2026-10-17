using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>project_checkpoint / project_rollback (PLAN.md E2): byte-exact copies of a project's files, and putting them back.</summary>
public static class CheckpointTests
{
    const char US = Wire.US;

    // a project with a byte undefined in 1252 (0x81) in the module, a form with its .frx
    static (string Dir, string Vbp, string Bas, string Frm, string Frx) Project()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-cp-").FullName;
        var vbp = Path.Combine(dir, "Basic.vbp"); var bas = Path.Combine(dir, "modA.bas"); var frm = Path.Combine(dir, "Form1.frm"); var frx = Path.Combine(dir, "Form1.frx");
        File.WriteAllBytes(vbp, Vb1252.Encoding.GetBytes("Type=Exe\r\nForm=Form1.frm\r\nModule=modA; modA.bas\r\n"));
        File.WriteAllBytes(bas, [.. Vb1252.Encoding.GetBytes("Attribute VB_Name = \"modA\"\r\n' "), 0x81, 0x0D, 0x0A]);
        File.WriteAllBytes(frm, Vb1252.Encoding.GetBytes("VERSION 5.00\r\nBegin VB.Form Form1\r\nEnd\r\n"));
        File.WriteAllBytes(frx, [1, 2, 3, 0, 255]);
        return (dir, vbp, bas, frm, frx);
    }

    [Test]
    public static void ACheckpointPutsEveryFileBackByteExact()
    {
        var (dir, vbp, bas, frm, frx) = Project();
        var store = new Checkpoints(Path.Combine(dir, "store"));
        try
        {
            var originals = new[] { vbp, bas, frm, frx }.ToDictionary(p => p, File.ReadAllBytes);
            var cp = store.Create(vbp, "before");
            Assert.Equal(4, cp.Files.Count(f => f.Copy is not null));
            Assert.Equal(0, Checkpoints.ChangedSince(cp).Count);

            File.WriteAllText(bas, "changed");
            File.WriteAllBytes(frx, [9]);
            File.WriteAllBytes(vbp, Vb1252.Encoding.GetBytes("Type=Exe\r\nForm=Form1.frm\r\nModule=modA; modA.bas\r\nClass=Class1; Class1.cls\r\n"));
            File.WriteAllText(Path.Combine(dir, "Class1.cls"), "VERSION 1.0 CLASS");
            Assert.Equal(3, Checkpoints.ChangedSince(cp).Count);

            var r = store.Restore(store.Find(vbp, "before")!);
            Assert.Equal(3, r.Restored.Count);
            Assert.Equal(1, r.Unchanged);
            foreach (var (p, bytes) in originals) Assert.True(File.ReadAllBytes(p).AsSpan().SequenceEqual(bytes), p + " byte-exact");
            Assert.True(r.AddedSince.Any(p => p.EndsWith("Class1.cls", StringComparison.OrdinalIgnoreCase)), "the class added since is listed");
            Assert.True(File.Exists(Path.Combine(dir, "Class1.cls")), "and left on disk");
        }
        finally { Directory.Delete(dir, true); }
    }

    [Test]
    public static void LatestAndLabelsFindCheckpointsAndOldOnesAreDropped()
    {
        var (dir, vbp, _, _, _) = Project();
        var store = new Checkpoints(Path.Combine(dir, "store"));
        try
        {
            for (int i = 0; i < Checkpoints.Keep + 2; i++) { store.Create(vbp, i == 0 ? "first" : null); Thread.Sleep(2); }
            Assert.Equal(Checkpoints.Keep, store.List(vbp).Count);
            Assert.True(store.Find(vbp, "first") is null, "the oldest went first");
            Assert.Equal(store.List(vbp)[0].Id, store.Find(vbp, "latest")!.Id);
        }
        finally { Directory.Delete(dir, true); }
    }

    // ---- the tools

    sealed class Env(FakeAddin? addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new(addin is null ? [] : [addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "Basic - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    static async Task<JsonObject> Call(IdeTools tools, string tool, JsonObject args)
    {
        var r = (await tools.All().First(t => t.Name == tool).InvokeAsync(args, CancellationToken.None)).ToJson();
        var text = (string)r["content"]![0]!["text"]!;
        return (bool)r["isError"]! ? new JsonObject { ["error"] = text } : JsonNode.Parse(text)!.AsObject();
    }

    [Test]
    public static async Task AClosedProjectIsCheckpointedAndRolledBackByPath()
    {
        var (dir, vbp, bas, _, _) = Project();
        try
        {
            var tools = new IdeTools(new Env(null)) { Roots = [dir], CheckpointFolder = Path.Combine(dir, "store") };
            var original = File.ReadAllBytes(bas);
            var made = await Call(tools, "project_checkpoint", new JsonObject { ["path"] = vbp, ["label"] = "safe" });
            Assert.True(made["error"] is null, made.ToJsonString());
            File.WriteAllText(bas, "oops");

            var noConfirm = await Call(tools, "project_rollback", new JsonObject { ["path"] = vbp, ["checkpoint"] = "safe" });
            Assert.True(((string?)noConfirm["error"] ?? "").Contains("confirm"), noConfirm.ToJsonString());
            var listed = await Call(tools, "project_checkpoint", new JsonObject { ["path"] = vbp, ["list"] = true });
            Assert.Equal(1, (int)listed["checkpoints"]![0]!["changedSince"]!);

            var back = await Call(tools, "project_rollback", new JsonObject { ["path"] = vbp, ["checkpoint"] = "safe", ["confirm"] = true });
            Assert.Equal("restored", (string)back["status"]!);
            Assert.True(File.ReadAllBytes(bas).AsSpan().SequenceEqual(original), "restored");
            var none = await Call(tools, "project_rollback", new JsonObject { ["path"] = vbp, ["checkpoint"] = "nope", ["confirm"] = true });
            Assert.True(((string?)none["error"] ?? "").Contains("safe"), "names the checkpoints there are: " + none.ToJsonString());
        }
        finally { Directory.Delete(dir, true); }
    }

    [Test]
    public static async Task AnOpenProjectIsClosedRestoredAndOpenedAgain()
    {
        var (dir, vbp, bas, _, _) = Project();
        try
        {
            bool open = true, dirty = false; var verbs = new List<string>();
            using var fake = new FakeAddin((verb, a) =>
            {
                verbs.Add(verb);
                switch (verb)
                {
                    case "PROJECTS": return ("OK", open ? $"1{US}Basic{US}0{US}{vbp}{US}0" : "");
                    case "COMPONENTS": return ("OK", $"modA{US}1{US}{(dirty ? 1 : 0)}{US}2{US}{bas}");
                    case "CLOSE":
                        if (dirty) File.WriteAllText(bas, "unsaved edit saved by the close");   // VBProjects.Remove saves (spike 31)
                        open = false; return ("QUEUED", "close");
                    case "OPEN": open = true; dirty = false; return ("QUEUED", "open");
                    default: return ("ERR", $"400{US}unexpected {verb}");
                }
            });
            var tools = new IdeTools(new Env(fake)) { Roots = [dir], CheckpointFolder = Path.Combine(dir, "store"), OpenWait = TimeSpan.FromSeconds(3) };
            var original = File.ReadAllBytes(bas);
            Assert.True((await Call(tools, "project_checkpoint", new JsonObject { ["project"] = "Basic" }))["error"] is null, "checkpoint of a saved open project");

            File.WriteAllText(bas, "saved later");   // the agent saved a change it now regrets...
            dirty = true;                            // ...and has more unsaved on top
            var refusedCp = await Call(tools, "project_checkpoint", new JsonObject { ["project"] = "Basic" });
            Assert.True(((string?)refusedCp["error"] ?? "").Contains("unsaved"), refusedCp.ToJsonString());
            var refused = await Call(tools, "project_rollback", new JsonObject { ["project"] = "Basic", ["checkpoint"] = "latest", ["confirm"] = true });
            Assert.True(((string?)refused["error"] ?? "").Contains("discard_unsaved"), refused.ToJsonString());
            Assert.True(!verbs.Contains("CLOSE"), "nothing closed when refused");

            var r = await Call(tools, "project_rollback", new JsonObject { ["project"] = "Basic", ["checkpoint"] = "latest", ["confirm"] = true, ["discard_unsaved"] = true });
            Assert.True(r["error"] is null, r.ToJsonString());
            Assert.True(File.ReadAllBytes(bas).AsSpan().SequenceEqual(original), "the checkpoint's bytes, not the saved or unsaved edit");
            Assert.True(verbs.IndexOf("CLOSE") < verbs.IndexOf("OPEN"), "closed, then opened again");
            Assert.Equal("Basic", (string)r["reopened"]!["project"]!);
        }
        finally { Directory.Delete(dir, true); }
    }
}
