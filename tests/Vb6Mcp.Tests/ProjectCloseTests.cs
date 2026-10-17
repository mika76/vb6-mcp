using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>
/// project_close with discard (spike 31): VBProjects.Remove saves changed components silently, so the server snapshots
/// the files the close writes and restores them afterwards. The fake add-in's CLOSE does what VB did live.
/// </summary>
public static class ProjectCloseTests
{
    const char US = Wire.US;

    sealed class Env(FakeAddin addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "Basic - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    [Test]
    public static async Task DiscardRestoresWhatTheCloseSaved()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-close-").FullName;
        try
        {
            var vbp = Path.Combine(dir, "Basic.vbp"); var bas = Path.Combine(dir, "modA.bas");
            File.WriteAllBytes(vbp, Vb1252.Encoding.GetBytes("Type=Exe\r\nModule=modA; modA.bas\r\n"));
            File.WriteAllBytes(bas, Vb1252.Encoding.GetBytes("Attribute VB_Name = \"modA\"\r\nOption Explicit\r\n"));
            var originalBas = File.ReadAllBytes(bas); var originalVbp = File.ReadAllBytes(vbp);
            bool open = true;
            using var fake = new FakeAddin((verb, a) => verb switch
            {
                "PROJECTS" => ("OK", open ? $"1{US}Basic{US}0{US}{vbp}{US}1" : ""),
                "COMPONENTS" => ("OK", $"modA{US}1{US}1{US}2{US}{bas}"),   // modA has unsaved changes
                "CLOSE" => Close(),
                _ => ("ERR", $"400{US}unexpected {verb}"),
            });
            (string, string) Close()
            {
                // what VB6 does on VBProjects.Remove: saves the changed module and the .vbp, writes never-saved ones
                File.WriteAllBytes(bas, Vb1252.Encoding.GetBytes("Attribute VB_Name = \"modA\"\r\nOption Explicit\r\n' unwanted edit\r\n"));
                File.WriteAllBytes(vbp, Vb1252.Encoding.GetBytes("Type=Exe\r\nModule=modA; modA.bas\r\nClass=Class1; Class1.cls\r\n"));
                File.WriteAllText(Path.Combine(dir, "Class1.cls"), "VERSION 1.0 CLASS");
                open = false;
                return ("QUEUED", "close");
            }
            var r = (await new IdeTools(new Env(fake)).All().First(t => t.Name == "project_close").InvokeAsync(
                new JsonObject { ["project"] = "Basic", ["discard"] = true }, CancellationToken.None)).ToJson();
            var text = (string)r["content"]![0]!["text"]!;
            Assert.True(!(bool)r["isError"]! && text.Contains("\"status\":\"closed\""), text);
            Assert.True(File.ReadAllBytes(bas).AsSpan().SequenceEqual(originalBas), "modA.bas restored byte-exact");
            Assert.True(File.ReadAllBytes(vbp).AsSpan().SequenceEqual(originalVbp), ".vbp restored byte-exact");
            Assert.True(!File.Exists(Path.Combine(dir, "Class1.cls")), "the file VB created for a never-saved component is gone");
        }
        finally { Directory.Delete(dir, true); }
    }
}
