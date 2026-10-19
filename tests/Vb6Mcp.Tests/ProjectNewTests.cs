using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>
/// project_new and project_save path (spike 43). The fake add-in's PRJNEW / PRJSAVEAS write what VB wrote live: the
/// component files first, then the .vbp. The checks that refuse before anything is sent are tested too.
/// </summary>
public static class ProjectNewTests
{
    const char US = Wire.US;

    sealed class Env(FakeAddin addin, IdeMode mode = IdeMode.Design) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, mode == IdeMode.Design ? "Microsoft Visual Basic [design]" : "Microsoft Visual Basic", mode, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    /// <summary>A fake IDE holding one project (or none) that PRJNEW / PRJSAVEAS save the way VB did.</summary>
    sealed class FakeIde
    {
        public string? PrjName, Vbp, Form;
        public bool ListForm = true;
        public List<string> Verbs = [];
        public string[]? LastArgs;

        public (string, string)? Handle(string verb, string[] a)
        {
            Verbs.Add(verb);
            switch (verb)
            {
                case "PROJECTS": return ("OK", PrjName is null ? "" : $"1{US}{PrjName}{US}0{US}{Vbp ?? ""}{US}0");
                case "COMPONENTS": return ("OK", $"Form1{US}5{US}0{US}2{US}{Form ?? ""}");
                case "PRJNEW": LastArgs = a; PrjName = a[1]; Save(a[2]); return ("QUEUED", "new");
                case "PRJSAVEAS": LastArgs = a; Save(a[1]); return ("QUEUED", "save");
                default: return ("ERR", $"400{US}unexpected {verb}");
            }
        }

        void Save(string vbp)
        {
            Form = Path.Combine(Path.GetDirectoryName(vbp)!, "Form1.frm");
            File.WriteAllBytes(Form, Vb1252.Encoding.GetBytes("VERSION 5.00\r\nBegin VB.Form Form1\r\nEnd\r\nAttribute VB_Name = \"Form1\"\r\n"));
            File.WriteAllBytes(vbp, Vb1252.Encoding.GetBytes("Type=Exe\r\n" + (ListForm ? "Form=Form1.frm\r\n" : "") + $"Startup=\"Form1\"\r\nName=\"{PrjName}\"\r\n"));
            Vbp = vbp;
        }
    }

    static async Task<(bool IsError, string Text)> Call(IdeTools tools, string tool, JsonObject args)
    {
        var r = (await tools.All().First(t => t.Name == tool).InvokeAsync(args, CancellationToken.None)).ToJson();
        return ((bool)r["isError"]!, (string)r["content"]![0]!["text"]!);
    }

    static IdeTools Tools(FakeAddin fake, string root, IdeMode mode = IdeMode.Design) => new(new Env(fake, mode)) { Roots = [root], OpenWait = TimeSpan.FromSeconds(3) };

    [Test]
    public static async Task CreatesTheProjectInANewFolder()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-new-").FullName;
        try
        {
            var ide = new FakeIde();
            using var fake = new FakeAddin(ide.Handle);
            var vbp = Path.Combine(dir, "App", "MyApp.vbp");
            // an IDE with no project has no [design] tag in its title (spike 43): that must not refuse
            var (err, text) = await Call(Tools(fake, dir, IdeMode.Unknown), "project_new", new JsonObject { ["kind"] = "exe", ["path"] = vbp });
            Assert.True(!err && text.Contains("\"status\":\"created\""), text);
            Assert.Equal("exe|MyApp|" + vbp + "|add", string.Join("|", ide.LastArgs!));   // name from the file name; nothing open: add
            Assert.True(File.Exists(vbp) && File.Exists(Path.Combine(dir, "App", "Form1.frm")), "the .vbp and Form1.frm are written");
            Assert.True(text.Contains("Form1.frm") && text.Contains("\"notInVbp\":null"), text);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Test]
    public static async Task AVbpWithoutAComponentIsReportedIncomplete()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-new-").FullName;
        try
        {
            var ide = new FakeIde { ListForm = false };   // what VBProject.SaveAs alone wrote (spike 43)
            using var fake = new FakeAddin(ide.Handle);
            var (err, text) = await Call(Tools(fake, dir), "project_new", new JsonObject { ["kind"] = "exe", ["path"] = Path.Combine(dir, "A.vbp"), ["name"] = "Alpha" });
            Assert.True(!err && text.Contains("\"status\":\"incomplete\"") && text.Contains("\"notInVbp\":[\"Form1\"]"), text);
            Assert.Equal("Alpha", ide.LastArgs![1]);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Test]
    public static async Task RefusesBeforeSendingAnything()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-new-").FullName;
        try
        {
            var ide = new FakeIde();
            using var fake = new FakeAddin(ide.Handle);
            var tools = Tools(fake, dir);
            File.WriteAllText(Path.Combine(dir, "Old.vbp"), "Type=Exe\r\n");
            File.WriteAllText(Path.Combine(dir, "Class1.cls"), "x");
            async Task Refused(JsonObject args, string expected)
            {
                var (err, text) = await Call(tools, "project_new", args);
                Assert.True(err && text.Contains(expected), $"expected '{expected}': {text}");
            }
            await Refused(new() { ["kind"] = "exe", ["path"] = Path.Combine(dir, "Old.vbp") }, "already exists");
            await Refused(new() { ["kind"] = "dll", ["path"] = Path.Combine(dir, "Lib.vbp") }, "Class1.cls already exists");   // Class1 would be saved over it
            await Refused(new() { ["kind"] = "exe", ["path"] = @"C:\Elsewhere\X.vbp" }, "outside the workspace roots");
            await Refused(new() { ["kind"] = "exe", ["path"] = Path.Combine(dir, "My App.vbp") }, "isn't a valid project name");
            await Refused(new() { ["kind"] = "std", ["path"] = Path.Combine(dir, "X.vbp") }, "kind must be");
            await Refused(new() { ["kind"] = "exe", ["path"] = Path.Combine(dir, "X.txt") }, "must be a .vbp");
            Assert.True(!ide.Verbs.Contains("PRJNEW"), "nothing was queued: " + string.Join(",", ide.Verbs));

            ide.PrjName = "Basic"; ide.Vbp = Path.Combine(dir, "Old.vbp");   // a project is open: mode is required
            await Refused(new() { ["kind"] = "exe", ["path"] = Path.Combine(dir, "N", "N.vbp") }, "pass mode 'add'");
            await Refused(new() { ["kind"] = "exe", ["path"] = Path.Combine(dir, "N", "Basic.vbp"), ["mode"] = "add" }, "already open");
            Assert.True(!ide.Verbs.Contains("PRJNEW") && !Directory.Exists(Path.Combine(dir, "N")), "nothing queued, no folder made");
        }
        finally { Directory.Delete(dir, true); }
    }

    [Test]
    public static async Task AnAddinRefusalRemovesTheFolderItMade()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-new-").FullName;
        try
        {
            using var fake = new FakeAddin((verb, a) => verb switch
            {
                "PROJECTS" => ("OK", ""),
                "PRJNEW" => ("ERR", $"400{US}x isn't under the workspace roots"),
                _ => ("ERR", $"400{US}unexpected {verb}"),
            });
            var (err, text) = await Call(Tools(fake, dir), "project_new", new JsonObject { ["kind"] = "ocx", ["path"] = Path.Combine(dir, "Ctl", "Ctl.vbp") });
            Assert.True(err && text.Contains("workspace roots"), text);
            Assert.True(!Directory.Exists(Path.Combine(dir, "Ctl")), "the empty folder is gone");
        }
        finally { Directory.Delete(dir, true); }
    }

    [Test]
    public static async Task ProjectSaveWithPathSavesANeverSavedProject()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-new-").FullName;
        try
        {
            var ide = new FakeIde { PrjName = "Project1" };   // File > New Project: no .vbp, Form1 without a file
            using var fake = new FakeAddin(ide.Handle);
            var tools = Tools(fake, dir);
            var (err, text) = await Call(tools, "project_save", new JsonObject());
            Assert.True(err && text.Contains("never been saved; pass path"), text);   // not "nothing-to-save"
            var vbp = Path.Combine(dir, "P", "Project1.vbp");
            (err, text) = await Call(tools, "project_save", new JsonObject { ["path"] = vbp });
            Assert.True(!err && text.Contains("\"status\":\"saved\""), text);
            Assert.Equal("Project1|" + vbp, string.Join("|", ide.LastArgs!));
            (err, text) = await Call(tools, "project_save", new JsonObject { ["path"] = Path.Combine(dir, "Q.vbp") });
            Assert.True(err && text.Contains("already saved"), text);
        }
        finally { Directory.Delete(dir, true); }
    }
}
