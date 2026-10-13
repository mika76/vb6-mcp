using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

public static class EditToolsTests
{
    const char US = Wire.US;
    const string Module = "Option Explicit\r\n\r\nPrivate Sub A()\r\nEnd Sub";

    /// <summary>A fake IDE holding one module that applies REPLACE like the add-in (with the old-text guard).</summary>
    sealed class FakeIde : IDisposable
    {
        public string Text = Module;
        public List<string[]> Replaces = [];
        public long ClickedHwnd;
        public IdeWindowState Window;
        public FakeAddin Addin;

        public FakeIde(IdeMode mode = IdeMode.Design, params DialogInfo[] dialogs)
        {
            Window = new(1, true, $"Basic - Microsoft Visual Basic [{mode.ToString().ToLowerInvariant()}]", mode, true, false, dialogs);
            Addin = new FakeAddin((verb, a) => verb switch
            {
                "PROJECTS" => ("OK", $"1{US}Basic{US}0{US}D:\\b\\Basic.vbp{US}1"),
                "READ" => ("OK", $"{Text.Split("\r\n").Length}{US}1{US}{Text}"),
                "PROC" => ("OK", $"3{US}3{US}2{US}{VbCode.Slice(Text, 3, 2)}"),
                "REPLACE" => Replace(a),
                "CLICK" => Click(a[0]),
                _ => ("ERR", $"400{US}unknown {verb}"),
            });
        }

        (string, string) Replace(string[] a)
        {
            Replaces.Add(a);
            int start = int.Parse(a[2]), count = int.Parse(a[3]);
            if (VbCode.Slice(Text, start, count) != a[4] && count > 0) return ("ERR", $"409{US}the target lines changed since they were read");
            var lines = Text.Split("\r\n").ToList();
            lines.RemoveRange(start - 1, count);
            if (a[5].Length > 0) lines.InsertRange(start - 1, a[5].Split("\r\n"));
            Text = string.Join("\r\n", lines);
            return ("OK", $"{lines.Count}{US}1");
        }

        (string, string) Click(string hex) { ClickedHwnd = Convert.ToInt64(hex, 16); Window = Window with { Dialogs = [] }; return ("OK", "clicked"); }

        public IReadOnlyList<Tool> Tools() => new IdeTools(new Env(this)).All();
        public void Dispose() => Addin.Dispose();

        sealed class Env(FakeIde ide) : IIdeEnvironment
        {
            public DiscoveryResult Scan() => new([ide.Addin.Instance], [], []);
            public int DeleteStale(DiscoveryResult result) => 0;
            public IdeWindowState ReadWindows(int pid) => ide.Window;
            public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
        }
    }

    static async Task<JsonObject> Call(IReadOnlyList<Tool> tools, string name, JsonObject args)
    {
        var res = (await tools.First(t => t.Name == name).InvokeAsync(args, CancellationToken.None)).ToJson();
        var text = (string)res["content"]![0]!["text"]!;
        return new JsonObject { ["isError"] = (bool)res["isError"]!, ["text"] = text, ["json"] = (bool)res["isError"]! ? null : JsonNode.Parse(text) };
    }

    [Test]
    public static async Task ReplaceWithCorrectHashEditsAndReturnsNewHash()
    {
        using var ide = new FakeIde();
        var r = await Call(ide.Tools(), "code_replace_lines", new JsonObject
        {
            ["component"] = "modA", ["start_line"] = 3, ["count"] = 2,
            ["new_text"] = "Private Sub A()\n    Beep\nEnd Sub\n", ["expected_hash"] = VbCode.Hash(Module),
        });
        Assert.Equal(false, (bool)r["isError"]!, (string?)r["text"]);
        Assert.Equal("Private Sub A()\r\nEnd Sub", ide.Replaces[0][4], "expected old lines sent as the guard");
        Assert.Equal("Private Sub A()\r\n    Beep\r\nEnd Sub", ide.Replaces[0][5], "normalized new text, no trailing break");
        Assert.Equal(VbCode.Hash(ide.Text), (string)r["json"]!["hash"]!);
    }

    [Test]
    public static async Task StaleHashIsRefusedWithoutSendingAnything()
    {
        using var ide = new FakeIde();
        var r = await Call(ide.Tools(), "code_insert", new JsonObject
        { ["component"] = "modA", ["line"] = 2, ["text"] = "' hi", ["expected_hash"] = "0000000000000000" });
        Assert.Equal(true, (bool)r["isError"]!);
        Assert.True(((string)r["text"]!).Contains("conflict"), "explains the conflict");
        Assert.Equal(0, ide.Replaces.Count, "nothing sent");
    }

    [Test]
    public static async Task EditsOutsideDesignModeAreRefused()
    {
        using var ide = new FakeIde(IdeMode.Break);
        var r = await Call(ide.Tools(), "code_insert", new JsonObject
        { ["component"] = "modA", ["line"] = 2, ["text"] = "' hi", ["expected_hash"] = VbCode.Hash(Module) });
        Assert.Equal(true, (bool)r["isError"]!);
        Assert.True(((string)r["text"]!).Contains("break mode"), "names the mode");
        Assert.Equal(0, ide.Replaces.Count, "nothing sent");
    }

    [Test]
    public static async Task ProcReplaceUsesTheProcRange()
    {
        using var ide = new FakeIde();
        var r = await Call(ide.Tools(), "proc_replace", new JsonObject
        { ["component"] = "modA", ["name"] = "A", ["new_text"] = "Private Sub A()\r\n    Beep\r\nEnd Sub", ["expected_hash"] = VbCode.Hash(Module) });
        Assert.Equal(false, (bool)r["isError"]!, (string?)r["text"]);
        Assert.SequenceEqual(new[] { "3", "2" }, ide.Replaces[0][2..4], "start and count from PROC");
        Assert.True(ide.Text.EndsWith("    Beep\r\nEnd Sub"), "procedure replaced");
    }

    [Test]
    public static async Task InvalidTextIsRefusedBeforeSending()
    {
        using var ide = new FakeIde();
        var r = await Call(ide.Tools(), "code_insert", new JsonObject
        { ["component"] = "modA", ["line"] = 2, ["text"] = new string('x', 2000), ["expected_hash"] = VbCode.Hash(Module) });
        Assert.Equal(true, (bool)r["isError"]!);
        Assert.True(((string)r["text"]!).Contains("1023"), "names the limit");
        Assert.Equal(0, ide.Replaces.Count, "nothing sent");
    }

    [Test]
    public static async Task DialogRespondClicksTheNamedButtonAndVerifies()
    {
        var dlg = new DialogInfo(0x1000, "Source Code Control", ["Add this project to SourceSafe?"],
            [new DialogButton("Yes", 0x1001), new DialogButton("No", 0x1002), new DialogButton("Help", 0x1003)]);
        using var ide = new FakeIde(IdeMode.Design, dlg);
        var bad = await Call(ide.Tools(), "ide_dialog_respond", new JsonObject { ["button"] = "Maybe" });
        Assert.True(((string)bad["text"]!).Contains("Buttons: Yes, No, Help"), "lists the buttons");
        var r = await Call(ide.Tools(), "ide_dialog_respond", new JsonObject { ["button"] = "no" });
        Assert.Equal(false, (bool)r["isError"]!, (string?)r["text"]);
        Assert.Equal(0x1002L, ide.ClickedHwnd);
        Assert.Equal(true, (bool)r["json"]!["dialogClosed"]!);
    }
}
