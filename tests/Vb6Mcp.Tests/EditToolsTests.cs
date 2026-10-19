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
        public bool FailAfterDelete;   // like an InsertLines error after DeleteLines in the add-in
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
            if (FailAfterDelete) { Text = string.Join("\r\n", lines); return ("ERR", $"500{US}Invalid procedure call or argument"); }
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

    /// <summary>Several ranges in one call, numbered against the text read, applied bottom-up (improvement notes 5).</summary>
    [Test]
    public static async Task SeveralEditsInOneCallUseTheOriginalNumbering()
    {
        using var ide = new FakeIde();
        var r = await Call(ide.Tools(), "code_replace_lines", new JsonObject
        {
            ["component"] = "modA", ["expected_hash"] = VbCode.Hash(Module),
            ["edits"] = new JsonArray(
                new JsonObject { ["start_line"] = 1, ["count"] = 1, ["new_text"] = "Option Explicit\n' header" },   // grows the top
                new JsonObject { ["start_line"] = 4, ["count"] = 1, ["new_text"] = "    Beep\nEnd Sub" }),            // still line 4 of the original
        });
        Assert.Equal(false, (bool)r["isError"]!, (string?)r["text"]);
        Assert.Equal("Option Explicit\r\n' header\r\n\r\nPrivate Sub A()\r\n    Beep\r\nEnd Sub", ide.Text);
        Assert.Equal(4, int.Parse(ide.Replaces[0][2]), "the lower edit goes first");
        Assert.Equal(2, r["json"]!["edits"]!.AsArray().Count);
    }

    /// <summary>
    /// Improvement notes 29 and 30: a fresh module reads as "Option Explicit\r\n" with countOfLines 2 (VB counts the empty
    /// last line). Replacing both lines with text holding "" must succeed in one go, with no error after it was applied.
    /// </summary>
    [Test]
    public static async Task ReplacingAFreshModuleWithAnEmptyLiteralSucceeds()
    {
        using var ide = new FakeIde();
        ide.Text = "Option Explicit\r\n";
        var r = await Call(ide.Tools(), "code_replace_lines", new JsonObject
        {
            ["component"] = "modA", ["start_line"] = 1, ["count"] = 2, ["expected_hash"] = VbCode.Hash(ide.Text),
            ["new_text"] = "Option Explicit\n\nPublic Function Nz(s As String) As String\n    If s = \"\" Then Nz = \"-\" Else Nz = s\nEnd Function\n",
        });
        Assert.Equal(false, (bool)r["isError"]!, (string?)r["text"]);
        Assert.Equal("Option Explicit\r\n", ide.Replaces[0][4], "the empty last line is part of the guard");
        Assert.Equal("Option Explicit\r\n\r\nPublic Function Nz(s As String) As String\r\n    If s = \"\" Then Nz = \"-\" Else Nz = s\r\nEnd Function", ide.Text);
        Assert.Equal(1, ide.Replaces.Count, "one edit, no retry");
        Assert.True(r["json"]!["codeHints"] is null, "\"\" is no escape");
        Assert.Equal(VbCode.Hash(ide.Text), (string)r["json"]!["hash"]!);
    }

    /// <summary>An add-in error after the old lines were deleted (improvement notes 30): the result says the edit is half-applied.</summary>
    [Test]
    public static async Task AHalfAppliedEditIsReportedAsSuch()
    {
        using var ide = new FakeIde { FailAfterDelete = true };
        var r = await Call(ide.Tools(), "code_replace_lines", new JsonObject
        { ["component"] = "modA", ["start_line"] = 3, ["count"] = 2, ["new_text"] = "Private Sub B()\nEnd Sub", ["expected_hash"] = VbCode.Hash(Module) });
        var text = (string)r["text"]!;
        Assert.True((bool)r["isError"]! && text.Contains("half-applied") && text.Contains("Invalid procedure call"), text);
    }

    [Test]
    public static void ReplaceLinesMatchesTheAddinsNumbering()
    {
        Assert.Equal("a\r\nb", VbCode.ReplaceLines("Option Explicit\r\n", 1, 2, "a\r\nb"));
        Assert.Equal("x\r\ny\r\nz", VbCode.ReplaceLines("x\r\nz", 2, 0, "y"));
        Assert.Equal("x", VbCode.ReplaceLines("x\r\ny", 2, 1, ""));
    }

    [Test]
    public static async Task OverlappingEditsAreRefusedBeforeAnythingIsSent()
    {
        using var ide = new FakeIde();
        var r = await Call(ide.Tools(), "code_replace_lines", new JsonObject
        {
            ["component"] = "modA", ["expected_hash"] = VbCode.Hash(Module),
            ["edits"] = new JsonArray(
                new JsonObject { ["start_line"] = 3, ["count"] = 2, ["new_text"] = "x" },
                new JsonObject { ["start_line"] = 4, ["count"] = 0, ["new_text"] = "y" }),
        });
        Assert.True((bool)r["isError"]! && ((string)r["text"]!).Contains("overlap"), (string)r["text"]!);
        Assert.Equal(0, ide.Replaces.Count, "nothing sent");
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
