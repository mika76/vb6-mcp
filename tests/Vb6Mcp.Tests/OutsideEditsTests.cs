using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>Edits no tool made (improvement notes 50), and clipped controls (33, 52).</summary>
public static class OutsideEditsTests
{
    const char US = Wire.US;

    [Test]
    public static void CompareIsNullForSameAndCaseOnlyText()
    {
        Assert.True(ToolEditLedger.Compare("m", "x = Err.Number", "x = Err.Number", "disk") is null, "same");
        Assert.True(ToolEditLedger.Compare("m", "x = Err.Number", "x = Err.number", "disk") is null, "case only");
    }

    [Test]
    public static void CompareReportsARealChangeWithItsDiff()
    {
        var o = ToolEditLedger.Compare("m", "a = 1\r\nb = 2", "a = 1\r\nnothe = 2", "last vb6-mcp edit")!;
        Assert.Equal("last vb6-mcp edit", o.Since);
        Assert.Equal(2, o.ChangedLines);   // one removed, one added
        Assert.True(o.Diff.Contains("+nothe = 2") && o.Diff.Contains("-b = 2"), o.Diff);
        Assert.True(o.MoreDiffLines is null, "short diff");
    }

    [Test]
    public static void CompareCapsALongDiff()
    {
        var before = string.Join("\r\n", Enumerable.Range(0, 100).Select(i => $"a{i} = {i}"));
        var after = string.Join("\r\n", Enumerable.Range(0, 100).Select(i => $"b{i} = {i}"));
        var o = ToolEditLedger.Compare("m", before, after, "disk")!;
        Assert.Equal(200, o.ChangedLines);
        Assert.Equal(ToolEditLedger.MaxDiffLines, o.Diff.Split('\n').Length);
        Assert.True(o.MoreDiffLines > 0, "the rest is counted");
    }

    [Test]
    public static void ClipStateTellsInsidePartlyAndHidden()
    {
        var clip = (L: 0, T: 0, R: 100, B: 100);
        Assert.True(AppWindows.ClipState((10, 10, 50, 50), clip) is null, "inside");
        Assert.Equal("partly", AppWindows.ClipState((80, 10, 150, 50), clip));
        Assert.Equal("hidden", AppWindows.ClipState((10, 120, 50, 160), clip));
        Assert.Equal("hidden", AppWindows.ClipState((10, 10, 10, 50), clip));   // no area
        Assert.Equal("hidden", AppWindows.ClipState((10, 10, 50, 50), (5, 5, 5, 5)));   // a zero-size container (2 x 2 PictureBox: item 52)
    }

    sealed class Env(FakeAddin addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "P - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    /// <summary>One dirty module "modA" whose file on disk holds <c>Option Explicit</c>; the IDE text starts equal to it.</summary>
    sealed class SaveIde : IDisposable
    {
        public string Dir = Directory.CreateTempSubdirectory("vb6mcp-outside-").FullName;
        public string Text = "Option Explicit\r\nx = 1";
        public bool Dirty = true;
        public FakeAddin Addin;
        public IReadOnlyList<Tool> Tools;

        public SaveIde()
        {
            var file = Path.Combine(Dir, "modA.bas");
            File.WriteAllBytes(file, Vb1252.Encoding.GetBytes("Attribute VB_Name = \"modA\"\r\nOption Explicit\r\nx = 1\r\n"));
            Addin = new FakeAddin((verb, a) => verb switch
            {
                "PROJECTS" => ("OK", $"1{US}P{US}0{US}{Path.Combine(Dir, "P.vbp")}{US}{(Dirty ? 1 : 0)}"),
                "COMPONENTS" => ("OK", $"modA{US}1{US}{(Dirty ? 1 : 0)}{US}{Text.Split("\r\n").Length}{US}{file}"),
                "READ" => ("OK", $"{Text.Split("\r\n").Length}{US}{(Dirty ? 1 : 0)}{US}{Text}"),
                "REPLACE" => Replace(a),
                "SAVE" => Save(),
                _ => ("ERR", $"400{US}unexpected {verb}"),
            });
            Tools = new IdeTools(new Env(Addin)) { SaveWait = TimeSpan.FromSeconds(5) }.All();
        }

        (string, string) Replace(string[] a)
        {
            int start = int.Parse(a[2]), count = int.Parse(a[3]);
            var lines = Text.Split("\r\n").ToList();
            lines.RemoveRange(start - 1, count);
            if (a[5].Length > 0) lines.InsertRange(start - 1, a[5].Split("\r\n"));
            Text = string.Join("\r\n", lines);
            Dirty = true;
            return ("OK", $"{lines.Count}{US}1");
        }

        (string, string) Save() { Dirty = false; return ("QUEUED", ""); }

        public void Dispose() { Addin.Dispose(); Directory.Delete(Dir, true); }
    }

    static async Task<(bool Error, string Text)> Call(IReadOnlyList<Tool> tools, string name, JsonObject args)
    {
        var res = (await tools.First(t => t.Name == name).InvokeAsync(args, CancellationToken.None)).ToJson();
        return ((bool)res["isError"]!, (string)res["content"]![0]!["text"]!);
    }

    [Test]
    public static async Task ProjectSaveRefusesTypingNoToolMadeUnlessAccepted()
    {
        using var ide = new SaveIde();
        ide.Text = "Option Explicit\r\nnothe = 1";   // stray keystrokes: differs from the file, no tool edit
        var refused = await Call(ide.Tools, "project_save", new JsonObject());
        Assert.True(refused.Error, "refused");
        Assert.True(refused.Text.Contains("modA") && refused.Text.Contains("+nothe = 1") && refused.Text.Contains("outside_edits"), refused.Text);
        Assert.True(ide.Dirty, "nothing was saved");

        var saved = await Call(ide.Tools, "project_save", new JsonObject { ["outside_edits"] = true });
        Assert.True(!saved.Error, saved.Text);
        var o = JsonNode.Parse(saved.Text)!;
        Assert.Equal("saved", (string)o["status"]!);
        Assert.Equal("modA", (string)o["outsideEdits"]![0]!["component"]!);
    }

    [Test]
    public static async Task ProjectSaveDoesNotRefuseAnEditMadeThroughTheTools()
    {
        using var ide = new SaveIde();
        var read = JsonNode.Parse((await Call(ide.Tools, "code_read", new JsonObject { ["component"] = "modA" })).Text)!;
        var edit = await Call(ide.Tools, "code_replace_lines", new JsonObject
        {
            ["component"] = "modA", ["start_line"] = 2, ["count"] = 1, ["new_text"] = "x = 2", ["expected_hash"] = (string)read["hash"]!,
        });
        Assert.True(!edit.Error, edit.Text);
        var saved = await Call(ide.Tools, "project_save", new JsonObject());
        Assert.True(!saved.Error, saved.Text);
        var o = JsonNode.Parse(saved.Text)!;
        Assert.Equal("saved", (string)o["status"]!);
        Assert.True(o["outsideEdits"] is null, "no outside edits reported");
    }

    [Test]
    public static async Task ProjectSaveWithComponentsRefusesTypingNoToolMade()
    {
        using var ide = new SaveIde();
        ide.Text = "Option Explicit\r\nnothe = 1";
        var refused = await Call(ide.Tools, "project_save", new JsonObject { ["components"] = new JsonArray("modA") });
        Assert.True(refused.Error && refused.Text.Contains("+nothe = 1"), refused.Text);
    }
}
