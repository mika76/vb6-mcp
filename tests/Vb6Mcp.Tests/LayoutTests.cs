using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>form_layout (PLAN.md E1): the arithmetic, and the tool against a fake designer.</summary>
public static class LayoutTests
{
    const char US = Wire.US, RS = Wire.RS;

    static Dictionary<string, LayoutBox> Boxes(params (string Name, double W, double H)[] c)
        => c.ToDictionary(x => x.Name, x => new LayoutBox(5000, 5000, x.W, x.H), StringComparer.OrdinalIgnoreCase);

    static IReadOnlyList<IReadOnlyList<string>> Rows(params string[][] r) => r;

    [Test]
    public static void ATwoColumnFormGetsColumnsFromItsWidestControls()
    {
        var current = Boxes(("lblName", 600, 255), ("txtName", 2000, 315), ("lblCity", 900, 255), ("txtCity", 1500, 315));
        var r = FormLayout.Compute(new LayoutSpec { Rows = Rows(["lblName", "txtName"], ["lblCity", "txtCity"]) }, current, (6000, 4000));
        // column 1 = 900 (widest label), column 2 = 2000; margin 120, gaps 120/90; rows 315 high
        Assert.Equal(new LayoutBox(120, 150, 900, 255), r.Boxes["lblName"]);   // centred in the 315 row: +30
        Assert.Equal(new LayoutBox(1140, 120, 2000, 315), r.Boxes["txtName"]);
        Assert.Equal(new LayoutBox(120, 555, 900, 255), r.Boxes["lblCity"]);
        Assert.Equal(new LayoutBox(1140, 525, 2000, 315), r.Boxes["txtCity"]);   // fill: as wide as the column
        Assert.Equal(new LayoutBox(120, 120, 3020, 720), r.Block);
        Assert.Equal(0, r.Warnings.Count);
    }

    [Test]
    public static void ButtonsGoToTheBottomRightCorner()
    {
        var current = Boxes(("cmdOK", 1200, 375), ("cmdCancel", 1200, 375));
        var r = FormLayout.Compute(new LayoutSpec { Rows = Rows(["cmdOK", "cmdCancel"]), Align = "right", Anchor = "bottom" }, current, (6000, 4000));
        Assert.Equal(new LayoutBox(3360, 3505, 1200, 375), r.Boxes["cmdOK"]);
        Assert.Equal(new LayoutBox(4680, 3505, 1200, 375), r.Boxes["cmdCancel"]);   // right edge 5880 = 6000 - margin
    }

    [Test]
    public static void ARepeatedNameSpansColumnsAndAnEmptyCellIsSkipped()
    {
        var current = Boxes(("lblA", 800, 255), ("txtA", 2000, 315), ("txtNotes", 500, 900), ("chkX", 1000, 255));
        var r = FormLayout.Compute(new LayoutSpec { Rows = Rows(["lblA", "txtA"], ["txtNotes", "txtNotes"], ["", "chkX"]) }, current, null);
        Assert.Equal(new LayoutBox(120, 525, 2920, 900), r.Boxes["txtNotes"]);   // 800 + 120 + 2000
        Assert.Equal(1040, r.Boxes["chkX"].Left);
        Assert.Equal(2000, r.Boxes["chkX"].Width);
    }

    [Test]
    public static void FixedSizesWinAndFillCanBeTurnedOff()
    {
        var current = Boxes(("a", 500, 300), ("b", 700, 300));
        var r = FormLayout.Compute(new LayoutSpec { Rows = Rows(["a", "b"]), ColumnWidths = [1000, null], Fill = false, Left = 0, Top = 0, VAlign = "top" }, current, null);
        Assert.Equal(new LayoutBox(0, 0, 500, 300), r.Boxes["a"]);
        Assert.Equal(new LayoutBox(1120, 0, 700, 300), r.Boxes["b"]);
    }

    [Test]
    public static void ABlockTooBigForTheContainerIsWarnedAbout()
    {
        var r = FormLayout.Compute(new LayoutSpec { Rows = Rows(["a", "b"]) }, Boxes(("a", 3000, 300), ("b", 3000, 300)), (4000, 3000));
        Assert.True(r.Warnings.Any(w => w.Contains("right edge")), string.Join("; ", r.Warnings));
        Assert.Equal((6360.0, 540.0), FormLayout.FitInner(r, 120));
    }

    [Test]
    public static void BadGridsAreRefused()
    {
        var current = Boxes(("a", 100, 100), ("b", 100, 100));
        foreach (var rows in new[] { Rows(["a", "nope"]), Rows(["a", "b", "a"]), Rows(["a"], ["a"]), Rows([""]) })
            try { FormLayout.Compute(new LayoutSpec { Rows = rows }, current, null); Assert.True(false, "refused: " + string.Join("/", rows.Select(r => string.Join(",", r)))); }
            catch (ToolException) { }
        try { FormLayout.Compute(new LayoutSpec { Rows = Rows(["a"]), Align = "right" }, current, null); Assert.True(false, "align right without a size"); }
        catch (ToolException ex) { Assert.True(ex.Message.Contains("size"), ex.Message); }
    }

    // ---- the tool, against a fake designer: a form (twips, 6000 x 4000 inside) with a frame and controls

    sealed class Env(FakeAddin addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "P - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    static FakeAddin Designer(List<string> sets)
    {
        var queued = new Dictionary<string, string[]>();
        return new FakeAddin((verb, args) =>
        {
            switch (verb)
            {
                case "PROJECTS": return ("OK", $"1{US}P{US}0{US}C:/x/P.vbp{US}0");
                case "DESIGN": queued[args[0]] = args; return ("QUEUED", args[0]);
                case "RESULT":
                    var q = queued[args[0]];
                    switch (q[1])
                    {
                        case "CONTROLS":
                            return ("OK", $"ok{US}0{US}" +
                                $"fraInfo{US}-1{US}VB.Frame{US}Frame{US}3{US}{US}120{US}120{US}3000{US}1500{RS}" +
                                $"lblName{US}-1{US}VB.Label{US}Label{US}1{US}fraInfo{US}0{US}0{US}600{US}255{RS}" +
                                $"txtName{US}-1{US}VB.TextBox{US}TextBox{US}2{US}fraInfo{US}0{US}0{US}2000{US}315{RS}" +
                                $"cmdOK{US}-1{US}VB.CommandButton{US}CommandButton{US}2{US}{US}0{US}0{US}1200{US}375{RS}" +
                                $"cmdCancel{US}-1{US}VB.CommandButton{US}CommandButton{US}2{US}{US}0{US}0{US}1200{US}375");
                        case "PROPS":
                            return ("OK", $"ok{US}0{US}ScaleWidth{US}n{US}6000{RS}ScaleHeight{US}n{US}4000{RS}ScaleMode{US}n{US}1{RS}Width{US}n{US}6240{RS}Height{US}n{US}4650");
                        case "SET":
                            sets.Add($"{q[4]}.{q[5]}={q[6]}");
                            return ("OK", $"ok{US}1{US}{q[5]}{US}n{US}{q[6]}");
                    }
                    return ("ERR", $"400{US}unexpected op");
                default: return ("ERR", $"400{US}unexpected {verb}");
            }
        });
    }

    static async Task<JsonObject> Call(FakeAddin fake, JsonObject args)
    {
        var r = (await new IdeTools(new Env(fake)).All().First(t => t.Name == "form_layout").InvokeAsync(args, CancellationToken.None)).ToJson();
        var text = (string)r["content"]![0]!["text"]!;
        return (bool)r["isError"]! ? new JsonObject { ["error"] = text } : JsonNode.Parse(text)!.AsObject();
    }

    static JsonArray Grid(params string[][] rows) => new(rows.Select(r => (JsonNode)new JsonArray(r.Select(c => (JsonNode)c).ToArray())).ToArray());

    [Test]
    public static async Task ButtonsAreMovedToTheFormsCornerThroughSet()
    {
        var sets = new List<string>();
        using var fake = Designer(sets);
        var r = await Call(fake, new JsonObject { ["component"] = "Form1", ["project"] = "P", ["rows"] = Grid(["cmdOK", "cmdCancel"]), ["align"] = "right", ["anchor"] = "bottom" });
        Assert.True(r["error"] is null, r.ToJsonString());
        Assert.Equal("cmdOK.Left=3360|cmdOK.Top=3505|cmdCancel.Left=4680|cmdCancel.Top=3505", string.Join("|", sets));   // widths already right
        Assert.Equal(4, (int)r["applied"]!);
    }

    [Test]
    public static async Task AFrameIsLaidOutBelowItsCaptionAndSizedToFit()
    {
        var sets = new List<string>();
        using var fake = Designer(sets);
        var r = await Call(fake, new JsonObject { ["component"] = "Form1", ["project"] = "P", ["container"] = "fraInfo", ["rows"] = Grid(["lblName", "txtName"]), ["size_container"] = true });
        Assert.True(r["error"] is null, r.ToJsonString());
        // first row at 180 (caption) + 120; frame = block right/bottom + margin
        Assert.Equal("fraInfo.Width=2960|fraInfo.Height=735|lblName.Left=120|lblName.Top=330|txtName.Left=840|txtName.Top=300", string.Join("|", sets));
    }

    [Test]
    public static async Task DryRunChangesNothingAndControlsOfAnotherContainerAreRefused()
    {
        var sets = new List<string>();
        using var fake = Designer(sets);
        var dry = await Call(fake, new JsonObject { ["component"] = "Form1", ["project"] = "P", ["rows"] = Grid(["cmdOK"], ["cmdCancel"]), ["dry_run"] = true });
        Assert.Equal(true, (bool)dry["dryRun"]!);
        Assert.Equal(0, sets.Count);
        var wrong = await Call(fake, new JsonObject { ["component"] = "Form1", ["project"] = "P", ["rows"] = Grid(["cmdOK", "txtName"]) });
        Assert.True(((string?)wrong["error"] ?? "").Contains("inside fraInfo"), wrong.ToJsonString());
        Assert.Equal(0, sets.Count);
    }

    /// <summary>Four controls moved by hand in one call: the result suggests the form_layout call, rows from where they are (round 6).</summary>
    [Test]
    public static async Task MovingSeveralControlsByHandSuggestsFormLayout()
    {
        var queued = new Dictionary<string, string[]>();
        using var fake = new FakeAddin((verb, args) =>
        {
            switch (verb)
            {
                case "PROJECTS": return ("OK", $"1{US}P{US}0{US}C:/x/P.vbp{US}0");
                case "DESIGN": queued[args[0]] = args; return ("QUEUED", args[0]);
                case "RESULT":
                    var q = queued[args[0]];
                    return q[1] switch
                    {
                        "CONTROLS" => ("OK", $"ok{US}1{US}" +
                            $"lblB{US}-1{US}VB.Label{US}Label{US}1{US}{US}120{US}600{US}800{US}255{RS}" +
                            $"txtA{US}-1{US}VB.TextBox{US}TextBox{US}2{US}{US}1000{US}150{US}2000{US}315{RS}" +
                            $"lblA{US}-1{US}VB.Label{US}Label{US}1{US}{US}120{US}120{US}800{US}255{RS}" +
                            $"txtB{US}-1{US}VB.TextBox{US}TextBox{US}2{US}{US}1000{US}630{US}2000{US}315"),
                        "SET" => ("OK", $"ok{US}1{US}{q[5]}{US}n{US}{q[6]}"),
                        _ => ("ERR", $"400{US}unexpected op"),
                    };
                default: return ("ERR", $"400{US}unexpected {verb}");
            }
        });
        var tool = new IdeTools(new Env(fake)).All().First(t => t.Name == "form_set_properties");
        JsonObject Move(string c) => new() { ["control"] = c, ["property"] = "Left", ["value"] = "130" };
        var r = (await tool.InvokeAsync(new JsonObject { ["component"] = "Form1", ["project"] = "P", ["changes"] = new JsonArray(Move("lblA"), Move("txtA"), Move("lblB"), Move("txtB")) }, CancellationToken.None)).ToJson();
        var text = (string)r["content"]![0]!["text"]!;
        Assert.True(text.Contains("form_layout component=Form1 rows=[[\\\"lblA\\\",\\\"txtA\\\"],[\\\"lblB\\\",\\\"txtB\\\"]]"), text);

        var three = (await tool.InvokeAsync(new JsonObject { ["component"] = "Form1", ["project"] = "P", ["changes"] = new JsonArray(Move("lblA"), Move("txtA"), Move("lblB")) }, CancellationToken.None)).ToJson();
        Assert.True(!((string)three["content"]![0]!["text"]!).Contains("form_layout component"), "fewer than 4: no suggestion");
    }

    /// <summary>Improvement notes 38: form_layout placed the controls itself, so its review must not tell the agent to use form_layout.</summary>
    [Test]
    public static async Task FormLayoutsOwnReviewDoesNotSuggestFormLayout()
    {
        var queued = new Dictionary<string, string[]>();
        using var fake = new FakeAddin((verb, args) =>
        {
            switch (verb)
            {
                case "PROJECTS": return ("OK", $"1{US}P{US}0{US}C:/x/P.vbp{US}0");
                case "DESIGN": queued[args[0]] = args; return ("QUEUED", args[0]);
                case "RESULT":
                    var q = queued[args[0]];
                    return q[1] switch
                    {
                        "CONTROLS" => ("OK", $"ok{US}1{US}" + string.Join(RS, new[] { "a", "b", "c", "d" }.Select(n => $"txt{n}{US}-1{US}VB.TextBox{US}TextBox{US}2{US}{US}5{US}5{US}2000{US}315"))),
                        "PROPS" => ("OK", $"ok{US}0{US}ScaleWidth{US}n{US}6000{RS}ScaleHeight{US}n{US}4000{RS}ScaleMode{US}n{US}1{RS}Width{US}n{US}6240{RS}Height{US}n{US}4650"),
                        "SET" => ("OK", $"ok{US}1{US}{q[5]}{US}n{US}{q[6]}"),
                        _ => ("ERR", $"400{US}unexpected op"),
                    };
                default: return ("ERR", $"400{US}unexpected {verb}");
            }
        });
        var r = await Call(fake, new JsonObject { ["component"] = "Form1", ["project"] = "P", ["rows"] = Grid(["txta", "txtb"], ["txtc", "txtd"]) });
        Assert.True(r["error"] is null, r.ToJsonString());
        Assert.True(!r.ToJsonString().Contains("by hand"), r.ToJsonString());
    }

    [Test]
    public static async Task TheFormIsResizedByItsBordersToo()
    {
        var sets = new List<string>();
        using var fake = Designer(sets);
        await Call(fake, new JsonObject { ["component"] = "Form1", ["project"] = "P", ["rows"] = Grid(["fraInfo"], ["cmdOK"]), ["size_container"] = true, ["fill"] = false });
        // inside 120+3000+120 = 3240 wide, 120+1500+90+375+120 = 2205 high; borders 240 x 650
        Assert.True(sets.Contains("Form1.Width=3480") && sets.Contains("Form1.Height=2855"), string.Join("|", sets));
    }
}
