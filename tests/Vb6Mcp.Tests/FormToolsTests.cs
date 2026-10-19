using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

public static class FormToolsTests
{
    const char US = Wire.US, RS = Wire.RS;

    sealed class Env(FakeAddin addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "P - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    /// <summary>A fake add-in whose deferred SET refuses Left=x and answers Font.Bold with every Font member, as VB does.</summary>
    static FakeAddin SetFake(List<string> sets) {
        var queued = new Dictionary<string, string[]>();
        return new FakeAddin((verb, args) =>
        {
            switch (verb)
            {
                case "PROJECTS": return ("OK", $"1{US}P{US}0{US}C:/x/P.vbp{US}0");
                case "DESIGN": queued[args[0]] = args; return ("QUEUED", args[0]);
                case "RESULT":
                    var q = queued[args[0]];   // id, SET, project, component, control, property, value
                    sets.Add($"{q[4]}.{q[5]}={q[6]}");
                    if (q[5] == "Left" && q[6] == "x") return ("OK", $"failed{US}-2147221104: not a number: x");
                    if (q[5] == "Font.Bold") return ("OK", $"ok{US}1{US}Font.Name{US}s{US}Arial{RS}Font.Bold{US}b{US}True{RS}Font.Size{US}n{US}9");
                    return ("OK", $"ok{US}1{US}{q[5]}{US}s{US}{q[6]}");
                default: return ("ERR", $"400{US}unexpected {verb}");
            }
        });
    }

    static async Task<JsonObject> Call(FakeAddin fake, JsonObject args)
    {
        var r = (await new IdeTools(new Env(fake)).All().First(t => t.Name == "form_set_properties").InvokeAsync(args, CancellationToken.None)).ToJson();
        var text = (string)r["content"]![0]!["text"]!;
        return (bool)r["isError"]! ? new JsonObject { ["error"] = text } : JsonNode.Parse(text)!.AsObject();
    }

    static JsonObject Change(string c, string p, string v) => new() { ["control"] = c, ["property"] = p, ["value"] = v };

    /// <summary>All changes go through, in order; a refused value is reported and the rest still run.</summary>
    [Test]
    public static async Task SetsSeveralPropertiesAndReportsEachOne()
    {
        var sets = new List<string>();
        using var fake = SetFake(sets);
        var r = await Call(fake, new JsonObject
        {
            ["component"] = "Form1", ["project"] = "P",
            ["changes"] = new JsonArray(Change("Form1", "Caption", "Hello"), Change("Text1", "Left", "x"), Change("Label1", "Font.Bold", "True")), ["verbose"] = true,
        });
        Assert.Equal("Form1.Caption=Hello|Text1.Left=x|Label1.Font.Bold=True", string.Join("|", sets));
        Assert.Equal(2, (int)r["applied"]!);
        Assert.Equal(1, (int)r["failed"]!);
        var results = r["results"]!.AsArray();
        Assert.Equal("Hello", (string)results[0]!["now"]!);
        Assert.Equal("not a number: x", (string)results[1]!["error"]!);
        Assert.Equal("True", (string)results[2]!["now"]!);   // only Font.Bold, not every Font member
    }

    /// <summary>Without verbose only the failed changes are listed (token breakdown: echoing every change was 15% of a run).</summary>
    [Test]
    public static async Task OnlyFailedChangesAreListedByDefault()
    {
        var sets = new List<string>();
        using var fake = SetFake(sets);
        var r = await Call(fake, new JsonObject { ["component"] = "Form1", ["project"] = "P", ["changes"] = new JsonArray(Change("Form1", "Caption", "Hello"), Change("Text1", "Left", "x")) });
        Assert.Equal(1, (int)r["applied"]!);
        Assert.Equal(1, r["results"]!.AsArray().Count);
        Assert.Equal("Left", (string)r["results"]![0]!["property"]!);
    }

    /// <summary>A malformed list changes nothing: every item is checked before the first SET.</summary>
    [Test]
    public static async Task AMalformedListChangesNothing()
    {
        var sets = new List<string>();
        using var fake = SetFake(sets);
        var r = await Call(fake, new JsonObject
        {
            ["component"] = "Form1", ["project"] = "P",
            ["changes"] = new JsonArray(Change("Form1", "Caption", "Hello"), new JsonObject { ["control"] = "Text1" }),
        });
        Assert.True(((string?)r["error"] ?? "").Contains("changes[1]"), r.ToJsonString());
        Assert.Equal(0, sets.Count);
    }

    /// <summary>Some clients send the list as a JSON string.</summary>
    [Test]
    public static async Task AcceptsTheListAsAJsonString()
    {
        var sets = new List<string>();
        using var fake = SetFake(sets);
        var r = await Call(fake, new JsonObject
        {
            ["component"] = "Form1", ["project"] = "P",
            ["changes"] = new JsonArray(Change("Form1", "BackColor", "&H00FAF7F5&")).ToJsonString(),
        });
        Assert.Equal(1, (int)r["applied"]!);
        Assert.Equal("Form1.BackColor=&H00FAF7F5&", string.Join("|", sets));
    }

    /// <summary>
    /// A fake designer like a production project's frmMain after T9 run 1: two buttons in Frame3, VB's answers for Font members
    /// (not counted as a change), ForeColor on a button (438) and button Style 0.
    /// </summary>
    static FakeAddin DesignerFake(List<string> sets, bool fontCounts = false)
    {
        var queued = new Dictionary<string, string[]>();
        bool dirty = false;
        return new FakeAddin((verb, args) =>
        {
            switch (verb)
            {
                case "PROJECTS": return ("OK", $"1{US}P{US}0{US}C:/x/P.vbp{US}0");
                case "DESIGN": queued[args[0]] = args; return ("QUEUED", args[0]);
                case "RESULT":
                    var q = queued[args[0]];   // id, op, project, component, ...
                    switch (q[1])
                    {
                        case "CONTROLS":
                            return ("OK", $"ok{US}{(dirty ? 1 : 0)}{US}" +
                                $"Frame3{US}-1{US}VB.Frame{US}Frame{US}3{US}{US}120{US}5000{US}9800{US}1400{RS}" +
                                $"cmdRefund{US}-1{US}VB.CommandButton{US}CommandButton{US}2{US}Frame3{US}8250{US}300{US}1260{US}975{RS}" +
                                $"cmdFixAccount{US}-1{US}VB.CommandButton{US}CommandButton{US}2{US}Frame3{US}8250{US}300{US}1260{US}975{RS}" +
                                $"lblInfo{US}-1{US}VB.Label{US}Label{US}2{US}Frame3{US}150{US}100{US}1000{US}200{RS}" +
                                $"lblNote{US}-1{US}VB.Label{US}Label{US}2{US}Frame3{US}160{US}110{US}1000{US}200{RS}" +
                                $"lblHeader{US}-1{US}VB.Label{US}Label{US}2{US}{US}0{US}4800{US}9000{US}600{RS}" +
                                $"lblBand{US}-1{US}VB.Label{US}Label{US}2{US}Frame3{US}0{US}0{US}9800{US}1400{RS}" +
                                $"imgIcon{US}0{US}VB.Image{US}Image{US}1{US}{US}0{US}0{US}480{US}480{RS}" +
                                $"imgIcon{US}1{US}VB.Image{US}Image{US}1{US}{US}0{US}0{US}480{US}480");
                        case "ADD":   // id, ADD, project, component, class, name, container, geometry
                            return ("OK", $"ok{US}1{US}{(q[5].Length > 0 ? q[5] : "Command1")}{US}-1{US}{q[4]}{US}CommandButton{US}2{US}{q[6]}{US}100{US}100{US}1200{US}400");
                        case "PROPS":
                            return ("OK", q[4] == "Form1" ? $"ok{US}{(dirty ? 1 : 0)}{US}Caption{US}s{US}Main{RS}Tag{US}s{US}x" : $"ok{US}{(dirty ? 1 : 0)}{US}Style{US}n{US}0");
                        case "SET":
                            sets.Add($"{q[4]}.{q[5]}={q[6]}");
                            if (q[5] == "ForeColor") return ("OK", $"failed{US}438: Object doesn't support this property or method");
                            if (!q[5].Contains('.') || fontCounts) dirty = true;   // VB doesn't count Font.* member changes
                            return ("OK", $"ok{US}{(dirty ? 1 : 0)}{US}{q[5]}{US}s{US}{q[6]}");
                    }
                    return ("ERR", $"400{US}unexpected op");
                default: return ("ERR", $"400{US}unexpected {verb}");
            }
        });
    }

    /// <summary>T9 run 1: a button moved onto its sibling is reported; the label and the control array on the same spot aren't.</summary>
    [Test]
    public static async Task ControlsThatEndUpOnTopOfEachOtherAreReported()
    {
        var sets = new List<string>();
        using var fake = DesignerFake(sets);
        var r = await Call(fake, new JsonObject
        {
            ["component"] = "Form1", ["project"] = "P",
            ["changes"] = new JsonArray(Change("cmdFixAccount", "Left", "8250")),
        });
        var o = r["overlaps"]!.AsArray();
        Assert.Equal(1, o.Count);
        Assert.Equal("cmdRefund,cmdFixAccount", string.Join(",", o[0]!["controls"]!.AsArray().Select(n => (string)n!)));
        Assert.True((bool)o[0]!["sameRectangle"]!, "same rectangle");
        Assert.True(r["hints"]!.AsArray().Any(h => ((string)h!).Contains("form_screenshot")), "screenshot hint after a move");
    }

    /// <summary>A Label under a frame is reported (VB draws it behind windowed controls); label on label isn't (T10).</summary>
    [Test]
    public static async Task ALabelUnderAWindowedControlIsReported()
    {
        var sets = new List<string>();
        using var fake = DesignerFake(sets);
        var r = await Call(fake, new JsonObject
        {
            ["component"] = "Form1", ["project"] = "P",
            ["changes"] = new JsonArray(Change("lblHeader", "Top", "4800"), Change("lblInfo", "Left", "150"), Change("lblBand", "Top", "0")),
        });
        var pairs = r["overlaps"]!.AsArray().Select(o => string.Join("/", o!["controls"]!.AsArray().Select(n => (string)n!))).ToList();
        Assert.True(pairs.Contains("Frame3/lblHeader"), string.Join("; ", pairs));
        Assert.True(!pairs.Any(p => p.Contains("lblNote")), "label on label isn't reported: " + string.Join("; ", pairs));
        Assert.True(!pairs.Any(p => p.Contains("lblBand")), "a label behind whole controls is a background: " + string.Join("; ", pairs));
    }

    /// <summary>Without a geometry change nothing is checked (no CONTROLS read), even if controls already overlap.</summary>
    [Test]
    public static async Task NoOverlapCheckWithoutGeometryChanges()
    {
        var sets = new List<string>();
        using var fake = DesignerFake(sets);
        var r = await Call(fake, new JsonObject { ["component"] = "Form1", ["project"] = "P", ["changes"] = new JsonArray(Change("Form1", "Caption", "Orders")) });
        Assert.True(r["overlaps"] is null && r["hints"] is null, r.ToJsonString());
    }

    /// <summary>A change only to Font.* members isn't counted by VB; the tool makes VB count it (Tag set to itself).</summary>
    [Test]
    public static async Task FontOnlyChangesAreMadeToCount()
    {
        var sets = new List<string>();
        using var fake = DesignerFake(sets);
        var r = await Call(fake, new JsonObject
        {
            ["component"] = "Form1", ["project"] = "P",
            ["changes"] = new JsonArray(Change("Form1", "Font.Name", "Segoe UI")),
        });
        Assert.Equal("Form1.Font.Name=Segoe UI|Form1.Tag=x", string.Join("|", sets));
        Assert.True((bool)r["dirty"]!, "dirty after the Tag rewrite");
        Assert.True(r["hints"]!.AsArray().Any(h => ((string)h!).Contains("doesn't change the controls")), "form font hint");
    }

    /// <summary>ForeColor on a button: VB's 438 plus what it means; BackColor on a Style 0 button gets a hint.</summary>
    [Test]
    public static async Task ButtonColourRulesAreExplained()
    {
        var sets = new List<string>();
        using var fake = DesignerFake(sets);
        var r = await Call(fake, new JsonObject
        {
            ["component"] = "Form1", ["project"] = "P",
            ["changes"] = new JsonArray(Change("cmdRefund", "ForeColor", "&H00FFFFFF&"), Change("cmdRefund", "BackColor", "&H00FF0000&")),
        });
        Assert.True(((string)r["results"]![0]!["error"]!).Contains("no ForeColor"), r.ToJsonString());
        Assert.True(r["hints"]!.AsArray().Any(h => ((string)h!).Contains("Style = 1")), r.ToJsonString());
    }

    /// <summary>form_add_control sets properties on the new control in the same call (benchmark T8); a refused one is reported.</summary>
    [Test]
    public static async Task AddControlSetsPropertiesInTheSameCall()
    {
        var sets = new List<string>();
        using var fake = DesignerFake(sets);
        var res = (await new IdeTools(new Env(fake)).All().First(t => t.Name == "form_add_control").InvokeAsync(new JsonObject
        {
            ["component"] = "Form1", ["project"] = "P", ["prog_id"] = "VB.CommandButton", ["name"] = "cmdOk",
            ["properties"] = new JsonObject { ["Caption"] = "OK", ["Default"] = true, ["ForeColor"] = "&H00FFFFFF&" },
        }, CancellationToken.None)).ToJson();
        var r = JsonNode.Parse((string)res["content"]![0]!["text"]!)!.AsObject();
        Assert.Equal("cmdOk.Caption=OK|cmdOk.Default=true|cmdOk.ForeColor=&H00FFFFFF&", string.Join("|", sets));
        var props = r["properties"]!.AsArray();
        Assert.True((bool)props[0]!["ok"]! && (bool)props[1]!["ok"]!, r.ToJsonString());
        Assert.True(!(bool)props[2]!["ok"]! && ((string)props[2]!["error"]!).Contains("no ForeColor"), r.ToJsonString());
    }

    /// <summary>
    /// Improvement notes 43: a control added without left/top sits at the default position with the others still to be
    /// arranged; its overlaps are counted, not listed (they grew with the form). With a position they're listed.
    /// </summary>
    [Test]
    public static async Task AddControlListsOverlapsOnlyWhenPlaced()
    {
        async Task<JsonObject> Add(JsonObject extra)
        {
            using var fake = DesignerFake([]);
            var args = new JsonObject { ["component"] = "Form1", ["project"] = "P", ["prog_id"] = "VB.CommandButton", ["name"] = "cmdRefund", ["container"] = "Frame3" };
            foreach (var (k, v) in extra) args[k] = v?.DeepClone();
            var res = (await new IdeTools(new Env(fake)).All().First(t => t.Name == "form_add_control").InvokeAsync(args, CancellationToken.None)).ToJson();
            return JsonNode.Parse((string)res["content"]![0]!["text"]!)!.AsObject();
        }
        var unplaced = await Add([]);
        Assert.True(unplaced["overlaps"] is null, unplaced.ToJsonString());
        Assert.True(unplaced["hints"]!.AsArray().Any(h => ((string)h!).Contains("overlaps 1 other control(s) (not listed")), unplaced.ToJsonString());

        var placed = await Add(new JsonObject { ["left"] = 8250.0, ["top"] = 300.0 });
        Assert.Equal(1, placed["overlaps"]!.AsArray().Count);
        var byProperty = await Add(new JsonObject { ["properties"] = new JsonObject { ["Left"] = "8250" } });
        Assert.Equal(1, byProperty["overlaps"]?.AsArray().Count, byProperty.ToJsonString());
    }
}
