using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

public static class DebugToolsTests
{
    const char US = Wire.US;

    [Test]
    public static void ResultAfterTakesOnlyTheNewOutput()
    {
        var before = "? 1+1\r\n 2 \r\n";
        Assert.Equal<string?>(null, IdeTools.ResultAfter(before, before, "? x"), "nothing new yet");
        Assert.Equal<string?>(null, IdeTools.ResultAfter(before + "? x\r\n", before, "? x"), "command echoed, no output yet");
        Assert.Equal("42", IdeTools.ResultAfter(before + "? x\r\n 42 \r\n", before, "? x"));
        // the same command evaluated again: take the output after the last occurrence
        var twice = "? x\r\n 41 \r\n? x\r\n 42 \r\n";
        Assert.Equal("42", IdeTools.ResultAfter(twice, "? x\r\n 41 \r\n", "? x"));
        Assert.Equal("Form1", IdeTools.ResultAfter("? Me.Caption\r\nForm1\r\n", "", "? Me.Caption"));
    }

    /// <summary>A fake IDE whose mode reacts to CMD like VB6 would.</summary>
    sealed class FakeDebugIde : IDisposable
    {
        public IdeMode Mode = IdeMode.Design;
        public List<string> Commands = [];
        public FakeAddin Addin;

        public FakeDebugIde(IdeMode mode)
        {
            Mode = mode;
            Addin = new FakeAddin((verb, a) => verb switch
            {
                "PROJECTS" => ("OK", $"1{US}Basic{US}0{US}D:\\b\\Basic.vbp{US}0"),
                "CMD" => Cmd(a[0]),
                "SETSEL" => ("OK", "selected"),
                _ => ("ERR", $"400{US}unknown {verb}"),
            });
        }

        (string, string) Cmd(string id)
        {
            Commands.Add(id);
            Mode = (id, Mode) switch
            {
                ("186", IdeMode.Design) => IdeMode.Run,
                ("186", IdeMode.Break) => IdeMode.Run,
                ("189", IdeMode.Run) => IdeMode.Break,
                ("228", _) => IdeMode.Design,
                _ => Mode,
            };
            return ("QUEUED", "command " + id);
        }

        public IReadOnlyList<Tool> Tools() => new IdeTools(new Env(this)) { CommandWait = TimeSpan.FromSeconds(2) }.All();
        public void Dispose() => Addin.Dispose();

        sealed class Env(FakeDebugIde ide) : IIdeEnvironment
        {
            public DiscoveryResult Scan() => new([ide.Addin.Instance], [], []);
            public int DeleteStale(DiscoveryResult result) => 0;
            public IdeWindowState ReadWindows(int pid) => new(pid, true, $"Basic - Microsoft Visual Basic [{ide.Mode.ToString().ToLowerInvariant()}]", ide.Mode, true, false, []);
            public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
        }
    }

    static async Task<(bool IsError, string Text)> Call(IReadOnlyList<Tool> tools, string name, JsonObject? args = null)
    {
        var r = (await tools.First(t => t.Name == name).InvokeAsync(args ?? [], CancellationToken.None)).ToJson();
        return ((bool)r["isError"]!, (string)r["content"]![0]!["text"]!);
    }

    [Test]
    public static async Task StartBreakContinueEnd()
    {
        using var ide = new FakeDebugIde(IdeMode.Design);
        var tools = ide.Tools();
        var start = await Call(tools, "run_start");
        Assert.True(!start.IsError && start.Text.Contains("\"mode\":\"run\""), start.Text);
        var brk = await Call(tools, "run_break");
        Assert.True(brk.Text.Contains("\"mode\":\"break\""), brk.Text);
        var cont = await Call(tools, "run_continue");
        Assert.True(cont.Text.Contains("\"mode\":\"run\""), cont.Text);
        var end = await Call(tools, "run_end");       // started by us, so no force needed
        Assert.True(!end.IsError && end.Text.Contains("\"mode\":\"design\""), end.Text);
        Assert.SequenceEqual(new[] { "186", "189", "186", "228" }, ide.Commands);
    }

    [Test]
    public static async Task EndingTheUsersRunNeedsForce()
    {
        using var ide = new FakeDebugIde(IdeMode.Break);   // the user is debugging
        var tools = ide.Tools();
        var refused = await Call(tools, "run_end");
        Assert.True(refused.IsError && refused.Text.Contains("force"), refused.Text);
        Assert.Equal(0, ide.Commands.Count, "nothing sent");
        var forced = await Call(tools, "run_end", new JsonObject { ["force"] = true });
        Assert.True(!forced.IsError, forced.Text);
    }

    [Test]
    public static async Task ModeGuardsAndBreakpointTracking()
    {
        using var ide = new FakeDebugIde(IdeMode.Design);
        var tools = ide.Tools();
        var step = await Call(tools, "step_over");
        Assert.True(step.IsError && step.Text.Contains("break mode"), step.Text);

        var on = await Call(tools, "bp_toggle", new JsonObject { ["component"] = "modMath", ["line"] = 8 });
        Assert.True(on.Text.Contains("\"breakpoint\":\"set\""), on.Text);
        var off = await Call(tools, "bp_toggle", new JsonObject { ["component"] = "modMath", ["line"] = 8 });
        Assert.True(off.Text.Contains("\"breakpoint\":\"cleared\""), off.Text);
        Assert.SequenceEqual(new[] { "51", "51" }, ide.Commands);
    }

    [Test]
    public static async Task AppInputNeedsARunningProgram()
    {
        using var ide = new FakeDebugIde(IdeMode.Design);
        var tools = ide.Tools();
        var click = await Call(tools, "app_click", new JsonObject { ["text"] = "Add" });
        Assert.True(click.IsError && click.Text.Contains("design mode"), click.Text);
        var set = await Call(tools, "app_set_text", new JsonObject { ["text"] = "x" });
        Assert.True(set.IsError && set.Text.Contains("design mode"), set.Text);
        var mouse = await Call(tools, "app_mouse", new JsonObject { ["action"] = "click", ["x"] = 1, ["y"] = 1 });
        Assert.True(mouse.IsError && mouse.Text.Contains("design mode"), mouse.Text);
        var scroll = await Call(tools, "app_scroll", new JsonObject { ["direction"] = "down" });
        Assert.True(scroll.IsError && scroll.Text.Contains("design mode"), scroll.Text);
        ide.Mode = IdeMode.Break;
        click = await Call(tools, "app_click", new JsonObject { ["text"] = "Add" });
        Assert.True(click.IsError && click.Text.Contains("break mode"), click.Text);
        Assert.Equal(0, ide.Commands.Count, "nothing sent");
    }

    [Test]
    public static void ValuesAreSplitOnlyWhenTheCountMatches()
    {
        Assert.SequenceEqual(new[] { "Long", "", "x y" }, IdeTools.SplitValues("Long|~||~|x y|~|", 3)!, "empty string value");
        Assert.SequenceEqual(new[] { "4", "11:outer/inner" }, IdeTools.SplitValues("4|~|11:outer/inner|~|", 2)!);
        Assert.SequenceEqual(new[] { " padded " }, IdeTools.SplitValues(" padded |~|", 1)!, "trailing separator keeps spaces");
        Assert.True(IdeTools.SplitValues("a|~|b|~|", 3) is null, "too few values");
        Assert.True(IdeTools.SplitValues("a|~|b|~|c", 2) is null, "missing final separator");
    }

    [Test]
    public static void RuntimeErrorDialogsAreParsed()
    {
        var d = new DialogInfo(1, "Microsoft Visual Basic", ["Run-time error '11':\r\n\r\nDivision by zero"], [new DialogButton("Debug", 2)]);
        Assert.True(IdeTools.RuntimeError(d) == (11, "Division by zero"), IdeTools.RuntimeError(d)?.ToString() ?? "null");
        var neg = new DialogInfo(1, "x", ["Run-time error '-2147217900':", "Automation error"], []);
        Assert.True(IdeTools.RuntimeError(neg) == (-2147217900, "Automation error"), IdeTools.RuntimeError(neg)?.ToString() ?? "null");
        Assert.True(IdeTools.RuntimeError(new DialogInfo(1, "x", ["Compile error:", "Variable not defined"], [])) is null);
    }
}
