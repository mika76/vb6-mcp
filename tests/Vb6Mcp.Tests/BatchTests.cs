using System.Text.Json.Nodes;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;

namespace Vb6Mcp.Tests;

/// <summary>batch_run: several tool calls in one, in order, through the same tool code (stand-in tools here).</summary>
public static class BatchTests
{
    static (Tool Batch, List<string> Ran) Make()
    {
        var ran = new List<string>();
        Tool T(string name, Func<JsonObject, ToolResult> body) =>
            new(name, name, Schema.Object(), (a, _) => { ran.Add(name + (a.Str("x") is { } x ? ":" + x : "")); return Task.FromResult(body(a)); });
        var tools = new List<Tool>
        {
            T("app_click", a => ToolResult.Ok(new { clicked = a.Str("x") })),
            T("app_screenshot", _ => ToolResult.Image([0x89, 0x50, 0x4E, 0x47], new { window = "Form1" })),
            T("app_wait", _ => throw new ToolException("timed out")),
            T("ide_launch", _ => ToolResult.Ok(new { })),
            T("code_read", _ => ToolResult.Ok(new { hash = "h1", added = new { name = "Command1" }, positions = new[] { new { left = 120 }, new { left = 960 } } })),
            T("edit", a => ToolResult.Ok(new { got = a.ToJsonString() })),
        };
        return (BatchTool.Create(tools), ran);
    }

    static JsonObject Call(string tool, JsonObject? args = null) => new() { ["tool"] = tool, ["args"] = args ?? new JsonObject() };

    [Test]
    public static async Task CallsRunInOrderAndImagesComeThrough()
    {
        var (batch, ran) = Make();
        var r = (await batch.InvokeAsync(new JsonObject { ["calls"] = new JsonArray(Call("app_click", new() { ["x"] = "OK" }), Call("app_screenshot")) }, CancellationToken.None)).ToJson();
        Assert.Equal("app_click:OK|app_screenshot", string.Join("|", ran));
        Assert.True(!(bool)r["isError"]!, r.ToJsonString());
        var content = r["content"]!.AsArray();
        Assert.True(((string)content[0]!["text"]!).StartsWith("[1] app_click: "), (string)content[0]!["text"]!);
        Assert.Equal("image", (string)content[1]!["type"]!);
        Assert.True(((string)content[^1]!["text"]!).Contains("\"ran\":2"), "summary last");
    }

    [Test]
    public static async Task AFailingCallStopsTheBatchUnlessToldToContinue()
    {
        var (batch, ran) = Make();
        var stop = (await batch.InvokeAsync(new JsonObject { ["calls"] = new JsonArray(Call("app_wait"), Call("app_click", new() { ["x"] = "after" })) }, CancellationToken.None)).ToJson();
        Assert.True((bool)stop["isError"]!, "a failed call makes the batch an error");
        Assert.Equal("app_wait", string.Join("|", ran));
        Assert.True(stop["content"]!.AsArray().Any(c => ((string?)c!["text"] ?? "").Contains("\"notRun\":1")), stop.ToJsonString());

        var (batch2, ran2) = Make();
        await batch2.InvokeAsync(new JsonObject { ["calls"] = new JsonArray(Call("app_wait"), Call("app_click", new() { ["x"] = "after" })), ["continue_on_error"] = true }, CancellationToken.None);
        Assert.Equal("app_wait|app_click:after", string.Join("|", ran2));
    }

    [Test]
    public static async Task ExcludedOrUnknownToolsRefuseTheWholeBatchBeforeAnythingRuns()
    {
        var (batch, ran) = Make();
        foreach (var bad in new[] { "ide_launch", "batch_run", "no_such_tool" })
        {
            var r = (await batch.InvokeAsync(new JsonObject { ["calls"] = new JsonArray(Call("app_click"), Call(bad)) }, CancellationToken.None)).ToJson();
            Assert.True((bool)r["isError"]!, bad);
        }
        Assert.Equal(0, ran.Count);
    }

    [Test]
    public static async Task ALaterCallUsesValuesFromAnEarlierResult()
    {
        var (batch, _) = Make();
        var r = (await batch.InvokeAsync(new JsonObject { ["calls"] = new JsonArray(Call("code_read"),
            Call("edit", new() { ["expected_hash"] = "$1.hash", ["name"] = "$1.added.name", ["left"] = "$1.positions[1].left", ["text"] = "Price = $1.00", ["nested"] = new JsonArray("$1.hash") })) }, CancellationToken.None)).ToJson();
        var got = (string)r["content"]![1]!["text"]!;
        Assert.True(!(bool)r["isError"]!, got);
        // values keep their JSON type; "$1.00" isn't a reference (a path starts with a letter)
        foreach (var want in new[] { "\\\"expected_hash\\\":\\\"h1\\\"", "\\\"name\\\":\\\"Command1\\\"", "\\\"left\\\":960", "Price = $1.00", "[\\\"h1\\\"]" })
            Assert.True(got.Contains(want), want + " in " + got);
    }

    [Test]
    public static async Task BadReferencesFailClearly()
    {
        var (batch, ran) = Make();
        var later = (await batch.InvokeAsync(new JsonObject { ["calls"] = new JsonArray(Call("edit", new() { ["x"] = "$2.hash" }), Call("code_read")) }, CancellationToken.None)).ToJson();
        Assert.True((bool)later["isError"]! && ran.Count == 0, "a reference to a later call refuses the batch up front");

        var missing = (await batch.InvokeAsync(new JsonObject { ["calls"] = new JsonArray(Call("code_read"), Call("edit", new() { ["x"] = "$1.nope" })) }, CancellationToken.None)).ToJson();
        var text = string.Join(" ", missing["content"]!.AsArray().Select(c => (string?)c!["text"]));
        Assert.True((bool)missing["isError"]! && text.Contains("has no nope") && text.Contains("hash"), text);

        var failed = (await batch.InvokeAsync(new JsonObject { ["calls"] = new JsonArray(Call("app_wait"), Call("edit", new() { ["x"] = "$1.hash" })), ["continue_on_error"] = true }, CancellationToken.None)).ToJson();
        Assert.True(string.Join(" ", failed["content"]!.AsArray().Select(c => (string?)c!["text"])).Contains("call 1 failed"), failed.ToJsonString());
    }

    [Test]
    public static async Task CompactCutsTheTextOfSuccessfulCallsButNotTheLast()
    {
        var (batch, _) = Make();
        var r = (await batch.InvokeAsync(new JsonObject { ["calls"] = new JsonArray(Call("edit", new() { ["x"] = new string('a', 500) }), Call("edit", new() { ["x"] = new string('b', 500) })), ["compact"] = true }, CancellationToken.None)).ToJson();
        var texts = r["content"]!.AsArray().Select(c => (string)c!["text"]!).ToList();
        Assert.True(texts[0].Contains("cut by compact") && texts[0].Length < 260, texts[0]);
        Assert.True(!texts[1].Contains("cut by compact") && texts[1].Contains(new string('b', 500)), "the last call in full");
    }
}
