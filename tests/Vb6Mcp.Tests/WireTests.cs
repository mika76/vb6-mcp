using Vb6Mcp.Core.Protocol;

namespace Vb6Mcp.Tests;

public static class WireTests
{
    const char US = Wire.US, RS = Wire.RS;

    [Test]
    public static async Task FrameRoundTripsIncludingAnsiAndNewlines()
    {
        string body = "Private Sub Form_Load()\r\n    MsgBox \"café €\"\r\nEnd Sub";
        var frame = Wire.EncodeFrame(body);
        // header counts bytes, and 1252 is single-byte: é and € are one byte each
        Assert.Equal($"{body.Length}\n", System.Text.Encoding.ASCII.GetString(frame, 0, frame.Length - body.Length));
        var read = await Wire.ReadFrameAsync(new MemoryStream(frame), CancellationToken.None);
        Assert.Equal(body, read);
    }

    [Test]
    public static async Task ReadsSeveralFramesFromOneStream()
    {
        var ms = new MemoryStream();
        foreach (var b in new[] { "one", "", "three" }) ms.Write(Wire.EncodeFrame(b));
        ms.Position = 0;
        Assert.Equal("one", await Wire.ReadFrameAsync(ms, CancellationToken.None));
        Assert.Equal("", await Wire.ReadFrameAsync(ms, CancellationToken.None));
        Assert.Equal("three", await Wire.ReadFrameAsync(ms, CancellationToken.None));
        Assert.Equal<string?>(null, await Wire.ReadFrameAsync(ms, CancellationToken.None));
    }

    [Test]
    public static async Task RejectsBadHeadersAndTruncation()
    {
        await AssertThrowsAsync(() => Wire.ReadFrameAsync(new MemoryStream("12x\nabc"u8.ToArray()), CancellationToken.None));
        await AssertThrowsAsync(() => Wire.ReadFrameAsync(new MemoryStream("10\nabc"u8.ToArray()), CancellationToken.None));
        await AssertThrowsAsync(() => Wire.ReadFrameAsync(new MemoryStream("999999999\n"u8.ToArray()), CancellationToken.None));
        await AssertThrowsAsync(() => Wire.ReadFrameAsync(new MemoryStream("\n"u8.ToArray()), CancellationToken.None));
    }

    [Test]
    public static void BuildsRequestAndRejectsSeparatorsInArguments()
    {
        var body = Wire.BuildRequest("tok", "7", "READ", "Project1", "Form1");
        Assert.Equal($"VB6MCP1{US}tok{US}7{US}READ{US}Project1{US}Form1", body);
        Assert.Throws<ProtocolException>(() => Wire.BuildRequest("tok", "1", "READ", $"bad{US}arg"));
        Assert.Throws<ProtocolException>(() => Wire.BuildRequest("tok", "1", "READ", $"bad{RS}arg"));
    }

    [Test]
    public static void ParsesResponsesKeepingSeparatorsInPayload()
    {
        var ok = Wire.ParseResponse($"VB6MCP1{US}7{US}OK{US}5{US}1{US}code with{US}more");
        Assert.Equal(ResponseStatus.Ok, ok.Status);
        Assert.Equal("7", ok.Id);
        Assert.SequenceEqual(new[] { "5", "1", $"code with{US}more" }, Wire.Fields(ok.Payload, 3));

        var err = Wire.ParseResponse($"VB6MCP1{US}8{US}ERR{US}404{US}component not found: Project1.Form9");
        Assert.Equal(ResponseStatus.Error, err.Status);
        Assert.Equal(404, err.ErrorCode);
        Assert.Equal("component not found: Project1.Form9", err.ErrorMessage);

        Assert.Equal(ResponseStatus.Queued, Wire.ParseResponse($"VB6MCP1{US}9{US}QUEUED{US}186").Status);
        Assert.Throws<ProtocolException>(() => Wire.ParseResponse($"NOPE{US}1{US}OK{US}"));
        Assert.Throws<ProtocolException>(() => Wire.ParseResponse($"VB6MCP1{US}1{US}MAYBE{US}"));
    }

    [Test]
    public static void SplitsTablePayloads()
    {
        var recs = Wire.Records($"Form1{US}5{US}1{RS}Module1{US}1{US}0");
        Assert.Equal(2, recs.Count);
        Assert.SequenceEqual(new[] { "Module1", "1", "0" }, recs[1]);
        Assert.Equal(0, Wire.Records("").Count);
    }

    static async Task AssertThrowsAsync(Func<Task> f)
    {
        try { await f(); }
        catch (ProtocolException) { return; }
        throw new AssertFailed("expected ProtocolException");
    }
}
