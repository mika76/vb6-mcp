using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>Crash notices with the next result (PLAN.md D2, improvement notes 25).</summary>
public static class CrashNoticeTests
{
    const string Wevtutil =
        "Event[0]:\r\n  Log Name: Application\r\n  Source: Application Error\r\n  Date: 2026-10-16T22:02:33.1230000Z\r\n  Event ID: 1000\r\n" +
        "  Description: \r\nFaulting application name: notepad.exe, version: 10.0\r\nFaulting process id: 0x1e8\r\n\r\n" +
        "Event[1]:\r\n  Log Name: Application\r\n  Source: Application Error\r\n  Date: 2026-10-16T22:02:33.4560000Z\r\n  Event ID: 1000\r\n" +
        "  Description: \r\nFaulting application name: VB6.EXE, version: 6.0.97.82, time stamp: 0x403acf8a\r\n" +
        "Faulting module name: KERNELBASE.dll, version: 10.0.22000.1, time stamp: 0x1\r\nException code: 0xc06d007e\r\nFault offset: 0x0012a0c2\r\n" +
        "Faulting process ID: 0x1e8\r\nFaulting module path: C:\\Windows\\System32\\KERNELBASE.dll\r\n";

    [Test]
    public static void TheApplicationErrorOfTheVb6ProcessIsFound()
    {
        var e = JsonNode.Parse(System.Text.Json.JsonSerializer.Serialize(IdeTools.ParseApplicationError(Wevtutil, 488)))!;   // 0x1e8
        Assert.Equal("KERNELBASE.dll, version: 10.0.22000.1, time stamp: 0x1", (string)e["faultingModule"]!);
        Assert.Equal("0xc06d007e", (string)e["exceptionCode"]!);
        Assert.True(IdeTools.ParseApplicationError(Wevtutil, 489) is null, "another pid");
    }

    sealed class Env(Func<DiscoveryResult> scan) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => scan();
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => throw new InvalidOperationException();
        public AddinClient Client(InstanceInfo instance) => throw new InvalidOperationException();
    }

    [Test]
    public static async Task ACrashIsReportedOnceWithTheNextResult()
    {
        var ide = new InstanceInfo(488, 1, new string('a', 32), 1, "0.1.0", "", "");
        DiscoveryResult state = new([ide], [], []);
        var tools = new IdeTools(new Env(() => state)) { ApplicationLogEntry = _ => new { exceptionCode = "0xc06d007e" } };
        var tool = tools.All().First(t => t.Name == "recovery_list");
        async Task<string> Call() => string.Join(" ", (await tool.InvokeAsync([], CancellationToken.None)).ToJson()["content"]!.AsArray().Select(c => (string?)c!["text"]));

        Assert.True(!(await Call()).Contains("ideCrashed"), "running: nothing to report");
        state = new([], [ide], []);   // discovery file left behind, process gone
        var crashed = await Call();
        Assert.True(crashed.Contains("ideCrashed") && crashed.Contains("0xc06d007e") && crashed.Contains("reopen=true"), crashed);
        Assert.True(!(await Call()).Contains("ideCrashed"), "reported once");

        // an IDE that closed normally (no file left) isn't a crash
        var other = ide with { Pid = 500 };
        state = new([other], [], []); await Call();
        state = new([], [], []);
        Assert.True(!(await Call()).Contains("ideCrashed"), "clean exit");
    }
}
