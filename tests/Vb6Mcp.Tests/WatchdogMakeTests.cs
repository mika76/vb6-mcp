using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;
using Vb6Mcp.Watchdog;

namespace Vb6Mcp.Tests;

/// <summary>build_make with no IDE running: the elevated watchdog starts /make (PLAN.md D1).</summary>
public static class WatchdogMakeTests
{
    [Test]
    public static void DefinesAreNameIntegerPairsOnly()
    {
        Assert.True(WatchdogActions.ValidDefines("") && WatchdogActions.ValidDefines("noterm=1:debug=0") && WatchdogActions.ValidDefines("x_1=-1"), "valid");
        foreach (var bad in new[] { "a=b", "1a=1", "a=1 & calc", "a=1:", "a", "a=1\" /x \"" })
            Assert.True(!WatchdogActions.ValidDefines(bad), bad);
    }

    sealed class Actions(string log) : IWatchdogActions
    {
        public List<string> Calls = [];
        public bool Elevated => true;
        public int ParentPid => 1;
        public bool IsVb6(int pid) => true;
        public int WindowPid(long hwnd) => 1;
        public Task<string> BreakAsync(int pid, CancellationToken ct) => throw new NotSupportedException();
        public string Press(long h) => throw new NotSupportedException();
        public string Kill(int pid) => throw new NotSupportedException();
        public string Quit(int pid) => throw new NotSupportedException();
        // a pid that isn't running: the build is "done" at once, and its log is already written
        public string Make(string vbp, string defines) { Calls.Add($"make {vbp} {defines}"); return string.Join(Wire.US, "2147483000", log, vbp); }
        public Task<string> SnapAsync(long h, string id, CancellationToken ct) => throw new NotSupportedException();
        public string Launch(string? project) => throw new NotSupportedException();
        public string Regsvr(string file, bool unregister) => throw new NotSupportedException();
        public string RegStatus(int pid) => throw new NotSupportedException();
    }

    sealed class NoIde(InstanceInfo? watchdog) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => throw new InvalidOperationException();
        public AddinClient Client(InstanceInfo instance) => throw new InvalidOperationException();
        public InstanceInfo? FindWatchdog() => watchdog;
    }

    [Test]
    public static async Task AClosedProjectBuildsThroughTheWatchdogWhenNoIdeRuns()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-wdmake-").FullName;
        var discovery = Path.Combine(Path.GetTempPath(), $"vb6mcp-wd-{Guid.NewGuid():N}.json");
        using var stop = new CancellationTokenSource();
        try
        {
            var vbp = Path.Combine(dir, "Basic.vbp");
            File.WriteAllBytes(vbp, Vb1252.Encoding.GetBytes("Type=Exe\r\nName=\"Basic\"\r\nExeName32=\"Basic.exe\"\r\n"));
            var log = Path.Combine(dir, "build.log");
            File.WriteAllText(log, "Build of 'Basic.exe' succeeded.");
            var actions = new Actions(log);
            _ = new WatchdogHost(actions, _ => { }, discovery).RunAsync(stop.Token);
            for (int i = 0; i < 50 && !File.Exists(discovery); i++) await Task.Delay(20);
            var wd = Discovery.TryRead(discovery)!;

            var tools = new IdeTools(new NoIde(wd)) { Roots = [dir] };
            var r = (await tools.All().First(t => t.Name == "build_make").InvokeAsync(new JsonObject { ["path"] = vbp, ["defines"] = new JsonObject { ["noterm"] = 1 } }, CancellationToken.None)).ToJson();
            var o = JsonNode.Parse((string)r["content"]![0]!["text"]!)!;
            Assert.Equal("succeeded", (string?)o["status"] ?? o.ToJsonString());
            Assert.True(((string)o["via"]!).Contains("watchdog"), (string)o["via"]!);
            Assert.Equal($"make {Path.GetFullPath(vbp)} noterm=1", string.Join("|", actions.Calls));

            var none = (await new IdeTools(new NoIde(null)) { Roots = [dir] }.All().First(t => t.Name == "build_make").InvokeAsync(new JsonObject { ["path"] = vbp }, CancellationToken.None)).ToJson();
            Assert.True((bool)none["isError"]! && ((string)none["content"]![0]!["text"]!).Contains("watchdog --stay"), "without IDE and watchdog: says how to get one");
        }
        finally { stop.Cancel(); Directory.Delete(dir, true); }
    }

    /// <summary>
    /// Spike 41: with no elevated VB6 and no elevated watchdog, com_register refuses (a regsvr32 at our level can only
    /// fail) and says how to elevate; it never shows a UAC prompt without elevate=true.
    /// </summary>
    [Test]
    public static async Task RegisteringWithoutElevationAsksForElevate()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-reg-").FullName;
        try
        {
            var dll = Path.Combine(dir, "A.dll");
            File.WriteAllBytes(dll, [0x4D, 0x5A]);
            var tool = new WorkspaceTools(new NoIde(null), [dir]).All().First(t => t.Name == "com_register");
            var r = (await tool.InvokeAsync(new JsonObject { ["file"] = dll, ["confirm"] = true }, CancellationToken.None)).ToJson();
            var text = (string)r["content"]![0]!["text"]!;
            Assert.True((bool)r["isError"]! && text.Contains("elevate=true"), text);
        }
        finally { Directory.Delete(dir, true); }
    }
}
