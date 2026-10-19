using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Watchdog;

namespace Vb6Mcp.Tests;

public static class WatchdogTests
{
    /// <summary>Fake actions: pid 4242 and window 0xABC are "VB6"; records what was asked.</summary>
    sealed class FakeActions : IWatchdogActions
    {
        public List<string> Calls = [];
        public bool Elevated => true;
        public int ParentPid => 4242;
        public bool IsVb6(int pid) => pid == 4242;
        public int WindowPid(long hwnd) => hwnd == 0xABC ? 4242 : 7;
        public Task<string> BreakAsync(int pid, CancellationToken ct) { Calls.Add($"break {pid}"); throw new WatchdogRefusedException("not in run mode"); }
        public string Press(long h) { Calls.Add($"press {h:X}"); return "pressed"; }
        public string Kill(int pid) { Calls.Add($"kill {pid}"); return "killed"; }
        public string Quit(int pid) { Calls.Add($"quit {pid}"); return "close-posted"; }
        public string Make(string vbp, string defines) { Calls.Add($"make {vbp} {defines}"); return $"77{Wire.US}C:\\x\\b.log{Wire.US}{vbp}"; }
        public Task<string> SnapAsync(long h, string id, CancellationToken ct) { Calls.Add($"snap {h:X} {id}"); return Task.FromResult($"C:\\x\\{id}.png{Wire.US}PrintWindow"); }
        public string Launch(string? project) { Calls.Add($"launch {project}"); return "99"; }
        public string Regsvr(string file, bool unregister) { Calls.Add($"regsvr {file} {unregister}"); return "77"; }
        public string RegStatus(int pid) { Calls.Add($"regstatus {pid}"); return "0"; }
    }

    static async Task<(FakeActions Actions, AddinClient Client, CancellationTokenSource Stop, string Discovery)> StartAsync()
    {
        var actions = new FakeActions();
        var discovery = Path.Combine(Path.GetTempPath(), $"vb6mcp-wd-{Guid.NewGuid():N}.json");
        var host = new WatchdogHost(actions, _ => { }, discovery);
        var stop = new CancellationTokenSource();
        _ = host.RunAsync(stop.Token);
        for (int i = 0; i < 50 && !File.Exists(discovery); i++) await Task.Delay(20);
        var info = Discovery.TryRead(discovery) ?? throw new InvalidOperationException("no discovery file");
        return (actions, new AddinClient(info), stop, discovery);
    }

    [Test]
    public static async Task PingAndDiscoveryFile()
    {
        var (_, client, stop, discovery) = await StartAsync();
        using (stop)
        {
            Assert.Equal("watchdog " + WatchdogHost.Version, client.Instance.AddinVersion, "discovery has the add-in file's fields");
            var f = (await client.SendAsync("PING")).Split(Wire.US);
            Assert.Equal(Environment.ProcessId.ToString(), f[0]);
            Assert.Equal("1", f[2], "elevated");
            Assert.Equal("4242", f[3], "parent");
            stop.Cancel();
            for (int i = 0; i < 50 && File.Exists(discovery); i++) await Task.Delay(20);
            Assert.True(!File.Exists(discovery), "discovery file deleted on shutdown");
        }
    }

    [Test]
    public static async Task TokenAndTargetsAreChecked()
    {
        var (actions, client, stop, _) = await StartAsync();
        using (stop)
        {
            var bad = new AddinClient(client.Instance with { Token = new string('0', 32) });
            Assert.Equal(401, (await bad.SendRawAsync("PING", [])).ErrorCode);
            Assert.Equal(400, (await client.SendRawAsync("KILL", ["4242"])).ErrorCode, "kill without force");
            Assert.Equal(400, (await client.SendRawAsync("KILL", ["1234", "force"])).ErrorCode, "not a VB6 process");
            Assert.Equal("killed", await client.SendAsync("KILL", "4242", "force"));
            Assert.Equal(400, (await client.SendRawAsync("QUIT", ["1234"])).ErrorCode, "quit: not a VB6 process");
            Assert.Equal("close-posted", await client.SendAsync("QUIT", "4242"));
            Assert.Equal(400, (await client.SendRawAsync("PRESS", ["DEF"])).ErrorCode, "window of another process");
            Assert.Equal("pressed", await client.SendAsync("PRESS", "ABC"));
            Assert.Equal(400, (await client.SendRawAsync("SNAP", ["ABC", "..\\evil"])).ErrorCode, "unsafe id");
            Assert.Equal(400, (await client.SendRawAsync("LAUNCH", ["C:\\nope\\x.txt"])).ErrorCode, "not an existing .vbp");
            var refused = await client.SendRawAsync("BREAK", ["4242"]);
            Assert.True(refused.ErrorCode == 409 && refused.ErrorMessage.Contains("run mode"), refused.ErrorMessage);
            Assert.SequenceEqual(new[] { "kill 4242", "quit 4242", "press ABC", "break 4242" }, actions.Calls);
        }
    }

    /// <summary>REGSVR needs the literal 'register' or 'unregister' (hard rules 9, 13); REGSTATUS a numeric pid.</summary>
    [Test]
    public static async Task RegsvrNeedsAnExplicitAction()
    {
        var (actions, client, stop, _) = await StartAsync();
        using (stop)
        {
            Assert.Equal(400, (await client.SendRawAsync("REGSVR", ["C:\\w\\a.dll"])).ErrorCode, "no action");
            Assert.Equal(400, (await client.SendRawAsync("REGSVR", ["C:\\w\\a.dll", "yes"])).ErrorCode, "not register/unregister");
            Assert.Equal(400, (await client.SendRawAsync("REGSTATUS", ["x"])).ErrorCode, "bad pid");
            Assert.Equal("77", await client.SendAsync("REGSVR", "C:\\w\\a.dll", "unregister"));
            Assert.Equal("0", await client.SendAsync("REGSTATUS", "77"));
            Assert.SequenceEqual(new[] { "regsvr C:\\w\\a.dll True", "regstatus 77" }, actions.Calls);
        }
    }

    /// <summary>The tray icons: 32x32, and the elevated one differs (the shield) so the state shows without hovering.</summary>
    [Test]
    public static void TrayIconsShowElevation()
    {
        using var plain = WatchdogTray.MakeIcon(false);
        using var admin = WatchdogTray.MakeIcon(true);
        using var a = plain.ToBitmap();
        using var b = admin.ToBitmap();
        Assert.True(a.Width == 32 && b.Height == 32, $"{a.Width}x{b.Height}");
        // the shield's yellow half sits in the lower right corner
        var px = b.GetPixel(27, 22);
        Assert.True(px.R > 200 && px.G > 150 && px.B < 100, "yellow shield: " + px);
        Assert.True(a.GetPixel(27, 22) != px, "no shield on the plain icon");
    }

    /// <summary>The real action checks the file before starting anything elevated: absolute, no '..', .dll/.ocx, under the roots.</summary>
    [Test]
    public static void RegsvrChecksTheFile()
    {
        var actions = new WatchdogActions(_ => { }, Path.GetTempPath());
        foreach (var bad in new[] { "a.dll", "C:\\w\\..\\a.dll", "C:\\w\\a.exe" })
            Assert.Throws<ArgumentException>(() => actions.Regsvr(bad, false));
        var outside = Path.Combine(Environment.SystemDirectory, "scrrun.dll");
        Assert.Throws<WatchdogRefusedException>(() => actions.Regsvr(outside, false));   // outside the roots
        Assert.Throws<WatchdogRefusedException>(() => actions.RegStatus(1));   // a pid it didn't start
    }
}
