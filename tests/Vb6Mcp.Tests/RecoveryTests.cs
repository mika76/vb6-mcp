using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Core.Recovery;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;
using Vb6Mcp.Watchdog;

namespace Vb6Mcp.Tests;

public static class RecoveryTests
{
    const char US = Wire.US, RS = Wire.RS;

    /// <summary>A fake IDE with one project and two components; modMath is dirty.</summary>
    sealed class FakeIde : IDisposable
    {
        public string MathText = "Option Explicit\r\nPublic x As Long";
        public bool MathDirty = true;
        public List<string> Replaced = [];
        public FakeAddin Addin;
        public FakeIde()
        {
            Addin = new FakeAddin((verb, a) => verb switch
            {
                "PROJECTS" => ("OK", $"1{US}Basic{US}0{US}D:\\b\\Basic.vbp{US}1"),
                "COMPONENTS" => ("OK", $"modMath{US}1{US}{(MathDirty ? 1 : 0)}{US}2{US}D:\\b\\modMath.bas{RS}frmMain{US}5{US}0{US}9{US}D:\\b\\frmMain.frm"),
                "READ" => ("OK", $"{MathText.Split("\r\n").Length}{US}{(MathDirty ? 1 : 0)}{US}{MathText}"),
                "REPLACE" => Replace(a),
                _ => ("ERR", $"400{US}unknown {verb}"),
            });
        }
        (string, string) Replace(string[] a)
        {
            if (a[4] != MathText) return ("ERR", $"409{US}changed");
            Replaced.Add(a[5]); MathText = a[5]; MathDirty = true;
            return ("OK", $"{MathText.Split("\r\n").Length}{US}1");
        }
        public void Dispose() => Addin.Dispose();
    }

    static RecoveryStore TempStore() => new(Path.Combine(Path.GetTempPath(), $"vb6mcp-rec-{Guid.NewGuid():N}"));

    static RecoveryPoller Poller(RecoveryStore store, FakeIde ide, IdeMode mode, TimeSpan idle, DiscoveryResult? scan = null) =>
        new(store, () => scan ?? new DiscoveryResult([ide.Addin.Instance], [], []), i => new AddinClient(i, TimeSpan.FromSeconds(2)),
            _ => mode, () => idle, _ => { });

    [Test]
    public static async Task PollerCopiesUnsavedCodeOnlyWhenIdle()
    {
        using var ide = new FakeIde();
        var store = TempStore();
        int pid = ide.Addin.Instance.Pid;

        await Poller(store, ide, IdeMode.Design, TimeSpan.FromSeconds(1)).PollOnceAsync(default);
        Assert.Equal(0, store.Copies(pid).Count, "user not idle: nothing read");
        await Poller(store, ide, IdeMode.Run, TimeSpan.FromMinutes(1)).PollOnceAsync(default);
        Assert.Equal(0, store.Copies(pid).Count, "program running: nothing read");

        await Poller(store, ide, IdeMode.Design, TimeSpan.FromMinutes(1)).PollOnceAsync(default);
        var copy = store.Copies(pid).Single();
        Assert.Equal("modMath", copy.Component, "only the dirty component");
        Assert.Equal("D:\\b\\modMath.bas", copy.File);
        Assert.Equal(ide.MathText, store.ReadText(pid, "Basic", "modMath"));

        ide.MathDirty = false;                                       // the user saved
        await Poller(store, ide, IdeMode.Break, TimeSpan.FromMinutes(1)).PollOnceAsync(default);
        Assert.Equal(0, store.Copies(pid).Count, "saved component's copy removed");
        Assert.True(store.ReadSession(pid) is { Crashed: false }, "session recorded");
    }

    [Test]
    public static async Task CrashedSessionsAreKeptAndMarked()
    {
        using var ide = new FakeIde();
        var store = TempStore();
        int pid = ide.Addin.Instance.Pid;
        await Poller(store, ide, IdeMode.Design, TimeSpan.FromMinutes(1)).PollOnceAsync(default);
        // the IDE's discovery file is left behind but the process is gone: crashed or killed
        await Poller(store, ide, IdeMode.Design, TimeSpan.FromMinutes(1), new DiscoveryResult([], [ide.Addin.Instance], [])).PollOnceAsync(default);
        Assert.True(store.ReadSession(pid) is { Crashed: true }, "marked crashed");
        Assert.Equal(1, store.Copies(pid).Count, "copies kept");
    }

    sealed class Env(FakeIde ide, RecoveryStore store, IdeMode mode, int targetPid) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([ide.Addin.Instance with { Pid = targetPid }], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, $"Basic - Microsoft Visual Basic [{mode.ToString().ToLowerInvariant()}]", mode, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance with { Port = ide.Addin.Port, Token = ide.Addin.Token }, TimeSpan.FromSeconds(2));
        public RecoveryStore Recovery => store;
    }

    static async Task<(bool IsError, string Text)> Call(IReadOnlyList<Tool> tools, string name, JsonObject args)
    {
        var r = (await tools.First(t => t.Name == name).InvokeAsync(args, CancellationToken.None)).ToJson();
        return ((bool)r["isError"]!, (string)r["content"]![0]!["text"]!);
    }

    [Test]
    public static async Task ListReadRestoreDiscard()
    {
        using var ide = new FakeIde();
        var store = TempStore();
        const int crashedPid = 999_991;                               // not a running process
        store.WriteSession(new RecoverySession(crashedPid, "", DateTime.Now, true, ["D:\\b\\Basic.vbp"]));
        store.WriteCopy(crashedPid, new RecoveryCopy("Basic", "D:\\b\\Basic.vbp", "modMath", 1, "D:\\b\\modMath.bas", 3, "h", DateTime.Now.AddMinutes(-5)),
                        "Option Explicit\r\nPublic x As Long\r\nPublic y As Long");
        ide.MathDirty = false;
        var tools = new IdeTools(new Env(ide, store, IdeMode.Design, 4242)).All();

        var list = await Call(tools, "recovery_list", []);
        Assert.True(!list.IsError && list.Text.Contains("\"crashed\":true") && list.Text.Contains("modMath"), list.Text);
        var read = await Call(tools, "recovery_read", new JsonObject { ["from_pid"] = crashedPid, ["component"] = "modMath" });
        Assert.True(read.Text.Contains("Public y As Long"), read.Text);

        var restored = await Call(tools, "recovery_restore", new JsonObject { ["from_pid"] = crashedPid, ["component"] = "modMath" });
        Assert.True(!restored.IsError && restored.Text.Contains("\"status\":\"restored\""), restored.Text);
        Assert.Equal("Option Explicit\r\nPublic x As Long\r\nPublic y As Long", ide.MathText, "IDE got the copy (by file, not name)");

        var refused = await Call(tools, "recovery_discard", new JsonObject { ["from_pid"] = crashedPid });
        Assert.True(refused.IsError && refused.Text.Contains("confirm"), refused.Text);
        var done = await Call(tools, "recovery_discard", new JsonObject { ["from_pid"] = crashedPid, ["confirm"] = true });
        Assert.True(!done.IsError && store.Copies(crashedPid).Count == 0, done.Text);
    }

    [Test]
    public static async Task RestoreNeedsDesignModeAndKillNeedsForce()
    {
        using var ide = new FakeIde();
        var store = TempStore();
        store.WriteCopy(999_992, new RecoveryCopy("Basic", "", "modMath", 1, "D:\\b\\modMath.bas", 1, "h", DateTime.Now), "x");
        var tools = new IdeTools(new Env(ide, store, IdeMode.Break, 4242)).All();
        var r = await Call(tools, "recovery_restore", new JsonObject { ["from_pid"] = 999_992, ["component"] = "modMath" });
        Assert.True(r.IsError && r.Text.Contains("design mode"), r.Text);
        var k = await Call(tools, "ide_kill", []);
        Assert.True(k.IsError && k.Text.Contains("force"), k.Text);
        Assert.Equal(0, ide.Replaced.Count);
    }

    sealed class BusyEnv(FakeAddin watchdog, Func<IdeMode> mode) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([new InstanceInfo(4242, 1, "t", 1, "x", "", "")], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        // the debugged program loops: the IDE doesn't respond
        public IdeWindowState ReadWindows(int pid) => new(pid, true, $"P - Microsoft Visual Basic [{mode().ToString().ToLowerInvariant()}]", mode(), false, true, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromMilliseconds(300));
        public InstanceInfo? FindWatchdog() => watchdog.Instance;
    }

    [Test]
    public static async Task BreakGoesThroughTheWatchdogWhenTheIdeIsBusy()
    {
        var mode = IdeMode.Run;
        var asked = new List<string>();
        using var wd = new FakeAddin((verb, a) => { asked.Add($"{verb} {string.Join(" ", a)}"); mode = IdeMode.Break; return ("OK", $"break{US}vb6"); });
        var tools = new IdeTools(new BusyEnv(wd, () => mode)).All();
        var r = await Call(tools, "run_break", []);
        Assert.True(!r.IsError && r.Text.Contains("watchdog") && r.Text.Contains("\"mode\":\"break\""), r.Text);
        Assert.SequenceEqual(new[] { "BREAK 4242" }, asked, "add-in not tried: the IDE doesn't respond");
    }
}
