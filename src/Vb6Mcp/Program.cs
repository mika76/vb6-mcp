// vb6mcp: MCP server for the VB6 IDE. See docs/architecture.md and docs/tools.md.
// Usage: vb6mcp [mcp]       MCP server on stdio (default; what Claude Code starts)
//        vb6mcp watchdog    elevated sidecar, started by the VB6 add-in (PLAN.md §3.1)
// Diagnostics go to stderr, or to the file named by VB6MCP_LOG (some clients hide stderr).
using System.Text;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.Watchdog;

namespace Vb6Mcp;

public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        string mode = args.Length > 0 ? args[0].ToLowerInvariant() : "mcp";
        if (mode == "watchdog") return await WatchdogMainAsync();
        if (mode != "mcp")
        {
            Console.Error.WriteLine($"vb6mcp: unknown mode '{mode}'. Usage: vb6mcp [mcp|watchdog]");
            return 2;
        }

        var log = OpenLog(Environment.GetEnvironmentVariable("VB6MCP_LOG"));

        var stdin = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
        var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { NewLine = "\n" };
        // Workspace roots for the disk tools (hard rule 8: paths are checked against them)
        var roots = (Environment.GetEnvironmentVariable("VB6MCP_ROOTS") ?? "")
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var env = new IdeEnvironment();
        var tools = new IdeTools(env) { Roots = roots }.All().Concat(new WorkspaceTools(env, roots).All()).ToList();
        log.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} vb6mcp {McpServer.ServerVersion} started, pid {Environment.ProcessId}, {tools.Count} tools, roots: {(roots.Length == 0 ? "(none)" : string.Join("; ", roots))}");
        await new McpServer(stdin, stdout, log, tools).RunAsync();
        return 0;
    }

    /// <summary>
    /// The watchdog: one per user session (named mutex), exits 2 minutes after the last VB6 is gone.
    /// The add-in starts it with CreateProcess (not VB Shell, which blocked the IDE ~9 s, spike 12),
    /// so it inherits VB6's elevation without a UAC prompt and keeps running if VB6 crashes.
    /// </summary>
    private static async Task<int> WatchdogMainAsync()
    {
        string baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp");
        Directory.CreateDirectory(baseDir);
        var logWriter = OpenLog(Path.Combine(baseDir, "watchdog.log"));
        void Log(string m) => logWriter.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{Environment.ProcessId}] {m}");

        using var mutex = new Mutex(true, @"Local\VB6McpWatchdog", out bool first);
        if (!first) { Log("another watchdog is running; exiting"); return 0; }
        SetProcessDPIAware();                       // window rectangles and captures in physical pixels

        var actions = new WatchdogActions(Log, Path.Combine(baseDir, "shots"));
        var host = new WatchdogHost(actions, Log, Path.Combine(baseDir, "watchdog.json"));
        using var cts = new CancellationTokenSource();
        _ = Task.Run(async () =>
        {
            var lastSeen = DateTime.UtcNow;
            while (!cts.IsCancellationRequested)
            {
                await Task.Delay(5000);
                if (System.Diagnostics.Process.GetProcessesByName("VB6").Length > 0) lastSeen = DateTime.UtcNow;
                else if (DateTime.UtcNow - lastSeen > TimeSpan.FromMinutes(2)) { Log("no VB6 for 2 minutes; exiting"); cts.Cancel(); }
            }
        });
        // Recovery copies of unsaved code, every 10 s while the user is idle (hard rule 12)
        var poller = new RecoveryPoller(new Core.Recovery.RecoveryStore(), () => Core.Bridge.Discovery.Scan(),
            i => new Core.Bridge.AddinClient(i, TimeSpan.FromSeconds(2)), pid => UI.IdeWindows.ParseMode(UI.IdeWindows.Read(pid, 200).Title),
            RecoveryPoller.SystemIdle, Log);
        _ = Task.Run(async () =>
        {
            while (!cts.IsCancellationRequested)
            {
                try { await poller.PollOnceAsync(cts.Token); }
                catch (Exception ex) when (ex is not OperationCanceledException) { Log("recovery poll failed: " + ex.Message); }
                try { await Task.Delay(10_000, cts.Token); } catch (OperationCanceledException) { }
            }
        });
        Log($"watchdog {WatchdogHost.Version} started, elevated={actions.Elevated}, parent={actions.ParentPid}");
        await host.RunAsync(cts.Token);
        Log("watchdog stopped");
        GC.KeepAlive(mutex);
        return 0;
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool SetProcessDPIAware();

    /// <summary>
    /// Several server instances can run at once (one per Claude session, plus `claude mcp get` health
    /// checks), all appending to the same log. An exclusive open made a second instance crash at
    /// startup, so open shared and fall back to stderr if the file can't be opened at all.
    /// </summary>
    private static TextWriter OpenLog(string? path)
    {
        if (path is not { Length: > 0 }) return Console.Error;
        try
        {
            var fs = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
            return TextWriter.Synchronized(new StreamWriter(fs, new UTF8Encoding(false)) { AutoFlush = true });
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Console.Error.WriteLine($"vb6mcp: can't open log '{path}' ({ex.Message}); logging to stderr");
            return Console.Error;
        }
    }
}
