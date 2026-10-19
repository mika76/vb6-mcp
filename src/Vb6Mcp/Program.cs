// vb6mcp: MCP server for the VB6 IDE. See docs/architecture.md and docs/tools.md.
// Usage: vb6mcp [mcp]       MCP server on stdio (default; what Claude Code starts)
//        vb6mcp watchdog    elevated sidecar, started by the VB6 add-in (PLAN.md §3.1)
//        vb6mcp watchdog --stay   the same, started by the user (elevated) and kept running with no VB6, so sessions
//                                 can start their own IDEs through it (ide_launch) without a UAC prompt each
//        vb6mcp hook        Claude Code PreToolUse hook: blocks Edit/Write on VB6 source files
//        vb6mcp install-hook [settings.json] [--no-idle-timeout]  registers this exe as that hook and sets Claude Code's idle timeout (install scripts)
//        vb6mcp uninstall-hook [settings.json] removes that hook again (used by uninstall.ps1)
//        vb6mcp describe <file> [filter]       prints a typelib (DLL/OCX/TLB/OLB) as VB-like declarations
//        vb6mcp windows <pid>                   prints the IDE's window tree (diagnosis)
// Diagnostics go to stderr, or to the file named by VB6MCP_LOG (some clients hide stderr).
using System.Text;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.Watchdog;

namespace Vb6Mcp;

public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        string mode = args.Length > 0 ? args[0].ToLowerInvariant() : "mcp";
        if (mode == "watchdog") return await WatchdogMainAsync(args.Skip(1).Any(x => x.Equals("--stay", StringComparison.OrdinalIgnoreCase)));
        if (mode == "hook") return Hook.Run(Console.In, Console.Error);
        if (mode == "install-hook")
            return Hook.Install(args.Skip(1).FirstOrDefault(x => !x.StartsWith("--", StringComparison.Ordinal)) ?? Hook.UserSettingsPath, Environment.ProcessPath!, Console.Out,
                               setIdleTimeout: !args.Contains("--no-idle-timeout"));
        if (mode == "uninstall-hook") return Hook.Uninstall(args.Length > 1 ? args[1] : Hook.UserSettingsPath, Console.Out);
        if (mode == "windows" && args.Length > 1 && int.TryParse(args[1], out var wpid))
        {
            // diagnosis: the IDE's window tree (classes, captions, rectangles), e.g. to see how a form designer is built
            foreach (var w in UI.IdeWindows.Descendants(wpid))
                Console.WriteLine($"{w.Hwnd:X8} parent {w.Parent:X8} {(w.Visible ? "visible" : "hidden ")} {w.Left},{w.Top} {w.Width}x{w.Height}  {w.Class}  \"{w.Caption}\"");
            return 0;
        }
        if (mode == "describe" && args.Length > 1)
        {
            Console.Out.Write(Core.Files.TypeLibDescriber.Render(Core.Files.TypeLibDescriber.Describe(args[1]), args.Length > 2 ? args[2] : null));
            return 0;
        }
        if (mode != "mcp")
        {
            Console.Error.WriteLine($"vb6mcp: unknown mode '{mode}'. Usage: vb6mcp [mcp|watchdog|hook|install-hook|uninstall-hook|describe <file>|windows <pid>]");
            return 2;
        }

        var log = OpenLog(Environment.GetEnvironmentVariable("VB6MCP_LOG"));
        var inheritedRuns = new List<int>();
        try
        {
            var taken = IdeOwnership.TakeOver(IdeOwnership.SessionKey(), runs: inheritedRuns);
            // a run counts only while it is still going: one the user ended and started again isn't the agent's
            inheritedRuns.RemoveAll(pid => UI.IdeWindows.Read(pid).Mode is not (UI.IdeMode.Run or UI.IdeMode.Break));
            log.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} start pid {Environment.ProcessId}: session {(IdeOwnership.SessionKey() is null ? "unknown (no handover)" : "set")}" +
                          (taken.Count > 0 ? $"; took over IDEs [{string.Join(", ", taken)}] and runs [{string.Join(", ", inheritedRuns)}] from the previous server" : ""));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { log.WriteLine($"handover failed: {ex.Message}"); }

        var stdin = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
        var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { NewLine = "\n" };
        // Workspace roots for the disk tools (hard rule 8: paths are checked against them)
        var roots = Roots();
        var env = new IdeEnvironment();
        var ledger = new ToolEditLedger();   // one for both: file_replace_lines edits the IDE too
        var workspace = new WorkspaceTools(env, roots) { Ledger = ledger };
        var ideTools = new IdeTools(env) { Roots = roots, MapProvider = workspace.CurrentMap, Ledger = ledger, InheritedRuns = inheritedRuns,
            EvalOpen = string.Equals(Environment.GetEnvironmentVariable("VB6MCP_EVAL"), "open", StringComparison.OrdinalIgnoreCase) };
        var tools = ideTools.All().Concat(workspace.All()).ToList();
        tools.Add(BatchTool.Create(tools));   // several calls in one (each with its own checks)
        // PLAN.md E7 (spike): the desktop app's MCP dialog offers Reconnect only for a failed server, so picking up a
        // reinstalled build meant restarting the Claude session. Ending the process after answering should make the
        // server show as failed, with Reconnect. Nothing in the IDE is touched, so it isn't a write tool.
        tools.Add(new Tool("server_restart",
            "End this vb6 MCP server process (after answering), e.g. after scripts\\install-server.ps1 installed a newer build. " +
            "The vb6 tools are then unavailable until the user reconnects the server (the MCP servers dialog, /mcp) or restarts the " +
            "session. IDEs, the watchdog and unsaved code are not affected. Only when the user asks for it. Wait for its answer before " +
            "the next vb6 call: calls that arrive while it ends are refused with 'restarting'.",
            Schema.Object(("confirm", "boolean", "Required: true.", true)),
            (a, _) =>
            {
                if (a.Bool("confirm") != true) throw new ToolException("ending the server needs confirm=true");
                // the next server of this session takes over this one's IDEs (improvement notes 53)
                IReadOnlyList<int> handed;
                try { handed = IdeOwnership.WriteHandover(IdeOwnership.SessionKey(), runs: ideTools.RunsStarted); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { handed = []; }
                McpServer.Restarting = true;   // new calls get "restarting, call again" right away (spike 42)
                log.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} server_restart: exiting pid {Environment.ProcessId}; handing over IDEs [{string.Join(", ", handed)}] (session {(IdeOwnership.SessionKey() is null ? "unknown" : "set")})");
                // after the reply is written, and after any call that is still running (at most 2 minutes)
                Task.Run(async () =>
                {
                    await Task.Delay(1000);
                    for (var until = DateTime.UtcNow.AddMinutes(2); McpServer.InFlight > 0 && DateTime.UtcNow < until;) await Task.Delay(200);
                    log.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} server_restart: exiting pid {Environment.ProcessId} now");
                    Environment.Exit(3);
                });
                return Task.FromResult(ToolResult.Ok(new { exiting = Environment.ProcessId, version = McpServer.ServerVersion, exe = Environment.ProcessPath,
                    handedOver = handed.Count > 0 ? handed : null,
                    note = "the server ends in a second; reconnect vb6 in the MCP servers dialog (/mcp) or restart the session" }));
            }));
        log.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} vb6mcp {McpServer.ServerVersion} started, pid {Environment.ProcessId}, {tools.Count} tools, roots: {(roots.Length == 0 ? "(none)" : string.Join("; ", roots))}");
        await new McpServer(stdin, stdout, log, tools).RunAsync();
        return 0;
    }

    /// <summary>
    /// The watchdog: one per user session (named mutex), exits 2 minutes after the last VB6 is gone.
    /// The add-in starts it with CreateProcess (not VB Shell, which blocked the IDE ~9 s, spike 12),
    /// so it inherits VB6's elevation without a UAC prompt and keeps running if VB6 crashes.
    /// </summary>
    private static async Task<int> WatchdogMainAsync(bool stay)
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
            DateTime? crashSeen = null;
            while (!cts.IsCancellationRequested)
            {
                await Task.Delay(5000);
                // A discovery file left behind by a gone process is a crash (spike 12b). Then stay 30 minutes instead of 2, so the
                // agent can relaunch through the watchdog without a UAC prompt once it notices (improvement notes 25, PLAN D2).
                try { if (Core.Bridge.Discovery.Scan().Stale.Count > 0) crashSeen ??= DateTime.UtcNow; } catch (IOException) { }
                var linger = crashSeen is { } c && DateTime.UtcNow - c < TimeSpan.FromMinutes(30) ? TimeSpan.FromMinutes(30) : TimeSpan.FromMinutes(2);
                if (System.Diagnostics.Process.GetProcessesByName("VB6").Length > 0) { lastSeen = DateTime.UtcNow; crashSeen = null; }
                // Started by the add-in it leaves 2 minutes after the last VB6. Started by the user with --stay it keeps running,
                // so a session can launch its own IDE through it (ide_launch, elevated without a UAC prompt).
                else if (!stay && DateTime.UtcNow - lastSeen > linger) { Log($"no VB6 for {linger.TotalMinutes:0} minutes; exiting"); cts.Cancel(); }
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
        // Sync engine: outside changes to files of loaded projects, every 2 s (PLAN §8)
        var sync = new SyncPoller(() => Core.Bridge.Discovery.Scan(), i => new Core.Bridge.AddinClient(i, TimeSpan.FromSeconds(2)),
            pid => UI.IdeWindows.ParseMode(UI.IdeWindows.Read(pid, 200).Title), Log, Core.SyncState.DefaultPath);
        _ = Task.Run(async () =>
        {
            while (!cts.IsCancellationRequested)
            {
                try { await sync.PollOnceAsync(cts.Token); }
                catch (Exception ex) when (ex is not OperationCanceledException) { Log("sync poll failed: " + ex.Message); }
                try { await Task.Delay(2_000, cts.Token); } catch (OperationCanceledException) { }
            }
        });
        Log($"watchdog {WatchdogHost.Version} started, elevated={actions.Elevated}, parent={actions.ParentPid}, exe={Environment.ProcessPath}{(stay ? ", --stay" : "")}");   // an elevated process's path is hidden from medium ones: log it (spike 26)
        await host.RunAsync(cts.Token);
        Log("watchdog stopped");
        GC.KeepAlive(mutex);
        return 0;
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool SetProcessDPIAware();

    /// <summary>
    /// Workspace roots: VB6MCP_ROOTS (semicolon-separated; development, install-server.ps1), else roots.txt next to this
    /// exe, one folder per line. The installer writes that file into Program Files, where only administrators can change
    /// it, and the add-in reads the same file next to its DLL (Phase 6): one admin-controlled list for both sides.
    /// </summary>
    public static string[] Roots()
    {
        var env = Environment.GetEnvironmentVariable("VB6MCP_ROOTS");
        if (!string.IsNullOrWhiteSpace(env)) return env.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var file = Path.Combine(AppContext.BaseDirectory, "roots.txt");
        if (!File.Exists(file)) return [];
        return File.ReadAllLines(file, Core.Vb1252.Encoding).Select(l => l.Trim()).Where(l => l.Length > 0).ToArray();
    }

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
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);   // %LOCALAPPDATA%\vb6mcp may not exist yet on a fresh install
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
