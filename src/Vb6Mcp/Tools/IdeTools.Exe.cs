using System.Collections.Concurrent;
using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json.Nodes;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Compiled programs (Phase 3): exe_run starts a built EXE so the app_* tools can drive it (argument
/// exe). It runs at the server's integrity, so the server reads and sends input itself
/// (<see cref="LocalInput"/>); only programs started here get input (hard rule 8).
/// </summary>
public sealed partial class IdeTools
{
    /// <summary>Programs started with exe_run: pid → path.</summary>
    private readonly ConcurrentDictionary<int, string> _exes = new();

    private IReadOnlyList<Tool> ExeTools() =>
    [
        new("exe_run",
            "Start a compiled program (.exe under the workspace roots) so the app_* tools can drive it: pass its pid as exe. " +
            "It runs non-elevated, like this server. Waits up to 15 s for its first window and returns the windows.",
            Schema.Object(("path", "string", "Full path of the .exe.", true), ("args", "string", "Command line arguments.", false)),
            (a, ct) => ExeRunAsync(a, ct)),
        new("exe_end",
            "Close a program started with exe_run: asks its windows to close (WM_CLOSE), waits 5 s; with force=true ends it if it is still running.",
            Schema.Object(("exe", "integer", "Its pid (from exe_run).", true), ("force", "boolean", "End the process if it doesn't close.", false)),
            (a, ct) => ExeEndAsync(a, ct)),
    ];

    private static bool IsAlive(int pid)
    {
        try { using var p = Process.GetProcessById(pid); return !p.HasExited; }
        catch (ArgumentException) { return false; }
        catch (InvalidOperationException) { return false; }
    }

    private async Task<ToolResult> ExeRunAsync(JsonObject a, CancellationToken ct)
    {
        var path = Path.GetFullPath(a.Req("path"));
        if (!path.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) throw new ToolException("path must be an .exe");
        if (Roots.Count == 0) throw new ToolException("no workspace roots are configured (VB6MCP_ROOTS); exe_run only starts programs under them");
        if (!Roots.Any(r => PathKey.Key(path).StartsWith(PathKey.Key(r).TrimEnd('\\') + "\\", StringComparison.Ordinal)))
            throw new ToolException($"{path} isn't under the workspace roots ({string.Join("; ", Roots)})");
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);

        var psi = new ProcessStartInfo(path) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(path)! };
        if (a.Str("args") is { Length: > 0 } args) psi.Arguments = args;
        Audit.Write("exe_run", 0, $"{path} {psi.Arguments}");
        Process proc;
        try { proc = Process.Start(psi) ?? throw new ToolException("the program didn't start"); }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 740)
        { throw new ToolException("the program requires elevation (manifest or compatibility setting); exe_run starts programs non-elevated"); }
        catch (Win32Exception ex) { throw new ToolException("the program didn't start: " + ex.Message); }
        int pid = proc.Id;
        _exes[pid] = path;

        var deadline = DateTime.UtcNow.AddSeconds(15);
        IReadOnlyList<AppWindow> windows = [];
        while (DateTime.UtcNow < deadline && IsAlive(pid) && (windows = AppWindows.List(pid)).Count == 0) await Task.Delay(250, ct);
        if (!IsAlive(pid))
        {
            _exes.TryRemove(pid, out _);
            return ToolResult.Ok(new { status = "exited", exe = pid, path, exitCode = SafeExitCode(proc) });
        }
        return ToolResult.Ok(new
        {
            status = windows.Count > 0 ? "running" : "no-window", exe = pid, path,
            windows = windows.Select(x => new { hwnd = Hex(x.Hwnd), x.Kind, x.Caption }),
            dialogs = IdeWindows.Dialogs(pid).Select(d => new { caption = d.Caption, text = string.Join("\n", d.Texts), buttons = d.Buttons.Select(b => b.Caption) }),
            note = $"drive it with the app_* tools and exe={pid}; close it with exe_end",
        });
    }

    private static int? SafeExitCode(Process p) { try { return p.HasExited ? p.ExitCode : null; } catch (InvalidOperationException) { return null; } }

    private async Task<ToolResult> ExeEndAsync(JsonObject a, CancellationToken ct)
    {
        int pid = a.Int("exe") ?? throw new ToolException("missing required argument 'exe'");
        if (!_exes.TryGetValue(pid, out var path)) throw new ToolException($"process {pid} wasn't started with exe_run");
        if (!IsAlive(pid)) { _exes.TryRemove(pid, out _); return ToolResult.Ok(new { status = "already-exited", exe = pid }); }
        Audit.Write("exe_end", pid, path);
        foreach (var w in AppWindows.List(pid).Where(x => x.Kind is "Form" or "MDIForm")) LocalInputClose(w.Hwnd);
        var deadline = DateTime.UtcNow.AddSeconds(5);
        while (DateTime.UtcNow < deadline && IsAlive(pid)) await Task.Delay(200, ct);
        if (IsAlive(pid))
        {
            if (a.Bool("force") != true)
                return ToolResult.Ok(new
                {
                    status = "still-running", exe = pid,
                    dialogs = IdeWindows.Dialogs(pid).Select(d => new { caption = d.Caption, text = string.Join("\n", d.Texts), buttons = d.Buttons.Select(b => b.Caption) }),
                    note = "it didn't close (a dialog such as 'save changes?' may be waiting); answer it, or pass force=true to end it",
                });
            using var p = Process.GetProcessById(pid);
            p.Kill();
            p.WaitForExit(3000);
        }
        _exes.TryRemove(pid, out _);
        return ToolResult.Ok(new { status = "ended", exe = pid });
    }

    /// <summary>Posts WM_CLOSE, the same as the user clicking the window's close box (QueryUnload/Unload run).</summary>
    private static void LocalInputClose(long hwnd) => PostMessage(new IntPtr(hwnd), 0x10, IntPtr.Zero, IntPtr.Zero);

    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
}
