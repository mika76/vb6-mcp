using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>
/// Crash notices (PLAN.md D2, improvement notes 25). VB6 crashed while the agent was busy elsewhere and it found out only
/// when its next call failed with "no VB6 IDE with the add-in has pid 488". Now the next result of any vb6 tool carries
/// ideCrashed: which IDE, its recovery copies, the Windows Application-log entry (faulting module, exception code) and how
/// to relaunch. A crash is a discovery file left behind by a process that is gone (spike 12b); a clean exit removes it.
/// Only IDEs this server saw running are reported, and each once.
/// </summary>
public sealed partial class IdeTools
{
    private readonly HashSet<int> _seenLive = [];
    private readonly HashSet<int> _endedOnPurpose = [];
    private readonly object _crashLock = new();

    /// <summary>Crashes since the last call, worked out before the tool runs (ide_list cleans up the discovery files).</summary>
    private List<object> CrashesSinceLastCall()
    {
        DiscoveryResult scan;
        try { scan = _env.Scan(); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return []; }
        var crashed = new List<int>();
        lock (_crashLock)
        {
            foreach (var s in scan.Stale)
                if (_seenLive.Remove(s.Pid) && !_endedOnPurpose.Remove(s.Pid)) crashed.Add(s.Pid);
            var live = scan.Live.Select(i => i.Pid).ToHashSet();
            _seenLive.RemoveWhere(p => !live.Contains(p));   // gone without a left-over file: closed normally
            _seenLive.UnionWith(live);
        }
        return crashed.Select(CrashNotice).ToList();
    }

    private ToolResult WithCrashes(List<object> crashes, ToolResult r)
        => crashes.Count == 0 ? r : r.With("ideCrashed", JsonSerializer.SerializeToNode(crashes, McpServer.Json));

    /// <summary>ide_kill and ide_quit end IDEs on purpose: not a crash.</summary>
    private void EndedOnPurpose(int pid) { lock (_crashLock) _endedOnPurpose.Add(pid); }

    private object CrashNotice(int pid)
    {
        var session = Store.Sessions().FirstOrDefault(s => s.Pid == pid && s.Note == null);
        bool watchdog = _env.FindWatchdog() is not null;
        return new
        {
            pid, recovery = RecoverySummary(pid), projects = session?.ProjectFiles is { Count: > 0 } p ? p : null,
            applicationLog = ApplicationLogEntry(pid),
            next = watchdog
                ? "relaunch with ide_launch reopen=true now: the watchdog starts VB6 without a UAC prompt (it stays up 30 minutes after a crash). Then recovery_list / recovery_restore"
                : "ide_launch reopen=true relaunches it, with a UAC prompt (no watchdog is running). Then recovery_list / recovery_restore",
        };
    }

    /// <summary>Reads the Application log (wevtutil ships with Windows; a medium process may read it). Tests replace it.</summary>
    public Func<int, object?> ApplicationLogEntry { get; init; } = pid => ParseApplicationError(RunWevtutil(), pid);

    private static string RunWevtutil()
    {
        try
        {
            var psi = new ProcessStartInfo("wevtutil.exe", "qe Application /q:\"*[System[Provider[@Name='Application Error'] and TimeCreated[timediff(@SystemTime) <= 7200000]]]\" /f:text /rd:true /c:20")
            { UseShellExecute = false, RedirectStandardOutput = true, CreateNoWindow = true };
            using var p = Process.Start(psi)!;
            var text = p.StandardOutput.ReadToEndAsync();
            return p.WaitForExit(5000) && text.Wait(1000) ? text.Result : "";
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException) { return ""; }
    }

    /// <summary>
    /// The "Application Error" event of a VB6 process, from wevtutil's text output: faulting module, exception code, offset
    /// and time. Matched by "Faulting process id: 0x&lt;pid hex&gt;"; null when there is none (e.g. it was killed, not crashed).
    /// </summary>
    public static object? ParseApplicationError(string wevtutilText, int pid)
    {
        var hex = "0x" + pid.ToString("x", CultureInfo.InvariantCulture);
        foreach (var ev in Regex.Split(wevtutilText, @"(?m)^Event\[\d+\]:?"))
        {
            if (!ev.Contains("VB6.EXE", StringComparison.OrdinalIgnoreCase)) continue;
            // Windows 11 writes "Faulting process ID" (checked on this machine); older ones "id"
            var pidLine = Regex.Match(ev, @"Faulting process id:\s*(0x[0-9a-fA-F]+)", RegexOptions.IgnoreCase);
            if (!pidLine.Success || !pidLine.Groups[1].Value.Equals(hex, StringComparison.OrdinalIgnoreCase)) continue;
            string? Field(string name) => Regex.Match(ev, $@"{Regex.Escape(name)}:\s*(.+)") is { Success: true } m ? m.Groups[1].Value.Trim() : null;
            return new
            {
                time = Field("Date"), faultingModule = Field("Faulting module name"), exceptionCode = Field("Exception code"),
                faultOffset = Field("Fault offset"), modulePath = Field("Faulting module path"),
            };
        }
        return null;
    }
}
