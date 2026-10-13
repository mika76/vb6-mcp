using System.Runtime.InteropServices;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Recovery;
using Vb6Mcp.UI;

namespace Vb6Mcp.Watchdog;

/// <summary>
/// Takes recovery copies of unsaved IDE code (PLAN.md §3.1). The watchdog polls the add-ins (it's a
/// client like the MCP server) instead of the add-in pushing heartbeats: that keeps the add-in thin.
/// Hard rule 12: copies are taken only while the user is idle and the IDE isn't running the program,
/// so a read on the IDE thread never interrupts typing or a run.
/// </summary>
public sealed class RecoveryPoller(
    RecoveryStore store,
    Func<DiscoveryResult> scan,
    Func<InstanceInfo, AddinClient> client,
    Func<int, IdeMode> mode,
    Func<TimeSpan> userIdle,
    Action<string> log)
{
    /// <summary>How long the user must have been idle (no keyboard or mouse input) before reading.</summary>
    public TimeSpan IdleNeeded { get; init; } = TimeSpan.FromSeconds(5);

    /// <summary>Copies of crashed sessions are kept this long.</summary>
    public TimeSpan KeepCrashed { get; init; } = TimeSpan.FromDays(7);

    /// <summary>One round: copy what changed in live IDEs, mark crashed ones, remove cleanly closed ones.</summary>
    public async Task PollOnceAsync(CancellationToken ct)
    {
        var s = scan();
        var livePids = s.Live.Select(i => i.Pid).ToHashSet();
        var stalePids = s.Stale.Select(i => i.Pid).ToHashSet();

        if (userIdle() >= IdleNeeded)
            foreach (var inst in s.Live)
            {
                if (ct.IsCancellationRequested) return;
                var m = mode(inst.Pid);
                if (m is not (IdeMode.Design or IdeMode.Break)) continue;     // never while the program runs
                try { await CopyAsync(inst); }
                catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException or Core.Protocol.ProtocolException or IOException)
                { log($"recovery: pid {inst.Pid} skipped: {ex.Message}"); }
            }

        foreach (var session in store.Sessions())
        {
            if (livePids.Contains(session.Pid)) continue;
            if (stalePids.Contains(session.Pid))
            {
                // discovery file left behind, process gone: crashed or killed (spike 12b). Keep the copies.
                if (!session.Crashed) { store.WriteSession(session with { Crashed = true }); log($"recovery: pid {session.Pid} crashed; {store.Copies(session.Pid).Count} copies kept"); }
                else if (DateTime.Now - session.Updated > KeepCrashed) { store.DeleteSession(session.Pid); log($"recovery: pid {session.Pid} pruned (older than {KeepCrashed.TotalDays:0} days)"); }
            }
            else if (!session.Crashed && !Discovery.IsVb6Process(session.Pid))
            {
                // closed cleanly: the add-in deleted its discovery file on OnDisconnection
                store.DeleteSession(session.Pid);
                log($"recovery: pid {session.Pid} closed cleanly; copies removed");
            }
        }
    }

    private async Task CopyAsync(InstanceInfo inst)
    {
        var c = client(inst);
        var projects = await c.ProjectsAsync();
        var keep = new HashSet<(string, string)>();
        int written = 0;
        foreach (var p in projects)
        {
            foreach (var comp in await c.ComponentsAsync(p.Name))
            {
                if (!comp.IsDirty) continue;
                keep.Add((p.Name, comp.Name));
                var code = await c.ReadAsync(p.Name, comp.Name);
                var hash = VbCode.Hash(code.Text);
                if (store.Copies(inst.Pid).FirstOrDefault(x => x.Project == p.Name && x.Component == comp.Name) is { } old && old.Hash == hash) continue;
                store.WriteCopy(inst.Pid, new RecoveryCopy(p.Name, p.FileName, comp.Name, comp.Type, comp.FileName, code.CountOfLines, hash, DateTime.Now), code.Text);
                written++;
            }
        }
        // a component that is no longer dirty was saved (or reverted): its copy isn't needed
        foreach (var old in store.Copies(inst.Pid).Where(x => !keep.Contains((x.Project, x.Component))))
            store.DeleteCopy(inst.Pid, old.Project, old.Component);
        store.WriteSession(new RecoverySession(inst.Pid, inst.Started, DateTime.Now, false, projects.Select(p => p.FileName).ToList()));
        if (written > 0) log($"recovery: pid {inst.Pid}: {written} copies updated, {keep.Count} unsaved components");
    }

    /// <summary>Time since the last keyboard or mouse input in this session (GetLastInputInfo).</summary>
    public static TimeSpan SystemIdle()
    {
        var info = new LASTINPUTINFO { cbSize = (uint)Marshal.SizeOf<LASTINPUTINFO>() };
        return GetLastInputInfo(ref info) ? TimeSpan.FromMilliseconds(unchecked((uint)Environment.TickCount - info.dwTime)) : TimeSpan.Zero;
    }

    [StructLayout(LayoutKind.Sequential)] private struct LASTINPUTINFO { public uint cbSize, dwTime; }
    [DllImport("user32.dll")] private static extern bool GetLastInputInfo(ref LASTINPUTINFO info);
}
