using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.UI;

namespace Vb6Mcp.Watchdog;

/// <summary>
/// The sync engine (PLAN.md §8): notices files of loaded projects that changed outside the IDE (git,
/// another editor, another IDE sharing the file) and applies the owner's policy: reload automatically
/// when it is safe, never lose unsaved IDE work.
/// <list type="bullet">
/// <item>Clean component, design mode: reloaded (add-in RELOAD "clean", which checks again just before).</item>
/// <item>Run or break mode: waits, reloads once the IDE is back in design mode.</item>
/// <item>Component with unsaved changes: a conflict; nothing is touched (sync_diff / sync_take_*).</item>
/// <item>.vbp changed, file deleted: reported only.</item>
/// </list>
/// The IDE's own saves are recognised through the add-in's record of AfterWriteFile (WRITES). A change
/// has to stay the same for one more poll before it counts, so half-written files aren't reloaded.
/// It lives in the watchdog because outside changes happen whether or not an agent session runs.
/// </summary>
public sealed class SyncPoller(
    Func<DiscoveryResult> scan,
    Func<InstanceInfo, AddinClient> client,
    Func<int, IdeMode> mode,
    Action<string> log,
    string statePath)
{
    private sealed record Stat(bool Exists, long Ticks, long Size);
    private sealed record Watched(string Path, string Project, string? Component, bool Dirty);

    private readonly Dictionary<int, long> _seq = [];
    private readonly Dictionary<(int Pid, string Key), Stat> _known = [];
    private readonly Dictionary<(int Pid, string Key), Stat> _pending = [];
    private readonly Dictionary<(int Pid, string Key), Watched> _awaitDesign = [];
    private readonly List<SyncConflict> _conflicts = [];
    private readonly List<SyncEvent> _events = [];
    private bool _changed;

    /// <summary>Events kept in the state file.</summary>
    public int KeepEvents { get; init; } = 100;

    private static Stat StatOf(string path)
    {
        var fi = new FileInfo(path);
        return fi.Exists ? new Stat(true, fi.LastWriteTimeUtc.Ticks, fi.Length) : new Stat(false, 0, 0);
    }

    private void Event(int pid, string kind, string file, string? project, string? component, string? note = null)
    {
        _events.Add(new SyncEvent(DateTime.Now, pid, kind, file, project, component, note));
        if (_events.Count > KeepEvents) _events.RemoveRange(0, _events.Count - KeepEvents);
        _changed = true;
        log($"sync: pid {pid} {kind} {project}.{component} {file}{(note is null ? "" : " (" + note + ")")}");
    }

    /// <summary>One round over all live IDEs.</summary>
    public async Task PollOnceAsync(CancellationToken ct)
    {
        var live = scan().Live;
        var livePids = live.Select(i => i.Pid).ToHashSet();
        foreach (var inst in live)
        {
            if (ct.IsCancellationRequested) return;
            try { await PollIdeAsync(inst); }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException or Core.Protocol.ProtocolException or IOException)
            { /* busy (program running a loop) or gone: try again next round */ }
        }
        // forget IDEs that are gone
        if (_conflicts.RemoveAll(c => !livePids.Contains(c.Pid)) > 0) _changed = true;
        foreach (var k in _known.Keys.Where(k => !livePids.Contains(k.Pid)).ToList()) _known.Remove(k);
        foreach (var pid in _seq.Keys.Where(p => !livePids.Contains(p)).ToList()) _seq.Remove(pid);
        if (_changed) { SyncState.Write(new SyncSnapshot(_conflicts.ToList(), _events.ToList(), DateTime.Now), statePath); _changed = false; }
    }

    private async Task PollIdeAsync(InstanceInfo inst)
    {
        var c = client(inst);
        int pid = inst.Pid;
        var (next, ownPaths) = await c.WritesAsync(_seq.GetValueOrDefault(pid));
        _seq[pid] = next;
        var own = ownPaths.Select(p => PathKey.Key(p)).ToHashSet();

        var watched = new List<Watched>();
        foreach (var p in await c.ProjectsAsync())
        {
            if (p.FileName.Length > 0) watched.Add(new Watched(p.FileName, p.Name, null, p.IsDirty));
            foreach (var comp in await c.ComponentsAsync(p.Name))
                if (comp.FileName.Length > 0) watched.Add(new Watched(comp.FileName, p.Name, comp.Name, comp.IsDirty));
        }
        var ideMode = mode(pid);

        // conflicts end when the component has no unsaved changes any more (saved over disk, or reloaded)
        foreach (var conflict in _conflicts.Where(x => x.Pid == pid).ToList())
        {
            var w = watched.FirstOrDefault(x => x.Component == conflict.Component && x.Project == conflict.Project);
            if (w is null || !w.Dirty)
            {
                _conflicts.Remove(conflict);
                Event(pid, "resolved", conflict.File, conflict.Project, conflict.Component,
                      w is null ? "component gone" : own.Contains(PathKey.Key(conflict.File)) ? "the IDE saved its version over the disk file" : "no unsaved changes any more");
            }
        }

        foreach (var w in watched)
        {
            var key = (pid, PathKey.Key(w.Path));
            var stat = StatOf(w.Path);
            if (!_known.TryGetValue(key, out var prev)) { _known[key] = stat; continue; }      // first sight
            if (own.Contains(key.Item2)) { _known[key] = stat; _pending.Remove(key); _awaitDesign.Remove(key); continue; }   // the IDE's own save
            if (stat == prev) { _pending.Remove(key); continue; }
            if (!stat.Exists) { _known[key] = stat; Event(pid, "deleted", w.Path, w.Project, w.Component); continue; }
            // changed outside the IDE: act only when it stayed the same for one more round
            if (!_pending.TryGetValue(key, out var seen) || seen != stat) { _pending[key] = stat; continue; }
            _pending.Remove(key);
            _known[key] = stat;

            if (w.Component is null) { Event(pid, "project-changed", w.Path, w.Project, null, "reload the project yourself if needed (it is disruptive)"); continue; }
            if (w.Dirty)
            {
                if (!_conflicts.Any(x => x.Pid == pid && x.Component == w.Component && x.Project == w.Project))
                {
                    _conflicts.Add(new SyncConflict(pid, w.Project, w.Component, w.Path, DateTime.Now));
                    Event(pid, "conflict", w.Path, w.Project, w.Component, "changed on disk while it has unsaved changes in the IDE; nothing was touched");
                }
                continue;
            }
            if (ideMode != IdeMode.Design) { _awaitDesign[key] = w; Event(pid, "pending", w.Path, w.Project, w.Component, "will reload in design mode"); continue; }
            await ReloadAsync(c, pid, w);
        }

        // reloads that waited for design mode
        if (ideMode == IdeMode.Design)
            foreach (var (key, w) in _awaitDesign.Where(x => x.Key.Pid == pid).ToList())
            {
                _awaitDesign.Remove(key);
                var now = watched.FirstOrDefault(x => x.Component == w.Component && x.Project == w.Project);
                if (now is null) continue;
                if (now.Dirty)
                {
                    _conflicts.Add(new SyncConflict(pid, w.Project, w.Component!, w.Path, DateTime.Now));
                    Event(pid, "conflict", w.Path, w.Project, w.Component, "edited in the IDE while the reload waited");
                }
                else await ReloadAsync(c, pid, now);
            }
    }

    private async Task ReloadAsync(AddinClient c, int pid, Watched w)
    {
        try
        {
            await c.ReloadAsync(w.Project, w.Component!, "clean");
            Event(pid, "reloaded", w.Path, w.Project, w.Component);
        }
        catch (AddinErrorException ex) { Event(pid, "not-reloaded", w.Path, w.Project, w.Component, ex.Message); }
    }
}
