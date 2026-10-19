using System.Globalization;
using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Core.Recovery;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Recovery copies of unsaved code taken by the watchdog (PLAN.md §3.1). The server reads the store
/// directly (it's in the user's profile); restoring writes into a live IDE through its add-in, so
/// the components end up unsaved there and the user decides whether to save (hard rule 4).
/// </summary>
public sealed partial class IdeTools
{
    private IReadOnlyList<Tool> RecoveryTools() =>
    [
        new("recovery_list",
            "Recovery copies of unsaved IDE code, taken by the watchdog while the user is idle: per VB6 session (pid) whether it is " +
            "still running or crashed/was killed, and per component when it was copied. Use after a hang, crash or ide_kill.",
            Schema.Object(("from_pid", "integer", "Only this session.", false)), (a, _) => Task.FromResult(RecoveryList(a))),
        new("recovery_read",
            "The text of one recovery copy, and whether it differs from the file on disk.",
            Schema.Object(("from_pid", "integer", "The session (pid of the VB6 that had the unsaved code).", true),
                          ("component", "string", ComponentDesc, true), ("project", "string", "Project name, if several copies share the component name.", false)),
            (a, _) => Task.FromResult(RecoveryRead(a))),
        new("recovery_restore",
            "Put a recovery copy back into a running IDE: replaces the code of the component with the same file in an open project " +
            "(design mode). The component is left unsaved; save with project_save once the user agrees.",
            Schema.Object(("from_pid", "integer", "The session the copy came from.", true), ("component", "string", ComponentDesc, true),
                          ("project", "string", "Project name of the copy, if needed.", false), ("pid", "integer", "Target IDE. " + PidDesc, false)),
            (a, _) => RecoveryRestoreAsync(a)),
        new("recovery_discard",
            "Delete the recovery copies of one session (e.g. after restoring them). Needs confirm=true.",
            Schema.Object(("from_pid", "integer", "The session.", true), ("confirm", "boolean", "Required: true.", true)),
            (a, _) => Task.FromResult(RecoveryDiscard(a))),
    ];

    private RecoveryStore Store => _env.Recovery;

    /// <summary>For ide_kill and recovery_list: the copies of one session with their age.</summary>
    private object RecoverySummary(int pid)
    {
        var copies = Store.Copies(pid);
        return new
        {
            copies = copies.Select(c => new { project = c.Project, component = c.Component, file = c.File, lines = c.Lines, taken = c.Taken.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), ageMinutes = (int)(DateTime.Now - c.Taken).TotalMinutes }),
            note = copies.Count == 0 ? "no recovery copies: nothing was unsaved at the last copy, or the watchdog wasn't running" : null,
        };
    }

    private ToolResult RecoveryList(JsonObject a)
    {
        int? only = a.Int("from_pid");
        var sessions = Store.Sessions().Where(s => only is null || s.Pid == only).Select(s => new
        {
            pid = s.Pid, running = s.Note == null && Discovery.IsVb6Process(s.Pid), crashed = s.Note == null && s.Crashed, note = s.Note,
            updated = s.Updated.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), projects = s.ProjectFiles,
            recovery = RecoverySummary(s.Pid),
        }).ToList();
        return ToolResult.Ok(new
        {
            sessions,
            note = sessions.Count == 0 ? "no recovery copies (the watchdog copies unsaved components while you are idle; a cleanly closed IDE's copies are removed)" : null,
        });
    }

    private RecoveryCopy FindCopy(JsonObject a, out int fromPid)
    {
        fromPid = a.Int("from_pid") ?? throw new ToolException("missing required argument 'from_pid'");
        var component = a.Req("component");
        var project = a.Str("project");
        var matches = Store.Copies(fromPid).Where(c => c.Component.Equals(component, StringComparison.OrdinalIgnoreCase)
                                                    && (project is null || c.Project.Equals(project, StringComparison.OrdinalIgnoreCase))).ToList();
        return matches.Count switch
        {
            0 => throw new ToolException($"no recovery copy of {component} in session {fromPid} (see recovery_list)"),
            1 => matches[0],
            _ => throw new ToolException($"{matches.Count} copies named {component}; pass project ({string.Join(", ", matches.Select(m => m.Project))})"),
        };
    }

    private ToolResult RecoveryRead(JsonObject a)
    {
        var copy = FindCopy(a, out int fromPid);
        var text = Store.ReadText(fromPid, copy.Project, copy.Component) ?? throw new ToolException("the copy's text file is missing");
        object disk;
        try
        {
            var f = VbSourceFile.Load(copy.File);
            disk = new { exists = true, sameAsDisk = f.Code == text, diskLines = f.CountOfLines };
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { disk = new { exists = false, error = ex.Message }; }
        return ToolResult.Ok(new
        {
            project = copy.Project, component = copy.Component, file = copy.File, lines = copy.Lines,
            taken = copy.Taken.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), disk, text,
        });
    }

    private async Task<ToolResult> RecoveryRestoreAsync(JsonObject a)
    {
        var copy = FindCopy(a, out int fromPid);
        var text = Store.ReadText(fromPid, copy.Project, copy.Component) ?? throw new ToolException("the copy's text file is missing");
        var inst = ResolveInstance(a);
        if (inst.Pid == fromPid) throw new ToolException("the copy comes from this same IDE: it still has the unsaved code (nothing to restore)");
        var w = _env.ReadWindows(inst.Pid);
        if (w.Mode != IdeMode.Design) throw new ToolException($"the target IDE is in {Mode(w.Mode)} mode; restoring edits code, which needs design mode");
        var client = _env.Client(inst);

        // The target component is identified by its file (hard rule 3), not by name
        (string Project, ComponentInfo Comp)? target = null;
        foreach (var p in await client.ProjectsAsync())
            foreach (var c in await client.ComponentsAsync(p.Name))
                if (c.FileName.Length > 0 && copy.File.Length > 0 && PathKey.Same(c.FileName, copy.File)) target = (p.Name, c);
        if (target is null) throw new ToolException($"no open project in IDE {inst.Pid} has {copy.File}; open {copy.ProjectFile} first (project_open)");

        var current = await client.ReadAsync(target.Value.Project, target.Value.Comp.Name);
        if (current.Text == text) return ToolResult.Ok(new { status = "already-identical", project = target.Value.Project, component = target.Value.Comp.Name });
        string prepared;
        try { prepared = VbCode.PrepareForInsert(text); }
        catch (ArgumentException ex) { throw new ToolException("the copy can't be inserted: " + ex.Message); }
        Audit.Write("recovery_restore", inst.Pid, $"{copy.Project}.{copy.Component} from session {fromPid} ({copy.Taken:yyyy-MM-dd HH:mm:ss})");
        var (lines, dirty) = await client.ReplaceAsync(target.Value.Project, target.Value.Comp.Name, 1, current.CountOfLines, current.Text, prepared);
        try { Ledger.Record(inst.Pid, target.Value.Project, target.Value.Comp.Name, (await client.ReadAsync(target.Value.Project, target.Value.Comp.Name)).Text); }
        catch (Exception ex) when (ex is AddinErrorException or AddinBusyException) { }
        return ToolResult.Ok(new
        {
            status = "restored", project = target.Value.Project, component = target.Value.Comp.Name, lines, dirty,
            from = new { pid = fromPid, taken = copy.Taken.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) },
            note = "unsaved in the IDE: review, then project_save; Edit > Undo in VB6 reverts it",
        });
    }

    private ToolResult RecoveryDiscard(JsonObject a)
    {
        int fromPid = a.Int("from_pid") ?? throw new ToolException("missing required argument 'from_pid'");
        if (a.Bool("confirm") != true) throw new ToolException("recovery_discard needs confirm=true: the copies are deleted for good");
        if (fromPid < RecoveryStore.KeptId(0) && Discovery.IsVb6Process(fromPid)) throw new ToolException($"VB6 {fromPid} is still running; its copies are kept up to date automatically");
        int n = Store.Copies(fromPid).Count;
        Store.DeleteSession(fromPid);
        return ToolResult.Ok(new { pid = fromPid, deleted = n });
    }
}
