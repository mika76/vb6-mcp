using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>
/// project_checkpoint / project_rollback (PLAN.md E2). A checkpoint is the project's files as saved on disk; unsaved IDE
/// changes can't be copied (VBIDE can only save them, spike 31), so an open project must be saved first. A rollback of an
/// open project closes it, writes the files back and opens it again, all with an explicit confirm (hard rule 9).
/// </summary>
public sealed partial class IdeTools
{
    /// <summary>Where checkpoints are kept (tests use a temporary folder).</summary>
    public string? CheckpointFolder { get; init; }

    private Checkpoints CheckpointStore => new(CheckpointFolder);

    private IReadOnlyList<Tool> CheckpointTools() =>
    [
        new("project_checkpoint",
            "Save a copy of a project's files (the .vbp, every component, .frx and resource files) as they are on disk, to go back " +
            "to with project_rollback if a change goes wrong. Take one before a large or risky change. A project open in the IDE " +
            "must have no unsaved changes (project_save first), since only saved files can be copied. list=true lists the " +
            "project's checkpoints and how many files changed since each. The last 20 per project are kept.",
            Schema.Object(("label", "string", "A short name for the checkpoint, e.g. 'before redesign'.", false),
                          ("list", "boolean", "List checkpoints instead of making one.", false),
                          ("path", "string", "A project's .vbp (instead of project, e.g. for a closed project).", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => CheckpointAsync(a)),
        new("project_rollback",
            "Put a project's files back as they were at a checkpoint (project_checkpoint). checkpoint is its id, its label or 'latest'. " +
            "Needs confirm=true. A project open in the IDE is closed, restored and opened again (design mode); if it has unsaved " +
            "changes they are lost, so that also needs discard_unsaved=true. Files added since the checkpoint are left on disk and " +
            "listed (the restored .vbp no longer names them).",
            Schema.Object(("checkpoint", "string", "Checkpoint id, label or 'latest'.", true),
                          ("confirm", "boolean", "Required: true.", true),
                          ("discard_unsaved", "boolean", "Allow losing unsaved IDE changes of the open project.", false),
                          ("path", "string", "A project's .vbp (instead of project).", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => RollbackAsync(a, ct)),
    ];

    /// <summary>The .vbp named by path, or the open project's.</summary>
    private async Task<string> CheckpointVbpAsync(JsonObject a)
    {
        if (a.Str("path") is { Length: > 0 } p)
        {
            var vbp = PathKey.Normalize(p);
            if (!vbp.EndsWith(".vbp", StringComparison.OrdinalIgnoreCase)) throw new ToolException("path must be a .vbp file");
            if (!File.Exists(vbp)) throw new ToolException("file not found: " + vbp);
            return vbp;
        }
        var (client, project) = await ResolveProjectAsync(a);
        var prj = (await client.ProjectsAsync()).FirstOrDefault(x => x.Name.Equals(project, StringComparison.OrdinalIgnoreCase))
            ?? throw new ToolException($"no project {project} in VB6 {client.Instance.Pid}");
        if (prj.FileName.Length == 0) throw new ToolException($"{project} has never been saved, so it has no files to copy; save it first (project_save)");
        return PathKey.Normalize(prj.FileName);
    }

    /// <summary>Where the project is open (if anywhere), and which IDEs have any of the given files loaded in another project.</summary>
    private async Task<(InstanceInfo? Inst, ProjectInfo? Project, List<string> Elsewhere)> WhereLoadedAsync(string vbp, IReadOnlyCollection<string> files)
    {
        InstanceInfo? at = null; ProjectInfo? open = null;
        var elsewhere = new List<string>();
        var set = new HashSet<string>(files, StringComparer.OrdinalIgnoreCase);
        foreach (var inst in _env.Scan().Live)
        {
            try
            {
                var client = _env.Client(inst);
                foreach (var p in await client.ProjectsAsync())
                {
                    if (SamePath(p.FileName, vbp)) { at = inst; open = p; continue; }
                    foreach (var c in await client.ComponentsAsync(p.Name))
                        if (c.FileName.Length > 0 && set.Contains(PathKey.Normalize(c.FileName))) elsewhere.Add($"{c.FileName} (VB6 {inst.Pid}, {p.Name})");
                }
            }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException)
            {
                throw new ToolException($"VB6 {inst.Pid} didn't answer, so it can't be ruled out that it has these files loaded; try again when it responds");
            }
        }
        return (at, open, elsewhere);
    }

    private async Task<ToolResult> CheckpointAsync(JsonObject a)
    {
        var vbp = await CheckpointVbpAsync(a);
        var store = CheckpointStore;
        if (a.Bool("list") == true)
            return ToolResult.Ok(new
            {
                vbp, checkpoints = store.List(vbp).Select(c => new { id = c.Id, label = c.Label, created = c.Created, files = c.Files.Count(f => f.Copy is not null), changedSince = Checkpoints.ChangedSince(c).Count }),
            });

        var (inst, open, _) = await WhereLoadedAsync(vbp, []);
        if (inst is not null && open is not null)
        {
            var dirty = (await _env.Client(inst).ComponentsAsync(open.Name)).Where(c => c.IsDirty).Select(c => c.Name).ToList();
            if (dirty.Count > 0 || open.IsDirty)
                throw new ToolException($"{open.Name} has unsaved changes in VB6 {inst.Pid} ({(dirty.Count > 0 ? string.Join(", ", dirty) : "project settings")}): a checkpoint copies the files on disk, " +
                    "which don't have them. Save first (project_save), then take the checkpoint.");
        }
        var cp = store.Create(vbp, a.Str("label"));
        var missing = cp.Files.Where(f => f.Copy is null).Select(f => f.Path).ToList();
        return ToolResult.Ok(new
        {
            vbp, id = cp.Id, label = cp.Label, files = cp.Files.Count - missing.Count, missing = missing.Count > 0 ? missing : null,
            note = "go back with project_rollback checkpoint=" + (cp.Label is null ? cp.Id : $"'{cp.Label}'") + " confirm=true",
        });
    }

    private async Task<ToolResult> RollbackAsync(JsonObject a, CancellationToken ct)
    {
        if (a.Bool("confirm") != true) throw new ToolException("rolling back overwrites the project's files: needs confirm=true");
        var vbp = await CheckpointVbpAsync(a);
        var store = CheckpointStore;
        var which = a.Req("checkpoint");
        var cp = store.Find(vbp, which) ?? throw new ToolException($"no checkpoint '{which}' for {vbp}" +
            (store.List(vbp) is { Count: > 0 } all ? $"; there are: {string.Join(", ", all.Take(10).Select(c => c.Id + (c.Label is null ? "" : $" ({c.Label})")))}" : "; take one with project_checkpoint"));
        // every file it writes must be under the roots (hard rule 8)
        if (Roots.Count == 0) throw new ToolException($"{vbp} is outside the workspace roots (none are configured)");
        if (cp.Files.FirstOrDefault(f => !ProjectMap.IsUnderRoots(f.Path, Roots)) is { } outside)
            throw new ToolException($"{outside.Path} is outside the workspace roots");

        var (inst, open, elsewhere) = await WhereLoadedAsync(vbp, cp.Files.Select(f => f.Path).ToList());
        // hard rule 4: never write under a file another IDE (or another project) has loaded
        if (elsewhere.Count > 0)
            throw new ToolException($"files of this checkpoint are loaded elsewhere: {string.Join("; ", elsewhere.Take(5))}. Close them there first.");

        if (inst is null || open is null)
        {
            Audit.Write("project_rollback", 0, $"{vbp} to {cp.Id}");
            return ToolResult.Ok(RollbackJson(vbp, cp, store.Restore(cp), reopened: null));
        }

        // open in an IDE: close (snapshot so VB's silent save on close can be undone), restore, open again
        if (OwnerOf(inst.Pid) is { } owner && !IdeOwnership.IsMe(owner))
            throw new ToolException($"{open.Name} is open in VB6 {inst.Pid}, which belongs to another session ({owner.Client}); project_rollback won't change it");
        RequireDesignMode(inst);
        var client = _env.Client(inst);
        var dirty = (await client.ComponentsAsync(open.Name)).Where(c => c.IsDirty).ToList();
        if ((dirty.Count > 0 || open.IsDirty) && a.Bool("discard_unsaved") != true)
            throw new ToolException($"{open.Name} has unsaved changes ({(dirty.Count > 0 ? string.Join(", ", dirty.Select(c => c.Name)) : "project settings")}) that the rollback would lose: pass discard_unsaved=true, or save them and take a checkpoint first");
        var snap = await TakeSnapshotAsync(inst, open.FileName, dirty);
        Audit.Write("project_rollback", inst.Pid, $"{open.Name} ({vbp}) to {cp.Id}{(dirty.Count > 0 ? $", discarding {string.Join(", ", dirty.Select(c => c.Name))}" : "")}");
        try { await client.CloseAsync(open.Name, discard: dirty.Count > 0 || open.IsDirty); }
        catch (AddinErrorException ex) when (ex.Code == 409) { throw new ToolException(ex.AddinMessage); }
        var closed = await WaitForProjectsAsync(inst, client, ps => !ps.Any(p => SamePath(p.FileName, vbp)), ct);
        if (closed.Dialog is { } d) return d;
        if (!closed.Done) throw new ToolException("the project didn't close, nothing was restored: " + (AddinLogSince(inst.Pid, closed.Started, "close failed") ?? "check ide_status"));
        await Task.Delay(500, ct);   // VB has finished writing once the project is gone
        RestoreSnapshot(snap);       // undo what the close wrote, and drop files VB created for never-saved components
        var result = store.Restore(cp);
        await client.OpenAsync(vbp, replace: false);
        var opened = await WaitForProjectsAsync(inst, client, ps => ps.Any(p => SamePath(p.FileName, vbp)), ct);
        if (opened.Dialog is { } d2) return d2;
        object reopened = opened.Done ? new { pid = inst.Pid, project = opened.Projects.First(p => SamePath(p.FileName, vbp)).Name }
            : new { pid = inst.Pid, error = AddinLogSince(inst.Pid, opened.Started, "open failed") ?? "the project didn't open again; open it with project_open" };
        return ToolResult.Ok(RollbackJson(vbp, cp, result, reopened));
    }

    private static object RollbackJson(string vbp, Checkpoint cp, RollbackResult r, object? reopened) => new
    {
        status = r.Failed.Count > 0 ? "partly restored" : "restored", vbp, checkpoint = cp.Id, label = cp.Label,
        restored = r.Restored, unchanged = r.Unchanged, failed = r.Failed.Count > 0 ? r.Failed : null,
        addedSince = r.AddedSince.Count > 0 ? r.AddedSince : null, reopened,
        note = r.AddedSince.Count > 0 ? "files added since the checkpoint are still on disk; the restored .vbp doesn't name them" : null,
    };
}
