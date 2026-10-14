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
/// Phase 4 sync tools (PLAN.md §8). The watchdog's sync engine reloads clean components whose files
/// changed outside the IDE and records conflicts (unsaved IDE changes + changed file); these tools
/// show that state and resolve conflicts, always explicitly (hard rule 9) and without losing work.
/// </summary>
public sealed partial class IdeTools
{
    private IReadOnlyList<Tool> SyncTools() =>
    [
        new("sync_status",
            "Files changed outside the IDE: conflicts (a component's file changed on disk while it has unsaved changes in the IDE; " +
            "nothing was touched) and recent sync events (automatic reloads of unmodified components, pending reloads, deleted files, " +
            "changed .vbp files). Kept by the watchdog's sync engine.",
            Schema.Object(("pid", "integer", "Only this IDE.", false)), (a, _) => Task.FromResult(SyncStatus(a))),
        new("sync_diff",
            "Differences between a component's code in the IDE (including unsaved changes) and its file on disk, as a unified diff " +
            "(--- ide, +++ disk).",
            Schema.Object(("component", "string", ComponentDesc, true), ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => SyncDiffAsync(a)),
        new("sync_take_disk",
            "Resolve a conflict in favour of the file on disk: reloads the component, DISCARDING its unsaved IDE changes. A recovery " +
            "copy of the IDE text is kept first (recovery_list / recovery_restore). Design mode. Needs confirm=true.",
            Schema.Object(("component", "string", ComponentDesc, true), ("confirm", "boolean", "Required: true.", true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => SyncTakeDiskAsync(a, ct)),
        new("sync_take_ide",
            "Resolve a conflict in favour of the IDE: saves the project (File > Save Project), writing the IDE's version over the changed " +
            "file. This saves every unsaved component of the project; the result lists them. Needs confirm=true.",
            Schema.Object(("component", "string", ComponentDesc, true), ("confirm", "boolean", "Required: true.", true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => SyncTakeIdeAsync(a, ct)),
    ];

    private ToolResult SyncStatus(JsonObject a)
    {
        int? pid = a.Int("pid");
        var s = SyncState.Read(SyncStatePath ?? SyncState.DefaultPath);
        return ToolResult.Ok(new
        {
            conflicts = s.Conflicts.Where(c => pid is null || c.Pid == pid).Select(c => new
            {
                pid = c.Pid, project = c.Project, component = c.Component, file = c.File,
                detected = c.Detected.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
            }),
            events = s.Events.Where(e => pid is null || e.Pid == pid).TakeLast(30).Select(e => new
            {
                time = e.Time.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), pid = e.Pid, kind = e.Kind, project = e.Project,
                component = e.Component, file = e.File, note = e.Note,
            }),
            updated = s.Updated == DateTime.MinValue ? null : s.Updated.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
            note = _env.FindWatchdog() is null ? "the watchdog isn't running, so outside changes aren't being watched now" : null,
        });
    }

    /// <summary>Where the watchdog writes the sync state; tests point it elsewhere.</summary>
    public string? SyncStatePath { get; init; }

    private async Task<(AddinClient Client, string Project, ComponentInfo Comp)> ResolveComponentAsync(JsonObject a)
    {
        var (client, project) = await ResolveProjectAsync(a);
        var name = a.Req("component");
        var comp = (await client.ComponentsAsync(project)).FirstOrDefault(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                   ?? throw new ToolException($"component {name} not found in {project}");
        if (comp.FileName.Length == 0) throw new ToolException($"{comp.Name} has never been saved to a file");
        return (client, project, comp);
    }

    private static string DiskCode(string file)
    {
        try { return VbSourceFile.Load(file).Code; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { throw new ToolException($"{file} can't be read: {ex.Message}"); }
    }

    private async Task<ToolResult> SyncDiffAsync(JsonObject a)
    {
        var (client, project, comp) = await ResolveComponentAsync(a);
        var ide = (await client.ReadAsync(project, comp.Name)).Text;
        var disk = DiskCode(comp.FileName);
        var diff = LineDiff.Unified(ide, disk, $"ide ({project}.{comp.Name}{(comp.IsDirty ? ", unsaved changes" : "")})", $"disk ({comp.FileName})");
        return ToolResult.Ok(new { project, component = comp.Name, file = comp.FileName, dirty = comp.IsDirty, same = diff.Length == 0, diff });
    }

    private async Task<ToolResult> SyncTakeDiskAsync(JsonObject a, CancellationToken ct)
    {
        if (a.Bool("confirm") != true) throw new ToolException("sync_take_disk needs confirm=true: it discards the component's unsaved IDE changes (a recovery copy is kept)");
        var (client, project, comp) = await ResolveComponentAsync(a);
        var inst = client.Instance;
        RequireMode(inst, IdeMode.Design);
        var ide = await client.ReadAsync(project, comp.Name);
        var disk = DiskCode(comp.FileName);
        // keep the IDE text first: nothing is ever lost (hard rule 4)
        string? kept = null;
        if (comp.IsDirty)
        {
            // in its own session: the watchdog prunes copies of components that are no longer unsaved, and
            // after the reload this one isn't (seen live)
            int keptId = RecoveryStore.KeptId(inst.Pid);
            Store.WriteSession(new RecoverySession(keptId, "", DateTime.Now, true, [], $"kept by sync_take_disk from VB6 pid {inst.Pid} before discarding its unsaved changes"));
            Store.WriteCopy(keptId, new RecoveryCopy(project, "", comp.Name, comp.Type, comp.FileName, ide.CountOfLines, VbCode.Hash(ide.Text), DateTime.Now), ide.Text);
            kept = $"recovery copy kept 7 days (recovery_read from_pid={keptId} component={comp.Name})";
        }
        Audit.Write("sync_take_disk", inst.Pid, $"{project}.{comp.Name} ({comp.FileName})");
        await client.ReloadAsync(project, comp.Name, "discard");
        string? now = null;
        for (int i = 0; i < 20 && now != disk; i++) { await Task.Delay(200, ct); now = (await client.ReadAsync(project, comp.Name)).Text; }
        return ToolResult.Ok(new
        {
            status = now == disk ? "reloaded" : "unknown", project, component = comp.Name, file = comp.FileName, ideMatchesDisk = now == disk, kept,
            note = now == disk ? null : "the IDE text doesn't match the disk file yet; see %LOCALAPPDATA%\\vb6mcp\\addin.log",
        });
    }

    private async Task<ToolResult> SyncTakeIdeAsync(JsonObject a, CancellationToken ct)
    {
        if (a.Bool("confirm") != true) throw new ToolException("sync_take_ide needs confirm=true: it overwrites the changed file on disk with the IDE's version");
        var (client, project, comp) = await ResolveComponentAsync(a);
        var inst = client.Instance;
        var alsoSaved = (await client.ComponentsAsync(project)).Where(c => c.IsDirty && c.Name != comp.Name).Select(c => c.Name).ToList();
        Audit.Write("sync_take_ide", inst.Pid, $"{project} save (conflict on {comp.Name})");
        await client.SaveAsync(project);
        bool clean = false;
        for (int i = 0; i < 25 && !clean; i++)
        {
            await Task.Delay(200, ct);
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count > 0) return ToolResult.Ok(new { status = "dialog", dialogs = Dialogs(w), note = "saving shows a dialog; it was not answered" });
            clean = !(await client.ComponentsAsync(project)).Any(c => c.IsDirty);
        }
        var ide = (await client.ReadAsync(project, comp.Name)).Text;
        return ToolResult.Ok(new
        {
            status = clean ? "saved" : "pending", project, component = comp.Name, file = comp.FileName,
            diskMatchesIde = DiskCode(comp.FileName) == ide, alsoSaved,
        });
    }
}
