using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Project-level tools (PLAN Phase 7 A1-A3): close (with discard), add components, project settings. Agents had to
/// write new files and .vbp lines by script, or ask the user to restart VB6, for want of these (improvement notes 1-3).
/// </summary>
public sealed partial class IdeTools
{
    private static readonly string[] IdeProps = ["name", "description", "helpfile", "helpcontextid", "startmode", "buildfilename", "startup"];

    private IReadOnlyList<Tool> ProjectTools() =>
    [
        new("project_close",
            "Close a project in the IDE. Refused while it has unsaved changes, unless discard=true, which throws those changes away " +
            "(recovery copies the watchdog made are kept). Design mode only.",
            Schema.Object(("project", "string", ProjectDesc, false), ("discard", "boolean", "Close even with unsaved changes, discarding them.", false),
                          ("pid", "integer", PidDesc, false)),
            (a, ct) => CloseProjectAsync(a, ct)),
        new("project_add_component",
            "Add a component to a project open in the IDE: a new one (kind: module, class, form, mdiform, usercontrol, propertypage, " +
            "with a name) or an existing source file (file: full path under the workspace roots). Unsaved until project_save, which " +
            "also adds it to the .vbp. Design mode only.",
            Schema.Object(("kind", "string", "module, class, form, mdiform, usercontrol or propertypage (for a new component).", false),
                          ("name", "string", "Name of the new component.", false),
                          ("file", "string", "Existing .bas/.cls/.frm/.ctl/.pag/.dsr file to add instead.", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => AddComponentAsync(a, ct)),
        new("project_settings",
            "Read or change project settings. For a project open in the IDE: name, description, helpfile, helpcontextid, startmode " +
            "(0 standalone, 1 ActiveX component), buildfilename and startup (a form, 'Sub Main' or '(None)') are changed live. " +
            "Settings only the .vbp has (CondComp for conditional compilation constants, Command32 for the command line, ...) need " +
            "the project closed: pass reopen=true to close it (it must have no unsaved changes), edit the .vbp and open it again. " +
            "For a closed project pass path: set uses .vbp key names (CondComp, Command32, Startup, ...). References, controls, " +
            "members and binary compatibility have their own tools.",
            Schema.Object(("set", "object", "Settings to change: {\"CondComp\": \"nologin = 1\", \"startup\": \"frmMain\"}.", false),
                          ("reopen", "boolean", "Allow closing and reopening an open project to change .vbp-only settings.", false),
                          ("path", "string", "A closed project's .vbp (instead of project).", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => ProjectSettingsAsync(a, ct)),
    ];

    // ---------- project_close ----------

    private async Task<ToolResult> CloseProjectAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireDesignMode(inst);
        var (client, project) = await ResolveProjectAsync(a);
        bool discard = a.Bool("discard") == true;   // hard rule 9: never a default
        var comps = await client.ComponentsAsync(project);
        var dirty = comps.Where(c => c.IsDirty).ToList();
        var prjFile = (await client.ProjectsAsync()).First(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase)).FileName;

        // VBIDE can't close without saving: VBProjects.Remove saves changed components silently, and IsDirty = False saves
        // too (spike 31). So a discard snapshots the files the close will write (the .vbp and each changed component's
        // file), lets VB close, then writes the snapshots back and deletes source files VB created for never-saved
        // components. Refused if another IDE has one of those files loaded (hard rule 4: no writes under a loaded file).
        Snapshot? snap = null;
        if (discard && (dirty.Count > 0 || (await client.ProjectsAsync()).Any(p => p.Name == project && p.IsDirty)))
        {
            if (prjFile.Length == 0) throw new ToolException("the project has never been saved, so a discard can't restore anything; VB would ask where to save it");
            snap = await TakeSnapshotAsync(inst, prjFile, dirty);
            Audit.Write("project_close", inst.Pid, $"{project} discard ({string.Join(", ", dirty.Select(c => c.Name))})");
        }
        try { await client.CloseAsync(project, discard); }
        catch (AddinErrorException ex) when (ex.Code == 409) { throw new ToolException(ex.AddinMessage); }
        var closed = await WaitForProjectsAsync(inst, client, ps => !ps.Any(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase)), ct);
        if (closed.Dialog is { } d) return d;
        if (!closed.Done)
            return ToolResult.Ok(new { status = "pending", project, error = AddinLogSince(inst.Pid, closed.Started, "close failed"),
                                       note = snap is null ? null : "nothing was restored: the project didn't close" });
        object? restored = null;
        if (snap is not null)
        {
            await Task.Delay(500, ct);   // VB has finished writing once the project is gone; a moment for the file system
            restored = RestoreSnapshot(snap);
        }
        return ToolResult.Ok(new
        {
            status = "closed", project, discarded = snap is not null ? dirty.Select(c => c.Name) : null, restored,
            projects = closed.Projects.Select(p => p.Name),
        });
    }

    /// <summary>Bytes of the files a close will write, and the source files that existed in the project folder.</summary>
    private sealed record Snapshot(Dictionary<string, byte[]?> Files, string Folder, HashSet<string> Existing);

    private static readonly string[] SourceExtensions = [".bas", ".cls", ".frm", ".frx", ".ctl", ".ctx", ".pag", ".pgx", ".dsr", ".dsx"];

    private async Task<Snapshot> TakeSnapshotAsync(InstanceInfo inst, string prjFile, IReadOnlyList<ComponentInfo> dirty)
    {
        var files = new Dictionary<string, byte[]?>(StringComparer.OrdinalIgnoreCase) { [prjFile] = File.ReadAllBytes(prjFile) };
        foreach (var c in dirty.Where(c => c.FileName.Length > 0))
        {
            files[c.FileName] = File.Exists(c.FileName) ? File.ReadAllBytes(c.FileName) : null;
            // a form's binary .frx (and .ctx/.pgx) is written along with it
            foreach (var bin in new[] { ".frx", ".ctx", ".pgx" }.Select(x => Path.ChangeExtension(c.FileName, x)))
                if (!files.ContainsKey(bin)) files[bin] = File.Exists(bin) ? File.ReadAllBytes(bin) : null;
        }
        // hard rule 4: another IDE with one of these files loaded would be written under
        foreach (var other in _env.Scan().Live.Where(i => i.Pid != inst.Pid))
        {
            try
            {
                var oc = _env.Client(other);
                foreach (var p in await oc.ProjectsAsync())
                {
                    if (files.ContainsKey(p.FileName)) throw new ToolException($"{p.FileName} is also open in VB6 {other.Pid}; close it there first (a discard restores files on disk)");
                    foreach (var c in await oc.ComponentsAsync(p.Name))
                        if (c.FileName.Length > 0 && files.ContainsKey(c.FileName)) throw new ToolException($"{c.FileName} is also loaded in VB6 {other.Pid} ({p.Name}); a discard would write under it");
                }
            }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException)
            {
                throw new ToolException($"VB6 {other.Pid} didn't answer, so it can't be ruled out that it has these files loaded; try again when it responds");
            }
        }
        var folder = Path.GetDirectoryName(prjFile)!;
        var existing = Directory.EnumerateFiles(folder).Where(p => SourceExtensions.Contains(Path.GetExtension(p).ToLowerInvariant())).ToHashSet(StringComparer.OrdinalIgnoreCase);
        return new Snapshot(files, folder, existing);
    }

    private static object RestoreSnapshot(Snapshot snap)
    {
        var restored = new List<string>(); var deleted = new List<string>(); var failed = new List<string>();
        foreach (var (path, bytes) in snap.Files)
        {
            try
            {
                if (bytes is null) { if (File.Exists(path)) { File.Delete(path); deleted.Add(path); } continue; }
                if (File.Exists(path) && File.ReadAllBytes(path).AsSpan().SequenceEqual(bytes)) continue;
                File.WriteAllBytes(path, bytes);
                restored.Add(path);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { failed.Add($"{path}: {ex.Message}"); }
        }
        // source files VB created for components that were never saved
        foreach (var p in Directory.EnumerateFiles(snap.Folder).Where(p => SourceExtensions.Contains(Path.GetExtension(p).ToLowerInvariant()) && !snap.Existing.Contains(p)))
        {
            try { File.Delete(p); deleted.Add(p); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { failed.Add($"{p}: {ex.Message}"); }
        }
        return new { files = restored, deleted = deleted.Count > 0 ? deleted : null, failed = failed.Count > 0 ? failed : null };
    }

    /// <summary>
    /// Polls PROJECTS until <paramref name="done"/> holds, a dialog appears, or OpenWait passes. <paramref name="answer"/>
    /// may handle an expected dialog (returns true when it clicked something); any other dialog is reported, not answered.
    /// </summary>
    private async Task<(bool Done, ToolResult? Dialog, IReadOnlyList<ProjectInfo> Projects, DateTime Started)> WaitForProjectsAsync(
        InstanceInfo inst, AddinClient client, Func<IReadOnlyList<ProjectInfo>, bool> done, CancellationToken ct,
        Func<AddinClient, DialogInfo, Task<bool>>? answer = null)
    {
        var started = DateTime.Now;
        var deadline = DateTime.UtcNow + OpenWait;
        IReadOnlyList<ProjectInfo> now = [];
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(300, ct);
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count > 0 && answer is not null && await answer(client, w.Dialogs[0])) continue;
            if (w.Dialogs.Count > 0)
                return (false, ToolResult.Ok(new { status = "dialog", note = "VB6 is showing a dialog; it was not answered", dialogs = Dialogs(w) }), now, started);
            try { now = await client.ProjectsAsync(); } catch (AddinBusyException) { continue; }
            if (done(now)) return (true, null, now, started);
        }
        return (false, null, now, started);
    }

    // ---------- project_add_component ----------

    private async Task<ToolResult> AddComponentAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireDesignMode(inst);
        var (client, project) = await ResolveProjectAsync(a);
        string? file = a.Str("file"), kind = a.Str("kind"), name = a.Str("name"), newFile = null;
        string expected;
        if (file is { Length: > 0 })
        {
            if (kind is not null || name is not null) throw new ToolException("pass either file, or kind and name");
            var path = PathKey.Normalize(file);
            if (Roots.Count == 0 || !ProjectMap.IsUnderRoots(path, Roots)) throw new ToolException($"{path} is outside the workspace roots");
            if (!File.Exists(path)) throw new ToolException("file not found: " + path);
            if (VbSourceFile.KindOf(path) is VbFileKind.Other) throw new ToolException("not a VB6 source file (.bas/.cls/.frm/.ctl/.pag/.dsr)");
            expected = VbSourceFile.Load(path).VbName ?? Path.GetFileNameWithoutExtension(path);
            try { await client.AddComponentFileAsync(project, path); }
            catch (AddinErrorException ex) when (ex.Code is 400 or 404 or 409) { throw new ToolException(ex.AddinMessage); }
        }
        else
        {
            if (kind is null || name is null) throw new ToolException("pass kind and name for a new component, or file for an existing one");
            expected = name;
            // A new component has no file, and File > Save Project then stops at a Save As dialog for it (verified
            // 2026-10-17). So it is saved right away, next to the .vbp, under its name: checked before adding anything.
            var prjFile = (await client.ProjectsAsync()).First(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase)).FileName;
            if (prjFile.Length > 0)
            {
                newFile = Path.Combine(Path.GetDirectoryName(prjFile)!, name + NewFileExtension(kind));
                if (File.Exists(newFile)) throw new ToolException($"{newFile} already exists; add it with file=..., or choose another name");
            }
            try { await client.AddComponentAsync(project, kind, name); }
            catch (AddinErrorException ex) when (ex.Code is 400 or 404 or 409) { throw new ToolException(ex.AddinMessage); }
        }
        // the add runs deferred: wait for the component to appear
        var deadline = DateTime.UtcNow + OpenWait;
        var started = DateTime.Now;
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(300, ct);
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count > 0) return ToolResult.Ok(new { status = "dialog", project, note = "VB6 is showing a dialog while adding; it was not answered", dialogs = Dialogs(w) });
            IReadOnlyList<ComponentInfo> comps;
            try { comps = await client.ComponentsAsync(project); } catch (AddinBusyException) { continue; }
            if (comps.FirstOrDefault(c => c.Name.Equals(expected, StringComparison.OrdinalIgnoreCase)) is { } added)
            {
                if (newFile is not null && added.FileName.Length == 0) added = await SaveNewComponentAsync(inst, client, project, added, newFile, ct);
                return ToolResult.Ok(new
                {
                    status = "added", project, component = added.Name, type = ComponentType(added.Type), file = added.FileName.Length > 0 ? added.FileName : null, dirty = added.IsDirty,
                    note = added.FileName.Length > 0 && newFile is not null
                        ? "its file is written; project_save adds it to the .vbp"
                        : "project_save adds it to the .vbp",
                });
            }
        }
        return ToolResult.Ok(new { status = "pending", project, error = AddinLogSince(inst.Pid, started, "add component failed") });
    }

    private static string NewFileExtension(string kind) => kind.ToLowerInvariant() switch
    {
        "module" => ".bas", "class" => ".cls", "form" or "mdiform" => ".frm", "usercontrol" => ".ctl", "propertypage" => ".pag",
        _ => throw new ToolException("kind must be module, class, form, mdiform, usercontrol or propertypage"),
    };

    /// <summary>Saves a just-added component to its new file (SAVECOMP with a path) and waits until VB reports the file.</summary>
    private async Task<ComponentInfo> SaveNewComponentAsync(InstanceInfo inst, AddinClient client, string project, ComponentInfo added, string path, CancellationToken ct)
    {
        try { await client.SaveComponentAsync(project, added.Name, path); }
        catch (AddinErrorException ex) when (ex.Code is 400 or 409) { throw new ToolException($"{added.Name} was added but couldn't be saved to {path}: {ex.AddinMessage}"); }
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(10);
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(300, ct);
            IReadOnlyList<ComponentInfo> comps;
            try { comps = await client.ComponentsAsync(project); } catch (AddinBusyException) { continue; }
            if (comps.FirstOrDefault(c => c.Name.Equals(added.Name, StringComparison.OrdinalIgnoreCase)) is { FileName.Length: > 0 } saved) return saved;
        }
        return added;
    }

    // ---------- project_settings ----------

    private async Task<ToolResult> ProjectSettingsAsync(JsonObject a, CancellationToken ct)
    {
        var set = a["set"] switch
        {
            null => [],
            JsonObject o => o.Select(kv => (Key: kv.Key, Value: kv.Value is JsonValue v ? (v.TryGetValue(out string? t) ? t : v.ToJsonString()) : null)).ToList(),
            _ => throw new ToolException("set must be an object: {\"CondComp\": \"nologin = 1\"}"),
        };
        if (set.FirstOrDefault(s => s.Value is null).Key is { } badKey) throw new ToolException($"setting {badKey}: the value must be text, a number or true/false");

        if (a.Str("path") is { Length: > 0 } vbpArg) return await ClosedProjectSettingsAsync(PathKey.Normalize(vbpArg), set!);

        var inst = ResolveInstance(a);
        var (client, project) = await ResolveProjectAsync(a);
        var prj = (await client.ProjectsAsync()).First(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase));
        var ideSet = set.Where(s => IdeProps.Contains(s.Key, StringComparer.OrdinalIgnoreCase)).ToList();
        var vbpSet = set.Where(s => !IdeProps.Contains(s.Key, StringComparer.OrdinalIgnoreCase)).ToList();
        if (set.Count > 0) RequireDesignMode(inst);
        // validate the .vbp-only keys before changing anything
        if (vbpSet.Count > 0)
        {
            if (prj.FileName.Length == 0) throw new ToolException("the project has never been saved, so it has no .vbp yet; save it first");
            var probe = Vb1252.Encoding.GetString(File.ReadAllBytes(prj.FileName));
            foreach (var (k, v) in vbpSet)
                try { probe = VbpKeys.Set(probe, k, v!); } catch (ArgumentException ex) { throw new ToolException($"{k}: {ex.Message}"); }
            if (a.Bool("reopen") != true)
                throw new ToolException($"{string.Join(", ", vbpSet.Select(s => s.Key))} {(vbpSet.Count == 1 ? "is" : "are")} only in the .vbp, which VB6 rewrites from memory when it saves. " +
                    "Changing them needs the project closed and reopened: pass reopen=true (the project must have no unsaved changes), or close it first and pass path.");
        }

        ProjectProps props = await client.ProjectPropsAsync(project);
        foreach (var (k, v) in ideSet)
        {
            try { props = await client.SetProjectPropAsync(project, k.ToLowerInvariant(), v!); }
            catch (AddinErrorException ex) when (ex.Code is 400 or 404 or 409) { throw new ToolException($"{k}: {ex.AddinMessage}"); }
            if (k.Equals("name", StringComparison.OrdinalIgnoreCase)) project = props.Name;
        }

        object? reopened = null;
        if (vbpSet.Count > 0)
        {
            var dirty = (await client.ComponentsAsync(project)).Where(c => c.IsDirty).Select(c => c.Name).ToList();
            var now = (await client.ProjectsAsync()).First(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase));
            if (dirty.Count > 0 || now.IsDirty)
                throw new ToolException($"{project} has unsaved changes ({(dirty.Count > 0 ? string.Join(", ", dirty) : "project settings")}); save it (project_save) before reopening" +
                    (ideSet.Count > 0 ? ". The IDE settings above were applied and are unsaved too." : ""));
            var path = now.FileName;
            try { await client.CloseAsync(project, discard: false); }
            catch (AddinErrorException ex) when (ex.Code == 409) { throw new ToolException(ex.AddinMessage); }
            var closed = await WaitForProjectsAsync(inst, client, ps => !ps.Any(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase)), ct);
            if (closed.Dialog is { } d) return d;
            if (!closed.Done) throw new ToolException("the project didn't close: " + (AddinLogSince(inst.Pid, closed.Started, "close failed") ?? "check ide_status"));
            var text = Vb1252.Encoding.GetString(File.ReadAllBytes(path));
            foreach (var (k, v) in vbpSet) text = VbpKeys.Set(text, k, v!);
            WriteVbp(path, text);
            await client.OpenAsync(path, replace: false);
            var opened = await WaitForProjectsAsync(inst, client, ps => ps.Any(p => SamePath(p.FileName, path)), ct);
            if (opened.Dialog is { } d2) return d2;
            if (!opened.Done) throw new ToolException($"the .vbp was changed, but the project didn't open again: {AddinLogSince(inst.Pid, opened.Started, "open failed") ?? "check ide_status"}");
            reopened = new { path, changed = vbpSet.Select(s => s.Key) };
        }

        var file = (await client.ProjectsAsync()).FirstOrDefault(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase))?.FileName;
        return ToolResult.Ok(new
        {
            source = "ide", project,
            settings = new { name = props.Name, description = props.Description, helpFile = props.HelpFile, helpContextId = props.HelpContextId, startMode = props.StartMode, buildFileName = props.BuildFileName, startup = props.Startup, type = props.Type },
            vbp = file is { Length: > 0 } && File.Exists(file) ? VbpSection(file) : null, reopened,
            note = ideSet.Count > 0 ? "IDE settings are unsaved; project_save writes them to the .vbp" : null,
        });
    }

    private async Task<ToolResult> ClosedProjectSettingsAsync(string vbp, IReadOnlyList<(string Key, string Value)> set)
    {
        if (!vbp.EndsWith(".vbp", StringComparison.OrdinalIgnoreCase)) throw new ToolException("path must be a .vbp file");
        if (Roots.Count == 0 || !ProjectMap.IsUnderRoots(vbp, Roots)) throw new ToolException($"{vbp} is outside the workspace roots");
        if (!File.Exists(vbp)) throw new ToolException("file not found: " + vbp);
        // hard rule 4: never write a file an IDE has loaded
        foreach (var inst in _env.Scan().Live)
        {
            try
            {
                if ((await _env.Client(inst).ProjectsAsync()).FirstOrDefault(p => SamePath(p.FileName, vbp)) is { } open)
                    throw new ToolException($"{vbp} is open in VB6 {inst.Pid} as {open.Name}; change it there (project_settings without path, reopen=true for .vbp-only keys)");
            }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException)
            {
                throw new ToolException($"VB6 {inst.Pid} didn't answer, so it can't be ruled out that it has {vbp} open; try again when it responds");
            }
        }
        if (set.Count > 0)
        {
            var text = Vb1252.Encoding.GetString(File.ReadAllBytes(vbp));
            foreach (var (k, v) in set)
                try { text = VbpKeys.Set(text, k, v); } catch (ArgumentException ex) { throw new ToolException($"{k}: {ex.Message}"); }
            WriteVbp(vbp, text);
        }
        return ToolResult.Ok(new { source = "disk", path = vbp, vbp = VbpSection(vbp), changed = set.Count > 0 ? set.Select(s => s.Key) : null });
    }

    /// <summary>The .vbp's project-section keys (members, references and controls left out: project_components and refs_list show those).</summary>
    private static Dictionary<string, string> VbpSection(string path)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in Vb1252.Encoding.GetString(File.ReadAllBytes(path)).Replace("\r\n", "\n").Split('\n'))
        {
            if (line.StartsWith('[')) break;
            int eq = line.IndexOf('=');
            if (eq <= 0) continue;
            var key = line[..eq];
            if (VbpKeys.Forbidden.Contains(key, StringComparer.OrdinalIgnoreCase) && key != "Type") continue;
            result.TryAdd(key, line[(eq + 1)..].Trim('"'));
        }
        return result;
    }

    private static void WriteVbp(string path, string text)
    {
        var tmp = path + ".vb6mcp.tmp";
        File.WriteAllBytes(tmp, Vb1252.Encoding.GetBytes(text));
        File.Move(tmp, path, overwrite: true);
    }
}
