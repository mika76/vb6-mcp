using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Editing tools: change live code in the IDE, save, answer dialogs. Safety rules:
/// - optimistic concurrency: every edit needs the module hash from code_read/proc_get; the server
///   re-reads and refuses on mismatch, and the add-in re-checks the exact target lines (409);
/// - design mode only (edits in break mode contributed to the spike 3 deadlock);
/// - VB6 limits and the 1252 encoding are checked before anything is sent (hard rules 1, 10);
/// - saving and dialog answers are explicit (hard rules 4, 9). Documented in docs/tools.md.
/// </summary>
public sealed partial class IdeTools
{
    /// <summary>What the tools last left in each component (improvement notes 50); shared with WorkspaceTools.</summary>
    public ToolEditLedger Ledger { get; init; } = new();

    private const string HashDesc = "The module hash from code_read or proc_get. The edit is refused if the module changed since.";

    /// <summary>ide_dialog_respond repeat: the most clicks in one call (a big project shows one load error per form).</summary>
    private const int MaxRepeatClicks = 500;

    private IReadOnlyList<Tool> EditTools() =>
    [
        new("code_replace_lines",
            "Replace lines of a component's code in the IDE (unsaved, undoable in VB6). count 0 inserts before start_line; " +
            "empty new_text deletes. Several ranges in one call: edits (all numbered against the text you read). Needs expected_hash " +
            "from code_read/proc_get; design mode only. Returns the new hash. Mechanical renames: find/replace (preview first, then apply=true).",
            FindReplace.WithSchema(EditList.WithSchema(Schema.Object(("component", "string", ComponentDesc, true), ("start_line", "integer", "First line to replace (1-based).", false),
                          ("count", "integer", "Number of lines to replace (0 = insert).", false), ("new_text", "string", "Replacement lines." + EditList.CodeRule, false),
                          ("edits", "array", EditList.Description, false),
                          ("expected_hash", "string", HashDesc, true), ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)))),
            (a, _) => ReplaceLinesAsync(a)),
        new("code_insert",
            "Insert lines into a component before line `line` (use line = countOfLines + 1 to append). Same rules as code_replace_lines.",
            Schema.Object(("component", "string", ComponentDesc, true), ("line", "integer", "Insert before this line (1-based).", true),
                          ("text", "string", "Lines to insert." + EditList.CodeRule, true), ("expected_hash", "string", HashDesc, true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => InsertAsync(a)),
        new("proc_replace",
            "Replace a whole procedure or property (from its first comment line to End Sub/Function/Property) with new_text, " +
            "exactly the range proc_get returns. Same rules as code_replace_lines.",
            Schema.Object(("component", "string", ComponentDesc, true), ("name", "string", "Procedure or property name.", true),
                          ("new_text", "string", "The complete new procedure, including comments above it." + EditList.CodeRule, true),
                          ("expected_hash", "string", HashDesc, true), ("kind", "string", "proc (default), get, let or set.", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => ProcReplaceAsync(a)),
        new("project_save",
            "Save a project like File > Save Project (all unsaved components and the .vbp). Explicit only: edits are never " +
            "saved automatically. A never-saved project (File > New Project) needs path: the new .vbp under the workspace roots, " +
            "its components saved next to it. Dialogs shown while saving are reported, not answered. " +
            "components: save only these components, without the .vbp (which File > Save Project rewrites, e.g. reference versions). " +
            "Code changed in the IDE by anything but these tools (stray typing, the user) is refused with its diff unless outside_edits=true.",
            Schema.Object(("project", "string", ProjectDesc, false), ("components", "array", "Save only these components (not the .vbp).", false),
                          ("path", "string", "New .vbp for a project that has never been saved.", false),
                          ("outside_edits", "boolean", "Save even code changes no vb6-mcp tool made (refused by default, with the diff).", false),
                          ("pid", "integer", PidDesc, false)),
            (a, ct) => SaveAsync(a, ct)),
        new("ide_dialog_respond",
            "Click a button in a dialog the IDE is showing (see ide_status). The button must be named explicitly; " +
            "pass dialog (its caption) when several dialogs are open. repeat=true keeps clicking it while the same message comes back for " +
            "form after form (a missing control or reference makes VB show \"Errors during load\" once per form), then says what the load " +
            "logs report as missing.",
            Schema.Object(("button", "string", "Button caption, e.g. OK, Yes, No, Cancel.", true),
                          ("dialog", "string", "Dialog caption; needed when more than one dialog is open.", false),
                          ("repeat", "boolean", "Keep clicking while the same kind of message comes back (at most 500).", false),
                          ("pid", "integer", PidDesc, false)),
            (a, ct) => DialogRespondAsync(a, ct)),
    ];

    // ---------- edits ----------

    private async Task<ToolResult> ReplaceLinesAsync(JsonObject a)
    {
        if (FindReplace.Args(a) is { } fr)
        {
            // find/replace (PLAN.md A6): the matches become one edit list against the same hash; a preview changes nothing
            IReadOnlyList<FoundLine> Find(string text) => FindReplace.Compute(text, fr.Find, fr.Replace, fr.WholeWord, fr.MatchCase, fr.InStrings, fr.InComments);
            if (!fr.Apply)
            {
                var (client, project) = await ResolveProjectAsync(a);
                var component = a.Req("component");
                var module = await client.ReadAsync(project, component);
                var hash = VbCode.Hash(module.Text);
                if (!string.Equals(hash, a.Req("expected_hash"), StringComparison.OrdinalIgnoreCase))
                    throw new ToolException($"conflict: {component} changed since you read it (expected hash {a.Str("expected_hash")}, now {hash}). Re-read it with code_read.");
                return ToolResult.Ok(new { source = "ide", project, component, preview = FindReplace.Preview(Find(module.Text)), hash });
            }
            return await EditManyAsync(a, async (client, target) =>
            {
                var found = Find((await client.ReadAsync(target.Project, target.Component)).Text);
                if (found.Count == 0) throw new ToolException("no matches, nothing was changed (whole words, code only by default)");
                return FindReplace.Edits(found);
            });
        }
        var edits = EditList.From(a);   // validated (lengths, overlaps) before anything is sent to the IDE
        return await EditManyAsync(a, (_, _) => Task.FromResult(edits));
    }

    private Task<ToolResult> InsertAsync(JsonObject a)
    {
        int line = a.Int("line") ?? throw new ToolException("missing required argument 'line'");
        return EditAsync(a, a.Str("text") ?? throw new ToolException("missing required argument 'text'"),
            (_, _) => Task.FromResult((line, 0)));
    }

    private Task<ToolResult> ProcReplaceAsync(JsonObject a)
    {
        var name = a.Req("name");
        int kind = ParseKind(a.Str("kind"));
        return EditAsync(a, a.Str("new_text") ?? throw new ToolException("missing required argument 'new_text'"),
            async (client, target) =>
            {
                var p = await client.ProcAsync(target.Project, target.Component, name, kind);
                var (_, start, count) = VbCode.TrimLeadingBlankLines(p.Text, p.ProcStart, p.ProcCount);
                return (start, count);
            });
    }

    /// <summary>Shared edit path for one range (code_insert, proc_replace): see <see cref="EditManyAsync"/>.</summary>
    private Task<ToolResult> EditAsync(JsonObject a, string newText, Func<AddinClient, (string Project, string Component), Task<(int Start, int Count)>> range)
    {
        string prepared;
        try { prepared = VbCode.PrepareForInsert(newText); }
        catch (ArgumentException ex) { throw new ToolException(ex.Message); }
        return EditManyAsync(a, async (client, target) =>
        {
            var (start, count) = await range(client, target);
            return [new LineEdit(start, count, prepared)];
        });
    }

    /// <summary>
    /// Resolve, check mode, check the module hash once, then replace each range bottom-up with its old text as guard (all
    /// ranges are numbered against the text the hash covers). If one fails, the result says which were applied: each is
    /// a separate VBIDE edit, undoable in VB6.
    /// </summary>
    private async Task<ToolResult> EditManyAsync(JsonObject a, Func<AddinClient, (string Project, string Component), Task<IReadOnlyList<LineEdit>>> ranges)
    {
        var inst = ResolveInstance(a);
        RequireDesignMode(inst);
        var (client, project) = await ResolveProjectAsync(a);
        var component = a.Req("component");
        var expected = a.Req("expected_hash");

        var module = await client.ReadAsync(project, component);
        var actual = VbCode.Hash(module.Text);
        if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
            throw new ToolException($"conflict: {component} changed since you read it (expected hash {expected}, now {actual}). Re-read it with code_read and redo the edit.");

        var edits = await ranges(client, (project, component));
        int total = module.CountOfLines;
        foreach (var e in edits)
            if (e.Start > total + 1 || e.Start + e.Count - 1 > total)
                throw new ToolException($"lines {e.Start}-{e.Start + e.Count - 1} are outside {component} ({total} lines); nothing was changed");
        var applied = new List<object>();
        var appliedEdits = new List<LineEdit>();
        foreach (var e in edits.OrderByDescending(x => x.Start))
        {
            string expectedOld = e.Count > 0 ? VbCode.Slice(module.Text, e.Start, e.Count) : "";
            try { await client.ReplaceAsync(project, component, e.Start, e.Count, expectedOld, e.Text); }
            catch (AddinErrorException ex) when (ex.Code == 409)
            {
                throw new ToolException($"conflict: lines {e.Start}-{e.Start + e.Count - 1} of {component} changed while editing ({ex.AddinMessage}). " +
                    (applied.Count == 0 ? "Nothing was changed. " : $"{applied.Count} later edit(s) were already applied (undo in VB6, or re-read). ") + "Re-read and retry.");
            }
            catch (Exception ex) when (ex is not ToolException)
            {
                // Any other failure (add-in error, busy, timeout) can leave the module part-way: say exactly what is in it
                // now, so the agent doesn't retry an edit that went in (improvement notes 30)
                throw new ToolException($"edit of lines {e.Start}-{e.Start + e.Count - 1} of {component} failed ({ex.Message}). " +
                    await DescribeAfterFailureAsync(client, project, component, module.Text, appliedEdits, e));
            }
            applied.Add(new { startLine = e.Start, oldLines = e.Count, newLines = e.NewLines });
            appliedEdits.Add(e);
        }
        applied.Reverse();   // report in file order

        // Everything below only describes an edit that is already in the IDE: none of it may turn that into an error,
        // or the agent retries and duplicates the edit (improvement notes 29: LiteralHints threw on "")
        CodeText after;
        try { after = await client.ReadAsync(project, component); }
        catch (Exception ex)
        {
            return ToolResult.Ok(new
            {
                source = "ide", project, component, replaced = applied.Count == 1 ? applied[0] : null, edits = applied.Count > 1 ? applied : null,
                note = $"applied, but re-reading the module failed ({ex.Message}); get the new hash with code_read before the next edit",
            });
        }
        Ledger.Record(inst.Pid, project, component, after.Text);   // improvement notes 50: what project_save compares with
        var typed = string.Join("\r\n", edits.Select(e => e.Text));
        // VB re-cases new text to the project's spelling of each name; a case-only edit can leave the hash unchanged
        IReadOnlyList<string> recased, hints;
        try { recased = VbCode.CaseChanges(typed, after.Text); } catch (Exception) { recased = []; }
        try { hints = VbCode.LiteralHints(typed); } catch (Exception) { hints = []; }
        return ToolResult.Ok(new
        {
            source = "ide", project, component, dirty = after.IsDirty, countOfLines = after.CountOfLines, hash = VbCode.Hash(after.Text),
            replaced = applied.Count == 1 ? applied[0] : null, edits = applied.Count > 1 ? applied : null,
            vbKeptSpelling = recased.Count > 0 ? recased : null,
            codeHints = hints.Count > 0 ? hints : null,
            note = "unsaved in the IDE; undo with Edit > Undo in VB6, save with project_save" + (recased.Count == 0 ? "" :
                $". VB kept its own spelling of {string.Join(", ", recased)}: it uses one spelling per name in the whole project, so a case-only change doesn't take" +
                (VbCode.Hash(after.Text) == VbCode.Hash(module.Text) ? " (the module is unchanged)" : "")),
        });
    }

    /// <summary>
    /// After an edit failed part-way through EditManyAsync: re-reads the module and compares it with the original, the
    /// original plus the edits that went in, and that plus the failed one (the add-in deletes the old lines, then inserts:
    /// an error in between leaves the old lines gone). Returns a sentence saying which it is.
    /// </summary>
    private static async Task<string> DescribeAfterFailureAsync(AddinClient client, string project, string component, string original, IReadOnlyList<LineEdit> applied, LineEdit failed)
    {
        string now;
        try { now = (await client.ReadAsync(project, component)).Text; }
        catch (Exception ex) { return $"Re-reading {component} failed too ({ex.Message}), so what was applied is unknown: read it with code_read before anything else."; }
        var withApplied = applied.Aggregate(original, (text, e) => VbCode.ReplaceLines(text, e.Start, e.Count, e.Text));
        string later = applied.Count == 0 ? "" : $"{applied.Count} later edit(s) went in before it. ";
        if (now == original) return "Nothing was changed. Re-read and retry.";
        if (now == withApplied) return later + "This edit was not applied. Re-read (new hash) and redo it.";
        if (now == VbCode.ReplaceLines(withApplied, failed.Start, failed.Count, failed.Text))
            return later + "This edit was applied too (the error came after it): don't redo it; re-read for the new hash.";
        return later + "The module is now in neither the old nor the new state (this edit may be half-applied: old lines deleted, new ones not inserted). " +
               $"Undo in VB6 (Edit > Undo) or read {component} with code_read and repair it; don't retry blindly.";
    }

    private void RequireDesignMode(InstanceInfo inst)
    {
        var w = _env.ReadWindows(inst.Pid);
        if (w.Mode != IdeMode.Design)
            throw new ToolException($"the IDE is in {w.Mode.ToString().ToLowerInvariant()} mode ('{w.Title}'); edits are only allowed in design mode");
        if (w.Dialogs.Count > 0)
            throw new ToolException($"the IDE is showing a dialog ('{w.Dialogs[0].Caption}'); answer it first (ide_dialog_respond)");
    }

    private static int ParseKind(string? kind) => (kind ?? "proc").ToLowerInvariant() switch
    {
        "proc" or "sub" or "function" => 0,
        "let" => 1,
        "set" => 2,
        "get" => 3,
        var k => throw new ToolException($"kind must be proc, get, let or set (got '{k}')"),
    };

    // ---------- save ----------

    /// <summary>How long project_save waits for the components to become clean (or a dialog).</summary>
    public TimeSpan SaveWait { get; init; } = TimeSpan.FromSeconds(20);

    private async Task<ToolResult> SaveAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireDesignMode(inst);
        var (client, project) = await ResolveProjectAsync(a);
        bool acceptOutside = a.Bool("outside_edits") == true;   // hard rule 9: explicit
        var prj = (await client.ProjectsAsync()).First(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase));
        if (a.Str("path") is { Length: > 0 } newPath)
        {
            if (prj.FileName.Length > 0) throw new ToolException($"{project} is already saved as {prj.FileName}; path is only for a project that has never been saved");
            if (a.List("components") is { Count: > 0 }) throw new ToolException("pass either path or components");
            return await SaveProjectAsAsync(inst, client, project, newPath, acceptOutside, ct);
        }
        // a never-saved project (File > New Project) would stop at a Save As dialog; unchanged, it isn't even dirty
        if (prj.FileName.Length == 0)
            throw new ToolException($"{project} has never been saved; pass path: the new .vbp under the workspace roots (its components are saved next to it)");
        if (a.List("components") is { Count: > 0 } only) return await SaveComponentsAsync(inst, client, project, only, acceptOutside, ct);
        var dirtyComps = (await client.ComponentsAsync(project)).Where(c => c.IsDirty).ToList();
        var dirtyBefore = dirtyComps.Select(c => c.Name).ToList();
        if (dirtyBefore.Count == 0 && !prj.IsDirty)
            return ToolResult.Ok(new { status = "nothing-to-save", project });
        var outside = await OutsideEditsAsync(inst, client, project, dirtyComps);
        if (outside.Count > 0 && !acceptOutside) throw OutsideEditsRefusal(outside);

        // VB rewrites the .vbp too when the project counts as changed (a production project, 2026-10-17: 17 reference lines, two
        // control versions bumped). Read it before and after, and say what changed.
        string? vbpBefore = ReadVbp(prj.FileName);
        try { await client.SaveAsync(project); }
        catch (AddinErrorException ex) when (ex.Code == 409) { throw new ToolException(ex.AddinMessage); }

        var deadline = DateTime.UtcNow + SaveWait;
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(300, ct);
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count > 0)
                return ToolResult.Ok(new { status = "dialog", project, note = "VB6 is showing a dialog while saving; it was not answered", dialogs = Dialogs(w) });
            IReadOnlyList<ComponentInfo> comps;
            try { comps = await client.ComponentsAsync(project); }
            catch (AddinBusyException) { continue; }
            if (comps.All(c => !c.IsDirty))
            {
                var vbp = VbpReport(vbpBefore, ReadVbp(prj.FileName));
                return ToolResult.Ok(new
                {
                    status = "saved", project, saved = dirtyBefore,
                    vbp = vbp is null ? null : new
                    {
                        changed = true, removed = vbp.Removed, added = vbp.Added,
                        versionChanges = vbp.VersionChanges.Select(v => new { kind = v.Kind, guid = v.Guid, name = v.Name, oldVersion = v.OldVersion, newVersion = v.NewVersion }),
                    },
                    note = vbp?.Note,
                    outsideEdits = outside.Count > 0 ? OutsideJson(outside) : null,
                });
            }
        }
        return ToolResult.Ok(new { status = "pending", project, note = $"components still unsaved after {SaveWait.TotalSeconds:0} s; check ide_status" });
    }

    /// <summary>
    /// Saves a never-saved project as a new .vbp (PRJSAVEAS, spike 43): its never-saved components go next to it under
    /// their names, then VBProject.SaveAs writes the .vbp. Nothing is overwritten.
    /// </summary>
    private async Task<ToolResult> SaveProjectAsAsync(InstanceInfo inst, AddinClient client, string project, string pathArg, bool acceptOutside, CancellationToken ct)
    {
        var path = NewVbpPath(pathArg);
        var folder = Path.GetDirectoryName(path)!;
        var comps = await client.ComponentsAsync(project);
        foreach (var c in comps.Where(c => c.FileName.Length == 0))
        {
            var target = Path.Combine(folder, c.Name + ComponentFileExtension(c.Type));
            if (File.Exists(target)) throw new ToolException($"{target} already exists, and {c.Name} would be saved there; choose another folder");
        }
        var outside = await OutsideEditsAsync(inst, client, project, comps.Where(c => c.IsDirty).ToList());
        if (outside.Count > 0 && !acceptOutside) throw OutsideEditsRefusal(outside);
        await QueueWithFolderAsync(folder, () => client.SaveProjectAsAsync(project, path));
        var saved = await WaitForProjectsAsync(inst, client, ps => ps.Any(p => SamePath(p.FileName, path)), ct);
        if (saved.Dialog is not null) return NewProjectDialog(inst, project, path);
        if (!saved.Done) return ToolResult.Ok(new { status = "pending", project, path, error = AddinLogSince(inst.Pid, saved.Started, "save project as failed") });
        return await SavedProjectResultAsync(client, saved.Projects.First(p => SamePath(p.FileName, path)), saved.Projects, "saved");
    }

    private static string ComponentFileExtension(int type) => type switch
    {
        1 => ".bas", 2 => ".cls", 5 or 6 => ".frm", 7 => ".pag", 8 => ".ctl", 9 => ".dob", _ => ".dsr",
    };

    /// <summary>
    /// Saves only the named components (SAVECOMP: VBComponent.SaveAs to its own file), leaving the .vbp alone: File > Save
    /// Project rewrote 17 reference lines of a production project's .vbp and bumped control versions (PLAN B8; verified 2026-10-17
    /// that SaveAs writes the component quietly and the .vbp keeps its bytes).
    /// </summary>
    private async Task<ToolResult> SaveComponentsAsync(InstanceInfo inst, AddinClient client, string project, IReadOnlyList<string> names, bool acceptOutside, CancellationToken ct)
    {
        var comps = await client.ComponentsAsync(project);
        var targets = names.Select(n => comps.FirstOrDefault(c => c.Name.Equals(n, StringComparison.OrdinalIgnoreCase))
                                         ?? throw new ToolException($"component not found: {project}.{n}")).ToList();
        if (targets.FirstOrDefault(c => c.FileName.Length == 0) is { } unsaved)
            throw new ToolException($"{unsaved.Name} has never been saved, so it has no file yet; save it with project_save (without components)");
        var dirty = targets.Where(c => c.IsDirty).ToList();
        if (dirty.Count == 0) return ToolResult.Ok(new { status = "nothing-to-save", project, components = names });
        var outside = await OutsideEditsAsync(inst, client, project, dirty);
        if (outside.Count > 0 && !acceptOutside) throw OutsideEditsRefusal(outside);
        foreach (var c in dirty) await client.SaveComponentAsync(project, c.Name);
        var deadline = DateTime.UtcNow + SaveWait;
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(300, ct);
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count > 0) return ToolResult.Ok(new { status = "dialog", project, note = "VB6 is showing a dialog while saving; it was not answered", dialogs = Dialogs(w) });
            IReadOnlyList<ComponentInfo> now;
            try { now = await client.ComponentsAsync(project); } catch (AddinBusyException) { continue; }
            if (dirty.All(d => now.Any(n => n.Name == d.Name && !n.IsDirty)))
                return ToolResult.Ok(new { status = "saved", project, saved = dirty.Select(d => d.Name), note = "only these components; the .vbp wasn't written", outsideEdits = outside.Count > 0 ? OutsideJson(outside) : null });
        }
        return ToolResult.Ok(new { status = "pending", project, note = $"still unsaved after {SaveWait.TotalSeconds:0} s; check ide_status" });
    }

    /// <summary>
    /// Code changes in <paramref name="comps"/> that no tool made (see <see cref="ToolEditLedger"/>): measured against the
    /// last tool edit, or against the file on disk for a component no tool edited. Never-saved components are skipped
    /// (nothing to compare with); so is one that can't be read, since this only adds a check.
    /// </summary>
    private async Task<List<OutsideEdit>> OutsideEditsAsync(InstanceInfo inst, AddinClient client, string project, IEnumerable<ComponentInfo> comps)
    {
        var found = new List<OutsideEdit>();
        foreach (var c in comps)
        {
            try
            {
                var now = (await client.ReadAsync(project, c.Name)).Text;
                OutsideEdit? o;
                if (Ledger.Last(inst.Pid, project, c.Name) is { } last) o = ToolEditLedger.Compare(c.Name, last, now, "last vb6-mcp edit");
                else if (c.FileName.Length > 0 && File.Exists(c.FileName)) o = ToolEditLedger.Compare(c.Name, VbSourceFile.Load(c.FileName).Code, now, "disk");
                else continue;
                if (o is not null) found.Add(o);
            }
            catch (Exception ex) when (ex is AddinErrorException or AddinBusyException or IOException or UnauthorizedAccessException) { }
        }
        return found;
    }

    private static object OutsideJson(IReadOnlyList<OutsideEdit> o) =>
        o.Select(e => new { component = e.Component, since = e.Since, changedLines = e.ChangedLines, diff = e.Diff, moreDiffLines = e.MoreDiffLines });

    private static ToolException OutsideEditsRefusal(IReadOnlyList<OutsideEdit> o) => new(
        "not saved: code changed in the IDE that no vb6-mcp tool edited (stray keystrokes reaching a code pane, or the user's own typing). " +
        "Check each diff: keep it with outside_edits=true, or undo it in VB6 (Edit > Undo) or with code_replace_lines.\n" +
        string.Join("\n", o.Select(e => $"{e.Component} ({e.ChangedLines} changed line(s) since {e.Since}):\n{e.Diff}" +
                                          (e.MoreDiffLines is int m ? $"\n... {m} more diff lines (sync_diff component={e.Component})" : ""))));

    private static string? ReadVbp(string path)
    {
        try { return path.Length > 0 && File.Exists(path) ? Vb1252.Encoding.GetString(File.ReadAllBytes(path)) : null; }
        catch (IOException) { return null; }
    }

    /// <summary>The .vbp lines a save changed, or null when it didn't change (or couldn't be read).</summary>
    private static VbpSaveReport? VbpReport(string? before, string? after)
    {
        if (before is null || after is null) return null;
        var c = VbpDiff.Compare(before, after);
        if (!c.Any) return null;
        string? note = c.Versions.Count == 0 ? null
            : "VB changed the version of " + string.Join(", ", c.Versions.Select(v => $"{v.Name} ({v.Kind}) {v.OldVersion} to {v.NewVersion}")) +
              " in the .vbp: it writes the version that is registered on this machine. Check this is wanted before committing the .vbp.";
        return new VbpSaveReport(true, c.Removed, c.Added, c.Versions, note);
    }

    private sealed record VbpSaveReport(bool Changed, IReadOnlyList<string> Removed, IReadOnlyList<string> Added, IReadOnlyList<VbpVersionChange> VersionChanges, string? Note);

    // ---------- dialogs ----------

    private async Task<ToolResult> DialogRespondAsync(JsonObject a, CancellationToken ct)
    {
        // A VB6 that is still starting can show a dialog before the add-in has loaded (Phase 2b live
        // test: another add-in's error box). Then there is no add-in to ask; the watchdog can click.
        InstanceInfo? inst = null; int targetPid;
        if (a.Int("pid") is int p && _env.Scan().Live.All(i => i.Pid != p) && Discovery.IsVb6Process(p)) targetPid = p;
        else { inst = ResolveInstance(a); targetPid = inst.Pid; }
        var button = a.Req("button").Replace("&", "", StringComparison.Ordinal);
        var dialogs = _env.ReadWindows(targetPid).Dialogs;
        if (dialogs.Count == 0) throw new ToolException("the IDE isn't showing a dialog");
        var caption = a.Str("dialog");
        var candidates = caption is null ? dialogs : dialogs.Where(d => d.Caption.Equals(caption, StringComparison.OrdinalIgnoreCase)).ToList();
        if (candidates.Count == 0) throw new ToolException($"no dialog captioned '{caption}'. Open: {string.Join(", ", dialogs.Select(d => d.Caption))}");
        if (candidates.Count > 1) throw new ToolException($"several dialogs are open ({string.Join(", ", candidates.Select(d => d.Caption))}); pass dialog");
        var dlg = candidates[0];
        var btn = dlg.Buttons.FirstOrDefault(b => b.Caption.Equals(button, StringComparison.OrdinalIgnoreCase))
                  ?? throw new ToolException($"dialog '{dlg.Caption}' has no button '{button}'. Buttons: {string.Join(", ", dlg.Buttons.Select(b => b.Caption))}");

        // Spike 41: a VB6 that isn't elevated runs at the server's own level, so the server can post the click itself.
        // Before the add-in has loaded there is no watchdog yet either (the add-in starts it), so another add-in's
        // startup error ("Data View") could only be answered by hand.
        bool ownLevel = Vb6Install.IsElevated(targetPid) == false;
        string ClickHere(long hwnd, string why)
        {
            Audit.Write("ide_dialog_respond", targetPid, $"server BM_CLICK {hwnd:X} ({why})");
            LocalInput.Click(hwnd);
            return $"server ({why}; VB6 isn't elevated)";
        }
        async Task<string> Click(long hwnd)
        {
            if (inst is null)
            {
                if (await PressViaWatchdogAsync(hwnd, targetPid)) return "watchdog (no add-in yet)";
                if (ownLevel) return ClickHere(hwnd, "no add-in yet");
                throw new ToolException($"VB6 {targetPid} has no add-in loaded yet and " + NoWatchdog);
            }
            try { await _env.Client(inst).ClickAsync(hwnd); return "add-in"; }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException)
            {
                // The add-in shares the IDE thread; if it can't answer, the elevated watchdog can still click (PLAN §3.1)
                if (await PressViaWatchdogAsync(hwnd, targetPid)) return "watchdog";
                if (ownLevel) return ClickHere(hwnd, "the add-in didn't answer");
                throw new ToolException($"the add-in didn't answer ({ex.Message}) and " + NoWatchdog);
            }
        }
        string via = await Click(btn.Hwnd);

        // Verify the effect: a dropped click looks like success (spike 10)
        await Task.Delay(500, ct);
        var after = _env.ReadWindows(targetPid);
        bool gone = after.Dialogs.All(d => d.Hwnd != dlg.Hwnd);
        // Improvement notes 11: the first click on "Continue Loading Project?" sometimes didn't take, and the same call a
        // minute later did (a dialog that had only just appeared, or an IDE still busy loading). Give it two more seconds
        // to close, then press once more: it is the same window, so the same answer.
        bool retried = false;
        for (var until = DateTime.UtcNow.AddSeconds(2); !gone && DateTime.UtcNow < until; gone = after.Dialogs.All(d => d.Hwnd != dlg.Hwnd))
        {
            await Task.Delay(250, ct);
            after = _env.ReadWindows(targetPid);
        }
        if (!gone && after.Dialogs.FirstOrDefault(d => d.Hwnd == dlg.Hwnd)?.Buttons.FirstOrDefault(b => b.Hwnd == btn.Hwnd) is { } again)
        {
            via = await Click(again.Hwnd);
            retried = true;
            for (var until = DateTime.UtcNow.AddSeconds(2); DateTime.UtcNow < until && !gone; gone = after.Dialogs.All(d => d.Hwnd != dlg.Hwnd))
            {
                await Task.Delay(250, ct);
                after = _env.ReadWindows(targetPid);
            }
        }

        // repeat: the same message for form after form (a missing reference or control: one "Errors during load" per form,
        // spike 36; the owner: "click ok until it stops"). Keep pressing the same button while the next dialog is the same
        // kind of message (its text with quoted paths blanked out), then sum up the load logs it named.
        if (a.Bool("repeat") == true && gone)
        {
            static string Kind(string text) => Regex.Replace(text, "'[^']*'", "'…'");
            var kind = Kind(string.Join("\n", dlg.Texts));
            var texts = new List<string> { string.Join("\n", dlg.Texts) };
            int clicks = 1;
            while (clicks < MaxRepeatClicks)
            {
                // the next form loads before its message appears: wait up to 5 s for it
                DialogInfo? next = null;
                for (var until = DateTime.UtcNow.AddSeconds(5); DateTime.UtcNow < until; await Task.Delay(250, ct))
                {
                    after = _env.ReadWindows(targetPid);
                    if (after.Dialogs.Count > 0) { next = after.Dialogs[0]; break; }
                }
                if (next is null || Kind(string.Join("\n", next.Texts)) != kind || next.Buttons.FirstOrDefault(b => b.Caption.Equals(button, StringComparison.OrdinalIgnoreCase)) is not { } nb)
                    break;
                await Click(nb.Hwnd);
                texts.Add(string.Join("\n", next.Texts));
                clicks++;
                await Task.Delay(300, ct);
            }
            after = _env.ReadWindows(targetPid);
            var logs = texts.Select(Core.Files.LoadLogs.LogOf).OfType<string>().ToList();
            return ToolResult.Ok(new
            {
                clicked = btn.Caption, dialog = dlg.Caption, repeated = clicks, stoppedBecause = clicks >= MaxRepeatClicks ? $"the limit of {MaxRepeatClicks} clicks" : after.Dialogs.Count > 0 ? "a different dialog" : "no more dialogs",
                distinctMessages = texts.Distinct().Count(),
                missing = logs.Count > 0 ? Core.Files.LoadLogs.Summarize(logs).Select(p => new { @class = p.Class, reason = p.Reason, controls = p.Controls, forms = p.Forms.Count, formNames = p.Forms.Take(10) }) : null,
                dialogsNow = Dialogs(after), via,
                note = logs.Count > 0 ? "missing lists what the load logs say is behind the messages; register the component (com_search, com_register) or fix the reference (refs_fix_broken), then reopen the project. Saving now would drop the controls that didn't load" : null,
            });
        }
        // A run_start that stopped at a dialog (SourceSafe, Debugging options) is still the agent's run once it starts
        if (_pendingStarts.Contains(targetPid))
        {
            if (after.Mode is IdeMode.Run or IdeMode.Break) { _runsWeStarted.Add(targetPid); _pendingStarts.Remove(targetPid); }
            else if (after.Dialogs.Count == 0) _pendingStarts.Remove(targetPid);
        }
        return ToolResult.Ok(new
        {
            clicked = btn.Caption, dialog = dlg.Caption, text = string.Join("\n", dlg.Texts), dialogClosed = gone, retried = retried ? true : (bool?)null,
            dialogsNow = Dialogs(after), via,
            note = gone ? null : "the dialog is still open after two presses. It may be waiting behind another window of the IDE (a dialog the user opened), " +
                                 "or the IDE is busy: check dialogsNow and ide_status, then try again",
        });
    }
}
