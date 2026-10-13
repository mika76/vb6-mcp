using System.Text.Json.Nodes;
using Vb6Mcp.Core;
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
    private const string HashDesc = "The module hash from code_read or proc_get. The edit is refused if the module changed since.";

    private IReadOnlyList<Tool> EditTools() =>
    [
        new("code_replace_lines",
            "Replace lines of a component's code in the IDE (unsaved, undoable in VB6). count 0 inserts before start_line; " +
            "empty new_text deletes. Needs expected_hash from code_read/proc_get; design mode only. Returns the new hash.",
            Schema.Object(("component", "string", ComponentDesc, true), ("start_line", "integer", "First line to replace (1-based).", true),
                          ("count", "integer", "Number of lines to replace (0 = insert).", true), ("new_text", "string", "Replacement lines (CR LF or LF).", true),
                          ("expected_hash", "string", HashDesc, true), ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => ReplaceLinesAsync(a)),
        new("code_insert",
            "Insert lines into a component before line `line` (use line = countOfLines + 1 to append). Same rules as code_replace_lines.",
            Schema.Object(("component", "string", ComponentDesc, true), ("line", "integer", "Insert before this line (1-based).", true),
                          ("text", "string", "Lines to insert.", true), ("expected_hash", "string", HashDesc, true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => InsertAsync(a)),
        new("proc_replace",
            "Replace a whole procedure or property (from its first comment line to End Sub/Function/Property) with new_text, " +
            "exactly the range proc_get returns. Same rules as code_replace_lines.",
            Schema.Object(("component", "string", ComponentDesc, true), ("name", "string", "Procedure or property name.", true),
                          ("new_text", "string", "The complete new procedure, including comments above it.", true),
                          ("expected_hash", "string", HashDesc, true), ("kind", "string", "proc (default), get, let or set.", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => ProcReplaceAsync(a)),
        new("project_save",
            "Save a project like File > Save Project (all unsaved components and the .vbp). Explicit only: edits are never " +
            "saved automatically. Refused for a never-saved project. Dialogs shown while saving are reported, not answered.",
            Schema.Object(("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => SaveAsync(a, ct)),
        new("ide_dialog_respond",
            "Click a button in a dialog the IDE is showing (see ide_status). The button must be named explicitly; " +
            "pass dialog (its caption) when several dialogs are open.",
            Schema.Object(("button", "string", "Button caption, e.g. OK, Yes, No, Cancel.", true),
                          ("dialog", "string", "Dialog caption; needed when more than one dialog is open.", false),
                          ("pid", "integer", PidDesc, false)),
            (a, ct) => DialogRespondAsync(a, ct)),
    ];

    // ---------- edits ----------

    private async Task<ToolResult> ReplaceLinesAsync(JsonObject a)
    {
        int start = a.Int("start_line") ?? throw new ToolException("missing required argument 'start_line'");
        int count = a.Int("count") ?? throw new ToolException("missing required argument 'count'");
        return await EditAsync(a, a.Str("new_text") ?? throw new ToolException("missing required argument 'new_text'"),
            (_, _) => Task.FromResult((start, count)));
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

    /// <summary>Shared edit path: resolve, check mode, validate text, check hash, replace with the old text as guard.</summary>
    private async Task<ToolResult> EditAsync(JsonObject a, string newText, Func<AddinClient, (string Project, string Component), Task<(int Start, int Count)>> range)
    {
        string prepared;
        try { prepared = VbCode.PrepareForInsert(newText); }
        catch (ArgumentException ex) { throw new ToolException(ex.Message); }

        var inst = ResolveInstance(a);
        RequireDesignMode(inst);
        var (client, project) = await ResolveProjectAsync(a);
        var component = a.Req("component");
        var expected = a.Req("expected_hash");

        var module = await client.ReadAsync(project, component);
        var actual = VbCode.Hash(module.Text);
        if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
            throw new ToolException($"conflict: {component} changed since you read it (expected hash {expected}, now {actual}). Re-read it with code_read and redo the edit.");

        var (start, count) = await range(client, (project, component));
        string expectedOld = count > 0 ? VbCode.Slice(module.Text, start, count) : "";
        try
        {
            await client.ReplaceAsync(project, component, start, count, expectedOld, prepared);
        }
        catch (AddinErrorException ex) when (ex.Code == 409)
        {
            throw new ToolException($"conflict: the target lines of {component} changed while editing ({ex.AddinMessage}). Re-read and retry.");
        }
        var after = await client.ReadAsync(project, component);
        int inserted = prepared.Length == 0 ? 0 : prepared.Split("\r\n").Length;
        return ToolResult.Ok(new
        {
            source = "ide", project, component, dirty = after.IsDirty, countOfLines = after.CountOfLines, hash = VbCode.Hash(after.Text),
            replaced = new { startLine = start, oldLines = count, newLines = inserted },
            note = "unsaved in the IDE; undo with Edit > Undo in VB6, save with project_save",
        });
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
        var dirtyBefore = (await client.ComponentsAsync(project)).Where(c => c.IsDirty).Select(c => c.Name).ToList();
        var prj = (await client.ProjectsAsync()).First(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase));
        if (dirtyBefore.Count == 0 && !prj.IsDirty)
            return ToolResult.Ok(new { status = "nothing-to-save", project });

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
                return ToolResult.Ok(new { status = "saved", project, saved = dirtyBefore });
        }
        return ToolResult.Ok(new { status = "pending", project, note = $"components still unsaved after {SaveWait.TotalSeconds:0} s; check ide_status" });
    }

    // ---------- dialogs ----------

    private async Task<ToolResult> DialogRespondAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        var button = a.Req("button").Replace("&", "", StringComparison.Ordinal);
        var dialogs = _env.ReadWindows(inst.Pid).Dialogs;
        if (dialogs.Count == 0) throw new ToolException("the IDE isn't showing a dialog");
        var caption = a.Str("dialog");
        var candidates = caption is null ? dialogs : dialogs.Where(d => d.Caption.Equals(caption, StringComparison.OrdinalIgnoreCase)).ToList();
        if (candidates.Count == 0) throw new ToolException($"no dialog captioned '{caption}'. Open: {string.Join(", ", dialogs.Select(d => d.Caption))}");
        if (candidates.Count > 1) throw new ToolException($"several dialogs are open ({string.Join(", ", candidates.Select(d => d.Caption))}); pass dialog");
        var dlg = candidates[0];
        var btn = dlg.Buttons.FirstOrDefault(b => b.Caption.Equals(button, StringComparison.OrdinalIgnoreCase))
                  ?? throw new ToolException($"dialog '{dlg.Caption}' has no button '{button}'. Buttons: {string.Join(", ", dlg.Buttons.Select(b => b.Caption))}");

        await _env.Client(inst).ClickAsync(btn.Hwnd);

        // Verify the effect: a dropped click looks like success (spike 10)
        await Task.Delay(500, ct);
        var after = _env.ReadWindows(inst.Pid);
        bool gone = after.Dialogs.All(d => d.Hwnd != dlg.Hwnd);
        return ToolResult.Ok(new { clicked = btn.Caption, dialog = dlg.Caption, text = string.Join("\n", dlg.Texts), dialogClosed = gone, dialogsNow = Dialogs(after) });
    }
}
