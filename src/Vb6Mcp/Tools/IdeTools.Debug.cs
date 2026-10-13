using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 2: run and debug. VB6 has no debugger API, so these combine IDE commands (command-bar ids,
/// deferred in the add-in: hard rule 14) with window reads (mode from the title, dialogs) and the
/// Immediate window (spikes 3, 5, 6/8). Every command observes its outcome, because a command that
/// was sent isn't proof that it worked (spike 10). Documented in docs/tools.md.
/// </summary>
public sealed partial class IdeTools
{
    // Command-bar ids from the spike 5 dump
    private const string CmdStart = "186", CmdBreak = "189", CmdEnd = "228", CmdStepInto = "188", CmdStepOver = "194",
                         CmdStepOut = "2559", CmdRunToCursor = "1811", CmdToggleBp = "51", CmdClearBps = "579";

    /// <summary>IDE pids where this server started the current run (hard rule 11: ending other runs needs force).</summary>
    private readonly HashSet<int> _runsWeStarted = [];
    /// <summary>Breakpoints set through this server, per pid. VBIDE can't list breakpoints, so this can drift.</summary>
    private readonly Dictionary<int, HashSet<(string Project, string Component, int Line)>> _breakpoints = [];

    /// <summary>How long run/step commands wait for the IDE to reach the expected state.</summary>
    public TimeSpan CommandWait { get; init; } = TimeSpan.FromSeconds(10);

    private IReadOnlyList<Tool> DebugTools() =>
    [
        new("run_start",
            "Start the program in the IDE (F5), or with full_compile=true as 'Start With Full Compile' (Ctrl+F5, compiles everything first). " +
            "Design mode only. Reports the resulting mode and any dialogs (compile errors, MsgBoxes, runtime errors).",
            Schema.Object(("full_compile", "boolean", "Compile all code before starting.", false), ("pid", "integer", PidDesc, false)),
            (a, ct) => RunStartAsync(a, ct)),
        new("run_break",
            "Break into the running program. Run mode only. Uses the IDE's Break command; if the program keeps the IDE busy (a loop), " +
            "the watchdog injects Ctrl+Break instead (VB6 comes to the front for a moment). The result says which was used.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => RunBreakAsync(a, ct)),
        new("run_continue",
            "Continue the program from break mode (F5).",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => SimpleCommandAsync(a, CmdStart, [IdeMode.Break], [IdeMode.Run, IdeMode.Design], ct)),
        new("run_end",
            "End the running program and return to design mode. Ending a run the agent didn't start needs force=true " +
            "(it may be the user's debugging session).",
            Schema.Object(("force", "boolean", "End even though this server didn't start the run.", false), ("pid", "integer", PidDesc, false)),
            (a, ct) => RunEndAsync(a, ct)),
        new("step_into", "Step into the next statement (F8). Break mode, or design mode to start at the first statement.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => SimpleCommandAsync(a, CmdStepInto, [IdeMode.Break, IdeMode.Design], [IdeMode.Break, IdeMode.Run, IdeMode.Design], ct)),
        new("step_over", "Step over the next statement (Shift+F8). Break mode only.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => SimpleCommandAsync(a, CmdStepOver, [IdeMode.Break], [IdeMode.Break, IdeMode.Run, IdeMode.Design], ct)),
        new("step_out", "Step out of the current procedure (Ctrl+Shift+F8). Break mode only.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => SimpleCommandAsync(a, CmdStepOut, [IdeMode.Break], [IdeMode.Break, IdeMode.Run, IdeMode.Design], ct)),
        new("run_to_cursor", "Run until the given line (Ctrl+F8). Break or design mode.",
            Schema.Object(("component", "string", ComponentDesc, true), ("line", "integer", "Line to run to (1-based).", true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => RunToCursorAsync(a, ct)),
        new("bp_toggle",
            "Toggle a breakpoint on a line (F9). Design or break mode. VBIDE can't list breakpoints, so the server tracks the ones it " +
            "toggled; breakpoints the user sets by hand aren't known.",
            Schema.Object(("component", "string", ComponentDesc, true), ("line", "integer", "Line (1-based).", true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => BreakpointToggleAsync(a, ct)),
        new("bp_clear_all", "Clear all breakpoints (Ctrl+Shift+F9), including ones the user set.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => BreakpointClearAsync(a, ct)),
        new("debug_where",
            "Where the IDE is: the active code pane's component, line and procedure (in break mode, the current statement), plus that line's text.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, _) => WhereAsync(a)),
        new("debug_eval",
            "Evaluate an expression in the Immediate window ('? expr') and return the result. Break mode (current scope) or design mode " +
            "(module-level/global only; it may run code, e.g. a function call).",
            Schema.Object(("expression", "string", "Expression, e.g. x or m_counter.Value * 2.", true), ("pid", "integer", PidDesc, false)),
            (a, ct) => EvalAsync(a, ct)),
        new("debug_immediate",
            "Read the Immediate window (Debug.Print output and earlier evaluations), optionally only the last N lines.",
            Schema.Object(("last", "integer", "Only the last N lines.", false), ("pid", "integer", PidDesc, false)),
            (a, _) => ImmediateReadAsync(a)),
        new("syntax_check",
            "Compile the code as it is in the IDE (unsaved edits included) via 'Start With Full Compile'. On a compile error, reports the " +
            "message and location and closes the error dialog. If it compiles, the program starts and is ended again right away " +
            "(startup code like Form_Load runs briefly). Design mode only.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => SyntaxCheckAsync(a, ct)),
    ];

    // ---------- run control ----------

    private async Task<ToolResult> RunStartAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireMode(inst, IdeMode.Design);
        bool full = a["full_compile"]?.GetValue<bool>() == true;
        await _env.Client(inst).CommandAsync(full ? "fullcompile" : CmdStart);
        var w = await WaitForAsync(inst, s => s.Mode is IdeMode.Run or IdeMode.Break || s.Dialogs.Count > 0, ct);
        if (w.Mode is IdeMode.Run or IdeMode.Break) _runsWeStarted.Add(inst.Pid);
        return ToolResult.Ok(State("started", w));
    }

    private async Task<ToolResult> RunEndAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireMode(inst, IdeMode.Run, IdeMode.Break);
        if (!_runsWeStarted.Contains(inst.Pid) && a["force"]?.GetValue<bool>() != true)
            throw new ToolException("this run wasn't started by the agent (it may be the user's debugging session); pass force=true to end it anyway");
        await _env.Client(inst).CommandAsync(CmdEnd);
        var w = await WaitForAsync(inst, s => s.Mode == IdeMode.Design, ct);
        if (w.Mode == IdeMode.Design) _runsWeStarted.Remove(inst.Pid);
        return ToolResult.Ok(State("ended", w));
    }

    private async Task<ToolResult> SimpleCommandAsync(JsonObject a, string cmd, IdeMode[] allowed, IdeMode[] expected, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        var before = RequireMode(inst, allowed);
        await _env.Client(inst).CommandAsync(cmd);
        // A step stays in break mode: when the starting mode is also an expected one, the short minimum
        // wait is enough; otherwise wait for the mode to change (e.g. break -> run/design on continue)
        bool sameModeOk = expected.Contains(before.Mode);
        var w = await WaitForAsync(inst, s => s.Dialogs.Count > 0 || (expected.Contains(s.Mode) && (sameModeOk || s.Mode != before.Mode)),
                                   ct, minWait: TimeSpan.FromMilliseconds(400));
        if (w.Mode is IdeMode.Run or IdeMode.Break && before.Mode == IdeMode.Design) _runsWeStarted.Add(inst.Pid);
        if (w.Mode == IdeMode.Design) _runsWeStarted.Remove(inst.Pid);
        return ToolResult.Ok(State("done", w));
    }

    private async Task<ToolResult> RunToCursorAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        var before = RequireMode(inst, IdeMode.Break, IdeMode.Design);
        var (client, project) = await ResolveProjectAsync(a);
        var component = a.Req("component");
        int line = a.Int("line") ?? throw new ToolException("missing required argument 'line'");
        await client.SetSelectionAsync(project, component, line);
        await client.CommandAsync(CmdRunToCursor);
        var w = await WaitForAsync(inst, s => s.Dialogs.Count > 0 || s.Mode == IdeMode.Break, ct, minWait: TimeSpan.FromMilliseconds(400));
        if (before.Mode == IdeMode.Design && w.Mode != IdeMode.Design) _runsWeStarted.Add(inst.Pid);
        return ToolResult.Ok(State("done", w));
    }

    // ---------- breakpoints ----------

    private async Task<ToolResult> BreakpointToggleAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireMode(inst, IdeMode.Design, IdeMode.Break);
        var (client, project) = await ResolveProjectAsync(a);
        var component = a.Req("component");
        int line = a.Int("line") ?? throw new ToolException("missing required argument 'line'");
        await client.SetSelectionAsync(project, component, line);
        await client.CommandAsync(CmdToggleBp);
        await Task.Delay(300, ct);
        var set = _breakpoints.TryGetValue(inst.Pid, out var s) ? s : _breakpoints[inst.Pid] = [];
        var key = (project, component, line);
        bool nowOn = set.Add(key) || !set.Remove(key);
        return ToolResult.Ok(new
        {
            project, component, line, breakpoint = nowOn ? "set" : "cleared",
            knownBreakpoints = set.Select(b => new { b.Project, b.Component, b.Line }),
            note = "based on the toggles made through this server; breakpoints set by hand in the IDE aren't known (VBIDE can't list them)",
        });
    }

    private async Task<ToolResult> BreakpointClearAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireMode(inst, IdeMode.Design, IdeMode.Break);
        await _env.Client(inst).CommandAsync(CmdClearBps);
        await Task.Delay(300, ct);
        _breakpoints.Remove(inst.Pid);
        return ToolResult.Ok(new { cleared = true });
    }

    // ---------- inspection ----------

    private async Task<ToolResult> WhereAsync(JsonObject a)
    {
        var inst = ResolveInstance(a);
        var w = _env.ReadWindows(inst.Pid);
        var client = _env.Client(inst);
        var where = await client.WhereAsync();
        string? lineText = null;
        try { lineText = (await client.ReadAsync(where.Project, where.Component, where.StartLine, 1)).Text; } catch (AddinErrorException) { }
        return ToolResult.Ok(new
        {
            mode = Mode(w.Mode), where.Project, where.Component, line = where.StartLine, column = where.StartColumn,
            procedure = where.Procedure.Length == 0 ? null : where.Procedure, lineText,
        });
    }

    private async Task<ToolResult> EvalAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireMode(inst, IdeMode.Break, IdeMode.Design);
        var expr = a.Req("expression").Trim();
        if (expr.Contains('\n') || expr.Contains('\r')) throw new ToolException("the expression must be a single line");
        var (status, result, w) = await EvaluateAsync(inst, expr, "debug_eval", ct);
        return status switch
        {
            "ok" => ToolResult.Ok(new { expression = expr, status, result }),
            "dialog" => ToolResult.Ok(new { expression = expr, status, note = "evaluating showed a dialog (e.g. a runtime error); it was not answered", dialogs = Dialogs(w!) }),
            _ => ToolResult.Ok(new { expression = expr, status, note = $"no result after {CommandWait.TotalSeconds:0} s; check debug_immediate and ide_status" }),
        };
    }

    /// <summary>
    /// Types "? expr" into the Immediate window and waits for the new output (spike 5; the read
    /// sequence is in docs/protocol.md). Status: ok (result set), dialog (window state returned), pending.
    /// </summary>
    private async Task<(string Status, string? Result, IdeWindowState? Window)> EvaluateAsync(InstanceInfo inst, string expr, string auditAs, CancellationToken ct)
    {
        long imm = _env.FindImmediate(inst.Pid);
        if (imm == 0) throw new ToolException("the Immediate window wasn't found");
        var client = _env.Client(inst);
        string before = await client.ImmediateReadAsync(imm);
        string command = "? " + expr;
        Audit.Write(auditAs, inst.Pid, command);
        await client.ImmediateAsync(imm, command);

        var deadline = DateTime.UtcNow + CommandWait;
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(300, ct);
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count > 0) return ("dialog", null, w);
            string now;
            try { now = await client.ImmediateReadAsync(imm); }
            catch (AddinBusyException) { continue; }                 // the evaluation is running code
            var result = ResultAfter(now, before, command);
            if (result != null) return ("ok", result, w);
        }
        return ("pending", null, null);
    }

    /// <summary>
    /// The lines the Immediate window printed after our command line. Null while nothing new has
    /// appeared. The window keeps old content, so only text after the last occurrence counts.
    /// </summary>
    public static string? ResultAfter(string now, string before, string command)
    {
        if (now.Length <= before.Length && now == before) return null;
        int at = now.LastIndexOf(command, StringComparison.Ordinal);
        if (at < 0) return null;
        var rest = now[(at + command.Length)..].TrimStart('\r', '\n');
        var lines = rest.Split("\r\n").Select(l => l.TrimEnd()).ToList();
        while (lines.Count > 0 && lines[^1].Length == 0) lines.RemoveAt(lines.Count - 1);
        return lines.Count == 0 ? null : string.Join("\n", lines.Select(l => l.Trim()));
    }

    private async Task<ToolResult> ImmediateReadAsync(JsonObject a)
    {
        var inst = ResolveInstance(a);
        long imm = _env.FindImmediate(inst.Pid);
        if (imm == 0) throw new ToolException("the Immediate window wasn't found");
        var text = await _env.Client(inst).ImmediateReadAsync(imm);
        var lines = text.Split("\r\n").ToList();
        while (lines.Count > 0 && lines[^1].Trim().Length == 0) lines.RemoveAt(lines.Count - 1);
        if (a.Int("last") is int n && n < lines.Count) lines = lines.Skip(lines.Count - n).ToList();
        return ToolResult.Ok(new { lines = lines.Count, text = string.Join("\r\n", lines) });
    }

    // ---------- syntax check ----------

    private async Task<ToolResult> SyntaxCheckAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireMode(inst, IdeMode.Design);
        var client = _env.Client(inst);
        await client.CommandAsync("fullcompile");
        var w = await WaitForAsync(inst, s => s.Mode is IdeMode.Run or IdeMode.Break || s.Dialogs.Count > 0, ct, TimeSpan.FromSeconds(60));

        var compileError = w.Dialogs.FirstOrDefault(d => d.Texts.Any(t => t.StartsWith("Compile error", StringComparison.OrdinalIgnoreCase)));
        if (compileError != null)
        {
            string message = string.Join(" ", compileError.Texts).Replace("\r\n", " ").Replace("Compile error:", "").Trim();
            var ok = compileError.Buttons.FirstOrDefault(b => b.Caption.Equals("OK", StringComparison.OrdinalIgnoreCase));
            if (ok != null) { await client.ClickAsync(ok.Hwnd); await Task.Delay(400, ct); }
            Where? where = null;
            try { where = await client.WhereAsync(); } catch (Exception ex) when (ex is AddinErrorException or AddinBusyException) { }
            string? lineText = null;
            if (where != null)
                try { lineText = (await client.ReadAsync(where.Project, where.Component, where.StartLine, 1)).Text; } catch (AddinErrorException) { }
            return ToolResult.Ok(new
            {
                compiles = false, message,
                location = where is null ? null : new { where.Project, where.Component, line = where.StartLine, column = where.StartColumn, endColumn = where.EndColumn, procedure = where.Procedure, lineText },
            });
        }
        if (w.Mode is IdeMode.Run or IdeMode.Break)
        {
            if (w.Dialogs.Count > 0)
                return ToolResult.Ok(new
                {
                    compiles = true, status = "running-with-dialog",
                    note = "it compiled and the program started, but it shows a dialog (e.g. a MsgBox in Form_Load); answer it, then run_end",
                    dialogs = Dialogs(w),
                });
            _runsWeStarted.Add(inst.Pid);
            await client.CommandAsync(CmdEnd);
            var after = await WaitForAsync(inst, s => s.Mode == IdeMode.Design, ct);
            if (after.Mode == IdeMode.Design) _runsWeStarted.Remove(inst.Pid);
            return ToolResult.Ok(new { compiles = true, status = after.Mode == IdeMode.Design ? "ended" : "still-running", mode = Mode(after.Mode) });
        }
        return ToolResult.Ok(new { compiles = (bool?)null, status = "unknown", note = "neither a compile error nor a run appeared", state = State("", w) });
    }

    // ---------- helpers ----------

    private IdeWindowState RequireMode(InstanceInfo inst, params IdeMode[] allowed)
    {
        var w = _env.ReadWindows(inst.Pid);
        if (!allowed.Contains(w.Mode))
            throw new ToolException($"the IDE is in {Mode(w.Mode)} mode; this needs {string.Join(" or ", allowed.Select(Mode))} mode");
        if (w.Dialogs.Count > 0)
            throw new ToolException($"the IDE is showing a dialog ('{w.Dialogs[0].Caption}': {string.Join(" ", w.Dialogs[0].Texts)}); answer it first (ide_dialog_respond)");
        return w;
    }

    private async Task<IdeWindowState> WaitForAsync(InstanceInfo inst, Func<IdeWindowState, bool> done, CancellationToken ct, TimeSpan? timeout = null, TimeSpan? minWait = null)
    {
        var start = DateTime.UtcNow;
        var deadline = start + (timeout ?? CommandWait);
        IdeWindowState w;
        do
        {
            await Task.Delay(250, ct);
            w = _env.ReadWindows(inst.Pid);
            if (done(w) && DateTime.UtcNow - start >= (minWait ?? TimeSpan.Zero)) return w;
        } while (DateTime.UtcNow < deadline);
        return w;
    }

    private static object State(string status, IdeWindowState w) => new
    {
        status, mode = Mode(w.Mode), responds = w.Responds, title = w.Title, dialogs = Dialogs(w),
        note = !w.Responds ? "the IDE isn't responding (program busy in a loop?)"
             : w.Dialogs.Any(d => RuntimeError(d) != null) ? "a runtime error occurred: debug_error (action debug shows the failing line)"
             : w.Dialogs.Count > 0 ? "a dialog is open; answer it with ide_dialog_respond" : null,
    };
}
