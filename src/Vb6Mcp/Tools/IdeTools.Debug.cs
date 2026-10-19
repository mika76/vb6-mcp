using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
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

    /// <summary>IDEs whose current run this server started (handed to the next server by server_restart, notes 56).</summary>
    public IReadOnlyList<int> RunsStarted => [.. _runsWeStarted];

    /// <summary>Runs the previous server of this session started (from the handover), counted as this server's.</summary>
    public IEnumerable<int> InheritedRuns { init { foreach (var pid in value) _runsWeStarted.Add(pid); } }
    /// <summary>IDEs where run_start stopped at a dialog in design mode; the run counts as ours if it starts once the dialog is answered (spike 24 live test).</summary>
    private readonly HashSet<int> _pendingStarts = [];
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
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => StepAsync(a, CmdStepInto, [IdeMode.Break, IdeMode.Design], ct)),
        new("step_over", "Step over the next statement (Shift+F8). Break mode only.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => StepAsync(a, CmdStepOver, [IdeMode.Break], ct)),
        new("step_out", "Step out of the current procedure (Ctrl+Shift+F8). Break mode only.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => StepAsync(a, CmdStepOut, [IdeMode.Break], ct)),
        new("run_to_cursor", "Run until the given line (Ctrl+F8). Break or design mode.",
            Schema.Object(("component", "string", ComponentDesc, true), ("line", "integer", "Line to run to (1-based).", true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => RunToCursorAsync(a, ct)),
        new("bp_toggle",
            "Toggle a breakpoint on a line (F9). Design, break or run mode (while the program runs too). VBIDE can't list breakpoints, so the server tracks the ones it " +
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
            "(module-level/global only; it may run code, e.g. a function call). For reading values only: expressions with side effects (database writes, " +
            "Shell or other process starts, file, registry or network access, scripts, SendKeys) are refused; test through the program's own UI.",
            Schema.Object(("expression", "string", "Expression, e.g. x or m_counter.Value * 2.", true),
                          ("allow_side_effect", "boolean", "Ask the user to approve an expression with a side effect (only when really needed).", false),
                          ("pid", "integer", PidDesc, false)),
            (a, ct) => EvalAsync(a, ct)),
        new("debug_immediate",
            "Read the Immediate window (Debug.Print output and earlier evaluations), optionally only the last N lines.",
            Schema.Object(("last", "integer", "Only the last N lines.", false), ("pid", "integer", PidDesc, false)),
            (a, _) => ImmediateReadAsync(a)),
        new("syntax_check",
            "Compile the code as it is in the IDE (unsaved edits included) via 'Start With Full Compile'. On a compile error, reports the " +
            "message and location and closes the error dialog. If it compiles, the program starts and is ended again right away " +
            "(startup code like Form_Load runs briefly). Design mode only. Other conditional compilation constants: project_settings CondComp with reopen=true.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => SyntaxCheckAsync(a, ct)),
    ];

    // ---------- run control ----------

    private async Task<ToolResult> RunStartAsync(JsonObject a, CancellationToken ct)
    {
        var settled = await SettleReloadsAsync(ResolveInstance(a).Pid, ct);
        return (await RunStartCoreAsync(a, ct)).With("reloads", settled);
    }

    private async Task<ToolResult> RunStartCoreAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireMode(inst, IdeMode.Design);
        bool full = a["full_compile"]?.GetValue<bool>() == true;
        await RememberDirtyBeforeRunAsync(inst);   // spike 32: a run marks forms dirty
        var missingDlls = await MissingDllsAsync(inst);
        await _env.Client(inst).CommandAsync(full ? "fullcompile" : CmdStart);
        var w = await WaitForAsync(inst, s => s.Mode is IdeMode.Run or IdeMode.Break || s.Dialogs.Count > 0, ct);
        if (w.Mode is IdeMode.Run or IdeMode.Break) _runsWeStarted.Add(inst.Pid);
        else if (w.Dialogs.Count > 0) _pendingStarts.Add(inst.Pid);
        return ToolResult.Ok(State("started", w)).With("missingDlls", missingDlls);
    }

    /// <summary>Per .vbp (path, write time): when it was checked and what was missing. Walking the DLLs takes a moment, so repeated runs reuse it.</summary>
    private readonly Dictionary<(string Vbp, DateTime Written), (DateTime At, IReadOnlyList<(string Name, Core.Files.MissingDll Missing)> Missing)> _dllChecks = [];

    /// <summary>
    /// Improvement notes 55 / spike 39: a control or reference whose DLL needs a DLL that isn't installed. Loaded at
    /// start-up it fails with error 48/429; delay-loaded, it ends the IDE (0xc06d007e, unsaved work lost) the moment
    /// the program first uses it. The run still starts (that code path may never run), but the agent hears about it
    /// first. Null when nothing is missing or the check can't run.
    /// </summary>
    private async Task<JsonNode?> MissingDllsAsync(InstanceInfo inst)
    {
        try
        {
            var projects = await _env.Client(inst).ProjectsAsync();
            var appFolders = new List<string>();
            if (Core.Vb6Install.FindExe() is { } vb6) appFolders.Add(Path.GetDirectoryName(vb6)!);
            IReadOnlyList<Core.Files.RegisteredTypeLib>? registered = null;
            var found = new List<(string Project, string Name, Core.Files.MissingDll Missing)>();
            foreach (var p in projects.Where(p => File.Exists(p.FileName)))
            {
                var key = (Core.Files.PathKey.Normalize(p.FileName), File.GetLastWriteTimeUtc(p.FileName));
                if (!_dllChecks.TryGetValue(key, out var c) || DateTime.UtcNow - c.At > TimeSpan.FromMinutes(10))
                {
                    registered ??= Core.Files.TypeLibRegistry.All();
                    var folders = appFolders.Append(Path.GetDirectoryName(key.Item1)!).ToList();
                    c = (DateTime.UtcNow, Core.Files.ProjectPreflight.MissingDlls(Core.Files.VbpFile.Load(key.Item1), registered, folders));
                    _dllChecks[key] = c;
                }
                found.AddRange(c.Missing.Select(m => (p.Name, m.Name, m.Missing)));
            }
            if (found.Count == 0) return null;
            bool delay = found.Any(f => f.Missing.Delay);
            bool unsaved = projects.Any(p => p.IsDirty);
            return JsonSerializer.SerializeToNode(new
            {
                dlls = found.Take(10).Select(f => new { project = f.Project, component = f.Name, chain = string.Join(" -> ", f.Missing.Chain), delayLoaded = f.Missing.Delay }),
                more = found.Count > 10 ? found.Count - 10 : (int?)null,
                hint = (delay
                    ? "A delay-loaded DLL is missing: when the program first uses that component, the IDE process ends (exception 0xc06d007e, spike 39)" +
                      (unsaved ? ", and unsaved changes are lost: project_save before going near that code" : "") + ". "
                    : "Loading these components fails with error 48 or 429. ") +
                    "Install the missing DLLs or avoid the code that uses the component; com_diagnose project=... has the details.",
            }, McpServer.Json);
        }
        catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException or IOException or UnauthorizedAccessException) { return null; }
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
        return ToolResult.Ok(State("ended", w)).With("markedChangedByRun", MarkedJson(await MarkedChangedByRunAsync(inst)));
    }

    /// <summary>
    /// A step command, plus where it stopped (component, line, procedure, the line's text) when the IDE is in break mode:
    /// stepping without it took a debug_where after every step (benchmark T18).
    /// </summary>
    private async Task<ToolResult> StepAsync(JsonObject a, string cmd, IdeMode[] allowed, CancellationToken ct)
    {
        var result = await SimpleCommandAsync(a, cmd, allowed, [IdeMode.Break, IdeMode.Run, IdeMode.Design], ct);
        if (result.IsError) return result;
        var inst = ResolveInstance(a);
        if (_env.ReadWindows(inst.Pid).Mode != IdeMode.Break) return result;
        try
        {
            var client = _env.Client(inst);
            var where = await client.WhereAsync();
            string? lineText = null;
            try { lineText = (await client.ReadAsync(where.Project, where.Component, where.StartLine, 1)).Text; } catch (AddinErrorException) { }
            return result.With("at", JsonSerializer.SerializeToNode(new
            {
                component = where.Component, line = where.StartLine, procedure = where.Procedure.Length == 0 ? null : where.Procedure, lineText,
            }, McpServer.Json));
        }
        catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException) { return result; }
    }

    private async Task<ToolResult> SimpleCommandAsync(JsonObject a, string cmd, IdeMode[] allowed, IdeMode[] expected, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        var before = RequireMode(inst, allowed);
        if (before.Mode == IdeMode.Design) await RememberDirtyBeforeRunAsync(inst);   // a step from design mode starts a run
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
        if (before.Mode == IdeMode.Design) await RememberDirtyBeforeRunAsync(inst);
        await client.CommandAsync(CmdRunToCursor);
        var w = await WaitForAsync(inst, s => s.Dialogs.Count > 0 || s.Mode == IdeMode.Break, ct, minWait: TimeSpan.FromMilliseconds(400));
        if (before.Mode == IdeMode.Design && w.Mode != IdeMode.Design) _runsWeStarted.Add(inst.Pid);
        return ToolResult.Ok(State("done", w));
    }

    // ---------- breakpoints ----------

    private async Task<ToolResult> BreakpointToggleAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        // run mode too: a breakpoint toggled while the program runs stops it, and clearing it works (spike 37)
        RequireMode(inst, IdeMode.Design, IdeMode.Break, IdeMode.Run);
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
        // not in run mode: Clear All Breakpoints did nothing while the program ran, and the next Add still stopped (spike 37)
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
        // Side effects (database writes, processes, files, registry, network, scripts, keystrokes) are refused unless the
        // owner configured VB6MCP_EVAL=open: the IDE runs elevated, so this is code execution with admin rights (PLAN.md E5)
        var effect = EvalPolicy.SideEffect(expr);
        // allow_side_effect needs the user's approval, and who asks depends on the client (spike 40): the Claude Code hook (owner's
        // choice, 2026-10-17: ask rather than refuse) only for Claude Code, since other clients don't run it; other clients through
        // MCP elicitation; without either it is refused, since an agent could pass the flag itself.
        if (effect is not null && !EvalOpen)
        {
            var how = EvalPolicy.Approval(EvalClientId(), EvalGuarded(), EvalClientElicits());
            if (how == EvalApproval.None || a.Bool("allow_side_effect") != true)
            {
                Audit.Write("debug_eval", inst.Pid, $"refused ({effect}): ? {expr}");
                throw new ToolException($"refused: this expression looks like it has a side effect ({effect}). Prefer testing through the program's own UI (app_click, app_set_text). " +
                    (how != EvalApproval.None
                        ? "If it is really needed, call again with allow_side_effect=true: the user is asked to approve it."
                        : "Side effects need the user's approval, which needs the vb6-mcp Claude Code hook (install-server.ps1 / install.ps1 install it) " +
                          "or a client that can ask the user (MCP elicitation); or the owner sets VB6MCP_EVAL=open."));
            }
            if (how == EvalApproval.Elicitation) await AskToEvalAsync(inst, expr, effect, ct);
            Audit.Write("debug_eval", inst.Pid, how == EvalApproval.Hook
                ? $"side effect approved by the user (Claude Code hook prompt) ({effect}) follows"
                : $"side effect approved by the user (MCP elicitation in {Mcp.McpServer.ClientName}) ({effect}) follows");
        }
        else if (effect is not null) Audit.Write("debug_eval", inst.Pid, $"side effect allowed by VB6MCP_EVAL=open ({effect}) follows");
        var (status, result, w) = await EvaluateAsync(inst, expr, "debug_eval", ct);
        var effectNote = effect is null ? null : $"this had a side effect ({effect}); undo it when you are done (e.g. put data back), and mention it in your report";
        if (status == "dialog" && await CloseOwnCompileErrorAsync(inst, w!, ct) is { } compileError)
            return ToolResult.Fail($"the expression doesn't compile: {compileError} (VB's dialog was closed; the IDE is in {_env.ReadWindows(inst.Pid).Mode.ToString().ToLowerInvariant()} mode)" +
                                   // printing a whole record gave "Type mismatch" or "Only user-defined types ... coerced" (T17)
                                   (compileError.Contains("Type mismatch", StringComparison.OrdinalIgnoreCase) || compileError.Contains("user-defined type", StringComparison.OrdinalIgnoreCase)
                                       ? ". If it is a user-defined type (a Type ... End Type record), print its fields one by one, e.g. CStr(items(1).Price)" : ""));
        return status switch
        {
            "ok" => ToolResult.Ok(new { expression = expr, status, result, sideEffect = effect, note = effectNote }),
            "dialog" => ToolResult.Ok(new { expression = expr, status, note = "evaluating showed a dialog (e.g. a runtime error); it was not answered", dialogs = Dialogs(w!) }),
            _ => ToolResult.Ok(new { expression = expr, status, note = $"no result after {CommandWait.TotalSeconds:0} s; check debug_immediate and ide_status" }),
        };
    }

    /// <summary>
    /// Asks the user to approve a side effect through the client's own form (MCP elicitation). The agent doesn't see the form,
    /// so it can't answer it; only "accept" with the box ticked runs the expression. Anything else (decline, cancel, an error,
    /// no answer in <see cref="EvalElicitTimeout"/>) refuses it.
    /// </summary>
    private async Task AskToEvalAsync(InstanceInfo inst, string expr, string effect, CancellationToken ct)
    {
        var message = $"An agent wants to run this in the VB6 IDE's Immediate window. VB6 runs as administrator, so it runs with admin rights.\n\n? {expr}\n\n" +
                      $"It looks like it has a side effect ({effect}). Run it?";
        var schema = new System.Text.Json.Nodes.JsonObject
        {
            ["type"] = "object",
            ["properties"] = new System.Text.Json.Nodes.JsonObject
            {
                ["approve"] = new System.Text.Json.Nodes.JsonObject
                {
                    ["type"] = "boolean", ["title"] = "Run it",
                    ["description"] = "Tick to run the expression above in the elevated IDE.",
                },
            },
            ["required"] = new System.Text.Json.Nodes.JsonArray("approve"),
        };
        System.Text.Json.Nodes.JsonObject answer;
        try { answer = await EvalElicit(message, schema, ct); }
        catch (Mcp.McpRequestException ex)
        {
            Audit.Write("debug_eval", inst.Pid, $"refused, no approval ({ex.Message}) ({effect}): ? {expr}");
            throw new ToolException($"refused: the user couldn't be asked to approve the side effect ({effect}): {ex.Message}");
        }
        if (EvalPolicy.Approves(answer)) return;
        // "not approved" rather than "declined by the user": Copilot CLI with -p (no user) answers decline by itself, within 1 ms
        var action = (string?)answer["action"] ?? "?";
        Audit.Write("debug_eval", inst.Pid, $"refused, not approved (MCP elicitation in {Mcp.McpServer.ClientName}: {action}) ({effect}): ? {expr}");
        throw new ToolException($"refused: the side effect was not approved ({effect}; answer: {action}" +
                                (action == "accept" ? ", the box wasn't ticked" : "") + "). Don't retry it; test through the program's own UI.");
    }

    /// <summary>
    /// Closes a "Compile error" message that our own Immediate line just caused (the only dialog open, with an OK button)
    /// and returns its text; null if the dialog is anything else, which is left for the agent (hard rule 9: a run-time
    /// error in the program is the agent's decision). Benchmark T17: debug_locals printed a private user-defined type,
    /// VB answered with a compile error, and the agent had to close it with ide_dialog_respond.
    /// </summary>
    private async Task<string?> CloseOwnCompileErrorAsync(InstanceInfo inst, IdeWindowState w, CancellationToken ct)
    {
        if (w.Dialogs.Count != 1) return null;
        var d = w.Dialogs[0];
        var text = string.Join(" ", d.Texts).Trim();
        if (!text.StartsWith("Compile error", StringComparison.OrdinalIgnoreCase)) return null;
        if (d.Buttons.FirstOrDefault(b => b.Caption.Equals("OK", StringComparison.OrdinalIgnoreCase)) is not { } ok) return null;
        try { await _env.Client(inst).ClickAsync(ok.Hwnd); }
        catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException) { await PressViaWatchdogAsync(ok.Hwnd, inst.Pid); }
        for (var until = DateTime.UtcNow.AddSeconds(3); DateTime.UtcNow < until; await Task.Delay(200, ct))
            if (_env.ReadWindows(inst.Pid).Dialogs.Count == 0) return Regex.Replace(text, @"^Compile error:?\s*", "", RegexOptions.IgnoreCase);
        return null;
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

    /// <summary>How long compile/run tools wait for reloads the watchdog has pending (it polls every 2 s).</summary>
    public TimeSpan ReloadSettle { get; init; } = TimeSpan.FromSeconds(8);

    /// <summary>
    /// A file changed on disk during a run is reloaded only in design mode, a few seconds after run_end. A syntax_check
    /// called straight after compiled the previous text (improvement notes 27). Before compiling or running, wait for the
    /// watchdog's pending reloads of this IDE (sync events: last event of a file is "pending"), and say what happened.
    /// </summary>
    public async Task<JsonNode?> SettleReloadsAsync(int pid, CancellationToken ct)
    {
        List<string> Pending(int pid) => Core.SyncState.Read(SyncStatePath).Events.Where(e => e.Pid == pid)
            .GroupBy(e => e.File, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.OrderBy(e => e.Time).Last().Kind == "pending").Select(g => g.Key).ToList();
        var pending = Pending(pid);
        if (pending.Count == 0) return null;
        var deadline = DateTime.UtcNow + ReloadSettle;
        var still = pending;
        while (still.Count > 0 && DateTime.UtcNow < deadline) { await Task.Delay(500, ct); still = Pending(pid); }
        return JsonSerializer.SerializeToNode(new
        {
            applied = pending.Except(still, StringComparer.OrdinalIgnoreCase).ToList(),
            stillPending = still.Count > 0 ? still : null,
            note = still.Count > 0 ? "these files changed on disk and aren't reloaded yet, so the IDE still has the previous text; check sync_status" : "waited for these reloads (files changed on disk) before going on",
        }, McpServer.Json);
    }

    private async Task<ToolResult> SyntaxCheckAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        var settled = await SettleReloadsAsync(inst.Pid, ct);
        // Code no tool edited gets compiled too: say so, the result may pass or fail because of it (improvement notes 50)
        var outside = new List<object>();
        try
        {
            var client = _env.Client(inst);
            foreach (var p in await client.ProjectsAsync())
                foreach (var o in await OutsideEditsAsync(inst, client, p.Name, (await client.ComponentsAsync(p.Name)).Where(c => c.IsDirty)))
                    outside.Add(new { project = p.Name, component = o.Component, since = o.Since, changedLines = o.ChangedLines, diff = o.Diff, moreDiffLines = o.MoreDiffLines });
        }
        catch (Exception ex) when (ex is AddinErrorException or AddinBusyException or AddinUnreachableException) { }
        return (await SyntaxCheckCoreAsync(a, ct)).With("reloads", settled)
            .With("outsideEdits", outside.Count > 0 ? JsonSerializer.SerializeToNode(outside, McpServer.Json) : null);
    }

    private async Task<ToolResult> SyntaxCheckCoreAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireMode(inst, IdeMode.Design);
        var client = _env.Client(inst);
        await RememberDirtyBeforeRunAsync(inst);   // a passing syntax_check runs the program briefly (spike 32)
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
            return ToolResult.Ok(new { compiles = true, status = after.Mode == IdeMode.Design ? "ended" : "still-running", mode = Mode(after.Mode) })
                .With("markedChangedByRun", MarkedJson(await MarkedChangedByRunAsync(inst)));
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
