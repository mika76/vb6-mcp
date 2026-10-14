using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 2 inspection: locals, call stack, runtime errors. VB6 has no debugger API, and the Locals
/// window is custom-drawn, so locals come from the procedure's declarations plus Immediate
/// evaluation; the call stack is read from the Call Stack dialog by the add-in. Documented in docs/tools.md.
/// </summary>
public sealed partial class IdeTools
{
    /// <summary>Separator between evaluated values; unlikely in real data, and checked when parsing.</summary>
    private const string Sep = "|~|";

    /// <summary>Longest Immediate line debug_locals types (the IDE's line limit is 1023, hard rule 10).</summary>
    private const int MaxEvalLine = 700;

    /// <summary>Types whose value CStr can print without raising an error (no Null, arrays, objects).</summary>
    private static readonly HashSet<string> s_scalarTypes = new(StringComparer.OrdinalIgnoreCase)
        { "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date", "Boolean", "Byte", "String" };

    private IReadOnlyList<Tool> InspectTools() =>
    [
        new("debug_locals",
            "Variables in scope at the current statement (break mode): the procedure's parameters, return value and Dim/Static " +
            "variables, plus the module's variables. Each with its runtime type (TypeName) and, for simple types, its value. " +
            "Arrays, objects and Null show their type only; inspect them with debug_eval. Reads nothing from disk and runs no user code.",
            Schema.Object(("module", "boolean", "Include module-level variables (default true).", false), ("pid", "integer", PidDesc, false)),
            (a, ct) => LocalsAsync(a, ct)),
        new("debug_callstack",
            "The call stack in break mode, most recent call first (from View > Call Stack, which is opened, read and closed again).",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, ct) => CallStackAsync(a, ct)),
        new("debug_output",
            "OutputDebugString output of the programs being worked on (compiled VB6 programs have no Debug.Print; programs that log " +
            "usually call OutputDebugString). The first call starts capturing (earlier output is lost); later calls return what " +
            "arrived, after after_id. By default only VB6.EXE (the program running in the IDE) and programs started with exe_run; " +
            "all=true shows every process. stop=true releases the capture buffer (only one monitor, e.g. DebugView, can own it).",
            Schema.Object(("after_id", "integer", "Only messages after this id (the last_id of the previous call).", false),
                          ("all", "boolean", "Messages of every process.", false), ("max", "integer", "At most this many (default 200).", false),
                          ("stop", "boolean", "Stop capturing and release the buffer.", false)),
            (a, _) => Task.FromResult(DebugOutputRead(a))),
        new("debug_error",
            "The runtime error dialog ('Run-time error 11: Division by zero'), if one is open: error number and message. With " +
            "action 'debug' it presses Debug and returns the failing line (then use debug_locals / debug_callstack); with action " +
            "'end' it ends the program (needs force=true if the agent didn't start the run).",
            Schema.Object(("action", "string", "debug or end. Without it, only reports the error.", false),
                          ("force", "boolean", "Allow 'end' for a run the agent didn't start.", false), ("pid", "integer", PidDesc, false)),
            (a, ct) => ErrorAsync(a, ct)),
    ];

    // ---------- debug_locals ----------

    private sealed class Var(string name, string scope, string kind, string? declaredType)
    {
        public string Name { get; } = name;
        public string Scope { get; } = scope;
        public string Kind { get; } = kind;
        public string? DeclaredType { get; } = declaredType;
        public string? Type { get; set; }
        public string? Value { get; set; }
        public int? Length { get; set; }
        public bool Truncated { get; set; }
    }

    private async Task<ToolResult> LocalsAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireMode(inst, IdeMode.Break);
        var client = _env.Client(inst);
        var where = await client.WhereAsync();
        if (where.Procedure.Length == 0) throw new ToolException($"the current statement ({where.Component}:{where.StartLine}) isn't inside a procedure");

        // The procedure's text: try each kind (Sub/Function, Get, Let, Set) and take the one containing the line
        ProcText? proc = null;
        for (int kind = 0; kind <= 3 && proc is null; kind++)
        {
            try
            {
                var p = await client.ProcAsync(where.Project, where.Component, where.Procedure, kind);
                if (where.StartLine >= p.ProcStart && where.StartLine < p.ProcStart + p.ProcCount) proc = p;
            }
            catch (AddinErrorException) { }
        }
        if (proc is null) throw new ToolException($"procedure {where.Procedure} wasn't found in {where.Component}");

        var vars = VbLocals.Parse(proc.Text)
            .Select(d => new Var(d.Name, "procedure", d.Kind.ToString().ToLowerInvariant() switch { "returnvalue" => "return", var k => k }, d.DeclaredType)).ToList();
        if (a.Bool("module") ?? true)
        {
            var local = vars.Select(v => v.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var m in await client.MembersAsync(where.Project, where.Component))
                if (m.MemberType == 3 && !local.Contains(m.Name))        // 3 = vbext_mt_Variable; a local of the same name hides it
                    vars.Add(new Var(m.Name, "module", m.Scope == 2 ? "public" : "private", null));
        }

        string? problem = null;
        // Pass 1: TypeName never raises an error, whatever the variable holds
        foreach (var chunk in Chunks(vars, v => $"TypeName({v.Name})"))
        {
            var r = await EvalListAsync(inst, chunk.Select(x => x.Expr).ToList(), ct);
            if (r.Error != null) { problem = r.Error; break; }
            for (int i = 0; i < chunk.Count; i++) chunk[i].Var.Type = r.Values![i];
        }
        // Pass 2: values, only for types CStr can print (CStr of Null raised "Invalid use of Null", Phase 2 test)
        if (problem is null)
        {
            var scalars = vars.Where(v => v.Type != null && s_scalarTypes.Contains(v.Type)).ToList();
            foreach (var chunk in Chunks(scalars, v => v.Type!.Equals("String", StringComparison.OrdinalIgnoreCase)
                         ? $"Len({v.Name}) & \":\" & Replace(Replace(Left$({v.Name}, 200), vbCr, \"\\r\"), vbLf, \"\\n\")"
                         : $"CStr({v.Name})"))
            {
                var r = await EvalListAsync(inst, chunk.Select(x => x.Expr).ToList(), ct);
                if (r.Error != null) { problem = r.Error; break; }
                for (int i = 0; i < chunk.Count; i++)
                {
                    var v = chunk[i].Var; var s = r.Values![i];
                    if (v.Type!.Equals("String", StringComparison.OrdinalIgnoreCase) && s.IndexOf(':') is int c and > 0
                        && int.TryParse(s[..c], NumberStyles.Integer, CultureInfo.InvariantCulture, out var len))
                    { v.Length = len; v.Value = s[(c + 1)..]; v.Truncated = len > 200; }
                    else v.Value = s;
                }
            }
            foreach (var v in vars.Where(v => v.Value is null && v.Type is "Empty" or "Null" or "Nothing")) v.Value = v.Type;
        }

        object Shape(Var v) => new
        {
            name = v.Name, kind = v.Kind, type = v.Type, value = v.Value, declared = v.DeclaredType,
            length = v.Length, truncated = v.Truncated ? true : (bool?)null,
        };
        return ToolResult.Ok(new
        {
            where.Project, where.Component, procedure = where.Procedure, line = where.StartLine,
            locals = vars.Where(v => v.Scope == "procedure").Select(Shape),
            module = vars.Where(v => v.Scope == "module").Select(Shape),
            note = problem ?? "values of arrays, objects and Null aren't shown: use debug_eval (e.g. UBound(items), items(0), obj.Name). " +
                             "Strings are cut at 200 characters (truncated, length); \\r and \\n stand for line breaks",
        });
    }

    /// <summary>Groups expressions into Immediate lines of at most <see cref="MaxEvalLine"/> characters.</summary>
    private static List<List<(Var Var, string Expr)>> Chunks(IEnumerable<Var> vars, Func<Var, string> expr)
    {
        var chunks = new List<List<(Var, string)>>();
        var current = new List<(Var, string)>(); int length = 0;
        foreach (var v in vars)
        {
            var e = expr(v);
            if (current.Count > 0 && length + e.Length + 10 > MaxEvalLine) { chunks.Add(current); current = []; length = 0; }
            current.Add((v, e)); length += e.Length + 10;
        }
        if (current.Count > 0) chunks.Add(current);
        return chunks;
    }

    /// <summary>
    /// Evaluates several expressions in one Immediate line, joined by <see cref="Sep"/>, with a final
    /// separator so trailing spaces survive the line trimming in <see cref="ResultAfter"/>.
    /// </summary>
    private async Task<(IReadOnlyList<string>? Values, string? Error)> EvalListAsync(InstanceInfo inst, IReadOnlyList<string> exprs, CancellationToken ct)
    {
        string line = string.Join($" & \"{Sep}\" & ", exprs) + $" & \"{Sep}\"";
        var (status, result, w) = await EvaluateAsync(inst, line, "debug_locals", ct);
        if (status == "dialog")
            return (null, $"evaluating showed a dialog ('{w!.Dialogs[0].Caption}': {string.Join(" ", w.Dialogs[0].Texts)}); it was not answered");
        if (status != "ok" || result is null) return (null, "the Immediate window didn't show a result in time");
        var parts = SplitValues(result, exprs.Count);
        return parts is null ? (null, "the values couldn't be separated (a value contains the separator?): " + result) : (parts, null);
    }

    /// <summary>Splits "v1|~|v2|~|" into exactly <paramref name="count"/> values, or null if that's not possible.</summary>
    public static IReadOnlyList<string>? SplitValues(string printed, int count)
    {
        var parts = printed.Replace("\n", "\\n").Split(Sep);
        // the line ends with a separator, so the last part is empty
        if (parts.Length != count + 1 || parts[^1].Length != 0) return null;
        return parts[..count];
    }

    // ---------- debug_output ----------

    /// <summary>One capture per server process: the DBWIN buffer can only have one owner.</summary>
    private static DebugOutput? s_dbwin;
    private static readonly object s_dbwinLock = new();

    private ToolResult DebugOutputRead(JsonObject a)
    {
        lock (s_dbwinLock)
        {
            if (a.Bool("stop") == true)
            {
                bool was = s_dbwin?.Running == true;
                s_dbwin?.Dispose(); s_dbwin = null;
                return ToolResult.Ok(new { capturing = false, note = was ? "capture stopped; the DBWIN buffer is free for DebugView" : "wasn't capturing" });
            }
            bool started = false;
            if (s_dbwin is null)
            {
                var mon = new DebugOutput();
                if (mon.Start() is { } why) { mon.Dispose(); throw new ToolException(why); }
                s_dbwin = mon; started = true;
            }
            long after = a.Int("after_id") ?? 0;
            bool all = a.Bool("all") ?? false;
            int max = Math.Clamp(a.Int("max") ?? 200, 1, 2000);
            var msgs = s_dbwin.Read(after, m => all || m.Process.Equals("VB6", StringComparison.OrdinalIgnoreCase) || _exes.ContainsKey(m.Pid), max);
            return ToolResult.Ok(new
            {
                capturing = true, since = s_dbwin.Started.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                messages = msgs.Select(m => new { id = m.Id, time = m.Time.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture), pid = m.Pid, process = m.Process, text = m.Text }),
                last_id = s_dbwin.LastId,
                note = started ? "capture started now; output written before this call isn't available. Call again (with after_id) to read new output" : null,
            });
        }
    }

    // ---------- debug_callstack ----------

    private async Task<ToolResult> CallStackAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        var w = _env.ReadWindows(inst.Pid);
        var dlg = w.Dialogs.FirstOrDefault(d => d.Caption == "Call Stack");
        bool opened = false;
        if (dlg is null)
        {
            RequireMode(inst, IdeMode.Break);
            await _env.Client(inst).CallStackAsync();
            w = await WaitForAsync(inst, s => s.Dialogs.Any(d => d.Caption == "Call Stack"), ct, TimeSpan.FromSeconds(5));
            dlg = w.Dialogs.FirstOrDefault(d => d.Caption == "Call Stack")
                  ?? throw new ToolException("the Call Stack dialog didn't open" + (w.Dialogs.Count > 0 ? $" (a dialog is open: '{w.Dialogs[0].Caption}')" : ""));
            opened = true;
        }
        var list = AppWindows.Controls(dlg.Hwnd).FirstOrDefault(c => c.Class == "ListBox")
                   ?? throw new ToolException("the Call Stack dialog has no list");
        var client = _env.Client(inst);
        var items = await client.ListItemsAsync(list.Hwnd);
        bool closed = false;
        if (opened && dlg.Buttons.FirstOrDefault(b => b.Caption == "Close") is { } close)
        {
            await client.ClickAsync(close.Hwnd);
            var after = await WaitForAsync(inst, s => s.Dialogs.All(d => d.Caption != "Call Stack"), ct, TimeSpan.FromSeconds(3));
            closed = after.Dialogs.All(d => d.Caption != "Call Stack");
        }
        return ToolResult.Ok(new
        {
            frames = items.Select((text, i) =>
            {
                var parts = text.Split('.');
                return parts.Length == 3
                    ? (object)new { depth = i, project = parts[0], component = parts[1], procedure = parts[2] }
                    : new { depth = i, text };
            }),
            note = opened && !closed ? "the Call Stack dialog is still open; close it with ide_dialog_respond button Close"
                 : !opened ? "the Call Stack dialog was already open and was left open" : null,
        });
    }

    // ---------- debug_error ----------

    /// <summary>Number and message of a VB runtime error dialog ("Run-time error '11':\n\nDivision by zero"), or null.</summary>
    public static (int Number, string Message)? RuntimeError(DialogInfo d)
    {
        var m = RuntimeErrorText().Match(string.Join("\n", d.Texts));
        return m.Success ? (int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture), m.Groups[2].Value.Trim()) : null;
    }

    [GeneratedRegex(@"Run-time error '(-?\d+)':\s*(.*)", RegexOptions.Singleline)]
    private static partial Regex RuntimeErrorText();

    private async Task<ToolResult> ErrorAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        var w = _env.ReadWindows(inst.Pid);
        var dlg = w.Dialogs.FirstOrDefault(d => RuntimeError(d) != null);
        if (dlg is null)
            return ToolResult.Ok(new { runtimeError = (object?)null, mode = Mode(w.Mode), dialogs = Dialogs(w), note = "no runtime error dialog is open" });
        var (number, message) = RuntimeError(dlg)!.Value;
        var action = a.Str("action")?.ToLowerInvariant();
        var client = _env.Client(inst);
        switch (action)
        {
            case null or "":
                return ToolResult.Ok(new
                {
                    runtimeError = new { number, message }, buttons = dlg.Buttons.Select(b => b.Caption), mode = Mode(w.Mode),
                    note = "pass action 'debug' to stop at the failing line, or 'end' to end the program",
                });
            case "debug":
            {
                var debug = dlg.Buttons.FirstOrDefault(b => b.Caption == "Debug") ?? throw new ToolException("the dialog has no Debug button (error in compiled code?)");
                await client.ClickAsync(debug.Hwnd);
                var after = await WaitForAsync(inst, s => s.Mode == IdeMode.Break && s.Dialogs.Count == 0, ct);
                Where? where = null; string? lineText = null;
                try
                {
                    where = await client.WhereAsync();
                    lineText = (await client.ReadAsync(where.Project, where.Component, where.StartLine, 1)).Text;
                }
                catch (Exception ex) when (ex is AddinErrorException or AddinBusyException) { }
                return ToolResult.Ok(new
                {
                    runtimeError = new { number, message }, status = after.Mode == IdeMode.Break ? "break" : Mode(after.Mode),
                    location = where is null ? null : new { where.Project, where.Component, line = where.StartLine, procedure = where.Procedure, lineText },
                    note = "the failing statement is the current line: inspect with debug_locals, debug_callstack, debug_eval; fix with proc_replace " +
                           "(only in design mode: run_end first) or move on with step_over / run_continue",
                });
            }
            case "end":
            {
                if (!_runsWeStarted.Contains(inst.Pid) && a.Bool("force") != true)
                    throw new ToolException("this run wasn't started by the agent (it may be the user's); pass force=true to end it");
                var end = dlg.Buttons.FirstOrDefault(b => b.Caption == "End") ?? throw new ToolException("the dialog has no End button");
                await client.ClickAsync(end.Hwnd);
                var after = await WaitForAsync(inst, s => s.Mode == IdeMode.Design, ct);
                if (after.Mode == IdeMode.Design) _runsWeStarted.Remove(inst.Pid);
                return ToolResult.Ok(new { runtimeError = new { number, message }, status = after.Mode == IdeMode.Design ? "ended" : Mode(after.Mode), dialogs = Dialogs(after) });
            }
            default:
                throw new ToolException("action must be debug or end");
        }
    }
}
