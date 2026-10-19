using System.Globalization;
using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 3: drive a running program, either the one running in the IDE or a compiled EXE started with
/// exe_run (argument exe). Reads (windows, controls, text) are done by the server: they cross the
/// elevation boundary (spike 10). Input into the IDE's program goes through the add-in, inside the
/// elevated IDE process; input into a compiled EXE (same integrity as the server) is sent by the
/// server (<see cref="LocalInput"/>). Everything is audited (hard rule 8). Documented in docs/tools.md.
/// </summary>
public sealed partial class IdeTools
{
    private const string WindowDesc = "Form or dialog: window handle (hex, from app_windows) or caption. Default: the program's only/first form.";
    private const string ExeDesc = "Act on a compiled program started with exe_run (its pid) instead of the program running in the IDE.";

    /// <summary>How long app_screenshot waits for the add-in's capture file.</summary>
    public TimeSpan SnapWait { get; init; } = TimeSpan.FromSeconds(5);

    /// <summary>
    /// The program an app_* tool acts on: the one running in an IDE (<paramref name="Ide"/>, input
    /// through its add-in) or a compiled EXE started with exe_run (input from the server).
    /// </summary>
    private sealed record Target(int Pid, InstanceInfo? Ide, string? Exe)
    {
        public bool IsIde => Ide != null;
    }

    /// <summary>Common arguments of the app_* tools: pid (IDE) and exe.</summary>
    private static (string, string, string, bool)[] TargetArgs() => [("pid", "integer", PidDesc, false), ("exe", "integer", ExeDesc, false)];

    private static JsonObject Args(params (string, string, string, bool)[] own) => Schema.Object([.. own, .. TargetArgs()]);

    private IReadOnlyList<Tool> AppTools() =>
    [
        new("app_windows",
            "List the windows (forms, dialogs) of the running program (in the IDE, or a compiled EXE with exe), with caption, class, position and size.",
            Args(), (a, _) => Task.FromResult(AppWindowsList(a))),
        new("app_tree",
            "Controls of a form of the running program: window handle, kind (CommandButton, TextBox, ...), text, position (relative to the " +
            "form's client area), visible, enabled. Lightweight controls (Label, Image, Shape, Line) have no window and only show in app_screenshot.",
            Args(("window", "string", WindowDesc, false)), (a, _) => Task.FromResult(AppTree(a))),
        new("app_click",
            "Click a command button, check box or option button of the running program. Address it by control (hex handle from app_tree) " +
            "or by text (its caption, optionally within window). Returns the state afterwards (mode, dialogs, windows). A control the user " +
            "can't see (hidden, or outside its container's visible area: app_tree's clipped) is refused unless force=true.",
            Args(("control", "string", "Control window handle (hex) from app_tree.", false), ("text", "string", "Control caption, e.g. Add or &Add.", false),
                 ("window", "string", WindowDesc, false), ("force", "boolean", "Click even a control the user can't see.", false)),
            (a, ct) => AppClickAsync(a, ct)),
        new("app_set_text",
            "Set the text of a text box in the running program (fires its Change event). Address it by control (hex handle from app_tree) " +
            "or by index (0-based among the form's text boxes, in app_tree order). Returns the text read back.",
            Args(("text", "string", "New text.", true), ("control", "string", "Text box window handle (hex).", false),
                 ("index", "integer", "0-based index among the form's text boxes in reading order (top to bottom, left to right); app_tree shows each box's index.", false), ("window", "string", WindowDesc, false)),
            (a, ct) => AppSetTextAsync(a, ct)),
        new("app_mouse",
            "Mouse input at a position in a form of the running program: click, double_click, right_click, or drag (from x,y to to_x,to_y). " +
            "x/y are pixels relative to the form's client area (the coordinates app_tree reports). By default the messages are posted to the " +
            "window under the point, without touching the user's mouse; that works for windowed controls (buttons, lists, picture boxes). " +
            "Labels and Images (no window) only react to real input: pass real_input=true, which brings the form to the front, moves the " +
            "real cursor, clicks, then puts the cursor and the active window back.",
            Args(("action", "string", "click, double_click, right_click or drag.", true), ("x", "integer", "X in the form's client area.", true),
                 ("y", "integer", "Y in the form's client area.", true), ("to_x", "integer", "Drag end X.", false), ("to_y", "integer", "Drag end Y.", false),
                 ("real_input", "boolean", "Use real mouse input (needed for Labels/Images): takes the foreground and moves the cursor briefly.", false),
                 ("window", "string", WindowDesc, false)),
            (a, ct) => AppMouseAsync(a, ct)),
        new("app_scroll",
            "Scroll a control of the running program (list box, multi-line text box, grid, scroll bar control, or a scrollable form). " +
            "direction up/down/left/right; amount line (default), page or end; count repeats (default 1). method 'scrollbar' (default) sends " +
            "scroll bar commands, 'wheel' sends mouse-wheel notches at the control's centre. Returns the scroll position before and after.",
            Args(("control", "string", "Window handle (hex) from app_tree. Default: the form itself.", false),
                 ("direction", "string", "up, down, left or right.", true), ("amount", "string", "line (default), page or end.", false),
                 ("count", "integer", "Repeat count, 1-100 (default 1).", false), ("method", "string", "scrollbar (default) or wheel.", false),
                 ("window", "string", WindowDesc, false)),
            (a, ct) => AppScrollAsync(a, ct)),
        new("app_select",
            "Select an item in a list box or combo box of the running program, by text (exact, case-insensitive) or index (0-based). " +
            "Raises the control's Click event like a user's click. Returns the selection read back.",
            Args(("control", "string", "List or combo box window handle (hex) from app_tree. Default: the form's only list or combo box, or the only one that has the item text.", false), ("text", "string", "Item text.", false),
                 ("index", "integer", "0-based item index.", false)),
            (a, ct) => AppSelectAsync(a, ct)),
        new("app_keys",
            "Type text or press keys in the running program. Plain text is typed as characters; braces name keys: {ENTER} {TAB} {ESC} " +
            "{BACKSPACE} {DELETE} {UP} {DOWN} {LEFT} {RIGHT} {HOME} {END} {PGUP} {PGDN} {F1}..{F12}, with modifiers {CTRL+S} " +
            "{SHIFT+TAB} {ALT+F}; {LBRACE}/{RBRACE} for literal braces. Goes to control, or to the program's focused control (never to the IDE). " +
            "Posted by default (doesn't disturb the user); real_input=true sends real keystrokes (brings the form to the front briefly).",
            Args(("keys", "string", "Text and {KEY} names.", true), ("control", "string", "Target window handle (hex): a control (app_tree) or a form (app_windows). Default: the focused control.", false),
                 ("real_input", "boolean", "Real keystrokes (SendInput) instead of posted messages.", false)),
            (a, ct) => AppKeysAsync(a, ct)),
        new("app_wait",
            "Wait until something happens in the running program or the IDE, up to timeout seconds: a window with a caption appears " +
            "(window) or disappears (window + gone), a control's text equals or contains a value (control + text / contains), the IDE " +
            "reaches a mode (mode: run, break, design), or a dialog appears (dialog: true). Returns what was seen and how long it took.",
            Args(("window", "string", "Window caption (substring, case-insensitive).", false), ("gone", "boolean", "Wait for the window to disappear.", false),
                 ("control", "string", "Control window handle (hex).", false), ("text", "string", "Exact text of control.", false),
                 ("contains", "string", "Text control should contain.", false), ("mode", "string", "run, break or design (IDE only).", false),
                 ("dialog", "boolean", "Wait for a dialog.", false), ("timeout", "integer", "Seconds (default 10, at most 120).", false)),
            (a, ct) => AppWaitAsync(a, ct)),
        new("app_screenshot",
            "Screenshot of a window as an image: a form or dialog of the running program (default), or the whole IDE (target 'ide'). " +
            "Shows what app_tree can't: Labels, Images, drawing, layout. Works when the window is covered.",
            Args(("window", "string", WindowDesc + " Or 'ide' for the IDE main window.", false),
                 ("max_width", "integer", "Scale down to at most this width (default 1024).", false)),
            (a, ct) => AppScreenshotAsync(a, ct)),
    ];

    // ---------- reads ----------

    private ToolResult AppWindowsList(JsonObject a)
    {
        var t = ResolveTarget(a);
        var windows = AppWindows.List(t.Pid);
        IdeMode? mode = t.IsIde ? _env.ReadWindows(t.Pid).Mode : null;
        return ToolResult.Ok(new
        {
            mode = mode is { } m ? Mode(m) : null, exe = t.Exe,
            windows = windows.Select(x => new { hwnd = Hex(x.Hwnd), x.Kind, x.Caption, x.Class, x.Left, x.Top, x.Width, x.Height, x.Enabled }),
            note = mode == IdeMode.Design ? "the IDE is in design mode: the program isn't running (run_start)" : null,
        });
    }

    private ToolResult AppTree(JsonObject a)
    {
        var t = ResolveTarget(a);
        var form = ResolveWindow(t, a.Str("window"));
        var controls = AppWindows.Controls(form.Hwnd);
        // each text box's app_set_text index (reading order), so the agent doesn't count them in window order (T18)
        var textIndex = AppWindows.TextBoxes(form.Hwnd).Select((c, i) => (c.Hwnd, i)).ToDictionary(x => x.Hwnd, x => x.i);
        return ToolResult.Ok(new
        {
            window = new { hwnd = Hex(form.Hwnd), form.Kind, form.Caption },
            // Compact on purpose (token breakdown, docs/benchmarks): no window class (Kind comes from it), and hidden/disabled
            // only when true. app_tree was up to 8% of a run's re-read tokens.
            controls = controls.Select(c =>
            {
                var o = new System.Text.Json.Nodes.JsonObject { ["hwnd"] = Hex(c.Hwnd), ["kind"] = c.Kind, ["text"] = c.Text, ["left"] = c.Left, ["top"] = c.Top, ["width"] = c.Width, ["height"] = c.Height };
                if (textIndex.TryGetValue(c.Hwnd, out var ti)) { o["index"] = ti; if (AppWindows.IsReadOnly(c.Hwnd)) o["readOnly"] = true; }
                if (c.Depth > 0) o["depth"] = c.Depth;
                if (!c.Visible) o["hidden"] = true;
                if (c.Clipped is { } clip) o["clipped"] = clip;   // "partly" or "hidden": outside a container's visible area
                if (!c.Enabled) o["disabled"] = true;
                return o;
            }),
            note = "Labels, Images, Shapes and Lines have no window and aren't listed; use app_screenshot to see them",
        });
    }

    // ---------- input ----------

    private async Task<ToolResult> AppClickAsync(JsonObject a, CancellationToken ct)
    {
        var t = ResolveTarget(a);
        RequireReady(t, allowDialogTarget: true);
        AppControl target;
        if (a.Str("control") is { Length: > 0 } hex) target = FindControlByHandle(t, ParseHex(hex));
        else
        {
            var text = a.Str("text") ?? throw new ToolException("pass control (handle from app_tree) or text (the caption)");
            var form = ResolveWindow(t, a.Str("window"));
            var wanted = text.Replace("&", "", StringComparison.Ordinal);
            var matches = AppWindows.Controls(form.Hwnd)
                .Where(c => c.Kind is "CommandButton" or "CheckBox" or "OptionButton" && c.Text.Replace("&", "", StringComparison.Ordinal).Equals(wanted, StringComparison.OrdinalIgnoreCase)).ToList();
            // a caption can belong to a visible button and a parked one (XmPanel moves unused buttons out of view)
            if (matches.Count > 1 && matches.Count(Seen) == 1) matches = [.. matches.Where(Seen)];
            target = matches.Count switch
            {
                0 => throw new ToolException($"no button, check box or option button captioned '{text}' in '{form.Caption}' (see app_tree)"),
                1 => matches[0],
                _ => throw new ToolException($"{matches.Count} controls are captioned '{text}'; pass control (handle) instead"),
            };
        }
        if (!target.Enabled) throw new ToolException($"'{target.Text}' is disabled");
        // improvement notes 33/52: the program acted on a click on a button no user could reach
        if (!Seen(target) && a.Bool("force") != true)
            throw new ToolException($"{target.Kind} '{target.Text}' ({Hex(target.Hwnd)}) can't be seen by the user: " +
                (!target.Visible ? "its window is hidden" : "it lies outside its container's visible area") +
                ". A user couldn't click it; pass force=true to click it anyway");
        Audit.Write("app_click", t.Pid, $"{Hex(target.Hwnd)} {target.Kind} '{target.Text}'");
        var before = await UiFingerprintAsync(t);
        if (t.Ide is { } inst) await _env.Client(inst).AppClickAsync(target.Hwnd);
        else LocalInput.Click(target.Hwnd);
        await Task.Delay(500, ct);              // the click and its event code run after the posted message
        var result = ToolResult.Ok(AfterState(t, $"clicked {target.Kind} '{target.Text}'"));
        // improvement notes 34: "clicked" was returned while a wedged form ignored the click. Say when nothing visible
        // changed (windows, captions, texts, enabled/visible states, dialogs); a click can still have effects not seen here.
        return before.Length > 0 && await UiFingerprintAsync(t) == before
            ? result.With("changed", System.Text.Json.Nodes.JsonValue.Create(false))
                    .With("verify", System.Text.Json.Nodes.JsonValue.Create("no change seen after the click: no window, text, enabled or visible state or list changed and no dialog appeared" + (t.IsIde ? " (check box states aren't compared for a program in the IDE). " : ". ") +
                          "The click may have done nothing (form busy or wedged, handler empty) or something not shown in windows (a file, the database, Debug.Print): check app_screenshot or debug_immediate"))
            : result;
    }

    private async Task<ToolResult> AppSetTextAsync(JsonObject a, CancellationToken ct)
    {
        var t = ResolveTarget(a);
        RequireReady(t, allowDialogTarget: TargetsDialog(t, a));
        var text = a.Str("text") ?? throw new ToolException("missing required argument 'text'");
        AppControl target;
        if (a.Str("control") is { Length: > 0 } hex) target = FindControlByHandle(t, ParseHex(hex));
        else
        {
            var form = ResolveWindow(t, a.Str("window"));
            var boxes = AppWindows.TextBoxes(form.Hwnd);   // reading order, as app_tree's index shows it
            int index = a.Int("index") ?? (boxes.Count == 1 ? 0 : throw new ToolException($"the form has {boxes.Count} text boxes; pass index or control"));
            if (index < 0 || index >= boxes.Count) throw new ToolException($"index {index} is out of range (the form has {boxes.Count} text boxes)");
            target = boxes[index];
        }
        if (target.Kind != "TextBox") throw new ToolException($"control {Hex(target.Hwnd)} is a {target.Kind}, not a text box");
        Audit.Write("app_set_text", t.Pid, $"{Hex(target.Hwnd)} '{text}'");
        if (t.Ide is { } inst) await _env.Client(inst).AppTextAsync(target.Hwnd, text);
        else await Task.Run(() => LocalInput.SetText(target.Hwnd, text), ct);
        await Task.Delay(400, ct);
        var now = AppWindows.Text(new IntPtr(target.Hwnd));
        // a Locked text box takes WM_SETTEXT, but no user could type there: say so (T18: the price went into the locked total)
        bool readOnly = AppWindows.IsReadOnly(target.Hwnd);
        return ToolResult.Ok(new { textNow = now, applied = now == text, readOnly = readOnly ? true : (bool?)null,
            note = readOnly ? $"text box {Hex(target.Hwnd)} is read-only (Locked): a user can't type into it. Was another box meant? app_tree shows each box's index" : null,
            state = AfterState(t, "text set") });
    }

    private async Task<ToolResult> AppMouseAsync(JsonObject a, CancellationToken ct)
    {
        var t = ResolveTarget(a);
        RequireReady(t, allowDialogTarget: false);
        var action = (a.Str("action") ?? "").ToLowerInvariant();
        string verb = action switch
        {
            "click" => "click", "double_click" => "dblclick", "right_click" => "rclick", "drag" => "drag",
            _ => throw new ToolException("action must be click, double_click, right_click or drag"),
        };
        int x = a.Int("x") ?? throw new ToolException("missing required argument 'x'");
        int y = a.Int("y") ?? throw new ToolException("missing required argument 'y'");
        var form = ResolveWindow(t, a.Str("window"));
        var (hwnd, lx, ly) = AppWindows.HitTest(form.Hwnd, x, y);
        string under = hwnd == form.Hwnd ? $"form '{form.Caption}'" : $"{AppWindows.Kind(ClassOf(hwnd))} {Hex(hwnd)}";
        bool real = a.Bool("real_input") ?? false;
        int? ex = null, ey = null;
        if (verb == "drag")
        {
            int tx = a.Int("to_x") ?? throw new ToolException("drag needs to_x and to_y");
            int ty = a.Int("to_y") ?? throw new ToolException("drag needs to_x and to_y");
            // the whole drag goes to the window it starts in (as with mouse capture), in its coordinates
            (ex, ey) = AppWindows.FormToClient(form.Hwnd, hwnd, tx, ty);
        }
        string? realNote = null;
        Audit.Write("app_mouse", t.Pid, $"{(real ? "real " : "")}{verb} {Hex(hwnd)} ({lx},{ly}){(ex is null ? "" : $"->({ex},{ey})")}");
        if (t.Ide is { } inst)
        {
            var client = _env.Client(inst);
            if (real)
            {
                // Spike 17: windowless controls only react to real input. The add-in can't wait for the
                // program to process it (hard rule 5), so the server waits, then always restores.
                string outcome;
                try
                {
                    await client.MouseRealAsync(hwnd, verb, lx, ly, ex, ey);   // deferred in the add-in (hard rule 14)
                    await Task.Delay(500, ct);
                }
                finally { outcome = await client.MouseRestoreAsync(); }
                if (outcome != "sent") throw new ToolException("real input wasn't sent: " + outcome);
            }
            else { await client.MouseAsync(hwnd, verb, lx, ly, ex, ey); await Task.Delay(500, ct); }
        }
        else if (real)
        {
            var outcome = await Task.Run(() => LocalInput.Real(hwnd, h => LocalInput.RealMouse(h, verb, lx, ly, ex ?? 0, ey ?? 0)), ct);
            if (!outcome.StartsWith("sent", StringComparison.Ordinal)) throw new ToolException("real input wasn't sent: " + outcome);
            realNote = outcome == "sent" ? null : outcome;
        }
        else { LocalInput.Mouse(hwnd, verb, lx, ly, ex ?? 0, ey ?? 0); await Task.Delay(500, ct); }
        return ToolResult.Ok(new
        {
            target = under, at = new { x = lx, y = ly }, method = real ? "real input (SendInput); cursor and active window restored" : "posted messages", warning = realNote,
            state = AfterState(t, $"{action} at ({x},{y})"),
            verify = real ? "check the effect with app_screenshot or app_tree"
                          : "posted input has no visible result of its own: check with app_screenshot or app_tree. A Label or Image ignores posted input; use real_input=true (spike 17)",
        });
    }

    private async Task<ToolResult> AppScrollAsync(JsonObject a, CancellationToken ct)
    {
        var t = ResolveTarget(a);
        RequireReady(t, allowDialogTarget: false);
        var direction = (a.Str("direction") ?? "").ToLowerInvariant();
        bool vertical = direction switch
        {
            "up" or "down" => true, "left" or "right" => false,
            _ => throw new ToolException("direction must be up, down, left or right"),
        };
        bool back = direction is "up" or "left";
        var amount = (a.Str("amount") ?? "line").ToLowerInvariant();
        int count = a.Int("count") ?? 1;
        if (count is < 1 or > 100) throw new ToolException("count must be 1-100");
        var method = (a.Str("method") ?? "scrollbar").ToLowerInvariant();
        if (method is not ("scrollbar" or "wheel")) throw new ToolException("method must be scrollbar or wheel");

        var form = ResolveWindow(t, a.Str("window"));
        long hwnd = a.Str("control") is { Length: > 0 } hex ? FindControlByHandle(t, ParseHex(hex)).Hwnd : form.Hwnd;
        string cls = ClassOf(hwnd);
        bool isBar = cls.Contains("ScrollBar", StringComparison.OrdinalIgnoreCase);
        // A window's own scroll bars can be read here. A scroll bar control of the elevated IDE only
        // by the add-in, because UIPI blocks SBM_* messages (spike 17); a compiled EXE's directly.
        async Task<(int Min, int Max, int Pos)?> Position()
        {
            if (!isBar || !t.IsIde) return AppWindows.ScrollPos(hwnd, vertical, isBar);
            try { return await _env.Client(t.Ide!).ScrollPosAsync(hwnd); } catch (AddinErrorException) { return null; }
        }
        var before = await Position();
        if (method == "wheel")
        {
            if (!vertical) throw new ToolException("the wheel scrolls up or down only");
            if (amount != "line") throw new ToolException("with method wheel, use count (notches) instead of amount");
            var c = AppWindows.Controls(form.Hwnd).FirstOrDefault(x => x.Hwnd == hwnd);
            var (cx, cy) = c is null ? (form.Width / 2, form.Height / 2) : AppWindows.FormToClient(form.Hwnd, hwnd, c.Left + c.Width / 2, c.Top + c.Height / 2);
            Audit.Write("app_scroll", t.Pid, $"wheel {Hex(hwnd)} {(back ? count : -count)}");
            if (t.Ide is { } inst) await _env.Client(inst).MouseAsync(hwnd, "wheel", cx, cy, back ? count : -count);
            else LocalInput.Mouse(hwnd, "wheel", cx, cy, back ? count : -count);
        }
        else
        {
            int code = (amount, back) switch
            {
                ("line", true) => 0, ("line", false) => 1, ("page", true) => 2, ("page", false) => 3,
                ("end", true) => 6, ("end", false) => 7,
                _ => throw new ToolException("amount must be line, page or end"),
            };
            if (amount == "end") count = 1;
            Audit.Write("app_scroll", t.Pid, $"{direction} {amount} x{count} {Hex(hwnd)} ({cls})");
            if (t.Ide is { } inst) await _env.Client(inst).ScrollAsync(hwnd, vertical, code, count);
            else LocalInput.Scroll(hwnd, vertical, code, count, isBar);
        }
        await Task.Delay(400, ct);
        var after = await Position();
        return ToolResult.Ok(new
        {
            target = $"{AppWindows.Kind(cls)} {Hex(hwnd)}",
            before = before is { } b ? new { b.Min, b.Max, b.Pos } : null,
            after = after is { } f ? new { f.Min, f.Max, f.Pos } : null,
            moved = before is { } b2 && after is { } f2 ? b2.Pos != f2.Pos : (bool?)null,
            state = AfterState(t, $"scrolled {direction}"),
            note = before is null ? "the window has no standard scroll bar to read; check the effect with app_screenshot" : null,
        });
    }

    private async Task<ToolResult> AppSelectAsync(JsonObject a, CancellationToken ct)
    {
        var t = ResolveTarget(a);
        RequireReady(t, allowDialogTarget: false);
        async Task<(string Kind, IReadOnlyList<string> Items, int Selected)> StateOf(long h)
        {
            bool combo = ClassOf(h).Contains("ComboBox", StringComparison.OrdinalIgnoreCase);
            return t.Ide is { } inst ? await _env.Client(inst).ListStateAsync(h) : await Task.Run(() => LocalInput.ListState(h, combo), ct);
        }
        // Without control (benchmark T17: three calls with only text): the form's only list or combo box, or with text the
        // only one that has that item (the Shop form has the product combo and the list of order lines)
        long hwnd;
        if (a.Str("control") is { Length: > 0 } hex) hwnd = FindControlByHandle(t, ParseHex(hex)).Hwnd;
        else
        {
            var lists = AppWindows.Controls(ResolveWindow(t, a.Str("window")).Hwnd).Where(c => c.Kind is "ListBox" or "ComboBox").ToList();
            if (lists.Count > 1 && a.Str("text") is { } wanted)
            {
                var having = new List<AppControl>();
                foreach (var l in lists)
                    if ((await StateOf(l.Hwnd)).Items.Any(s => s.Equals(wanted, StringComparison.OrdinalIgnoreCase))) having.Add(l);
                if (having.Count > 0) lists = having;
            }
            hwnd = lists.Count == 1 ? lists[0].Hwnd
                : throw new ToolException(lists.Count == 0 ? "the form has no list or combo box" : $"the form has {lists.Count} list and combo boxes{(a.Str("text") is null ? "" : $" with '{a.Str("text")}'")} ({string.Join(", ", lists.Select(l => $"{l.Kind} {Hex(l.Hwnd)}"))}); pass control");
        }
        bool isCombo = ClassOf(hwnd).Contains("ComboBox", StringComparison.OrdinalIgnoreCase);
        if (!isCombo && !ClassOf(hwnd).Contains("ListBox", StringComparison.OrdinalIgnoreCase))
            throw new ToolException($"control {Hex(hwnd)} is a {AppWindows.Kind(ClassOf(hwnd))}, not a list box or combo box");
        Task<(string Kind, IReadOnlyList<string> Items, int Selected)> State() => StateOf(hwnd);
        var (kind, items, before) = await State();
        int index;
        if (a.Int("index") is int i) index = i;
        else if (a.Str("text") is { } text)
        {
            var matches = items.Select((s, n) => (s, n)).Where(x => x.s.Equals(text, StringComparison.OrdinalIgnoreCase)).ToList();
            index = matches.Count switch
            {
                0 => throw new ToolException($"no item '{text}'. Items: {string.Join(", ", items.Take(30).Select(s => $"'{s}'"))}{(items.Count > 30 ? ", ..." : "")}"),
                1 => matches[0].n,
                _ => throw new ToolException($"{matches.Count} items are '{text}'; pass index ({string.Join(", ", matches.Select(m => m.n))})"),
            };
        }
        else throw new ToolException("pass text or index");
        if (index < 0 || index >= items.Count) throw new ToolException($"index {index} is out of range ({items.Count} items)");
        Audit.Write("app_select", t.Pid, $"{Hex(hwnd)} [{index}] '{items[index]}'");
        if (t.Ide is { } inst2) await _env.Client(inst2).SelectAsync(hwnd, index);
        else await Task.Run(() => LocalInput.Select(hwnd, index, isCombo), ct);
        await Task.Delay(400, ct);
        var (_, _, now) = await State();
        return ToolResult.Ok(new
        {
            control = Hex(hwnd), kind, index, item = items[index], selectedBefore = before, selectedNow = now, applied = now == index,
            state = AfterState(t, $"selected '{items[index]}'"),
        });
    }

    private async Task<ToolResult> AppKeysAsync(JsonObject a, CancellationToken ct)
    {
        var t = ResolveTarget(a);
        RequireReady(t, allowDialogTarget: TargetsDialog(t, a));
        string events;
        try { events = KeySyntax.ToEvents(a.Req("keys"), a.Bool("real_input") == true ? LocalInput.KeyFor : null); }
        catch (FormatException ex) { throw new ToolException(ex.Message); }
        // control may also be a form (app_windows): a form with only Labels has no control to type into (improvement notes 54)
        long hwnd = a.Str("control") is { Length: > 0 } hex
            ? (AppWindows.List(t.Pid).FirstOrDefault(w => w.Hwnd == ParseHex(hex)) is { } formTarget ? formTarget.Hwnd : FindControlByHandle(t, ParseHex(hex)).Hwnd)
            : 0;
        bool real = a.Bool("real_input") ?? false;
        string? realNote = null;
        if (hwnd == 0 && t.IsIde) hwnd = ProgramFocusInIde(t, real);
        Audit.Write("app_keys", t.Pid, $"{(real ? "real " : "")}{(hwnd == 0 ? "focused" : Hex(hwnd))} {a.Req("keys")}");
        long target;
        if (t.Ide is { } inst)
        {
            var client = _env.Client(inst);
            if (real)
            {
                string outcome;
                try
                {
                    target = await client.KeysAsync(hwnd, events, true);    // deferred in the add-in (hard rule 14)
                    await Task.Delay(500, ct);
                }
                finally { outcome = await client.MouseRestoreAsync(); }
                if (outcome != "sent") throw new ToolException("real input wasn't sent: " + outcome);
            }
            else { target = await client.KeysAsync(hwnd, events, false); await Task.Delay(400, ct); }
        }
        else
        {
            var exeForm = ResolveWindow(t, null);
            int thread = LocalInput.WindowThread(exeForm.Hwnd);
            if (hwnd == 0) hwnd = LocalInput.Focused(thread).ToInt64();
            if (hwnd == 0 && AppWindows.Controls(exeForm.Hwnd).Count == 0) hwnd = exeForm.Hwnd;   // only Labels: the form takes the keys
            if (hwnd == 0) throw new ToolException("no control of the program has the focus; pass the control (from app_tree)");
            target = hwnd;
            if (real)
            {
                var outcome = await Task.Run(() => LocalInput.Real(hwnd, h => LocalInput.RealKeys(h, events)), ct);
                if (!outcome.StartsWith("sent", StringComparison.Ordinal)) throw new ToolException("real input wasn't sent: " + outcome);
                realNote = outcome == "sent" ? null : outcome;
            }
            else
            {
                try { await Task.Run(() => LocalInput.Keys(hwnd, events, thread), ct); }
                catch (InvalidOperationException ex) { throw new ToolException(ex.Message); }
                await Task.Delay(300, ct);
            }
        }
        string? textNow = null;
        var cls = AppWindows.ClassOf(target);
        if (AppWindows.Kind(cls) is "TextBox" or "ComboBox") textNow = AppWindows.Text(new IntPtr(target));
        return ToolResult.Ok(new
        {
            target = $"{AppWindows.Kind(cls)} {Hex(target)}", method = real ? "real input (SendInput)" : "posted messages", textNow, warning = realNote,
            state = AfterState(t, "keys sent"),
            verify = "check the effect with app_tree, app_screenshot or app_wait",
        });
    }

    private async Task<ToolResult> AppWaitAsync(JsonObject a, CancellationToken ct)
    {
        var t = ResolveTarget(a);
        int timeout = Math.Clamp(a.Int("timeout") ?? 10, 1, 120);
        var window = a.Str("window"); bool gone = a.Bool("gone") ?? false;
        var control = a.Str("control") is { Length: > 0 } hex ? ParseHex(hex) : (long?)null;
        var text = a.Str("text"); var contains = a.Str("contains");
        var modeWanted = a.Str("mode")?.ToLowerInvariant();
        bool dialog = a.Bool("dialog") ?? false;
        if (window is null && control is null && modeWanted is null && !dialog) throw new ToolException("pass window, control (+ text/contains), mode or dialog");
        if (control != null && text is null && contains is null) throw new ToolException("with control, pass text or contains");
        if (modeWanted is not (null or "run" or "break" or "design")) throw new ToolException("mode must be run, break or design");
        if (modeWanted != null && !t.IsIde) throw new ToolException("mode is about the IDE; a compiled program has none");

        var start = DateTime.UtcNow;
        string? seen = null;
        HashSet<long>? dialogsAtStart = null;
        while (true)
        {
            var w = _env.ReadWindows(t.Pid);
            var dialogs = t.IsIde ? w.Dialogs : IdeWindows.Dialogs(t.Pid);
            dialogsAtStart ??= [.. dialogs.Select(d => d.Hwnd)];
            if (modeWanted != null && Mode(w.Mode) == modeWanted) seen = $"mode {modeWanted}";
            else if (dialog && dialogs.Count > 0) seen = $"dialog '{dialogs[0].Caption}'";
            else if (window != null)
            {
                var hit = AppWindows.List(t.Pid).FirstOrDefault(x => x.Caption.Contains(window, StringComparison.OrdinalIgnoreCase));
                if (gone ? hit is null : hit is not null) seen = gone ? $"window '{window}' gone" : $"window '{hit!.Caption}' ({Hex(hit.Hwnd)})";
            }
            else if (control is long h)
            {
                var now = AppWindows.Text(new IntPtr(h));
                if ((text != null && now == text) || (contains != null && now.Contains(contains, StringComparison.Ordinal))) seen = $"text '{now}'";
            }
            // improvement notes 32: a MsgBox that blocks the program won't let the awaited window appear: stop waiting
            var blocking = seen is null && !dialog ? dialogs.FirstOrDefault(d => !dialogsAtStart.Contains(d.Hwnd)) : null;
            if (seen != null || blocking != null || (DateTime.UtcNow - start).TotalSeconds >= timeout)
                return ToolResult.Ok(new
                {
                    met = seen != null, seen, waitedMs = (int)(DateTime.UtcNow - start).TotalMilliseconds,
                    mode = t.IsIde ? Mode(w.Mode) : null, dialogs = dialogs.Select(d => new { caption = d.Caption, text = string.Join("\n", d.Texts), buttons = d.Buttons.Select(b => b.Caption) }),
                    note = blocking != null ? $"stopped waiting: dialog '{blocking.Caption}' appeared, and the program waits for it to be answered"
                         : seen is null ? $"not seen within {timeout} s" : null,
                });
            await Task.Delay(200, ct);
        }
    }

    private async Task<ToolResult> AppScreenshotAsync(JsonObject a, CancellationToken ct)
    {
        var t = ResolveTarget(a);
        int maxWidth = a.Int("max_width") ?? 1024;
        var which = a.Str("window");
        long hwnd; string caption;
        if (string.Equals(which, "ide", StringComparison.OrdinalIgnoreCase))
        {
            if (!t.IsIde) throw new ToolException("window 'ide' is for the IDE, not a compiled program");
            hwnd = IdeWindows.MainWindow(t.Pid);
            if (hwnd == 0) throw new ToolException("the IDE main window wasn't found");
            caption = _env.ReadWindows(t.Pid).Title;
        }
        else { var w = ResolveWindow(t, which); hwnd = w.Hwnd; caption = w.Caption; }

        string method;
        // a compiled EXE and a non-elevated IDE can be captured by the server; an elevated IDE can't (spike 15)
        System.Drawing.Bitmap? bmp = t.IsIde && Vb6Install.IsElevated(t.Pid) == true ? null : Screenshots.CaptureLocal(hwnd);
        if (bmp != null) method = "server PrintWindow";
        else if (!t.IsIde) throw new ToolException("the window couldn't be captured");
        else
        {
            // Elevated IDE: capture in-process through the add-in (deferred), then wait for the file.
            // A busy or hung IDE can't run the deferred capture: then the watchdog captures (PLAN §3.1).
            string? path = null;
            if (_env.ReadWindows(t.Pid).Responds)
            {
                try
                {
                    string id = "s" + DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture);
                    path = await _env.Client(t.Ide!).SnapAsync(hwnd, id);
                    var deadline = DateTime.UtcNow + SnapWait;
                    while (!File.Exists(path) && DateTime.UtcNow < deadline) await Task.Delay(100, ct);
                }
                catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException) { path = null; }
            }
            if (path != null && File.Exists(path))
            {
                bmp = Screenshots.LoadBmp(path);
                try { File.Delete(path); } catch (IOException) { }
                method = "add-in PrintWindow (elevated IDE)";
            }
            else
            {
                var viaWd = await SnapViaWatchdogAsync(hwnd, maxWidth, ct)
                            ?? throw new ToolException("the IDE didn't produce the screenshot (busy or hung?) and " + NoWatchdog);
                return ToolResult.Image(viaWd.Png, new { window = caption, hwnd = Hex(hwnd), method = viaWd.Method, clientOrigin = ClientOrigin(hwnd), note = CoordinatesNote });
            }
        }
        using (bmp)
        {
            var png = Screenshots.ToPng(bmp, maxWidth);
            double scale = maxWidth > 0 && bmp.Width > maxWidth ? (double)maxWidth / bmp.Width : 1;
            return ToolResult.Image(png, new { window = caption, hwnd = Hex(hwnd), width = bmp.Width, height = bmp.Height, method,
                                               clientOrigin = ClientOrigin(hwnd), imageScale = scale < 1 ? Math.Round(scale, 4) : (double?)null, note = CoordinatesNote });
        }
    }

    // improvement notes 35: the image is the whole window, app_mouse takes client coordinates; a click read off the
    // image missed by the title bar and border (about 8, 31 px) and hit another control
    private static object ClientOrigin(long hwnd) { var (x, y) = AppWindows.ClientOffset(hwnd); return new { x, y }; }

    private const string CoordinatesNote = "the image shows the whole window; for app_mouse use x = image x / imageScale - clientOrigin.x (same for y)";

    // ---------- helpers ----------

    private static string ClassOf(long hwnd) => AppWindows.ClassOf(hwnd);

    /// <summary>The program to act on: exe (a process exe_run started) or the IDE's program.</summary>
    private Target ResolveTarget(JsonObject a)
    {
        if (a.Int("exe") is int pid)
        {
            if (!_exes.TryGetValue(pid, out var path))
                throw new ToolException($"process {pid} wasn't started with exe_run; input only goes to programs the agent started (hard rule 8)");
            if (!IsAlive(pid)) { _exes.TryRemove(pid, out _); throw new ToolException($"{Path.GetFileName(path)} (pid {pid}) has exited"); }
            return new Target(pid, null, path);
        }
        var inst = ResolveInstance(a);
        return new Target(inst.Pid, inst, null);
    }

    /// <summary>
    /// app_keys without control on a program running in the IDE: the program shares the IDE's UI thread, so "the focused
    /// control" can be the Immediate window or a code pane (improvement notes 51: {F6} went to the Immediate window after
    /// debug_immediate). Returns the focused control when it is in one of the program's windows; otherwise, for real input,
    /// the program's form (bringing it to the front gives the focus back to its own control); posted input is refused.
    /// </summary>
    private long ProgramFocusInIde(Target t, bool real)
    {
        var form = ResolveWindow(t, null);
        long focused = LocalInput.Focused(LocalInput.WindowThread(form.Hwnd)).ToInt64();
        var root = LocalInput.Root(focused);
        if (focused != 0 && AppWindows.List(t.Pid).Any(w => w.Hwnd == root)) return focused;
        // a form without windowed controls (only Labels: "Press any key") takes the keys itself (Form_KeyDown/KeyPress)
        if (real || AppWindows.Controls(form.Hwnd).Count == 0) return form.Hwnd;
        throw new ToolException("the keyboard focus is in the IDE (" + (focused == 0 ? "no focused window" : $"window {Hex(focused)}") +
            "), not in the program, so posted keys would go to VB6 itself; pass control (from app_tree) or real_input=true (brings the program's form to the front)");
    }

    private AppWindow ResolveWindow(Target t, string? which)
    {
        var windows = AppWindows.List(t.Pid);
        if (windows.Count == 0)
            throw new ToolException(t.IsIde
                ? $"the program has no visible windows (IDE mode: {Mode(_env.ReadWindows(t.Pid).Mode)}; start it with run_start)"
                : $"{Path.GetFileName(t.Exe)} has no visible windows (yet)");
        if (which is { Length: > 0 })
        {
            if (TryParseHex(which, out var h) && windows.FirstOrDefault(w => w.Hwnd == h) is { } byHandle) return byHandle;
            return windows.FirstOrDefault(w => w.Caption.Equals(which, StringComparison.OrdinalIgnoreCase))
                   ?? throw new ToolException($"no window '{which}'. Open: {string.Join(", ", windows.Select(w => $"'{w.Caption}' ({Hex(w.Hwnd)})"))}");
        }
        var forms = windows.Where(w => w.Kind is "Form" or "MDIForm").ToList();
        return forms.Count >= 1 ? forms[0] : windows[0];
    }

    /// <summary>What a user could see change: the program's windows and dialogs with their controls' text and states.</summary>
    private async Task<string> UiFingerprintAsync(Target t)
    {
        try
        {
            var sb = new System.Text.StringBuilder();
            foreach (var w in AppWindows.List(t.Pid))
            {
                sb.Append(w.Hwnd).Append('|').Append(w.Caption).Append('|').Append(w.Enabled).Append('\n');
                foreach (var c in AppWindows.Controls(w.Hwnd))
                {
                    sb.Append(c.Hwnd).Append('|').Append(c.Text).Append('|').Append(c.Enabled).Append(c.Visible).Append(c.Left).Append(',').Append(c.Top);
                    // UIPI blocks LB_*/CB_*/BM_* from the server to an elevated IDE (live check of notes 34: a log ListBox
                    // grew and the click looked like "no change"). Lists go through the add-in there; check states are
                    // read only for compiled programs, which run at the server's integrity.
                    if (c.Kind is "ListBox" or "ComboBox")
                        sb.Append('|').Append(t.Ide is { } inst ? await ListFingerprintAsync(inst, c.Hwnd) : AppWindows.ItemCount(c.Hwnd, c.Kind).ToString(CultureInfo.InvariantCulture));
                    else if (c.Kind is "CheckBox" or "OptionButton" && !t.IsIde) sb.Append('|').Append(AppWindows.CheckState(c.Hwnd));
                    sb.Append('\n');
                }
            }
            return sb.ToString();
        }
        catch (Exception ex) when (ex is ToolException or InvalidOperationException) { return ""; }
    }

    private async Task<string> ListFingerprintAsync(InstanceInfo inst, long hwnd)
    {
        try { var (_, items, selected) = await _env.Client(inst).ListStateAsync(hwnd); return items.Count + ":" + selected; }
        catch (Exception ex) when (ex is AddinErrorException or AddinBusyException or AddinUnreachableException) { return "?"; }
    }

    /// <summary>Whether a user could see (and so click) the control: its window is visible and not clipped away.</summary>
    private static bool Seen(AppControl c) => c.Visible && c.Clipped != "hidden";

    private static AppControl FindControlByHandle(Target t, long hwnd)
    {
        foreach (var w in AppWindows.List(t.Pid))
            if (AppWindows.Controls(w.Hwnd).FirstOrDefault(c => c.Hwnd == hwnd) is { } c) return c;
        throw new ToolException($"control {Hex(hwnd)} wasn't found in the program's windows (see app_tree)");
    }

    /// <summary>The IDE must be in run mode; a compiled program must still run. No open dialog unless allowed.</summary>
    /// <summary>
    /// Whether the call is aimed at an open dialog itself (improvement notes 36: the program's own InputBox): window names
    /// it (caption or handle), or control is one of its child windows. Then the dialog isn't in the way, it's the target.
    /// </summary>
    private bool TargetsDialog(Target t, JsonObject a)
    {
        var dialogs = t.IsIde ? _env.ReadWindows(t.Pid).Dialogs : IdeWindows.Dialogs(t.Pid);
        if (dialogs.Count == 0) return false;
        if (a.Str("window") is { Length: > 0 } w &&
            dialogs.Any(d => d.Caption.Equals(w, StringComparison.OrdinalIgnoreCase) || (TryParseHex(w, out var h) && d.Hwnd == h))) return true;
        return a.Str("control") is { Length: > 0 } c && TryParseHex(c, out var ch) && dialogs.Any(d => d.Hwnd == LocalInput.Root(ch));
    }

    private void RequireReady(Target t, bool allowDialogTarget)
    {
        if (t.IsIde)
        {
            var w = _env.ReadWindows(t.Pid);
            if (w.Mode != IdeMode.Run)
                throw new ToolException($"the IDE is in {Mode(w.Mode)} mode; the program must be running (run_start, or run_continue from break)");
            if (!allowDialogTarget && w.Dialogs.Count > 0)
                throw new ToolException($"a dialog is open ('{w.Dialogs[0].Caption}'); answer it first (ide_dialog_respond)");
        }
        else if (!allowDialogTarget && IdeWindows.Dialogs(t.Pid) is { Count: > 0 } d)
            throw new ToolException($"the program shows a dialog ('{d[0].Caption}'); answer it first (app_click on its button)");
    }

    private object AfterState(Target t, string what)
    {
        if (!t.IsIde)
        {
            bool alive = IsAlive(t.Pid);
            return new
            {
                action = what, running = alive,
                dialogs = alive ? IdeWindows.Dialogs(t.Pid).Select(d => new { caption = d.Caption, text = string.Join("\n", d.Texts), buttons = d.Buttons.Select(b => b.Caption) }) : null,
                windows = alive ? AppWindows.List(t.Pid).Where(x => x.Kind != "Dialog").Select(x => new { hwnd = Hex(x.Hwnd), x.Kind, x.Caption }) : null,
                note = alive ? null : "the program has exited",
            };
        }
        var w = _env.ReadWindows(t.Pid);
        return new
        {
            action = what, mode = Mode(w.Mode), responds = w.Responds, dialogs = Dialogs(w),
            windows = AppWindows.List(t.Pid).Where(x => x.Kind != "Dialog").Select(x => new { hwnd = Hex(x.Hwnd), x.Kind, x.Caption }),
            note = w.Mode == IdeMode.Break ? "the program stopped in break mode (breakpoint or error?): debug_where"
                 : !w.Responds ? "the IDE isn't responding (the program may be busy)" : null,
        };
    }

    private static string Hex(long h) => h.ToString("X", CultureInfo.InvariantCulture);
    // Agents pass control names (txtPrice) here (benchmark T16, 3 failed calls): a running program's control names
    // aren't visible from outside it, so say where the handle comes from
    private static long ParseHex(string s) => TryParseHex(s, out var v) ? v
        : throw new ToolException($"'{s}' isn't a window handle (hex). Control names can't be seen from outside the running program: take the control's hwnd from app_tree (kind, text and position tell them apart)");
    private static bool TryParseHex(string s, out long v)
        => long.TryParse(s.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? s[2..] : s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v);
}
