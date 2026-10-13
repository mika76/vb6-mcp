using System.Globalization;
using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 3: drive the program running in the IDE. Reads (windows, controls, text) are done by the
/// server: they cross the elevation boundary (spike 10). Input goes through the add-in, inside the IDE
/// process, and is audited (hard rule 8). Screenshots: the server captures itself when VB6 isn't
/// elevated, otherwise the add-in captures in-process (spike 15). Documented in docs/tools.md.
/// </summary>
public sealed partial class IdeTools
{
    private const string WindowDesc = "Form or dialog: window handle (hex, from app_windows) or caption. Default: the program's only/first form.";

    /// <summary>How long app_screenshot waits for the add-in's capture file.</summary>
    public TimeSpan SnapWait { get; init; } = TimeSpan.FromSeconds(5);

    private IReadOnlyList<Tool> AppTools() =>
    [
        new("app_windows",
            "List the windows (forms, dialogs) of the program running in the IDE, with caption, class, position and size.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, _) => Task.FromResult(AppWindowsList(a))),
        new("app_tree",
            "Controls of a form of the running program: window handle, kind (CommandButton, TextBox, ...), text, position (relative to the " +
            "form's client area), visible, enabled. Lightweight controls (Label, Image, Shape, Line) have no window and only show in app_screenshot.",
            Schema.Object(("window", "string", WindowDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => Task.FromResult(AppTree(a))),
        new("app_click",
            "Click a command button, check box or option button of the running program. Address it by control (hex handle from app_tree) " +
            "or by text (its caption, optionally within window). Returns the state afterwards (mode, dialogs, windows).",
            Schema.Object(("control", "string", "Control window handle (hex) from app_tree.", false), ("text", "string", "Control caption, e.g. Add or &Add.", false),
                          ("window", "string", WindowDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => AppClickAsync(a, ct)),
        new("app_set_text",
            "Set the text of a text box in the running program (fires its Change event). Address it by control (hex handle from app_tree) " +
            "or by index (0-based among the form's text boxes, in app_tree order). Returns the text read back.",
            Schema.Object(("text", "string", "New text.", true), ("control", "string", "Text box window handle (hex).", false),
                          ("index", "integer", "0-based index among the form's text boxes.", false), ("window", "string", WindowDesc, false),
                          ("pid", "integer", PidDesc, false)),
            (a, ct) => AppSetTextAsync(a, ct)),
        new("app_mouse",
            "Mouse input at a position in a form of the running program: click, double_click, right_click, or drag (from x,y to to_x,to_y). " +
            "x/y are pixels relative to the form's client area (the coordinates app_tree reports). By default the messages are posted to the " +
            "window under the point, without touching the user's mouse; that works for windowed controls (buttons, lists, picture boxes). " +
            "Labels and Images (no window) only react to real input: pass real_input=true, which brings the form to the front, moves the " +
            "real cursor, clicks, then puts the cursor and the active window back.",
            Schema.Object(("action", "string", "click, double_click, right_click or drag.", true), ("x", "integer", "X in the form's client area.", true),
                          ("y", "integer", "Y in the form's client area.", true), ("to_x", "integer", "Drag end X.", false), ("to_y", "integer", "Drag end Y.", false),
                          ("real_input", "boolean", "Use real mouse input (needed for Labels/Images): takes the foreground and moves the cursor briefly.", false),
                          ("window", "string", WindowDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => AppMouseAsync(a, ct)),
        new("app_scroll",
            "Scroll a control of the running program (list box, multi-line text box, grid, scroll bar control, or a scrollable form). " +
            "direction up/down/left/right; amount line (default), page or end; count repeats (default 1). method 'scrollbar' (default) sends " +
            "scroll bar commands, 'wheel' sends mouse-wheel notches at the control's centre. Returns the scroll position before and after.",
            Schema.Object(("control", "string", "Window handle (hex) from app_tree. Default: the form itself.", false),
                          ("direction", "string", "up, down, left or right.", true), ("amount", "string", "line (default), page or end.", false),
                          ("count", "integer", "Repeat count, 1-100 (default 1).", false), ("method", "string", "scrollbar (default) or wheel.", false),
                          ("window", "string", WindowDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => AppScrollAsync(a, ct)),
        new("app_screenshot",
            "Screenshot of a window as an image: a form or dialog of the running program (default), or the whole IDE (target 'ide'). " +
            "Shows what app_tree can't: Labels, Images, drawing, layout. Works when the window is covered.",
            Schema.Object(("window", "string", WindowDesc + " Or 'ide' for the IDE main window.", false),
                          ("max_width", "integer", "Scale down to at most this width (default 1024).", false), ("pid", "integer", PidDesc, false)),
            (a, ct) => AppScreenshotAsync(a, ct)),
    ];

    private ToolResult AppWindowsList(JsonObject a)
    {
        var inst = ResolveInstance(a);
        var w = _env.ReadWindows(inst.Pid);
        var windows = AppWindows.List(inst.Pid);
        return ToolResult.Ok(new
        {
            mode = Mode(w.Mode),
            windows = windows.Select(x => new { hwnd = Hex(x.Hwnd), x.Kind, x.Caption, x.Class, x.Left, x.Top, x.Width, x.Height, x.Enabled }),
            note = w.Mode == IdeMode.Design ? "the IDE is in design mode: the program isn't running (run_start)" : null,
        });
    }

    private ToolResult AppTree(JsonObject a)
    {
        var inst = ResolveInstance(a);
        var form = ResolveWindow(inst, a.Str("window"));
        var controls = AppWindows.Controls(form.Hwnd);
        return ToolResult.Ok(new
        {
            window = new { hwnd = Hex(form.Hwnd), form.Kind, form.Caption },
            controls = controls.Select(c => new { hwnd = Hex(c.Hwnd), c.Kind, c.Text, c.Left, c.Top, c.Width, c.Height, c.Visible, c.Enabled, c.Depth, c.Class }),
            note = "Labels, Images, Shapes and Lines have no window and aren't listed; use app_screenshot to see them",
        });
    }

    private async Task<ToolResult> AppClickAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireRunningWithoutDialog(inst, allowDialogTarget: true);
        AppControl target;
        if (a.Str("control") is { Length: > 0 } hex) target = FindControlByHandle(inst, ParseHex(hex));
        else
        {
            var text = a.Str("text") ?? throw new ToolException("pass control (handle from app_tree) or text (the caption)");
            var form = ResolveWindow(inst, a.Str("window"));
            var wanted = text.Replace("&", "", StringComparison.Ordinal);
            var matches = AppWindows.Controls(form.Hwnd)
                .Where(c => c.Kind is "CommandButton" or "CheckBox" or "OptionButton" && c.Text.Replace("&", "", StringComparison.Ordinal).Equals(wanted, StringComparison.OrdinalIgnoreCase)).ToList();
            target = matches.Count switch
            {
                0 => throw new ToolException($"no button, check box or option button captioned '{text}' in '{form.Caption}' (see app_tree)"),
                1 => matches[0],
                _ => throw new ToolException($"{matches.Count} controls are captioned '{text}'; pass control (handle) instead"),
            };
        }
        if (!target.Enabled) throw new ToolException($"'{target.Text}' is disabled");
        Audit.Write("app_click", inst.Pid, $"{Hex(target.Hwnd)} {target.Kind} '{target.Text}'");
        await _env.Client(inst).AppClickAsync(target.Hwnd);
        await Task.Delay(500, ct);              // the click and its event code run after the posted message
        return ToolResult.Ok(AfterState(inst, $"clicked {target.Kind} '{target.Text}'"));
    }

    private async Task<ToolResult> AppSetTextAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireRunningWithoutDialog(inst, allowDialogTarget: false);
        var text = a.Str("text") ?? throw new ToolException("missing required argument 'text'");
        AppControl target;
        if (a.Str("control") is { Length: > 0 } hex) target = FindControlByHandle(inst, ParseHex(hex));
        else
        {
            var form = ResolveWindow(inst, a.Str("window"));
            var boxes = AppWindows.Controls(form.Hwnd).Where(c => c.Kind == "TextBox").ToList();
            int index = a.Int("index") ?? (boxes.Count == 1 ? 0 : throw new ToolException($"the form has {boxes.Count} text boxes; pass index or control"));
            if (index < 0 || index >= boxes.Count) throw new ToolException($"index {index} is out of range (the form has {boxes.Count} text boxes)");
            target = boxes[index];
        }
        if (target.Kind != "TextBox") throw new ToolException($"control {Hex(target.Hwnd)} is a {target.Kind}, not a text box");
        Audit.Write("app_set_text", inst.Pid, $"{Hex(target.Hwnd)} '{text}'");
        await _env.Client(inst).AppTextAsync(target.Hwnd, text);
        await Task.Delay(400, ct);
        var now = AppWindows.Text(new IntPtr(target.Hwnd));
        var state = AfterState(inst, "text set");
        return ToolResult.Ok(new { textNow = now, applied = now == text, state });
    }

    private async Task<ToolResult> AppScreenshotAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        int maxWidth = a.Int("max_width") ?? 1024;
        var which = a.Str("window");
        long hwnd; string caption;
        if (string.Equals(which, "ide", StringComparison.OrdinalIgnoreCase))
        {
            hwnd = IdeWindows.MainWindow(inst.Pid);
            if (hwnd == 0) throw new ToolException("the IDE main window wasn't found");
            caption = _env.ReadWindows(inst.Pid).Title;
        }
        else { var w = ResolveWindow(inst, which); hwnd = w.Hwnd; caption = w.Caption; }

        string method;
        System.Drawing.Bitmap? bmp = Vb6Install.IsElevated(inst.Pid) == true ? null : Screenshots.CaptureLocal(hwnd);
        if (bmp != null) method = "server PrintWindow";
        else
        {
            // Elevated IDE: capture in-process through the add-in (deferred), then wait for the file.
            // A busy or hung IDE can't run the deferred capture: then the watchdog captures (PLAN §3.1).
            string? path = null;
            if (_env.ReadWindows(inst.Pid).Responds)
            {
                try
                {
                    string id = "s" + DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture);
                    path = await _env.Client(inst).SnapAsync(hwnd, id);
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
                return ToolResult.Image(viaWd.Png, new { window = caption, hwnd = Hex(hwnd), method = viaWd.Method });
            }
        }
        using (bmp)
        {
            var png = Screenshots.ToPng(bmp, maxWidth);
            return ToolResult.Image(png, new { window = caption, hwnd = Hex(hwnd), width = bmp.Width, height = bmp.Height, method });
        }
    }

    private async Task<ToolResult> AppMouseAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireRunningWithoutDialog(inst, allowDialogTarget: false);
        var action = (a.Str("action") ?? "").ToLowerInvariant();
        string verb = action switch
        {
            "click" => "click", "double_click" => "dblclick", "right_click" => "rclick", "drag" => "drag",
            _ => throw new ToolException("action must be click, double_click, right_click or drag"),
        };
        int x = a.Int("x") ?? throw new ToolException("missing required argument 'x'");
        int y = a.Int("y") ?? throw new ToolException("missing required argument 'y'");
        var form = ResolveWindow(inst, a.Str("window"));
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
        var client = _env.Client(inst);
        Audit.Write("app_mouse", inst.Pid, $"{(real ? "real " : "")}{verb} {Hex(hwnd)} ({lx},{ly}){(ex is null ? "" : $"->({ex},{ey})")}");
        if (real)
        {
            // Spike 17: windowless controls only react to real input. The add-in can't wait for the
            // program to process it (hard rule 5), so the server waits, then always restores.
            try
            {
                await client.MouseRealAsync(hwnd, verb, lx, ly, ex, ey);
                await Task.Delay(400, ct);
            }
            finally { await client.MouseRestoreAsync(); }
        }
        else
        {
            await client.MouseAsync(hwnd, verb, lx, ly, ex, ey);
            await Task.Delay(500, ct);
        }
        return ToolResult.Ok(new
        {
            target = under, at = new { x = lx, y = ly }, method = real ? "real input (SendInput); cursor and active window restored" : "posted messages",
            state = AfterState(inst, $"{action} at ({x},{y})"),
            verify = real ? "check the effect with app_screenshot or app_tree"
                          : "posted input has no visible result of its own: check with app_screenshot or app_tree. A Label or Image ignores posted input; use real_input=true (spike 17)",
        });
    }

    private async Task<ToolResult> AppScrollAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireRunningWithoutDialog(inst, allowDialogTarget: false);
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

        var form = ResolveWindow(inst, a.Str("window"));
        long hwnd = a.Str("control") is { Length: > 0 } hex ? FindControlByHandle(inst, ParseHex(hex)).Hwnd : form.Hwnd;
        string cls = ClassOf(hwnd);
        bool isBar = cls.Contains("ScrollBar", StringComparison.OrdinalIgnoreCase);
        var client = _env.Client(inst);
        // a window's own scroll bars can be read here; a scroll bar control only by the add-in, because
        // UIPI blocks SBM_* messages to the elevated IDE (spike 17 notes)
        async Task<(int Min, int Max, int Pos)?> Position()
        {
            if (!isBar) return AppWindows.ScrollPos(hwnd, vertical, false);
            try { return await client.ScrollPosAsync(hwnd); } catch (AddinErrorException) { return null; }
        }
        var before = await Position();
        if (method == "wheel")
        {
            if (!vertical) throw new ToolException("the wheel scrolls up or down only");
            if (amount != "line") throw new ToolException("with method wheel, use count (notches) instead of amount");
            var c = AppWindows.Controls(form.Hwnd).FirstOrDefault(x => x.Hwnd == hwnd);
            var (cx, cy) = c is null ? (form.Width / 2, form.Height / 2) : AppWindows.FormToClient(form.Hwnd, hwnd, c.Left + c.Width / 2, c.Top + c.Height / 2);
            Audit.Write("app_scroll", inst.Pid, $"wheel {Hex(hwnd)} {(back ? count : -count)}");
            await client.MouseAsync(hwnd, "wheel", cx, cy, back ? count : -count);
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
            Audit.Write("app_scroll", inst.Pid, $"{direction} {amount} x{count} {Hex(hwnd)} ({cls})");
            await client.ScrollAsync(hwnd, vertical, code, count);
        }
        await Task.Delay(400, ct);
        var after = await Position();
        return ToolResult.Ok(new
        {
            target = $"{AppWindows.Kind(cls)} {Hex(hwnd)}",
            before = before is { } b ? new { b.Min, b.Max, b.Pos } : null,
            after = after is { } f ? new { f.Min, f.Max, f.Pos } : null,
            moved = before is { } b2 && after is { } f2 ? b2.Pos != f2.Pos : (bool?)null,
            state = AfterState(inst, $"scrolled {direction}"),
            note = before is null ? "the window has no standard scroll bar to read; check the effect with app_screenshot" : null,
        });
    }

    // ---------- helpers ----------

    private static string ClassOf(long hwnd) => AppWindows.ClassOf(hwnd);

    private AppWindow ResolveWindow(InstanceInfo inst, string? which)
    {
        var windows = AppWindows.List(inst.Pid);
        if (windows.Count == 0)
            throw new ToolException($"the program has no visible windows (IDE mode: {Mode(_env.ReadWindows(inst.Pid).Mode)}; start it with run_start)");
        if (which is { Length: > 0 })
        {
            if (TryParseHex(which, out var h) && windows.FirstOrDefault(w => w.Hwnd == h) is { } byHandle) return byHandle;
            return windows.FirstOrDefault(w => w.Caption.Equals(which, StringComparison.OrdinalIgnoreCase))
                   ?? throw new ToolException($"no window '{which}'. Open: {string.Join(", ", windows.Select(w => $"'{w.Caption}' ({Hex(w.Hwnd)})"))}");
        }
        var forms = windows.Where(w => w.Kind is "Form" or "MDIForm").ToList();
        return forms.Count >= 1 ? forms[0] : windows[0];
    }

    private static AppControl FindControlByHandle(InstanceInfo inst, long hwnd)
    {
        foreach (var w in AppWindows.List(inst.Pid))
            if (AppWindows.Controls(w.Hwnd).FirstOrDefault(c => c.Hwnd == hwnd) is { } c) return c;
        throw new ToolException($"control {Hex(hwnd)} wasn't found in the program's windows (see app_tree)");
    }

    private void RequireRunningWithoutDialog(InstanceInfo inst, bool allowDialogTarget)
    {
        var w = _env.ReadWindows(inst.Pid);
        if (w.Mode != IdeMode.Run)
            throw new ToolException($"the IDE is in {Mode(w.Mode)} mode; the program must be running (run_start, or run_continue from break)");
        if (!allowDialogTarget && w.Dialogs.Count > 0)
            throw new ToolException($"a dialog is open ('{w.Dialogs[0].Caption}'); answer it first (ide_dialog_respond)");
    }

    private object AfterState(InstanceInfo inst, string what)
    {
        var w = _env.ReadWindows(inst.Pid);
        return new
        {
            action = what, mode = Mode(w.Mode), responds = w.Responds, dialogs = Dialogs(w),
            windows = AppWindows.List(inst.Pid).Where(x => x.Kind != "Dialog").Select(x => new { hwnd = Hex(x.Hwnd), x.Kind, x.Caption }),
            note = w.Mode == IdeMode.Break ? "the program stopped in break mode (breakpoint or error?): debug_where"
                 : !w.Responds ? "the IDE isn't responding (the program may be busy)" : null,
        };
    }

    private static string Hex(long h) => h.ToString("X", CultureInfo.InvariantCulture);
    private static long ParseHex(string s) => TryParseHex(s, out var v) ? v : throw new ToolException($"'{s}' isn't a window handle (hex)");
    private static bool TryParseHex(string s, out long v)
        => long.TryParse(s.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? s[2..] : s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v);
}
