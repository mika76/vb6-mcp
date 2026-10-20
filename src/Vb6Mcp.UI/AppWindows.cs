using System.Runtime.InteropServices;
using System.Text;

namespace Vb6Mcp.UI;

/// <summary>A top-level window of the running program (form or dialog).</summary>
public sealed record AppWindow(long Hwnd, string Class, string Kind, string Caption, int Left, int Top, int Width, int Height, bool Enabled);

/// <summary>A control window inside a form. Lightweight controls (Label, Image, Shape, Line) have no window and don't appear.</summary>
public sealed record AppControl(long Hwnd, long Parent, string Class, string Kind, string Text, int Left, int Top, int Width, int Height, bool Visible, bool Enabled, int Depth, string? Clipped = null);

/// <summary>
/// Channel C reads of the running program's windows. A program run from the IDE lives inside
/// VB6.EXE ("Thunder*" classes); a compiled EXE uses "ThunderRT6*" classes. Reads work from the
/// non-elevated server (spike 10); input goes through the add-in.
/// </summary>
public static class AppWindows
{
    /// <summary>IDE windows that are never part of the running program.</summary>
    private static readonly string[] s_ideClasses =
        ["wndclass_desked_gsk", "IDEOwner", "CompatWindowDesktopReplacement", "Static", "VBFloatingPalette", "ThunderMain", "ThunderRT6Main"];

    /// <summary>Visible forms and dialogs of the program running in process <paramref name="pid"/>.</summary>
    public static IReadOnlyList<AppWindow> List(int pid)
    {
        var list = new List<AppWindow>();
        EnumWindows((h, _) =>
        {
            GetWindowThreadProcessId(h, out var p);
            if (p != (uint)pid || !IsWindowVisible(h)) return true;
            string cls = ClassName(h);
            if (s_ideClasses.Contains(cls)) return true;
            GetWindowRect(h, out var r);
            list.Add(new AppWindow(h.ToInt64(), cls, Kind(cls), Caption(h), r.L, r.T, r.R - r.L, r.B - r.T, IsWindowEnabled(h)));
            return true;
        }, IntPtr.Zero);
        return list;
    }

    /// <summary>Child control windows of a form, depth-first, with positions relative to the form's client area.</summary>
    public static IReadOnlyList<AppControl> Controls(long form)
    {
        var root = new IntPtr(form);
        var origin = new POINT();
        ClientToScreen(root, ref origin);
        var list = new List<AppControl>();
        // clip: the part of the screen every ancestor's client area lets through. A control outside it can't be seen or
        // clicked by the user, though its window is "visible" (improvement notes 33: XmPanel hides buttons below its
        // bottom edge; 52: a control inside a 2 x 2 PictureBox).
        void Walk(IntPtr parent, int depth, (int L, int T, int R, int B) clip)
        {
            for (var c = GetWindow(parent, 5 /*GW_CHILD*/); c != IntPtr.Zero; c = GetWindow(c, 2 /*GW_HWNDNEXT*/))
            {
                string cls = ClassName(c);
                GetWindowRect(c, out var r);
                var rect = (r.L, r.T, r.R, r.B);
                list.Add(new AppControl(c.ToInt64(), parent.ToInt64(), cls, Kind(cls), Text(c), r.L - origin.X, r.T - origin.Y,
                                        r.R - r.L, r.B - r.T, IsWindowVisible(c), IsWindowEnabled(c), depth, ClipState(rect, clip)));
                Walk(c, depth + 1, Intersect(clip, ClientScreenRect(c)));
            }
        }
        Walk(root, 0, ClientScreenRect(root));
        return list;
    }

    /// <summary>BM_GETCHECK of a check box or option button: 0 unchecked, 1 checked, 2 grayed; -1 when it doesn't answer.</summary>
    public static int CheckState(long hwnd) =>
        SendMessageTimeout(new IntPtr(hwnd), 0x00F0 /*BM_GETCHECK*/, IntPtr.Zero, IntPtr.Zero, 2 /*SMTO_ABORTIFHUNG*/, 500, out var r) == IntPtr.Zero ? -1 : (int)r;

    /// <summary>Items in a ListBox (LB_GETCOUNT) or ComboBox (CB_GETCOUNT); -1 when it doesn't answer or isn't one.</summary>
    public static int ItemCount(long hwnd, string kind)
    {
        uint msg = kind switch { "ListBox" => 0x018B, "ComboBox" => 0x0146, _ => 0 };
        if (msg == 0) return -1;
        return SendMessageTimeout(new IntPtr(hwnd), msg, IntPtr.Zero, IntPtr.Zero, 2 /*SMTO_ABORTIFHUNG*/, 500, out var r) == IntPtr.Zero ? -1 : (int)r;
    }

    /// <summary>Where <paramref name="control"/>'s window rectangle lies inside <paramref name="window"/>'s (both screen rectangles, so relative to the window's top-left corner).</summary>
    public static (int X, int Y, int Width, int Height) RectInWindow(long control, long window)
    {
        GetWindowRect(new IntPtr(control), out var c);
        GetWindowRect(new IntPtr(window), out var w);
        return (c.L - w.L, c.T - w.T, c.R - c.L, c.B - c.T);
    }

    /// <summary>Where a window's client area starts, relative to the window's top-left corner (title bar and border).</summary>
    public static (int X, int Y) ClientOffset(long hwnd)
    {
        var h = new IntPtr(hwnd);
        GetWindowRect(h, out var r);
        var p = new POINT();
        ClientToScreen(h, ref p);
        return (p.X - r.L, p.Y - r.T);
    }

    /// <summary>Less of a control than this many pixels across or down counts as hidden: too small to see or click.</summary>
    public const int MinVisiblePixels = 4;

    /// <summary>
    /// null when <paramref name="rect"/> lies inside <paramref name="clip"/>, "partly" when part of it is cut off, "hidden"
    /// when nothing of it shows (or it has no area), or less than <see cref="MinVisiblePixels"/> in either direction.
    /// Rectangles are left, top, right, bottom (right/bottom exclusive).
    /// </summary>
    public static string? ClipState((int L, int T, int R, int B) rect, (int L, int T, int R, int B) clip)
    {
        var i = Intersect(rect, clip);
        // improvement notes 52, live re-test: a button in a 2 x 2-pixel PictureBox was "partly" clipped and app_click
        // clicked it, though nobody could see or click those 4 pixels
        if (i.R - i.L < MinVisiblePixels || i.B - i.T < MinVisiblePixels) return "hidden";
        return i == rect ? null : "partly";
    }

    private static (int L, int T, int R, int B) Intersect((int L, int T, int R, int B) a, (int L, int T, int R, int B) b) =>
        (Math.Max(a.L, b.L), Math.Max(a.T, b.T), Math.Min(a.R, b.R), Math.Min(a.B, b.B));

    private static (int L, int T, int R, int B) ClientScreenRect(IntPtr h)
    {
        GetClientRect(h, out var c);
        var p = new POINT();
        ClientToScreen(h, ref p);
        return (p.X, p.Y, p.X + c.R, p.Y + c.B);
    }

    /// <summary>Friendly control kind from the VB window class ("ThunderCommandButton", "ThunderRT6TextBox", ...).</summary>
    public static string Kind(string cls)
    {
        var c = cls.StartsWith("ThunderRT6", StringComparison.Ordinal) ? cls["ThunderRT6".Length..]
              : cls.StartsWith("Thunder", StringComparison.Ordinal) ? cls["Thunder".Length..] : cls;
        return c switch
        {
            "FormDC" or "Form" => "Form",
            "MDIForm" => "MDIForm",
            "CommandButton" or "Button" => "CommandButton",
            "TextBox" or "Edit" => "TextBox",
            "CheckBox" => "CheckBox",
            "OptionButton" => "OptionButton",
            "ComboBox" => "ComboBox",
            "ListBox" => "ListBox",
            "Frame" => "Frame",
            "PictureBox" or "PictureBoxDC" => "PictureBox",
            "VScrollBar" => "VScrollBar",
            "HScrollBar" => "HScrollBar",
            "Timer" => "Timer",
            "DriveListBox" or "DirListBox" or "FileListBox" => c,
            "#32770" => "Dialog",
            _ => cls,
        };
    }

    /// <summary>
    /// Window text via WM_GETTEXT with a timeout (works across the elevation boundary, spike 10, and
    /// doesn't hang on a busy program). For a text box this is its content.
    /// </summary>
    public static string Text(IntPtr h)
    {
        if (SendMessageTimeout(h, 0x0E /*WM_GETTEXTLENGTH*/, IntPtr.Zero, IntPtr.Zero, 0x2, 500, out var len) == IntPtr.Zero) return "";
        int n = (int)Math.Min(len.ToInt64(), 32768);
        if (n == 0) return "";
        var sb = new StringBuilder(n + 1);
        SendMessageTimeout(h, 0x0D /*WM_GETTEXT*/, new IntPtr(sb.Capacity), sb, 0x2, 500, out _);
        return sb.ToString();
    }

    /// <summary>Window class of a window.</summary>
    public static string ClassOf(long hwnd) => ClassName(new IntPtr(hwnd));

    /// <summary>True for an edit control with ES_READONLY: a VB TextBox with Locked = True, which a user can't type into.</summary>
    public static bool IsReadOnly(long hwnd) => (GetWindowLongW(new IntPtr(hwnd), -16) & 0x0800) != 0;   // GWL_STYLE, ES_READONLY

    /// <summary>
    /// A form's text boxes in reading order (top to bottom, then left to right): what app_set_text's index counts. The
    /// window order put a locked total box first, and an agent typed the price into it (benchmark T18).
    /// </summary>
    public static List<AppControl> TextBoxes(long form) => [.. Controls(form).Where(c => c.Kind == "TextBox").OrderBy(c => c.Top).ThenBy(c => c.Left)];

    private static string ClassName(IntPtr h) { var s = new StringBuilder(128); _ = GetClassName(h, s, s.Capacity); return s.ToString(); }
    private static string Caption(IntPtr h) { var s = new StringBuilder(512); _ = GetWindowText(h, s, s.Capacity); return s.ToString(); }

    /// <summary>
    /// The deepest visible child window of <paramref name="form"/> under a point given relative to the
    /// form's client area (the coordinates app_tree reports), and the point in that window's client
    /// coordinates. A point over a Label or Image returns the window underneath (usually the form or a
    /// frame); VB's runtime routes mouse messages there to the windowless control.
    /// </summary>
    public static (long Hwnd, int X, int Y) HitTest(long form, int x, int y)
    {
        var target = new IntPtr(form);
        var p = new POINT { X = x, Y = y };
        while (true)
        {
            var child = ChildWindowFromPointEx(target, p, 0x1 /*CWP_SKIPINVISIBLE*/ | 0x4 /*CWP_SKIPTRANSPARENT*/);
            if (child == IntPtr.Zero || child == target) break;
            // p is in target's client coordinates; convert to the child's
            ClientToScreen(target, ref p);
            ScreenToClient(child, ref p);
            target = child;
        }
        return (target.ToInt64(), p.X, p.Y);
    }

    /// <summary>A point relative to the form's client area, in <paramref name="hwnd"/>'s client coordinates.</summary>
    public static (int X, int Y) FormToClient(long form, long hwnd, int x, int y)
    {
        var p = new POINT { X = x, Y = y };
        ClientToScreen(new IntPtr(form), ref p);
        ScreenToClient(new IntPtr(hwnd), ref p);
        return (p.X, p.Y);
    }

    /// <summary>
    /// Scroll position (min, max, pos) of a scroll bar control (SB_CTL) or of a window's own vertical or
    /// horizontal scroll bar; null if it has none. Works across processes, so effects of scroll input can be verified.
    /// </summary>
    public static (int Min, int Max, int Pos)? ScrollPos(long hwnd, bool vertical, bool isScrollBarControl)
    {
        var si = new SCROLLINFO { cbSize = (uint)Marshal.SizeOf<SCROLLINFO>(), fMask = 0x1 /*SIF_RANGE*/ | 0x4 /*SIF_POS*/ };
        int bar = isScrollBarControl ? 2 /*SB_CTL*/ : vertical ? 1 /*SB_VERT*/ : 0 /*SB_HORZ*/;
        return GetScrollInfo(new IntPtr(hwnd), bar, ref si) ? (si.nMin, si.nMax, si.nPos) : null;
    }

    [StructLayout(LayoutKind.Sequential)] private struct SCROLLINFO { public uint cbSize, fMask; public int nMin, nMax; public uint nPage; public int nPos, nTrackPos; }
    [DllImport("user32.dll")] private static extern bool GetScrollInfo(IntPtr hwnd, int bar, ref SCROLLINFO si);
    [DllImport("user32.dll")] private static extern int GetWindowLongW(IntPtr hwnd, int index);
    [DllImport("user32.dll")] private static extern IntPtr ChildWindowFromPointEx(IntPtr parent, POINT p, uint flags);
    [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr hwnd, ref POINT p);
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] private struct POINT { public int X, Y; }
    private delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc proc, IntPtr lParam);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr h, uint cmd);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd, out RECT r);
    [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr hwnd, out RECT r);
    [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr hwnd, ref POINT p);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, IntPtr w, StringBuilder l, uint flags, uint timeout, out IntPtr result);
    [DllImport("user32.dll")] private static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
}
