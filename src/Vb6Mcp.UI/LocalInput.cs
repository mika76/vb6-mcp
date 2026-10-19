using System.Runtime.InteropServices;
using System.Text;

namespace Vb6Mcp.UI;

/// <summary>
/// Input from the server itself, for compiled programs started with exe_run. They run at the server's
/// (medium) integrity, so its messages and SendInput reach them, unlike the elevated IDE (spike 10),
/// where the add-in does the same things in-process. Mirrors the add-in's input (Input.bas): posted
/// messages by default, real input (SendInput) on request. Only for windows of processes the server
/// started (checked by the caller, hard rule 8).
/// </summary>
public static class LocalInput
{
    // Windows-1252 (characters of the key events) needs the code-pages provider; the server registers
    // it in Core (Vb1252), but this class must not depend on that having happened first.
    static LocalInput() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    public static void Click(long hwnd) => PostMessage(H(hwnd), BM_CLICK, 0, 0);

    /// <summary>WM_SETTEXT with a timeout: fires the control's Change event in the program.</summary>
    public static void SetText(long hwnd, string text) => SendMessageTimeout(H(hwnd), WM_SETTEXT, IntPtr.Zero, text, SMTO_ABORTIFHUNG, 2000, out _);

    /// <summary>Posted mouse input in hwnd's client coordinates (same sequences as the add-in's MOUSE verb).</summary>
    public static void Mouse(long hwnd, string action, int x, int y, int x2 = 0, int y2 = 0)
    {
        var h = H(hwnd);
        switch (action)
        {
            case "click": Post(h, WM_MOUSEMOVE, 0, x, y); Post(h, WM_LBUTTONDOWN, 1, x, y); Post(h, WM_LBUTTONUP, 0, x, y); break;
            case "dblclick":
                Post(h, WM_MOUSEMOVE, 0, x, y); Post(h, WM_LBUTTONDOWN, 1, x, y); Post(h, WM_LBUTTONUP, 0, x, y);
                Post(h, WM_LBUTTONDBLCLK, 1, x, y); Post(h, WM_LBUTTONUP, 0, x, y); break;
            case "rclick": Post(h, WM_MOUSEMOVE, 0, x, y); Post(h, WM_RBUTTONDOWN, 2, x, y); Post(h, WM_RBUTTONUP, 0, x, y); break;
            case "drag":
                Post(h, WM_MOUSEMOVE, 0, x, y); Post(h, WM_LBUTTONDOWN, 1, x, y);
                for (int i = 1; i <= 10; i++) Post(h, WM_MOUSEMOVE, 1, x + (x2 - x) * i / 10, y + (y2 - y) * i / 10);
                Post(h, WM_LBUTTONUP, 0, x2, y2); break;
            case "wheel":
                var p = new POINT { X = x, Y = y }; ClientToScreen(h, ref p);
                for (int i = 0; i < Math.Abs(x2); i++) PostMessage(h, WM_MOUSEWHEEL, (Math.Sign(x2) * 120) << 16, MakeLParam(p.X, p.Y));
                break;
            default: throw new ArgumentException("unknown mouse action " + action);
        }
    }

    /// <summary>Scroll bar commands; a scroll bar control notifies its parent (as the add-in's SCROLL).</summary>
    public static void Scroll(long hwnd, bool vertical, int code, int count, bool isScrollBar)
    {
        var h = H(hwnd);
        var target = isScrollBar ? GetParent(h) : h;
        nint lp = isScrollBar ? h : 0;
        uint msg = vertical ? WM_VSCROLL : WM_HSCROLL;
        for (int i = 0; i < count; i++) PostMessage(target, msg, code, lp);
        PostMessage(target, msg, SB_ENDSCROLL, lp);
    }

    /// <summary>Items and selection of a list or combo box. Standard list messages are marshalled across processes.</summary>
    public static (string Kind, IReadOnlyList<string> Items, int Selected) ListState(long hwnd, bool isCombo)
    {
        var h = H(hwnd);
        uint getCount = isCombo ? CB_GETCOUNT : LB_GETCOUNT, getSel = isCombo ? CB_GETCURSEL : LB_GETCURSEL;
        uint getLen = isCombo ? CB_GETLBTEXTLEN : LB_GETTEXTLEN, getText = isCombo ? CB_GETLBTEXT : LB_GETTEXT;
        int n = (int)Send(h, getCount, 0), sel = (int)Send(h, getSel, 0);
        var items = new List<string>();
        for (int i = 0; i < n; i++)
        {
            int len = Math.Max(0, (int)Send(h, getLen, i));
            var sb = new StringBuilder(len + 1);
            SendMessageTimeout(h, getText, i, sb, SMTO_ABORTIFHUNG, 1000, out _);
            items.Add(sb.ToString());
        }
        return (isCombo ? "combo" : "list", items, sel);
    }

    /// <summary>Selects an item and posts *_SELCHANGE to the parent, where VB raises Click (as the add-in's SELECT).</summary>
    public static void Select(long hwnd, int index, bool isCombo)
    {
        var h = H(hwnd);
        Send(h, isCombo ? CB_SETCURSEL : LB_SETCURSEL, index);
        PostMessage(GetParent(h), WM_COMMAND, (GetDlgCtrlID(h) & 0xFFFF) | (1 << 16), h);
    }

    /// <summary>
    /// Posted key events (the add-in's KEYS format). The add-in can defer between messages inside the
    /// program's thread; from outside, the server attaches to the program's input thread to set the
    /// modifier state and pauses after each key down, so the WM_CHAR that TranslateMessage generates
    /// arrives before the next posted message (Phase 3 finding). Runs synchronously on one thread
    /// (thread attachment is per thread). hwnd 0: the program's focused control.
    /// </summary>
    public static long Keys(long hwnd, string events, int programThread)
    {
        var h = hwnd != 0 ? H(hwnd) : Focused(programThread);
        if (h == IntPtr.Zero) throw new InvalidOperationException("no control of the program has the focus; pass the control (from app_tree)");
        uint me = GetCurrentThreadId();
        bool attached = AttachThreadInput(me, (uint)programThread, true);
        try
        {
            int mods = 0;
            foreach (var tok in events.Split(';'))
            {
                char kind = tok[0]; int v = int.Parse(tok.AsSpan(1));
                switch (kind)
                {
                    case 'S':
                        mods = v;
                        var ks = new byte[256]; GetKeyboardState(ks);
                        ks[0x10] = ks[0xA0] = (byte)((v & 1) != 0 ? 0x80 : 0);
                        ks[0x11] = ks[0xA2] = (byte)((v & 2) != 0 ? 0x80 : 0);
                        ks[0x12] = ks[0xA4] = (byte)((v & 4) != 0 ? 0x80 : 0);
                        SetKeyboardState(ks);
                        break;
                    case 'D':
                    case 'U':
                        nint lp = 1 | ((nint)MapVirtualKey((uint)v, 0) << 16);
                        if (v is >= 0x21 and <= 0x28 or 0x2D or 0x2E) lp |= 0x1000000;
                        if ((mods & 4) != 0) lp |= 0x20000000;
                        if (kind == 'U') lp |= unchecked((nint)0xC0000000);
                        uint msg = (mods & 4) != 0 ? (kind == 'D' ? WM_SYSKEYDOWN : WM_SYSKEYUP) : (kind == 'D' ? WM_KEYDOWN : WM_KEYUP);
                        PostMessage(h, msg, v, lp);
                        if (kind == 'D') Thread.Sleep(60);         // let TranslateMessage queue its WM_CHAR first
                        break;
                    case 'C': PostMessage(h, WM_CHAR, v, 1); break;
                }
            }
            Thread.Sleep(60);
            if (mods != 0)
            {
                var ks = new byte[256]; GetKeyboardState(ks);
                ks[0x10] = ks[0xA0] = ks[0x11] = ks[0xA2] = ks[0x12] = ks[0xA4] = 0;
                SetKeyboardState(ks);
            }
        }
        finally { if (attached) AttachThreadInput(me, (uint)programThread, false); }
        return h.ToInt64();
    }

    /// <summary>The focused window of a thread (GetGUIThreadInfo works across processes).</summary>
    public static IntPtr Focused(int thread)
    {
        var gti = new GUITHREADINFO { cbSize = Marshal.SizeOf<GUITHREADINFO>() };
        return GetGUIThreadInfo((uint)thread, ref gti) ? gti.hwndFocus : IntPtr.Zero;
    }

    /// <summary>
    /// The virtual key and Shift that type <paramref name="c"/> on the current keyboard layout (VkKeyScan), for real input;
    /// null when it needs Ctrl or Alt (AltGr characters) or the layout has no key for it: then it goes as a Unicode
    /// character. Caps Lock inverts Shift for letters, as it would for the user's own typing.
    /// </summary>
    public static (int Vk, bool Shift)? KeyFor(char c)
    {
        if (c < ' ') return null;
        short r = VkKeyScanW(c);
        if (r == -1 || (r & 0x0600) != 0) return null;     // no key, or needs Ctrl/Alt
        int vk = r & 0xFF; bool shift = (r & 0x0100) != 0;
        if (char.IsLetter(c) && (GetKeyState(0x14 /*VK_CAPITAL*/) & 1) != 0) shift = !shift;
        return (vk, shift);
    }

    /// <summary>The top-level window a window belongs to (GetAncestor GA_ROOT), 0 for 0.</summary>
    public static long Root(long hwnd) => hwnd == 0 ? 0 : GetAncestor(H(hwnd), 2).ToInt64();

    /// <summary>Thread that owns a window.</summary>
    public static int WindowThread(long hwnd) => (int)GetWindowThreadProcessId(H(hwnd), out _);

    /// <summary>
    /// Real input: brings hwnd's form to the front (attached to the foreground thread's input, spike 17),
    /// runs <paramref name="send"/> (SendInput), then restores cursor and foreground window. Returns
    /// "sent" or why not. Synchronous on one thread.
    /// </summary>
    public static string Real(long hwnd, Func<IntPtr, string?> send)
    {
        var h = H(hwnd);
        var top = GetAncestor(h, 2);
        GetCursorPos(out var cursor);
        var previous = GetForegroundWindow();
        try
        {
            uint fgThread = GetWindowThreadProcessId(previous, out _), me = GetCurrentThreadId();
            bool attached = fgThread != 0 && fgThread != me && AttachThreadInput(me, fgThread, true);
            SetForegroundWindow(top);
            if (attached) AttachThreadInput(me, fgThread, false);
            Thread.Sleep(100);
            // SetForegroundWindow alone is often refused (the user is active in another program);
            // SwitchToThisWindow (what Alt+Tab uses) gets there (Phase 3 live test, spike 19)
            for (int i = 0; i < 10 && GetForegroundWindow() != top; i++) { SwitchToThisWindow(top, true); Thread.Sleep(100); }
            if (GetForegroundWindow() != top) { RestoreForeground(previous, cursor); return "couldn't bring the form to the front"; }
            var why = send(h);
            Thread.Sleep(400);                                   // let the program process the input
            if (why != null) return why;
            return RestoreForeground(previous, cursor) ? "sent" : "sent (the previous window couldn't be brought back to the front)";
        }
        catch { RestoreForeground(previous, cursor); throw; }
    }

    /// <summary>
    /// Puts the cursor and the previous foreground window back and checks that it worked. Windows may
    /// refuse while the user is typing in another window (Phase 3 live test); SwitchToThisWindow is the
    /// fallback (spike 19).
    /// </summary>
    private static bool RestoreForeground(IntPtr previous, POINT cursor)
    {
        SetCursorPos(cursor.X, cursor.Y);
        if (previous == IntPtr.Zero) return true;
        uint fgThread = GetWindowThreadProcessId(GetForegroundWindow(), out _), me = GetCurrentThreadId();
        bool attached = fgThread != 0 && fgThread != me && AttachThreadInput(me, fgThread, true);
        SetForegroundWindow(previous);
        if (attached) AttachThreadInput(me, fgThread, false);
        Thread.Sleep(50);
        if (GetForegroundWindow() == previous) return true;
        SwitchToThisWindow(previous, true);
        Thread.Sleep(150);
        return GetForegroundWindow() == previous;
    }

    /// <summary>SendInput mouse sequence at hwnd's client point (same actions as the add-in's MOUSEREAL).</summary>
    public static string? RealMouse(IntPtr h, string action, int x, int y, int x2, int y2)
    {
        var p = new POINT { X = x, Y = y }; ClientToScreen(h, ref p);
        var p2 = new POINT { X = x2, Y = y2 }; ClientToScreen(h, ref p2);
        var hit = WindowFromPoint(p);
        GetWindowThreadProcessId(hit, out var hitPid); GetWindowThreadProcessId(h, out var pid);
        if (hitPid != pid) return "the point is covered by another program's window";
        var ins = new List<INPUT> { Move(p.X, p.Y) };
        switch (action)
        {
            case "click": ins.Add(Btn(0x2)); ins.Add(Btn(0x4)); break;
            case "dblclick": ins.Add(Btn(0x2)); ins.Add(Btn(0x4)); ins.Add(Btn(0x2)); ins.Add(Btn(0x4)); break;
            case "rclick": ins.Add(Btn(0x8)); ins.Add(Btn(0x10)); break;
            case "drag":
                // One batch, like the add-in. A paced version (moves 15 ms apart) made no difference in
                // the live test (spike 20), so the simple form stays.
                ins.Add(Btn(0x2));
                for (int i = 1; i <= 10; i++) ins.Add(Move(p.X + (p2.X - p.X) * i / 10, p.Y + (p2.Y - p.Y) * i / 10));
                ins.Add(Btn(0x4)); break;
        }
        return SendAll(ins);
    }

    private static string? SendAll(List<INPUT> ins)
    {
        var arr = ins.ToArray();
        uint sent = SendInput((uint)arr.Length, arr, Marshal.SizeOf<INPUT>());
        return sent == arr.Length ? null : $"SendInput sent {sent} of {arr.Length} events";
    }

    /// <summary>SendInput keyboard events (the add-in's KEYS format), focusing h first.</summary>
    public static string? RealKeys(IntPtr h, string events)
    {
        uint thread = GetWindowThreadProcessId(h, out _), me = GetCurrentThreadId();
        if (AttachThreadInput(me, thread, true)) { SetFocus(h); AttachThreadInput(me, thread, false); }
        var ins = new List<INPUT>();
        int mods = 0;
        foreach (var tok in events.Split(';'))
        {
            char kind = tok[0]; int v = int.Parse(tok.AsSpan(1));
            switch (kind)
            {
                case 'S':
                    for (int b = 0; b < 3; b++)
                    {
                        ushort vk = (ushort)(0x10 + b);
                        bool now = (v & (1 << b)) != 0, was = (mods & (1 << b)) != 0;
                        if (now && !was) ins.Add(Key(vk, 0, false));
                        if (!now && was) ins.Add(Key(vk, 0, true));
                    }
                    mods = v; break;
                case 'D': ins.Add(Key((ushort)v, 0, false)); break;
                case 'U': ins.Add(Key((ushort)v, 0, true)); break;
                case 'C':
                    char c = Encoding.GetEncoding(1252).GetString([(byte)v])[0];
                    ins.Add(Key(0, c, false)); ins.Add(Key(0, c, true)); break;
            }
        }
        for (int b = 0; b < 3; b++) if ((mods & (1 << b)) != 0) ins.Add(Key((ushort)(0x10 + b), 0, true));   // never leave a modifier down
        var arr = ins.ToArray();
        uint sent = SendInput((uint)arr.Length, arr, Marshal.SizeOf<INPUT>());
        return sent == arr.Length ? null : $"SendInput sent {sent} of {arr.Length} events";
    }

    // ---------- helpers ----------

    private static IntPtr H(long hwnd) => new(hwnd);
    private static void Post(IntPtr h, uint msg, nint w, int x, int y) => PostMessage(h, msg, w, MakeLParam(x, y));
    private static nint MakeLParam(int lo, int hi) => (nint)(((hi & 0xFFFF) << 16) | (lo & 0xFFFF));
    private static nint Send(IntPtr h, uint msg, nint w) { SendMessageTimeout(h, msg, w, IntPtr.Zero, SMTO_ABORTIFHUNG, 1000, out var r); return r; }

    private static INPUT Move(int sx, int sy)
    {
        int vx = GetSystemMetrics(76), vy = GetSystemMetrics(77), vw = GetSystemMetrics(78), vh = GetSystemMetrics(79);
        return new INPUT { type = 0, u = new InputUnion { mi = new MOUSEINPUT { dx = (sx - vx) * 65535 / (vw - 1), dy = (sy - vy) * 65535 / (vh - 1), dwFlags = 0x1 | 0x8000 | 0x4000 } } };
    }
    private static INPUT Btn(uint flags) => new() { type = 0, u = new InputUnion { mi = new MOUSEINPUT { dwFlags = flags } } };
    private static INPUT Key(ushort vk, int uni, bool up)
    {
        uint flags = uni != 0 ? 0x4u : (vk is >= 0x21 and <= 0x28 or 0x2D or 0x2E ? 0x1u : 0);
        if (up) flags |= 0x2;
        return new INPUT { type = 1, u = new InputUnion { ki = new KEYBDINPUT { wVk = uni != 0 ? (ushort)0 : vk, wScan = uni != 0 ? (ushort)uni : (ushort)MapVirtualKey(vk, 0), dwFlags = flags } } };
    }

    private const uint WM_SETTEXT = 0x0C, WM_KEYDOWN = 0x100, WM_KEYUP = 0x101, WM_CHAR = 0x102, WM_SYSKEYDOWN = 0x104, WM_SYSKEYUP = 0x105,
        WM_COMMAND = 0x111, WM_HSCROLL = 0x114, WM_VSCROLL = 0x115, WM_MOUSEMOVE = 0x200, WM_LBUTTONDOWN = 0x201, WM_LBUTTONUP = 0x202,
        WM_LBUTTONDBLCLK = 0x203, WM_RBUTTONDOWN = 0x204, WM_RBUTTONUP = 0x205, WM_MOUSEWHEEL = 0x20A, BM_CLICK = 0xF5,
        LB_SETCURSEL = 0x186, LB_GETCURSEL = 0x188, LB_GETTEXT = 0x189, LB_GETTEXTLEN = 0x18A, LB_GETCOUNT = 0x18B,
        CB_GETCOUNT = 0x146, CB_GETCURSEL = 0x147, CB_GETLBTEXT = 0x148, CB_GETLBTEXTLEN = 0x149, CB_SETCURSEL = 0x14E,
        SMTO_ABORTIFHUNG = 0x2;
    private const int SB_ENDSCROLL = 8;

    [StructLayout(LayoutKind.Sequential)] private struct POINT { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)]
    private struct GUITHREADINFO { public int cbSize, flags; public IntPtr hwndActive, hwndFocus, hwndCapture, hwndMenuOwner, hwndMoveSize, hwndCaret; public RECT rcCaret; }
    [StructLayout(LayoutKind.Sequential)] private struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr extra; }
    [StructLayout(LayoutKind.Sequential)] private struct KEYBDINPUT { public ushort wVk, wScan; public uint dwFlags, time; public IntPtr extra; }
    [StructLayout(LayoutKind.Explicit)] private struct InputUnion { [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; }
    [StructLayout(LayoutKind.Sequential)] private struct INPUT { public uint type; public InputUnion u; }

    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr h, uint msg, nint w, nint l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern short VkKeyScanW(char c);
    [DllImport("user32.dll")] private static extern short GetKeyState(int vk);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, IntPtr w, string l, uint flags, uint timeout, out IntPtr r);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, nint w, StringBuilder l, uint flags, uint timeout, out IntPtr r);
    [DllImport("user32.dll")] private static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, nint w, IntPtr l, uint flags, uint timeout, out nint r);
    [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr h, uint flags);
    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr h);
    [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr h, ref POINT p);
    [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(POINT p);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread, ref GUITHREADINFO gti);
    [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint a, uint b, bool attach);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern bool GetKeyboardState(byte[] ks);
    [DllImport("user32.dll")] private static extern bool SetKeyboardState(byte[] ks);
    [DllImport("user32.dll")] private static extern uint MapVirtualKey(uint code, uint mapType);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] private static extern void SwitchToThisWindow(IntPtr h, bool altTab);
    [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr h);
    [DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT p);
    [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] private static extern int GetSystemMetrics(int i);
    [DllImport("user32.dll")] private static extern uint SendInput(uint n, INPUT[] inputs, int size);
}
