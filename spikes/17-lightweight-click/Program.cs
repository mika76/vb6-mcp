// Spike 17: which mouse messages make VB6 raise Click on a windowless control (Label)?
// Run ELEVATED (like VB6) while tests/fixtures/Ui runs in the IDE. lblClick_Click writes its click
// count into the form caption, so each variant's effect is read with GetWindowText.
// Usage: Spike17 [x y] [--sendinput]   (x, y: point on lblClick in the form's client area; default 60 228)
using System.Runtime.InteropServices;
using System.Text;

int x = args.Length >= 2 && int.TryParse(args[0], out var ax) ? ax : 60;
int y = args.Length >= 2 && int.TryParse(args[1], out var ay) ? ay : 228;
bool allowSendInput = args.Contains("--sendinput");

IntPtr form = IntPtr.Zero;
EnumWindows((h, _) => { if (ClassName(h) == "ThunderFormDC" && Caption(h).StartsWith("UI fixture")) { form = h; return false; } return true; }, IntPtr.Zero);
if (form == IntPtr.Zero) { Console.WriteLine("UI fixture form not found: run tests/fixtures/Ui in the IDE first"); return 1; }
Console.WriteLine($"form {form:X} '{Caption(form)}', point ({x},{y})");

const uint WM_MOUSEMOVE = 0x200, WM_LBUTTONDOWN = 0x201, WM_LBUTTONUP = 0x202, WM_SETCURSOR = 0x20, WM_NCHITTEST = 0x84;
IntPtr lp = (IntPtr)((y << 16) | (x & 0xFFFF));
var screen = new POINT { X = x, Y = y }; ClientToScreen(form, ref screen);
IntPtr lpScreen = (IntPtr)((screen.Y << 16) | (screen.X & 0xFFFF));
IntPtr setCursor(uint mouseMsg) => (IntPtr)(((int)mouseMsg << 16) | 1 /*HTCLIENT*/);

void Post3() { PostMessage(form, WM_MOUSEMOVE, 0, lp); PostMessage(form, WM_LBUTTONDOWN, 1, lp); PostMessage(form, WM_LBUTTONUP, 0, lp); }

var variants = new List<(string Name, Action Run)>
{
    ("A posted move/down/up (what MOUSE does now)", Post3),
    ("B posted WM_SETCURSOR before move and down", () =>
    {
        PostMessage(form, WM_SETCURSOR, form, setCursor(WM_MOUSEMOVE)); PostMessage(form, WM_MOUSEMOVE, 0, lp);
        PostMessage(form, WM_SETCURSOR, form, setCursor(WM_LBUTTONDOWN)); PostMessage(form, WM_LBUTTONDOWN, 1, lp);
        PostMessage(form, WM_LBUTTONUP, 0, lp);
    }),
    ("C sent WM_NCHITTEST + WM_SETCURSOR, then posted move/down/up", () =>
    {
        SendMessageTimeout(form, WM_NCHITTEST, 0, lpScreen, 2, 1000, out _);
        SendMessageTimeout(form, WM_SETCURSOR, form, setCursor(WM_MOUSEMOVE), 2, 1000, out _);
        Post3();
    }),
    ("D sent move/down/up (SendMessage)", () =>
    {
        SendMessageTimeout(form, WM_MOUSEMOVE, 0, lp, 2, 1000, out _);
        SendMessageTimeout(form, WM_LBUTTONDOWN, 1, lp, 2, 1000, out _);
        SendMessageTimeout(form, WM_LBUTTONUP, 0, lp, 2, 1000, out _);
    }),
    ("E real cursor moved onto the label (SetCursorPos), then posted move/down/up, cursor restored", () =>
    {
        GetCursorPos(out var old); SetCursorPos(screen.X, screen.Y); Thread.Sleep(50);
        Post3(); Thread.Sleep(300); SetCursorPos(old.X, old.Y);
    }),
};
// Capture hypothesis: VB raises a Label Click only if the form still has the mouse capture at button-up,
// and only the foreground window may capture the mouse. So: bring the form to the front first.
bool Foreground()
{
    // a background process may not take the foreground; a synthetic Alt press lifts that lock
    keybd_event(0x12, 0, 0, 0); keybd_event(0x12, 0, 2, 0);
    SetForegroundWindow(form); Thread.Sleep(200);
    bool ok = GetForegroundWindow() == form;
    Console.WriteLine($"       foreground: {(ok ? "form" : "NOT the form")} ({GetForegroundWindow():X})");
    return ok;
}
variants.Add(("G form brought to the front, then posted move/down/up", () => { Foreground(); Post3(); }));
variants.Add(("H form in front, posted down, capture checked, posted up", () =>
{
    Foreground(); PostMessage(form, WM_MOUSEMOVE, 0, lp); PostMessage(form, WM_LBUTTONDOWN, 1, lp); Thread.Sleep(300);
    var gti = new GUITHREADINFO { cbSize = Marshal.SizeOf<GUITHREADINFO>() };
    GetGUIThreadInfo(GetWindowThreadProcessId(form, out _), ref gti);
    Console.WriteLine($"       after down: capture={gti.hwndCapture:X} (form {form:X})");
    PostMessage(form, WM_LBUTTONUP, 0, lp);
}));
// Key-state hypothesis: posted messages do not update the thread's key state, and VB may check
// GetKeyState(VK_LBUTTON). Attach to the form's thread so both share key state, set the button state by hand.
void SetLButton(bool down)
{
    var ks = new byte[256]; GetKeyboardState(ks);
    ks[0x01] = (byte)(down ? 0x80 : 0); SetKeyboardState(ks);
}
variants.Add(("I attached to the form thread, VK_LBUTTON set down/up around posted down/up (form in background)", () =>
{
    uint vbThread = GetWindowThreadProcessId(form, out _), me = GetCurrentThreadId();
    bool attached = AttachThreadInput(me, vbThread, true);
    Console.WriteLine($"       attached: {attached}");
    SetLButton(true); PostMessage(form, WM_MOUSEMOVE, 1, lp); PostMessage(form, WM_LBUTTONDOWN, 1, lp); Thread.Sleep(300);
    SetLButton(false); PostMessage(form, WM_LBUTTONUP, 0, lp); Thread.Sleep(300);
    AttachThreadInput(me, vbThread, false);
}));
if (allowSendInput)
    variants.Add(("F SendInput click (real input; brings the form to the front)", () =>
    {
        GetCursorPos(out var old); Foreground(); SetCursorPos(screen.X, screen.Y); Thread.Sleep(50);
        var inputs = new INPUT[] { new() { type = 0, mi = new() { dwFlags = 0x2 } }, new() { type = 0, mi = new() { dwFlags = 0x4 } } };
        SendInput(2, inputs, Marshal.SizeOf<INPUT>()); Thread.Sleep(300); SetCursorPos(old.X, old.Y);
    }));

string? only = args.FirstOrDefault(a => a.StartsWith("--only="))?["--only=".Length..];   // e.g. --only=I or --only=AI
foreach (var (name, run) in variants.Where(v => only is null || only.Contains(v.Name[0])))
{
    string before = Caption(form);
    run();
    Thread.Sleep(700);
    string after = Caption(form);
    Console.WriteLine($"{(after != before ? "CLICK " : "none  ")} {name}   [{before}] -> [{after}]");
}
return 0;

static string ClassName(IntPtr h) { var s = new StringBuilder(128); GetClassName(h, s, s.Capacity); return s.ToString(); }
static string Caption(IntPtr h) { var s = new StringBuilder(256); GetWindowText(h, s, s.Capacity); return s.ToString(); }

delegate bool EnumProc(IntPtr h, IntPtr l);
[StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
[StructLayout(LayoutKind.Sequential)] struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr extra; }
[StructLayout(LayoutKind.Sequential)] struct INPUT { public uint type; public MOUSEINPUT mi; }   // x86: MOUSEINPUT is the largest member of the union
partial class Program
{
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern IntPtr SendMessageTimeout(IntPtr h, uint m, IntPtr w, IntPtr l, uint f, uint t, out IntPtr r);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h, ref POINT p);
    [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT p);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern uint SendInput(uint n, INPUT[] i, int size);
}
[StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }
[StructLayout(LayoutKind.Sequential)] struct GUITHREADINFO { public int cbSize, flags; public IntPtr hwndActive, hwndFocus, hwndCapture, hwndMenuOwner, hwndMoveSize, hwndCaret; public RECT rcCaret; }
partial class Program
{
    [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, int extra);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool GetGUIThreadInfo(uint thread, ref GUITHREADINFO info);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
}
partial class Program
{
    [DllImport("user32.dll")] static extern bool GetKeyboardState(byte[] ks);
    [DllImport("user32.dll")] static extern bool SetKeyboardState(byte[] ks);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool attach);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
}
