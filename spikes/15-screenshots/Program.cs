// Spike 15: can the NON-elevated side screenshot the elevated VB6 IDE, its designer and code
// windows, dialogs and the running program's forms? Two methods per window:
//   PrintWindow (window renders itself: works when covered; may be blocked or hang)
//   Screen copy  (pixels from the screen: works when hung; only what is visible)
// Usage: ShotSpike [pid] [outDir]

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

Native.SetProcessDPIAware();
int pid = args.Length > 0 ? int.Parse(args[0]) : Process.GetProcessesByName("VB6").First().Id;
string outDir = args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "out");
Directory.CreateDirectory(outDir);
Console.WriteLine($"capturing VB6 pid {pid} -> {outDir}");

int i = 0;
foreach (var h in Native.Targets((uint)pid))
    Console.WriteLine($"[{++i}] " + Capture(h, outDir, i));

static string Capture(IntPtr h, string dir, int idx)
{
    Native.GetWindowRect(h, out var r);
    int w = r.R - r.L, hh = r.B - r.T;
    if (w <= 0 || hh <= 0) return "zero size";
    string name = $"{idx:00}_{Native.Cls(h)}";
    var sb = new StringBuilder($"{Native.Cls(h)} '{Native.Cap(h)}' {w}x{hh} at {r.L},{r.T}");

    var sw = Stopwatch.StartNew();
    using (var bmp = new Bitmap(w, hh, PixelFormat.Format32bppArgb))
    {
        bool ok;
        using (var g = Graphics.FromImage(bmp)) { var dc = g.GetHdc(); ok = Native.PrintWindow(h, dc, 2); g.ReleaseHdc(dc); }
        sw.Stop();
        bmp.Save(Path.Combine(dir, name + "_print.png"), ImageFormat.Png);
        sb.Append($"\n    PrintWindow: ok={ok} {sw.ElapsedMilliseconds}ms, {Blankness(bmp)}");
    }
    sw.Restart();
    using (var bmp = new Bitmap(w, hh, PixelFormat.Format32bppArgb))
    {
        using (var g = Graphics.FromImage(bmp)) g.CopyFromScreen(r.L, r.T, 0, 0, new Size(w, hh));
        sw.Stop();
        bmp.Save(Path.Combine(dir, name + "_screen.png"), ImageFormat.Png);
        sb.Append($"\n    Screen:      {sw.ElapsedMilliseconds}ms, {Blankness(bmp)}");
    }
    return sb.ToString();
}

// Share of the most common colour in a sample grid: ~100% means a blank/failed capture
static string Blankness(Bitmap b)
{
    var counts = new Dictionary<int, int>(); int n = 0;
    for (int y = 0; y < b.Height; y += Math.Max(1, b.Height / 40))
        for (int x = 0; x < b.Width; x += Math.Max(1, b.Width / 40))
        { int c = b.GetPixel(x, y).ToArgb(); counts[c] = counts.GetValueOrDefault(c) + 1; n++; }
    return $"{100.0 * counts.Values.Max() / n:0}% one colour, {counts.Count} colours";
}

static class Native
{
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc p, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    public static string Cls(IntPtr h) { var s = new StringBuilder(64); GetClassName(h, s, 64); return s.ToString(); }
    public static string Cap(IntPtr h) { var s = new StringBuilder(256); GetWindowText(h, s, 256); return s.ToString(); }

    // Visible top-level windows of the process + the IDE's designer/code/project/properties panes
    public static List<IntPtr> Targets(uint pid)
    {
        var l = new List<IntPtr>();
        EnumWindows((h, _) =>
        {
            GetWindowThreadProcessId(h, out var p);
            if (p != pid || !IsWindowVisible(h)) return true;
            string c = Cls(h);
            if (c is "CompatWindowDesktopReplacement" or "IDEOwner" or "Static") return true;
            l.Add(h);
            if (c == "wndclass_desked_gsk")
                EnumChildWindows(h, (k, _) =>
                {
                    if (IsWindowVisible(k) && Cls(k) is "DesignerWindow" or "VbaWindow" or "PROJECT" or "wndclass_pbrs") l.Add(k);
                    return true;
                }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
        return l;
    }
}
