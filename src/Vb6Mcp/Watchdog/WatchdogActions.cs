using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.UI;

namespace Vb6Mcp.Watchdog;

/// <summary>
/// The watchdog's Win32 actions. They run outside any VB6 thread, so they work while an IDE is busy
/// or hung (spike 12). As an elevated child of VB6 the watchdog may send input to VB6 (spike 10).
/// </summary>
public sealed class WatchdogActions(Action<string> log, string shotsDir) : IWatchdogActions
{
    public bool Elevated { get; } = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
    public int ParentPid { get; } = GetParentPid();

    public bool IsVb6(int pid) => Discovery.IsVb6Process(pid);

    public int WindowPid(long hwnd)
    {
        GetWindowThreadProcessId(new IntPtr(hwnd), out var pid);
        return (int)pid;
    }

    /// <summary>
    /// Injected Ctrl+Break drops a busy loop into break mode (spike 12). Keyboard input goes to the
    /// foreground window, and Ctrl+Break in a console window kills the program running there, so
    /// the input is only sent while VB6 (or Windows' "ghost" stand-in for its unresponsive window)
    /// is in front. The previous foreground window is restored afterwards.
    /// </summary>
    public async Task<string> BreakAsync(int pid, CancellationToken ct)
    {
        long main = IdeWindows.MainWindow(pid);
        if (main == 0) throw new WatchdogRefusedException("the IDE main window wasn't found");
        var mode = IdeWindows.ParseMode(IdeWindows.Read(pid, 200).Title);
        if (mode != IdeMode.Run) throw new WatchdogRefusedException($"the IDE is in {mode.ToString().ToLowerInvariant()} mode, not run mode");

        var previous = GetForegroundWindow();
        string front = BringToFront(new IntPtr(main), pid);
        if (front != "vb6" && front != "ghost")
        {
            Restore(previous, pid);
            throw new WatchdogRefusedException($"VB6 couldn't be brought to the front ({front}); Ctrl+Break would go to another program");
        }
        keybd_event(VK_CONTROL, 0, 0, 0);
        keybd_event(VK_CANCEL, 0x46, KEYEVENTF_EXTENDEDKEY, 0);
        keybd_event(VK_CANCEL, 0x46, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
        log($"Ctrl+Break injected for {pid} (front: {front})");

        var deadline = DateTime.UtcNow.AddSeconds(3);
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(150, ct);
            if (IdeWindows.ParseMode(IdeWindows.Read(pid, 200).Title) == IdeMode.Break)
            {
                Restore(previous, pid);
                return $"break{Wire.US}{front}";
            }
        }
        Restore(previous, pid);
        throw new WatchdogRefusedException($"Ctrl+Break was sent (front: {front}) but the IDE didn't enter break mode within 3 s");
    }

    /// <summary>"vb6", "ghost", or what is in front instead.</summary>
    private static string Classify(IntPtr fg, int pid)
    {
        GetWindowThreadProcessId(fg, out var fgPid);
        if (fgPid == (uint)pid) return "vb6";
        var cls = new StringBuilder(64); GetClassName(fg, cls, 64);
        return cls.ToString() == "Ghost" ? "ghost" : $"'{cls}' of process {fgPid}";
    }

    /// <summary>
    /// Brings the IDE to the front unless VB6 already is (typically the debugged program's own form).
    /// Never attaches to a thread of the VB6 process: that thread is the one that is busy, and
    /// SetForegroundWindow while attached to it blocked the watchdog (Phase 2b live test). The
    /// foreground change runs on its own thread with a timeout, in case it still waits on VB6.
    /// </summary>
    private static string BringToFront(IntPtr main, int pid)
    {
        string now = Classify(GetForegroundWindow(), pid);
        if (now is "vb6" or "ghost") return now;
        // A window that doesn't respond doesn't easily take the foreground (Phase 2b live test: plain
        // SetForegroundWindow left Claude in front). Try stronger ways one after the other.
        var tried = new List<string>();
        SetForegroundSafely(main, pid);
        if ((now = Classify(GetForegroundWindow(), pid)) is "vb6" or "ghost") return now;
        tried.Add("SetForegroundWindow: " + now);
        // what Alt+Tab uses; it also switches to unresponsive windows (Windows shows a ghost)
        SwitchToThisWindow(main, true);
        Thread.Sleep(300);
        if ((now = Classify(GetForegroundWindow(), pid)) is "vb6" or "ghost") return now;
        tried.Add("SwitchToThisWindow: " + now);
        // an Alt press lifts the foreground lock (spike 17); Alt alone does nothing harmful in the front program
        keybd_event(VK_MENU, 0, 0, 0); keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        SetForegroundSafely(main, pid);
        if ((now = Classify(GetForegroundWindow(), pid)) is "vb6" or "ghost") return now;
        tried.Add("Alt + SetForegroundWindow: " + now);
        return string.Join("; ", tried);
    }

    /// <summary>Gives the foreground back, unless it was VB6 already (nothing to restore).</summary>
    private static void Restore(IntPtr window, int pid)
    {
        if (window == IntPtr.Zero || Classify(window, pid) is "vb6" or "ghost") return;
        SetForegroundSafely(window, pid);
    }

    /// <summary>
    /// SetForegroundWindow, attached to the current foreground thread's input (a background process
    /// may not take the foreground otherwise, spike 17) unless that thread belongs to VB6. On a
    /// separate thread with a 1 s timeout; a stuck thread is abandoned rather than blocking the request.
    /// </summary>
    private static void SetForegroundSafely(IntPtr window, int vb6Pid)
    {
        var t = new Thread(() =>
        {
            uint fgThread = GetWindowThreadProcessId(GetForegroundWindow(), out var fgPid), me = GetCurrentThreadId();
            bool attached = fgPid != (uint)vb6Pid && fgThread != 0 && fgThread != me && AttachThreadInput(me, fgThread, true);
            SetForegroundWindow(window);
            if (attached) AttachThreadInput(me, fgThread, false);
        }) { IsBackground = true };
        t.Start();
        t.Join(1000);
        Thread.Sleep(150);
    }
    /// <summary>Clicks a dialog button of a VB6 process (works while the add-in can't answer).</summary>
    public string Press(long buttonHwnd)
    {
        var h = new IntPtr(buttonHwnd);
        var cls = new StringBuilder(64); GetClassName(h, cls, 64);
        if (!cls.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase))
            throw new WatchdogRefusedException($"window {buttonHwnd:X} is a {cls}, not a dialog button");
        PostMessage(h, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
        return "pressed";
    }

    /// <summary>Ends a VB6 process. Last resort; the server only asks with force (hard rule 12).</summary>
    public string Kill(int pid)
    {
        using var p = Process.GetProcessById(pid);
        p.Kill();
        bool gone = p.WaitForExit(5000);
        log($"killed {pid}: {(gone ? "exited" : "still running after 5 s")}");
        return gone ? "killed" : "kill-sent";
    }

    /// <summary>
    /// PNG of a window. PrintWindow captures covered windows (spike 15) but blocks on a hung window,
    /// so it runs with a timeout and falls back to a copy of the screen area (only what is visible).
    /// </summary>
    public async Task<string> SnapAsync(long hwnd, string id, CancellationToken ct)
    {
        Directory.CreateDirectory(shotsDir);
        var path = Path.Combine(shotsDir, id + ".png");
        var capture = Task.Run(() => Screenshots.CaptureLocal(hwnd), ct);
        string method;
        Bitmap? bmp = await Task.WhenAny(capture, Task.Delay(3000, ct)) == capture ? capture.Result : null;
        if (bmp != null) method = "PrintWindow";
        else
        {
            GetWindowRect(new IntPtr(hwnd), out var r);
            if (r.R <= r.L || r.B <= r.T) throw new WatchdogRefusedException("the window has no area to capture");
            bmp = new Bitmap(r.R - r.L, r.B - r.T, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp)) g.CopyFromScreen(r.L, r.T, 0, 0, bmp.Size);
            method = "screen copy (window not responding: shows only what is visible)";
        }
        using (bmp) bmp.Save(path, ImageFormat.Png);
        return $"{path}{Wire.US}{method}";
    }

    /// <summary>Starts VB6.EXE. As the watchdog is elevated, the new IDE is too, with no UAC prompt (spike 12b).</summary>
    public string Launch(string? project)
    {
        var exe = Vb6Install.FindExe() ?? throw new WatchdogRefusedException("VB6.EXE not found");
        var psi = new ProcessStartInfo(exe) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(project ?? exe)! };
        if (project != null) psi.Arguments = $"\"{project}\"";
        try
        {
            using var p = Process.Start(psi) ?? throw new WatchdogRefusedException("VB6 didn't start");
            return p.Id.ToString(CultureInfo.InvariantCulture);
        }
        catch (Win32Exception ex) { throw new WatchdogRefusedException("VB6 didn't start: " + ex.Message); }
    }

    private static int GetParentPid()
    {
        var pbi = new PROCESS_BASIC_INFORMATION();
        return NtQueryInformationProcess(Process.GetCurrentProcess().Handle, 0, ref pbi, Marshal.SizeOf(pbi), out _) == 0
            ? pbi.InheritedFromUniqueProcessId.ToInt32() : 0;
    }

    private const byte VK_CONTROL = 0x11, VK_CANCEL = 0x03, VK_MENU = 0x12;
    private const uint KEYEVENTF_EXTENDEDKEY = 0x1, KEYEVENTF_KEYUP = 0x2, BM_CLICK = 0xF5;

    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)]
    private struct PROCESS_BASIC_INFORMATION { public IntPtr Reserved1, PebBaseAddress, Reserved2a, Reserved2b, UniqueProcessId, InheritedFromUniqueProcessId; }

    [DllImport("ntdll.dll")] private static extern int NtQueryInformationProcess(IntPtr h, int cls, ref PROCESS_BASIC_INFORMATION pbi, int size, out int ret);
    [DllImport("user32.dll")] private static extern void keybd_event(byte vk, byte scan, uint flags, int extra);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] private static extern void SwitchToThisWindow(IntPtr h, bool altTab);
    [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint a, uint b, bool attach);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out RECT r);
}
