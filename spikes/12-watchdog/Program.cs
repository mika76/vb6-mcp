// Spike 12: elevated watchdog sidecar. Started by the VB6 add-in (inherits elevation, no UAC).
// One instance per session (named mutex), loopback TCP + token, discovery file
// %LOCALAPPDATA%\vb6mcp\watchdog.json, log %LOCALAPPDATA%\vb6mcp\watchdog.log.
// Exits 2 minutes after the last VB6 is gone.
// Verbs: PING | STATUS | SCREENSHOT <pid> [ide|dialog|form|<hex hwnd>] | BREAK <pid> | KILL <pid> force

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;

string baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp");
Directory.CreateDirectory(baseDir);
string logPath = Path.Combine(baseDir, "watchdog.log");
void Log(string m) { try { File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{Environment.ProcessId}] {m}\r\n"); } catch { } }

using var mutex = new Mutex(true, @"Local\VB6McpWatchdog", out bool firstInstance);
if (!firstInstance) { Log("another watchdog is running; exiting"); return; }

Native.SetProcessDPIAware();
bool elevated = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
int parentPid = Native.ParentPid();
DateTime started = DateTime.Now;
string token = Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();

var listener = new TcpListener(IPAddress.Loopback, 0);
listener.Start();
int port = ((IPEndPoint)listener.LocalEndpoint).Port;
string discovery = Path.Combine(baseDir, "watchdog.json");
File.WriteAllText(discovery, $"{{\"pid\":{Environment.ProcessId},\"port\":{port},\"token\":\"{token}\",\"version\":\"spike12\"}}");
Log($"started port={port} elevated={elevated} parent={parentPid}");

// Idle exit: no VB6 for 2 minutes
_ = Task.Run(async () =>
{
    var lastSeen = DateTime.Now;
    while (true)
    {
        await Task.Delay(5000);
        if (Process.GetProcessesByName("VB6").Length > 0) lastSeen = DateTime.Now;
        else if ((DateTime.Now - lastSeen).TotalMinutes >= 2)
        {
            Log("no VB6 for 2 minutes; exiting");
            try { File.Delete(discovery); } catch { }
            Environment.Exit(0);
        }
    }
});

while (true)
{
    var client = await listener.AcceptTcpClientAsync();
    _ = Task.Run(() => Serve(client));
}

void Serve(TcpClient client)
{
    using (client)
    {
        var s = client.GetStream();
        var line = ReadLine(s);
        string reply; bool ok = true;
        try
        {
            var sp = line.IndexOf(' ');
            if (sp < 0 || line[..sp] != token) { ok = false; reply = "bad token"; }
            else
            {
                var parts = line[(sp + 1)..].Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                reply = parts[0].ToUpperInvariant() switch
                {
                    "PING" => $"pong pid={Environment.ProcessId} elevated={elevated} parent={parentPid} parentAlive={Alive(parentPid)} uptime={(int)(DateTime.Now - started).TotalSeconds}s",
                    "STATUS" => Status(),
                    "SCREENSHOT" => Screenshot(int.Parse(parts[1]), parts.Length > 2 ? parts[2] : "ide"),
                    "BREAK" => Break(int.Parse(parts[1])),
                    "KILL" => parts.Length > 2 && parts[2] == "force" ? Kill(int.Parse(parts[1])) : throw new Exception("KILL needs 'force'"),
                    "PRESS" => Press(int.Parse(parts[1]), string.Join(' ', parts.Skip(2))),
                    "LAUNCH" => Launch(string.Join(' ', parts.Skip(1))),
                    _ => throw new Exception("unknown verb " + parts[0]),
                };
            }
        }
        catch (Exception ex) { ok = false; reply = ex.Message; }
        var body = Encoding.Latin1.GetBytes(reply);
        var head = Encoding.Latin1.GetBytes($"{(ok ? "OK" : "ERR")} {body.Length}\r\n");
        s.Write(head); s.Write(body);
        Log($"{line.Split(' ').Skip(1).FirstOrDefault()} -> {(ok ? "OK" : "ERR")} {reply.Split('\n')[0]}");
    }
}

static string ReadLine(NetworkStream s)
{
    var sb = new StringBuilder(); int b;
    while ((b = s.ReadByte()) >= 0 && b != '\n') if (b != '\r') sb.Append((char)b);
    return sb.ToString();
}

static bool Alive(int pid) { try { return !Process.GetProcessById(pid).HasExited; } catch { return false; } }

// ---- health ----
static string Status()
{
    var sb = new StringBuilder();
    foreach (var p in Process.GetProcessesByName("VB6"))
    {
        var main = Native.MainIde((uint)p.Id);
        string title = Native.Cap(main);                       // cached caption: no message sent
        // The mode tag isn't always last: a maximized MDI child appends " - [Form1 (Code)]"
        var m = System.Text.RegularExpressions.Regex.Match(title, @"\[(design|run|break)\]");
        string mode = m.Success ? m.Groups[1].Value : "?";
        var sw = Stopwatch.StartNew();
        bool responds = Native.Responds(main, 1000);
        sw.Stop();
        bool hungFlag = Native.IsHungAppWindow(main);
        var dialogs = Native.Dialogs((uint)p.Id);
        string state = responds ? "responsive"
            : mode == "run" ? "busy-running"
            : dialogs.Count > 0 ? "blocked-dialog"
            : "hung";
        sb.AppendLine($"pid={p.Id} state={state} mode={mode} responds={responds} ({sw.ElapsedMilliseconds}ms) isHung={hungFlag} title='{title}'");
        foreach (var d in dialogs) sb.AppendLine($"    dialog: {d}");
    }
    // Crash detection: add-in discovery files whose process is gone (a clean unload deletes the file)
    var inst = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp", "instances");
    if (Directory.Exists(inst))
        foreach (var f in Directory.GetFiles(inst, "*.json"))
            if (int.TryParse(Path.GetFileNameWithoutExtension(f), out var dead) && !Alive(dead))
                sb.AppendLine($"pid={dead} state=crashed (discovery file left behind, last written {File.GetLastWriteTime(f):HH:mm:ss})");
    return sb.Length == 0 ? "no VB6 running" : sb.ToString();
}

// ---- screenshots: PrintWindow with timeout, screen copy fallback ----
string Screenshot(int pid, string target)
{
    IntPtr h = target switch
    {
        "ide" => Native.MainIde((uint)pid),
        "dialog" => Native.FirstTopLevel((uint)pid, "#32770"),
        "form" => Native.FirstTopLevel((uint)pid, "ThunderFormDC", "ThunderRT6FormDC"),
        _ => new IntPtr(Convert.ToInt64(target, 16)),
    };
    if (h == IntPtr.Zero) throw new Exception("no window for target " + target);
    Native.GetWindowRect(h, out var r);
    int w = r.R - r.L, hh = r.B - r.T;
    var bmp = new Bitmap(w, hh, PixelFormat.Format32bppArgb);
    string method = "PrintWindow";
    var sw = Stopwatch.StartNew();
    var pw = Task.Run(() =>
    {
        using var g = Graphics.FromImage(bmp);
        var dc = g.GetHdc();
        try { return Native.PrintWindow(h, dc, 2); } finally { g.ReleaseHdc(dc); }
    });
    bool done = pw.Wait(3000);
    if (!done || !pw.Result)
    {
        method = done ? "screen (PrintWindow failed)" : "screen (PrintWindow timed out)";
        bmp = new Bitmap(w, hh, PixelFormat.Format32bppArgb);   // abandon the blocked bitmap
        using var g = Graphics.FromImage(bmp);
        g.CopyFromScreen(r.L, r.T, 0, 0, new Size(w, hh));
    }
    sw.Stop();
    string dir = Path.Combine(baseDir, "shots");
    Directory.CreateDirectory(dir);
    string file = Path.Combine(dir, $"{DateTime.Now:HHmmss}_{pid}_{target}.png");
    bmp.Save(file, ImageFormat.Png);
    return $"{file} {w}x{hh} method={method} {sw.ElapsedMilliseconds}ms";
}

// ---- break a busy loop from outside the IDE thread: inject Ctrl+Break ----
static string Break(int pid)
{
    var main = Native.MainIde((uint)pid);
    string before = Native.Cap(main);
    Native.CtrlBreak();
    Thread.Sleep(1500);
    return $"injected Ctrl+Break; title before='{before}' after='{Native.Cap(main)}'";
}

// Start VB6 on a project. We're elevated, so the child inherits elevation (no UAC prompt expected).
static string Launch(string project)
{
    project = project.Trim('"');
    var ext = Path.GetExtension(project).ToLowerInvariant();
    if (ext is not (".vbp" or ".vbg")) throw new Exception("LAUNCH needs a .vbp or .vbg path");
    if (!File.Exists(project)) throw new Exception("not found: " + project);
    const string vb6 = @"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE";
    var p = Process.Start(new ProcessStartInfo(vb6, $"\"{project}\"") { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(project)! })!;
    var sw = Stopwatch.StartNew();
    IntPtr main = IntPtr.Zero;
    while (sw.ElapsedMilliseconds < 20000 && (main = Native.MainIde((uint)p.Id)) == IntPtr.Zero) Thread.Sleep(250);
    Thread.Sleep(1000);
    var dialogs = Native.Dialogs((uint)p.Id);
    return $"launched pid={p.Id} window={(main != IntPtr.Zero)} after {sw.ElapsedMilliseconds}ms title='{Native.Cap(main)}' dialogs={dialogs.Count}" +
           string.Concat(dialogs.Select(d => "\n    dialog: " + d));
}

// Click a button (by caption, '&' ignored) in a visible dialog of the VB6 process
static string Press(int pid, string caption)
{
    var btn = Native.DialogButton((uint)pid, caption);
    if (btn == IntPtr.Zero) throw new Exception($"no dialog button '{caption}'");
    Native.Click(btn);
    Thread.Sleep(500);
    return $"pressed '{caption}'; dialogs now: {Native.Dialogs((uint)pid).Count}";
}

static string Kill(int pid)
{
    var p = Process.GetProcessById(pid);
    if (!p.ProcessName.Equals("VB6", StringComparison.OrdinalIgnoreCase)) throw new Exception("not a VB6 process");
    p.Kill();
    p.WaitForExit(5000);
    return $"killed {pid}; exited={p.HasExited}";
}

static class Native
{
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] struct PBI { public IntPtr Reserved1, PebBase, Reserved2a, Reserved2b, UniqueProcessId, InheritedFrom; }
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc p, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsHungAppWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
    [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("ntdll.dll")] static extern int NtQueryInformationProcess(IntPtr h, int cls, ref PBI info, int len, out int ret);

    public static string Cls(IntPtr h) { var s = new StringBuilder(64); GetClassName(h, s, 64); return s.ToString(); }
    public static string Cap(IntPtr h) { var s = new StringBuilder(512); GetWindowText(h, s, 512); return s.ToString(); }

    public static int ParentPid()
    {
        var pbi = new PBI();
        NtQueryInformationProcess(Process.GetCurrentProcess().Handle, 0, ref pbi, Marshal.SizeOf<PBI>(), out _);
        return pbi.InheritedFrom.ToInt32();
    }

    public static IntPtr MainIde(uint pid) => FirstTopLevel(pid, "wndclass_desked_gsk");

    public static IntPtr FirstTopLevel(uint pid, params string[] classes)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((h, _) =>
        {
            GetWindowThreadProcessId(h, out var p);
            if (p == pid && IsWindowVisible(h) && classes.Contains(Cls(h))) { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static bool Responds(IntPtr h, uint timeoutMs)
        => SendMessageTimeout(h, 0 /*WM_NULL*/, IntPtr.Zero, IntPtr.Zero, 0x2 /*SMTO_ABORTIFHUNG*/, timeoutMs, out _) != IntPtr.Zero;

    public static List<string> Dialogs(uint pid)
    {
        var l = new List<string>();
        EnumWindows((h, _) =>
        {
            GetWindowThreadProcessId(h, out var p);
            if (p == pid && IsWindowVisible(h) && Cls(h) == "#32770")
            {
                var texts = new List<string>();
                EnumChildWindows(h, (c, _) => { var t = Cap(c); if (t.Length > 0) texts.Add($"{Cls(c)}:'{t.Replace("\r\n", " ")}'"); return true; }, IntPtr.Zero);
                l.Add($"{h.ToInt64():X} '{Cap(h)}' {string.Join(" | ", texts)}");
            }
            return true;
        }, IntPtr.Zero);
        return l;
    }

    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);

    public static IntPtr DialogButton(uint pid, string caption)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((h, _) =>
        {
            GetWindowThreadProcessId(h, out var p);
            if (p != pid || !IsWindowVisible(h) || Cls(h) != "#32770") return true;
            EnumChildWindows(h, (c, _) =>
            {
                if (Cls(c) == "Button" && string.Equals(Cap(c).Replace("&", ""), caption, StringComparison.OrdinalIgnoreCase)) { found = c; return false; }
                return true;
            }, IntPtr.Zero);
            return found == IntPtr.Zero;
        }, IntPtr.Zero);
        return found;
    }

    // BM_CLICK is posted: SendMessage would wait for the dialog's thread
    public static void Click(IntPtr button) => PostMessage(button, 0x00F5 /*BM_CLICK*/, IntPtr.Zero, IntPtr.Zero);

    public static void CtrlBreak()
    {
        const byte VK_CONTROL = 0x11, VK_CANCEL = 0x03; const uint KEYUP = 0x2;
        keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
        keybd_event(VK_CANCEL, 0x46, 0, UIntPtr.Zero);
        keybd_event(VK_CANCEL, 0x46, KEYUP, UIntPtr.Zero);
        keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero);
    }
}
