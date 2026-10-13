using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace Vb6Mcp.UI;

/// <summary>IDE mode as shown in the main window title.</summary>
public enum IdeMode { Unknown, Design, Run, Break }

/// <summary>A button of a dialog: caption without '&amp;' and its window handle (for clicking through the add-in).</summary>
public sealed record DialogButton(string Caption, long Hwnd);

/// <summary>A visible dialog (#32770) of the VB6 process: caption, static texts, buttons.</summary>
public sealed record DialogInfo(long Hwnd, string Caption, IReadOnlyList<string> Texts, IReadOnlyList<DialogButton> Buttons);

/// <summary>What can be read about an IDE without the add-in (works while the IDE thread is busy).</summary>
public sealed record IdeWindowState(int Pid, bool HasMainWindow, string Title, IdeMode Mode, bool Responds, bool IsHung, IReadOnlyList<DialogInfo> Dialogs);

/// <summary>
/// Channel C reads of the VB6 IDE from the non-elevated side. Reads cross the elevation boundary
/// (spike 10); input doesn't, so nothing here sends input.
/// </summary>
public static partial class IdeWindows
{
    /// <summary>Window class of the VB6 IDE main window.</summary>
    public const string MainWindowClass = "wndclass_desked_gsk";

    /// <summary>
    /// Mode from the title. The tag isn't always last: with a maximized MDI child the title is
    /// "Project1 - Microsoft Visual Basic [run] - [Form1 (Code)]" (spike 12).
    /// </summary>
    public static IdeMode ParseMode(string title)
    {
        var m = ModeTag().Match(title);
        return !m.Success ? IdeMode.Unknown : m.Groups[1].Value switch
        {
            "design" => IdeMode.Design,
            "run" => IdeMode.Run,
            "break" => IdeMode.Break,
            _ => IdeMode.Unknown,
        };
    }

    [GeneratedRegex(@"\[(design|run|break)\]")]
    private static partial Regex ModeTag();

    /// <summary>Reads title, mode, responsiveness and dialogs. <paramref name="probeMs"/> bounds the WM_NULL probe.</summary>
    public static IdeWindowState Read(int pid, uint probeMs = 1000)
    {
        var main = FindTopLevel((uint)pid, MainWindowClass);
        if (main == IntPtr.Zero)
            return new IdeWindowState(pid, false, "", IdeMode.Unknown, false, false, Dialogs(pid));
        // GetWindowText on another process's top-level window returns the cached caption without
        // sending a message, so it works even when the IDE thread is busy (spike 12).
        string title = Caption(main);
        bool responds = SendMessageTimeout(main, 0 /*WM_NULL*/, IntPtr.Zero, IntPtr.Zero, 0x2 /*SMTO_ABORTIFHUNG*/, probeMs, out _) != IntPtr.Zero;
        return new IdeWindowState(pid, true, title, ParseMode(title), responds, IsHungAppWindow(main), Dialogs(pid));
    }

    /// <summary>Visible #32770 dialogs of the process with their static texts and buttons.</summary>
    public static IReadOnlyList<DialogInfo> Dialogs(int pid)
    {
        var list = new List<DialogInfo>();
        EnumWindows((h, _) =>
        {
            GetWindowThreadProcessId(h, out var p);
            if (p != (uint)pid || !IsWindowVisible(h) || ClassName(h) != "#32770") return true;
            var texts = new List<string>();
            var buttons = new List<DialogButton>();
            EnumChildWindows(h, (c, _) =>
            {
                if (!IsWindowVisible(c)) return true;
                string cls = ClassName(c), cap = Caption(c);
                if (cap.Length == 0) return true;
                if (cls == "Button") buttons.Add(new DialogButton(cap.Replace("&", "", StringComparison.Ordinal), c.ToInt64()));
                else if (cls == "Static") texts.Add(cap);
                return true;
            }, IntPtr.Zero);
            list.Add(new DialogInfo(h.ToInt64(), Caption(h), texts, buttons));
            return true;
        }, IntPtr.Zero);
        return list;
    }

    private static IntPtr FindTopLevel(uint pid, string cls)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((h, _) =>
        {
            GetWindowThreadProcessId(h, out var p);
            if (p == pid && IsWindowVisible(h) && ClassName(h) == cls) { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private static string ClassName(IntPtr h) { var s = new StringBuilder(128); _ = GetClassName(h, s, s.Capacity); return s.ToString(); }
    private static string Caption(IntPtr h) { var s = new StringBuilder(1024); _ = GetWindowText(h, s, s.Capacity); return s.ToString(); }

    private delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc proc, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr lParam);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsHungAppWindow(IntPtr hwnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder s, int n);
    [DllImport("user32.dll")] private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
}
