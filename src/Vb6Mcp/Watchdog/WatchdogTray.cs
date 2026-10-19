using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Vb6Mcp.Watchdog;

/// <summary>
/// The watchdog's notification-area icon: shows that it runs and whether it runs as administrator (a shield on the
/// icon, and the tooltip), with a menu to open its log, restart it as administrator, or stop it. It runs on its own STA
/// thread with a WinForms message loop, so it never touches the watchdog's verb threads, and the watchdog works the
/// same without it (--no-tray). Only the user acts through it: nothing here is reachable over TCP.
/// </summary>
public sealed class WatchdogTray : IDisposable
{
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new();
    private ContextMenuStrip? _menu;
    private NotifyIcon? _icon;

    /// <param name="elevated">Whether this watchdog runs elevated: the shield and the menu depend on it.</param>
    /// <param name="stay">Started with --stay (by the user) rather than by an add-in: shown in the tooltip.</param>
    /// <param name="logPath">watchdog.log, for "Open log".</param>
    /// <param name="restartElevated">Starts an elevated copy (UAC prompt); true when it started, so this one should stop.</param>
    /// <param name="stop">Stops this watchdog.</param>
    public WatchdogTray(bool elevated, bool stay, string logPath, Func<bool> restartElevated, Action stop)
    {
        _thread = new Thread(() => Run(elevated, stay, logPath, restartElevated, stop)) { IsBackground = true, Name = "tray" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        _ready.Wait(TimeSpan.FromSeconds(5));
    }

    private void Run(bool elevated, bool stay, string logPath, Func<bool> restartElevated, Action stop)
    {
        try
        {
            _menu = new ContextMenuStrip();
            _ = _menu.Handle;                       // a handle to BeginInvoke on from Dispose (another thread)
            var header = new ToolStripMenuItem($"vb6-mcp watchdog {WatchdogHost.Version}: {(elevated ? "administrator" : "not elevated")}") { Enabled = false };
            _menu.Items.Add(header);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add("Open log", null, (_, _) => Open(logPath));
            if (!elevated)
                _menu.Items.Add("Restart as administrator", null, (_, _) => { if (restartElevated()) stop(); });
            _menu.Items.Add("Stop watchdog", null, (_, _) => stop());
            _icon = new NotifyIcon
            {
                Icon = MakeIcon(elevated),
                // NotifyIcon.Text is limited to 63 characters
                Text = $"vb6-mcp watchdog: {(elevated ? "administrator" : "not elevated")}{(stay ? ", stays" : "")}",
                ContextMenuStrip = _menu,
                Visible = true,
            };
            _ready.Set();
            Application.Run();
        }
        catch (Exception) { _ready.Set(); }        // no tray (no desktop, say): the watchdog itself carries on
    }

    private static void Open(string path)
    {
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch (Exception ex) when (ex is Win32Exception or FileNotFoundException) { }
    }

    /// <summary>
    /// A drawn icon: a dark "VB" tile, with a yellow-blue shield in the corner when elevated (the Windows convention for
    /// "runs as administrator"), so the state shows without hovering.
    /// </summary>
    public static Icon MakeIcon(bool elevated)
    {
        using var bmp = new Bitmap(32, 32);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            using var tile = new SolidBrush(Color.FromArgb(0x2B, 0x3A, 0x55));
            using var path = Rounded(new Rectangle(1, 1, 30, 30), 6);
            g.FillPath(tile, path);
            using var font = new Font("Segoe UI", 13f, FontStyle.Bold, GraphicsUnit.Pixel);
            g.DrawString("VB", font, Brushes.White, elevated ? new PointF(1, 2) : new PointF(5, 7));
            if (elevated)
            {
                // shield: left half blue, right half yellow, white outline
                var shield = new GraphicsPath();
                shield.AddLines(new[] { new PointF(21, 15), new PointF(31, 18), new PointF(31, 24), new PointF(26, 31), new PointF(21, 24) });
                shield.AddLines(new[] { new PointF(21, 24), new PointF(16, 31), new PointF(11, 24), new PointF(11, 18), new PointF(21, 15) });
                using var white = new Pen(Color.White, 2f);
                g.DrawPath(white, shield);
                g.SetClip(new RectangleF(11, 14, 10, 18));
                g.FillPath(new SolidBrush(Color.FromArgb(0x1E, 0x6F, 0xD9)), shield);
                g.SetClip(new RectangleF(21, 14, 11, 18));
                g.FillPath(new SolidBrush(Color.FromArgb(0xF5, 0xC2, 0x18)), shield);
                g.ResetClip();
            }
        }
        return Icon.FromHandle(bmp.GetHicon());
    }

    private static GraphicsPath Rounded(Rectangle r, int radius)
    {
        var p = new GraphicsPath();
        int d = radius * 2;
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    /// <summary>Removes the icon (else it lingers in the notification area until the mouse passes over it) and ends the loop.</summary>
    public void Dispose()
    {
        var menu = _menu;
        if (menu is null || !menu.IsHandleCreated) return;
        try
        {
            menu.Invoke(() =>
            {
                if (_icon is not null) { _icon.Visible = false; _icon.Dispose(); }
                Application.ExitThread();
            });
        }
        catch (Exception ex) when (ex is InvalidOperationException or ObjectDisposedException) { }
        _thread.Join(TimeSpan.FromSeconds(2));
    }
}
