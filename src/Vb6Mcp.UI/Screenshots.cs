using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace Vb6Mcp.UI;

/// <summary>
/// Screenshot helpers. Capturing an elevated VB6 window from the non-elevated server fails
/// (PrintWindow is blocked, spike 15), so for elevated IDEs the add-in captures in-process and writes
/// a BMP; for non-elevated IDEs the server captures itself. Both end up as PNG for MCP image content.
/// </summary>
public static class Screenshots
{
    /// <summary>
    /// Captures a window with PrintWindow from this process. Returns null when it fails (e.g. the
    /// window belongs to an elevated process).
    /// </summary>
    public static Bitmap? CaptureLocal(long hwnd)
    {
        var h = new IntPtr(hwnd);
        if (!GetWindowRect(h, out var r) || r.R <= r.L || r.B <= r.T) return null;
        var bmp = new Bitmap(r.R - r.L, r.B - r.T, PixelFormat.Format32bppArgb);
        bool ok;
        using (var g = Graphics.FromImage(bmp)) { var dc = g.GetHdc(); ok = PrintWindow(h, dc, 2 /*PW_RENDERFULLCONTENT*/); g.ReleaseHdc(dc); }
        if (ok) return bmp;
        bmp.Dispose();
        return null;
    }

    /// <summary>Loads a BMP written by the add-in (fully into memory, so the file can be deleted).</summary>
    public static Bitmap LoadBmp(string path)
    {
        using var fs = new MemoryStream(File.ReadAllBytes(path));
        using var tmp = new Bitmap(fs);
        return new Bitmap(tmp);
    }

    /// <summary>PNG bytes, scaled down proportionally when wider than <paramref name="maxWidth"/>.</summary>
    public static byte[] ToPng(Bitmap source, int maxWidth)
    {
        Bitmap img = source;
        if (maxWidth > 0 && source.Width > maxWidth)
        {
            int h = Math.Max(1, source.Height * maxWidth / source.Width);
            img = new Bitmap(maxWidth, h, PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(img);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawImage(source, 0, 0, maxWidth, h);
        }
        using var ms = new MemoryStream();
        img.Save(ms, ImageFormat.Png);
        if (!ReferenceEquals(img, source)) img.Dispose();
        return ms.ToArray();
    }

    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int L, T, R, B; }
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
}
