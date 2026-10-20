using System.Drawing;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

public static class AppToolsTests
{
    [Test]
    public static void ScreenshotCropKeepsFullResolutionAndClipsToTheImage()
    {
        // improvement notes 48: only the part that matters, unscaled
        using var bmp = new Bitmap(100, 50);
        bmp.SetPixel(12, 7, Color.Red);
        using (var part = Screenshots.Crop(bmp, 10, 5, 20, 10))
        {
            Assert.True(part is not null && part.Width == 20 && part.Height == 10, "20 x 10");
            Assert.Equal(Color.Red.ToArgb(), part!.GetPixel(2, 2).ToArgb());
        }
        using (var edge = Screenshots.Crop(bmp, 90, 40, 30, 30))
            Assert.True(edge is not null && edge.Width == 10 && edge.Height == 10, "clipped to the image");
        Assert.True(Screenshots.Crop(bmp, 200, 0, 10, 10) is null, "outside: nothing");
        using var round = Screenshots.FromPng(Screenshots.ToPng(bmp, 0));
        Assert.Equal(Color.Red.ToArgb(), round.GetPixel(12, 7).ToArgb());
    }

    [Test]
    public static void ControlKindsFromWindowClasses()
    {
        Assert.Equal("CommandButton", AppWindows.Kind("ThunderCommandButton"));       // in the IDE
        Assert.Equal("CommandButton", AppWindows.Kind("ThunderRT6CommandButton"));    // compiled
        Assert.Equal("TextBox", AppWindows.Kind("ThunderRT6TextBox"));
        Assert.Equal("Form", AppWindows.Kind("ThunderFormDC"));
        Assert.Equal("Form", AppWindows.Kind("ThunderRT6FormDC"));
        Assert.Equal("Dialog", AppWindows.Kind("#32770"));
        Assert.Equal("VScrollBar", AppWindows.Kind("ThunderVScrollBar"));
        Assert.Equal("PictureBox", AppWindows.Kind("ThunderPictureBoxDC"));      // seen live (Phase 3)
        Assert.Equal("MSFlexGridWndClass", AppWindows.Kind("MSFlexGridWndClass"), "unknown classes pass through");
    }

    [Test]
    public static void PngIsScaledToMaxWidth()
    {
        using var bmp = new Bitmap(800, 400);
        using var big = Image.FromStream(new MemoryStream(Screenshots.ToPng(bmp, 400)));
        Assert.Equal(400, big.Width);
        Assert.Equal(200, big.Height);
        using var same = Image.FromStream(new MemoryStream(Screenshots.ToPng(bmp, 1024)));
        Assert.Equal(800, same.Width, "never scaled up");
    }

    [Test]
    public static void AuditAppendsOneLinePerAction()
    {
        var old = Audit.LogPath;
        var path = Path.Combine(Path.GetTempPath(), $"vb6mcp-audit-{Guid.NewGuid():N}.log");
        try
        {
            Audit.LogPath = path;
            Audit.Write("app_set_text", 42, "1A2B 'two\r\nlines'");
            Audit.Write("app_click", 42, "1A2C CommandButton 'Add'");
            var lines = File.ReadAllLines(path);
            Assert.Equal(2, lines.Length, "newlines in details are escaped");
            Assert.True(lines[0].Contains("app_set_text pid=42") && lines[0].Contains(@"two\r\nlines"), lines[0]);
        }
        finally { Audit.LogPath = old; File.Delete(path); }
    }
}
