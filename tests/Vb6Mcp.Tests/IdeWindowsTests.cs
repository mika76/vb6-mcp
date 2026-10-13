using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

public static class IdeWindowsTests
{
    [Test]
    public static void ParsesModeFromRealTitles()
    {
        // Titles observed in spikes 2, 3 and 12
        Assert.Equal(IdeMode.Design, IdeWindows.ParseMode("Project1 - Microsoft Visual Basic [design]"));
        Assert.Equal(IdeMode.Run, IdeWindows.ParseMode("Project1 - Microsoft Visual Basic [run]"));
        Assert.Equal(IdeMode.Break, IdeWindows.ParseMode("Project1 - Microsoft Visual Basic [break]"));
        Assert.Equal(IdeMode.Run, IdeWindows.ParseMode("Project1 - Microsoft Visual Basic [run] - [Form1 (Code)]"));
        Assert.Equal(IdeMode.Design, IdeWindows.ParseMode("Microsoft Visual Basic [design]"));
        Assert.Equal(IdeMode.Unknown, IdeWindows.ParseMode("Microsoft Visual Basic"));
        Assert.Equal(IdeMode.Unknown, IdeWindows.ParseMode(""));
    }

    [Test, Category("IDE")]
    public static void ReadsAtLeastOneRunningIde()
    {
        var vb = System.Diagnostics.Process.GetProcessesByName("VB6");
        Assert.True(vb.Length > 0, "start VB6 first");
        var s = IdeWindows.Read(vb[0].Id);
        Assert.True(s.HasMainWindow, "main window found");
        Assert.True(s.Mode != IdeMode.Unknown, $"mode parsed from '{s.Title}'");
    }
}
