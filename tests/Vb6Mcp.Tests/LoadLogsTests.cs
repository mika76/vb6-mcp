using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

/// <summary>Summing up VB6's form load logs (spike 36): many "Errors during load", a few missing classes.</summary>
public static class LoadLogsTests
{
    [Test]
    public static void ManyLogsBecomeAFewMissingClasses()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-loadlogs-").FullName;
        try
        {
            void Log(string form, string text) => File.WriteAllBytes(Path.Combine(dir, form + ".log"), Vb1252.Encoding.GetBytes(text));
            // the shapes seen in the owner's workspace
            Log("frmAdj", "Line 16: Class XmPanelCtrl.XmPanel of control XmPanel1 was not a loaded control class.\r\n");
            Log("frmLocations", "Line 33: Class XmTextCtrl.XmText of control XmText1 was not a loaded control class.\r\n" +
                                        "Line 56: Class XmTextCtrl.XmText of control XmText2 was not a loaded control class.\r\n");
            Log("frmFloats", "Line 12: Class XmTextCtrl.XmText of control XmText1 was not a loaded control class.\r\n");
            Log("frmMain", "Line 691: Cannot load control WinHook1; license not found.\r\n");
            var logs = new[] { "frmAdj", "frmLocations", "frmFloats", "frmMain" }.Select(f => Path.Combine(dir, f + ".log"));

            var s = LoadLogs.Summarize(logs);
            Assert.Equal("XmTextCtrl.XmText", s[0].Class);
            Assert.Equal(3, s[0].Controls);
            Assert.Equal(2, s[0].Forms.Count);
            Assert.True(s.Any(p => p.Class == "XmPanelCtrl.XmPanel" && p.Controls == 1), "the other class");
            Assert.True(s.Any(p => p.Class == "control WinHook1" && p.Reason.StartsWith("licence")), "the licence case");
            Assert.Equal(@"D:\p\frmMain.log", LoadLogs.LogOf(@"Errors during load. Refer to 'D:\p\frmMain.log' for details"));
        }
        finally { Directory.Delete(dir, true); }
    }
}
