using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Tools;

namespace Vb6Mcp.Tests;

public static class BuildTests
{
    [Test]
    public static void LogsFromTheSpike()
    {
        // Real logs from spike 21
        var ok = BuildLog.Parse("Build of 'Basic.exe' succeeded.");
        Assert.True(ok.Succeeded && ok.Errors.Count == 0, "success");
        var bad = BuildLog.Parse("Compile Error in File 'D:\\s\\modMath.bas', Line 8 : Variable not defined\r\nBuild of 'Basic.exe' failed.");
        Assert.True(!bad.Succeeded && bad.Errors.Count == 1, "one error; the failed line is the status");
        Assert.Equal(9, bad.Errors[0].Line ?? 0, "/make counts IDE lines from 0; the tools use 1-based IDE lines");
        Assert.Equal("compile", bad.Errors[0].Kind);
        var compat = BuildLog.Parse("There is a compatibility error between the current project and the version-compatible component: 'Hello' in the 'clsGreeter' class module has arguments and/or a return type that is incompatible with a similar declaration in the version-compatible component.\r\nBuild of 'ComFixture.dll' failed.");
        Assert.True(!compat.Succeeded && compat.Errors[0].Kind == "compatibility", "compatibility error");
        Assert.Equal("clsGreeter", compat.Errors[0].Class);
        Assert.Equal("Hello", compat.Errors[0].Member);
        var open = BuildLog.Parse("The file 'D:\\x\\Basic.vbp' is already open.\r\nThe file 'D:\\x\\Basic.vbp' is already open.");
        Assert.True(!open.Succeeded && open.Errors.Count == 1 && open.Errors[0].Kind == "other", "repeated line kept once");
    }
    [Test]
    public static void OutputFileDefaultsByProjectType()
    {
        var dll = VbpFile.Parse(@"D:\p\Com.vbp", "Type=OleDll\r\nName=\"ComFixture\"\r\n");
        Assert.Equal(@"D:\p\ComFixture.dll", IdeTools.OutputFileOf(dll));
        var named = VbpFile.Parse(@"D:\p\App.vbp", "Type=Exe\r\nName=\"App\"\r\nExeName32=\"Out.exe\"\r\nPath32=\"..\\bin\"\r\n");
        Assert.Equal(@"D:\bin\Out.exe", IdeTools.OutputFileOf(named));
    }
}
