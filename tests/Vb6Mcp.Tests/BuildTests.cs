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

    /// <summary>build_in_ide when the code compiles but the make fails (improvement notes 20): what can be seen from here.</summary>
    [Test]
    public static void MakeFailureCausesNameTheOutputInUseOrMissingFolder()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-make-").FullName;
        try
        {
            var exe = Path.Combine(dir, "App.exe");
            Assert.True(IdeTools.MakeFailureCauses(Path.Combine(dir, "nope", "App.exe"))[0].Contains("doesn't exist"), "missing folder");
            Assert.True(IdeTools.MakeFailureCauses(exe)[0].StartsWith("no cause visible"), "nothing wrong");
            File.WriteAllBytes(exe, [0x4D, 0x5A]);
            using (new FileStream(exe, FileMode.Open, FileAccess.Read, FileShare.Read))   // a running program holds its exe like this
                Assert.True(IdeTools.MakeFailureCauses(exe).Any(c => c.Contains("in use")), "held open");
        }
        finally { Directory.Delete(dir, true); }
    }

    [Test]
    public static void OutputFileDefaultsByProjectType()
    {
        var dll = VbpFile.Parse(@"D:\p\Com.vbp", "Type=OleDll\r\nName=\"ComFixture\"\r\n");
        Assert.Equal(@"D:\p\ComFixture.dll", IdeTools.OutputFileOf(dll));
        var named = VbpFile.Parse(@"D:\p\App.vbp", "Type=Exe\r\nName=\"App\"\r\nExeName32=\"Out.exe\"\r\nPath32=\"..\\bin\"\r\n");
        Assert.Equal(@"D:\bin\Out.exe", IdeTools.OutputFileOf(named));
    }

    /// <summary>A form VB couldn't load: the build log only points at frmX.log; the error carries that path (B2).</summary>
    [Test]
    public static void LoadErrorsPointAtTheFormLog()
    {
        var r = BuildLog.Parse("Errors during load. Refer to 'D:\\x\\frmTransferHistory.log' for details\r\n" +
                               "'D:\\x\\frmTransferHistory.frm' could not be loaded\r\nBuild of 'App.exe' failed.\r\n");
        Assert.True(!r.Succeeded, "failed");
        var load = r.Errors.Single(e => e.Kind == "load");
        Assert.Equal("D:\\x\\frmTransferHistory.log", load.File);
    }
}
