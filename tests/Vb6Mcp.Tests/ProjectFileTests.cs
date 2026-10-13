using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

public static class ProjectFileTests
{
    static string RepoRoot()
    {
        var dir = AppContext.BaseDirectory;
        while (!File.Exists(Path.Combine(dir, "vb6-mcp.slnx"))) dir = Path.GetDirectoryName(dir)!;
        return dir;
    }

    [Test]
    public static void PathKeyNormalizesRelativeSlashesAndCase()
    {
        var a = PathKey.Normalize(@"..\Shared/Utils.bas", @"C:\Work\App1");
        Assert.Equal(@"C:\Work\Shared\Utils.bas", a);
        Assert.True(PathKey.Same(@"c:\work\shared\UTILS.BAS", a), "case-insensitive identity");
        Assert.True(PathKey.Same(@"C:\Work\App2\..\Shared\Utils.bas", a), "same file via another project");
        Assert.Equal(@"C:\Work\App1\x.bas", PathKey.Normalize("\"x.bas\"", @"C:\Work\App1"), "quotes stripped");
    }

    [Test]
    public static void ParsesTheBasicFixture()
    {
        var vbp = VbpFile.Load(Path.Combine(RepoRoot(), "tests", "fixtures", "Basic", "Basic.vbp"));
        Assert.Equal("Basic", vbp.Name);
        Assert.Equal("Exe", vbp.Type);
        Assert.Equal("frmMain", vbp.Startup);
        Assert.SequenceEqual(new[] { "Form", "Module", "Class" }, vbp.Members.Select(m => m.Kind));
        Assert.Equal("modMath", vbp.Members[1].Name);
        Assert.Equal<string?>(null, vbp.Members[0].Name, "Form= lines have no name");
        Assert.True(vbp.Members.All(m => File.Exists(m.FullPath)), "member paths resolve to existing files");
    }

    [Test]
    public static void ParsesSharedPathsReferencesObjectsAndOutput()
    {
        var text = string.Join("\r\n",
            "Type=OleDll",
            @"Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\Windows\SysWOW64\stdole2.tlb#OLE Automation",
            "Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX",
            "Form=Forms\\frmA.frm",
            @"Module=modShared; ..\..\Shared\modShared.bas",
            "Class=clsX; clsX.cls",
            "Name=\"MyLib\"",
            "ExeName32=\"MyLib.dll\"",
            "Path32=\"..\\bin\"",
            "CompatibleMode=\"2\"",
            "CompatibleEXE32=\"..\\compat\\MyLib.dll\"",
            "[MS Transaction Server]",
            "AutoRefresh=1");
        var vbp = VbpFile.Parse(@"C:\Src\Libs\MyLib\MyLib.vbp", text);
        Assert.Equal("MyLib", vbp.Name);
        Assert.Equal(@"C:\Src\Shared\modShared.bas", vbp.Members[1].FullPath, "climbs out of the project folder");
        Assert.Equal(@"C:\Src\Libs\MyLib\Forms\frmA.frm", vbp.Members[0].FullPath);
        Assert.Equal(1, vbp.References.Count);
        Assert.True(vbp.References[0].StartsWith(@"*\G{00020430"), "reference kept raw");
        Assert.Equal(1, vbp.Objects.Count);
        Assert.Equal(@"C:\Src\Libs\bin\MyLib.dll", vbp.OutputPath);
        Assert.Equal(2, vbp.CompatibleMode);
        Assert.Equal(@"C:\Src\Libs\compat\MyLib.dll", vbp.CompatibleFile);
        Assert.Equal(3, vbp.Members.Count, "lines after [section] ignored");
    }
}
