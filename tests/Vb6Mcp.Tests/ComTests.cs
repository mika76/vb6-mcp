using Vb6Mcp.Tools;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

public static class ComTests
{
    static VbpFile Vbp(string dir, string text) => VbpFile.Parse(Path.Combine(dir, "P.vbp"), text.Replace("\n", "\r\n"));

    /// <summary>Scripting.Dictionary is registered on every Windows: ProgID -> CLSID -> scrrun.dll, with its typelib.</summary>
    [Test]
    public static void ReadsClassRegistrations()
    {
        var clsid = ComRegistry.ClsidOf("Scripting.Dictionary");
        Assert.True(clsid is not null, "Scripting.Dictionary registered");
        var c = ComRegistry.Class(clsid!.Value)!;
        Assert.True(c.ServerExists && c.ServerFile!.EndsWith("scrrun.dll", StringComparison.OrdinalIgnoreCase), c.ServerFile ?? "(none)");
        Assert.True(c.TypeLib is not null, "class names its typelib");
        // the class is a coclass of the typelib in its own server file
        var d = TypeLibDescriber.Describe(c.ServerFile!);
        Assert.True(d.Info.LibId == c.TypeLib && d.Types.Any(t => t.Kind == "coclass" && t.Guid == clsid), "coclass in the file's typelib");
        Assert.True(ComRegistry.FindProgIds("Scripting.Dict").Any(p => p.Clsid == clsid), "found by search");
        Assert.True(ComRegistry.ClsidOf("No.Such.ProgId.vb6mcp") is null, "unknown ProgID");
    }

    /// <summary>com_register's verdict (improvement notes 18): a class nothing registers doesn't fail it; one registered elsewhere does.</summary>
    [Test]
    public static void RegistrationVerdict()
    {
        Assert.True(WorkspaceTools.Verified(false, 0, allHere: false, anyHere: true, libHere: true, elsewhere: 0, missing: 1, classes: 6), "CapPRO: 5 of 6, the sixth unregistered anywhere");
        Assert.True(!WorkspaceTools.Verified(false, 0, false, true, true, elsewhere: 1, missing: 0, classes: 6), "a class pointing at another file");
        Assert.True(!WorkspaceTools.Verified(false, 0, false, false, libHere: false, 0, 0, 2), "typelib not registered here");
        Assert.True(!WorkspaceTools.Verified(false, 5, false, true, true, 0, 1, 6), "regsvr32 failed");
        Assert.True(!WorkspaceTools.Verified(false, 0, false, true, true, 0, missing: 3, classes: 3), "no class registered at all");
        Assert.True(WorkspaceTools.Verified(true, 0, false, anyHere: false, false, 0, 0, 3) && !WorkspaceTools.Verified(true, 0, false, anyHere: true, false, 0, 0, 3), "unregister");
    }

    /// <summary>
    /// Default members (improvement notes 14): Dictionary's is Item; MSComm makes the hidden alias _CommPort the default,
    /// which is reported as the visible CommPort.
    /// </summary>
    [Test]
    public static void DefaultMembersAreMarked()
    {
        var win = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        var text = TypeLibDescriber.Render(TypeLibDescriber.Describe(Path.Combine(win, "SysWOW64", "scrrun.dll")), "Dictionary");
        Assert.True(text.Split('\n').Any(l => l.StartsWith("Class Dictionary") && l.Contains("default member: Item")), text);
        Assert.True(text.Split('\n').Any(l => l.Contains("Property Get Item(") && l.Contains("' default member")), "the member itself is marked: " + text);
        var comm = Path.Combine(win, "SysWOW64", "MSCOMM32.OCX");
        if (File.Exists(comm))
            Assert.True(TypeLibDescriber.Render(TypeLibDescriber.Describe(comm)).Contains("default member: CommPort (as hidden _CommPort)"), "MSComm");
    }

    /// <summary>
    /// A VB-built DLL hides each class's own interface (_clsPricing): its methods are still listed by default, and the
    /// interface's restricted IDispatch plumbing isn't (benchmark T15). The committed compatibility file of tests\fixtures\Shared.
    /// </summary>
    [Test]
    public static void AVbClassShowsItsMethods()
    {
        var dll = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "fixtures", "Shared", "SharedLib", "compat", "SharedLib.dll"));
        var text = TypeLibDescriber.Render(TypeLibDescriber.Describe(dll));
        Assert.True(text.Contains("Function PriceText(ByVal net As Currency, ByVal taxRate As Double) As String"), text);
        Assert.True(!text.Contains("QueryInterface"), "IDispatch plumbing stays hidden: " + text);
        // improvement notes 45: not even with include_hidden; the hidden class interface itself still shows
        var all = TypeLibDescriber.Render(TypeLibDescriber.Describe(dll), includeHidden: true);
        foreach (var m in new[] { "QueryInterface", "AddRef", "Release(", "GetTypeInfoCount", "GetTypeInfo(", "GetIDsOfNames", "Invoke(" })
            Assert.True(!all.Contains(m), $"{m} with include_hidden: " + all);
        Assert.True(all.Contains("Function PriceText("), all);
    }

    /// <summary>
    /// PE machine type: the 32-bit scrrun.dll in SysWOW64 vs the 64-bit one (System32, reached as Sysnative from this
    /// 32-bit process). VB6 programs are 32-bit, so com_diagnose flags 64-bit servers.
    /// </summary>
    [Test]
    public static void BitnessAndDebugRegistrations()
    {
        var win = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        Assert.Equal("x86", ComRegistry.Bitness(Path.Combine(win, "SysWOW64", "scrrun.dll")));
        var native = Path.Combine(win, Environment.Is64BitProcess ? "System32" : "Sysnative", "scrrun.dll");
        if (File.Exists(native)) Assert.Equal("x64", ComRegistry.Bitness(native));
        Assert.True(ComRegistry.Bitness(Path.Combine(win, "win.ini")) is null, "not a PE file");
        Assert.True(ComRegistry.IsVbDebugServer(@"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6DEBUG.DLL"), "debugger stub");
        Assert.True(ComRegistry.IsProjectTypeLib(new RegisteredTypeLib(Guid.NewGuid(), 1, 0, "", @"D:\x\P.vbp")), "typelib registered to a .vbp");
    }

    [Test]
    public static void CompatFindings()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-compat-").FullName;
        try
        {
            File.WriteAllText(Path.Combine(dir, "Lib.dll"), "x");
            Directory.CreateDirectory(Path.Combine(dir, "compat"));
            File.WriteAllText(Path.Combine(dir, "compat", "Lib.dll"), "x");
            // generated, never hand-written (hard rule 13)
            Guid a = Guid.NewGuid(), b = Guid.NewGuid();
            var none = Compat.Check(Vbp(dir, "Type=OleDll\nName=\"Lib\"\nExeName32=\"Lib.dll\"\nCompatibleMode=\"0\"\n"), _ => new TypeLibInfo(a, 1, 0, "Lib"), []);
            Assert.True(none.Findings.Any(f => f.Text.Contains("not binary")), "no binary compatibility warned");

            var ok = Compat.Check(Vbp(dir, "Type=OleDll\nName=\"Lib\"\nExeName32=\"Lib.dll\"\nCompatibleMode=\"2\"\nCompatibleEXE32=\"compat\\Lib.dll\"\n"),
                _ => new TypeLibInfo(a, 1, 0, "Lib"), [new RegisteredTypeLib(a, 1, 0, "Lib", Path.Combine(dir, "Lib.dll"))]);
            Assert.Equal(0, ok.Findings.Count, string.Join("; ", ok.Findings.Select(f => f.Text)));

            var broken = Compat.Check(Vbp(dir, "Type=OleDll\nName=\"Lib\"\nExeName32=\"Lib.dll\"\nCompatibleMode=\"2\"\nCompatibleEXE32=\"compat\\Lib.dll\"\n"),
                // the compatibility file is the one whose parent folder is "compat" (the temp folder's name contains "compat" too)
                f => new TypeLibInfo(Path.GetFileName(Path.GetDirectoryName(f)) == "compat" ? a : b, 1, 0, "Lib"), []);
            Assert.True(broken.Findings.Any(f => f.Level == "error" && f.Text.Contains("differs")), "different LIBIDs = error");
            Assert.True(broken.Findings.Any(f => f.Text.Contains("isn't registered")), "unregistered warned");

            var self = Compat.Check(Vbp(dir, "Type=OleDll\nName=\"Lib\"\nExeName32=\"Lib.dll\"\nCompatibleMode=\"2\"\nCompatibleEXE32=\"Lib.dll\"\n"), _ => new TypeLibInfo(a, 1, 0, "Lib"), []);
            Assert.True(self.Findings.Any(f => f.Text.Contains("build output itself")), "reference = output warned");

            var missing = Compat.Check(Vbp(dir, "Type=Control\nName=\"Ctl\"\nExeName32=\"Ctl.ocx\"\nCompatibleMode=\"2\"\nCompatibleEXE32=\"gone\\Ctl.ocx\"\n"), _ => null, []);
            Assert.True(missing.Findings.Any(f => f.Level == "error" && f.Text.Contains("doesn't exist")), "missing reference file");

            Assert.Equal("1", Vbp(dir, "Type=OleDll\nMajorVer=1\n").Settings["MajorVer"], "version settings kept");
        }
        finally { Directory.Delete(dir, true); }
    }
}
