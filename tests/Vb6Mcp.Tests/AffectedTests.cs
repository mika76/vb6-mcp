using Microsoft.Win32;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

public static class AffectedTests
{
    static void Write(string path, string text)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, Vb1252.Encoding.GetBytes(text.Replace("\n", "\r\n")));
    }

    /// <summary>
    /// The reader against a typelib every Windows has: stdole2.tlb. Its LIBID must be the key under
    /// HKCR\TypeLib (32-bit view) through which VB resolves "Reference=*\G{...}#2.0#..." lines.
    /// </summary>
    [Test]
    public static void ReadsTypeLibIdentityWithoutRegistering()
    {
        var tlb = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.SystemX86), "stdole2.tlb");
        var t = TypeLibInfo.Read(tlb);
        Assert.True(t is not null, "stdole2.tlb read");
        Assert.Equal("stdole", t!.Name);
        Assert.Equal(2, t.Major);
        using var key = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32).OpenSubKey($@"TypeLib\{t.LibId:B}\{t.Major:x}.{t.Minor:x}");
        Assert.True(key is not null, "registered under the same LIBID and version");
        // the line VB writes for it parses back to the same identity
        var dep = ComDependency.Parse($@"*\G{t.LibId:B}#2.0#0#{tlb}#OLE Automation");
        Assert.True(dep is not null && dep.LibId == t.LibId && dep.Major == 2, "Reference= line parsed");
        Assert.True(TypeLibInfo.Read(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.SystemX86), "notepad.exe")) is null, "no typelib: null");
        Assert.True(TypeLibInfo.Read(@"C:\no\such\file.dll") is null, "missing file: null");
    }

    [Test]
    public static void ParsesReferenceAndObjectLines()
    {
        var g = Guid.NewGuid();
        Assert.Equal(1, ComDependency.Parse($"{g:B}#1.0#0; MyCtl.ocx")!.Major, "Object= line");
        Assert.Equal(11, ComDependency.Parse($@"*\G{g:B}#b.0#0#C:\x.dll#X")!.Major, "hex major like the registry");
        Assert.True(ComDependency.Parse(@"*\A..\Lib\Lib.vbp") is null, "project reference not parsed");
        Assert.True(ComDependency.Parse("garbage") is null, "garbage");
    }

    /// <summary>
    /// Shared module used by a DLL and an EXE; the EXE and a second EXE reference the DLL; an OCX using the DLL
    /// is used by a third EXE (Object= line). Expected order: DLL, then its users, then the OCX's user.
    /// </summary>
    [Test]
    public static void RebuildOrderFollowsTypeLibReferences()
    {
        var root = Directory.CreateTempSubdirectory("vb6mcp-aff-").FullName;
        try
        {
            // GUIDs here are generated, never hand-written (hard rule 13); the typelib reader is faked
            Guid lib = Guid.NewGuid(), ctl = Guid.NewGuid(), other = Guid.NewGuid();
            Write(Path.Combine(root, "Shared", "modShared.bas"), "Attribute VB_Name = \"modShared\"\n");
            Write(Path.Combine(root, "Lib", "Lib.vbp"), "Type=OleDll\nModule=modShared; ..\\Shared\\modShared.bas\nName=\"Lib\"\nExeName32=\"Lib.dll\"\nCompatibleMode=\"2\"\n");
            Write(Path.Combine(root, "App", "App.vbp"), $"Type=Exe\nReference=*\\G{lib:B}#1.0#0#..\\Lib\\Lib.dll#Lib\nModule=modShared; ..\\Shared\\modShared.bas\nName=\"App\"\n");
            Write(Path.Combine(root, "Ctl", "Ctl.vbp"), $"Type=Control\nReference=*\\G{lib:B}#1.0#0#..\\Lib\\Lib.dll#Lib\nName=\"Ctl\"\nExeName32=\"Ctl.ocx\"\n");
            Write(Path.Combine(root, "Host", "Host.vbp"), $"Type=Exe\nObject={ctl:B}#1.0#0; Ctl.ocx\nName=\"Host\"\n");
            Write(Path.Combine(root, "Old", "Old.vbp"), $"Type=Exe\nReference=*\\G{lib:B}#2.0#0#Lib.dll#Lib\nName=\"Old\"\n");      // other major
            Write(Path.Combine(root, "Free", "Free.vbp"), $"Type=Exe\nReference=*\\G{other:B}#1.0#0#x.dll#X\nName=\"Free\"\n");
            var map = ProjectMap.Build([root]);
            TypeLibInfo? Fake(VbpFile p) => p.Name switch { "Lib" => new(lib, 1, 0, "Lib"), "Ctl" => new(ctl, 1, 0, "Ctl"), _ => null };

            var r = Affected.Find(map, [Path.Combine(root, "shared", "MODSHARED.BAS")], Fake);
            Assert.SequenceEqual(new[] { "Lib", "App", "Ctl", "Host" }, r.Order.Select(x => x.Project.Name));
            var app = r.Order.Single(x => x.Project.Name == "App");
            Assert.Equal(1, app.Includes.Count, "App includes the file itself");
            Assert.SequenceEqual(new[] { "Lib" }, app.Uses.Select(u => u.Name));
            Assert.SequenceEqual(new[] { "Ctl" }, r.Order.Single(x => x.Project.Name == "Host").Uses.Select(u => u.Name));
            Assert.Equal(0, r.Warnings.Count, string.Join("; ", r.Warnings));

            // a DLL that was never built: its users can't be found, and that's a warning
            var unbuilt = Affected.Find(map, [Path.Combine(root, "Shared", "modShared.bas")], _ => null);
            Assert.SequenceEqual(new[] { "App", "Lib" }, unbuilt.Order.Select(x => x.Project.Name));
            Assert.True(unbuilt.Warnings.Any(w => w.StartsWith("Lib: no typelib")), string.Join("; ", unbuilt.Warnings));

            var none = Affected.Find(map, [Path.Combine(root, "Shared", "nothing.bas")], Fake);
            Assert.True(none.Order.Count == 0 && none.Warnings.Count == 1, "unused file: warning");
        }
        finally { Directory.Delete(root, true); }
    }

    [Test]
    public static void CyclesAreReportedNotLooped()
    {
        var root = Directory.CreateTempSubdirectory("vb6mcp-aff-").FullName;
        try
        {
            Guid a = Guid.NewGuid(), b = Guid.NewGuid();
            Write(Path.Combine(root, "A", "A.vbp"), $"Type=OleDll\nModule=m; m.bas\nReference=*\\G{b:B}#1.0#0#B.dll#B\nName=\"A\"\n");
            Write(Path.Combine(root, "B", "B.vbp"), $"Type=OleDll\nReference=*\\G{a:B}#1.0#0#A.dll#A\nName=\"B\"\n");
            var map = ProjectMap.Build([root]);
            var r = Affected.Find(map, [Path.Combine(root, "A", "m.bas")], p => new TypeLibInfo(p.Name == "A" ? a : b, 1, 0, p.Name));
            Assert.SequenceEqual(new[] { "A", "B" }, r.Order.Select(x => x.Project.Name));
            Assert.True(r.Warnings.Any(w => w.StartsWith("circular")), string.Join("; ", r.Warnings));
        }
        finally { Directory.Delete(root, true); }
    }
}
