using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

/// <summary>Hints for unregistered controls and references (PLAN.md C3, improvement notes 19).</summary>
public static class LibraryLookupTests
{
    [Test]
    public static void KeysIgnoreVersionsSuffixesAndNoiseWords()
    {
        Assert.Equal("capturepro", LibraryLookup.Key("CAPTUREPROLib"));
        Assert.Equal("capturepro", LibraryLookup.Key("CAPTUREPRO3Lib"));
        Assert.Equal("capturepro", LibraryLookup.Key("CapturePRO 3.0 Control"));
        Assert.Equal("capturepro", LibraryLookup.Key("CapturePRO3.dll"));
        Assert.Equal("cappro", LibraryLookup.Key("cappro.dll"));
    }

    [Test]
    public static void SimilarRegisteredLibrariesAreOtherLibIds()
    {
        Guid v2 = Guid.NewGuid(), v3 = Guid.NewGuid();
        var registered = new List<RegisteredTypeLib>
        {
            new(v3, 3, 0, "CapturePRO 3.0 Control", @"C:\Windows\SysWOW64\CapturePRO3.dll"),
            new(Guid.NewGuid(), 1, 0, "Microsoft Scripting Runtime", @"C:\Windows\SysWOW64\scrrun.dll"),
        };
        // the v2 file on disk names its typelib CAPTUREPROLib; the project only says cappro.dll
        var similar = LibraryLookup.Similar(["CAPTUREPROLib", "cappro.dll"], registered, v2);
        Assert.Equal(1, similar.Count);
        Assert.Equal(v3, similar[0].LibId);
        Assert.Equal(0, LibraryLookup.Similar(["cappro.dll"], registered, v2).Count);   // the file name alone isn't like CapturePRO
    }

    [Test]
    public static void AFileCarryingTheLibraryIsFoundByName()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-lib-").FullName;
        try
        {
            var sub = Path.Combine(dir, "Tools", "Lib"); Directory.CreateDirectory(sub);
            var scrrun = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SysWOW64", "scrrun.dll");
            File.Copy(scrrun, Path.Combine(sub, "myscript.dll"));
            var lib = TypeLibInfo.Read(scrrun)!.LibId;
            var found = LibraryLookup.FilesWithLib(lib, ["MyScript.dll"], [(dir, true)]);
            Assert.Equal(1, found.Count);
            Assert.True(found[0].Path.EndsWith("myscript.dll", StringComparison.OrdinalIgnoreCase) && found[0].Name == "Scripting", found[0].ToString());
            Assert.Equal(0, LibraryLookup.FilesWithLib(lib, ["MyScript.dll"], [(dir, false)]).Count);   // not recursive: not found
            Assert.Equal(0, LibraryLookup.FilesWithLib(Guid.NewGuid(), ["MyScript.dll"], [(dir, true)]).Count);   // another LIBID
        }
        finally { Directory.Delete(dir, true); }
    }
}
