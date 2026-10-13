using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

public static class ProjectMapTests
{
    static void Write(string path, string text)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, Vb1252.Encoding.GetBytes(text.Replace("\n", "\r\n")));
    }

    [Test]
    public static void IndexesSharedFilesAcrossProjects()
    {
        var root = Directory.CreateTempSubdirectory("vb6mcp-map-").FullName;
        try
        {
            Write(Path.Combine(root, "Shared", "modShared.bas"), "Attribute VB_Name = \"modShared\"\n");
            Write(Path.Combine(root, "App1", "App1.vbp"), "Type=Exe\nModule=modShared; ..\\Shared\\modShared.bas\nModule=modA; modA.bas\nName=\"App1\"\n");
            Write(Path.Combine(root, "Libs", "Lib2", "Lib2.vbp"), "Type=OleDll\nModule=modShared; ..\\..\\Shared\\modShared.bas\nName=\"Lib2\"\n");
            Write(Path.Combine(root, "App1", "bin", "Ignored.vbp"), "Type=Exe\nName=\"Ignored\"\n");      // bin is skipped

            var map = ProjectMap.Build([root]);
            Assert.SequenceEqual(new[] { "App1", "Lib2" }, map.Projects.Select(p => p.Name).OrderBy(n => n));
            var users = map.ProjectsUsing(Path.Combine(root, "shared", "MODSHARED.BAS"));      // other spelling
            Assert.SequenceEqual(new[] { "App1", "Lib2" }, users.Select(p => p.Name).OrderBy(n => n));
            Assert.Equal(1, map.SharedFiles().Count(), "one shared file");
            Assert.Equal(1, map.ProjectsUsing(Path.Combine(root, "App1", "modA.bas")).Count);
            Assert.True(ProjectMap.IsUnderRoots(Path.Combine(root, "App1", "x.bas"), [root]), "inside the root");
            Assert.True(!ProjectMap.IsUnderRoots(@"C:\Windows\win.ini", [root]), "outside the root");
        }
        finally { Directory.Delete(root, true); }
    }

    /// <summary>
    /// Spike 9 (scan time): set VB6MCP_ROOTS to the real workspace roots and run --cat Scan.
    /// Prints counts and timing only, no file contents.
    /// </summary>
    [Test, Category("Scan")]
    public static void ScanRealRoots()
    {
        var roots = (Environment.GetEnvironmentVariable("VB6MCP_ROOTS") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries);
        Assert.True(roots.Length > 0, "set VB6MCP_ROOTS");
        var map = ProjectMap.Build(roots);
        Console.WriteLine($"        roots: {string.Join("; ", map.Roots)}");
        Console.WriteLine($"        projects: {map.Projects.Count}, member files: {map.Projects.Sum(p => p.Members.Count)}, " +
                          $"shared files: {map.SharedFiles().Count()}, errors: {map.Errors.Count}, scan: {map.ScanTime.TotalMilliseconds:0} ms");
        var types = map.Projects.GroupBy(p => p.Type).Select(g => $"{g.Key}={g.Count()}");
        Console.WriteLine($"        types: {string.Join(", ", types)}");
    }
}
