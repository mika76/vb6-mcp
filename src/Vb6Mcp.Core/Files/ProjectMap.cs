using System.Diagnostics;

namespace Vb6Mcp.Core.Files;

/// <summary>
/// In-memory index of every .vbp under the workspace roots: file (normalized path) → projects that
/// include it. This is all that's left of the "code server" idea (PLAN.md §9): shared source files
/// make "which projects use this file?" essential, and parsing .vbp files is cheap enough to rescan.
/// </summary>
public sealed class ProjectMap
{
    private static readonly string[] s_skipDirs = [".git", ".vs", "bin", "obj", "node_modules"];

    private readonly Dictionary<string, List<VbpFile>> _byFile = new(StringComparer.Ordinal);

    public IReadOnlyList<string> Roots { get; }
    public IReadOnlyList<VbpFile> Projects { get; }
    /// <summary>.vbp files that couldn't be read, with the reason.</summary>
    public IReadOnlyList<(string Path, string Error)> Errors { get; }
    public TimeSpan ScanTime { get; }
    public DateTime ScannedAt { get; }

    private ProjectMap(IReadOnlyList<string> roots, List<VbpFile> projects, List<(string, string)> errors, TimeSpan time)
    {
        Roots = roots; Projects = projects; Errors = errors; ScanTime = time; ScannedAt = DateTime.UtcNow;
        foreach (var p in projects)
            foreach (var m in p.Members)
            {
                var key = m.FullPath.ToUpperInvariant();
                if (!_byFile.TryGetValue(key, out var list)) _byFile[key] = list = [];
                if (!list.Contains(p)) list.Add(p);
            }
    }

    /// <summary>Scans the roots recursively for .vbp files (skips .git, .vs, bin, obj, node_modules).</summary>
    public static ProjectMap Build(IEnumerable<string> roots)
    {
        var sw = Stopwatch.StartNew();
        var rootList = roots.Select(r => PathKey.Normalize(r)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var projects = new List<VbpFile>();
        var errors = new List<(string, string)>();
        foreach (var root in rootList.Where(Directory.Exists))
            foreach (var vbp in EnumerateVbp(root))
            {
                try { projects.Add(VbpFile.Load(vbp)); }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { errors.Add((vbp, ex.Message)); }
            }
        return new ProjectMap(rootList, projects, errors, sw.Elapsed);
    }

    private static IEnumerable<string> EnumerateVbp(string root)
    {
        var pending = new Stack<string>();
        pending.Push(root);
        var opts = new EnumerationOptions { IgnoreInaccessible = true, RecurseSubdirectories = false, AttributesToSkip = FileAttributes.ReparsePoint };
        while (pending.Count > 0)
        {
            var dir = pending.Pop();
            foreach (var f in Directory.EnumerateFiles(dir, "*.vbp", opts)) yield return f;
            foreach (var d in Directory.EnumerateDirectories(dir, "*", opts))
                if (!s_skipDirs.Contains(Path.GetFileName(d), StringComparer.OrdinalIgnoreCase)) pending.Push(d);
        }
    }

    /// <summary>Projects whose member list includes <paramref name="file"/> (any spelling of its path).</summary>
    public IReadOnlyList<VbpFile> ProjectsUsing(string file)
        => _byFile.TryGetValue(PathKey.Key(file), out var list) ? list : [];

    /// <summary>Files included by more than one project.</summary>
    public IEnumerable<(string File, IReadOnlyList<VbpFile> Projects)> SharedFiles()
        => _byFile.Where(kv => kv.Value.Count > 1)
                  .Select(kv => (kv.Value[0].Members.First(m => m.FullPath.Equals(kv.Key, StringComparison.OrdinalIgnoreCase)).FullPath, (IReadOnlyList<VbpFile>)kv.Value));

    /// <summary>True if <paramref name="path"/> lies under one of the roots.</summary>
    public static bool IsUnderRoots(string path, IEnumerable<string> roots)
    {
        var key = PathKey.Key(path);
        return roots.Any(r => { var rk = PathKey.Key(r).TrimEnd('\\') + "\\"; return key.StartsWith(rk, StringComparison.Ordinal); });
    }
}
