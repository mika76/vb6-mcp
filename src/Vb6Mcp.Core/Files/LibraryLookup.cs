using System.Diagnostics;
using System.Text.RegularExpressions;

namespace Vb6Mcp.Core.Files;

/// <summary>A file on disk that carries a wanted type library (not registered, or registered elsewhere).</summary>
public sealed record LibraryFile(string Path, int Major, int Minor, string Name);

/// <summary>
/// Hints for a control or reference that isn't registered (PLAN.md C3, improvement notes 19). The owner had CapturePRO 3
/// (CapturePRO3.dll, CAPTUREPRO3Lib) registered for a project that wants 2.0 (CapPRO.dll, CAPTUREPROLib, another LIBID),
/// and the v2 DLLs were already on disk, unregistered. Two lookups help: files that carry the wanted LIBID (searched by
/// the file names the project names, so it stays cheap), and registered libraries with a similar name (another version
/// or edition). Identity is still the LIBID (hard rule 13); names only point the agent and the user at candidates.
/// </summary>
public static class LibraryLookup
{
    /// <summary>
    /// Files named like <paramref name="fileNames"/> under <paramref name="folders"/> (each searched recursively or not)
    /// whose type library has <paramref name="libId"/>. Bounded by <paramref name="budget"/>.
    /// </summary>
    public static IReadOnlyList<LibraryFile> FilesWithLib(Guid libId, IEnumerable<string> fileNames, IEnumerable<(string Folder, bool Recursive)> folders, TimeSpan? budget = null)
    {
        var names = new HashSet<string>(fileNames.Where(n => n.Length > 0).Select(System.IO.Path.GetFileName).OfType<string>(), StringComparer.OrdinalIgnoreCase);
        var found = new List<LibraryFile>();
        if (names.Count == 0) return found;
        var clock = Stopwatch.StartNew();
        var limit = budget ?? TimeSpan.FromSeconds(8);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (folder, recursive) in folders)
        {
            if (!Directory.Exists(folder)) continue;
            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(folder, "*", new EnumerationOptions { RecurseSubdirectories = recursive, IgnoreInaccessible = true, AttributesToSkip = FileAttributes.ReparsePoint });
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { continue; }
            foreach (var f in files)
            {
                if (clock.Elapsed > limit) return found;
                if (!names.Contains(System.IO.Path.GetFileName(f)) || !seen.Add(PathKey.Normalize(f))) continue;
                if (TypeLibInfo.Read(f) is { } t && t.LibId == libId) found.Add(new LibraryFile(PathKey.Normalize(f), t.Major, t.Minor, t.Name));
            }
        }
        return found;
    }

    /// <summary>Registered libraries (other LIBIDs) whose description or file name is like one of <paramref name="names"/>.</summary>
    public static IReadOnlyList<RegisteredTypeLib> Similar(IEnumerable<string> names, IReadOnlyList<RegisteredTypeLib> registered, Guid exclude)
    {
        var keys = names.Select(Key).Where(k => k.Length >= 4).Distinct().ToList();
        if (keys.Count == 0) return [];
        bool Like(string candidate) => Key(candidate) is { Length: >= 4 } k && keys.Any(n => k.Contains(n, StringComparison.Ordinal) || n.Contains(k, StringComparison.Ordinal));
        return registered.Where(t => t.LibId != exclude && (Like(t.Description) || (t.Win32Path is { } p && Like(System.IO.Path.GetFileNameWithoutExtension(p)))))
            .GroupBy(t => (t.LibId, t.Major, t.Minor)).Select(g => g.First()).Take(10).ToList();
    }

    // Words that say nothing about which product a library is
    static readonly HashSet<string> Noise = new(StringComparer.OrdinalIgnoreCase)
        { "lib", "library", "type", "control", "controls", "activex", "ocx", "dll", "object", "objects", "component", "components", "for", "the", "version", "microsoft", "sp", "x" };

    /// <summary>A comparable key: lower-case letter runs without version numbers and noise words ("CAPTUREPRO3Lib", "CapturePRO 3.0 Control" → capturepro).</summary>
    public static string Key(string name)
    {
        var stem = System.IO.Path.GetFileNameWithoutExtension(name.Trim());
        stem = Regex.Replace(stem, @"(?<=[a-z0-9])Lib$", "", RegexOptions.IgnoreCase);   // CAPTUREPROLib, MSComctlLib
        var words = Regex.Matches(stem, "[A-Za-z]+").Select(m => m.Value).Where(w => !Noise.Contains(w));
        return string.Concat(words).ToLowerInvariant();
    }
}
