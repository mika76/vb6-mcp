using System.Text.RegularExpressions;

namespace Vb6Mcp.Core.Files;

/// <summary>A reference or control whose version VB changed in the .vbp (e.g. mscomctl 2.1 to 2.2).</summary>
public sealed record VbpVersionChange(string Kind, string Guid, string Name, string OldVersion, string NewVersion);

/// <summary>What changed in a .vbp between two reads: lines removed and added, and version changes of references and controls.</summary>
public sealed record VbpChanges(IReadOnlyList<string> Removed, IReadOnlyList<string> Added, IReadOnlyList<VbpVersionChange> Versions)
{
    public bool Any => Removed.Count > 0 || Added.Count > 0;
}

/// <summary>
/// Compares two versions of a .vbp. Saving a project that VB counts as changed rewrites the .vbp as well as the dirty
/// components. In a production project (2026-10-17) that rewrote 17 Reference=/Object= lines nobody asked for: absolute hint
/// paths, re-cased OCX names, and control versions bumped to what's registered (mscomctl 2.1 to 2.2). project_save
/// reports it with this.
/// </summary>
public static partial class VbpDiff
{
    public static VbpChanges Compare(string before, string after)
    {
        var a = Lines(before); var b = Lines(after);
        // multiset difference keeps duplicates honest without depending on line order (VB may reorder)
        var removed = Subtract(a, b); var added = Subtract(b, a);
        var versions = new List<VbpVersionChange>();
        foreach (var r in removed)
        {
            if (Parse(r) is not { } old) continue;
            var now = added.Select(Parse).FirstOrDefault(n => n is not null && n.Value.Kind == old.Kind && n.Value.Guid.Equals(old.Guid, StringComparison.OrdinalIgnoreCase));
            if (now is { } n && n.Version != old.Version) versions.Add(new VbpVersionChange(old.Kind, old.Guid, n.Name.Length > 0 ? n.Name : old.Name, old.Version, n.Version));
        }
        return new VbpChanges(removed, added, versions);
    }

    private static List<string> Lines(string text) => text.Replace("\r\n", "\n").Split('\n').Where(l => l.Length > 0).ToList();

    private static List<string> Subtract(List<string> x, List<string> y)
    {
        var left = y.GroupBy(l => l).ToDictionary(g => g.Key, g => g.Count());
        var result = new List<string>();
        foreach (var l in x)
        {
            if (left.TryGetValue(l, out int n) && n > 0) left[l] = n - 1;
            else result.Add(l);
        }
        return result;
    }

    // Object={GUID}#2.1#0; MSCOMCTL.OCX      Reference=*\G{GUID}#2.0#0#C:\path\x.dll#Description
    private static (string Kind, string Guid, string Version, string Name)? Parse(string line)
    {
        var m = RefLine().Match(line);
        if (!m.Success) return null;
        string kind = m.Groups["k"].Value.Equals("Object", StringComparison.OrdinalIgnoreCase) ? "control" : "reference";
        string rest = m.Groups["rest"].Value;
        string name = kind == "control" ? (rest.Contains(';') ? rest[(rest.IndexOf(';') + 1)..].Trim() : "")
                                        : rest.Split('#') is { Length: > 0 } parts ? parts[^1] : "";
        return (kind, m.Groups["g"].Value, m.Groups["v"].Value, name);
    }

    [GeneratedRegex(@"^(?<k>Object|Reference)=(\*\\G)?(?<g>\{[0-9A-Fa-f-]{36}\})#(?<v>[0-9A-Fa-f]+\.[0-9A-Fa-f]+)(?<rest>.*)$", RegexOptions.IgnoreCase)]
    private static partial Regex RefLine();
}
