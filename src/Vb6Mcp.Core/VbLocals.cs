using System.Text;
using System.Text.RegularExpressions;

namespace Vb6Mcp.Core;

/// <summary>What declared a procedure-level name.</summary>
public enum LocalKind { Parameter, ReturnValue, Local, Static }

/// <summary>A name declared inside one procedure: a parameter, the function's return value, or a Dim/Static/ReDim.</summary>
public sealed record LocalDecl(string Name, LocalKind Kind, string? DeclaredType, bool IsArray);

/// <summary>
/// The names declared in one procedure, for debug_locals. VBIDE has no API for the Locals window
/// (it is custom-drawn, Phase 2), and <c>CodeModule.Members</c> lists only module-level members,
/// so the procedure's own declarations are scanned here: the signature, Dim, Static and ReDim.
/// This is deliberately small: it reads declarations only, never expressions (AGENTS.md principle 3).
/// </summary>
public static partial class VbLocals
{
    /// <summary>Declarations of the procedure in <paramref name="procText"/> (as proc_get returns it, comments above included).</summary>
    public static IReadOnlyList<LocalDecl> Parse(string procText)
    {
        var result = new List<LocalDecl>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        void Add(LocalDecl d) { if (seen.Add(d.Name)) result.Add(d); }

        bool inBody = false;
        foreach (var stmt in Statements(procText))
        {
            if (!inBody)
            {
                var sig = Signature().Match(stmt);
                if (!sig.Success) continue;
                inBody = true;
                foreach (var p in SplitTopLevel(sig.Groups["params"].Value))
                    if (ParseParam(p) is { } param) Add(param);
                var kind = sig.Groups["kind"].Value;
                if (kind.Equals("Function", StringComparison.OrdinalIgnoreCase) || kind.StartsWith("Property Get", StringComparison.OrdinalIgnoreCase))
                    Add(new LocalDecl(sig.Groups["name"].Value, LocalKind.ReturnValue, AsType(sig.Groups["rest"].Value), sig.Groups["rest"].Value.TrimStart().StartsWith("()")));
                continue;
            }
            if (EndProc().IsMatch(stmt)) break;
            var dim = DimStatement().Match(stmt);
            if (!dim.Success) continue;
            var keyword = dim.Groups["kw"].Value;
            var lk = keyword.Equals("Static", StringComparison.OrdinalIgnoreCase) ? LocalKind.Static : LocalKind.Local;
            foreach (var part in SplitTopLevel(dim.Groups["list"].Value))
            {
                var v = VarDecl().Match(part.Trim());
                if (!v.Success) continue;
                Add(new LocalDecl(v.Groups["name"].Value, lk, AsType(v.Groups["rest"].Value), v.Groups["dims"].Success));
            }
        }
        return result;
    }

    private static LocalDecl? ParseParam(string p)
    {
        var m = Param().Match(p.Trim());
        return m.Success ? new LocalDecl(m.Groups["name"].Value, LocalKind.Parameter, AsType(m.Groups["rest"].Value), m.Groups["arr"].Success || m.Groups["pa"].Success) : null;
    }

    private static string? AsType(string rest)
    {
        var m = AsClause().Match(rest);
        return m.Success ? m.Groups["type"].Value : null;
    }

    /// <summary>Logical statements: continuations joined, comments removed, split at ':' outside strings.</summary>
    public static IEnumerable<string> Statements(string text)
    {
        var joined = new StringBuilder();
        foreach (var raw in text.Replace("\r\n", "\n").Split('\n'))
        {
            var line = StripComment(raw);
            if (line.EndsWith(" _", StringComparison.Ordinal) || line == "_") { joined.Append(line[..^1]).Append(' '); continue; }
            joined.Append(line);
            foreach (var s in SplitColons(joined.ToString())) if (s.Trim().Length > 0) yield return s.Trim();
            joined.Clear();
        }
        if (joined.Length > 0) foreach (var s in SplitColons(joined.ToString())) if (s.Trim().Length > 0) yield return s.Trim();
    }

    private static string StripComment(string line)
    {
        bool inString = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"') inString = !inString;
            else if (!inString && c == '\'') return line[..i].TrimEnd();
        }
        var t = line.TrimStart();
        if (t.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase) || t.Equals("Rem", StringComparison.OrdinalIgnoreCase)) return "";
        return line.TrimEnd();
    }

    private static IEnumerable<string> SplitColons(string s)
    {
        bool inString = false; int start = 0;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '"') inString = !inString;
            // ':' followed by '=' is a named argument (x:=1), not a statement separator
            else if (!inString && s[i] == ':' && (i + 1 >= s.Length || s[i + 1] != '=')) { yield return s[start..i]; start = i + 1; }
        }
        yield return s[start..];
    }

    /// <summary>Splits at commas outside parentheses and strings.</summary>
    private static IEnumerable<string> SplitTopLevel(string s)
    {
        int depth = 0, start = 0; bool inString = false;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '"') inString = !inString;
            else if (inString) continue;
            else if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == ',' && depth == 0) { yield return s[start..i]; start = i + 1; }
        }
        if (s.Trim().Length > 0) yield return s[start..];
    }

    [GeneratedRegex(@"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?<name>[A-Za-z_]\w*)\s*\((?<params>.*)\)(?<rest>.*)$", RegexOptions.IgnoreCase)]
    private static partial Regex Signature();

    [GeneratedRegex(@"^End\s+(?:Sub|Function|Property)\b", RegexOptions.IgnoreCase)]
    private static partial Regex EndProc();

    [GeneratedRegex(@"^(?<kw>Dim|Static|ReDim(?:\s+Preserve)?)\s+(?<list>.+)$", RegexOptions.IgnoreCase)]
    private static partial Regex DimStatement();

    [GeneratedRegex(@"^(?:WithEvents\s+)?(?<name>[A-Za-z_]\w*)[%&!#@$]?\s*(?<dims>\([^)]*\))?(?<rest>.*)$", RegexOptions.IgnoreCase)]
    private static partial Regex VarDecl();

    [GeneratedRegex(@"^(?:Optional\s+)?(?:(?:ByVal|ByRef)\s+)?(?<pa>ParamArray\s+)?(?<name>[A-Za-z_]\w*)[%&!#@$]?\s*(?<arr>\(\s*\))?(?<rest>.*)$", RegexOptions.IgnoreCase)]
    private static partial Regex Param();

    [GeneratedRegex(@"\bAs\s+(?:New\s+)?(?<type>[A-Za-z_][\w.]*(?:\s*\*\s*\d+)?)", RegexOptions.IgnoreCase)]
    private static partial Regex AsClause();
}
