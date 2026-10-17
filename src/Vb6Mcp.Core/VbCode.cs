using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.Text;

namespace Vb6Mcp.Core;

/// <summary>
/// Helpers for code text exchanged with the IDE: hashing for optimistic concurrency, validation of
/// text before it is written (AGENTS.md hard rules 1 and 10), and procedure-range trimming.
/// </summary>
public static partial class VbCode
{
    /// <summary>VB6's hard limit on physical line length.</summary>
    public const int MaxLineLength = 1023;
    /// <summary>VB6 allows at most 24 line continuations in one logical line.</summary>
    public const int MaxContinuations = 24;

    private static readonly Encoding s_strict1252;

    // Field initializers run before the static constructor body, so the strict encoding must be created here, after
    // Vb1252.Encoding has registered the code-pages provider. As an initializer it threw (TypeInitializationException)
    // whenever an edit tool was the first thing a server process did (improvement notes session, 2026-10-17).
    static VbCode()
    {
        _ = Vb1252.Encoding;
        s_strict1252 = Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
    }

    /// <summary>Short content hash (first 16 hex chars of SHA-256 over the 1252 bytes). Write tools compare it.</summary>
    public static string Hash(string text)
        => Convert.ToHexString(SHA256.HashData(Vb1252.Encoding.GetBytes(text)))[..16].ToLowerInvariant();

    /// <summary>
    /// Normalizes and validates text to be inserted: LF/CR → CR LF, one trailing line break removed
    /// (InsertLines would otherwise add an empty line), then checks the VB6 limits and the encoding.
    /// Throws <see cref="ArgumentException"/> with a message meant for the agent.
    /// </summary>
    public static string PrepareForInsert(string text)
    {
        var t = text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");
        if (t.EndsWith("\r\n", StringComparison.Ordinal)) t = t[..^2];
        if (t.Contains('\x1F') || t.Contains('\x1E')) throw new ArgumentException("text contains control characters 0x1E/0x1F");
        try { s_strict1252.GetBytes(t); }
        catch (EncoderFallbackException ex)
        {
            throw new ArgumentException($"text contains a character that Windows-1252 can't represent: U+{(int)ex.CharUnknown:X4} (VB6 source is ANSI)");
        }
        var lines = t.Split("\r\n");
        int run = 0;
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].Length > MaxLineLength)
                throw new ArgumentException($"line {i + 1} has {lines[i].Length} characters; VB6 allows at most {MaxLineLength}");
            run = lines[i].EndsWith(" _", StringComparison.Ordinal) ? run + 1 : 0;
            if (run > MaxContinuations)
                throw new ArgumentException($"more than {MaxContinuations} line continuations ending at line {i + 1}");
        }
        return t;
    }

    /// <summary>Lines of a module text (CR LF separated), 1-based slice. Out-of-range parts are clamped.</summary>
    public static string Slice(string text, int startLine, int count)
    {
        var lines = text.Length == 0 ? [] : text.Split("\r\n");
        int first = Math.Max(1, startLine) - 1;
        int n = Math.Max(0, Math.Min(count, lines.Length - first));
        return n <= 0 ? "" : string.Join("\r\n", lines, first, n);
    }

    /// <summary>
    /// VB6's procedure start includes the blank lines above it (it starts right after the previous
    /// procedure). Drops those, so the range starts at the first comment or the declaration.
    /// </summary>
    /// <summary>
    /// Identifiers the agent typed that VB now spells differently. VB keeps one spelling per name for the whole project
    /// and re-cases new text to it, so retyping Err.number as Err.Number can leave the module, and its hash, unchanged
    /// (improvement notes item 24). Returns "typed -> kept" for each name in <paramref name="typed"/> that appears in
    /// <paramref name="module"/> only with another spelling. Text in string literals and comments is ignored.
    /// </summary>
    public static IReadOnlyList<string> CaseChanges(string typed, string module)
    {
        var kept = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var id in Identifiers(module))
        {
            if (!kept.TryGetValue(id, out var set)) kept[id] = set = new HashSet<string>(StringComparer.Ordinal);
            set.Add(id);
        }
        var result = new List<string>();
        foreach (var id in Identifiers(typed).Distinct(StringComparer.Ordinal))
            if (kept.TryGetValue(id, out var spellings) && !spellings.Contains(id))
                result.Add($"{id} -> {spellings.First()}");
        return result;
    }

    /// <summary>
    /// C-style escapes in VB string literals of code an agent submits (owner's suggestion, 2026-10-17). VB6 has no escape
    /// sequences: "a\nb" is a backslash and an n, and "C:\\Temp" has two backslashes. Agents used to C, JSON or JavaScript
    /// write both. Returns one hint per kind found, empty when the code is fine. A leading \\ (a UNC path) is left alone.
    /// </summary>
    public static IReadOnlyList<string> LiteralHints(string code)
    {
        var hints = new List<string>();
        bool newline = false, doubled = false;
        foreach (Match m in StringLiteral().Matches(code))
        {
            var s = m.Groups[1].Value;
            if (Regex.IsMatch(s, @"\\[nrt]") && !Regex.IsMatch(s, @"^[A-Za-z]:\\|^\\\\")) newline = true;   // a path like "C:\new" is fine
            if (s.IndexOf(@"\\", 1, StringComparison.Ordinal) > 0) doubled = true;
        }
        if (newline) hints.Add(@"A string literal contains \n, \r or \t: VB6 has no escape sequences, so that is a backslash and a letter. Use vbCrLf, vbCr, vbLf or vbTab (e.g. ""a"" & vbCrLf & ""b"").");
        if (doubled) hints.Add(@"A string literal contains \\: in VB6 a backslash is literal, so write a path with single backslashes (""C:\Temp\x.txt""). Only a UNC path starts with \\.");
        return hints;
    }

    [GeneratedRegex("\"((?:[^\"\\r\\n]|\"\")*)\"")]
    private static partial Regex StringLiteral();

    /// <summary>
    /// When two texts differ only by letter case (VB re-cased identifiers project-wide, improvement notes 23): the
    /// identifiers spelled differently, as "disk -> ide" (from -> to). Null when they differ in more than case.
    /// </summary>
    public static IReadOnlyList<string>? CaseOnlyRespellings(string from, string to)
    {
        if (!string.Equals(from, to, StringComparison.OrdinalIgnoreCase)) return null;
        var toSpelling = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var id in Identifiers(to)) toSpelling.TryAdd(id, id);
        return Identifiers(from).Distinct(StringComparer.Ordinal)
            .Where(id => toSpelling.TryGetValue(id, out var t) && t != id).Select(id => $"{id} -> {toSpelling[id]}")
            .Distinct(StringComparer.Ordinal).ToList();
    }

    // Identifiers outside string literals and comments ("..." and ' to end of line; Rem is rare enough to ignore)
    private static IEnumerable<string> Identifiers(string text)
    {
        foreach (var raw in text.Split('\n'))
        {
            var line = new StringBuilder();
            bool inString = false;
            foreach (char ch in raw)
            {
                if (ch == '"') { inString = !inString; line.Append(' '); continue; }
                if (!inString && ch == '\'') break;
                line.Append(inString ? ' ' : ch);
            }
            foreach (Match m in IdentifierPattern().Matches(line.ToString())) yield return m.Value;
        }
    }

    [GeneratedRegex(@"[A-Za-z_][A-Za-z0-9_]*")]
    private static partial Regex IdentifierPattern();

    public static (string Text, int StartLine, int Count) TrimLeadingBlankLines(string procText, int startLine, int count)
    {
        var lines = procText.Split("\r\n");
        int skip = 0;
        while (skip < lines.Length - 1 && lines[skip].Trim().Length == 0) skip++;
        return (string.Join("\r\n", lines.Skip(skip)), startLine + skip, count - skip);
    }
}
