using Vb6Mcp.Core;

namespace Vb6Mcp.Tools;

/// <summary>A page of a module's code: the text returned, where it starts, and where to continue if it was cut.</summary>
public sealed record CodePage(string Text, int StartLine, int Lines, bool Truncated, int? NextStartLine);

/// <summary>
/// Caps what code_read and file_read return. Without a range they used to return the whole module, and a 185 KB class
/// came back as one JSON line the client couldn't take (improvement notes item 13). A default read now stops at
/// <see cref="MaxLines"/> lines or <see cref="MaxChars"/> characters, whichever comes first, and says where to continue.
/// An explicit start_line/count is honoured as asked, up to the same character cap. The hash always covers the whole text.
/// </summary>
public static class Paging
{
    public const int MaxLines = 2000;
    public const int MaxChars = 100_000;

    public static CodePage Page(string code, int? startLine, int? count)
    {
        var all = code.Length == 0 ? [] : code.Split("\r\n");
        int start = Math.Max(1, startLine ?? 1);
        int want = count ?? (startLine is null ? MaxLines : int.MaxValue);
        int end = (int)Math.Min((long)start - 1 + want, all.Length);   // exclusive, 0-based
        var lines = new List<string>();
        int chars = 0;
        for (int i = start - 1; i < end; i++)
        {
            if (lines.Count > 0 && chars + all[i].Length + 2 > MaxChars) break;   // always return at least one line
            lines.Add(all[i]);
            chars += all[i].Length + 2;
        }
        int last = start - 1 + lines.Count;            // last line returned, 1-based
        bool truncated = last < all.Length && (count is null || last < start - 1 + want);
        return new CodePage(string.Join("\r\n", lines), start, lines.Count, truncated, truncated ? last + 1 : null);
    }

    /// <summary>The note for a cut page, telling the agent how to read on.</summary>
    public static string? Note(CodePage p, int total)
        => p.Truncated ? $"showing lines {p.StartLine}-{p.StartLine + p.Lines - 1} of {total}; read on with start_line={p.NextStartLine} (the hash covers the whole text)" : null;
}
