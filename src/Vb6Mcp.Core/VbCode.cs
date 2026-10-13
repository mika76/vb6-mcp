using System.Security.Cryptography;
using System.Text;

namespace Vb6Mcp.Core;

/// <summary>
/// Helpers for code text exchanged with the IDE: hashing for optimistic concurrency, validation of
/// text before it is written (AGENTS.md hard rules 1 and 10), and procedure-range trimming.
/// </summary>
public static class VbCode
{
    /// <summary>VB6's hard limit on physical line length.</summary>
    public const int MaxLineLength = 1023;
    /// <summary>VB6 allows at most 24 line continuations in one logical line.</summary>
    public const int MaxContinuations = 24;

    private static readonly Encoding s_strict1252 = Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

    static VbCode() => _ = Vb1252.Encoding;   // registers the code-pages provider before s_strict1252 is used

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
    public static (string Text, int StartLine, int Count) TrimLeadingBlankLines(string procText, int startLine, int count)
    {
        var lines = procText.Split("\r\n");
        int skip = 0;
        while (skip < lines.Length - 1 && lines[skip].Trim().Length == 0) skip++;
        return (string.Join("\r\n", lines.Skip(skip)), startLine + skip, count - skip);
    }
}
