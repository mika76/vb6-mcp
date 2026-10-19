namespace Vb6Mcp.Core.Files;

/// <summary>
/// Matching a search text against names the way people mean it: at the start of a word. Improvement notes 44:
/// com_search with a company's two-letter prefix matched every MSXML library and filled the 50-entry cap with substrings.
/// Words are split at anything that isn't a letter or digit and where a lower-case letter is followed by an
/// upper-case one, so "XmText.ocx" has the words Xm, Text, ocx and "MSXML2" has only MSXML2.
/// </summary>
public static class WordMatch
{
    /// <summary>How well a search text matches a name: anywhere, at the start of a word, or a whole word.</summary>
    public enum Rank { None, Substring, WordPrefix, Word }

    /// <summary>The best match of <paramref name="text"/> in <paramref name="name"/> (case-insensitive).</summary>
    public static Rank Of(string text, string? name)
    {
        if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(text) || !name.Contains(text, StringComparison.OrdinalIgnoreCase)) return Rank.None;
        var best = Rank.Substring;
        foreach (var w in Words(name))
        {
            if (w.Equals(text, StringComparison.OrdinalIgnoreCase)) return Rank.Word;
            if (w.StartsWith(text, StringComparison.OrdinalIgnoreCase)) best = Rank.WordPrefix;
        }
        // a text that spans words ("Common Controls", "MSComctlLib.List") counts when it starts at a word
        if (best == Rank.Substring && !text.All(char.IsLetterOrDigit))
            for (int i = name.IndexOf(text, StringComparison.OrdinalIgnoreCase); i >= 0; i = name.IndexOf(text, i + 1, StringComparison.OrdinalIgnoreCase))
                if (StartsWord(name, i)) return Rank.WordPrefix;
        return best;
    }

    /// <summary>The best rank over several names.</summary>
    public static Rank Of(string text, params string?[] names) => names.Select(n => Of(text, n)).DefaultIfEmpty(Rank.None).Max();

    private static bool StartsWord(string s, int i)
        => i == 0 || !char.IsLetterOrDigit(s[i - 1]) || (char.IsLower(s[i - 1]) && char.IsUpper(s[i]));

    private static IEnumerable<string> Words(string s)
    {
        int start = -1;
        for (int i = 0; i <= s.Length; i++)
        {
            bool alnum = i < s.Length && char.IsLetterOrDigit(s[i]);
            if (alnum && start >= 0 && char.IsLower(s[i - 1]) && char.IsUpper(s[i])) { yield return s[start..i]; start = i; }
            else if (alnum && start < 0) start = i;
            else if (!alnum && start >= 0) { yield return s[start..i]; start = -1; }
        }
    }
}
