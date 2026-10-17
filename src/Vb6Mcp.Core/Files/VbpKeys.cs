namespace Vb6Mcp.Core.Files;

/// <summary>
/// Sets project-section keys of a .vbp on disk, byte-exactly (Windows-1252, CR LF, every other line untouched). For the
/// settings VBIDE doesn't expose: conditional compilation constants (CondComp) and the command line (Command32), among
/// others (improvement notes 2, 7, 22: the agent edited CondComp by script). Lines that define what the project is
/// made of or depends on are refused (hard rules 2 and 13): members, references, controls, and binary compatibility.
/// </summary>
public static class VbpKeys
{
    /// <summary>Keys this helper never changes: they have their own tools, or carry GUIDs and compatibility settings.</summary>
    public static readonly string[] Forbidden =
        [.. VbpFile.MemberKeys, "Reference", "Object", "ResFile32", "CompatibleMode", "CompatibleEXE32", "VersionCompatible32", "Type"];

    /// <summary>Keys whose value VB writes in quotes.</summary>
    static readonly string[] Quoted = ["CondComp", "Command32", "Name", "Title", "ExeName32", "HelpFile", "Description", "Path32", "VersionComments",
                                       "VersionCompanyName", "VersionFileDescription", "VersionLegalCopyright", "VersionLegalTrademarks", "VersionProductName"];

    /// <summary>The .vbp text with <paramref name="key"/> set to <paramref name="value"/> (added to the project section if missing).</summary>
    public static string Set(string text, string key, string value)
    {
        if (Forbidden.Any(k => k.Equals(key, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"{key} isn't changed by this tool: members, references, controls, the type and binary compatibility have their own tools (hard rules 2 and 13)");
        if (key.Length == 0 || !key.All(c => char.IsAsciiLetterOrDigit(c) || c == '_')) throw new ArgumentException($"'{key}' isn't a .vbp key");
        if (value.Contains('\r') || value.Contains('\n') || value.Contains('"')) throw new ArgumentException("the value can't contain line breaks or double quotes");
        if (value.Any(c => c > 0xFF)) throw new ArgumentException("the value must be Windows-1252 text");

        bool crlf = text.Contains("\r\n");
        string nl = crlf ? "\r\n" : "\n";
        var lines = text.Split(nl).ToList();
        // the project section ends at the first [section] line (e.g. [MS Transaction Server])
        int end = lines.FindIndex(l => l.StartsWith('['));
        if (end < 0) end = lines.Count;
        for (int i = 0; i < end; i++)
        {
            int eq = lines[i].IndexOf('=');
            if (eq <= 0 || !lines[i][..eq].Equals(key, StringComparison.OrdinalIgnoreCase)) continue;
            bool quoted = lines[i].Length > eq + 1 && lines[i][eq + 1] == '"';
            lines[i] = lines[i][..eq] + "=" + (quoted ? $"\"{value}\"" : value);
            return string.Join(nl, lines);
        }
        // new key: before the trailing blank lines / sections of the project section
        int at = end;
        while (at > 0 && lines[at - 1].Length == 0) at--;
        lines.Insert(at, key + "=" + (Quoted.Contains(key, StringComparer.OrdinalIgnoreCase) ? $"\"{value}\"" : value));
        return string.Join(nl, lines);
    }

    /// <summary>A key's value from the project section (quotes removed), or null.</summary>
    public static string? Get(string text, string key)
    {
        foreach (var line in text.Replace("\r\n", "\n").Split('\n'))
        {
            if (line.StartsWith('[')) break;
            int eq = line.IndexOf('=');
            if (eq > 0 && line[..eq].Equals(key, StringComparison.OrdinalIgnoreCase)) return line[(eq + 1)..].Trim('"');
        }
        return null;
    }
}
