namespace Vb6Mcp.Core.Files;

/// <summary>Which kind of VB6 source file, by extension.</summary>
public enum VbFileKind { Module, Class, Form, UserControl, PropertyPage, Designer, Other }

/// <summary>
/// A VB6 source file split the way the IDE sees it. The IDE's code module shows only <see cref="Code"/>;
/// it hides the header (VERSION / Object= / Begin..End designer block / class BEGIN..END block)
/// and every "Attribute ..." line, including the ones inside procedures
/// (e.g. "Attribute Value.VB_UserMemId = 0"). Keeping these apart lets tools read code like the IDE
/// does and, later, write code back without touching what the IDE hides (AGENTS.md hard rule 2).
/// </summary>
public sealed class VbSourceFile
{
    /// <summary>Raw file bytes, kept for byte-exact round trips.</summary>
    public required byte[] Bytes { get; init; }
    public required string Path { get; init; }
    public required VbFileKind Kind { get; init; }
    /// <summary>Lines before the code (header/designer block and the leading Attribute lines).</summary>
    public required IReadOnlyList<string> Header { get; init; }
    /// <summary>The code module text as the IDE shows it (CR LF separated, no Attribute lines).</summary>
    public required string Code { get; init; }
    /// <summary>Attribute lines that appeared inside the code, with the 1-based code line they follow.</summary>
    public required IReadOnlyList<(int AfterCodeLine, string Line)> InlineAttributes { get; init; }
    /// <summary>VB_Name from the header, i.e. the component name.</summary>
    public string? VbName { get; init; }

    public int CountOfLines => Code.Length == 0 ? 0 : Code.Split("\r\n").Length;

    public static VbFileKind KindOf(string path) => System.IO.Path.GetExtension(path).ToLowerInvariant() switch
    {
        ".bas" => VbFileKind.Module,
        ".cls" => VbFileKind.Class,
        ".frm" => VbFileKind.Form,
        ".ctl" => VbFileKind.UserControl,
        ".pag" => VbFileKind.PropertyPage,
        ".dsr" => VbFileKind.Designer,
        _ => VbFileKind.Other,
    };

    public static VbSourceFile Load(string path) => Parse(path, File.ReadAllBytes(path));

    /// <summary>Splits file bytes into header, code and inline attributes.</summary>
    public static VbSourceFile Parse(string path, byte[] bytes)
    {
        var text = Vb1252.Encoding.GetString(bytes);
        // Files are CR LF; tolerate LF-only files when reading (they'd be broken for VB6 anyway).
        var lines = text.Replace("\r\n", "\n").Split('\n').ToList();
        if (lines.Count > 0 && lines[^1].Length == 0) lines.RemoveAt(lines.Count - 1);   // final line break

        int i = 0;
        // Header blocks: "VERSION x" + optional Object= lines + one Begin...End (forms) or BEGIN...END (classes)
        if (i < lines.Count && lines[i].StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase))
        {
            i++;
            while (i < lines.Count && lines[i].StartsWith("Object", StringComparison.OrdinalIgnoreCase)) i++;
            if (i < lines.Count && lines[i].TrimStart().StartsWith("Begin", StringComparison.OrdinalIgnoreCase))
            {
                int depth = 0;
                for (; i < lines.Count; i++)
                {
                    var t = lines[i].Trim();
                    if (t.StartsWith("Begin", StringComparison.OrdinalIgnoreCase) || t.StartsWith("BeginProperty", StringComparison.OrdinalIgnoreCase)) depth++;
                    else if (t.Equals("End", StringComparison.OrdinalIgnoreCase) || t.StartsWith("EndProperty", StringComparison.OrdinalIgnoreCase)) depth--;
                    if (depth == 0) { i++; break; }
                }
            }
        }
        // Leading Attribute lines (VB_Name, VB_GlobalNameSpace, ...)
        string? vbName = null;
        while (i < lines.Count && IsAttribute(lines[i]))
        {
            vbName ??= ParseVbName(lines[i]);
            i++;
        }
        var header = lines.Take(i).ToList();

        var code = new List<string>();
        var inline = new List<(int, string)>();
        for (; i < lines.Count; i++)
        {
            if (IsAttribute(lines[i])) inline.Add((code.Count, lines[i]));
            else code.Add(lines[i]);
        }
        vbName ??= header.Select(ParseVbName).FirstOrDefault(n => n != null);
        return new VbSourceFile
        {
            Bytes = bytes, Path = path, Kind = KindOf(path), Header = header,
            Code = string.Join("\r\n", code), InlineAttributes = inline, VbName = vbName,
        };
    }

    private static bool IsAttribute(string line) => line.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase);

    private static string? ParseVbName(string line)
    {
        const string prefix = "Attribute VB_Name = \"";
        return line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && line.EndsWith('"') ? line[prefix.Length..^1] : null;
    }
}
