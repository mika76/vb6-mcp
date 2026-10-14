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

    /// <summary>
    /// The file's bytes with code lines [<paramref name="startLine"/>, +<paramref name="count"/>) (1-based, as the
    /// IDE shows the code) replaced by <paramref name="newText"/> (CR LF separated; count 0 inserts). Everything
    /// outside the range stays byte-exact: header, designer block, other lines and the final line break
    /// (hard rule 2). Hidden inline Attribute lines inside the range (e.g. "Attribute Value.VB_UserMemId = 0"
    /// after a property's first line) are put back after the same code line in the new text; if that line
    /// isn't there any more the edit is refused, because the attribute would be lost or misplaced.
    /// The caller validates the new text (1252, line limits: VbCode.PrepareForInsert).
    /// </summary>
    public byte[] ReplaceCode(int startLine, int count, string newText)
    {
        var text = Vb1252.Encoding.GetString(Bytes);
        bool finalBreak = text.EndsWith('\n');
        var lines = text.Replace("\r\n", "\n").Split('\n').ToList();
        if (finalBreak) lines.RemoveAt(lines.Count - 1);

        // physical index of each code line (header and inline attributes skipped)
        var codeAt = new List<int>();
        for (int i = Header.Count; i < lines.Count; i++) if (!IsAttribute(lines[i])) codeAt.Add(i);
        int n = codeAt.Count;
        if (startLine < 1 || startLine > n + 1) throw new ArgumentException($"start line {startLine} is outside the code (1..{n + 1})");
        if (count < 0 || startLine - 1 + count > n) throw new ArgumentException($"lines {startLine}..{startLine + count - 1} go past the end of the code ({n} lines)");

        var newLines = newText.Length == 0 ? new List<string>() : newText.Replace("\r\n", "\n").Split('\n').ToList();
        int first = startLine - 1;
        int physStart = first < n ? codeAt[first] : lines.Count;
        int physEnd = count == 0 ? physStart : codeAt[first + count - 1] + 1;    // exclusive
        // inline attributes between the first and last replaced code line
        var output = new List<string>(newLines);
        for (int i = physStart; i < physEnd; i++)
        {
            if (!IsAttribute(lines[i])) continue;
            int prevCode = i - 1;
            while (prevCode >= physStart && IsAttribute(lines[prevCode])) prevCode--;
            if (prevCode < physStart) { output.Insert(0, lines[i]); continue; }   // belongs before the range
            int at = output.IndexOf(lines[prevCode]);
            if (at < 0) throw new ArgumentException($"the range contains the hidden line '{lines[i]}' after '{lines[prevCode].Trim()}', which the new text doesn't keep; edit it in the IDE or keep that line");
            output.Insert(at + 1, lines[i]);
        }
        lines.RemoveRange(physStart, physEnd - physStart);
        lines.InsertRange(physStart, output);
        var result = string.Join("\r\n", lines) + (finalBreak || lines.Count == 0 ? "\r\n" : "");
        return Vb1252.Encoding.GetBytes(result);
    }

    private static bool IsAttribute(string line) => line.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase);

    private static string? ParseVbName(string line)
    {
        const string prefix = "Attribute VB_Name = \"";
        return line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && line.EndsWith('"') ? line[prefix.Length..^1] : null;
    }
}
