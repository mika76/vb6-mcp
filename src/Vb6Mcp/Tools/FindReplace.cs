using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>One line find/replace changes: its number (1-based), the text before and after.</summary>
public sealed record FoundLine(int Line, string Old, string New);

/// <summary>
/// find/replace for code_replace_lines and file_replace_lines (PLAN.md A6, improvement notes 5: renaming m_comm.Output to
/// m_comm.OutputText on about 50 lines took about 50 calls, each needing the previous hash). The matches become one edit
/// list against the hash the agent read, so it is one call and the usual conflict checks hold. By default only code is
/// changed: text inside string literals and comments stays, since a rename of an identifier shouldn't touch them.
/// </summary>
public static class FindReplace
{
    public const int MaxPreview = 200;

    public const string Description = "Instead of line ranges: replace every occurrence of find with replace (whole words, any case, " +
        "code only: not in strings or comments, unless the options say so). Without apply=true it only previews the changed lines.";

    /// <summary>The lines that change. Throws on bad options.</summary>
    public static IReadOnlyList<FoundLine> Compute(string code, string find, string replace, bool wholeWord = true, bool matchCase = false, bool inStrings = false, bool inComments = false)
    {
        if (find.Length == 0) throw new ToolException("find is empty");
        if (find.Contains('\n') || replace.Contains('\n')) throw new ToolException("find and replace work within a line; use edits for multi-line changes");
        var pattern = Regex.Escape(find);
        if (wholeWord) pattern = (char.IsLetterOrDigit(find[0]) || find[0] == '_' ? @"(?<![A-Za-z0-9_])" : "") + pattern + (char.IsLetterOrDigit(find[^1]) || find[^1] == '_' ? @"(?![A-Za-z0-9_])" : "");
        var re = new Regex(pattern, (matchCase ? RegexOptions.None : RegexOptions.IgnoreCase) | RegexOptions.CultureInvariant);
        var lines = code.Replace("\r\n", "\n").Split('\n');
        var result = new List<FoundLine>();
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var kinds = Kinds(line);
            var sb = new StringBuilder();
            int last = 0;
            foreach (Match m in re.Matches(line))
            {
                bool ok = Enumerable.Range(m.Index, m.Length).All(p => kinds[p] switch { 'c' => true, 's' => inStrings, _ => inComments });
                if (!ok) continue;
                sb.Append(line, last, m.Index - last).Append(replace);
                last = m.Index + m.Length;
            }
            if (last == 0 && sb.Length == 0) continue;
            sb.Append(line, last, line.Length - last);
            if (sb.ToString() != line) result.Add(new FoundLine(i + 1, line, sb.ToString()));
        }
        return result;
    }

    /// <summary>Per character: c = code, s = inside a string literal (quotes included), r = comment (' or Rem to the end).</summary>
    static char[] Kinds(string line)
    {
        var k = new char[line.Length];
        bool inString = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (inString) { k[i] = 's'; if (c == '"') { if (i + 1 < line.Length && line[i + 1] == '"') { k[++i] = 's'; } else inString = false; } continue; }
            if (c == '"') { inString = true; k[i] = 's'; continue; }
            bool rem = (c is 'R' or 'r') && string.Compare(line, i, "Rem", 0, 3, StringComparison.OrdinalIgnoreCase) == 0
                       && (i + 3 == line.Length || !(char.IsLetterOrDigit(line[i + 3]) || line[i + 3] == '_'))
                       && line[..i].TrimEnd() is var before && (before.Length == 0 || before.EndsWith(':'));
            if (c == '\'' || rem) { for (int j = i; j < line.Length; j++) k[j] = 'r'; break; }
            k[i] = 'c';
        }
        return k;
    }

    /// <summary>The find/replace arguments of a call, or null when it has none.</summary>
    public static (string Find, string Replace, bool WholeWord, bool MatchCase, bool InStrings, bool InComments, bool Apply)? Args(JsonObject a)
    {
        if (a["find"] is null) return null;
        if (a["start_line"] is not null || a["edits"] is not null || a["new_text"] is not null)
            throw new ToolException("pass either find/replace or start_line/count/new_text or edits, not several");
        return (a.Str("find") ?? throw new ToolException("find must be text"), a.Str("replace") ?? throw new ToolException("missing required argument 'replace'"),
                a.Bool("whole_word") ?? true, a.Bool("match_case") ?? false, a.Bool("in_strings") ?? false, a.Bool("in_comments") ?? false, a.Bool("apply") == true);
    }

    /// <summary>Line edits for the changed lines (one per line, bottom-up), checked like any other edits.</summary>
    public static IReadOnlyList<LineEdit> Edits(IReadOnlyList<FoundLine> found)
    {
        if (found.Count > EditList.MaxEdits) throw new ToolException($"{found.Count} lines would change; at most {EditList.MaxEdits} per call: narrow find");
        return found.Select(f =>
        {
            try { return new LineEdit(f.Line, 1, Core.VbCode.PrepareForInsert(f.New)); }
            catch (ArgumentException ex) { throw new ToolException($"line {f.Line}: {ex.Message}"); }
        }).OrderByDescending(e => e.Start).ToList();
    }

    /// <summary>The preview result part: the changed lines (at most 200) and how many.</summary>
    public static object Preview(IReadOnlyList<FoundLine> found) => new
    {
        lines = found.Count, changes = found.Take(MaxPreview).Select(f => new { line = f.Line, old = f.Old, @new = f.New }),
        truncated = found.Count > MaxPreview ? true : (bool?)null,
        note = found.Count == 0 ? "no matches (whole words, code only by default)" : "nothing changed yet: call again with the same arguments and apply=true",
    };

    /// <summary>Adds the find/replace parameters to an edit tool's schema.</summary>
    public static JsonObject WithSchema(JsonObject schema)
    {
        var p = schema["properties"]!.AsObject();
        p["find"] = new JsonObject { ["type"] = "string", ["description"] = Description };
        p["replace"] = new JsonObject { ["type"] = "string", ["description"] = "Replacement for find." };
        p["whole_word"] = new JsonObject { ["type"] = "boolean", ["description"] = "Match whole words only (default true)." };
        p["match_case"] = new JsonObject { ["type"] = "boolean", ["description"] = "Match case (default false: VB names ignore case)." };
        p["in_strings"] = new JsonObject { ["type"] = "boolean", ["description"] = "Also replace inside string literals (default false)." };
        p["in_comments"] = new JsonObject { ["type"] = "boolean", ["description"] = "Also replace inside comments (default false)." };
        p["apply"] = new JsonObject { ["type"] = "boolean", ["description"] = "Make the find/replace changes (default false: preview only)." };
        return schema;
    }
}
