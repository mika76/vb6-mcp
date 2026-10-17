using System.Text;
using System.Text.RegularExpressions;

namespace Vb6Mcp.Core;

/// <summary>One statement that names the identifier: first line (1-based, IDE numbering), last line, the statement with
/// its ` _` continuations joined, and whether it declares the identifier (Sub/Function/Property/Declare/Event).</summary>
public sealed record VbRef(int Line, int EndLine, string Statement, bool Declaration);

/// <summary>
/// Finds the statements that use an identifier in VB6 code (symbol_refs, PLAN.md C5). Improvement notes 4: the agent
/// needed every caller of Add_To_Audit across six projects with the arguments they pass, and grep missed calls
/// continued over several lines. So statements are joined across ` _` first, and matches inside string literals and
/// comments don't count. No parser beyond that (principle 3): a whole-word, case-insensitive match, as VB names are.
/// </summary>
public static class VbRefs
{
    static readonly Regex Ident = new(@"^[A-Za-z][A-Za-z0-9_]*$", RegexOptions.CultureInvariant);

    public static bool IsIdentifier(string name) => Ident.IsMatch(name);

    public static IReadOnlyList<VbRef> Find(string code, string name)
    {
        if (!IsIdentifier(name)) throw new ArgumentException("not a VB identifier: " + name);
        var word = new Regex($@"(?<![A-Za-z0-9_]){Regex.Escape(name)}(?![A-Za-z0-9_])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        var decl = new Regex($@"^\s*(?:(?:Public|Private|Friend|Global|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set)|Declare\s+(?:Sub|Function)|Event)\s+{Regex.Escape(name)}(?![A-Za-z0-9_])",
                             RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        var lines = code.Replace("\r\n", "\n").Split('\n');
        var found = new List<VbRef>();
        for (int i = 0; i < lines.Length; i++)
        {
            int first = i;
            var stmt = new StringBuilder();
            var masked = new StringBuilder();
            while (true)
            {
                var m = Mask(lines[i]);
                bool continued = m.TrimEnd().EndsWith(" _", StringComparison.Ordinal) || m.TrimEnd() == "_";
                if (continued && i + 1 < lines.Length)
                {
                    stmt.Append(lines[i].TrimEnd()[..^1].TrimEnd()).Append(' ');
                    masked.Append(m.TrimEnd()[..^1].TrimEnd()).Append(' ');
                    i++;
                    continue;
                }
                stmt.Append(lines[i].TrimEnd());
                masked.Append(m);
                break;
            }
            var text = masked.ToString();
            if (!word.IsMatch(text)) continue;
            var s = Regex.Replace(stmt.ToString().Trim(), @"\s{2,}", " ");
            found.Add(new VbRef(first + 1, i + 1, s, decl.IsMatch(text)));
        }
        return found;
    }

    /// <summary>The line with string literals blanked and any comment (' or Rem) removed, so only code is matched.</summary>
    static string Mask(string line)
    {
        var sb = new StringBuilder(line.Length);
        bool inString = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (inString)
            {
                if (c == '"') { if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append("  "); i++; continue; } inString = false; sb.Append('"'); }
                else sb.Append(' ');
                continue;
            }
            if (c == '"') { inString = true; sb.Append('"'); continue; }
            if (c == '\'') break;
            // Rem starts a comment as a statement (line start or after ':')
            if ((c is 'R' or 'r') && string.Compare(line, i, "Rem", 0, 3, StringComparison.OrdinalIgnoreCase) == 0
                && (i + 3 == line.Length || !char.IsLetterOrDigit(line[i + 3]) && line[i + 3] != '_')
                && sb.ToString().TrimEnd() is var before && (before.Length == 0 || before.EndsWith(':'))) break;
            sb.Append(c);
        }
        return sb.ToString();
    }
}
