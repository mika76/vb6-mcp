using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>One line-range edit: replace <see cref="Count"/> lines from <see cref="Start"/> (0 = insert) with prepared text.</summary>
public sealed record LineEdit(int Start, int Count, string Text)
{
    /// <summary>Lines the prepared text adds (0 for a deletion).</summary>
    public int NewLines => Text.Length == 0 ? 0 : Text.Split("\r\n").Length;
}

/// <summary>
/// The edits of code_replace_lines / file_replace_lines: either the single start_line/count/new_text, or an `edits` list
/// (improvement notes item 5: a 50-line mechanical rename took 50 calls, each needing the previous hash). All ranges are
/// numbered against the same original text, checked for overlaps up front, and applied bottom-up, so earlier line numbers
/// stay valid while later ones are edited.
/// </summary>
public static class EditList
{
    public const int MaxEdits = 500;

    /// <summary>
    /// Appended to every parameter that takes VB6 code (owner, 2026-10-17: CR LF and backslashes needed fixing by hand).
    /// The tools convert line breaks themselves; escapes in string literals are the agent's to get right.
    /// </summary>
    public const string CodeRule = " VB6 (Windows-only): line breaks become CR LF; string literals have no escapes, so use vbCrLf/vbTab instead of \\n/\\t and single backslashes in paths (\"C:\\Temp\").";

    /// <summary>Description of the edits parameter, shared by code_replace_lines and file_replace_lines.</summary>
    public const string Description = "Instead of start_line/count/new_text: a list of {start_line, count, new_text}, all numbered " +
        "against the text you read (same expected_hash), not overlapping; applied bottom-up, one new hash back. At most 500.";

    /// <summary>Gives the edits parameter its object item schema (Schema.Object only makes string lists).</summary>
    public static JsonObject WithSchema(JsonObject schema)
    {
        schema["properties"]!["edits"]!["items"] = new JsonObject
        {
            ["type"] = "object",
            ["properties"] = new JsonObject
            {
                ["start_line"] = new JsonObject { ["type"] = "integer" },
                ["count"] = new JsonObject { ["type"] = "integer" },
                ["new_text"] = new JsonObject { ["type"] = "string" },
            },
            ["required"] = new JsonArray("start_line", "count", "new_text"),
        };
        return schema;
    }

    public static IReadOnlyList<LineEdit> From(JsonObject a)
    {
        var list = a["edits"] switch
        {
            null => null,
            JsonArray arr => arr,
            JsonValue v when v.TryGetValue(out string? text) => JsonNode.Parse(text) as JsonArray,   // some clients send the list as a JSON string
            _ => throw new ToolException("edits must be a list of {start_line, count, new_text} objects"),
        };
        if (list is null)
        {
            int start = a.Int("start_line") ?? throw new ToolException("missing required argument 'start_line' (or pass edits)");
            int count = a.Int("count") ?? throw new ToolException("missing required argument 'count'");
            return [new LineEdit(start, count, Prepare(a.Str("new_text") ?? throw new ToolException("missing required argument 'new_text'"), null))];
        }
        if (a["start_line"] is not null || a["new_text"] is not null) throw new ToolException("pass either edits or start_line/count/new_text, not both");
        if (list.Count == 0) throw new ToolException("edits is empty");
        if (list.Count > MaxEdits) throw new ToolException($"at most {MaxEdits} edits per call ({list.Count} given)");
        var edits = list.Select((n, i) =>
        {
            if (n is not JsonObject o || o.Int("start_line") is not int s || o.Int("count") is not int c || o.Str("new_text") is not { } t)
                throw new ToolException($"edits[{i}] needs start_line, count and new_text");
            if (s < 1 || c < 0) throw new ToolException($"edits[{i}]: start_line must be 1 or more and count 0 or more");
            return new LineEdit(s, c, Prepare(t, i));
        }).ToList();
        Validate(edits);
        return edits.OrderByDescending(e => e.Start).ThenByDescending(e => e.Count).ToList();
    }

    /// <summary>
    /// Refuses ranges that overlap or touch the same start: their meaning would depend on the order (insert before or after
    /// the replaced lines?), so the agent must merge them.
    /// </summary>
    public static void Validate(IReadOnlyList<LineEdit> edits)
    {
        var sorted = edits.OrderBy(e => e.Start).ThenBy(e => e.Count).ToList();
        for (int i = 1; i < sorted.Count; i++)
        {
            var (p, c) = (sorted[i - 1], sorted[i]);
            if (c.Start < p.Start + Math.Max(p.Count, 1))
                throw new ToolException($"edits overlap: lines {p.Start}-{p.Start + Math.Max(p.Count, 1) - 1} and {c.Start}-{c.Start + Math.Max(c.Count, 1) - 1}; merge them into one edit");
        }
    }

    private static string Prepare(string text, int? index)
    {
        try { return VbCode.PrepareForInsert(text); }
        catch (ArgumentException ex) { throw new ToolException(index is null ? ex.Message : $"edits[{index}]: {ex.Message}"); }
    }
}
