using System.Text.Json.Nodes;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>A control's rectangle in its container's units (twips unless the container's ScaleMode says otherwise).</summary>
public readonly record struct LayoutBox(double Left, double Top, double Width, double Height);

/// <summary>
/// What form_layout arranges: a grid of control names. One row of names is a row of buttons; one name per row is a
/// column. The same name in adjacent cells of a row spans those columns; "" is an empty cell.
/// </summary>
public sealed class LayoutSpec
{
    public required IReadOnlyList<IReadOnlyList<string>> Rows { get; init; }
    /// <summary>Top-left of the block; null = placed by <see cref="Align"/> / <see cref="Anchor"/> and the margin.</summary>
    public double? Left { get; init; }
    public double? Top { get; init; }
    public double GapX { get; init; } = 120;
    public double GapY { get; init; } = 90;
    /// <summary>Distance from the container's edges when the block is placed by Align/Anchor or the container is sized.</summary>
    public double Margin { get; init; } = 120;
    /// <summary>Fixed column widths / row heights; a null entry (or a missing one) is worked out from the controls.</summary>
    public IReadOnlyList<double?> ColumnWidths { get; init; } = [];
    public IReadOnlyList<double?> RowHeights { get; init; } = [];
    /// <summary>Give every control the width of its cell (all text boxes of a column the same width).</summary>
    public bool Fill { get; init; } = true;
    /// <summary>Within a row: top, center (default: a label sits level with the text box beside it) or bottom.</summary>
    public string VAlign { get; init; } = "center";
    /// <summary>The block across the container: left (default), center or right.</summary>
    public string Align { get; init; } = "left";
    /// <summary>The block in the container: top (default) or bottom.</summary>
    public string Anchor { get; init; } = "top";
}

/// <summary>Where each control goes, the block's rectangle, and warnings (the block runs past the container).</summary>
public sealed record LayoutResult(IReadOnlyDictionary<string, LayoutBox> Boxes, LayoutBox Block, IReadOnlyList<string> Warnings);

/// <summary>
/// form_layout's arithmetic, separate from the IDE so it is unit-tested. Benchmark T9–T12 redesigns placed 39–60
/// controls one at a time with twips worked out by hand, which is where overlapping controls came from (PLAN.md E1).
/// Column widths and row heights default to the largest control in them, so the agent names an arrangement and the
/// sizes it already gave the controls (or VB's defaults) decide the rest.
/// </summary>
public static class FormLayout
{
    public const int MaxCells = 400;

    /// <summary>
    /// Positions for the spec. <paramref name="current"/> holds the controls' rectangles now (keys case-insensitive,
    /// Name or Name(index)); <paramref name="inner"/> is the container's inside size (null when unknown: then only
    /// left/top placement works); <paramref name="topInset"/> is space taken at the top (a Frame's caption).
    /// </summary>
    public static LayoutResult Compute(LayoutSpec spec, IReadOnlyDictionary<string, LayoutBox> current, (double Width, double Height)? inner, double topInset = 0)
    {
        var rows = spec.Rows;
        if (rows.Count == 0 || rows.All(r => r.All(c => string.IsNullOrWhiteSpace(c)))) throw new ToolException("rows is empty: give a list of rows, each a list of control names");
        int cols = rows.Max(r => r.Count);
        if (rows.Count * cols > MaxCells) throw new ToolException($"at most {MaxCells} cells ({rows.Count} rows x {cols} columns given)");
        if (spec.VAlign is not ("top" or "center" or "bottom")) throw new ToolException("valign must be top, center or bottom");
        if (spec.Align is not ("left" or "center" or "right")) throw new ToolException("align must be left, center or right");
        if (spec.Anchor is not ("top" or "bottom")) throw new ToolException("anchor must be top or bottom");
        if (spec.GapX < 0 || spec.GapY < 0 || spec.Margin < 0) throw new ToolException("gaps and margin can't be negative");

        // each control: one row, one run of adjacent columns
        var cells = new Dictionary<string, (int Row, int Col, int Span)>(StringComparer.OrdinalIgnoreCase);
        for (int r = 0; r < rows.Count; r++)
            for (int c = 0; c < rows[r].Count; c++)
            {
                var name = rows[r][c]?.Trim() ?? "";
                if (name.Length == 0) continue;
                if (!current.ContainsKey(name)) throw new ToolException($"{name} is not a control in this container (form_controls lists them)");
                if (cells.TryGetValue(name, out var at))
                {
                    if (at.Row == r && at.Col + at.Span == c) { cells[name] = at with { Span = at.Span + 1 }; continue; }
                    throw new ToolException($"{name} appears twice; to span columns, repeat it in adjacent cells of one row");
                }
                cells[name] = (r, c, 1);
            }

        double Given(IReadOnlyList<double?> list, int i) => i < list.Count && list[i] is { } v ? v : double.NaN;
        var colW = new double[cols];
        for (int c = 0; c < cols; c++)
        {
            colW[c] = Given(spec.ColumnWidths, c);
            if (double.IsNaN(colW[c]))
                colW[c] = cells.Where(kv => kv.Value.Col == c && kv.Value.Span == 1).Select(kv => current[kv.Key].Width).DefaultIfEmpty(0).Max();
        }
        // a spanning control wider than its columns widens the last of them (only auto columns)
        foreach (var (name, (_, col, span)) in cells.Where(kv => kv.Value.Span > 1))
        {
            double have = colW.Skip(col).Take(span).Sum() + spec.GapX * (span - 1), need = current[name].Width;
            int last = col + span - 1;
            if (need > have && double.IsNaN(Given(spec.ColumnWidths, last)) && !spec.Fill) colW[last] += need - have;
        }
        var rowH = new double[rows.Count];
        for (int r = 0; r < rows.Count; r++)
        {
            rowH[r] = Given(spec.RowHeights, r);
            if (double.IsNaN(rowH[r])) rowH[r] = cells.Where(kv => kv.Value.Row == r).Select(kv => current[kv.Key].Height).DefaultIfEmpty(0).Max();
        }

        double blockW = colW.Sum() + spec.GapX * (cols - 1), blockH = rowH.Sum() + spec.GapY * (rows.Count - 1);
        double NeedInner(string what) => inner is null ? throw new ToolException($"{what} needs the container's size, which isn't known; give left/top instead") : 0;
        double left = spec.Left ?? spec.Align switch
        {
            "right" => NeedInner("align right") + inner!.Value.Width - spec.Margin - blockW,
            "center" => NeedInner("align center") + (inner!.Value.Width - blockW) / 2,
            _ => spec.Margin,
        };
        double top = spec.Top ?? (spec.Anchor == "bottom" ? NeedInner("anchor bottom") + inner!.Value.Height - spec.Margin - blockH : topInset + spec.Margin);

        var warnings = new List<string>();
        if (left < 0 || top < topInset) warnings.Add($"the block ({R(blockW)} x {R(blockH)}) doesn't fit: it starts at {R(left)}, {R(top)}");
        if (inner is { } box)
        {
            if (left + blockW > box.Width) warnings.Add($"the block runs {R(left + blockW - box.Width)} past the container's right edge (inside width {R(box.Width)}): use size_container, narrower column_widths or fewer columns");
            if (top + blockH > box.Height) warnings.Add($"the block runs {R(top + blockH - box.Height)} past the container's bottom edge (inside height {R(box.Height)}): use size_container or fewer rows");
        }

        var boxes = new Dictionary<string, LayoutBox>(StringComparer.OrdinalIgnoreCase);
        foreach (var (name, (row, col, span)) in cells)
        {
            var now = current[name];
            double x = left + colW.Take(col).Sum() + spec.GapX * col;
            double cellW = colW.Skip(col).Take(span).Sum() + spec.GapX * (span - 1);
            double y = top + rowH.Take(row).Sum() + spec.GapY * row;
            double h = now.Height;
            y += spec.VAlign switch { "center" => (rowH[row] - h) / 2, "bottom" => rowH[row] - h, _ => 0 };
            boxes[name] = new LayoutBox(R(x), R(y), R(spec.Fill ? cellW : now.Width), R(h));
        }
        return new LayoutResult(boxes, new LayoutBox(R(left), R(top), R(blockW), R(blockH)), warnings);
    }

    /// <summary>The container's inside size that fits the block plus the margin on the right and bottom.</summary>
    public static (double Width, double Height) FitInner(LayoutResult r, double margin) => (r.Block.Left + r.Block.Width + margin, r.Block.Top + r.Block.Height + margin);

    /// <summary>Reads the rows argument: a list of lists of names, or a list of strings (each one a row of a single control).</summary>
    public static IReadOnlyList<IReadOnlyList<string>> ParseRows(JsonNode? node)
    {
        var arr = node switch
        {
            JsonArray a => a,
            JsonValue v when v.TryGetValue(out string? text) => JsonNode.Parse(text) as JsonArray,   // some clients send JSON as a string
            _ => null,
        } ?? throw new ToolException("rows must be a list of rows, each a list of control names: [[\"lblName\", \"txtName\"], [\"lblCity\", \"txtCity\"]]");
        return arr.Select((row, i) => row switch
        {
            JsonArray cells => (IReadOnlyList<string>)cells.Select(c => c is JsonValue cv && cv.TryGetValue(out string? s) ? s : c is null ? "" : throw new ToolException($"rows[{i}]: cells must be control names (\"\" for an empty cell)")).ToList(),
            JsonValue v when v.TryGetValue(out string? one) => [one],
            _ => throw new ToolException($"rows[{i}] must be a list of control names"),
        }).ToList();
    }

    /// <summary>Reads column_widths / row_heights: numbers, with null for "work it out".</summary>
    public static IReadOnlyList<double?> ParseSizes(JsonNode? node, string what) => node switch
    {
        null => [],
        JsonArray a => a.Select(n => n is null ? (double?)null : n is JsonValue v && v.TryGetValue(out double d) && d >= 0 ? d : throw new ToolException($"{what}: numbers (or null for automatic)")).ToList(),
        _ => throw new ToolException($"{what} must be a list of numbers"),
    };

    private static double R(double v) => Math.Round(v, MidpointRounding.AwayFromZero);
}
