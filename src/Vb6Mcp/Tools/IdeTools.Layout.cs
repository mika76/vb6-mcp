using System.Globalization;
using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>form_layout (PLAN.md E1): arrange controls by naming a grid; the server works out the twips.</summary>
public sealed partial class IdeTools
{
    // A Frame's children are placed from its outer top-left, and its caption takes the top. 180 twips clears an 8 pt caption
    // (the default font); with the margin the first row starts at 300.
    const double FrameCaptionInset = 180;

    private Tool FormLayoutTool() => new("form_layout",
        "Arrange controls of one container (the form, or a Frame/PictureBox) as a grid, instead of working out Left/Top by hand. rows is a " +
        "list of rows, each a list of control names: [[\"lblName\",\"txtName\"],[\"lblCity\",\"txtCity\"]] makes a two-column form; one row " +
        "[[\"cmdOK\",\"cmdCancel\"]] with align=right and anchor=bottom puts buttons in the bottom-right corner. \"\" is an empty cell; the same " +
        "name in adjacent cells of a row spans those columns. Column widths and row heights default to the largest control in them (size the " +
        "controls first, or give column_widths/row_heights; null = automatic); with fill (default true) every control takes its cell's width. " +
        "Controls are centred vertically in their row (labels level with the text box beside them). size_container fits the frame or form " +
        "around the block. dry_run=true returns the positions without changing anything. Units are the container's (twips). The result lists " +
        "overlaps with controls outside the grid; check with form_screenshot. Unsaved; save with project_save.",
        LayoutSchema(Schema.Object(("component", "string", ComponentDesc, true),
            ("rows", "array", "List of rows, each a list of control names (Name or Name(index)); \"\" = empty cell.", true),
            ("container", "string", "Frame or PictureBox whose controls these are (default: the form itself).", false),
            ("left", "number", "Left of the block (default: by align and margin).", false), ("top", "number", "Top of the block (default: by anchor and margin).", false),
            ("gap_x", "number", "Space between columns (default 120).", false), ("gap_y", "number", "Space between rows (default 90).", false),
            ("margin", "number", "Distance from the container's edges (default 120).", false),
            ("column_widths", "array", "Width per column; null = widest control in it.", false), ("row_heights", "array", "Height per row; null = tallest control in it.", false),
            ("fill", "boolean", "Every control takes its cell's width (default true).", false),
            ("valign", "string", "top, center (default) or bottom within the row.", false),
            ("align", "string", "left (default), center or right in the container.", false),
            ("anchor", "string", "top (default) or bottom of the container.", false),
            ("size_container", "boolean", "Resize the container to fit the block plus margin (needs align left and anchor top).", false),
            ("dry_run", "boolean", "Return the positions only.", false),
            ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false))),
        (a, _) => FormLayoutAsync(a));

    private static JsonObject LayoutSchema(JsonObject schema)
    {
        schema["properties"]!["rows"]!["items"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" } };
        foreach (var k in new[] { "column_widths", "row_heights" })
            schema["properties"]![k]!["items"] = new JsonObject { ["type"] = new JsonArray("number", "null") };
        return schema;
    }

    private async Task<ToolResult> FormLayoutAsync(JsonObject a)
    {
        double? Dbl(string k) => a[k] is JsonValue v ? (v.TryGetValue(out double d) ? d : v.TryGetValue(out string? s) && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var p) ? p : throw new ToolException($"{k} must be a number")) : null;
        // everything is checked before the IDE is asked, so a malformed call changes nothing
        var spec = new LayoutSpec
        {
            Rows = FormLayout.ParseRows(a["rows"]),
            Left = Dbl("left"), Top = Dbl("top"),
            GapX = Dbl("gap_x") ?? 120, GapY = Dbl("gap_y") ?? 90, Margin = Dbl("margin") ?? 120,
            ColumnWidths = FormLayout.ParseSizes(a["column_widths"], "column_widths"), RowHeights = FormLayout.ParseSizes(a["row_heights"], "row_heights"),
            Fill = a.Bool("fill") ?? true,
            VAlign = a.Str("valign")?.ToLowerInvariant() ?? "center", Align = a.Str("align")?.ToLowerInvariant() ?? "left", Anchor = a.Str("anchor")?.ToLowerInvariant() ?? "top",
        };
        bool sizeContainer = a.Bool("size_container") == true, dryRun = a.Bool("dry_run") == true;
        if (sizeContainer && ((spec.Align != "left" && spec.Left is null) || (spec.Anchor != "top" && spec.Top is null)))
            throw new ToolException("size_container fits the container around a block placed from the top-left: use align left and anchor top (or give left/top)");

        var (client, project, component) = await DesignTargetAsync(a);
        string container = a.Str("container") ?? "";
        bool onForm = container.Length == 0 || container.Equals(component, StringComparison.OrdinalIgnoreCase);
        if (onForm) container = "";
        var controls = Wire.Records((await DesignAsync(client, "CONTROLS", project, component)).Payload);
        string Key(string[] r) => r[0] + (At(r, 1) is { Length: > 0 } ix && ix != "-1" ? $"({ix})" : "");

        var current = new Dictionary<string, LayoutBox>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in controls.Where(r => string.Equals(At(r, 5), container, StringComparison.OrdinalIgnoreCase)))
            current[Key(r)] = new LayoutBox(Number(At(r, 6)) ?? 0, Number(At(r, 7)) ?? 0, Number(At(r, 8)) ?? 0, Number(At(r, 9)) ?? 0);
        // a name that is in another container gets its own message: moving between containers is form_set_property Container
        foreach (var name in spec.Rows.SelectMany(r => r).Where(n => !string.IsNullOrWhiteSpace(n) && !current.ContainsKey(n.Trim())).Distinct(StringComparer.OrdinalIgnoreCase))
            if (controls.FirstOrDefault(r => Key(r).Equals(name.Trim(), StringComparison.OrdinalIgnoreCase)) is { } other)
                throw new ToolException($"{name} is inside {(At(other, 5) is { Length: > 0 } c ? c : "the form")}, not {(onForm ? "the form" : container)}: lay out one container per call");

        // the container's inside size, and what to set to resize it
        var warnings = new List<string>();
        (double Width, double Height)? inner = null;
        double topInset = 0;
        (string Name, double ExtraW, double ExtraH)? sizeVia = null;   // Width/Height = inside size + extra (borders, title bar)
        var frame = onForm ? null : controls.FirstOrDefault(r => Key(r).Equals(container, StringComparison.OrdinalIgnoreCase))
            ?? throw new ToolException($"{container} is not a control on {component}");
        if (frame is not null && At(frame, 2) == "VB.Frame")
        {
            inner = (Number(At(frame, 8)) ?? 0, Number(At(frame, 9)) ?? 0);
            topInset = FrameCaptionInset;
            sizeVia = (container, 0, 0);
        }
        else
        {
            string target = onForm ? component : container;
            var props = Wire.Records((await DesignAsync(client, "PROPS", project, component, target)).Payload)
                .GroupBy(r => r[0], StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => Number(At(g.First(), 2)), StringComparer.OrdinalIgnoreCase);
            double? P(string k) => props.TryGetValue(k, out var v) ? v : null;
            if (P("ScaleWidth") is { } sw && P("ScaleHeight") is { } sh)
            {
                inner = (sw, sh);
                if (P("ScaleMode") is { } mode && mode != 1)
                    warnings.Add($"{(onForm ? component : container)} has ScaleMode {mode}, not twips: positions, gaps and margin are in its units");
                else if (P("Width") is { } w && P("Height") is { } h)
                    sizeVia = (onForm ? component : container, w - sw, h - sh);
            }
        }

        var layout = FormLayout.Compute(spec, current, inner, topInset);
        warnings.AddRange(sizeContainer ? layout.Warnings.Where(w => !w.Contains("past the container")) : layout.Warnings);

        var changes = new List<(string Control, string Property, string Value)>();
        object? containerSize = null;
        if (sizeContainer)
        {
            if (sizeVia is not { } via) throw new ToolException($"the size of {(onForm ? component : container)} can't be set from here (no Width/ScaleWidth in twips); leave out size_container");
            var (fw, fh) = FormLayout.FitInner(layout, spec.Margin);
            double newW = Math.Round(fw + via.ExtraW), newH = Math.Round(fh + via.ExtraH);
            changes.Add((via.Name, "Width", newW.ToString(CultureInfo.InvariantCulture)));
            changes.Add((via.Name, "Height", newH.ToString(CultureInfo.InvariantCulture)));
            containerSize = new { name = via.Name, width = newW, height = newH };
        }
        foreach (var (name, box) in layout.Boxes)
        {
            var now = current[name];
            void Add(string p, double was, double want) { if (Math.Abs(was - want) >= 1) changes.Add((name, p, want.ToString(CultureInfo.InvariantCulture))); }
            Add("Left", now.Left, box.Left); Add("Top", now.Top, box.Top); Add("Width", now.Width, box.Width); Add("Height", now.Height, box.Height);
        }
        var positions = layout.Boxes.OrderBy(kv => kv.Value.Top).ThenBy(kv => kv.Value.Left)
            .Select(kv => new { control = kv.Key, left = kv.Value.Left, top = kv.Value.Top, width = kv.Value.Width, height = kv.Value.Height }).ToList();
        var block = new { left = layout.Block.Left, top = layout.Block.Top, width = layout.Block.Width, height = layout.Block.Height };
        if (dryRun)
            return ToolResult.Ok(new { source = "ide", project, component, container = onForm ? component : container, dryRun = true, block, positions, containerSize, wouldSet = changes.Count, warnings = warnings.Count > 0 ? warnings : null });
        if (changes.Count == 0)
            return ToolResult.Ok(new { source = "ide", project, component, container = onForm ? component : container, block, positions, applied = 0, note = "already laid out like this", warnings = warnings.Count > 0 ? warnings : null });

        Audit.Write("form_layout", client.Instance.Pid, $"{project}.{component} {(onForm ? "" : container + " ")}{layout.Boxes.Count} controls, {changes.Count} changes");
        var (results, done, dirty, stopped, failed) = await ApplyChangesAsync(client, project, component, changes);
        var review = stopped is null && done.Count > 0 ? await ReviewAsync(client, project, component, dirty, done, null, byLayout: true) : (Dirty: dirty, Hints: new List<string>(), Overlaps: new List<object>());
        return ToolResult.Ok(new
        {
            source = "ide", project, component, container = onForm ? component : container, dirty = review.Dirty, block, positions, containerSize,
            applied = done.Count, failed, notAttempted = changes.Count - results.Count,
            errors = failed > 0 ? results.Where(r => r.GetType().GetProperty("ok")?.GetValue(r) is false).ToList() : null, stopped,
            warnings = warnings.Count > 0 ? warnings : null, hints = review.Hints.Count > 0 ? review.Hints : null, overlaps = review.Overlaps.Count > 0 ? review.Overlaps : null,
            note = "unsaved in the IDE; save with project_save",
        });
    }
}
