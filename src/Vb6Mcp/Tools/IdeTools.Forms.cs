using System.Globalization;
using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 5 form designer tools (PLAN.md §10): the controls of a form, UserControl or property page as
/// the IDE holds them right now (unsaved changes included), their properties, and changes to them.
/// Served by the add-in's DESIGN verb, which is deferred (hard rule 14: a UserControl's own code runs
/// when its properties are read or set), and design mode only (error 393 in run mode).
/// </summary>
public sealed partial class IdeTools
{
    private const string ControlDesc = "Control name, or Name(index) for an element of a control array.";

    private IReadOnlyList<Tool> FormTools() =>
    [
        new("form_controls",
            "List the controls of a form, UserControl or property page as the designer holds them now (unsaved changes included): name, " +
            "index (control arrays), ProgID, container, position and size (in the container's units, usually twips). Design mode; the " +
            "component's designer must be open in the IDE.",
            Schema.Object(("component", "string", ComponentDesc, true), ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => FormControlsAsync(a)),
        new("form_properties",
            "All design-time properties of one control (or of the form itself: pass the form's name) with their current values. Font " +
            "members are listed as Font.Name, Font.Size, Font.Bold, ...; pictures as Picture.Handle (0 = none), ...; indexed properties (List, ItemData) with their items.",
            Schema.Object(("component", "string", ComponentDesc, true), ("control", "string", ControlDesc, true),
                          ("names", "array", "Only these properties (names or prefixes ending in '.', e.g. Caption, Left, Font.); default all (about 50-100).", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => FormPropertiesAsync(a)),
        new("form_set_property",
            "Set a design-time property of a control (or of the form: pass its name). The value is text, converted to the property's " +
            "type: numbers ('1200', '&H8000000F&' for colours), booleans ('True'/'False'), strings as given. Font members: Font.Name, " +
            "Font.Size, Font.Bold, Font.Italic, Font.Underline, Font.Strikethrough, Font.Charset. Pictures (Picture, Icon, MouseIcon): value '@picture:<full path>' (a file under the workspace roots), or '@picture:' to clear. Lists (List, ItemData): one item per line. Renaming (property Name) doesn't rename " +
            "the control's event procedures, just like in the IDE. VB6 rules: a form's Font doesn't change the controls already on it; a " +
            "CommandButton shows BackColor only with Style = 1 and has no ForeColor. The result has hints and overlaps (controls hiding each " +
            "other) when they apply. For several changes use form_set_properties (one call). Unsaved; save with project_save.",
            Schema.Object(("component", "string", ComponentDesc, true), ("control", "string", ControlDesc, true),
                          ("property", "string", "Property name, e.g. Caption, Left, BackColor, Font.Bold.", true),
                          ("value", "string", "New value as text.", true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => FormSetPropertyAsync(a)),
        new("form_set_properties",
            "Set several design-time properties in one call, on one form/UserControl: changes is a list of {control, property, value}, " +
            "values as for form_set_property (control = the form's own name for form properties). Applied in order; a value VB refuses " +
            "is reported and the rest still run (nothing is rolled back). Stops early if the IDE gets busy or shows a dialog. VB6 rules: " +
            "the form's Font doesn't change the controls already on it (set their Font.* too); a CommandButton shows BackColor only with " +
            "Style = 1 and has no ForeColor. After moves the result lists overlaps (controls hiding each other) and hints; check the layout " +
            "with form_screenshot. To arrange many controls in rows and columns use form_layout. Unsaved; save with project_save.",
            ChangesSchema(Schema.Object(("component", "string", ComponentDesc, true),
                          ("changes", "array", "List of {control, property, value} objects (at most 500).", true),
                          ("verbose", "boolean", "List every change with the value VB read back (default: only failed ones).", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false))),
            (a, _) => FormSetPropertiesAsync(a)),
        FormLayoutTool(),
        new("form_add_control",
            "Add a control to a form/UserControl in the designer. prog_id is the class name as a .frm file shows it (Begin <class>): " +
            "VB.CommandButton, VB.TextBox, VB.Label, VB.Frame, or for ActiveX controls e.g. MSComctlLib.ListView (not the registered ProgID " +
            "MSComctlLib.ListViewCtrl.2; add the component first with refs_add_component, which takes that one). Optional name, container (a Frame or " +
            "PictureBox to put it in), position/size and properties ({Caption: ..., Font.Bold: True}, set like form_set_properties). For a group of " +
            "controls, give sizes only and arrange them afterwards with form_layout. With left/top, the result lists overlaps with sibling controls. Unsaved; save with project_save.",
            Schema.Object(("component", "string", ComponentDesc, true), ("prog_id", "string", "Class name as in a .frm: VB.CommandButton, MSComctlLib.ListView.", true),
                          ("name", "string", "Control name (default: VB's, e.g. Command1).", false),
                          ("container", "string", "Container control to add it to (default: the form).", false),
                          ("left", "number", "Left.", false), ("top", "number", "Top.", false), ("width", "number", "Width.", false), ("height", "number", "Height.", false),
                          ("properties", "object", "Properties to set on the new control, name to value as text (Caption, BackColor, Font.Name...).", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => FormAddControlAsync(a)),
        new("form_screenshot",
            "Screenshot of a form, UserControl or property page as it looks in the designer now (unsaved changes included): " +
            "layout, fonts, colours, without running the program. Opens the component's designer in the IDE (maximized inside the " +
            "IDE unless maximize=false), captures it and crops to the form. If the form is larger than the visible designer area, " +
            "the image is cut off and the result says so. Design mode.",
            Schema.Object(("component", "string", ComponentDesc, true), ("maximize", "boolean", "Maximize the designer window first (default true).", false),
                          ("max_width", "integer", "Scale the image down to at most this width (default 1024).", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => FormScreenshotAsync(a, ct)),
        new("form_remove_control",
            "Remove a control (and any controls inside it) from the designer. Its event procedures stay in the code, as when deleting by " +
            "hand. Unsaved (Edit > Undo in VB6 can bring it back). Needs confirm=true.",
            Schema.Object(("component", "string", ComponentDesc, true), ("control", "string", ControlDesc, true), ("confirm", "boolean", "Required: true.", true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => FormRemoveControlAsync(a)),
    ];

    private async Task<(AddinClient Client, string Project, string Component)> DesignTargetAsync(JsonObject a)
    {
        RequireDesignMode(ResolveInstance(a));
        var (client, project) = await ResolveProjectAsync(a);
        return (client, project, a.Req("component"));
    }

    private static async Task<(bool Dirty, string Payload)> DesignAsync(AddinClient client, string op, params string[] args)
    {
        try { return await client.DesignAsync(op, args, blockedBy: () => BlockingDialog(client.Instance.Pid)); }
        catch (AddinErrorException ex)
        {
            // our own add-in errors carry vbObjectError + n ("-2147221104: not a number: abc"); the number means nothing to the agent
            throw new ToolException(System.Text.RegularExpressions.Regex.Replace(ex.AddinMessage, @"^-21472\d{5}: ", ""));
        }
    }

    /// <summary>
    /// A dialog that opened while a designer op waited, described for the agent, or null. Reading or opening a designer
    /// runs the code of UserControls on it (spike 23); in a real project (spike 27) that showed a "Login Error" box,
    /// and the deferred op only finishes once the dialog is answered.
    /// </summary>
    private static string? BlockingDialog(int pid)
    {
        var d = IdeWindows.Dialogs(pid).FirstOrDefault();
        return d is null ? null
            : $"a dialog opened in the IDE while the designer was being used, probably code of a UserControl on the form: " +
              $"'{d.Caption}': {string.Join(" ", d.Texts)} [buttons: {string.Join(", ", d.Buttons.Select(b => b.Caption))}]. " +
              "Answer it with ide_dialog_respond; the designer request then finishes by itself, so try again afterwards.";
    }

    private async Task<ToolResult> FormControlsAsync(JsonObject a)
    {
        var (client, project, component) = await DesignTargetAsync(a);
        var (dirty, payload) = await DesignAsync(client, "CONTROLS", project, component);
        var controls = Wire.Records(payload).Select(ControlJson).ToList();
        return ToolResult.Ok(new { source = "ide", project, component, dirty, count = controls.Count, controls });
    }

    private async Task<ToolResult> FormPropertiesAsync(JsonObject a)
    {
        var (client, project, component) = await DesignTargetAsync(a);
        var control = a.Req("control");
        var (_, payload) = await DesignAsync(client, "PROPS", project, component, control);
        var names = a.List("names");
        var props = PropsJson(payload);
        if (names.Count > 0)
            props = props.Where(p => p.GetType().GetProperty("name")?.GetValue(p) is string n &&
                names.Any(x => x.EndsWith('.') ? n.StartsWith(x, StringComparison.OrdinalIgnoreCase) : n.Equals(x, StringComparison.OrdinalIgnoreCase))).ToList();
        return ToolResult.Ok(new { source = "ide", project, component, control, properties = props });
    }

    private async Task<ToolResult> FormSetPropertyAsync(JsonObject a)
    {
        var (client, project, component) = await DesignTargetAsync(a);
        var control = a.Req("control");
        var property = a.Req("property");
        var value = a.Str("value") ?? throw new ToolException("missing required argument 'value'");
        string payload; bool dirty;
        try { (dirty, payload) = await DesignAsync(client, "SET", project, component, control, property, CheckPicture(value)); }
        catch (ToolException ex) { throw new ToolException(ex.Message + ExplainSetError(property, ex.Message)); }
        var review = await ReviewAsync(client, project, component, dirty, [(control, property, value)], null);
        return ToolResult.Ok(new
        {
            source = "ide", project, component, control, dirty = review.Dirty, now = PropsJson(payload),
            hints = review.Hints.Count > 0 ? review.Hints : null, overlaps = review.Overlaps.Count > 0 ? review.Overlaps : null,
            note = "unsaved in the IDE; save with project_save",
        });
    }

    /// <summary>A '@picture:&lt;path&gt;' value with its path checked and normalized; other values unchanged.</summary>
    private string CheckPicture(string value)
    {
        if (!value.StartsWith("@picture:", StringComparison.Ordinal) || value.Length <= 9) return value;
        // the picture file is read by the elevated IDE: same roots rule as every other path (hard rule 8); the add-in checks again
        var pic = PathKey.Normalize(value[9..]);
        if (Roots.Count == 0 || !ProjectMap.IsUnderRoots(pic, Roots)) throw new ToolException($"{pic} is outside the workspace roots");
        if (!File.Exists(pic)) throw new ToolException("picture not found: " + pic);
        return "@picture:" + pic;
    }

    /// <summary>The changes argument's schema: a list of {control, property, value} objects (Schema.Object only makes string lists).</summary>
    private static JsonObject ChangesSchema(JsonObject schema)
    {
        schema["properties"]!["changes"]!["items"] = new JsonObject
        {
            ["type"] = "object",
            ["properties"] = new JsonObject
            {
                ["control"] = new JsonObject { ["type"] = "string" },
                ["property"] = new JsonObject { ["type"] = "string" },
                ["value"] = new JsonObject { ["type"] = "string" },
            },
            ["required"] = new JsonArray("control", "property", "value"),
        };
        return schema;
    }

    /// <summary>
    /// form_set_properties. Restyling six real forms took 107–160 form_set_property calls (benchmark T8), and each MCP
    /// round trip costs the agent tokens. The add-in calls are cheap, so the server loops the existing deferred DESIGN SET
    /// (no new verb): one MCP call, the same per-property behaviour, conversion and checks.
    /// </summary>
    private async Task<ToolResult> FormSetPropertiesAsync(JsonObject a)
    {
        var (client, project, component) = await DesignTargetAsync(a);
        var items = a["changes"] switch
        {
            JsonArray arr => arr,
            JsonValue v when v.TryGetValue(out string? text) => JsonNode.Parse(text) as JsonArray,   // some clients send the list as a JSON string
            _ => null,
        } ?? throw new ToolException("changes must be a list of {control, property, value} objects");
        if (items.Count == 0) throw new ToolException("changes is empty");
        if (items.Count > 500) throw new ToolException($"at most 500 changes per call ({items.Count} given)");
        // validate everything before changing anything, so a malformed list changes nothing
        var changes = items.Select((n, i) => n is JsonObject o && o.Str("control") is { Length: > 0 } c && o.Str("property") is { Length: > 0 } p && o.Str("value") is { } v
            ? (Control: c, Property: p, Value: v) : throw new ToolException($"changes[{i}] needs control, property and value (strings)")).ToList();

        var (results, done, dirty, stopped, failed) = await ApplyChangesAsync(client, project, component, changes);
        var review = stopped is null && done.Count > 0 ? await ReviewAsync(client, project, component, dirty, done, null) : (Dirty: dirty, Hints: new List<string>(), Overlaps: new List<object>());
        return ToolResult.Ok(new
        {
            source = "ide", project, component, dirty = review.Dirty, applied = results.Count - failed, failed, notAttempted = changes.Count - results.Count,
            // Only failures by default: echoing every change was 15% of a redesign run's re-read tokens (docs/benchmarks, token breakdown)
            results = a.Bool("verbose") == true ? results : results.Where(r => r.GetType().GetProperty("ok")?.GetValue(r) is false).ToList(),
            stopped, hints = review.Hints.Count > 0 ? review.Hints : null, overlaps = review.Overlaps.Count > 0 ? review.Overlaps : null,
            note = "unsaved in the IDE; save with project_save" + (a.Bool("verbose") == true ? "" : ". results lists failed changes only (verbose=true: all, with values read back)"),
        });
    }

    /// <summary>
    /// Sets each property through the deferred DESIGN SET, in order (form_set_properties, form_layout). A value VB refuses is
    /// reported and the rest still run; a dialog or a busy IDE stops the loop, since it would hold every later SET too.
    /// </summary>
    private async Task<(List<object> Results, List<(string Control, string Property, string Value)> Done, bool Dirty, string? Stopped, int Failed)> ApplyChangesAsync(
        AddinClient client, string project, string component, IReadOnlyList<(string Control, string Property, string Value)> changes)
    {
        var results = new List<object>();
        var done = new List<(string Control, string Property, string Value)>();
        bool dirty = false; string? stopped = null; int failed = 0;
        foreach (var (control, property, value) in changes)
        {
            try
            {
                var (d, payload) = await DesignAsync(client, "SET", project, component, control, property, CheckPicture(value));
                dirty = d;
                done.Add((control, property, value));
                // only the property that was set (Font.Bold returns every Font member); lists and anything unexpected in full
                var own = Wire.Records(payload).Where(r => string.Equals(r[0], property, StringComparison.OrdinalIgnoreCase)).ToList();
                bool single = own.Count == 1 && At(own[0], 1) != "i";
                results.Add(new { control, property, ok = true, now = single ? At(own[0], 2) : null, nowAll = single ? null : PropsJson(payload) });
            }
            catch (ToolException ex) when (!ex.Message.StartsWith("a dialog opened", StringComparison.Ordinal))
            {
                results.Add(new { control, property, ok = false, error = ex.Message + ExplainSetError(property, ex.Message) });
                failed++;
            }
            catch (Exception ex) when (ex is ToolException or AddinBusyException or AddinUnreachableException)
            {
                // a dialog or a busy IDE holds every later SET too: stop and say what was not attempted
                results.Add(new { control, property, ok = false, error = ex.Message });
                failed++;
                stopped = ex.Message;
                break;
            }
        }
        return (results, done, dirty, stopped, failed);
    }

    static readonly string[] GeometryProps = ["Left", "Top", "Width", "Height", "Visible", "Container"];
    // Images, shapes and lines are often meant to sit under other controls (a background, a border); timers aren't drawn
    // at run time. Labels are windowless too: VB draws them on the container, behind every windowed control, so a label
    // under a frame or button is hidden (T10: header labels under frames). Label on label is common and not reported.
    static readonly string[] NotCheckedForOverlap = ["VB.Image", "VB.Shape", "VB.Line", "VB.Timer"];
    const double MinOverlapTwips = 30;   // 2 pixels: smaller ones were edge-touching noise in T10 (15 twips = 1 pixel)

    /// <summary>VB's error for a property, plus what it means when we know (verified on this machine, see docs/tools.md).</summary>
    private static string ExplainSetError(string property, string error)
        => property.Equals("ForeColor", StringComparison.OrdinalIgnoreCase) && error.StartsWith("438", StringComparison.Ordinal)
            ? " (VB6 command buttons have no ForeColor: their caption colour can't be changed. Style = 1 allows BackColor and a picture; for coloured text use a Label or a PictureBox.)"
            : "";

    /// <summary>
    /// What to tell the agent after designer changes, and one repair:
    /// - A change made only through an object member (Font.Bold) isn't counted by VB: IsDirty stays False and File >
    ///   Save Project skips the form, so the change is silently lost (verified, docs/tools.md "Designer rules"). Setting the
    ///   form's Tag to its own value makes VB count it.
    /// - A form's Font doesn't change the controls already on it (verified through VBIDE).
    /// - A CommandButton shows BackColor only with Style = 1 (verified with form_screenshot).
    /// - Sibling controls that overlap after a move or an added control: they hide each other (benchmark T9 lost a
    ///   button that way). One CONTROLS read, only when geometry changed or a control was added.
    /// </summary>
    private async Task<(bool Dirty, List<string> Hints, List<object> Overlaps)> ReviewAsync(AddinClient client, string project, string component, bool dirty,
        IReadOnlyList<(string Control, string Property, string Value)> done, string? added, bool byLayout = false, bool listOverlaps = true)
    {
        var hints = new List<string>();
        var overlaps = new List<object>();
        bool IsForm(string c) => string.Equals(c, component, StringComparison.OrdinalIgnoreCase);

        if (!dirty && done.Any(d => d.Property.Contains('.')))
        {
            try
            {
                var tag = Wire.Records((await DesignAsync(client, "PROPS", project, component, component)).Payload).FirstOrDefault(r => r[0] == "Tag" && At(r, 1) == "s");
                if (tag is not null) dirty = (await DesignAsync(client, "SET", project, component, component, "Tag", At(tag, 2))).Dirty;
            }
            catch (Exception ex) when (ex is ToolException or AddinBusyException) { hints.Add("Couldn't make VB count the change: " + ex.Message); }
            if (!dirty) hints.Add("VB doesn't count a change made only to Font.* members as a change, so Save would skip this form. Change any other property (e.g. Tag) before saving.");
        }
        if (done.Any(d => IsForm(d.Control) && d.Property.StartsWith("Font.", StringComparison.OrdinalIgnoreCase)))
            hints.Add("The form's Font doesn't change the controls already on it: each keeps its own font. Set Font.* on the controls too.");

        bool geometry = added is not null || done.Any(d => GeometryProps.Contains(d.Property, StringComparer.OrdinalIgnoreCase));
        var backColors = done.Where(d => d.Property.Equals("BackColor", StringComparison.OrdinalIgnoreCase) && !IsForm(d.Control)).ToList();
        if (!geometry && backColors.Count == 0) return (dirty, hints, overlaps);

        IReadOnlyList<string[]> controls;
        try { controls = Wire.Records((await DesignAsync(client, "CONTROLS", project, component)).Payload); }
        catch (Exception ex) when (ex is ToolException or AddinBusyException)
        {
            // e.g. a UserControl's dialog opened when VB reloaded the designer: the changes are made, the check isn't
            hints.Add("The changes were made, but the layout check couldn't run: " + ex.Message);
            return (dirty, hints, overlaps);
        }
        string Key(string[] r) => r[0] + (At(r, 1) is { Length: > 0 } ix && ix != "-1" ? $"({ix})" : "");
        string Base(string name) => name.Split('(')[0];

        // CommandButton BackColor without Style = 1 (Graphical) shows nothing
        foreach (var b in backColors.Select(d => d.Control).Distinct(StringComparer.OrdinalIgnoreCase).Take(10))
        {
            if (controls.FirstOrDefault(r => string.Equals(Key(r), b, StringComparison.OrdinalIgnoreCase) || string.Equals(r[0], b, StringComparison.OrdinalIgnoreCase)) is not { } rec
                || At(rec, 2) != "VB.CommandButton") continue;
            if (done.Any(d => string.Equals(d.Control, b, StringComparison.OrdinalIgnoreCase) && d.Property.Equals("Style", StringComparison.OrdinalIgnoreCase) && d.Value.Trim() == "1")) continue;
            try
            {
                var style = Wire.Records((await DesignAsync(client, "PROPS", project, component, b)).Payload).FirstOrDefault(r => r[0] == "Style");
                if (style is not null && At(style, 2) != "1")
                    hints.Add($"{b} is a CommandButton with Style = {At(style, 2)}: its BackColor only shows with Style = 1 (Graphical).");
            }
            catch (ToolException) { }
        }
        if (!geometry) return (dirty, hints, overlaps);

        var touched = new HashSet<string>(done.Where(d => GeometryProps.Contains(d.Property, StringComparer.OrdinalIgnoreCase)).Select(d => Base(d.Control)), StringComparer.OrdinalIgnoreCase);
        if (added is not null) touched.Add(Base(added));
        // the .frm class (VB.Label) is the progId field; the className field is the short name (Label) (spike 29)
        var boxes = controls.Where(r => !NotCheckedForOverlap.Contains(At(r, 2)) && Number(At(r, 8)) is > 0 && Number(At(r, 9)) is > 0)
            .Select(r => (Name: Key(r), Base: r[0], Container: At(r, 5), Label: At(r, 2) == "VB.Label",
                          L: Number(At(r, 6)) ?? 0, T: Number(At(r, 7)) ?? 0, W: Number(At(r, 8))!.Value, H: Number(At(r, 9))!.Value)).ToList();
        for (int i = 0; i < boxes.Count; i++)
            for (int j = i + 1; j < boxes.Count; j++)
            {
                var (x, y) = (boxes[i], boxes[j]);
                if (x.Container != y.Container || string.Equals(x.Base, y.Base, StringComparison.OrdinalIgnoreCase)) continue;   // a control array is stacked on purpose
                if (!touched.Contains(x.Base) && !touched.Contains(y.Base)) continue;
                double w = Math.Min(x.L + x.W, y.L + y.W) - Math.Max(x.L, y.L), h = Math.Min(x.T + x.H, y.T + y.H) - Math.Max(x.T, y.T);
                if (w < MinOverlapTwips || h < MinOverlapTwips || (x.Label && y.Label)) continue;
                // A Label that wholly contains the other control is a background (a coloured header band or status strip
                // with buttons on it, T12): meant to be behind. Only partial overlaps hide part of a label.
                static bool Inside((double L, double T, double W, double H) i, (double L, double T, double W, double H) o)
                    => i.L >= o.L && i.T >= o.T && i.L + i.W <= o.L + o.W && i.T + i.H <= o.T + o.H;
                if ((x.Label && Inside((y.L, y.T, y.W, y.H), (x.L, x.T, x.W, x.H))) || (y.Label && Inside((x.L, x.T, x.W, x.H), (y.L, y.T, y.W, y.H)))) continue;
                bool same = x.L == y.L && x.T == y.T && x.W == y.W && x.H == y.H;
                string? hidden = x.Label ? x.Name : y.Label ? y.Name : null;
                overlaps.Add(new
                {
                    controls = new[] { x.Name, y.Name }, container = x.Container.Length > 0 ? x.Container : component, overlap = new { width = w, height = h }, sameRectangle = same,
                    note = hidden is null ? null : $"{hidden} is a Label: VB draws it behind windowed controls, so the overlapped part is hidden",
                });
                if (overlaps.Count >= 20) break;
            }
        if (overlaps.Count > 0 && !listOverlaps)
        {
            // Improvement notes 43: thirty adds at the default position listed every overlap with the controls still
            // waiting to be arranged, so the results grew with the form (O(n²) tokens). The count is enough there.
            hints.Add($"The new control overlaps {(overlaps.Count >= 20 ? "20 or more" : overlaps.Count.ToString(CultureInfo.InvariantCulture))} other control(s) (not listed: no left/top was given). " +
                      "Arrange the controls with form_layout once they're all added; it reports what still overlaps.");
            overlaps.Clear();
        }
        if (overlaps.Count > 0)
            hints.Add("Some controls overlap (see overlaps): the one in front hides the other. Move them apart, unless they're never visible at the same time. Code may show controls that are hidden at design time.");
        // not after form_layout itself: it placed them (improvement notes 38)
        if (!byLayout && LayoutSuggestion(component, controls, done) is { } suggestion) hints.Add(suggestion);
        hints.Add("Check the layout with form_screenshot.");
        return (dirty, hints, overlaps);
    }

    /// <summary>Controls moved by hand in one container before form_layout is suggested.</summary>
    const int LayoutHintMin = 4;

    /// <summary>
    /// When one call moved several controls of a container by hand (Left/Top), the form_layout call that arranges them, rows
    /// taken from where they ended up. Benchmark T12 round 6: agents worked out twips in Python and set them with
    /// form_set_properties although form_layout was loaded; the hint comes exactly when they do that. Rows: controls whose
    /// Top is within 120 twips of the row's first control, each row left to right.
    /// </summary>
    private static string? LayoutSuggestion(string component, IReadOnlyList<string[]> controls, IReadOnlyList<(string Control, string Property, string Value)> done)
    {
        var moved = done.Where(d => d.Property.Equals("Left", StringComparison.OrdinalIgnoreCase) || d.Property.Equals("Top", StringComparison.OrdinalIgnoreCase))
            .Select(d => d.Control).Distinct(StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (moved.Count < LayoutHintMin) return null;
        string Key(string[] r) => r[0] + (At(r, 1) is { Length: > 0 } ix && ix != "-1" ? $"({ix})" : "");
        var group = controls.Where(r => moved.Contains(Key(r)) && !NotCheckedForOverlap.Contains(At(r, 2)))
            .Select(r => (Name: Key(r), Container: At(r, 5), L: Number(At(r, 6)) ?? 0, T: Number(At(r, 7)) ?? 0))
            .GroupBy(x => x.Container).OrderByDescending(g => g.Count()).FirstOrDefault(g => g.Count() >= LayoutHintMin);
        if (group is null) return null;
        var rows = new List<List<string>>();
        double rowTop = double.NaN;
        foreach (var c in group.OrderBy(x => x.T).ThenBy(x => x.L))
        {
            if (rows.Count == 0 || Math.Abs(c.T - rowTop) > 120) { rows.Add([]); rowTop = c.T; }
            rows[^1].Add(c.Name);
        }
        foreach (var row in rows) row.Sort((x, y) => group.First(c => c.Name == x).L.CompareTo(group.First(c => c.Name == y).L));
        var json = "[" + string.Join(",", rows.Select(r => "[" + string.Join(",", r.Select(n => "\"" + n + "\"")) + "]")) + "]";
        return $"You placed {group.Count()} controls by hand. form_layout arranges rows and columns in one call with even spacing (sizes from the controls, no twips to work out): " +
               $"form_layout component={component}{(group.Key.Length > 0 ? $" container={group.Key}" : "")} rows={json} (dry_run=true previews; gap_x/gap_y/column_widths adjust it).";
    }

    private async Task<ToolResult> FormAddControlAsync(JsonObject a)
    {
        var (client, project, component) = await DesignTargetAsync(a);
        string Num(string k) => a[k] is JsonValue v && v.TryGetValue(out double d) ? d.ToString(CultureInfo.InvariantCulture) : "";
        var geometry = string.Join(",", Num("left"), Num("top"), Num("width"), Num("height")).TrimEnd(',');
        // properties are checked before the control is added, so a malformed argument adds nothing
        var props = a["properties"] switch
        {
            null => [],
            JsonObject o => o.Select(kv => (Name: kv.Key, Value: kv.Value is JsonValue v ? (v.TryGetValue(out string? t) ? t : v.ToJsonString()) : null)).ToList(),
            _ => throw new ToolException("properties must be an object: {\"Caption\": \"OK\", \"Font.Bold\": \"True\"}"),
        };
        if (props.FirstOrDefault(p => p.Value is null).Name is { } bad)
            throw new ToolException($"property {bad}: the value must be text, a number or true/false");
        var (dirty, payload) = await DesignAsync(client, "ADD", project, component, a.Req("prog_id"), a.Str("name") ?? "", a.Str("container") ?? "", geometry);
        var rec = payload.Split(Wire.US);
        string added = rec[0] + (rec.Length > 1 && rec[1] is { Length: > 0 } ix && ix != "-1" ? $"({ix})" : "");
        // Set the properties on the new control (benchmark T8: captions etc. cost a round trip each). A value VB refuses
        // is reported; the control stays, as with form_set_properties.
        var set = new List<object>();
        var done = new List<(string Control, string Property, string Value)>();
        foreach (var (name, value) in props)
        {
            try
            {
                var (d, _) = await DesignAsync(client, "SET", project, component, added, name, CheckPicture(value!));
                dirty = d; done.Add((added, name, value!));
                set.Add(new { property = name, ok = true });
            }
            catch (ToolException ex) { set.Add(new { property = name, ok = false, error = ex.Message + ExplainSetError(name, ex.Message) }); }
        }
        bool placed = Num("left").Length > 0 || Num("top").Length > 0 ||
                      done.Any(d => d.Property.Equals("Left", StringComparison.OrdinalIgnoreCase) || d.Property.Equals("Top", StringComparison.OrdinalIgnoreCase));
        var review = await ReviewAsync(client, project, component, dirty, done, added, listOverlaps: placed);
        return ToolResult.Ok(new
        {
            source = "ide", project, component, dirty = review.Dirty, added = ControlJson(rec), properties = set.Count > 0 ? set : null,
            hints = review.Hints.Count > 0 ? review.Hints : null, overlaps = review.Overlaps.Count > 0 ? review.Overlaps : null,
            note = "unsaved in the IDE; save with project_save",
        });
    }

    private async Task<ToolResult> FormRemoveControlAsync(JsonObject a)
    {
        if (a.Bool("confirm") != true) throw new ToolException("removing a control needs confirm=true");
        var (client, project, component) = await DesignTargetAsync(a);
        var control = a.Req("control");
        var (dirty, _) = await DesignAsync(client, "REMOVE", project, component, control);
        return ToolResult.Ok(new { source = "ide", project, component, removed = control, dirty, note = "unsaved in the IDE; event procedures of the control stay in the code" });
    }

    /// <summary>
    /// form_screenshot. The add-in opens the designer (DESIGN SHOW: VBComponent.Activate, deferred because a form hosting a
    /// UserControl runs its code); the server finds the designer MDI child by the exact caption the add-in reports and, in
    /// it, the form surface (the largest visible child that isn't a scroll bar). It captures the whole IDE window (an
    /// elevated IDE through the add-in, spike 15: child panes captured on their own come out misaligned) and crops.
    /// </summary>
    private async Task<ToolResult> FormScreenshotAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireDesignMode(inst);
        var (client, project) = await ResolveProjectAsync(a);
        var component = a.Req("component");
        bool maximize = a.Bool("maximize") ?? true;
        // the add-in returns the designer window's own caption (VBIDE DesignerWindow.Caption: "frmMain (Form)", no project prefix, spike 28)
        var (dirty, caption) = await DesignAsync(client, "SHOW", maximize ? [project, component, "max"] : [project, component]);
        await Task.Delay(400, ct);      // let the MDI child appear and paint

        // Match that exact caption among the MDI children; Activate brought ours to the top of the MDI z-order, and
        // EnumChildWindows walks in z-order, so if two windows somehow share a caption the first one is ours.
        var wins = IdeWindows.Descendants(inst.Pid);
        var mdi = wins.FirstOrDefault(w => w.Class == "MDIClient");
        var mdiChildren = wins.Where(w => w.Visible && (mdi is null || w.Parent == mdi.Hwnd)).ToList();
        var designer = mdiChildren.FirstOrDefault(w => string.Equals(w.Caption, caption, StringComparison.OrdinalIgnoreCase))
            ?? throw new ToolException($"the designer window '{caption}' wasn't found after opening it; MDI windows: " +
                                       string.Join("; ", mdiChildren.Select(w => $"'{w.Caption}' ({w.Class})")));
        var parentOf = wins.ToDictionary(w => w.Hwnd, w => w.Parent);
        bool Under(long h) { for (int i = 0; i < 50 && parentOf.TryGetValue(h, out var p); i++) { if (p == designer.Hwnd) return true; h = p; } return false; }
        var surface = wins.Where(w => w.Visible && Under(w.Hwnd) && !w.Class.Contains("ScrollBar", StringComparison.OrdinalIgnoreCase))
                          .OrderByDescending(w => (long)w.Width * w.Height).FirstOrDefault() ?? designer;

        var main = IdeWindows.MainRect(inst.Pid) ?? throw new ToolException("the IDE main window wasn't found");
        // What of the form can be seen: the designer's client area (no title bar or scroll bars), inside the MDI area
        // (a maximized designer's frame sticks out of it).
        var view = IdeWindows.ClientRect(designer.Hwnd) ?? (designer.Left, designer.Top, designer.Right, designer.Bottom);
        if (mdi is not null && IdeWindows.ClientRect(mdi.Hwnd) is { } m)
            view = (Math.Max(view.Left, m.Left), Math.Max(view.Top, m.Top), Math.Min(view.Right, m.Right), Math.Min(view.Bottom, m.Bottom));
        int formW = surface.Width, formH = surface.Height;
        bool bigger = formW > view.Right - view.Left || formH > view.Bottom - view.Top;

        // A form bigger than the designer: scroll the designer a page at a time and stitch the visible pieces (spike 28:
        // WM_VSCROLL/WM_HSCROLL posted to the DesignerWindow move the form; a capture of the form window alone is
        // misaligned, spike 15). The scroll position goes back to the top left afterwards.
        var add = _env.Client(inst);
        using var canvas = new System.Drawing.Bitmap(formW, formH);
        using var covered = new System.Drawing.Region(System.Drawing.Rectangle.Empty);
        int captures = 0;
        if (bigger)
        {
            Audit.Write("form_screenshot", inst.Pid, $"scroll designer {designer.Hwnd:X} of {component} to stitch {formW}x{formH}");
            await add.ScrollAsync(designer.Hwnd, true, 6, 1);    // SB_TOP
            await add.ScrollAsync(designer.Hwnd, false, 6, 1);
            await Task.Delay(200, ct);
        }
        try
        {
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    var sr = IdeWindows.WindowRect(surface.Hwnd) ?? throw new ToolException("the form window disappeared");
                    using (var bmp = await CaptureIdeAsync(inst, ct))
                    {
                        captures++;
                        double scale = (double)bmp.Width / Math.Max(1, main.Right - main.Left);   // 1 unless DPI virtualization
                        int l = Math.Max(sr.Left, view.Left), t = Math.Max(sr.Top, view.Top);
                        int r = Math.Min(sr.Right, view.Right), btm = Math.Min(sr.Bottom, view.Bottom);
                        var crop = new System.Drawing.Rectangle((int)((l - main.Left) * scale), (int)((t - main.Top) * scale), (int)((r - l) * scale), (int)((btm - t) * scale));
                        crop.Intersect(new System.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height));
                        if (r > l && btm > t && crop.Width > 0 && crop.Height > 0)
                        {
                            var dest = new System.Drawing.Rectangle(l - sr.Left, t - sr.Top, r - l, btm - t);
                            using var g = System.Drawing.Graphics.FromImage(canvas);
                            g.DrawImage(bmp, dest, crop, System.Drawing.GraphicsUnit.Pixel);
                            covered.Union(dest);
                        }
                    }
                    if (!bigger || sr.Right <= view.Right) break;
                    await add.ScrollAsync(designer.Hwnd, false, 3, 1);    // SB_PAGEDOWN on the horizontal bar: a page right
                    await Task.Delay(150, ct);
                    if (IdeWindows.WindowRect(surface.Hwnd) is { } moved && moved.Left == sr.Left) break;   // can't scroll further
                }
                if (!bigger) break;
                var now = IdeWindows.WindowRect(surface.Hwnd) ?? throw new ToolException("the form window disappeared");
                if (now.Bottom <= view.Bottom) break;
                await add.ScrollAsync(designer.Hwnd, false, 6, 1);    // back to the left edge
                await add.ScrollAsync(designer.Hwnd, true, 3, 1);     // a page down
                await Task.Delay(150, ct);
                if (IdeWindows.WindowRect(surface.Hwnd) is { } down && down.Top == now.Top) break;
            }
        }
        finally
        {
            if (bigger)
            {
                try { await add.ScrollAsync(designer.Hwnd, true, 6, 1); await add.ScrollAsync(designer.Hwnd, false, 6, 1); }
                catch (Exception) { /* the image is still valid; the designer just stays scrolled */ }
            }
        }

        // The tiles are rectangles, so the covered area is the sum of the region's scans. Anything missing (the designer
        // couldn't scroll that far) makes the image incomplete, and the result says so.
        double area = covered.GetRegionScans(new System.Drawing.Drawing2D.Matrix()).Sum(rc => (double)rc.Width * rc.Height);
        if (area <= 0) throw new ToolException("the form isn't visible in the designer window");
        bool complete = area >= (double)formW * formH - 0.5;
        var png = Screenshots.ToPng(canvas, a.Int("max_width") ?? 1024);
        return ToolResult.Image(png, new
        {
            project, component, designer = designer.Caption, formWindowClass = surface.Class,
            formSize = new { width = formW, height = formH }, clipped = !complete, stitched = captures > 1 ? captures : (int?)null, dirty,
            note = string.Join(" ", new[]
            {
                !complete ? "Part of the form couldn't be captured (the designer didn't scroll that far), so those parts of the image are blank." : null,
                // spike 28: opening a production project's frmMain ran its UserControls' code and marked the form changed
                dirty ? "The component has unsaved changes. Opening a designer runs the code of UserControls on it, which can mark the form changed, as opening it by hand does; don't save just for that." : null,
            }.Where(n => n != null)) is { Length: > 0 } n ? n : null,
        });
    }

    /// <summary>The whole IDE main window as a bitmap: in-process through the add-in for an elevated IDE (spike 15), else locally.</summary>
    private async Task<System.Drawing.Bitmap> CaptureIdeAsync(InstanceInfo inst, CancellationToken ct)
    {
        long hwnd = IdeWindows.MainWindow(inst.Pid);
        if (hwnd == 0) throw new ToolException("the IDE main window wasn't found");
        if (Vb6Install.IsElevated(inst.Pid) != true && Screenshots.CaptureLocal(hwnd) is { } local) return local;
        string id = "f" + DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture);
        string path = await _env.Client(inst).SnapAsync(hwnd, id);
        var deadline = DateTime.UtcNow + SnapWait;
        while (!File.Exists(path) && DateTime.UtcNow < deadline) await Task.Delay(100, ct);
        if (!File.Exists(path)) throw new ToolException("the IDE didn't produce the screenshot in time");
        var bmp = Screenshots.LoadBmp(path);
        try { File.Delete(path); } catch (IOException) { }
        return bmp;
    }

    // name US index US progId US className US controlType US container US left US top US width US height
    private static object ControlJson(string[] f) => new
    {
        name = f[0], index = f.Length > 1 && f[1] is { Length: > 0 } i && i != "-1" ? int.Parse(i, CultureInfo.InvariantCulture) : (int?)null,
        progId = At(f, 2), className = At(f, 3),
        kind = At(f, 4) switch { "1" => "lightweight", "2" => "standard", "3" => "container", var k => k },
        container = At(f, 5) is { Length: > 0 } c ? c : null,
        left = Number(At(f, 6)), top = Number(At(f, 7)), width = Number(At(f, 8)), height = Number(At(f, 9)),
    };

    // name US kind US value per record; kinds s/n/b/o/i/e
    private static List<object> PropsJson(string payload)
    {
        var records = Wire.Records(payload);
        // Indexed properties: IndexedValue past the end returns "" (or 0) instead of raising (spike 25 live test), so the
        // add-in reads up to its cap. The list ends after the last non-empty List item; ItemData has the same count.
        string[] Items(string[] f) => At(f, 2).Length == 0 ? [] : At(f, 2).Split('\n');
        int? listCount = records.FirstOrDefault(f => f[0] == "List" && At(f, 1) == "i") is { } list
            ? Array.FindLastIndex(Items(list), s => s.Length > 0) + 1 : null;
        string[] Trim(string[] f)
        {
            var items = Items(f);
            int n = listCount ?? Array.FindLastIndex(items, s => s.Length > 0) + 1;
            return items.Take(Math.Min(n, items.Length)).ToArray();
        }
        return records.Select(f => (object)new
        {
            name = f[0],
            type = At(f, 1) switch { "s" => "string", "n" => "number", "b" => "boolean", "o" => "object", "i" => "indexed", "e" => "error", var k => k },
            value = At(f, 1) == "i" ? null : At(f, 2),
            items = At(f, 1) == "i" ? Trim(f) : null,
        }).ToList();
    }

    private static string At(string[] f, int i) => i < f.Length ? f[i] : "";
    private static double? Number(string s) => double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : null;
}
