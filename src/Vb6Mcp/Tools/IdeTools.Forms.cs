using System.Globalization;
using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;

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
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => FormPropertiesAsync(a)),
        new("form_set_property",
            "Set a design-time property of a control (or of the form: pass its name). The value is text, converted to the property's " +
            "type: numbers ('1200', '&H8000000F&' for colours), booleans ('True'/'False'), strings as given. Font members: Font.Name, " +
            "Font.Size, Font.Bold, Font.Italic, Font.Underline, Font.Strikethrough, Font.Charset. Pictures (Picture, Icon, MouseIcon): value '@picture:<full path>' (a file under the workspace roots), or '@picture:' to clear. Lists (List, ItemData): one item per line. Renaming (property Name) doesn't rename " +
            "the control's event procedures, just like in the IDE. Unsaved; save with project_save.",
            Schema.Object(("component", "string", ComponentDesc, true), ("control", "string", ControlDesc, true),
                          ("property", "string", "Property name, e.g. Caption, Left, BackColor, Font.Bold.", true),
                          ("value", "string", "New value as text.", true),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => FormSetPropertyAsync(a)),
        new("form_add_control",
            "Add a control to a form/UserControl in the designer. prog_id is the class name as a .frm file shows it (Begin <class>): " +
            "VB.CommandButton, VB.TextBox, VB.Label, VB.Frame, or for ActiveX controls e.g. MSComctlLib.ListView (not the registered ProgID " +
            "MSComctlLib.ListViewCtrl.2; add the component first with refs_add_component, which takes that one). Optional name, container (a Frame or " +
            "PictureBox to put it in) and position/size. Unsaved; save with project_save.",
            Schema.Object(("component", "string", ComponentDesc, true), ("prog_id", "string", "Class name as in a .frm: VB.CommandButton, MSComctlLib.ListView.", true),
                          ("name", "string", "Control name (default: VB's, e.g. Command1).", false),
                          ("container", "string", "Container control to add it to (default: the form).", false),
                          ("left", "number", "Left.", false), ("top", "number", "Top.", false), ("width", "number", "Width.", false), ("height", "number", "Height.", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => FormAddControlAsync(a)),
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
        try { return await client.DesignAsync(op, args); }
        catch (AddinErrorException ex)
        {
            // our own add-in errors carry vbObjectError + n ("-2147221104: not a number: abc"); the number means nothing to the agent
            throw new ToolException(System.Text.RegularExpressions.Regex.Replace(ex.AddinMessage, @"^-21472\d{5}: ", ""));
        }
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
        return ToolResult.Ok(new { source = "ide", project, component, control, properties = PropsJson(payload) });
    }

    private async Task<ToolResult> FormSetPropertyAsync(JsonObject a)
    {
        var (client, project, component) = await DesignTargetAsync(a);
        var control = a.Req("control");
        var property = a.Req("property");
        var value = a.Str("value") ?? throw new ToolException("missing required argument 'value'");
        if (value.StartsWith("@picture:", StringComparison.Ordinal) && value.Length > 9)
        {
            // the picture file is read by the elevated IDE: same roots rule as every other path (hard rule 8); the add-in checks again
            var pic = PathKey.Normalize(value[9..]);
            if (Roots.Count == 0 || !ProjectMap.IsUnderRoots(pic, Roots)) throw new ToolException($"{pic} is outside the workspace roots");
            if (!File.Exists(pic)) throw new ToolException("picture not found: " + pic);
            value = "@picture:" + pic;
        }
        var (dirty, payload) = await DesignAsync(client, "SET", project, component, control, property, value);
        return ToolResult.Ok(new { source = "ide", project, component, control, dirty, now = PropsJson(payload), note = "unsaved in the IDE; save with project_save" });
    }

    private async Task<ToolResult> FormAddControlAsync(JsonObject a)
    {
        var (client, project, component) = await DesignTargetAsync(a);
        string Num(string k) => a[k] is JsonValue v && v.TryGetValue(out double d) ? d.ToString(CultureInfo.InvariantCulture) : "";
        var geometry = string.Join(",", Num("left"), Num("top"), Num("width"), Num("height")).TrimEnd(',');
        var (dirty, payload) = await DesignAsync(client, "ADD", project, component, a.Req("prog_id"), a.Str("name") ?? "", a.Str("container") ?? "", geometry);
        return ToolResult.Ok(new { source = "ide", project, component, dirty, added = ControlJson(payload.Split(Wire.US)), note = "unsaved in the IDE; save with project_save" });
    }

    private async Task<ToolResult> FormRemoveControlAsync(JsonObject a)
    {
        if (a.Bool("confirm") != true) throw new ToolException("removing a control needs confirm=true");
        var (client, project, component) = await DesignTargetAsync(a);
        var control = a.Req("control");
        var (dirty, _) = await DesignAsync(client, "REMOVE", project, component, control);
        return ToolResult.Ok(new { source = "ide", project, component, removed = control, dirty, note = "unsaved in the IDE; event procedures of the control stay in the code" });
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
