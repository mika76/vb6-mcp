using System.Globalization;
using System.Text.RegularExpressions;

namespace Vb6Mcp.Core.Files;

/// <summary>A control (or the form itself) in a .frm/.ctl designer block, with its position in its container's units.</summary>
/// <param name="Class">As the file writes it: VB.TextBox, XmTextCtrl.XmText.</param>
/// <param name="Name">Name, with (Index) for a control array element.</param>
/// <param name="Left">Left in the container's units (null for the form).</param>
/// <param name="Top">Top in the container's units.</param>
/// <param name="Width">Width (ClientWidth for the form).</param>
/// <param name="Height">Height (ClientHeight for the form).</param>
/// <param name="ScaleMode">The container's own ScaleMode (1 twips, 3 pixels; null when not set: twips for a form or PictureBox).</param>
/// <param name="Children">Controls inside it, in file order.</param>
public sealed record FrmControl(string Class, string Name, double? Left, double? Top, double? Width, double? Height, int? ScaleMode, List<FrmControl> Children)
{
    /// <summary>Whether the class isn't one of VB's intrinsic controls (a UserControl or another ActiveX control).</summary>
    public bool IsActiveX => !Class.StartsWith("VB.", StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// Reads the designer block of a .frm (improvement notes 42): which controls a form has, their class and position. Used to
/// name the windows of a running form, which VBIDE can't read in run mode (error 393, AGENTS.md) and which Windows only
/// knows by class (every UserControl is "ThunderRT6UserControlDC"). Read only; nothing here writes a file.
/// </summary>
public static class FrmDesign
{
    static readonly Regex BeginRx = new(@"^\s*Begin\s+([\w.]+)\s+(\w+)", RegexOptions.CultureInvariant);
    static readonly Regex PropRx = new(@"^\s*(\w+)\s*=\s*(-?[\d.]+)", RegexOptions.CultureInvariant);

    /// <summary>The form's tree (root = the form), or null when the file has no designer block.</summary>
    public static FrmControl? Parse(string text)
    {
        var stack = new Stack<(FrmControl Node, Dictionary<string, double> Props)>();
        FrmControl? root = null;
        int propertyDepth = 0;   // BeginProperty Font ... EndProperty: not controls
        foreach (var raw in text.Replace("\r\n", "\n").Split('\n'))
        {
            var line = raw.Trim();
            if (line.StartsWith("BeginProperty", StringComparison.OrdinalIgnoreCase)) { propertyDepth++; continue; }
            if (line.StartsWith("EndProperty", StringComparison.OrdinalIgnoreCase)) { propertyDepth--; continue; }
            if (propertyDepth > 0) continue;
            if (BeginRx.Match(raw) is { Success: true } b)
            {
                var node = new FrmControl(b.Groups[1].Value, b.Groups[2].Value, null, null, null, null, null, []);
                stack.Push((node, new(StringComparer.OrdinalIgnoreCase)));
                continue;
            }
            if (line.Equals("End", StringComparison.OrdinalIgnoreCase) && stack.Count > 0)
            {
                var (n, p) = stack.Pop();
                double? P(string k) => p.TryGetValue(k, out var v) ? v : null;
                var done = n with
                {
                    Name = p.TryGetValue("Index", out var ix) ? $"{n.Name}({ix.ToString(CultureInfo.InvariantCulture)})" : n.Name,
                    Left = P("Left"), Top = P("Top"), Width = P(stack.Count == 0 ? "ClientWidth" : "Width"), Height = P(stack.Count == 0 ? "ClientHeight" : "Height"),
                    ScaleMode = p.TryGetValue("ScaleMode", out var sm) ? (int)sm : null,
                };
                if (stack.Count == 0) { root = done; break; }
                stack.Peek().Node.Children.Add(done);
                continue;
            }
            if (stack.Count > 0 && PropRx.Match(line) is { Success: true } m &&
                double.TryParse(m.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var val))
                stack.Peek().Props.TryAdd(m.Groups[1].Value, val);
            if (line.StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase)) break;
        }
        return root;
    }

    /// <summary>A window to name: its handle, its parent window, and its rectangle in pixels relative to the parent window.</summary>
    public sealed record Window(long Hwnd, long Parent, int Left, int Top, int Width, int Height);

    /// <summary>
    /// Names windows by matching them to the designer tree: a window whose parent is matched to a container gets the child
    /// of that container with the same size (within 1 pixel) and the nearest position (within <paramref name="slack"/>
    /// pixels: borders of a PictureBox or Frame). Units: twips (<paramref name="twipsPerPixel"/>) unless the container's
    /// ScaleMode is 3 (pixels); other ScaleModes aren't matched. Controls moved by code at run time don't match and stay
    /// unnamed. Returns hwnd to control.
    /// </summary>
    public static Dictionary<long, FrmControl> Match(FrmControl form, long formHwnd, IReadOnlyList<Window> windows, double twipsPerPixel, int slack = 3)
    {
        var map = new Dictionary<long, FrmControl> { [formHwnd] = form };
        var scaleOf = new Dictionary<FrmControl, double?> { [form] = Scale(form.ScaleMode, twipsPerPixel) };
        var used = new HashSet<FrmControl>();
        // parents before children: walk until nothing new matches
        bool progress = true;
        while (progress)
        {
            progress = false;
            foreach (var w in windows)
            {
                if (map.ContainsKey(w.Hwnd) || !map.TryGetValue(w.Parent, out var container) || scaleOf[container] is not double unit) continue;
                FrmControl? best = null; double bestDistance = double.MaxValue;
                foreach (var c in container.Children)
                {
                    if (used.Contains(c) || c.Width is null || c.Height is null) continue;
                    if (Math.Abs(c.Width.Value / unit - w.Width) > 1 || Math.Abs(c.Height.Value / unit - w.Height) > 1) continue;
                    double dx = (c.Left ?? 0) / unit - w.Left, dy = (c.Top ?? 0) / unit - w.Top;
                    if (Math.Abs(dx) > slack || Math.Abs(dy) > slack) continue;
                    double d = dx * dx + dy * dy;
                    if (d < bestDistance) { best = c; bestDistance = d; }
                }
                if (best is null) continue;
                map[w.Hwnd] = best; used.Add(best);
                // a Frame has no ScaleMode: its children are in the units of the container it sits in
                scaleOf[best] = best.Class.Equals("VB.Frame", StringComparison.OrdinalIgnoreCase) ? unit : Scale(best.ScaleMode, twipsPerPixel);
                progress = true;
            }
        }
        map.Remove(formHwnd);
        return map;
    }

    private static double? Scale(int? scaleMode, double twipsPerPixel) => scaleMode switch
    {
        null or 1 => twipsPerPixel,
        3 => 1,
        _ => null,   // points, inches, user scales: not converted
    };
}
