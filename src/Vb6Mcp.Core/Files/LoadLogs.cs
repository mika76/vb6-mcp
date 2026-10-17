using System.Text.RegularExpressions;

namespace Vb6Mcp.Core.Files;

/// <summary>One missing control class (or licence) across a project's load logs: how many controls and forms, and why.</summary>
public sealed record LoadProblem(string Class, string Reason, int Controls, IReadOnlyList<string> Forms);

/// <summary>
/// Sums up the &lt;form&gt;.log files VB6 writes when controls can't load. A missing reference or control makes VB ask
/// "Errors during load" once per form (spike 36; the owner: "because there are a lot of forms it asks the same message many
/// many times"), so the useful answer is the few classes behind all of them. Lines seen in the owner's workspace:
/// "Line 16: Class XmPanelCtrl.XmPanel of control XmPanel1 was not a loaded control class." and
/// "Line 691: Cannot load control WinHook1; license not found."
/// </summary>
public static class LoadLogs
{
    static readonly Regex NotLoaded = new(@"Class\s+(\S+)\s+of control\s+(\S+)\s+was not a loaded control class", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    static readonly Regex License = new(@"Cannot load control\s+(\S+?);\s*license not found", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static IReadOnlyList<LoadProblem> Summarize(IEnumerable<string> logFiles)
    {
        var found = new List<(string Class, string Reason, string Form)>();
        foreach (var file in logFiles.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (!File.Exists(file)) continue;
            var form = Path.GetFileNameWithoutExtension(file);
            foreach (var line in Vb1252.Encoding.GetString(File.ReadAllBytes(file)).Replace("\r\n", "\n").Split('\n'))
            {
                if (NotLoaded.Match(line) is { Success: true } m)
                    found.Add((m.Groups[1].Value, "not a loaded control class: the control's component isn't registered (or a different version is)", form));
                else if (License.Match(line) is { Success: true } l)
                    found.Add(($"control {l.Groups[1].Value.TrimEnd('.')}", "licence not found: the product's design-time licence isn't installed (spike 27)", form));
                else if (line.Trim().Length > 0)
                    found.Add((line.Trim().Length > 120 ? line.Trim()[..120] : line.Trim(), "other", form));
            }
        }
        return found.GroupBy(f => (f.Class, f.Reason))
            .Select(g => new LoadProblem(g.Key.Class, g.Key.Reason, g.Count(), g.Select(x => x.Form).Distinct(StringComparer.OrdinalIgnoreCase).ToList()))
            .OrderByDescending(p => p.Controls).ToList();
    }

    /// <summary>The .log path a VB6 "Errors during load. Refer to '…' for details" message names, or null.</summary>
    public static string? LogOf(string dialogText)
        => Regex.Match(dialogText, @"Refer to '([^']+\.log)'", RegexOptions.IgnoreCase) is { Success: true } m ? m.Groups[1].Value : null;
}
