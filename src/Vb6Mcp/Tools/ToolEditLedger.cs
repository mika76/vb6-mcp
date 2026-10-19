using System.Collections.Concurrent;
using Vb6Mcp.Core;

namespace Vb6Mcp.Tools;

/// <summary>
/// The code text each IDE component had right after the last edit this server made to it (code_replace_lines,
/// code_insert, proc_replace, file_replace_lines routed to the IDE, recovery_restore). project_save and syntax_check
/// compare the IDE's text with it to find changes no tool made: stray keystrokes that reached a code pane, or the
/// user's own typing (improvement notes 50: a selection left on a compile error turned "c" into "nothe", and the next
/// project_save wrote it without a word). Reads don't update it: a code_read after a stray keystroke would hide it.
/// One ledger per server process; it starts empty, so after a restart every change counts as "not made by the tools".
/// </summary>
public sealed class ToolEditLedger
{
    private readonly ConcurrentDictionary<string, string> _text = new(StringComparer.OrdinalIgnoreCase);

    private static string Key(int pid, string project, string component) => $"{pid}|{project}|{component}";

    /// <summary>Remember <paramref name="text"/> as what the tools left in the component.</summary>
    public void Record(int pid, string project, string component, string text) => _text[Key(pid, project, component)] = text;

    /// <summary>The text the tools last left in the component, or null when no tool edited it.</summary>
    public string? Last(int pid, string project, string component) => _text.TryGetValue(Key(pid, project, component), out var t) ? t : null;

    /// <summary>The most diff lines one component reports (the rest is counted).</summary>
    public const int MaxDiffLines = 40;

    /// <summary>
    /// How <paramref name="now"/> differs from <paramref name="baseline"/> (the last tool edit, or the file on disk for a
    /// component no tool edited), or null when it doesn't, or only in the case of identifiers (VB re-cases a name in every
    /// module when any module declares it with a new spelling: that is the tools' own edit spreading, not typing).
    /// </summary>
    public static OutsideEdit? Compare(string component, string baseline, string now, string since)
    {
        if (baseline == now || string.Equals(baseline, now, StringComparison.OrdinalIgnoreCase)) return null;
        var diff = LineDiff.Unified(baseline, now, since, "ide now").Replace("\r\n", "\n").Split('\n');
        int changed = diff.Count(l => (l.StartsWith('+') || l.StartsWith('-')) && !l.StartsWith("+++") && !l.StartsWith("---"));
        var shown = string.Join("\n", diff.Take(MaxDiffLines));
        return new OutsideEdit(component, since, changed, shown, diff.Length > MaxDiffLines ? diff.Length - MaxDiffLines : null);
    }
}

/// <summary>A component whose code changed in the IDE without a tool edit (see <see cref="ToolEditLedger"/>).</summary>
/// <param name="Component">Component name.</param>
/// <param name="Since">"last vb6-mcp edit" or "disk": what the change is measured against.</param>
/// <param name="ChangedLines">Lines added plus lines removed.</param>
/// <param name="Diff">Unified diff, baseline to the IDE's text, capped at <see cref="ToolEditLedger.MaxDiffLines"/>.</param>
/// <param name="MoreDiffLines">Diff lines left out, or null.</param>
public sealed record OutsideEdit(string Component, string Since, int ChangedLines, string Diff, int? MoreDiffLines);
