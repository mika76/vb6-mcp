using System.Collections.Concurrent;
using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Forms a run marked as changed (spike 32). Running the program in the IDE marks the forms it loaded as dirty although
/// nothing changed: saving them writes identical bytes. Both T12 rerun agents saved, ran the program to test, and left
/// four "unsaved" forms, so a build from disk refused. The server remembers what was dirty when a run started and, in
/// design mode afterwards, names the components that became dirty only during the run and whose code still equals the
/// file (no designer change is possible in run mode). It never saves them itself (hard rule 4).
/// </summary>
public sealed partial class IdeTools
{
    private readonly ConcurrentDictionary<int, HashSet<string>> _dirtyBeforeRun = new();

    /// <summary>Call just before a run starts (run_start, a step or run-to-cursor from design mode, syntax_check).</summary>
    private async Task RememberDirtyBeforeRunAsync(InstanceInfo inst)
    {
        try
        {
            var client = _env.Client(inst);
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in await client.ProjectsAsync())
                foreach (var c in await client.ComponentsAsync(p.Name))
                    if (c.IsDirty) set.Add($"{p.Name}.{c.Name}");
            _dirtyBeforeRun[inst.Pid] = set;
        }
        catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException) { _dirtyBeforeRun.TryRemove(inst.Pid, out _); }
    }

    /// <summary>
    /// Components the last run marked changed: dirty now, not dirty before the run, code equal to the file. Null when no run
    /// was started through this server, the IDE isn't in design mode, or there are none.
    /// </summary>
    private async Task<List<(string Project, string Component, bool DesignerOpen)>?> MarkedChangedByRunAsync(InstanceInfo inst)
    {
        if (!_dirtyBeforeRun.TryGetValue(inst.Pid, out var before) || _env.ReadWindows(inst.Pid).Mode != IdeMode.Design) return null;
        var marked = new List<(string, string, bool)>();
        // Improvement notes 46: a form the run never loaded was listed, and its designer was open (form_screenshot had
        // opened it). VB probably re-initialises the UserControls of open designers when the mode changes; not confirmed
        // live yet, so the result only says which listed forms had an open designer. Designer windows are MDI children
        // captioned "frmMain (Form)", without the project (spike 28).
        var designers = DesignerNames(IdeWindows.Descendants(inst.Pid).Select(w => w.Caption));
        try
        {
            var client = _env.Client(inst);
            foreach (var p in await client.ProjectsAsync())
                foreach (var c in await client.ComponentsAsync(p.Name))
                {
                    if (!c.IsDirty || before.Contains($"{p.Name}.{c.Name}") || c.FileName.Length == 0 || !File.Exists(c.FileName)) continue;
                    string disk;
                    try { disk = DiskCode(c.FileName); } catch (Mcp.ToolException) { continue; }
                    if ((await client.ReadAsync(p.Name, c.Name)).Text == disk) marked.Add((p.Name, c.Name, designers.Contains(c.Name)));
                }
        }
        catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException) { return null; }
        return marked.Count > 0 ? marked : null;
    }

    /// <summary>
    /// Component names of open designer windows, not code windows ("P - frmMain (Code)"). A designer is captioned
    /// "P - frmMain (Form)" when it isn't maximized and "frmMain (Form)" when it is (live re-test 2026-10-19; spike 28
    /// saw only the second form).
    /// </summary>
    public static HashSet<string> DesignerNames(IEnumerable<string> captions)
        => captions.Where(c => c.EndsWith(')') && c.Contains(" (") && !c.EndsWith("(Code)", StringComparison.OrdinalIgnoreCase))
                   .Select(c => c.Contains(" - ", StringComparison.Ordinal) ? c[(c.LastIndexOf(" - ", StringComparison.Ordinal) + 3)..] : c)
                   .Select(c => c[..c.LastIndexOf(" (", StringComparison.Ordinal)])
                   .Where(n => n.Length > 0 && !n.Contains(' '))
                   .ToHashSet(StringComparer.OrdinalIgnoreCase);

    /// <summary>The result part for MarkedChangedByRunAsync: names, why, and the call that clears it.</summary>
    private static JsonNode? MarkedJson(List<(string Project, string Component, bool DesignerOpen)>? marked)
    {
        if (marked is null) return null;
        var open = marked.Where(m => m.DesignerOpen).Select(m => (JsonNode)$"{m.Project}.{m.Component}").ToArray();
        return new JsonObject
        {
            ["components"] = new JsonArray(marked.Select(m => (JsonNode)$"{m.Project}.{m.Component}").ToArray()),
            ["designerOpen"] = open.Length > 0 ? new JsonArray(open) : null,
            ["note"] = "VB marks forms the program loaded as changed after a run, though nothing changed (saving writes the same bytes, spike 32)" +
                       (open.Length > 0 ? "; forms whose designer is open (designerOpen) can be marked even if the program didn't load them" : "") + ". " +
                       "If they were saved before the run, save them again so a build from disk sees no unsaved changes: project_save components=[" +
                       string.Join(", ", marked.Select(m => m.Component)) + "] (only these files, not the .vbp)",
        };
    }
}
