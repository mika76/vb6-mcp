using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Disk-side tools over the workspace roots (VB6MCP_ROOTS): the project map (which projects use a
/// file; PLAN §9, spike 9) and reading source files from disk. Disk is the secondary channel
/// (principle 4): when an open IDE has the file loaded, its memory is the truth (hard rule 4), and
/// file_read says so. Documented in docs/tools.md.
/// </summary>
public sealed partial class WorkspaceTools(IIdeEnvironment env, IReadOnlyList<string> roots)
{
    /// <summary>Cached map older than this is rebuilt on the next request.</summary>
    public static readonly TimeSpan MaxMapAge = TimeSpan.FromMinutes(5);

    private readonly object _lock = new();
    private ProjectMap? _map;

    public IReadOnlyList<Tool> All() =>
    [
        new("ws_projects",
            "List the VB6 projects (.vbp) under the workspace roots (VB6MCP_ROOTS): name, type, path, member count, output file, " +
            "binary compatibility. Optional filter (substring of name or path). The map is cached; refresh=true rescans.",
            Schema.Object(("filter", "string", "Only projects whose name or path contains this text.", false),
                          ("refresh", "boolean", "Rescan the roots instead of using the cached map.", false)),
            (a, _) => Task.FromResult(Projects(a))),
        new("ws_file_projects",
            "Which projects include a source file (.bas/.cls/.frm/...)? Paths are compared normalized (relative paths, drive mappings, case), " +
            "so shared files reached via ..\\..\\Shared are found from every project.",
            Schema.Object(("path", "string", "Full path of the source file.", true),
                          ("refresh", "boolean", "Rescan the roots first.", false)),
            (a, _) => Task.FromResult(FileProjects(a))),
        new("file_read",
            "Read a VB6 source file from disk the way the IDE shows it (code only: no designer block, no hidden Attribute lines). " +
            "For files loaded in a running IDE, the IDE's memory is the truth: the result says so, and code_read gives the live text.",
            Schema.Object(("path", "string", "Full path of the .bas/.cls/.frm/.ctl/.pag/.dsr file.", true),
                          ("start_line", "integer", "First line to read (1-based).", false), ("count", "integer", "Number of lines.", false)),
            (a, _) => FileReadAsync(a)),
        new("file_replace_lines",
            "Edit a VB6 source file of a closed project on disk: replace count lines from start_line (1-based, the code as the IDE " +
            "shows it, as file_read returns it; count 0 inserts) with new_text. The header, designer block and hidden Attribute lines " +
            "stay byte-exact; the file stays Windows-1252 with CR LF. expected_hash (from file_read) guards against changes since. " +
            "If the file is open in a running IDE the edit goes to the IDE instead (its memory is the truth; it stays unsaved there).",
            Schema.Object(("path", "string", "Full path of the .bas/.cls/.frm/.ctl/.pag/.dsr file.", true), ("start_line", "integer", "First line (1-based).", true),
                          ("count", "integer", "Lines to replace (0 = insert before start_line).", true), ("new_text", "string", "Replacement text (lines separated by line breaks).", true),
                          ("expected_hash", "string", "The hash from file_read.", true), ("pid", "integer", "IDE to edit in, when the file is open in several.", false)),
            (a, _) => FileReplaceAsync(a)),
        new("ws_affected",
            "Which projects a change to (shared) source files affects, in rebuild order: the projects that include the files, then every project " +
            "that references their DLL/OCX (matched by typelib GUID and major version), dependencies first. Says where each project is open " +
            "(build those with build_in_ide; closed ones with build_make).",
            Schema.Object(("paths", "array", "Changed source files (full paths).", true),
                          ("refresh", "boolean", "Rescan the roots first.", false)),
            (a, _) => AffectedAsync(a)),
        .. ComTools(),
    ];

    private async Task<ToolResult> AffectedAsync(JsonObject a)
    {
        var paths = a.List("paths").Select(p => PathKey.Normalize(p)).ToList();
        if (paths.Count == 0) throw new ToolException("missing required argument 'paths'");
        foreach (var p in paths) RequireUnderRoots(p);
        var result = Affected.Find(Map(a["refresh"]?.GetValue<bool>() == true), paths, Affected.TypeLibOf);

        // Where the projects are open (best effort): open ones build in their IDE (/make refuses them, spike 21)
        var open = new List<(int Pid, ProjectInfo Prj)>();
        var unknown = new List<int>();
        foreach (var inst in env.Scan().Live)
        {
            try { foreach (var prj in await env.Client(inst).ProjectsAsync()) open.Add((inst.Pid, prj)); }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException) { unknown.Add(inst.Pid); }
        }
        var warnings = result.Warnings.ToList();
        if (unknown.Count > 0) warnings.Add($"VB6 {string.Join(", ", unknown)} didn't answer, so projects open there may show as closed");
        return ToolResult.Ok(new
        {
            files = paths,
            order = result.Order.Select((x, i) => new
            {
                step = i + 1, x.Project.Name, x.Project.Type, path = x.Project.Path, output = x.Project.OutputPath,
                compatibility = x.Project.CompatibleMode switch { 0 => "none", 1 => "project", 2 => "binary", var m => m.ToString() },
                includes = x.Includes,
                uses = x.Uses.Select(u => u.Name),
                typeLib = x.TypeLib is { } t ? new { libId = t.LibId.ToString("B").ToUpperInvariant(), version = $"{t.Major}.{t.Minor}", t.Name } : null,
                openIn = open.Where(o => o.Prj.FileName.Length > 0 && PathKey.Same(o.Prj.FileName, x.Project.Path)).Select(o => new { pid = o.Pid, project = o.Prj.Name, dirty = o.Prj.IsDirty }),
            }),
            warnings,
            note = "references are matched by the typelib in each DLL/OCX's current build (or compatibility file); project references inside groups (*\\A) aren't followed yet",
        });
    }

    /// <summary>The cached project map, or null without workspace roots (for IdeTools: refs_fix_broken).</summary>
    public ProjectMap? CurrentMap() => roots.Count == 0 ? null : Map(false);

    private ProjectMap Map(bool refresh)
    {
        if (roots.Count == 0)
            throw new ToolException("no workspace roots configured: set VB6MCP_ROOTS (semicolon-separated folders), e.g. via scripts\\install-server.ps1 -Roots");
        lock (_lock)
        {
            if (refresh || _map is null || DateTime.UtcNow - _map.ScannedAt > MaxMapAge) _map = ProjectMap.Build(roots);
            return _map;
        }
    }

    private ToolResult Projects(JsonObject a)
    {
        var map = Map(a["refresh"]?.GetValue<bool>() == true);
        var filter = a.Str("filter");
        var list = map.Projects
            .Where(p => filter is null || p.Name.Contains(filter, StringComparison.OrdinalIgnoreCase) || p.Path.Contains(filter, StringComparison.OrdinalIgnoreCase))
            .OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).ToList();
        const int limit = 200;
        return ToolResult.Ok(new
        {
            roots = map.Roots, scannedMs = (int)map.ScanTime.TotalMilliseconds, total = map.Projects.Count, matched = list.Count,
            projects = list.Take(limit).Select(p => new
            {
                p.Name, p.Type, path = p.Path, members = p.Members.Count, output = p.OutputPath,
                compatibility = p.CompatibleMode switch { 0 => "none", 1 => "project", 2 => "binary", var m => m.ToString() },
                compatibleFile = p.CompatibleFile,
            }),
            truncated = list.Count > limit,
            errors = map.Errors.Select(e => new { e.Path, e.Error }),
        });
    }

    private ToolResult FileProjects(JsonObject a)
    {
        var path = PathKey.Normalize(a.Req("path"));
        RequireUnderRoots(path);
        var users = Map(a["refresh"]?.GetValue<bool>() == true).ProjectsUsing(path);
        return ToolResult.Ok(new
        {
            path, shared = users.Count > 1,
            projects = users.Select(p => new { p.Name, p.Type, path = p.Path, asMember = p.Members.First(m => PathKey.Same(m.FullPath, path)).RelativePath }),
        });
    }

    private async Task<ToolResult> FileReadAsync(JsonObject a)
    {
        var path = PathKey.Normalize(a.Req("path"));
        RequireUnderRoots(path);
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);
        if (VbSourceFile.KindOf(path) == VbFileKind.Other) throw new ToolException("not a VB6 source file (.bas/.cls/.frm/.ctl/.pag/.dsr)");
        var f = VbSourceFile.Load(path);
        int start = a.Int("start_line") ?? 1;
        string text = a.Int("start_line") is null && a.Int("count") is null ? f.Code : VbCode.Slice(f.Code, start, a.Int("count") ?? int.MaxValue);
        var loaded = await FindInIdesAsync(path);
        return ToolResult.Ok(new
        {
            source = "disk", path, component = f.VbName, kind = f.Kind.ToString(), countOfLines = f.CountOfLines, startLine = start, hash = VbCode.Hash(f.Code),
            // named objects: System.Text.Json skips tuple fields
            openInIde = loaded.Select(l => new { l.pid, l.project, l.component, l.dirty }),
            warning = loaded.Any(l => l.dirty) ? "a running IDE has unsaved changes to this file; its memory is the truth (use code_read)" : null,
            text,
        });
    }

    /// <summary>
    /// file_replace_lines. A file loaded in an IDE is edited there (hard rule 4: the IDE's memory is the truth,
    /// and a disk write under a loaded component would be lost or conflict). Otherwise the code part of the file
    /// is replaced on disk, everything else byte-exact (VbSourceFile.ReplaceCode).
    /// </summary>
    private async Task<ToolResult> FileReplaceAsync(JsonObject a)
    {
        var path = PathKey.Normalize(a.Req("path"));
        RequireUnderRoots(path);
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);
        if (VbSourceFile.KindOf(path) is VbFileKind.Other) throw new ToolException("not a VB6 source file (.bas/.cls/.frm/.ctl/.pag/.dsr)");
        int start = a.Int("start_line") ?? throw new ToolException("missing required argument 'start_line'");
        int count = a.Int("count") ?? throw new ToolException("missing required argument 'count'");
        var expected = a.Req("expected_hash");
        string prepared;
        try { prepared = VbCode.PrepareForInsert(a.Str("new_text") ?? throw new ToolException("missing required argument 'new_text'")); }
        catch (ArgumentException ex) { throw new ToolException(ex.Message); }

        // Every running IDE must answer: a busy one could have the file open (and unsaved)
        var (loaded, unknown) = await FindInIdesStrictAsync(path);
        if (loaded.Count == 0 && unknown.Count > 0)
            throw new ToolException($"VB6 {string.Join(", ", unknown)} didn't answer (busy?), so it can't be ruled out that the file is open there with unsaved changes; try again when it responds");
        if (loaded.Count > 0)
        {
            var target = a.Int("pid") is int pid ? loaded.FirstOrDefault(l => l.Inst.Pid == pid) : loaded.Count == 1 ? loaded[0] : default;
            if (target.Inst is null)
                throw new ToolException($"the file is open in several IDEs ({string.Join(", ", loaded.Select(l => l.Inst.Pid))}); pass pid");
            var client = env.Client(target.Inst);
            var ide = await client.ReadAsync(target.Project, target.Component);
            if (VbCode.Hash(ide.Text) != expected)
                throw new ToolException($"the file is open in VB6 {target.Inst.Pid} as {target.Project}.{target.Component}{(target.Dirty ? " with unsaved changes" : "")}, and its text there " +
                                        "doesn't match expected_hash. Read it with code_read and edit with code_replace_lines");
            if (env.ReadWindows(target.Inst.Pid).Mode != IdeMode.Design)
                throw new ToolException($"the file is open in VB6 {target.Inst.Pid}, which isn't in design mode; edits need design mode");
            var old = VbCode.Slice(ide.Text, start, count);
            var (lines, dirty) = await client.ReplaceAsync(target.Project, target.Component, start, count, old, prepared);
            var now = await client.ReadAsync(target.Project, target.Component);
            return ToolResult.Ok(new
            {
                source = "ide", routedTo = new { pid = target.Inst.Pid, project = target.Project, component = target.Component }, path, countOfLines = lines, dirty,
                hash = VbCode.Hash(now.Text), note = "the file is open in the IDE, so the edit went there (unsaved); save with project_save",
            });
        }

        var f = VbSourceFile.Load(path);
        if (VbCode.Hash(f.Code) != expected) throw new ToolException("the file changed since it was read (hash mismatch); read it again with file_read");
        byte[] bytes;
        try { bytes = f.ReplaceCode(start, count, prepared); }
        catch (ArgumentException ex) { throw new ToolException(ex.Message); }
        var tmp = path + ".vb6mcp.tmp";
        File.WriteAllBytes(tmp, bytes);
        File.Move(tmp, path, overwrite: true);
        var after = VbSourceFile.Load(path);
        return ToolResult.Ok(new { source = "disk", path, component = after.VbName, countOfLines = after.CountOfLines, hash = VbCode.Hash(after.Code) });
    }

    /// <summary>Like FindInIdesAsync, but also lists the IDEs that couldn't be asked.</summary>
    private async Task<(List<(InstanceInfo Inst, string Project, string Component, bool Dirty)> Loaded, List<int> Unknown)> FindInIdesStrictAsync(string path)
    {
        var found = new List<(InstanceInfo, string, string, bool)>();
        var unknown = new List<int>();
        foreach (var inst in env.Scan().Live)
        {
            try
            {
                var client = env.Client(inst);
                foreach (var prj in await client.ProjectsAsync())
                    foreach (var c in await client.ComponentsAsync(prj.Name))
                        if (c.FileName.Length > 0 && PathKey.Same(c.FileName, path)) found.Add((inst, prj.Name, c.Name, c.IsDirty));
            }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException) { unknown.Add(inst.Pid); }
        }
        return (found, unknown);
    }

    /// <summary>Where a file is loaded in running IDEs (best effort; unresponsive IDEs are skipped).</summary>
    private async Task<List<(int pid, string project, string component, bool dirty)>> FindInIdesAsync(string path)
    {
        var found = new List<(int, string, string, bool)>();
        foreach (var inst in env.Scan().Live)
        {
            try
            {
                var client = env.Client(inst);
                foreach (var prj in await client.ProjectsAsync())
                    foreach (var c in await client.ComponentsAsync(prj.Name))
                        if (c.FileName.Length > 0 && PathKey.Same(c.FileName, path)) found.Add((inst.Pid, prj.Name, c.Name, c.IsDirty));
            }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException) { }
        }
        return found;
    }

    private void RequireUnderRoots(string path)
    {
        if (roots.Count == 0)
            throw new ToolException("no workspace roots configured: set VB6MCP_ROOTS (semicolon-separated folders)");
        if (!ProjectMap.IsUnderRoots(path, roots))
            throw new ToolException($"{path} is outside the workspace roots ({string.Join("; ", roots)})");
    }
}
