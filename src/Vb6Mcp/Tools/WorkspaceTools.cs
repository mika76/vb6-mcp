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
            "For files loaded in a running IDE, the IDE's memory is the truth: the result says so, and code_read gives the live text. " +
            "Without a range, at most 2000 lines are returned; nextStartLine says where to read on.",
            Schema.Object(("path", "string", "Full path of the .bas/.cls/.frm/.ctl/.pag/.dsr file.", true),
                          ("start_line", "integer", "First line to read (1-based).", false), ("count", "integer", "Number of lines.", false)),
            (a, _) => FileReadAsync(a)),
        new("file_replace_lines",
            "Edit a VB6 source file of a closed project on disk: replace count lines from start_line (1-based, the code as the IDE " +
            "shows it, as file_read returns it; count 0 inserts) with new_text. The header, designer block and hidden Attribute lines " +
            "stay byte-exact; the file stays Windows-1252 with CR LF. expected_hash (from file_read) guards against changes since. " +
            "Several ranges in one call: edits; mechanical renames: find/replace (preview, then apply=true). If the file is open in a running IDE the edit goes to the IDE instead (its memory is the " +
            "truth; it stays unsaved there).",
            FindReplace.WithSchema(EditList.WithSchema(Schema.Object(("path", "string", "Full path of the .bas/.cls/.frm/.ctl/.pag/.dsr file.", true), ("start_line", "integer", "First line (1-based).", false),
                          ("count", "integer", "Lines to replace (0 = insert before start_line).", false), ("new_text", "string", "Replacement text." + EditList.CodeRule, false),
                          ("edits", "array", EditList.Description, false),
                          ("drop_var_help_ids", "boolean", "Allow the edit to drop the hidden 'Attribute x.VB_VarHelpID' of a WithEvents variable it rewrites (VB writes it again on save).", false),
                          ("expected_hash", "string", "The hash from file_read.", true), ("pid", "integer", "IDE to edit in, when the file is open in several.", false)))),
            (a, _) => FileReplaceAsync(a)),
        new("ws_affected",
            "Which projects a change to (shared) source files affects, in rebuild order: the projects that include the files, then every project " +
            "that references their DLL/OCX (matched by typelib GUID and major version), dependencies first. Says where each project is open " +
            "(build those with build_in_ide; closed ones with build_make).",
            Schema.Object(("paths", "array", "Changed source files (full paths).", true),
                          ("refresh", "boolean", "Rescan the roots first.", false)),
            (a, _) => AffectedAsync(a)),
        new("symbol_refs",
            "Every statement that uses a name (a procedure, variable, constant...) across the projects under the workspace roots: file, " +
            "line, the full statement with ' _' continuations joined (so the arguments passed are visible), and which projects include " +
            "the file. Whole-word, case-insensitive; matches in strings and comments don't count. Files a running IDE has with unsaved " +
            "changes are read from the IDE. Declarations are marked. symbol_find finds only declarations.",
            Schema.Object(("name", "string", "The identifier, e.g. Add_To_Audit.", true),
                          ("project", "string", "Only projects whose name contains this, or one .vbp path.", false),
                          ("declarations", "boolean", "Include declarations (default true).", false),
                          ("refresh", "boolean", "Rescan the roots first.", false)),
            (a, _) => RefsAsync(a)),
        FormResourcesTool(),
        ResReadTool(),
        .. ComTools(),
    ];

    /// <summary>symbol_refs stops after this many statements.</summary>
    public const int MaxRefs = 300;

    /// <summary>
    /// symbol_refs (PLAN.md C5). A disk scan of every member of the selected projects, with the IDE's text for components
    /// that are dirty in a running IDE (hard rule 4: its memory is the truth). Each shared file is scanned once.
    /// </summary>
    private async Task<ToolResult> RefsAsync(JsonObject a)
    {
        var name = a.Req("name").Trim();
        if (!VbRefs.IsIdentifier(name)) throw new ToolException($"'{name}' isn't a VB identifier (letters, digits and _, starting with a letter); pass one name, e.g. Add_To_Audit");
        var map = Map(a.Bool("refresh") == true);
        IEnumerable<VbpFile> projects = map.Projects;
        if (a.Str("project") is { Length: > 0 } spec)
            projects = spec.EndsWith(".vbp", StringComparison.OrdinalIgnoreCase)
                ? projects.Where(p => PathKey.Same(p.Path, PathKey.Normalize(spec)))
                : projects.Where(p => p.Name.Contains(spec, StringComparison.OrdinalIgnoreCase));
        var selected = projects.ToList();
        if (selected.Count == 0) throw new ToolException($"no project under the roots matches '{a.Str("project")}'");
        var files = selected.SelectMany(p => p.Members).Select(m => m.FullPath)
            .Where(p => VbSourceFile.KindOf(p) != VbFileKind.Other).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        // components with unsaved changes in a running IDE: read those from the IDE
        var dirty = new Dictionary<string, (InstanceInfo Inst, string Project, string Component)>(StringComparer.OrdinalIgnoreCase);
        var warnings = new List<string>();
        foreach (var inst in env.Scan().Live)
        {
            try
            {
                var client = env.Client(inst);
                foreach (var prj in await client.ProjectsAsync())
                    foreach (var c in await client.ComponentsAsync(prj.Name))
                        if (c.IsDirty && c.FileName.Length > 0) dirty[PathKey.Normalize(c.FileName)] = (inst, prj.Name, c.Name);
            }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException)
            {
                warnings.Add($"VB6 {inst.Pid} didn't answer: unsaved changes there aren't included (the files were read from disk)");
            }
        }

        bool withDecl = a.Bool("declarations") != false;
        var results = new List<object>();
        int total = 0; bool truncated = false;
        foreach (var path in files.OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
        {
            string code, source = "disk";
            try
            {
                if (dirty.TryGetValue(path, out var d)) { code = (await env.Client(d.Inst).ReadAsync(d.Project, d.Component)).Text; source = "ide"; }
                else if (File.Exists(path)) code = VbSourceFile.Load(path).Code;
                else continue;
            }
            catch (Exception ex) when (ex is IOException or AddinBusyException or AddinUnreachableException or AddinErrorException)
            {
                warnings.Add($"{path}: {ex.Message}");
                continue;
            }
            var refs = VbRefs.Find(code, name).Where(r => withDecl || !r.Declaration).ToList();
            if (refs.Count == 0) continue;
            if (total + refs.Count > MaxRefs) { refs = refs.Take(MaxRefs - total).ToList(); truncated = true; }
            total += refs.Count;
            results.Add(new
            {
                path, source, projects = map.ProjectsUsing(path).Select(p => p.Name),
                refs = refs.Select(r => new { line = r.Line, endLine = r.EndLine > r.Line ? r.EndLine : (int?)null, declaration = r.Declaration ? true : (bool?)null, statement = r.Statement }),
            });
            if (truncated) break;
        }
        return ToolResult.Ok(new
        {
            name, projects = selected.Count, filesScanned = files.Count, count = total, truncated = truncated ? true : (bool?)null,
            files = results, warnings = warnings.Count > 0 ? warnings : null,
            note = truncated ? $"stopped at {MaxRefs} statements; narrow it with project" : null,
        });
    }

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
        var page = Paging.Page(f.Code, a.Int("start_line"), a.Int("count"));
        var loaded = await FindInIdesAsync(path);
        return ToolResult.Ok(new
        {
            source = "disk", path, component = f.VbName, kind = f.Kind.ToString(), countOfLines = f.CountOfLines, startLine = page.StartLine, hash = VbCode.Hash(f.Code),
            truncated = page.Truncated ? true : (bool?)null, nextStartLine = page.NextStartLine, note = Paging.Note(page, f.CountOfLines),
            // named objects: System.Text.Json skips tuple fields
            openInIde = loaded.Select(l => new { l.pid, l.project, l.component, l.dirty }),
            warning = loaded.Any(l => l.dirty) ? "a running IDE has unsaved changes to this file; its memory is the truth (use code_read)" : null,
            text = page.Text,
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
        var fr = FindReplace.Args(a);
        var edits = fr is null ? EditList.From(a) : [];   // one range or a list, validated before anything changes
        // find/replace (PLAN.md A6): edits from the matches in the text the hash covers; without apply only a preview
        IReadOnlyList<FoundLine> Find(string text) => FindReplace.Compute(text, fr!.Value.Find, fr.Value.Replace, fr.Value.WholeWord, fr.Value.MatchCase, fr.Value.InStrings, fr.Value.InComments);
        IReadOnlyList<LineEdit> FoundEdits(IReadOnlyList<FoundLine> found)
            => found.Count == 0 ? throw new ToolException("no matches, nothing was changed (whole words, code only by default)") : FindReplace.Edits(found);
        var expected = a.Req("expected_hash");

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
            if (fr is { Apply: false })
                return ToolResult.Ok(new { source = "ide", routedTo = new { pid = target.Inst.Pid, project = target.Project, component = target.Component }, path, preview = FindReplace.Preview(Find(ide.Text)), hash = expected });
            if (fr is not null) edits = FoundEdits(Find(ide.Text));
            if (env.ReadWindows(target.Inst.Pid).Mode != IdeMode.Design)
                throw new ToolException($"the file is open in VB6 {target.Inst.Pid}, which isn't in design mode; edits need design mode");
            int lines = 0; bool dirty = false; int done = 0;
            foreach (var e in edits)   // bottom-up: the original numbering stays valid above each edit
            {
                var old = e.Count > 0 ? VbCode.Slice(ide.Text, e.Start, e.Count) : "";
                try { (lines, dirty) = await client.ReplaceAsync(target.Project, target.Component, e.Start, e.Count, old, e.Text); }
                catch (AddinErrorException ex) when (ex.Code == 409)
                {
                    throw new ToolException($"conflict at lines {e.Start}-{e.Start + e.Count - 1} ({ex.AddinMessage}); " +
                        (done == 0 ? "nothing was changed" : $"{done} later edit(s) were already applied in the IDE (undo there, or re-read)"));
                }
                done++;
            }
            var now = await client.ReadAsync(target.Project, target.Component);
            return ToolResult.Ok(new
            {
                source = "ide", routedTo = new { pid = target.Inst.Pid, project = target.Project, component = target.Component }, path, countOfLines = lines, dirty,
                hash = VbCode.Hash(now.Text), note = "the file is open in the IDE, so the edit went there (unsaved); save with project_save",
                codeHints = VbCode.LiteralHints(string.Join("\r\n", edits.Select(e => e.Text))) is { Count: > 0 } ideHints ? ideHints : null,
            });
        }

        var f = VbSourceFile.Load(path);
        if (VbCode.Hash(f.Code) != expected) throw new ToolException("the file changed since it was read (hash mismatch); read it again with file_read");
        if (fr is { Apply: false }) return ToolResult.Ok(new { source = "disk", path, preview = FindReplace.Preview(Find(f.Code)), hash = expected });
        if (fr is not null) edits = FoundEdits(Find(f.Code));
        // All edits are applied in memory, bottom-up, before the file is written once: a failing edit changes nothing.
        byte[] bytes = f.Bytes;
        foreach (var e in edits)
        {
            try { bytes = f.ReplaceCode(e.Start, e.Count, e.Text, a.Bool("drop_var_help_ids") == true); }
            catch (ArgumentException ex) { throw new ToolException(edits.Count == 1 ? ex.Message : $"edit at line {e.Start}: {ex.Message} (nothing was changed)"); }
            f = VbSourceFile.Parse(path, bytes);
        }
        var tmp = path + ".vb6mcp.tmp";
        File.WriteAllBytes(tmp, bytes);
        File.Move(tmp, path, overwrite: true);
        var after = VbSourceFile.Load(path);
        return ToolResult.Ok(new
        {
            source = "disk", path, component = after.VbName, countOfLines = after.CountOfLines, hash = VbCode.Hash(after.Code),
            codeHints = VbCode.LiteralHints(string.Join("\r\n", edits.Select(e => e.Text))) is { Count: > 0 } ch ? ch : null,
        });
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
