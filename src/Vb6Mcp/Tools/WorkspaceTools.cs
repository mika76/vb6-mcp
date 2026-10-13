using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>
/// Disk-side tools over the workspace roots (VB6MCP_ROOTS): the project map (which projects use a
/// file; PLAN §9, spike 9) and reading source files from disk. Disk is the secondary channel
/// (principle 4): when an open IDE has the file loaded, its memory is the truth (hard rule 4), and
/// file_read says so. Documented in docs/tools.md.
/// </summary>
public sealed class WorkspaceTools(IIdeEnvironment env, IReadOnlyList<string> roots)
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
    ];

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
            source = "disk", path, component = f.VbName, kind = f.Kind.ToString(), countOfLines = f.CountOfLines, startLine = start,
            // named objects: System.Text.Json skips tuple fields
            openInIde = loaded.Select(l => new { l.pid, l.project, l.component, l.dirty }),
            warning = loaded.Any(l => l.dirty) ? "a running IDE has unsaved changes to this file; its memory is the truth (use code_read)" : null,
            text,
        });
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
