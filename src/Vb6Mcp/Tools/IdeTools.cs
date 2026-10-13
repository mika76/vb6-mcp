using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 1 tools: find IDEs, read their state, read live code and its outline.
/// Every code result says where it came from (source "ide") and whether the component is dirty
/// (AGENTS.md hard rule 4). Documented in docs/tools.md.
/// </summary>
public sealed partial class IdeTools(IIdeEnvironment env)
{
    private const string PidDesc = "VB6 process id. Optional when exactly one IDE is running (see ide_list).";
    private const string ProjectDesc = "Project name (VBProject.Name). Optional when the IDE has exactly one project.";
    private const string ComponentDesc = "Component name (VBComponent.Name), e.g. Form1 or Module1.";

    // Primary-constructor parameters are only in scope in this part of the partial class
    private readonly IIdeEnvironment _env = env;

    /// <summary>Workspace roots (VB6MCP_ROOTS): exe_run only starts programs under them (hard rule 8).</summary>
    public IReadOnlyList<string> Roots { get; init; } = [];

    public IReadOnlyList<Tool> All() => [.. ReadTools(), .. EditTools(), .. LaunchTools(), .. DebugTools(), .. InspectTools(), .. WatchdogTools(), .. RecoveryTools(), .. AppTools(), .. ExeTools()];

    private IReadOnlyList<Tool> ReadTools() =>
    [
        new("ide_list",
            "List running VB6 IDEs that have the vb6-mcp add-in: pid, mode (design/run/break), responsiveness, " +
            "open dialogs, projects. Also reports IDEs that crashed or were killed (left-over discovery files) and cleans those up.",
            Schema.Object(), (_, _) => ListAsync()),
        new("ide_status",
            "State of one IDE without needing its UI thread: title, mode, responds, hung flag, open dialogs (text and buttons), " +
            "and whether the add-in answers.",
            Schema.Object(("pid", "integer", PidDesc, false)), (a, _) => StatusAsync(a)),
        new("project_components",
            "Components of a project in the IDE: name, type, unsaved (dirty) flag, line count, file name.",
            Schema.Object(("pid", "integer", PidDesc, false), ("project", "string", ProjectDesc, false)), (a, _) => ComponentsAsync(a)),
        new("code_read",
            "Read a component's code as the IDE holds it right now, including unsaved edits. Works in design, run and break mode. " +
            "Optional start_line/count (1-based) read part of the module.",
            Schema.Object(("component", "string", ComponentDesc, true), ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false),
                          ("start_line", "integer", "First line to read (1-based).", false), ("count", "integer", "Number of lines.", false)),
            (a, _) => ReadAsync(a)),
        new("symbol_outline",
            "Outline of a component from the IDE's own parser (CodeModule.Members): procedures, properties, events, variables, constants " +
            "with scope and line spans. Cheap way to navigate large modules without reading them.",
            Schema.Object(("component", "string", ComponentDesc, true), ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => OutlineAsync(a)),
        new("symbol_find",
            "Find where a name is declared in the projects open in running IDEs (procedures, properties, variables, constants, events), " +
            "from the IDE's own parser. Exact name match (case-insensitive) unless contains=true. Use pid/project to narrow the search.",
            Schema.Object(("name", "string", "Name to look for.", true), ("contains", "boolean", "Match names containing the text.", false),
                          ("project", "string", "Only this project.", false), ("pid", "integer", "Only this IDE.", false)),
            (a, _) => FindAsync(a)),
        new("project_open",
            "Open a .vbp or .vbg in a running IDE. If the IDE already has projects, pass mode 'add' (adds to a group) or " +
            "'replace' (closes them; refused while any has unsaved changes). Waits up to 20 s for the project to appear; " +
            "dialogs shown while opening (missing references, source control, ...) are reported, never answered. " +
            "Starting VB6 itself needs the watchdog (not available yet).",
            Schema.Object(("path", "string", "Full path of the .vbp or .vbg file.", true),
                          ("mode", "string", "add or replace; required when the IDE already has a project.", false),
                          ("pid", "integer", PidDesc, false)),
            (a, ct) => OpenAsync(a, ct)),
        new("proc_get",
            "Read one procedure or property (with the comment lines above it) from the IDE, including unsaved edits.",
            Schema.Object(("component", "string", ComponentDesc, true), ("name", "string", "Procedure or property name.", true),
                          ("kind", "string", "proc (Sub/Function, default), get, let or set.", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => ProcAsync(a)),
    ];

    // ---------- tools ----------

    private async Task<ToolResult> ListAsync()
    {
        var scan = _env.Scan();
        var ides = new List<object>();
        foreach (var inst in scan.Live)
        {
            var w = _env.ReadWindows(inst.Pid);
            object? projects = null; string? addinError = null;
            if (w.Responds)
            {
                try { projects = (await _env.Client(inst).ProjectsAsync()).Select(p => new { p.Name, type = ProjectType(p.Type), file = p.FileName, dirty = p.IsDirty }).ToList(); }
                catch (Exception ex) { addinError = ex.Message; }
            }
            else addinError = "IDE not responding; add-in not queried";
            ides.Add(new { pid = inst.Pid, elevated = Core.Vb6Install.IsElevated(inst.Pid), mode = Mode(w.Mode), responds = w.Responds, title = w.Title, dialogs = Dialogs(w), addin = inst.AddinVersion, projects, addinError });
        }
        var crashed = scan.Stale.Select(s => new { pid = s.Pid, lastStarted = s.Started }).ToList();
        int cleaned = _env.DeleteStale(scan);
        return ToolResult.Ok(new { ides, crashed, cleanedUpDiscoveryFiles = cleaned, watchdog = await WatchdogInfoAsync() });
    }

    private async Task<ToolResult> StatusAsync(System.Text.Json.Nodes.JsonObject a)
    {
        var inst = ResolveInstance(a);
        var w = _env.ReadWindows(inst.Pid);
        string addin;
        try { var p = await _env.Client(inst).PingAsync(); addin = $"answers (v{p.AddinVersion}, protocol {p.Protocol})"; }
        catch (AddinBusyException) { addin = "busy: no answer within timeout (program running a loop, or IDE blocked)"; }
        catch (Exception ex) { addin = "unreachable: " + ex.Message; }
        return ToolResult.Ok(new { pid = inst.Pid, elevated = Core.Vb6Install.IsElevated(inst.Pid), title = w.Title, mode = Mode(w.Mode), responds = w.Responds, hung = w.IsHung, dialogs = Dialogs(w), addin,
                                   watchdog = await WatchdogInfoAsync(),
                                   note = !w.Responds && w.Mode == IdeMode.Run ? "the program is busy (loop?): run_break can interrupt it through the watchdog" : null });
    }

    private async Task<ToolResult> ComponentsAsync(System.Text.Json.Nodes.JsonObject a)
    {
        var (client, project) = await ResolveProjectAsync(a);
        var comps = await client.ComponentsAsync(project);
        return ToolResult.Ok(new
        {
            project,
            components = comps.Select(c => new { c.Name, type = ComponentType(c.Type), dirty = c.IsDirty, lines = c.CountOfLines, file = c.FileName }),
        });
    }

    private async Task<ToolResult> ReadAsync(System.Text.Json.Nodes.JsonObject a)
    {
        var (client, project) = await ResolveProjectAsync(a);
        var component = a.Req("component");
        // Always read the whole module: the hash covers the module, so write tools can detect any change
        var code = await client.ReadAsync(project, component);
        int start = a.Int("start_line") ?? 1;
        string text = a.Int("start_line") is null && a.Int("count") is null ? code.Text : Core.VbCode.Slice(code.Text, start, a.Int("count") ?? int.MaxValue);
        return ToolResult.Ok(new { source = "ide", project, component, dirty = code.IsDirty, countOfLines = code.CountOfLines, hash = Core.VbCode.Hash(code.Text), startLine = start, text });
    }

    private async Task<ToolResult> OutlineAsync(System.Text.Json.Nodes.JsonObject a)
    {
        var (client, project) = await ResolveProjectAsync(a);
        var component = a.Req("component");
        var members = await client.MembersAsync(project, component);
        return ToolResult.Ok(new
        {
            source = "ide",
            project,
            component,
            members = members.Select(m => new
            {
                m.Name,
                kind = MemberType(m.MemberType),
                procKind = m.ProcKind is int k ? ProcKind(k) : null,
                scope = Scope(m.Scope),
                @static = m.IsStatic,
                line = m.ProcBody ?? m.CodeLocation,
                startLine = m.ProcStart,
                lines = m.ProcCount,
            }),
        });
    }

    private async Task<ToolResult> ProcAsync(System.Text.Json.Nodes.JsonObject a)
    {
        var (client, project) = await ResolveProjectAsync(a);
        var component = a.Req("component");
        var name = a.Req("name");
        int kind = (a.Str("kind") ?? "proc").ToLowerInvariant() switch
        {
            "proc" or "sub" or "function" => 0,
            "let" => 1,
            "set" => 2,
            "get" => 3,
            var k => throw new ToolException($"kind must be proc, get, let or set (got '{k}')"),
        };
        var p = await client.ProcAsync(project, component, name, kind);
        var (text, start, count) = Core.VbCode.TrimLeadingBlankLines(p.Text, p.ProcStart, p.ProcCount);
        var module = await client.ReadAsync(project, component);
        return ToolResult.Ok(new { source = "ide", project, component, name, kind = ProcKind(kind), dirty = module.IsDirty, hash = Core.VbCode.Hash(module.Text), startLine = start, bodyLine = p.ProcBody, lines = count, text });
    }

    /// <summary>How long project_open waits for the project (or a dialog) to show up.</summary>
    public TimeSpan OpenWait { get; init; } = TimeSpan.FromSeconds(20);

    private async Task<ToolResult> OpenAsync(System.Text.Json.Nodes.JsonObject a, CancellationToken ct)
    {
        var path = Path.GetFullPath(a.Req("path"));
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext is not (".vbp" or ".vbg")) throw new ToolException("path must be a .vbp or .vbg file");
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);

        var inst = ResolveInstance(a);
        var client = _env.Client(inst);
        var before = await client.ProjectsAsync();
        if (before.Any(p => SamePath(p.FileName, path)))
            return ToolResult.Ok(new { status = "already-open", path, projects = before.Select(p => p.Name) });

        var mode = a.Str("mode")?.ToLowerInvariant();
        if (before.Count > 0 && mode is not ("add" or "replace"))
            throw new ToolException($"the IDE already has {string.Join(", ", before.Select(p => p.Name))} open; pass mode 'add' or 'replace'");
        bool replace = mode == "replace";
        if (replace && before.FirstOrDefault(p => p.IsDirty) is { } dirty)
            throw new ToolException($"project {dirty.Name} has unsaved changes; replace would make VB6 ask to save. Save or discard it first");

        await client.OpenAsync(path, replace);

        // The open runs deferred in the IDE; watch for the project or a dialog.
        var deadline = DateTime.UtcNow + OpenWait;
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(300, ct);
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count > 0)
                return ToolResult.Ok(new { status = "dialog", path, note = "VB6 is showing a dialog while opening; it was not answered", dialogs = Dialogs(w) });
            IReadOnlyList<ProjectInfo> now;
            try { now = await client.ProjectsAsync(); }
            catch (AddinBusyException) { continue; }        // IDE busy loading
            if (now.FirstOrDefault(p => SamePath(p.FileName, path)) is { } opened)
                return ToolResult.Ok(new { status = "opened", path, project = opened.Name, type = ProjectType(opened.Type), projects = now.Select(p => p.Name) });
        }
        return ToolResult.Ok(new { status = "pending", path, note = $"not visible after {OpenWait.TotalSeconds:0} s; check ide_status" });
    }

    private static bool SamePath(string a, string b)
        => a.Length > 0 && string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);

    /// <summary>symbol_find stops after this many matches.</summary>
    private const int MaxMatches = 200;

    private async Task<ToolResult> FindAsync(System.Text.Json.Nodes.JsonObject a)
    {
        var name = a.Req("name");
        bool contains = a["contains"]?.GetValue<bool>() == true;
        var onlyProject = a.Str("project");
        var live = _env.Scan().Live;
        if (a.Int("pid") is int pid) live = live.Where(i => i.Pid == pid).ToList();
        if (live.Count == 0) throw new ToolException("no VB6 IDE with the vb6-mcp add-in is running");

        var matches = new List<object>();
        var skipped = new List<string>();
        foreach (var inst in live)
        {
            var client = _env.Client(inst);
            try
            {
                foreach (var prj in await client.ProjectsAsync())
                {
                    if (onlyProject != null && !prj.Name.Equals(onlyProject, StringComparison.OrdinalIgnoreCase)) continue;
                    foreach (var comp in await client.ComponentsAsync(prj.Name))
                    {
                        if (comp.CountOfLines < 0) continue;                      // no code module (e.g. resource file)
                        foreach (var m in await client.MembersAsync(prj.Name, comp.Name))
                        {
                            bool hit = contains ? m.Name.Contains(name, StringComparison.OrdinalIgnoreCase) : m.Name.Equals(name, StringComparison.OrdinalIgnoreCase);
                            if (!hit) continue;
                            matches.Add(new
                            {
                                pid = inst.Pid, project = prj.Name, component = comp.Name, componentType = ComponentType(comp.Type),
                                m.Name, kind = MemberType(m.MemberType), procKind = m.ProcKind is int k ? ProcKind(k) : null,
                                scope = Scope(m.Scope), line = m.ProcBody ?? m.CodeLocation,
                            });
                            if (matches.Count >= MaxMatches) return ToolResult.Ok(new { name, contains, truncated = true, matches, skipped });
                        }
                    }
                }
            }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException) { skipped.Add($"pid {inst.Pid}: {ex.Message}"); }
        }
        return ToolResult.Ok(new { name, contains, truncated = false, matches, skipped });
    }

    // ---------- resolution ----------

    private InstanceInfo ResolveInstance(System.Text.Json.Nodes.JsonObject a)
    {
        var live = _env.Scan().Live;
        if (a.Int("pid") is int pid)
            return live.FirstOrDefault(i => i.Pid == pid) ?? throw new ToolException($"no VB6 IDE with the add-in has pid {pid}. Running: {Pids(live)}");
        return live.Count switch
        {
            0 => throw new ToolException("no VB6 IDE with the vb6-mcp add-in is running (start VB6 and check Add-Ins > Add-In Manager)"),
            1 => live[0],
            _ => throw new ToolException($"several VB6 IDEs are running ({Pids(live)}); pass pid"),
        };
    }

    private async Task<(AddinClient Client, string Project)> ResolveProjectAsync(System.Text.Json.Nodes.JsonObject a)
    {
        var client = _env.Client(ResolveInstance(a));
        if (a.Str("project") is { Length: > 0 } p) return (client, p);
        var projects = await client.ProjectsAsync();
        return projects.Count switch
        {
            0 => throw new ToolException("the IDE has no project open"),
            1 => (client, projects[0].Name),
            _ => throw new ToolException($"the IDE has several projects ({string.Join(", ", projects.Select(x => x.Name))}); pass project"),
        };
    }

    private static string Pids(IEnumerable<InstanceInfo> live) => string.Join(", ", live.Select(i => i.Pid));

    // ---------- names for VBIDE enum values (docs/protocol.md) ----------

    private static object Dialogs(IdeWindowState w) => w.Dialogs.Select(d => new
    {
        caption = d.Caption, text = string.Join("\n", d.Texts), buttons = d.Buttons.Select(b => b.Caption),
        runtimeError = RuntimeError(d) is { } e ? new { number = e.Number, message = e.Message, next = "debug_error" } : null,
    });
    private static string Mode(IdeMode m) => m.ToString().ToLowerInvariant();
    public static string ProjectType(int t) => t switch { 0 => "StandardExe", 1 => "ActiveXExe", 2 => "ActiveXDll", 3 => "ActiveXControl", _ => t.ToString() };
    public static string ComponentType(int t) => t switch
    {
        1 => "StdModule", 2 => "ClassModule", 3 => "MSForm", 4 => "ResFile", 5 => "Form", 6 => "MDIForm",
        7 => "PropertyPage", 8 => "UserControl", 9 => "DocObject", 10 => "RelatedDocument", 11 => "ActiveXDesigner", _ => t.ToString(),
    };
    public static string MemberType(int t) => t switch { 1 => "method", 2 => "property", 3 => "variable", 4 => "event", 5 => "const", _ => t.ToString() };
    public static string Scope(int s) => s switch { 1 => "private", 2 => "public", 3 => "friend", _ => s.ToString() };
    public static string ProcKind(int k) => k switch { 0 => "proc", 1 => "let", 2 => "set", 3 => "get", _ => k.ToString() };
}
