using System.Globalization;
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
    /// <summary>The workspace project map (WorkspaceTools.CurrentMap); null when there are no roots.</summary>
    public Func<Core.Files.ProjectMap?> MapProvider { get; init; } = () => null;

    public IReadOnlyList<Tool> All() =>
        [.. new IReadOnlyList<Tool>[] { ReadTools(), EditTools(), LaunchTools(), DebugTools(), InspectTools(), WatchdogTools(), RecoveryTools(), AppTools(), ExeTools(), BuildTools(), SyncTools(), FormTools(), RefTools(), ProjectTools(), CheckpointTools(), OwnershipTools() }
            .SelectMany(x => x).Select(t => WriteTools.Contains(t.Name) ? t.Guarded(a => RequireNotOwnedByOthers(t.Name, a)) : t)
            .Select(t => t.Around(CrashesSinceLastCall, WithCrashes))];   // crash notices ride along with the next result (D2)

    // ---------- IDE ownership (PLAN Phase 7 A9) ----------

    /// <summary>Where owner files live; tests use their own folder.</summary>
    public string? OwnersFolder { get; init; }

    /// <summary>
    /// Tools that change, run or save something in an IDE. They refuse an IDE another live session owns. Reading tools
    /// (code_read, ide_status, form_controls, screenshots of the running program, ...) work on every IDE.
    /// </summary>
    internal static readonly HashSet<string> WriteTools = new(StringComparer.Ordinal)
    {
        "code_replace_lines", "code_insert", "proc_replace", "project_save", "project_open", "project_close", "project_add_component", "project_settings", "project_rollback",
        "ide_dialog_respond", "ide_kill", "ide_quit", "run_start", "run_break", "run_continue", "run_end", "step_into", "step_over", "step_out", "run_to_cursor",
        "bp_toggle", "bp_clear_all", "debug_eval", "debug_immediate", "debug_locals", "syntax_check", "recovery_restore", "recovery_discard",
        "app_click", "app_set_text", "app_mouse", "app_scroll", "app_select", "app_keys", "build_in_ide", "sync_take_disk", "sync_take_ide",
        "form_set_property", "form_set_properties", "form_layout", "form_add_control", "form_remove_control", "form_screenshot",
        "refs_add", "refs_remove", "refs_add_component", "refs_fix_broken",
    };

    /// <summary>VB6MCP_EVAL=open: debug_eval runs expressions with side effects (flagged, audited); by default it refuses them (PLAN.md E5).</summary>
    public bool EvalOpen { get; init; }

    /// <summary>Whether the Claude Code hook guards debug_eval, so an allow_side_effect call was approved by the user (tests replace it).</summary>
    public Func<bool> EvalGuarded { get; init; } = () => Hook.GuardsEval();

    private IdeOwner MyOwner() => IdeOwnership.Me(Mcp.McpServer.ClientName);

    /// <summary>The IDE's live owner, or null.</summary>
    private IdeOwner? OwnerOf(int pid) => IdeOwnership.Read(pid, OwnersFolder, OwnerAlive);

    /// <summary>Whether an owner's server process still runs (tests replace it).</summary>
    public Func<IdeOwner, bool>? OwnerAlive { get; init; }

    private void RequireNotOwnedByOthers(string tool, System.Text.Json.Nodes.JsonObject a)
    {
        if (tool == "project_settings" && a["set"] is null) return;            // reading settings
        if (tool == "project_rollback" && a["path"] is not null) return;      // checks the owner of the IDE the project is open in, if any
        InstanceInfo inst;
        try { inst = ResolveInstance(a); } catch (ToolException) { return; }   // the tool itself reports a missing or ambiguous IDE
        if (OwnerOf(inst.Pid) is { } owner && !IdeOwnership.IsMe(owner))
            throw new ToolException($"VB6 {inst.Pid} belongs to another session ({owner.Client}, server pid {owner.ServerPid}, since {owner.Since.ToLocalTime():HH:mm}), " +
                $"so {tool} won't change it. Start your own IDE with ide_launch (through the watchdog, no UAC prompt), or pass the pid of one you own. Reading tools still work on it.");
    }

    private IReadOnlyList<Tool> OwnershipTools() =>
    [
        new("ide_claim",
            "Make an IDE yours, so other sessions' write tools leave it alone (ide_launch claims the IDEs it starts). Works on an " +
            "unowned IDE or one whose owner session has ended; an IDE another live session owns needs force=true (tell the user).",
            Schema.Object(("pid", "integer", "The IDE to claim.", true), ("force", "boolean", "Take it from another live session.", false)),
            (a, _) => Task.FromResult(Claim(a))),
        new("ide_release",
            "Give up ownership of an IDE you claimed or started, so other sessions can change it again.",
            Schema.Object(("pid", "integer", "The IDE to release.", true)),
            (a, _) => Task.FromResult(Release(a))),
    ];

    private ToolResult Claim(System.Text.Json.Nodes.JsonObject a)
    {
        int pid = a.Int("pid") ?? throw new ToolException("missing required argument 'pid'");
        if (_env.Scan().Live.All(i => i.Pid != pid)) throw new ToolException($"no VB6 IDE with the add-in has pid {pid}");
        if (OwnerOf(pid) is { } owner && !IdeOwnership.IsMe(owner))
        {
            if (a.Bool("force") != true)
                throw new ToolException($"VB6 {pid} belongs to another live session ({owner.Client}, server pid {owner.ServerPid}); pass force=true only if the user agrees");
            Audit.Write("ide_claim", pid, $"taken from {owner.Client} (server pid {owner.ServerPid})");
        }
        IdeOwnership.Claim(pid, MyOwner(), OwnersFolder);
        return ToolResult.Ok(new { status = "claimed", pid, owner = Mcp.McpServer.ClientName });
    }

    private ToolResult Release(System.Text.Json.Nodes.JsonObject a)
    {
        int pid = a.Int("pid") ?? throw new ToolException("missing required argument 'pid'");
        if (OwnerOf(pid) is not { } owner) return ToolResult.Ok(new { status = "not-owned", pid });
        if (!IdeOwnership.IsMe(owner)) throw new ToolException($"VB6 {pid} belongs to another session ({owner.Client}); only it can release it");
        IdeOwnership.Release(pid, OwnersFolder);
        return ToolResult.Ok(new { status = "released", pid });
    }

    /// <summary>"you", the other session's name, or null for an unowned IDE (for ide_list).</summary>
    private object? OwnerJson(int pid) => OwnerOf(pid) is not { } o ? null
        : IdeOwnership.IsMe(o) ? "you" : new { client = o.Client, serverPid = o.ServerPid, since = o.Since.ToLocalTime().ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) };

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
            "Optional start_line/count (1-based) read part of the module. Without a range, at most 2000 lines are returned; " +
            "nextStartLine says where to read on.",
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
            "Before opening, the project's controls are checked: an unregistered one would make VB6 load an empty project, so the open " +
            "is refused (force=true opens anyway); unregistered references are warnings. After opening, status opened-incomplete " +
            "means VB6 loaded fewer components than the .vbp lists (with the load logs). To start VB6 itself use ide_launch.",
            Schema.Object(("path", "string", "Full path of the .vbp or .vbg file.", true),
                          ("mode", "string", "add or replace; required when the IDE already has a project.", false),
                          ("force", "boolean", "Open even though controls of the project aren't registered.", false),
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
            ides.Add(new { pid = inst.Pid, owner = OwnerJson(inst.Pid), elevated = Core.Vb6Install.IsElevated(inst.Pid), mode = Mode(w.Mode), responds = w.Responds, title = w.Title, dialogs = Dialogs(w), addin = inst.AddinVersion, projects, addinError });
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
        }).With("markedChangedByRun", MarkedJson(await MarkedChangedByRunAsync(client.Instance)));
    }

    private async Task<ToolResult> ReadAsync(System.Text.Json.Nodes.JsonObject a)
    {
        var (client, project) = await ResolveProjectAsync(a);
        var component = a.Req("component");
        // Always read the whole module: the hash covers the module, so write tools can detect any change
        var code = await client.ReadAsync(project, component);
        var page = Paging.Page(code.Text, a.Int("start_line"), a.Int("count"));   // capped without a range (improvement notes 13)
        return ToolResult.Ok(new
        {
            source = "ide", project, component, dirty = code.IsDirty, countOfLines = code.CountOfLines, hash = Core.VbCode.Hash(code.Text), startLine = page.StartLine,
            truncated = page.Truncated ? true : (bool?)null, nextStartLine = page.NextStartLine, note = Paging.Note(page, code.CountOfLines), text = page.Text,
        });
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

        // Preflight (improvement notes 8, 17): an unregistered control makes VB6 load an empty project without a word
        // (AddFromFile answers its "Continue Loading Project?" with No). Refuse that unless forced; warn about references.
        Core.Files.VbpFile? vbp = ext == ".vbp" ? Core.Files.VbpFile.Load(path) : null;
        var missing = vbp is null ? [] : Core.Files.ProjectPreflight.Missing(vbp, Core.Files.TypeLibRegistry.All());
        var missingControls = missing.Where(m => m.Kind == "control").ToList();
        if (missingControls.Count > 0 && a.Bool("force") != true)
            throw new ToolException($"{missingControls.Count} control(s) of this project aren't registered on this machine: " +
                string.Join("; ", missingControls.Select(m => Describe(m) + Hints(m, vbp!))) + ". VB6 would load the project without its components (the " +
                "\"Continue Loading Project?\" prompt resolves as No). Register them (com_search, com_register; com_diagnose project=... " +
                "checks everything), or pass force=true to open anyway.");
        var warnings = missing.Where(m => m.Kind == "reference").Select(m => $"reference {Describe(m)}: VB6 will show it as MISSING{Hints(m, vbp!)}").ToList();
        var started = DateTime.Now;

        await client.OpenAsync(path, replace);

        // The open runs deferred in the IDE; watch for the project or a dialog.
        var deadline = DateTime.UtcNow + OpenWait;
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(300, ct);
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count > 0)
                return ToolResult.Ok(new { status = "dialog", path, note = "VB6 is showing a dialog while opening; it was not answered. Load errors come once per form: ide_dialog_respond repeat=true clicks through them and says what is missing", dialogs = Dialogs(w) });
            IReadOnlyList<ProjectInfo> now;
            try { now = await client.ProjectsAsync(); }
            catch (AddinBusyException) { continue; }        // IDE busy loading
            if (now.FirstOrDefault(p => SamePath(p.FileName, path)) is { } opened)
            {
                // "Opened" isn't "loaded": compare with the .vbp's members (a production project loaded 0 of 172 and said nothing)
                object? incomplete = null;
                if (vbp is not null)
                {
                    IReadOnlyList<ComponentInfo>? comps = null;   // best effort: the open itself succeeded
                    try { comps = await client.ComponentsAsync(opened.Name); } catch (Exception ex) when (ex is AddinBusyException or AddinErrorException) { }
                    int expected = Core.Files.ProjectPreflight.ComponentCount(vbp);
                    if (comps is not null && comps.Count < expected)
                    {
                        var loadedFiles = comps.Select(c => c.FileName).Where(n => n.Length > 0).ToHashSet(StringComparer.OrdinalIgnoreCase);
                        incomplete = new
                        {
                            loaded = comps.Count, expected,
                            notLoaded = vbp.Members.Where(m => !m.Kind.Equals("RelatedDoc", StringComparison.OrdinalIgnoreCase) && !loadedFiles.Contains(m.FullPath))
                                                   .Select(m => m.FullPath).Take(20),
                            loadLogs = LoadLogsSince(vbp, started),
                        };
                    }
                }
                var after = await SettleAsync(inst.Pid, ct);   // e.g. SourceSafe's "Add this project?" comes after the open (notes 31)
                return ToolResult.Ok(new
                {
                    status = incomplete is null ? "opened" : "opened-incomplete", path, project = opened.Name, type = ProjectType(opened.Type), projects = now.Select(p => p.Name),
                    incomplete, warnings = warnings.Count > 0 ? warnings : null,
                    dialogs = after.Dialogs.Count > 0 ? Dialogs(after) : null,
                    note = incomplete is not null ? "VB6 loaded only part of the project; see loadLogs. Don't save it: the .vbp would lose the missing members. Close it and fix the cause first."
                         : after.Dialogs.Count > 0 ? "opened, and VB6 shows a dialog after it (e.g. source control): answer it with ide_dialog_respond before the next call" : null,
                });
            }
        }
        // The add-in logs why a deferred open failed (e.g. 60082 "A project with the name 'AppA' is already loaded")
        var why = AddinLogSince(inst.Pid, started, "open failed");
        return ToolResult.Ok(new { status = why is null ? "pending" : "failed", path, error = why, note = why is null ? $"not visible after {OpenWait.TotalSeconds:0} s; check ide_status" : null });
    }

    /// <summary>
    /// Where the missing library might be (PLAN.md C3, improvement notes 19): files on disk that carry its LIBID (searched
    /// by the file name the .vbp gives, in the project folder, SysWOW64 and the roots) and registered libraries with a
    /// similar name, which are other libraries (another version or edition), not substitutes.
    /// </summary>
    private string Hints(Core.Files.MissingDependency m, Core.Files.VbpFile vbp)
    {
        var parts = m.Raw.Split('#');
        string fileName = m.Kind == "control" ? m.Name : parts.Length >= 4 ? Path.GetFileName(parts[3]) : "";
        var folders = new List<(string, bool)> { (Path.GetDirectoryName(vbp.Path)!, true), (Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SysWOW64"), false) };
        if (m.Kind == "reference" && parts.Length >= 4 && Path.GetDirectoryName(parts[3]) is { Length: > 0 } hintDir) folders.Add((hintDir, false));
        folders.AddRange(Roots.Select(r => (r, true)));
        var files = Core.Files.LibraryLookup.FilesWithLib(m.LibId, [fileName], folders);
        var similar = Core.Files.LibraryLookup.Similar(files.Select(x => x.Name).Append(fileName).Append(m.Name), Core.Files.TypeLibRegistry.All(), m.LibId);
        var text = "";
        if (files.Count > 0) text += $" [on disk with this library, not registered: {string.Join(", ", files.Take(3).Select(x => $"{x.Path} (v{x.Major}.{x.Minor})"))}: com_register it]";
        if (similar.Count > 0) text += $" [registered libraries with a similar name, but a different library: {string.Join(", ", similar.Take(3).Select(t => $"{t.Description} v{t.Major}.{t.Minor} ({Path.GetFileName(t.Win32Path ?? "")})"))}]";
        return text;
    }

    private static string Describe(Core.Files.MissingDependency m)
        => $"{m.Name} ({m.LibId.ToString("B").ToUpperInvariant()} v{m.Major}" + (m.OtherVersions.Count > 0 ? $"; registered: v{string.Join(", v", m.OtherVersions)}" : "") + ")";

    /// <summary>The .log files VB6 wrote next to the project's forms while loading (e.g. "Cannot load control ...; license not found").</summary>
    private static object[] LoadLogsSince(Core.Files.VbpFile vbp, DateTime since)
    {
        var dirs = vbp.Members.Select(m => Path.GetDirectoryName(m.FullPath)).OfType<string>().Append(Path.GetDirectoryName(vbp.Path)!).Distinct(StringComparer.OrdinalIgnoreCase);
        var logs = new List<object>();
        foreach (var d in dirs)
        {
            try
            {
                foreach (var log in Directory.EnumerateFiles(d, "*.log").Where(l => File.GetLastWriteTime(l) >= since.AddSeconds(-2)).Take(10))
                    logs.Add(new { file = log, lines = Core.Vb1252.Encoding.GetString(File.ReadAllBytes(log)).Replace("\r\n", "\n").Split('\n').Where(l => l.Trim().Length > 0).Take(20) });
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        }
        return logs.ToArray();
    }

    /// <summary>The newest add-in log line of this IDE since <paramref name="since"/> containing <paramref name="marker"/>, or null.</summary>
    private static string? AddinLogSince(int pid, DateTime since, string marker)
    {
        var file = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp", "addin.log");
        try
        {
            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            if (fs.Length > 256 * 1024) fs.Seek(-256 * 1024, SeekOrigin.End);
            var tail = new StreamReader(fs).ReadToEnd().Split('\n');
            string tag = $"[{pid}] ";
            foreach (var line in tail.Reverse())
            {
                if (!line.Contains(tag) || !line.Contains(marker, StringComparison.OrdinalIgnoreCase)) continue;
                if (line.Length >= 19 && DateTime.TryParse(line[..19], System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var t) && t < since.AddSeconds(-2)) return null;
                return line[(line.IndexOf(tag, StringComparison.Ordinal) + tag.Length)..].Trim();
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        return null;
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
        if (live.Count > 1 && live.Where(i => OwnerOf(i.Pid) is { } o && IdeOwnership.IsMe(o)).ToList() is { Count: 1 } mine)
            return mine[0];   // several IDEs, exactly one of them this session's: that one (A9)
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
