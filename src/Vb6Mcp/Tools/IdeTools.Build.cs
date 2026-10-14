using System.Diagnostics;
using System.Globalization;
using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 4 builds (spike 21). VB6.EXE /make refuses a project that is open in any IDE ("already
/// open"), so: open projects are built in their IDE with VBProject.MakeCompiledFile (from memory);
/// closed projects with /make, started by an add-in as an elevated child (DLL/OCX registration
/// without a UAC prompt). Documented in docs/tools.md.
/// </summary>
public sealed partial class IdeTools
{
    /// <summary>How long a build may take.</summary>
    public TimeSpan BuildWait { get; init; } = TimeSpan.FromMinutes(5);

    private IReadOnlyList<Tool> BuildTools() =>
    [
        new("build_make",
            "Build a project the way the command line does. A closed project (path: .vbp under the workspace roots) is built with " +
            "VB6.EXE /make from the files on disk, started elevated by the add-in, so an ActiveX DLL/OCX is registered on success. " +
            "A project open in an IDE can't be built by /make (VB6 refuses it), so a saved open project is built in that IDE instead " +
            "(File > Make); one with unsaved changes is refused (use build_in_ide to compile memory). Returns status, errors, the log " +
            "(for /make) and the output file. Dialogs are reported, never answered.",
            Schema.Object(("path", "string", "Full path of a .vbp (closed or open project).", false), ("project", "string", ProjectDesc + " (an open project)", false),
                          ("pid", "integer", PidDesc, false)),
            (a, ct) => BuildMakeAsync(a, ct)),
        new("build_in_ide",
            "Compile an open project from the IDE's memory (unsaved edits included) with File > Make, to its output file. Design mode. " +
            "On a compile error, the message and location (found with a full compile) are returned. Other dialogs are reported, never answered.",
            Schema.Object(("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => BuildInIdeAsync(a, ct)),
    ];

    /// <summary>The output file a .vbp builds: ExeName32 (+ Path32), or VB's default name next to the .vbp.</summary>
    public static string? OutputFileOf(VbpFile vbp) => vbp.OutputPath ?? (Path.GetDirectoryName(vbp.Path) is { } dir
        ? Path.Combine(dir, vbp.Name + vbp.Type.ToLowerInvariant() switch { "oledll" => ".dll", "control" => ".ocx", _ => ".exe" })
        : null);

    private static object FileState(string? path) => path is null ? new { path = (string?)null, exists = false, modified = (string?)null }
        : File.Exists(path) ? new { path = (string?)path, exists = true, modified = (string?)File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) }
        : new { path = (string?)path, exists = false, modified = (string?)null };

    private async Task<ToolResult> BuildMakeAsync(JsonObject a, CancellationToken ct)
    {
        // Which project: an open one (by name or path) or a closed .vbp
        string? path = a.Str("path") is { Length: > 0 } p0 ? Path.GetFullPath(p0) : null;
        if (path != null && !path.EndsWith(".vbp", StringComparison.OrdinalIgnoreCase)) throw new ToolException("path must be a .vbp");
        InstanceInfo? openIn = null; ProjectInfo? openPrj = null;
        if (path is null)
        {
            var (client, project) = await ResolveProjectAsync(a);
            openIn = client.Instance;
            openPrj = (await client.ProjectsAsync()).FirstOrDefault(x => x.Name.Equals(project, StringComparison.OrdinalIgnoreCase))
                      ?? throw new ToolException($"project {project} isn't open");
        }
        else
            foreach (var inst in _env.Scan().Live)
            {
                try
                {
                    if ((await _env.Client(inst).ProjectsAsync()).FirstOrDefault(x => x.FileName.Length > 0 && PathKey.Same(x.FileName, path)) is { } hit)
                    { openIn = inst; openPrj = hit; break; }
                }
                catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException) { }
            }

        if (openIn != null && openPrj != null)
        {
            var client = _env.Client(openIn);
            var dirty = (await client.ComponentsAsync(openPrj.Name)).Where(c => c.IsDirty).Select(c => c.Name).ToList();
            if (openPrj.IsDirty || dirty.Count > 0)
                throw new ToolException($"{openPrj.Name} is open in VB6 (pid {openIn.Pid}) with unsaved changes ({(dirty.Count > 0 ? string.Join(", ", dirty) : "project settings")}). " +
                                        "/make can't build an open project, and building the saved files would ignore the changes: save first (project_save, " +
                                        "with the user's agreement) or use build_in_ide to compile what is in memory");
            // Saved and open: memory equals disk, and VB6 refuses /make of an open project (spike 21)
            var r = await InIdeBuildAsync(client, openPrj.Name, ct);
            return ToolResult.Ok(new { via = $"VB6 pid {openIn.Pid} (File > Make): /make can't build a project that is open in an IDE", result = r });
        }

        // Closed project: /make through an add-in (elevated child), roots checked here and in the add-in
        if (Roots.Count == 0) throw new ToolException("no workspace roots are configured (VB6MCP_ROOTS); build_make only builds projects under them");
        if (!Roots.Any(r => PathKey.Key(path!).StartsWith(PathKey.Key(r).TrimEnd('\\') + "\\", StringComparison.Ordinal)))
            throw new ToolException($"{path} isn't under the workspace roots ({string.Join("; ", Roots)})");
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);
        var launcher = _env.Scan().Live.OrderByDescending(i => Vb6Install.IsElevated(i.Pid) == true).FirstOrDefault()
                       ?? throw new ToolException("no VB6 IDE with the add-in is running; one is needed to start the elevated build (ide_launch)");
        VbpFile? vbp = null;
        try { vbp = VbpFile.Load(path!); } catch (IOException) { }
        var output = vbp is null ? null : OutputFileOf(vbp);
        var before = FileState(output);
        Audit.Write("build_make", launcher.Pid, path!);
        var (pid, logPath, vbpPath) = await _env.Client(launcher).MakeAsync(path!);

        var start = DateTime.UtcNow;
        while (DateTime.UtcNow - start < BuildWait && Alive(pid))
        {
            await Task.Delay(300, ct);
            var dialogs = IdeWindows.Dialogs(pid);
            if (dialogs.Count > 0)
                return ToolResult.Ok(new
                {
                    status = "dialog", vbp = vbpPath, buildPid = pid, log = logPath,
                    dialogs = dialogs.Select(d => new { hwnd = Hex(d.Hwnd), caption = d.Caption, text = string.Join("\n", d.Texts), buttons = d.Buttons.Select(b => b.Caption) }),
                    note = "the build is waiting on a dialog (e.g. a binary compatibility decision); it is not answered automatically. Tell the user",
                });
        }
        if (Alive(pid)) return ToolResult.Ok(new { status = "timeout", vbp = vbpPath, buildPid = pid, log = logPath, note = $"still running after {BuildWait.TotalMinutes:0} min" });

        string log = File.Exists(logPath) ? Vb1252.Encoding.GetString(File.ReadAllBytes(logPath)) : "";
        var parsed = BuildLog.Parse(log);
        return ToolResult.Ok(new
        {
            status = parsed.Succeeded ? "succeeded" : "failed", vbp = vbpPath, via = $"VB6.EXE /make (started by VB6 pid {launcher.Pid})",
            seconds = Math.Round((DateTime.UtcNow - start).TotalSeconds, 1),
            output = new { before, after = FileState(output) },
            errors = parsed.Errors.Select(e => new { kind = e.Kind, file = e.File, line = e.Line, fileLine = FileLineOf(e.File, e.Line), message = e.Message, @class = e.Class, member = e.Member }),
            log = log.Trim(), logFile = logPath,
            note = parsed.Errors.Any(e => e.Kind == "compatibility")
                ? "binary compatibility would break; /make refuses (nothing was written or registered). Make the declaration compatible, or decide with the user to break compatibility (in the IDE: Project Properties or the Make dialog)"
                : vbp?.Type is "OleDll" or "Control" or "OleExe" && parsed.Succeeded ? "the component was registered by the build (VB6 does that on success)" : null,
        });
    }

    /// <summary>The line in the file on disk for a 1-based IDE code line: hidden header and inline Attribute lines added back.</summary>
    private static int? FileLineOf(string? file, int? ideLine)
    {
        if (file is null || ideLine is not int line) return null;
        try
        {
            var f = VbSourceFile.Load(file);
            return f.Header.Count + line + f.InlineAttributes.Count(x => x.AfterCodeLine < line);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return null; }
    }

    private static bool Alive(int pid)
    {
        try { using var p = Process.GetProcessById(pid); return !p.HasExited; }
        catch (ArgumentException) { return false; }
        catch (InvalidOperationException) { return false; }
        catch (System.ComponentModel.Win32Exception) { return true; }   // elevated process: can't query, but it exists
    }

    private async Task<ToolResult> BuildInIdeAsync(JsonObject a, CancellationToken ct)
    {
        var (client, project) = await ResolveProjectAsync(a);
        return ToolResult.Ok(await InIdeBuildAsync(client, project, ct));
    }

    /// <summary>
    /// File > Make in the IDE (MakeCompiledFile, deferred in the add-in). A compile error shows no
    /// dialog: the call fails with E_FAIL (spike 21), seen through MAKESTATUS; then a full compile
    /// (syntax_check) finds the message and location.
    /// </summary>
    private async Task<object> InIdeBuildAsync(AddinClient client, string project, CancellationToken ct)
    {
        var inst = client.Instance;
        RequireMode(inst, IdeMode.Design);
        var output = await client.MakeInIdeAsync(project);        // queued; the IDE compiles from memory
        var before = FileState(output);
        DateTime? beforeTime = File.Exists(output) ? File.GetLastWriteTime(output) : null;
        Audit.Write("build_in_ide", inst.Pid, $"{project} -> {output}");

        var start = DateTime.UtcNow;
        while (DateTime.UtcNow - start < BuildWait)
        {
            await Task.Delay(400, ct);
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count > 0)
                return new
                {
                    status = "dialog", project, output = new { before, now = FileState(output) }, dialogs = Dialogs(w),
                    note = "File > Make shows a dialog (replace file, binary compatibility, ...); it is not answered automatically",
                };
            (string State, string? Message) s;
            try { s = await client.MakeStatusAsync(); } catch (AddinBusyException) { continue; }   // busy compiling
            if (s.State == "ok")
            {
                bool written = File.Exists(output) && (beforeTime is null || File.GetLastWriteTime(output) > beforeTime);
                return new
                {
                    status = written ? "succeeded" : "no-output", project, seconds = Math.Round((DateTime.UtcNow - start).TotalSeconds, 1),
                    output = new { before, after = FileState(output) },
                    note = written ? null : "MakeCompiledFile returned but the output file didn't change",
                };
            }
            if (s.State == "failed")
            {
                // Locate the error the way spike 6 does: Start With Full Compile shows the compile error box
                object? check = null;
                try { check = (await SyntaxCheckAsync(new JsonObject { ["pid"] = inst.Pid }, ct)).ToJson()["content"]?[0]?["text"]?.GetValue<string>(); }
                catch (ToolException ex) { check = ex.Message; }
                var checkNode = check is string json ? JsonNode.Parse(json) : null;
                // A binary compatibility break shows as a decision dialog in the full compile (spike 21):
                // never answered automatically (hard rule 13)
                var compat = _env.ReadWindows(inst.Pid).Dialogs.FirstOrDefault(d => d.Buttons.Any(b => b.Caption.StartsWith("Break compatibility", StringComparison.OrdinalIgnoreCase)));
                if (compat != null)
                    return new
                    {
                        status = "compatibility-decision", project, output = new { before, now = FileState(output) },
                        dialog = new { caption = compat.Caption, text = string.Join("\n", compat.Texts), buttons = compat.Buttons.Select(b => b.Caption) },
                        note = "the change breaks binary compatibility. This is the user's decision (hard rule 13): ask them. Cancel leaves everything as it was; " +
                               "or make the declaration compatible again",
                    };
                return new
                {
                    status = "failed", project, message = s.Message, output = new { before, now = FileState(output) },
                    compileCheck = (object?)checkNode ?? check,
                    note = "MakeCompiledFile reports a compile error only as a generic failure; compileCheck has the message and location from a full compile",
                };
            }
        }
        return new { status = "timeout", project, output = new { before, now = FileState(output) }, note = "no result; see %LOCALAPPDATA%\\vb6mcp\\addin.log" };
    }
}
