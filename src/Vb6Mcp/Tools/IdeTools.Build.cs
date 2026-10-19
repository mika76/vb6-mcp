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
                          ("defines", "object", "Conditional compilation constants for this build (/make /d), e.g. {\"noterm\": 1}. Closed projects only.", false),
                          ("pid", "integer", PidDesc, false)),
            (a, ct) => BuildMakeAsync(a, ct)),
        new("build_in_ide",
            "Compile an open project from the IDE's memory (unsaved edits included) with File > Make, to its output file. Design mode. " +
            "On a compile error, the message and location (found with a full compile) are returned. Other dialogs are reported, never answered. " +
            "Other conditional compilation constants: project_settings set CondComp with reopen=true (VBIDE can't set them in memory), or build_make defines on a closed project.",
            Schema.Object(("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, ct) => BuildInIdeAsync(a, ct)),
    ];

    /// <summary>The output file a .vbp builds: ExeName32 (+ Path32), or VB's default name next to the .vbp.</summary>
    public static string? OutputFileOf(VbpFile vbp) => vbp.OutputPath ?? (Path.GetDirectoryName(vbp.Path) is { } dir
        ? Path.Combine(dir, vbp.Name + vbp.Type.ToLowerInvariant() switch { "oledll" => ".dll", "control" => ".ocx", _ => ".exe" })
        : null);

    /// <summary>
    /// Why a make can fail although the code compiles, as far as a non-elevated process can tell: the output folder is
    /// missing, the output file is held open (a running copy of the program; spike 26: only an exclusive open tells),
    /// or the file is read-only. Folders only administrators can write are noted, since the elevated IDE may still write there.
    /// </summary>
    public static List<string> MakeFailureCauses(string? output)
    {
        var causes = new List<string>();
        if (output is null) { causes.Add("the output file isn't known (no ExeName32 in the .vbp)"); return causes; }
        var dir = Path.GetDirectoryName(output);
        if (dir is null || !Directory.Exists(dir)) { causes.Add($"the output folder doesn't exist: {dir}"); return causes; }
        if (File.Exists(output))
        {
            if (File.GetAttributes(output).HasFlag(FileAttributes.ReadOnly)) causes.Add($"{output} is read-only");
            try { using var fs = new FileStream(output, FileMode.Open, FileAccess.Read, FileShare.None); }
            catch (IOException) { causes.Add($"{output} is in use (a running copy of the program, or a process that loaded the DLL/OCX): end it and build again"); }
            catch (UnauthorizedAccessException) { causes.Add($"{output} can't be opened from here (permissions), so whether it's in use is unknown"); }
        }
        try { var probe = Path.Combine(dir, $".vb6mcp-probe-{Environment.ProcessId}"); File.WriteAllBytes(probe, []); File.Delete(probe); }
        catch (UnauthorizedAccessException) { causes.Add($"{dir} isn't writable without elevation; the elevated IDE can usually still write there, unless it's protected"); }
        catch (IOException ex) { causes.Add($"{dir}: {ex.Message}"); }
        if (causes.Count == 0) causes.Add("no cause visible from here: the output folder is writable and the output file isn't in use");
        return causes;
    }

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
            {
                // spike 32: forms the last run marked changed (same code as on disk) are the usual cause after testing
                var marked = await MarkedChangedByRunAsync(openIn);
                var byRun = marked?.Where(m => m.Project == openPrj.Name).Select(m => m.Component).ToList() ?? [];
                throw new ToolException($"{openPrj.Name} is open in VB6 (pid {openIn.Pid}) with unsaved changes ({(dirty.Count > 0 ? string.Join(", ", dirty) : "project settings")}). " +
                                        "/make can't build an open project, and building the saved files would ignore the changes: save first (project_save, " +
                                        "with the user's agreement) or use build_in_ide to compile what is in memory" +
                    (byRun.Count > 0 ? $". {string.Join(", ", byRun)} {(byRun.Count == 1 ? "was" : "were")} only marked changed by the last run (VB does that to forms the program loaded; " +
                                       $"saving writes the same bytes, spike 32): project_save components=[{string.Join(", ", byRun)}]" : ""));
            }
            // Saved and open: memory equals disk, and VB6 refuses /make of an open project (spike 21)
            if (Defines(a) is not null)
                throw new ToolException("defines work only for /make of a closed project: VBIDE can't set conditional compilation constants of an open one. Close it (project_close), or change CondComp with project_settings");
            var r = await InIdeBuildAsync(client, openPrj.Name, ct);
            return ToolResult.Ok(new { via = $"VB6 pid {openIn.Pid} (File > Make): /make can't build a project that is open in an IDE", result = r });
        }

        // Closed project: /make through an add-in (elevated child), roots checked here and in the add-in
        if (Roots.Count == 0) throw new ToolException("no workspace roots are configured (VB6MCP_ROOTS); build_make only builds projects under them");
        if (!Roots.Any(r => PathKey.Key(path!).StartsWith(PathKey.Key(r).TrimEnd('\\') + "\\", StringComparison.Ordinal)))
            throw new ToolException($"{path} isn't under the workspace roots ({string.Join("; ", Roots)})");
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);
        // The elevated build is started by an IDE's add-in, or, with no IDE running, by the elevated watchdog (PLAN.md D1).
        // Starting VB6.EXE from here would show a UAC prompt (spike 2), so without either there is no build.
        var launcher = _env.Scan().Live.OrderByDescending(i => Vb6Install.IsElevated(i.Pid) == true).FirstOrDefault();
        var watchdog = launcher is null ? _env.FindWatchdog() : null;
        if (launcher is null && watchdog is null)
            throw new ToolException("no VB6 IDE with the add-in and no watchdog is running, and the build must be started elevated: start VB6 (ide_launch, " +
                                    "a UAC prompt unless a watchdog runs), or ask the user to start the watchdog once (vb6mcp watchdog --stay; builds then need no IDE)");
        VbpFile? vbp = null;
        try { vbp = VbpFile.Load(path!); } catch (IOException) { }
        var output = vbp is null ? null : OutputFileOf(vbp);
        var before = FileState(output);
        var defines = Defines(a);
        Audit.Write("build_make", launcher?.Pid ?? 0, path! + (defines is null ? "" : " /d " + defines) + (launcher is null ? " (via the watchdog)" : ""));
        int pid; string logPath, vbpPath;
        if (launcher is not null) (pid, logPath, vbpPath) = await _env.Client(launcher).MakeAsync(path!, defines);
        else
        {
            var r = await new AddinClient(watchdog!, TimeSpan.FromSeconds(10)).SendRawAsync("MAKE", [path!, defines ?? ""], ct);
            if (r.Status == Core.Protocol.ResponseStatus.Error)
                throw new ToolException(r.ErrorMessage.StartsWith("unknown verb", StringComparison.Ordinal)
                    ? "no VB6 IDE is running, and the running watchdog is older than this server (no MAKE verb): restart it, or start VB6 (ide_launch)"
                    : "the watchdog couldn't start the build: " + r.ErrorMessage);
            var f = r.Payload.Split(Core.Protocol.Wire.US);
            (pid, logPath, vbpPath) = (int.Parse(f[0], CultureInfo.InvariantCulture), f[1], f[2]);
        }
        string via = launcher is not null ? $"VB6.EXE /make (started by VB6 pid {launcher.Pid})" : $"VB6.EXE /make (started by the watchdog, pid {watchdog!.Pid}: no IDE was running)";

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
        // Spike 41: an ActiveX build by a VB6 that isn't elevated writes the DLL but can't register it. The log then has
        // "Error accessing the system registry" next to "Build of '…' succeeded", and nothing is registered anywhere.
        bool? builderElevated = Vb6Install.IsElevated(launcher?.Pid ?? watchdog!.Pid);
        bool registryError = parsed.Errors.Any(e => e.Message.Contains("Error accessing the system registry", StringComparison.OrdinalIgnoreCase));
        string? registryNote = !registryError ? null :
            $"VB6 wrote {output ?? "the output file"} but couldn't register it (\"Error accessing the system registry\"). Registering an ActiveX build needs administrator rights" +
            (builderElevated == false ? ", and the VB6 that ran /make isn't elevated" : "") +
            $". Register it from an elevated prompt (regsvr32 \"{output}\"), or build with VB6 running as administrator (spike 41)";
        return ToolResult.Ok(new
        {
            status = parsed.Succeeded ? "succeeded" : "failed", vbp = vbpPath, via, registered = registryError ? false : (bool?)null,
            seconds = Math.Round((DateTime.UtcNow - start).TotalSeconds, 1),
            output = new { before, after = FileState(output) },
            errors = parsed.Errors.Select(e => new
            {
                kind = e.Kind, file = e.File, line = e.Line, fileLine = FileLineOf(e.File, e.Line), message = e.Message, @class = e.Class, member = e.Member,
                details = e.Kind == "load" ? LoadLogLines(e.File) : null,
            }),
            log = log.Trim(), logFile = logPath,
            note = registryNote ?? (parsed.Errors.Any(e => e.Kind == "compatibility")
                ? "binary compatibility would break; /make refuses (nothing was written or registered). Make the declaration compatible, or decide with the user to break compatibility (in the IDE: Project Properties or the Make dialog)"
                : vbp?.Type is "OleDll" or "Control" or "OleExe" && parsed.Succeeded ? "the component was registered by the build (VB6 does that on success)" : null),
        });
    }

    /// <summary>
    /// The lines of a form's load log (VB writes <c>frmMain.log</c> next to the form when a control can't be loaded, e.g.
    /// "Line 12: Cannot load control WinHook1; license not found", spike 27). The build log only points at it.
    /// </summary>
    private static string[]? LoadLogLines(string? path)
    {
        try
        {
            if (path is null || !File.Exists(path)) return null;
            return Vb1252.Encoding.GetString(File.ReadAllBytes(path)).Replace("\r\n", "\n").Split('\n')
                .Select(l => l.Trim()).Where(l => l.Length > 0).Take(50).ToArray();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return null; }
    }

    /// <summary>
    /// build_make's defines ({"noterm": 1, "debug": 0}) as /make's /d argument ("noterm=1:debug=0"), or null. Names are
    /// VB identifiers and values integers: they end up on an elevated command line (the add-in checks the same).
    /// </summary>
    private static string? Defines(JsonObject a)
    {
        if (a["defines"] is null) return null;
        if (a["defines"] is not JsonObject o || o.Count == 0) throw new ToolException("defines must be an object of name: integer, e.g. {\"noterm\": 1}");
        var parts = new List<string>();
        foreach (var (k, v) in o)
        {
            if (k.Length == 0 || !char.IsAsciiLetter(k[0]) || !k.All(c => char.IsAsciiLetterOrDigit(c) || c == '_'))
                throw new ToolException($"defines: '{k}' isn't a valid constant name");
            if (v is not JsonValue jv || !(jv.TryGetValue(out int n) || (jv.TryGetValue(out string? t) && int.TryParse(t, out n))))
                throw new ToolException($"defines: {k} must be an integer (VB conditional constants are numbers; True is -1)");
            parts.Add($"{k}={n.ToString(CultureInfo.InvariantCulture)}");
        }
        return string.Join(":", parts);
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
        var settled = await SettleReloadsAsync(ResolveInstance(a).Pid, ct);
        return (await BuildInIdeCoreAsync(a, ct)).With("reloads", settled);
    }

    private async Task<ToolResult> BuildInIdeCoreAsync(JsonObject a, CancellationToken ct)
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
        // An empty (or not loaded) project: MakeCompiledFile then opens Project Properties and "Must have startup form or
        // Sub Main()" instead of building (improvement notes 9)
        if ((await client.ComponentsAsync(project)).Count == 0)
            throw new ToolException($"{project} has no components loaded, so there is nothing to build (VB would open Project Properties instead). " +
                                    "If the .vbp has members, the project didn't load completely: check project_open's load logs");
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
                // The full compile passed: the make step itself failed (improvement notes 20: a production project got E_FAIL from
                // MakeCompiledFile while its code compiled). Check the usual causes we can see from here.
                if (checkNode?["compiles"]?.GetValue<bool?>() == true)
                    return new
                    {
                        status = "make-failed", project, message = s.Message, output = new { before, now = FileState(output) },
                        compileCheck = checkNode, causes = MakeFailureCauses(output),
                        note = "the code compiles (full compile passed); the make step itself failed. See causes; the native compiler's own errors (C2.EXE, LINK.EXE) " +
                               "aren't reported by VBIDE: try File > Make by hand to see them, or build_make after saving and closing the project",
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
