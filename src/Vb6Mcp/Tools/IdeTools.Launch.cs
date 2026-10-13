using System.ComponentModel;
using System.Diagnostics;
using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>
/// ide_launch: start VB6, elevated or not (spike 16: elevation is per-machine configuration, and the
/// server should work for other users too). Documented in docs/tools.md.
/// </summary>
public sealed partial class IdeTools
{
    /// <summary>How long ide_launch waits for the add-in of the new IDE (includes time to approve UAC).</summary>
    public TimeSpan LaunchWait { get; init; } = TimeSpan.FromSeconds(90);

    private IReadOnlyList<Tool> LaunchTools() =>
    [
        new("ide_launch",
            "Start a new VB6 IDE, optionally opening a .vbp/.vbg (which skips VB6's New Project chooser). elevation: 'auto' (default: " +
            "as VB6 is configured on this machine), 'elevated' (as administrator) or 'normal' (not elevated). Starting VB6 elevated " +
            "shows a UAC prompt the user must approve. Waits until the vb6-mcp add-in answers and reports startup dialogs without answering them.",
            Schema.Object(("project", "string", "Full path of a .vbp or .vbg to open.", false),
                          ("elevation", "string", "auto (default), elevated or normal.", false)),
            (a, ct) => LaunchAsync(a, ct)),
    ];

    private async Task<ToolResult> LaunchAsync(JsonObject a, CancellationToken ct)
    {
        var exe = Vb6Install.FindExe() ?? throw new ToolException("VB6.EXE not found (registry ProductDir, .vbp association, default folder); set VB6MCP_VB6_EXE");
        var elevation = (a.Str("elevation") ?? "auto").ToLowerInvariant();
        if (elevation is not ("auto" or "elevated" or "normal")) throw new ToolException("elevation must be auto, elevated or normal");

        string? project = a.Str("project") is { Length: > 0 } p ? Path.GetFullPath(p) : null;
        if (project != null)
        {
            if (Path.GetExtension(project).ToLowerInvariant() is not (".vbp" or ".vbg")) throw new ToolException("project must be a .vbp or .vbg file");
            if (!File.Exists(project)) throw new ToolException("file not found: " + project);
            // VB6 refuses a project that is already open in another IDE ("already open" dialog, spike 12b)
            foreach (var inst in _env.Scan().Live)
            {
                try
                {
                    if ((await _env.Client(inst).ProjectsAsync()).Any(x => SamePath(x.FileName, project)))
                        throw new ToolException($"{project} is already open in VB6 (pid {inst.Pid}); VB6 would refuse to open it again");
                }
                catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException) { }
            }
        }

        var before = _env.Scan().Live.Select(i => i.Pid).ToHashSet();
        var vb6Before = Process.GetProcessesByName("VB6").Select(x => x.Id).ToHashSet();
        string via = "shell";
        Audit.Write("ide_launch", 0, $"{exe} elevation={elevation} project={project ?? "(none)"}");
        // An elevated watchdog starts VB6 elevated without a UAC prompt (spike 12b)
        if (elevation != "normal" && _env.FindWatchdog() is { } wdInfo)
        {
            var r = await _env.Client(wdInfo).SendRawAsync("LAUNCH", [project ?? ""], ct);
            if (r.Status == Core.Protocol.ResponseStatus.Ok) via = "watchdog (no UAC prompt)";
        }
        if (via == "shell")
        try { Start(exe, project, elevation); }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
        { throw new ToolException("the UAC prompt was declined or timed out; VB6 was not started"); }
        catch (Win32Exception ex) when (ex.NativeErrorCode == 740)
        {
            throw new ToolException("VB6 can't be started non-elevated on this machine: Windows requires elevation for this VB6.EXE " +
                "(a compatibility setting on that path, e.g. 'Run this program as an administrator' or a compatibility mode). " +
                "Use elevation 'auto' or 'elevated', or change VB6.EXE's compatibility settings (spike 16).");
        }

        // Wait for the add-in of the new IDE: a new live discovery entry (UAC approval can take a while).
        // VB6 loads add-ins one after the other, so a dialog from another add-in (or a startup prompt)
        // holds ours back: report it with the new pid instead of just timing out (Phase 2b live test).
        var deadline = DateTime.UtcNow + LaunchWait;
        InstanceInfo? inst2 = null;
        while (DateTime.UtcNow < deadline && inst2 is null)
        {
            await Task.Delay(500, ct);
            inst2 = _env.Scan().Live.FirstOrDefault(i => !before.Contains(i.Pid));
            if (inst2 != null) break;
            foreach (var newPid in Process.GetProcessesByName("VB6").Select(x => x.Id).Where(id => !vb6Before.Contains(id)))
            {
                var nw = _env.ReadWindows(newPid);
                if (nw.Dialogs.Count > 0)
                    return ToolResult.Ok(new
                    {
                        status = "dialog-before-addin", pid = newPid, via, mode = nw.Mode.ToString().ToLowerInvariant(), dialogs = Dialogs(nw),
                        note = $"VB6 started but shows a dialog before the vb6-mcp add-in loaded (e.g. another add-in's error at startup). " +
                               $"Answer it with ide_dialog_respond pid={newPid} (through the watchdog), then ide_list",
                    });
            }
        }
        if (inst2 is null)
            return ToolResult.Ok(new
            {
                status = "no-addin", exe, elevation, via,
                note = $"VB6 was started but the vb6-mcp add-in didn't connect within {LaunchWait.TotalSeconds:0} s. " +
                       "Check Add-Ins > Add-In Manager ('VB6 MCP Bridge', Load on Startup) and %LOCALAPPDATA%\\vb6mcp\\addin.log",
            });

        // Give a project a moment to load, then report what the user sees
        await Task.Delay(1500, ct);
        var w = _env.ReadWindows(inst2.Pid);
        IReadOnlyList<ProjectInfo> projects = [];
        try { projects = await _env.Client(inst2).ProjectsAsync(); } catch (AddinBusyException) { }
        return ToolResult.Ok(new
        {
            status = "started", pid = inst2.Pid, elevated = Vb6Install.IsElevated(inst2.Pid), elevation, exe, via,
            mode = w.Mode.ToString().ToLowerInvariant(), title = w.Title,
            projects = projects.Select(x => x.Name), dialogs = Dialogs(w),
            note = w.Dialogs.Count > 0 ? "VB6 is showing a dialog (e.g. the New Project chooser or a source-control prompt); answer it with ide_dialog_respond" : null,
        });
    }

    /// <summary>
    /// auto: ShellExecute, so Windows applies the machine's configuration and prompts for UAC when
    /// VB6 is set to elevate (CreateProcess would just fail with 740). elevated: verb "runas".
    /// normal: CreateProcess with __COMPAT_LAYER=RunAsInvoker; fails with 740 where elevation is forced.
    /// </summary>
    private static Process Start(string exe, string? project, string elevation)
    {
        var psi = new ProcessStartInfo(exe) { WorkingDirectory = Path.GetDirectoryName(project ?? exe)! };
        if (project != null) psi.Arguments = $"\"{project}\"";
        switch (elevation)
        {
            case "elevated": psi.UseShellExecute = true; psi.Verb = "runas"; break;
            case "normal": psi.UseShellExecute = false; psi.Environment["__COMPAT_LAYER"] = "RunAsInvoker"; break;
            default: psi.UseShellExecute = true; break;
        }
        return Process.Start(psi) ?? throw new ToolException("VB6 didn't start");
    }
}
