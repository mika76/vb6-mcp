using System.Globalization;
using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 2b: what the server does through the elevated watchdog when an IDE's own add-in can't
/// answer (PLAN.md §3.1). The watchdog is a sidecar: every tool here also says what happens without it.
/// </summary>
public sealed partial class IdeTools
{
    private IReadOnlyList<Tool> WatchdogTools() =>
    [
        new("ide_kill",
            "Last resort: end a VB6 process that is hung or stuck (needs force=true; never done automatically). Unsaved work is lost " +
            "except for the recovery copies the watchdog took; the result lists them with their age (see recovery_list). Try " +
            "run_break and ide_dialog_respond first. Needs the watchdog (it runs elevated like VB6).",
            Schema.Object(("force", "boolean", "Required: true.", true), ("pid", "integer", PidDesc, false)),
            (a, ct) => KillAsync(a, ct)),
        new("ide_quit",
            "Close a VB6 IDE normally (File > Exit), e.g. one this session started with ide_launch and no longer needs. Refused while " +
            "anything is unsaved, unless discard=true (the unsaved changes are thrown away, as project_close discard does). Only for an " +
            "IDE this session owns; another one needs force=true (only when the user asked). Design mode. Needs the watchdog.",
            Schema.Object(("discard", "boolean", "Close even with unsaved changes, discarding them.", false),
                          ("force", "boolean", "Close an IDE this session doesn't own.", false), ("pid", "integer", PidDesc, false)),
            (a, ct) => QuitAsync(a, ct)),
    ];

    /// <summary>How long ide_quit waits for VB6 to exit after WM_CLOSE.</summary>
    public TimeSpan QuitWait { get; init; } = TimeSpan.FromSeconds(15);

    /// <summary>
    /// ide_quit (PLAN.md C7). Unsaved work is never lost silently (hard rule 4): refused, or discarded per project
    /// through project_close's snapshot path with discard=true. The close itself is WM_CLOSE from the elevated watchdog
    /// (a non-elevated post to the IDE is refused, spike 10). Afterwards the owner file goes; ide_list removes the
    /// discovery file of the gone process.
    /// </summary>
    private async Task<ToolResult> QuitAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        RequireDesignMode(inst);
        if (OwnerOf(inst.Pid) is not { } owner || !IdeOwnership.IsMe(owner))
            if (a.Bool("force") != true)
                throw new ToolException($"VB6 {inst.Pid} isn't this session's (it didn't start or claim it): closing it could end the user's work. Pass force=true only if the user asked");
        var wd = Watchdog() ?? throw new ToolException(NoWatchdog + "; close VB6 yourself (File > Exit)");
        var client = _env.Client(inst);
        var unsaved = new List<string>();
        var projects = await client.ProjectsAsync();
        foreach (var p in projects)
        {
            var dirty = (await client.ComponentsAsync(p.Name)).Where(c => c.IsDirty).Select(c => $"{p.Name}.{c.Name}").ToList();
            if (dirty.Count == 0 && p.IsDirty) dirty.Add($"{p.Name} (project settings)");
            unsaved.AddRange(dirty);
        }
        var discarded = new List<object>();
        if (unsaved.Count > 0)
        {
            if (a.Bool("discard") != true)
                throw new ToolException($"unsaved changes: {string.Join(", ", unsaved)}. Save them (project_save), or pass discard=true to throw them away");
            foreach (var p in projects.Where(p => unsaved.Any(u => u.StartsWith(p.Name + ".", StringComparison.Ordinal) || u.StartsWith(p.Name + " ", StringComparison.Ordinal))))
            {
                var r = (await CloseProjectAsync(new JsonObject { ["pid"] = inst.Pid, ["project"] = p.Name, ["discard"] = true }, ct)).ToJson();
                var text = (string?)r["content"]?[0]?["text"] ?? "";
                if ((bool?)r["isError"] == true || !text.Contains("\"status\":\"closed\""))
                    return ToolResult.Ok(new { status = "not-quit", pid = inst.Pid, note = $"discarding {p.Name} didn't finish, so VB6 was left open", closeResult = JsonNode.Parse(text) });
                discarded.Add(new { project = p.Name, result = JsonNode.Parse(text) });
            }
        }
        Audit.Write("ide_quit", inst.Pid, $"WM_CLOSE{(discarded.Count > 0 ? $", discarded {string.Join(", ", unsaved)}" : "")}");
        EndedOnPurpose(inst.Pid);
        var q = await wd.SendRawAsync("QUIT", [inst.Pid.ToString(CultureInfo.InvariantCulture)], ct);
        if (q.Status == ResponseStatus.Error)
            throw new ToolException(q.ErrorMessage.StartsWith("unknown verb", StringComparison.Ordinal)
                ? "the running watchdog is older than this server (no QUIT verb): restart it (scripts\\build-addin.ps1 -WatchdogOnly), or close VB6 yourself"
                : "the watchdog couldn't close VB6: " + q.ErrorMessage);
        var deadline = DateTime.UtcNow + QuitWait;
        DateTime? dialogSince = null;
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(300, ct);
            if (!ProcessAlive(inst.Pid))
            {
                IdeOwnership.Release(inst.Pid, OwnersFolder);
                return ToolResult.Ok(new { status = "quit", pid = inst.Pid, discarded = discarded.Count > 0 ? discarded : null });
            }
            // A dialog can flash up while VB shuts down (the launch integration test saw one, and VB still exited): report
            // only one that stays for 2 s while the process lives on
            var w = _env.ReadWindows(inst.Pid);
            if (w.Dialogs.Count == 0) { dialogSince = null; continue; }
            dialogSince ??= DateTime.UtcNow;
            if (DateTime.UtcNow - dialogSince >= TimeSpan.FromSeconds(2))
                return ToolResult.Ok(new { status = "dialog", pid = inst.Pid, dialogs = Dialogs(w), note = "VB6 shows a dialog while closing (another add-in, or a save prompt); it was not answered" });
        }
        return ToolResult.Ok(new { status = "still-running", pid = inst.Pid, note = $"VB6 didn't exit within {QuitWait.TotalSeconds:0} s; check ide_status" });
    }

    /// <summary>Whether a process still runs (tests replace it). A medium process can open an elevated one for SYNCHRONIZE (spike 24).</summary>
    public Func<int, bool> ProcessAlive { get; init; } = pid =>
    {
        try { using var p = System.Diagnostics.Process.GetProcessById(pid); return !p.HasExited; }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception) { return false; }
    };

    /// <summary>A client for the running watchdog, or null. 10 s timeout: BREAK alone waits up to 3 s for break mode.</summary>
    private AddinClient? Watchdog() => _env.FindWatchdog() is { } wd ? new AddinClient(wd, TimeSpan.FromSeconds(10)) : null;

    /// <summary>Watchdog summary for ide_list / ide_status.</summary>
    private async Task<object?> WatchdogInfoAsync()
    {
        var wd = Watchdog();
        if (wd is null) return null;
        try
        {
            var f = (await wd.SendAsync("PING")).Split(Wire.US);
            return new { pid = wd.Instance.Pid, version = f.Length > 1 ? f[1] : "?", elevated = f.Length > 2 && f[2] == "1" };
        }
        catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException or ProtocolException)
        {
            return new { pid = wd.Instance.Pid, error = ex.Message };
        }
    }

    private const string NoWatchdog = "the watchdog isn't running (the add-in starts it from its 'watchdog' folder; see %LOCALAPPDATA%\\vb6mcp\\addin.log)";

    // ---------- run_break with the watchdog fallback ----------

    /// <summary>
    /// Break: the IDE's Break command through the add-in when the IDE responds; when the debugged
    /// program keeps the IDE thread busy, the add-in can't take the command, so the watchdog injects
    /// Ctrl+Break (spike 12). Either way the result is verified from the title.
    /// </summary>
    private async Task<ToolResult> RunBreakAsync(JsonObject a, CancellationToken ct)
    {
        var inst = ResolveInstance(a);
        var w = _env.ReadWindows(inst.Pid);
        if (w.Mode != IdeMode.Run) throw new ToolException($"the IDE is in {Mode(w.Mode)} mode; run_break needs run mode");
        if (w.Responds && w.Dialogs.Count == 0)
        {
            try
            {
                await _env.Client(inst).CommandAsync(CmdBreak);
                var after = await WaitForAsync(inst, s => s.Mode == IdeMode.Break, ct, TimeSpan.FromSeconds(3));
                if (after.Mode == IdeMode.Break) return ToolResult.Ok(new { method = "IDE Break command", state = State("break", after) });
            }
            catch (AddinBusyException) { }                     // became busy in between: fall through to the watchdog
        }
        var wd = Watchdog() ?? throw new ToolException("the IDE isn't responding (the program is busy) and " + NoWatchdog);
        Audit.Write("run_break", inst.Pid, "watchdog Ctrl+Break");
        var r = await wd.SendRawAsync("BREAK", [inst.Pid.ToString(CultureInfo.InvariantCulture)], ct);
        var state = _env.ReadWindows(inst.Pid);
        if (r.Status == ResponseStatus.Error)
            throw new ToolException($"the watchdog couldn't break: {r.ErrorMessage} (mode now: {Mode(state.Mode)})");
        return ToolResult.Ok(new
        {
            method = "watchdog: injected Ctrl+Break" + (r.Payload.EndsWith("ghost", StringComparison.Ordinal) ? " (into Windows' stand-in for the unresponsive window)" : ""),
            state = State(state.Mode == IdeMode.Break ? "break" : "sent", state),
        });
    }

    // ---------- ide_kill ----------

    private async Task<ToolResult> KillAsync(JsonObject a, CancellationToken ct)
    {
        if (a.Bool("force") != true) throw new ToolException("ide_kill needs force=true: it ends VB6 and loses unsaved work that has no recovery copy");
        var inst = ResolveInstance(a);
        var wd = Watchdog() ?? throw new ToolException(NoWatchdog + "; end VB6 with Task Manager instead");
        var before = _env.ReadWindows(inst.Pid);
        Audit.Write("ide_kill", inst.Pid, $"force (title '{before.Title}', responds={before.Responds})");
        EndedOnPurpose(inst.Pid);   // its discovery file stays behind, but this isn't a crash to report
        var r = await wd.SendRawAsync("KILL", [inst.Pid.ToString(CultureInfo.InvariantCulture), "force"], ct);
        if (r.Status == ResponseStatus.Error) throw new ToolException("the watchdog couldn't kill VB6: " + r.ErrorMessage);
        return ToolResult.Ok(new
        {
            pid = inst.Pid, status = r.Payload, wasTitle = before.Title,
            recovery = RecoverySummary(inst.Pid),
            // Killed in run/break mode, an ActiveX project's classes stay registered to VB6DEBUG.DLL (spike 24)
            comLeftovers = before.Mode is IdeMode.Run or IdeMode.Break
                ? Core.Files.TypeLibRegistry.All().Where(Core.Files.ComRegistry.IsProjectTypeLib).Select(t => t.Win32Path).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                : null,
            comNote = before.Mode is IdeMode.Run or IdeMode.Break
                ? "it was killed while running: ActiveX projects it was debugging stay registered to the VB6 debugger (comLeftovers lists their .vbp files); com_orphans shows the repair"
                : null,
            note = "the IDE's discovery file stays behind and is reported as crashed by ide_list; reopen with ide_launch, then restore with recovery_restore",
        });
    }

    // ---------- fallbacks used by other tools ----------

    /// <summary>Clicks a dialog button through the watchdog (the add-in didn't answer).</summary>
    private async Task<bool> PressViaWatchdogAsync(long buttonHwnd, int pid)
    {
        var wd = Watchdog();
        if (wd is null) return false;
        Audit.Write("ide_dialog_respond", pid, $"watchdog PRESS {buttonHwnd:X}");
        var r = await wd.SendRawAsync("PRESS", [buttonHwnd.ToString("X", CultureInfo.InvariantCulture)]);
        return r.Status == ResponseStatus.Ok;
    }

    /// <summary>PNG of a window through the watchdog (the add-in is busy or gone); null without a watchdog.</summary>
    private async Task<(byte[] Png, string Method)?> SnapViaWatchdogAsync(long hwnd, int maxWidth, CancellationToken ct)
    {
        var wd = Watchdog();
        if (wd is null) return null;
        string id = "w" + DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture);
        var r = await wd.SendRawAsync("SNAP", [hwnd.ToString("X", CultureInfo.InvariantCulture), id], ct);
        if (r.Status == ResponseStatus.Error) throw new ToolException("the watchdog couldn't capture the window: " + r.ErrorMessage);
        var f = r.Payload.Split(Wire.US);
        using var bmp = new System.Drawing.Bitmap(f[0]);
        var png = Screenshots.ToPng(bmp, maxWidth);
        bmp.Dispose();
        try { File.Delete(f[0]); } catch (IOException) { }
        return (png, "watchdog " + (f.Length > 1 ? f[1] : "capture"));
    }
}
