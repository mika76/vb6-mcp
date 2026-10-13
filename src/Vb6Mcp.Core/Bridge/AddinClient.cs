using System.Globalization;
using System.Net.Sockets;
using Vb6Mcp.Core.Protocol;

namespace Vb6Mcp.Core.Bridge;

/// <summary>
/// Talks to one add-in instance over loopback TCP (docs/protocol.md). Every request has a timeout:
/// the add-in runs on the IDE thread, which can be busy (debugged program in a loop) or blocked,
/// so a timeout becomes <see cref="AddinBusyException"/> instead of a hang (AGENTS.md hard rule 6).
/// </summary>
public sealed class AddinClient(InstanceInfo instance, TimeSpan? timeout = null)
{
    private int _nextId;

    public InstanceInfo Instance { get; } = instance;
    public TimeSpan Timeout { get; } = timeout ?? TimeSpan.FromSeconds(5);

    /// <summary>Sends one request on a fresh connection and returns the raw response. ERR is returned, not thrown.</summary>
    public async Task<Response> SendRawAsync(string verb, string[] args, CancellationToken ct = default)
    {
        string id = Interlocked.Increment(ref _nextId).ToString(CultureInfo.InvariantCulture);
        var body = Wire.BuildRequest(Instance.Token, id, verb, args);
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        cts.CancelAfter(Timeout);
        try
        {
            using var tcp = new TcpClient();
            await tcp.ConnectAsync("127.0.0.1", Instance.Port, cts.Token).ConfigureAwait(false);
            var stream = tcp.GetStream();
            await stream.WriteAsync(Wire.EncodeFrame(body), cts.Token).ConfigureAwait(false);
            var reply = await Wire.ReadFrameAsync(stream, cts.Token).ConfigureAwait(false)
                        ?? throw new ProtocolException("add-in closed the connection without replying");
            var response = Wire.ParseResponse(reply);
            if (response.Id != id) throw new ProtocolException($"response id {response.Id} doesn't match request id {id}");
            return response;
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            throw new AddinBusyException(Instance.Pid, verb, Timeout);
        }
        catch (SocketException ex)
        {
            throw new AddinUnreachableException(Instance.Pid, ex.Message);
        }
    }

    /// <summary>Sends a request and throws <see cref="AddinErrorException"/> on ERR. Returns the payload.</summary>
    public async Task<string> SendAsync(string verb, params string[] args)
    {
        var r = await SendRawAsync(verb, args).ConfigureAwait(false);
        if (r.Status == ResponseStatus.Error) throw new AddinErrorException(r.ErrorCode, r.ErrorMessage);
        return r.Payload;
    }

    // ---- typed verbs (docs/protocol.md "Verbs (v1)") ----

    public async Task<PingResult> PingAsync()
    {
        var f = Wire.Fields(await SendAsync("PING").ConfigureAwait(false), 3);
        return new PingResult(Int(f[0]), f[1], Int(f[2]));
    }

    public async Task<IReadOnlyList<ProjectInfo>> ProjectsAsync()
        => Wire.Records(await SendAsync("PROJECTS").ConfigureAwait(false))
            .Select(r => new ProjectInfo(Int(r[0]), r[1], Int(r[2]), r[3], r[4] == "1")).ToList();

    public async Task<IReadOnlyList<ComponentInfo>> ComponentsAsync(string project)
        => Wire.Records(await SendAsync("COMPONENTS", project).ConfigureAwait(false))
            .Select(r => new ComponentInfo(r[0], Int(r[1]), r[2] == "1", Int(r[3]), r[4])).ToList();

    public async Task<CodeText> ReadAsync(string project, string component, int? startLine = null, int? count = null)
    {
        string[] args = startLine is int s && count is int c
            ? [project, component, s.ToString(CultureInfo.InvariantCulture), c.ToString(CultureInfo.InvariantCulture)]
            : [project, component];
        var f = Wire.Fields(await SendAsync("READ", args).ConfigureAwait(false), 3);
        return new CodeText(Int(f[0]), f[1] == "1", f[2]);
    }

    public async Task<IReadOnlyList<MemberInfo>> MembersAsync(string project, string component)
        => Wire.Records(await SendAsync("MEMBERS", project, component).ConfigureAwait(false))
            .Select(r => new MemberInfo(r[0], Int(r[1]), Int(r[2]), r[3] == "1", Int(r[4]),
                                        NInt(r[5]), NInt(r[6]), NInt(r[7]), NInt(r[8]))).ToList();

    public async Task<ProcText> ProcAsync(string project, string component, string name, int procKind)
    {
        var f = Wire.Fields(await SendAsync("PROC", project, component, name, procKind.ToString(CultureInfo.InvariantCulture)).ConfigureAwait(false), 4);
        return new ProcText(Int(f[0]), Int(f[1]), Int(f[2]), f[3]);
    }

    /// <summary>
    /// Queues opening a project (deferred in the add-in, hard rule 14). Returns once queued; the caller
    /// watches PROJECTS and dialogs for the outcome. <paramref name="replace"/> closes the open projects first.
    /// </summary>
    public async Task OpenAsync(string path, bool replace)
    {
        var r = await SendRawAsync("OPEN", [path, replace ? "replace" : "add"]).ConfigureAwait(false);
        if (r.Status == ResponseStatus.Error) throw new AddinErrorException(r.ErrorCode, r.ErrorMessage);
        if (r.Status != ResponseStatus.Queued) throw new ProtocolException("OPEN should answer QUEUED");
    }

    /// <summary>
    /// Replaces <paramref name="count"/> lines at <paramref name="startLine"/> with <paramref name="newText"/>
    /// (count 0 = insert, empty text = delete), but only if those lines still equal <paramref name="expectedOld"/>
    /// (409 otherwise). Returns the new line count and dirty flag.
    /// </summary>
    public async Task<(int CountOfLines, bool IsDirty)> ReplaceAsync(string project, string component, int startLine, int count, string expectedOld, string newText)
    {
        var f = Wire.Fields(await SendAsync("REPLACE", project, component,
            startLine.ToString(CultureInfo.InvariantCulture), count.ToString(CultureInfo.InvariantCulture), expectedOld, newText).ConfigureAwait(false), 2);
        return (Int(f[0]), f[1] == "1");
    }

    /// <summary>Queues File &gt; Save Project for the project (deferred in the add-in).</summary>
    public async Task SaveAsync(string project)
    {
        var r = await SendRawAsync("SAVE", [project]).ConfigureAwait(false);
        if (r.Status == ResponseStatus.Error) throw new AddinErrorException(r.ErrorCode, r.ErrorMessage);
        if (r.Status != ResponseStatus.Queued) throw new ProtocolException("SAVE should answer QUEUED");
    }

    /// <summary>Clicks a dialog button of the IDE process (posted BM_CLICK inside the IDE).</summary>
    public Task ClickAsync(long buttonHwnd) => SendAsync("CLICK", buttonHwnd.ToString("X", CultureInfo.InvariantCulture));

    // ---- debugging (Phase 2) ----

    /// <summary>Queues a debug command by id (allow-listed in the add-in) or "fullcompile".</summary>
    public async Task CommandAsync(string idOrFullCompile)
    {
        var r = await SendRawAsync("CMD", [idOrFullCompile]).ConfigureAwait(false);
        if (r.Status == ResponseStatus.Error) throw new AddinErrorException(r.ErrorCode, r.ErrorMessage);
        if (r.Status != ResponseStatus.Queued) throw new ProtocolException("CMD should answer QUEUED");
    }

    /// <summary>Shows the code pane and puts the cursor at the start of the line.</summary>
    public Task SetSelectionAsync(string project, string component, int line)
        => SendAsync("SETSEL", project, component, line.ToString(CultureInfo.InvariantCulture));

    /// <summary>Active code pane: project, component, selection and procedure.</summary>
    public async Task<Where> WhereAsync()
    {
        var f = Wire.Fields(await SendAsync("WHERE").ConfigureAwait(false), 7);
        return new Where(f[0], f[1], Int(f[2]), Int(f[3]), Int(f[4]), Int(f[5]), f[6]);
    }

    /// <summary>Queues typing a line + Enter into the Immediate window (deferred: it can run code).</summary>
    public async Task ImmediateAsync(long immediateHwnd, string line)
    {
        var r = await SendRawAsync("IMM", [immediateHwnd.ToString("X", CultureInfo.InvariantCulture), line]).ConfigureAwait(false);
        if (r.Status == ResponseStatus.Error) throw new AddinErrorException(r.ErrorCode, r.ErrorMessage);
    }

    /// <summary>
    /// The Immediate window's whole text, via the clipboard (the window is custom-drawn). Several
    /// requests with the IDE idle in between, because VB6 updates command Enabled states only when
    /// idle and completes Copy asynchronously (Phase 2 live tests): IMMSELECT, pause, IMMCOPY, then
    /// IMMPEEK until the copy has arrived, and always IMMRESTORE to give the user's clipboard back.
    /// </summary>
    public async Task<string> ImmediateReadAsync(long immediateHwnd)
    {
        string hwnd = immediateHwnd.ToString("X", CultureInfo.InvariantCulture);
        for (int attempt = 1; ; attempt++)
        {
            try { await SendAsync("IMMSELECT", hwnd).ConfigureAwait(false); break; }
            catch (AddinErrorException ex) when (ex.Code == 409 && attempt < 3) { await Task.Delay(200).ConfigureAwait(false); }
        }
        try
        {
            await Task.Delay(150).ConfigureAwait(false);
            if (await SendAsync("IMMCOPY").ConfigureAwait(false) == "empty") return "";
            for (int i = 0; i < 20; i++)                        // up to ~2 s for the asynchronous copy
            {
                await Task.Delay(100).ConfigureAwait(false);
                var text = await SendAsync("IMMPEEK").ConfigureAwait(false);
                if (text.Length > 0) return text;
            }
            return "";
        }
        finally
        {
            await SendAsync("IMMRESTORE").ConfigureAwait(false);
        }
    }

    /// <summary>Queues View > Call Stack (a modal dialog; read it with <see cref="ListItemsAsync"/>, close it with <see cref="ClickAsync"/>).</summary>
    public async Task CallStackAsync()
    {
        var r = await SendRawAsync("CALLSTACK", []).ConfigureAwait(false);
        if (r.Status == ResponseStatus.Error) throw new AddinErrorException(r.ErrorCode, r.ErrorMessage);
        if (r.Status != ResponseStatus.Queued) throw new ProtocolException("CALLSTACK should answer QUEUED");
    }

    /// <summary>Items of a list box of the IDE process (read in-process: UIPI blocks LB_* from here).</summary>
    public async Task<IReadOnlyList<string>> ListItemsAsync(long hwnd)
    {
        var payload = await SendAsync("LISTITEMS", hwnd.ToString("X", CultureInfo.InvariantCulture)).ConfigureAwait(false);
        return payload.Length == 0 ? [] : payload.Split(Wire.RS);
    }

    // ---- the running program (Phase 3) ----

    /// <summary>Posts a click to a command button / check box / option button of the IDE process.</summary>
    public Task AppClickAsync(long hwnd) => SendAsync("APPCLICK", hwnd.ToString("X", CultureInfo.InvariantCulture));

    /// <summary>Queues setting a text box's text (deferred: it fires the Change event).</summary>
    public async Task AppTextAsync(long hwnd, string text)
    {
        var r = await SendRawAsync("APPTEXT", [hwnd.ToString("X", CultureInfo.InvariantCulture), text]).ConfigureAwait(false);
        if (r.Status == ResponseStatus.Error) throw new AddinErrorException(r.ErrorCode, r.ErrorMessage);
    }

    /// <summary>Queues a screenshot of a window of the IDE process; returns the BMP path it will be written to.</summary>
    public async Task<string> SnapAsync(long hwnd, string id)
    {
        var r = await SendRawAsync("SNAP", [hwnd.ToString("X", CultureInfo.InvariantCulture), id]).ConfigureAwait(false);
        if (r.Status == ResponseStatus.Error) throw new AddinErrorException(r.ErrorCode, r.ErrorMessage);
        return r.Payload;
    }

    /// <summary>
    /// Posts mouse input to a window of the IDE process, in that window's client coordinates.
    /// action: click, dblclick, rclick, drag (to x2, y2) or wheel (x2 = notches, negative = down).
    /// </summary>
    public Task MouseAsync(long hwnd, string action, int x, int y, int? x2 = null, int? y2 = null)
    {
        var args = new List<string> { hwnd.ToString("X", CultureInfo.InvariantCulture), action, S(x), S(y) };
        if (x2 is { } a) args.Add(S(a));
        if (y2 is { } b) args.Add(S(b));
        return SendAsync("MOUSE", [.. args]);
    }

    /// <summary>Posts scroll commands (SB_LINEUP=0 .. SB_BOTTOM=7) to a window or scroll bar control of the IDE process.</summary>
    public Task ScrollAsync(long hwnd, bool vertical, int code, int count)
        => SendAsync("SCROLL", hwnd.ToString("X", CultureInfo.InvariantCulture), vertical ? "v" : "h", S(code), S(count));

    /// <summary>
    /// Real mouse input (SendInput, spike 17): brings the form to the front and moves the cursor.
    /// Always follow with <see cref="MouseRestoreAsync"/> once the program has processed it.
    /// </summary>
    public Task MouseRealAsync(long hwnd, string action, int x, int y, int? x2 = null, int? y2 = null)
    {
        var args = new List<string> { hwnd.ToString("X", CultureInfo.InvariantCulture), action, S(x), S(y) };
        if (x2 is { } a && y2 is { } b) { args.Add(S(a)); args.Add(S(b)); }
        return SendAsync("MOUSEREAL", [.. args]);
    }

    /// <summary>Puts the cursor and the previously active window back after <see cref="MouseRealAsync"/>.</summary>
    public Task MouseRestoreAsync() => SendAsync("MOUSERESTORE");

    /// <summary>Range and position of a scroll bar control (read by the add-in: UIPI blocks SBM_* messages from the server).</summary>
    public async Task<(int Min, int Max, int Pos)> ScrollPosAsync(long hwnd)
    {
        var f = (await SendAsync("SCROLLPOS", hwnd.ToString("X", CultureInfo.InvariantCulture)).ConfigureAwait(false)).Split(Wire.US);
        return (Int(f[0]), Int(f[1]), Int(f[2]));
    }

    private static string S(int v) => v.ToString(CultureInfo.InvariantCulture);
    private static int Int(string s) => int.Parse(s, CultureInfo.InvariantCulture);
    private static int? NInt(string s) => s.Length == 0 ? null : int.Parse(s, CultureInfo.InvariantCulture);
}

public sealed record PingResult(int Pid, string AddinVersion, int Protocol);
public sealed record ProjectInfo(int Index, string Name, int Type, string FileName, bool IsDirty);
public sealed record ComponentInfo(string Name, int Type, bool IsDirty, int CountOfLines, string FileName);
public sealed record CodeText(int CountOfLines, bool IsDirty, string Text);
public sealed record MemberInfo(string Name, int MemberType, int Scope, bool IsStatic, int CodeLocation,
                                int? ProcKind, int? ProcStart, int? ProcBody, int? ProcCount);
public sealed record ProcText(int ProcStart, int ProcBody, int ProcCount, string Text);
public sealed record Where(string Project, string Component, int StartLine, int StartColumn, int EndLine, int EndColumn, string Procedure);

/// <summary>The add-in didn't answer in time: the IDE thread is busy (run mode loop) or blocked.</summary>
public sealed class AddinBusyException(int pid, string verb, TimeSpan timeout)
    : Exception($"VB6 (pid {pid}) didn't answer '{verb}' within {timeout.TotalSeconds:0.#} s: the IDE is busy or blocked");

/// <summary>Couldn't connect: the add-in isn't listening (IDE closed, add-in unloaded, stale discovery file).</summary>
public sealed class AddinUnreachableException(int pid, string detail)
    : Exception($"VB6 add-in (pid {pid}) is unreachable: {detail}");

/// <summary>The add-in answered ERR (docs/protocol.md "Error codes").</summary>
public sealed class AddinErrorException(int code, string message) : Exception($"add-in error {code}: {message}")
{
    public int Code { get; } = code;
    public string AddinMessage { get; } = message;
}
