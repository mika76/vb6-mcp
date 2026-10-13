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
