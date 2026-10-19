using System.Globalization;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;

namespace Vb6Mcp.Watchdog;

/// <summary>What the watchdog can do; the real implementation is <see cref="WatchdogActions"/>, tests use fakes.</summary>
public interface IWatchdogActions
{
    bool Elevated { get; }
    int ParentPid { get; }
    /// <summary>True if pid is a running VB6.EXE: every action is limited to VB6 processes (hard rule 8).</summary>
    bool IsVb6(int pid);
    /// <summary>Process id owning a window, 0 if none.</summary>
    int WindowPid(long hwnd);
    Task<string> BreakAsync(int pid, CancellationToken ct);
    string Press(long buttonHwnd);
    string Kill(int pid);
    /// <summary>Posts WM_CLOSE to the IDE's main window (ide_quit).</summary>
    string Quit(int pid);
    /// <summary>Starts VB6.EXE /make for a closed project under the roots (build_make without an IDE).</summary>
    string Make(string vbp, string defines);
    Task<string> SnapAsync(long hwnd, string id, CancellationToken ct);
    string Launch(string? project);
    /// <summary>Starts regsvr32 [/u] for a .dll/.ocx under the roots at the watchdog's level; returns its pid (spike 41).</summary>
    string Regsvr(string file, bool unregister);
    /// <summary>The exit code of a regsvr32 started by Regsvr, or "running".</summary>
    string RegStatus(int pid);
}

/// <summary>Raised by actions for an expected refusal; becomes ERR 409 with the message.</summary>
public sealed class WatchdogRefusedException(string message) : Exception(message);

/// <summary>
/// The watchdog's TCP endpoint (PLAN.md §3.1, docs/protocol.md "Watchdog"). Same framing, magic and
/// token check as the add-in. Unlike the add-in it runs on its own threads, so it keeps working while
/// a VB6 IDE thread is busy or hung, which is the whole point of it.
/// </summary>
public sealed class WatchdogHost(IWatchdogActions actions, Action<string> log, string discoveryPath)
{
    /// <summary>The exe's version (it was a constant that stayed 0.1.0 when the release became 0.2.0).</summary>
    public static readonly string Version = typeof(WatchdogHost).Assembly.GetName().Version?.ToString(3) ?? "0.0.0";
    private readonly DateTime _started = DateTime.Now;

    public string Token { get; } = Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();
    public int Port { get; private set; }

    /// <summary>Starts listening on 127.0.0.1 (random port), writes the discovery file and serves until cancelled.</summary>
    public async Task RunAsync(CancellationToken ct)
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);      // loopback only (hard rule 8)
        listener.Start();
        Port = ((IPEndPoint)listener.LocalEndpoint).Port;
        WatchdogDiscovery.Write(discoveryPath, Environment.ProcessId, Port, Token, Version);
        log($"listening on 127.0.0.1:{Port}, elevated={actions.Elevated}, parent={actions.ParentPid}");
        try
        {
            while (!ct.IsCancellationRequested)
            {
                var client = await listener.AcceptTcpClientAsync(ct);
                _ = Task.Run(() => ServeAsync(client, ct), ct);
            }
        }
        catch (OperationCanceledException) { }
        finally
        {
            listener.Stop();
            try { File.Delete(discoveryPath); } catch (IOException) { }
        }
    }

    private async Task ServeAsync(TcpClient client, CancellationToken ct)
    {
        using (client)
        {
            try
            {
                var stream = client.GetStream();
                var body = await Wire.ReadFrameAsync(stream, ct);
                if (body is null) return;
                var req = Wire.ParseRequest(body);
                string reply = req is null ? Wire.BuildResponse("?", ResponseStatus.Error, $"400{Wire.US}not a VB6MCP1 request")
                             : req.Token != Token ? Wire.BuildResponse(req.Id, ResponseStatus.Error, $"401{Wire.US}bad token")
                             : await DispatchAsync(req, ct);
                await stream.WriteAsync(Wire.EncodeFrame(reply), ct);
            }
            catch (Exception ex) when (ex is IOException or ProtocolException or OperationCanceledException or SocketException) { }
        }
    }

    /// <summary>Runs one verb. Every action is checked to target a VB6 process and logged (hard rule 8).</summary>
    public async Task<string> DispatchAsync(Request req, CancellationToken ct)
    {
        string Ok(string payload) => Wire.BuildResponse(req.Id, ResponseStatus.Ok, payload);
        string Err(int code, string msg) => Wire.BuildResponse(req.Id, ResponseStatus.Error, $"{code}{Wire.US}{msg}");
        var a = req.Args;
        try
        {
            switch (req.Verb.ToUpperInvariant())
            {
                case "PING":
                    return Ok(string.Join(Wire.US, Environment.ProcessId.ToString(CultureInfo.InvariantCulture), Version,
                        actions.Elevated ? "1" : "0", actions.ParentPid.ToString(CultureInfo.InvariantCulture),
                        ((int)(DateTime.Now - _started).TotalSeconds).ToString(CultureInfo.InvariantCulture)));
                case "BREAK":
                {
                    int pid = Vb6Pid(a, 0);
                    log($"BREAK {pid}");
                    return Ok(await actions.BreakAsync(pid, ct));
                }
                case "PRESS":
                {
                    long hwnd = Vb6Window(a, 0);
                    log($"PRESS {hwnd:X}");
                    return Ok(actions.Press(hwnd));
                }
                case "KILL":
                {
                    int pid = Vb6Pid(a, 0);
                    // destructive: the literal word "force" is required, never a default (hard rules 9, 12)
                    if (a.Length < 2 || a[1] != "force") return Err(400, "KILL needs the argument 'force'");
                    log($"KILL {pid} force");
                    return Ok(actions.Kill(pid));
                }
                case "QUIT":
                {
                    int pid = Vb6Pid(a, 0);
                    log($"QUIT {pid}");
                    return Ok(actions.Quit(pid));
                }
                case "MAKE":
                {
                    if (a.Length < 1 || a.Length > 2) return Err(400, "MAKE takes the .vbp path and optional defines");
                    log($"MAKE {a[0]}{(a.Length > 1 && a[1].Length > 0 ? " /d " + a[1] : "")}");
                    return Ok(actions.Make(a[0], a.Length > 1 ? a[1] : ""));
                }
                case "REGSVR":
                {
                    // registration changes the machine: the literal "register" or "unregister" is required (hard rules 9, 13)
                    if (a.Length != 2 || a[1] is not ("register" or "unregister")) return Err(400, "REGSVR takes the file and 'register' or 'unregister'");
                    log($"REGSVR {a[1]} {a[0]}");
                    return Ok(actions.Regsvr(a[0], a[1] == "unregister"));
                }
                case "REGSTATUS":
                {
                    if (a.Length != 1 || !int.TryParse(a[0], NumberStyles.None, CultureInfo.InvariantCulture, out var regPid)) return Err(400, "REGSTATUS takes a pid");
                    return Ok(actions.RegStatus(regPid));
                }
                case "SNAP":
                {
                    long hwnd = Vb6Window(a, 0);
                    if (a.Length < 2 || !IsSafeId(a[1])) return Err(400, "id must be 1-40 characters of A-Z a-z 0-9 _ -");
                    return Ok(await actions.SnapAsync(hwnd, a[1], ct));
                }
                case "LAUNCH":
                {
                    string? project = a.Length > 0 && a[0].Length > 0 ? a[0] : null;
                    if (project != null && (!Path.IsPathFullyQualified(project) || Path.GetExtension(project).ToLowerInvariant() is not (".vbp" or ".vbg") || !File.Exists(project)))
                        return Err(400, "project must be the full path of an existing .vbp or .vbg");
                    log($"LAUNCH {project ?? "(no project)"}");
                    return Ok(actions.Launch(project));
                }
                default:
                    return Err(400, "unknown verb " + req.Verb);
            }
        }
        catch (WatchdogRefusedException ex) { return Err(409, ex.Message); }
        catch (ArgumentException ex) { return Err(400, ex.Message); }
        catch (Exception ex) { log($"{req.Verb} failed: {ex}"); return Err(500, ex.Message); }
    }

    private int Vb6Pid(string[] a, int i)
    {
        if (a.Length <= i || !int.TryParse(a[i], NumberStyles.None, CultureInfo.InvariantCulture, out var pid)) throw new ArgumentException("missing or bad pid");
        if (!actions.IsVb6(pid)) throw new ArgumentException($"process {pid} isn't a running VB6.EXE");
        return pid;
    }

    private long Vb6Window(string[] a, int i)
    {
        if (a.Length <= i || !long.TryParse(a[i], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var hwnd)) throw new ArgumentException("missing or bad window handle (hex)");
        int pid = actions.WindowPid(hwnd);
        if (pid == 0 || !actions.IsVb6(pid)) throw new ArgumentException($"window {hwnd:X} doesn't belong to a VB6 process");
        return hwnd;
    }

    private static bool IsSafeId(string id) => id.Length is >= 1 and <= 40 && id.All(c => char.IsAsciiLetterOrDigit(c) || c is '_' or '-');
}
