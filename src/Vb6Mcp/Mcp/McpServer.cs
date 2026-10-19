using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Vb6Mcp.Mcp;

/// <summary>
/// Hand-written MCP server: JSON-RPC 2.0, newline-delimited, over stdio (spike 1, no SDK).
/// stdout carries protocol messages only; diagnostics go to the log writer (stderr or a file).
/// Requests are handled concurrently so a slow tool call doesn't block ping or other calls;
/// writes to stdout are serialized.
/// </summary>
public sealed class McpServer(TextReader input, TextWriter output, TextWriter log, IReadOnlyList<Tool> tools)
{
    public const string ServerName = "vb6-mcp";

    /// <summary>The client's name and version from initialize (clientInfo), e.g. "claude-code 2.1"; used to name IDE owners.</summary>
    public static string ClientName { get; private set; } = "unknown client";

    /// <summary>The client's clientInfo.name alone ("claude-code", "copilot-cli", "codex-mcp-client"), "" before initialize.</summary>
    public static string ClientId { get; private set; } = "";

    /// <summary>
    /// Whether the client can show the user a form (MCP elicitation, form mode). 2025-06-18 clients advertise
    /// "elicitation": {}; 2025-11-25 ones "elicitation": {"form": {}, "url": {}}, where {} still means form (spike 40).
    /// </summary>
    public static bool ClientElicits { get; private set; }

    /// <summary>The running server, for tools that send the client a request (elicitation); null in tests that call tools directly.</summary>
    public static McpServer? Current { get; private set; }

    /// <summary>
    /// Whether an initialize "capabilities" object offers form elicitation: "elicitation" present, and either empty
    /// (2025-06-18, or 2025-11-25's "{} means form") or naming "form". A client offering only "url" can't show our form.
    /// </summary>
    public static bool OffersFormElicitation(JsonObject? capabilities) =>
        capabilities?["elicitation"] is JsonObject e && (e.Count == 0 || e.ContainsKey("form"));
    public static readonly string ServerVersion = typeof(McpServer).Assembly.GetName().Version?.ToString(3) ?? "0.0.0";

    /// <summary>Protocol versions we accept, newest first. Unknown requested versions get the newest (spike 1).</summary>
    public static readonly string[] SupportedVersions = ["2025-06-18", "2025-03-26", "2024-11-05"];

    public static readonly JsonSerializerOptions Json = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    /// <summary>
    /// Sent in the initialize result ("instructions"); Claude Code shows it to the model up front. Without it, the
    /// benchmark (docs/benchmarks) saw the agent try Edit on VB6 files first (blocked by the hook), then discover these
    /// tools, which roughly doubled tokens on edit tasks. Short on purpose: it costs tokens on every request.
    /// </summary>
    public const string Instructions =
        "Tools for the Visual Basic 6 IDE and VB6 projects. VB6 sources (.bas .cls .frm .ctl .vbp) are Windows-1252 with CR LF and " +
        "hidden header/Attribute lines: change them only with these tools, never Edit, Write or shell commands. Never touch .frx .ctx " +
        ".pgx .dsx .res (binary, VB rewrites them): read them with form_resources, res_read. Open project (ide_list): code_read, " +
        "proc_get, code_replace_lines, code_insert, proc_replace, project_save; project_checkpoint/project_rollback around big changes. " +
        "Forms: form_controls, form_add_control, form_set_properties; rows and columns with form_layout, not hand-computed Left/Top. " +
        "Closed project: file_read, file_replace_lines. Build: build_make. Run: run_start, app_click, app_screenshot, run_end. " +
        "ws_affected: what to rebuild; symbol_refs: every use of a name; com_describe: a DLL/OCX; batch_run: several steps in one call. " +
        "Read/Grep on sources is fine. Windows only: files you write use CR LF; VB strings have no escapes (vbCrLf, not \\n; \"C:\\Temp\", not \"C:\\\\Temp\").";

    private static int s_inFlight;

    /// <summary>
    /// Tool calls running now. server_restart waits for them before ending the process: a call the client sent right after
    /// it (an ide_launch, 2026-10-17) was cut off mid-way and the client saw "Connection closed".
    /// </summary>
    public static int InFlight => Volatile.Read(ref s_inFlight);

    private readonly SemaphoreSlim _writeLock = new(1, 1);
    private readonly Dictionary<string, Tool> _tools = tools.ToDictionary(t => t.Name);

    /// <summary>Our own requests to the client that wait for a response, by id ("vb6mcp-1", ...).</summary>
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, TaskCompletionSource<JsonObject>> _outgoing = new();
    private int _nextRequestId;

    public async Task RunAsync(CancellationToken ct = default)
    {
        Current = this;
        var pending = new List<Task>();
        string? line;
        while ((line = await input.ReadLineAsync(ct).ConfigureAwait(false)) != null)
        {
            if (line.Length == 0) continue;
            Log("<- " + line);
            var request = line;   // capture this iteration's value, not the loop variable
            pending.Add(Task.Run(() => HandleLineAsync(request, ct), ct));
            pending.RemoveAll(t => t.IsCompleted);
        }
        // nobody can answer our requests any more
        foreach (var id in _outgoing.Keys)
            if (_outgoing.TryRemove(id, out var waiting)) waiting.TrySetException(new IOException("the client closed the connection"));
        await Task.WhenAll(pending).ConfigureAwait(false);
        Log("stdin closed");
    }

    /// <summary>
    /// Sends the client a JSON-RPC request and waits for its response: the "result" object, or an exception for a JSON-RPC
    /// error, the timeout or a closed connection. Tool calls run concurrently with the read loop, so a tool can await this
    /// while the loop goes on reading (the response arrives as a line without "method").
    /// </summary>
    public async Task<JsonObject> RequestAsync(string method, JsonObject @params, TimeSpan timeout, CancellationToken ct)
    {
        var id = "vb6mcp-" + Interlocked.Increment(ref _nextRequestId);
        var tcs = new TaskCompletionSource<JsonObject>(TaskCreationOptions.RunContinuationsAsynchronously);
        _outgoing[id] = tcs;
        try
        {
            await SendAsync(new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["method"] = method, ["params"] = @params }).ConfigureAwait(false);
            var response = await tcs.Task.WaitAsync(timeout, ct).ConfigureAwait(false);
            if (response["error"] is JsonObject err)
                throw new McpRequestException($"{method} failed: {(string?)err["message"] ?? err.ToJsonString()} ({err["code"]})");
            return response["result"] as JsonObject ?? [];
        }
        catch (TimeoutException) { throw new McpRequestException($"{method}: no answer from the client within {timeout.TotalSeconds:0} s"); }
        finally { _outgoing.TryRemove(id, out _); }
    }

    private void CompleteOutgoing(JsonObject msg)
    {
        var id = msg["id"] is JsonValue v ? v.ToString() : null;
        if (id is not null && _outgoing.TryRemove(id, out var tcs)) tcs.TrySetResult(msg);
        else Log("response to an unknown request id: " + id);
    }

    private async Task HandleLineAsync(string line, CancellationToken ct)
    {
        JsonObject? msg;
        try { msg = JsonNode.Parse(line) as JsonObject; }
        catch (JsonException ex) { await SendAsync(Error(null, -32700, "Parse error: " + ex.Message)).ConfigureAwait(false); return; }
        if (msg is null) { await SendAsync(Error(null, -32600, "Invalid request")).ConfigureAwait(false); return; }

        var method = (string?)msg["method"];
        if (method is null) { CompleteOutgoing(msg); return; }   // a response to one of our requests (elicitation)
        var id = msg["id"]?.DeepClone();
        bool isNotification = !msg.ContainsKey("id");

        JsonObject? reply;
        try { reply = await DispatchAsync(method, msg["params"] as JsonObject, id, ct).ConfigureAwait(false); }
        catch (Exception ex) { reply = Error(id, -32603, "Internal error: " + ex.Message); }

        if (isNotification) return;                  // never answer notifications
        // Unknown methods must get -32601: the desktop client probes server/discover and falls back on it (spike 1)
        await SendAsync(reply ?? Error(id, -32601, "Method not found: " + method)).ConfigureAwait(false);
    }

    private async Task<JsonObject?> DispatchAsync(string method, JsonObject? p, JsonNode? id, CancellationToken ct)
    {
        switch (method)
        {
            case "initialize":
                var requested = (string?)p?["protocolVersion"];
                if ((string?)p?["clientInfo"]?["name"] is { Length: > 0 } client)
                {
                    ClientId = client;
                    ClientName = (client + " " + ((string?)p?["clientInfo"]?["version"] ?? "")).Trim();
                }
                ClientElicits = OffersFormElicitation(p?["capabilities"] as JsonObject);
                return Result(id, new JsonObject
                {
                    ["protocolVersion"] = requested is not null && SupportedVersions.Contains(requested) ? requested : SupportedVersions[0],
                    ["capabilities"] = new JsonObject { ["tools"] = new JsonObject() },
                    ["serverInfo"] = new JsonObject { ["name"] = ServerName, ["version"] = ServerVersion },
                    ["instructions"] = Instructions,
                });
            case "ping":
                return Result(id, new JsonObject());
            case "tools/list":
                return Result(id, new JsonObject { ["tools"] = new JsonArray(_tools.Values.Select(t => (JsonNode)t.Definition()).ToArray()) });
            case "tools/call":
                var name = (string?)p?["name"] ?? "";
                var args = p?["arguments"] as JsonObject ?? [];
                Interlocked.Increment(ref s_inFlight);
                try
                {
                    var result = _tools.TryGetValue(name, out var tool)
                        ? await tool.InvokeAsync(args, ct).ConfigureAwait(false)
                        : ToolResult.Fail($"Unknown tool: {name}");
                    return Result(id, result.ToJson());
                }
                finally { Interlocked.Decrement(ref s_inFlight); }
            default:
                return null;
        }
    }

    private async Task SendAsync(JsonObject reply)
    {
        var json = reply.ToJsonString(Json);
        await _writeLock.WaitAsync().ConfigureAwait(false);
        try
        {
            await output.WriteLineAsync(json).ConfigureAwait(false);
            await output.FlushAsync().ConfigureAwait(false);
        }
        finally { _writeLock.Release(); }
        Log("-> " + (json.Length > 500 ? json[..500] + "..." : json));
    }

    private void Log(string message)
    {
        lock (log) { log.WriteLine($"{DateTime.Now:HH:mm:ss.fff} {message}"); log.Flush(); }
    }

    /// <summary>
    /// Asks the user through the client (MCP elicitation/create, form mode, protocol 2025-06-18): the client shows
    /// <paramref name="message"/> and a form for <paramref name="requestedSchema"/>; the agent never sees it. Returns the
    /// result: "action" (accept, decline, cancel) and, on accept, "content".
    /// </summary>
    public Task<JsonObject> ElicitAsync(string message, JsonObject requestedSchema, TimeSpan timeout, CancellationToken ct) =>
        RequestAsync("elicitation/create", new JsonObject { ["message"] = message, ["requestedSchema"] = requestedSchema }, timeout, ct);

    private static JsonObject Result(JsonNode? id, JsonNode result) => new() { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = result };

    private static JsonObject Error(JsonNode? id, int code, string message) =>
        new() { ["jsonrpc"] = "2.0", ["id"] = id, ["error"] = new JsonObject { ["code"] = code, ["message"] = message } };
}

/// <summary>A request we sent the client failed: a JSON-RPC error, no answer in time, or the connection closed.</summary>
public sealed class McpRequestException(string message) : Exception(message);
