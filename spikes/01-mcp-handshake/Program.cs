// Spike 1: minimal hand-written MCP server (JSON-RPC 2.0 over stdio, newline-delimited).
// No SDK, BCL only. One tool: vb6_processes.
// stdout carries protocol messages only; all diagnostics go to stderr.

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Encodings.Web;
using System.Text.Json.Nodes;

const string ServerName = "vb6-mcp-spike";
const string ServerVersion = "0.0.1";
string[] supportedVersions = ["2025-06-18", "2025-03-26", "2024-11-05"];

var stdin = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true, NewLine = "\n" };
var jsonOut = new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
// Diagnostics: stderr, or a file when VB6MCP_LOG is set (handy when a client hides stderr).
var logPath = Environment.GetEnvironmentVariable("VB6MCP_LOG");
TextWriter log = logPath is null ? Console.Error : new StreamWriter(logPath, append: true) { AutoFlush = true };

log.WriteLine($"[{ServerName}] started, pid {Environment.ProcessId}");

string? line;
while ((line = stdin.ReadLine()) != null)
{
    if (line.Length == 0) continue;
    log.WriteLine($"<- {line}");
    JsonObject? msg;
    try { msg = JsonNode.Parse(line) as JsonObject; }
    catch (JsonException ex) { Send(Error(null, -32700, "Parse error: " + ex.Message)); continue; }
    if (msg is null) { Send(Error(null, -32600, "Invalid request")); continue; }

    var id = msg["id"]?.DeepClone();
    var method = (string?)msg["method"];
    var prms = msg["params"] as JsonObject;

    if (method is null) continue;          // a response to us; we never send requests
    bool isNotification = !msg.ContainsKey("id");

    JsonObject? reply;
    try { reply = Handle(method, prms, id); }
    catch (Exception ex) { reply = Error(id, -32603, "Internal error: " + ex.Message); }

    if (isNotification) continue;          // never answer notifications
    Send(reply ?? Error(id, -32601, "Method not found: " + method));
}
log.WriteLine($"[{ServerName}] stdin closed, exiting");

JsonObject? Handle(string method, JsonObject? p, JsonNode? id) => method switch
{
    "initialize" => Result(id, new JsonObject
    {
        ["protocolVersion"] = PickVersion((string?)p?["protocolVersion"]),
        ["capabilities"] = new JsonObject { ["tools"] = new JsonObject() },
        ["serverInfo"] = new JsonObject { ["name"] = ServerName, ["version"] = ServerVersion },
    }),
    "notifications/initialized" => null,
    "ping" => Result(id, new JsonObject()),
    "tools/list" => Result(id, new JsonObject { ["tools"] = new JsonArray(ToolDef()) }),
    "tools/call" => Result(id, CallTool((string?)p?["name"], p?["arguments"] as JsonObject)),
    _ => null,
};

string PickVersion(string? requested) =>
    requested is not null && supportedVersions.Contains(requested) ? requested : supportedVersions[0];

JsonObject ToolDef() => new()
{
    ["name"] = "vb6_processes",
    ["description"] = "List running VB6 IDE processes (VB6.EXE) with pid, main window title and whether the process is elevated.",
    ["inputSchema"] = new JsonObject { ["type"] = "object", ["properties"] = new JsonObject() },
};

JsonObject CallTool(string? name, JsonObject? args)
{
    if (name != "vb6_processes")
        return ToolText($"Unknown tool: {name}", isError: true);

    var list = new JsonArray();
    foreach (var proc in Process.GetProcessesByName("VB6"))
    {
        list.Add(new JsonObject
        {
            ["pid"] = proc.Id,
            ["title"] = proc.MainWindowTitle,
            ["elevated"] = Token.IsElevated(proc.Id),
        });
    }
    return ToolText(new JsonObject { ["count"] = list.Count, ["processes"] = list }.ToJsonString(jsonOut));
}

static JsonObject ToolText(string text, bool isError = false) => new()
{
    ["content"] = new JsonArray(new JsonObject { ["type"] = "text", ["text"] = text }),
    ["isError"] = isError,
};

static JsonObject Result(JsonNode? id, JsonNode result) =>
    new() { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = result };

static JsonObject Error(JsonNode? id, int code, string message) =>
    new() { ["jsonrpc"] = "2.0", ["id"] = id, ["error"] = new JsonObject { ["code"] = code, ["message"] = message } };

void Send(JsonObject reply)
{
    var json = reply.ToJsonString(jsonOut);
    log.WriteLine($"-> {json}");
    stdout.WriteLine(json);
}

static class Token
{
    const uint ProcessQueryLimitedInformation = 0x1000, TokenQuery = 0x8;
    const int TokenIntegrityLevel = 25;

    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool OpenProcessToken(IntPtr proc, uint access, out IntPtr token);
    [DllImport("advapi32.dll", SetLastError = true)] static extern bool GetTokenInformation(IntPtr token, int cls, IntPtr buf, int len, out int ret);
    [DllImport("advapi32.dll")] static extern IntPtr GetSidSubAuthority(IntPtr sid, uint index);
    [DllImport("advapi32.dll")] static extern IntPtr GetSidSubAuthorityCount(IntPtr sid);

    // true = high integrity or above, false = below, null = could not determine
    public static bool? IsElevated(int pid)
    {
        IntPtr proc = OpenProcess(ProcessQueryLimitedInformation, false, pid);
        if (proc == IntPtr.Zero) return null;
        try
        {
            if (!OpenProcessToken(proc, TokenQuery, out var token)) return null;
            try
            {
                GetTokenInformation(token, TokenIntegrityLevel, IntPtr.Zero, 0, out int size);
                IntPtr buf = Marshal.AllocHGlobal(size);
                try
                {
                    if (!GetTokenInformation(token, TokenIntegrityLevel, buf, size, out _)) return null;
                    IntPtr sid = Marshal.ReadIntPtr(buf);
                    int count = Marshal.ReadByte(GetSidSubAuthorityCount(sid));
                    int rid = Marshal.ReadInt32(GetSidSubAuthority(sid, (uint)(count - 1)));
                    return rid >= 0x3000;
                }
                finally { Marshal.FreeHGlobal(buf); }
            }
            finally { CloseHandle(token); }
        }
        finally { CloseHandle(proc); }
    }
}
