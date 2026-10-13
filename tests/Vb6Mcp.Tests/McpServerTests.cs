using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

public static class McpServerTests
{
    const char US = Wire.US, RS = Wire.RS;

    /// <summary>Feeds JSON-RPC lines to a server and returns the responses keyed by id.</summary>
    static async Task<Dictionary<string, JsonObject>> Run(IReadOnlyList<Tool> tools, params string[] lines)
    {
        var output = new StringWriter();
        await new McpServer(new StringReader(string.Join("\n", lines) + "\n"), output, TextWriter.Null, tools).RunAsync();
        return output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(l => (JsonObject)JsonNode.Parse(l)!)
            .ToDictionary(o => o["id"]?.ToJsonString() ?? "null");
    }

    [Test]
    public static async Task HandshakeListAndUnknownMethods()
    {
        var r = await Run([],
            """{"jsonrpc":"2.0","id":"server-discover-probe-1","method":"server/discover","params":{}}""",
            """{"jsonrpc":"2.0","id":0,"method":"initialize","params":{"protocolVersion":"2025-11-25","capabilities":{}}}""",
            """{"jsonrpc":"2.0","method":"notifications/initialized"}""",
            """{"jsonrpc":"2.0","id":1,"method":"ping"}""",
            """{"jsonrpc":"2.0","id":2,"method":"tools/list"}""",
            """{bad json""");

        // the desktop client's probe must get -32601 so it falls back to initialize (spike 1)
        Assert.Equal(-32601, (int)r["\"server-discover-probe-1\""]["error"]!["code"]!);
        // unknown requested version -> our newest
        Assert.Equal(McpServer.SupportedVersions[0], (string)r["0"]["result"]!["protocolVersion"]!);
        Assert.True(r["1"]["result"] is JsonObject, "ping result");
        Assert.Equal(0, r["2"]["result"]!["tools"]!.AsArray().Count);
        Assert.Equal(-32700, (int)r["null"]["error"]!["code"]!);
        Assert.Equal(5, r.Count, "the notification got no reply");
    }

    [Test]
    public static async Task ToolsAreListedWithSchemas()
    {
        var r = await Run(new IdeTools(new FakeEnv(null)).All(), """{"jsonrpc":"2.0","id":1,"method":"tools/list"}""");
        var names = r["1"]["result"]!["tools"]!.AsArray().Select(t => (string)t!["name"]!).ToList();
        Assert.SequenceEqual(new[] { "ide_list", "ide_status", "project_components", "code_read", "symbol_outline", "project_open", "proc_get",
                                    "code_replace_lines", "code_insert", "proc_replace", "project_save", "ide_dialog_respond" }, names);
        var codeRead = r["1"]["result"]!["tools"]!.AsArray().First(t => (string)t!["name"]! == "code_read")!;
        Assert.Equal("component", (string)codeRead["inputSchema"]!["required"]![0]!);
    }

    [Test]
    public static async Task CodeReadGoesThroughTheAddin()
    {
        using var fake = new FakeAddin((verb, args) => verb switch
        {
            "PROJECTS" => ("OK", $"1{US}Project1{US}0{US}D:\\p\\Project1.vbp{US}1"),
            "READ" => ("OK", $"5{US}1{US}Option Explicit"),
            _ => ("ERR", $"400{US}unknown"),
        });
        var tools = new IdeTools(new FakeEnv(fake)).All();
        var r = await Run(tools, """{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"code_read","arguments":{"component":"Form1"}}}""");
        var res = r["1"]["result"]!;
        Assert.Equal(false, (bool)res["isError"]!);
        var payload = JsonNode.Parse((string)res["content"]![0]!["text"]!)!;
        Assert.Equal("ide", (string)payload["source"]!);
        Assert.Equal("Project1", (string)payload["project"]!, "project defaulted to the only one");
        Assert.Equal(true, (bool)payload["dirty"]!);
        Assert.Equal("Option Explicit", (string)payload["text"]!);
    }

    [Test]
    public static async Task ToolErrorsAreIsErrorResults()
    {
        var tools = new IdeTools(new FakeEnv(null)).All();
        var r = await Run(tools,
            """{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"code_read","arguments":{"component":"Form1"}}}""",
            """{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"nope","arguments":{}}}""");
        Assert.Equal(true, (bool)r["1"]["result"]!["isError"]!);
        Assert.True(((string)r["1"]["result"]!["content"]![0]!["text"]!).Contains("no VB6 IDE"), "explains that no IDE is running");
        Assert.Equal(true, (bool)r["2"]["result"]!["isError"]!);
    }

    [Test]
    public static async Task ProjectOpenQueuesAndWaitsForTheProject()
    {
        var vbp = Path.Combine(Path.GetTempPath(), $"vb6mcp-open-{Guid.NewGuid():N}.vbp");
        File.WriteAllText(vbp, "Type=Exe");
        try
        {
            bool opened = false;
            using var fake = new FakeAddin((verb, args) => verb switch
            {
                "PROJECTS" => ("OK", opened ? $"1{US}Basic{US}0{US}{vbp}{US}0" : ""),
                "OPEN" => args[1] == "add" ? Open() : ("ERR", $"400{US}unexpected mode"),
                _ => ("ERR", $"400{US}unknown"),
            });
            (string, string) Open() { opened = true; return ("QUEUED", "open"); }

            var tools = new IdeTools(new FakeEnv(fake)) { OpenWait = TimeSpan.FromSeconds(3) }.All();
            var r = await Run(tools, Call(1, "project_open", new JsonObject { ["path"] = vbp }));
            var res = r["1"]["result"]!;
            Assert.Equal(false, (bool)res["isError"]!, (string?)res["content"]![0]!["text"]);
            var payload = JsonNode.Parse((string)res["content"]![0]!["text"]!)!;
            Assert.Equal("opened", (string)payload["status"]!);
            Assert.Equal("Basic", (string)payload["project"]!);
        }
        finally { File.Delete(vbp); }
    }

    [Test]
    public static async Task ProjectOpenNeedsModeWhenProjectsAreOpen()
    {
        var vbp = Path.Combine(Path.GetTempPath(), $"vb6mcp-open-{Guid.NewGuid():N}.vbp");
        File.WriteAllText(vbp, "Type=Exe");
        try
        {
            using var fake = new FakeAddin((verb, _) => verb == "PROJECTS" ? ("OK", $"1{US}Other{US}0{US}D:\\x\\Other.vbp{US}1") : ("ERR", $"400{US}x"));
            var tools = new IdeTools(new FakeEnv(fake)).All();
            var r = await Run(tools,
                Call(1, "project_open", new JsonObject { ["path"] = vbp }),
                Call(2, "project_open", new JsonObject { ["path"] = vbp, ["mode"] = "replace" }));
            Assert.True(((string)r["1"]["result"]!["content"]![0]!["text"]!).Contains("pass mode"), "asks for a mode");
            Assert.True(((string)r["2"]["result"]!["content"]![0]!["text"]!).Contains("unsaved"), "refuses replace over unsaved work");
        }
        finally { File.Delete(vbp); }
    }

    /// <summary>A tools/call request line.</summary>
    static string Call(int id, string tool, JsonObject args)
        => new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["method"] = "tools/call", ["params"] = new JsonObject { ["name"] = tool, ["arguments"] = args } }.ToJsonString();

    /// <summary>One fake IDE (or none) backed by a <see cref="FakeAddin"/>.</summary>
    sealed class FakeEnv(FakeAddin? addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new(addin is null ? [] : [addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "Project1 - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }
}
