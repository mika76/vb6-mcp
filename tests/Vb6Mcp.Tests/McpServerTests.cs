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
        Assert.SequenceEqual(new[] { "ide_list", "ide_status", "project_components", "code_read", "symbol_outline", "symbol_find", "project_open", "proc_get",
                                    "code_replace_lines", "code_insert", "proc_replace", "project_save", "ide_dialog_respond", "ide_launch",
                                    "run_start", "run_break", "run_continue", "run_end", "step_into", "step_over", "step_out", "run_to_cursor",
                                    "bp_toggle", "bp_clear_all", "debug_where", "debug_eval", "debug_immediate", "syntax_check", "debug_locals", "debug_callstack", "debug_error",
                                    "ide_kill", "recovery_list", "recovery_read", "recovery_restore", "recovery_discard",
                                    "app_windows", "app_tree", "app_click", "app_set_text", "app_mouse", "app_scroll", "app_screenshot" }, names);
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
    public static async Task SymbolFindSearchesAllComponents()
    {
        using var fake = new FakeAddin((verb, args) => verb switch
        {
            "PROJECTS" => ("OK", $"1{US}Basic{US}0{US}D:\\b\\Basic.vbp{US}0"),
            "COMPONENTS" => ("OK", $"modMath{US}1{US}0{US}13{US}D:\\b\\modMath.bas{RS}res{US}4{US}0{US}-1{US}D:\\b\\x.res{RS}clsCounter{US}2{US}0{US}18{US}D:\\b\\clsCounter.cls"),
            "MEMBERS" => args[1] switch
            {
                "modMath" => ("OK", $"AddNumbers{US}1{US}2{US}0{US}7{US}0{US}5{US}7{US}6{RS}g_callCount{US}3{US}2{US}1{US}4{US}{US}{US}{US}"),
                "clsCounter" => ("OK", $"Increment{US}1{US}2{US}0{US}16{US}0{US}15{US}16{US}4"),
                _ => ("ERR", $"500{US}members of a resource file must not be requested"),
            },
            _ => ("ERR", $"400{US}unknown"),
        });
        var tools = new IdeTools(new FakeEnv(fake)).All();
        var r = await Run(tools,
            Call(1, "symbol_find", new JsonObject { ["name"] = "addnumbers" }),
            Call(2, "symbol_find", new JsonObject { ["name"] = "c", ["contains"] = true }));
        var exact = JsonNode.Parse((string)r["1"]["result"]!["content"]![0]!["text"]!)!;
        Assert.Equal(1, exact["matches"]!.AsArray().Count);
        Assert.Equal("modMath", (string)exact["matches"]![0]!["component"]!);
        Assert.Equal(7, (int)exact["matches"]![0]!["line"]!);
        var partial = JsonNode.Parse((string)r["2"]["result"]!["content"]![0]!["text"]!)!;
        Assert.SequenceEqual(new[] { "g_callCount", "Increment" }, partial["matches"]!.AsArray().Select(m => (string)m!["Name"]!));
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
