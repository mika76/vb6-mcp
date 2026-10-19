using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>debug_eval's database guard (PLAN.md E5): what counts as a write, and readonly refusing it.</summary>
public static class EvalPolicyTests
{
    [Test]
    public static void WritesAreSpottedAndReadsAreNot()
    {
        foreach (var w in new[]
        {
            "SqlCon.Execute(\"update orders set retries = 3 where order_number = 183\")",   // T12 batch run 2, verbatim shape
            "SqlCon.Execute(\"UPDATE dbo.orders SET active = 0\")",
            "cn.Execute(\"insert into log values (1)\")",
            "cn.Execute(\"DELETE FROM orders WHERE id = 1\")",
            "cn.Execute(\"exec sp_reset_order 184\")",
            "cn.Execute(\"drop table x\")",
            "rs.AddNew", "rs.Update", "rs.Delete", "cn.CommitTrans",
        })
            Assert.True(EvalPolicy.DatabaseWrite(w) is not null, w);
        foreach (var r in new[]
        {
            "SqlCon.Execute(\"select count(*) from orders\")(0)",
            "rs.Fields(\"retries\").Value",
            "rs.RecordCount", "m_counter.Value * 2", "lastUpdate", "frmMain.cmdUpdate.Caption", "rs.UpdateCount",
        })
            Assert.True(EvalPolicy.DatabaseWrite(r) is null, r);
    }

    [Test]
    public static void OtherSideEffectsAreSpottedToo()
    {
        foreach (var w in new[]
        {
            "Shell(\"cmd /c del x\")", "CreateObject(\"WScript.Shell\").Run(\"calc\")", "GetObject(\"winmgmts:\").Get(\"Win32_Process\")",
            "CreateObject(\"Scripting.FileSystemObject\").DeleteFile(\"C:\\x\")", "Kill \"C:\\x\"", "SaveSetting \"a\", \"b\", \"c\", \"d\"",
            "CreateObject(\"WScript.Shell\").RegWrite(\"HKCU\\x\", 1)", "CreateObject(\"MSXML2.XMLHTTP\")", "URLDownloadToFile(0, u, f, 0, 0)",
            "CreateObject(\"MSScriptControl.ScriptControl\")", "SendKeys \"%{F4}\"",
        })
            Assert.True(EvalPolicy.SideEffect(w) is not null, w);
        foreach (var r in new[] { "frmMain.lblShell.Caption", "Len(sFileName)", "rs.Fields(\"Name\").Value", "Dir$(\"C:\\x\\*.txt\")", "Environ$(\"TEMP\")", "UBound(arr)" })
            Assert.True(EvalPolicy.SideEffect(r) is null, r);
    }

    sealed class Env(FakeAddin addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "P - Microsoft Visual Basic [break]", IdeMode.Break, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    [Test]
    public static async Task ReadonlyRefusesAWriteBeforeTypingAnything()
    {
        var verbs = new List<string>();
        using var fake = new FakeAddin((verb, _) => { verbs.Add(verb); return ("ERR", $"400{Wire.US}unexpected {verb}"); });
        const string update = "SqlCon.Execute(\"update orders set active = 0\")";
        async Task<string> Eval(IdeTools tools, bool allow) => (string)(await tools.All().First(t => t.Name == "debug_eval").InvokeAsync(
            new JsonObject { ["expression"] = update, ["allow_side_effect"] = allow }, CancellationToken.None)).ToJson()["content"]![0]!["text"]!;

        // Claude Code without the hook: refused even with allow_side_effect (an agent could pass it itself)
        var noHook = new IdeTools(new Env(fake)) { EvalGuarded = () => false, EvalClientId = () => "claude-code", EvalClientElicits = () => false };
        var text = await Eval(noHook, allow: true);
        Assert.True(text.Contains("side effect") && text.Contains("hook"), text);
        // Claude Code with the hook: without allow it says how to ask; with allow (the user approved the prompt) it goes ahead
        var hooked = new IdeTools(new Env(fake)) { EvalGuarded = () => true, EvalClientId = () => "claude-code", EvalClientElicits = () => false };
        Assert.True((await Eval(hooked, allow: false)).Contains("allow_side_effect=true"), "says how to ask");
        Assert.True(!verbs.Contains("IMMEDIATE") && !verbs.Contains("IMMREAD"), "nothing sent to the Immediate window so far: " + string.Join(",", verbs));
        Assert.True(!(await Eval(hooked, allow: true)).StartsWith("refused", StringComparison.Ordinal), "approved: not refused");
    }

    /// <summary>
    /// Spike 40: Copilot CLI on a machine with the Claude Code hook installed passed allow_side_effect=true and its
    /// CreateTextFile ran in the elevated IDE with no prompt. The hook counts only for Claude Code now.
    /// </summary>
    [Test]
    public static async Task TheHookDoesNotVouchForOtherClients()
    {
        var verbs = new List<string>();
        using var fake = new FakeAddin((verb, _) => { verbs.Add(verb); return ("ERR", $"400{Wire.US}unexpected {verb}"); });
        var copilot = new IdeTools(new Env(fake)) { EvalGuarded = () => true, EvalClientId = () => "copilot-cli", EvalClientElicits = () => false };
        var text = await EvalText(copilot, Fso, allow: true);
        Assert.True(text.StartsWith("refused", StringComparison.Ordinal) && text.Contains("elicitation"), text);
        Assert.True(!verbs.Contains("IMMEDIATE"), "nothing typed: " + string.Join(",", verbs));
    }

    [Test]
    public static async Task OtherClientsAreAskedThroughElicitation()
    {
        var verbs = new List<string>();
        using var fake = new FakeAddin((verb, _) => { verbs.Add(verb); return ("ERR", $"400{Wire.US}unexpected {verb}"); });
        var asked = new List<string>();
        IdeTools Client(JsonObject answer) => new(new Env(fake))
        {
            EvalGuarded = () => true, EvalClientId = () => "copilot-cli", EvalClientElicits = () => true,
            EvalElicit = (message, schema, _) =>
            {
                asked.Add(message);
                Assert.True(schema["properties"]?["approve"] is not null, schema.ToJsonString());
                return Task.FromResult(answer);
            },
        };

        // no flag: not asked, told how to ask
        var text = await EvalText(Client(Accept(true)), Fso, allow: false);
        Assert.True(text.Contains("allow_side_effect=true") && asked.Count == 0, text);
        // declined, cancelled, accepted with the box unticked: refused, nothing typed into the Immediate window
        foreach (var answer in new[] { new JsonObject { ["action"] = "decline" }, new JsonObject { ["action"] = "cancel" }, Accept(false) })
        {
            text = await EvalText(Client(answer), Fso, allow: true);
            Assert.True(text.StartsWith("refused: the side effect was not approved", StringComparison.Ordinal), text);
        }
        Assert.Equal(3, asked.Count);
        Assert.True(asked[0].Contains(Fso) && asked[0].Contains("file change"), "the form shows the expression and the category: " + asked[0]);
        Assert.True(!verbs.Contains("IMMEDIATE"), "nothing typed: " + string.Join(",", verbs));
        // the client failed to ask (an error, no answer in time): refused
        var broken = new IdeTools(new Env(fake))
        {
            EvalGuarded = () => false, EvalClientId = () => "opencode", EvalClientElicits = () => true,
            EvalElicit = (_, _, _) => throw new Vb6Mcp.Mcp.McpRequestException("elicitation/create: no answer from the client within 300 s"),
        };
        text = await EvalText(broken, Fso, allow: true);
        Assert.True(text.StartsWith("refused: the user couldn't be asked", StringComparison.Ordinal), text);
        Assert.True(!verbs.Contains("IMMEDIATE"), "nothing typed: " + string.Join(",", verbs));
        // accepted with the box ticked: it goes ahead
        text = await EvalText(Client(Accept(true)), Fso, allow: true);
        Assert.True(!text.StartsWith("refused", StringComparison.Ordinal), text);
    }

    [Test]
    public static void ApprovalDependsOnTheClient()
    {
        Assert.Equal(EvalApproval.Hook, EvalPolicy.Approval("claude-code", hookInstalled: true, clientElicits: false));
        Assert.Equal(EvalApproval.Hook, EvalPolicy.Approval("claude-code", hookInstalled: true, clientElicits: true));
        Assert.Equal(EvalApproval.None, EvalPolicy.Approval("claude-code", hookInstalled: false, clientElicits: false));
        Assert.Equal(EvalApproval.Elicitation, EvalPolicy.Approval("claude-code", hookInstalled: false, clientElicits: true));
        Assert.Equal(EvalApproval.None, EvalPolicy.Approval("copilot-cli", hookInstalled: true, clientElicits: false));
        Assert.Equal(EvalApproval.Elicitation, EvalPolicy.Approval("copilot-cli", hookInstalled: true, clientElicits: true));
        Assert.Equal(EvalApproval.None, EvalPolicy.Approval("", hookInstalled: true, clientElicits: false));
        Assert.Equal(EvalApproval.None, EvalPolicy.Approval("Claude-Code", hookInstalled: true, clientElicits: false));   // the exact name only

        Assert.True(EvalPolicy.Approves(Accept(true)), "accept + ticked");
        Assert.True(!EvalPolicy.Approves(Accept(false)), "accept, unticked");
        Assert.True(!EvalPolicy.Approves(new JsonObject { ["action"] = "accept" }), "accept, no content");
        Assert.True(!EvalPolicy.Approves(new JsonObject { ["action"] = "accept", ["content"] = new JsonObject { ["approve"] = "true" } }), "a string isn't a tick");
        Assert.True(!EvalPolicy.Approves(new JsonObject { ["action"] = "decline", ["content"] = new JsonObject { ["approve"] = true } }), "decline");
        Assert.True(!EvalPolicy.Approves(null), "null");

        static JsonObject Caps(string json) => (JsonObject)JsonNode.Parse(json)!;
        Assert.True(Vb6Mcp.Mcp.McpServer.OffersFormElicitation(Caps("""{"elicitation":{}}""")), "2025-06-18 shape");
        Assert.True(Vb6Mcp.Mcp.McpServer.OffersFormElicitation(Caps("""{"sampling":{},"elicitation":{"form":{},"url":{}}}""")), "Copilot CLI");
        Assert.True(Vb6Mcp.Mcp.McpServer.OffersFormElicitation(Caps("""{"elicitation":{"form":{"applyDefaults":true}}}""")), "OpenCode");
        Assert.True(!Vb6Mcp.Mcp.McpServer.OffersFormElicitation(Caps("""{"elicitation":{"url":{}}}""")), "url only");
        Assert.True(!Vb6Mcp.Mcp.McpServer.OffersFormElicitation(Caps("""{"roots":{"listChanged":true}}""")), "Claude Code 2.1.300");
        Assert.True(!Vb6Mcp.Mcp.McpServer.OffersFormElicitation(null), "none");
    }

    const string Fso = "CreateObject(\"Scripting.FileSystemObject\").CreateTextFile(\"C:\\Temp\\x.txt\") Is Nothing";

    static JsonObject Accept(bool approve) => new() { ["action"] = "accept", ["content"] = new JsonObject { ["approve"] = approve } };

    static async Task<string> EvalText(IdeTools tools, string expression, bool allow) => (string)(await tools.All().First(t => t.Name == "debug_eval").InvokeAsync(
        new JsonObject { ["expression"] = expression, ["allow_side_effect"] = allow }, CancellationToken.None)).ToJson()["content"]![0]!["text"]!;
}
