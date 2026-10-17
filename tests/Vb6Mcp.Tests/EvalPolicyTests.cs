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

        // no hook guarding debug_eval: refused even with allow_side_effect (an agent could pass it itself)
        var noHook = new IdeTools(new Env(fake)) { EvalGuarded = () => false };
        var text = await Eval(noHook, allow: true);
        Assert.True(text.Contains("side effect") && text.Contains("hook"), text);
        // hook installed: without allow it says how to ask; with allow (the user approved the prompt) it goes ahead
        var hooked = new IdeTools(new Env(fake)) { EvalGuarded = () => true };
        Assert.True((await Eval(hooked, allow: false)).Contains("allow_side_effect=true"), "says how to ask");
        Assert.True(!verbs.Contains("IMMEDIATE") && !verbs.Contains("IMMREAD"), "nothing sent to the Immediate window so far: " + string.Join(",", verbs));
        Assert.True(!(await Eval(hooked, allow: true)).StartsWith("refused", StringComparison.Ordinal), "approved: not refused");
    }
}
