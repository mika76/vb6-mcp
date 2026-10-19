using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>Edits whose identifiers VB re-cased to the project's spelling (improvement notes 23, 24).</summary>
public static class CaseChangeTests
{
    [Test]
    public static void ARetypedNameThatVbKeptIsReported()
    {
        var module = "Private Sub A()\r\n    MsgBox Err.number\r\nEnd Sub";
        var changes = VbCode.CaseChanges("    MsgBox Err.Number", module);
        Assert.Equal("Number -> number", string.Join("|", changes));
    }

    [Test]
    public static void MatchingSpellingAndStringsAndCommentsAreIgnored()
    {
        var module = "Private Sub A()\r\n    x = \"Number\" ' Number here\r\n    y = Err.Number\r\nEnd Sub";
        Assert.Equal(0, VbCode.CaseChanges("    y = Err.Number\r\n    z = \"NUMBER\" ' NUMBER", module).Count);
    }

    [Test]
    public static void CaseOnlyDifferencesListTheRespelledNames()
    {
        Assert.Equal("Number -> number|CMD -> cmd", string.Join("|", VbCode.CaseOnlyRespellings("x = Err.Number: CMD = 1", "x = Err.number: cmd = 1")!));
        Assert.True(VbCode.CaseOnlyRespellings("x = 1", "x = 2") is null, "more than case");
    }

    /// <summary>C-style escapes in VB string literals (owner, 2026-10-17): VB6 has none.</summary>
    [Test]
    public static void CStyleEscapesInStringLiteralsAreHinted()
    {
        Assert.Equal(1, VbCode.LiteralHints("MsgBox \"line1\\nline2\"").Count);
        Assert.True(VbCode.LiteralHints("f = \"C:\\\\Temp\\\\x.txt\"")[0].Contains("single backslashes"), "doubled backslashes");
        Assert.Equal(2, VbCode.LiteralHints("a = \"x\\ty\": b = \"D:\\\\data\"").Count);
        foreach (var fine in new[] { "f = \"C:\\Temp\\new.txt\"", "u = \"\\\\server\\share\"", "MsgBox \"a\" & vbCrLf & \"b\"", "' comment with \\n", "s = \"say \"\"hi\"\"\"", "s = \"\"", "If s = \"\" Then s = \"x\"", "s = \"\\\"" })
            Assert.Equal(0, VbCode.LiteralHints(fine).Count, fine);
    }

    const char US = Wire.US, RS = Wire.RS;

    sealed class Env(FakeAddin addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "P - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    /// <summary>sync_diff without a component: the project summary (one same, one re-cased, one really changed).</summary>
    [Test]
    public static async Task TheProjectSummaryShowsCaseOnlyComponents()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-case-").FullName;
        try
        {
            string F(string n, string code) { var p = Path.Combine(dir, n + ".bas"); File.WriteAllBytes(p, Vb1252.Encoding.GetBytes($"Attribute VB_Name = \"{n}\"\r\n{code}\r\n")); return p; }
            var a = F("modA", "x = 1"); var b = F("modB", "y = Err.Number"); var c = F("modC", "z = 1");
            var ide = new Dictionary<string, string> { ["modA"] = "x = 1", ["modB"] = "y = Err.number", ["modC"] = "z = 2\r\nw = 3" };
            using var fake = new FakeAddin((verb, args) => verb switch
            {
                "PROJECTS" => ("OK", $"1{US}P{US}0{US}{Path.Combine(dir, "P.vbp")}{US}0"),
                "COMPONENTS" => ("OK", $"modA{US}1{US}0{US}1{US}{a}{RS}modB{US}1{US}1{US}1{US}{b}{RS}modC{US}1{US}1{US}1{US}{c}"),
                "READ" => ("OK", $"1{US}0{US}{ide[args[1]]}"),
                _ => ("ERR", $"400{US}unexpected {verb}"),
            });
            var r = (await new IdeTools(new Env(fake)).All().First(t => t.Name == "sync_diff").InvokeAsync(new JsonObject { ["project"] = "P" }, CancellationToken.None)).ToJson();
            var o = JsonNode.Parse((string)r["content"]![0]!["text"]!)!;
            Assert.True(!(bool)r["isError"]!, o.ToJsonString());
            Assert.Equal(1, (int)o["same"]!); Assert.Equal(1, (int)o["caseOnly"]!); Assert.Equal(1, (int)o["differs"]!);
            var rows = o["components"]!.AsArray();
            Assert.Equal("Number -> number", (string)rows.First(x => (string)x!["component"]! == "modB")!["respelled"]![0]!);
        }
        finally { Directory.Delete(dir, true); }
    }
}
