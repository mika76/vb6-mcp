using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>symbol_refs (PLAN.md C5, improvement notes 4): statements joined across continuations, strings and comments ignored.</summary>
public static class RefsTests
{
    const string Code =
        "Public Sub Add_To_Audit(ByVal Action As String, Optional ByVal Comment As String = \"\")\r\n" +
        "End Sub\r\n" +
        "Private Sub Pay()\r\n" +
        "    Add_To_Audit \"Pay\", _\r\n" +
        "        \"amount \" & _\r\n" +
        "        CStr(x)\r\n" +
        "    MsgBox \"Add_To_Audit failed\"   ' Add_To_Audit in a comment\r\n" +
        "    Rem Add_To_Audit here too\r\n" +
        "    Call modAudit.add_to_audit(\"x\"): y = 1\r\n" +
        "    Add_To_AuditEx \"no\"\r\n" +
        "End Sub";

    [Test]
    public static void CallsAreFoundWithTheirContinuationsJoined()
    {
        var refs = VbRefs.Find(Code, "Add_To_Audit");
        Assert.Equal(3, refs.Count);
        Assert.True(refs[0].Declaration && refs[0].Line == 1, "the declaration");
        Assert.Equal(4, refs[1].Line);
        Assert.Equal(6, refs[1].EndLine);
        Assert.Equal("Add_To_Audit \"Pay\", \"amount \" & CStr(x)", refs[1].Statement);
        Assert.True(!refs[1].Declaration, "a call");
        Assert.Equal(9, refs[2].Line);   // qualified and in another case; the string, comment, Rem and the longer name don't count
    }

    [Test]
    public static void OnlyIdentifiersAreAccepted()
    {
        Assert.True(VbRefs.IsIdentifier("x_1") && !VbRefs.IsIdentifier("a.b") && !VbRefs.IsIdentifier("1x") && !VbRefs.IsIdentifier("a b"), "identifier check");
    }

    sealed class NoIde : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => throw new InvalidOperationException();
        public AddinClient Client(InstanceInfo instance) => throw new InvalidOperationException();
    }

    [Test]
    public static async Task TheToolScansEveryProjectAndNamesThem()
    {
        var root = Directory.CreateTempSubdirectory("vb6mcp-refs-").FullName;
        try
        {
            var shared = Path.Combine(root, "Shared"); Directory.CreateDirectory(shared);
            File.WriteAllBytes(Path.Combine(shared, "modAudit.bas"), Vb1252.Encoding.GetBytes("Attribute VB_Name = \"modAudit\"\r\n" + Code + "\r\n"));
            foreach (var p in new[] { "A", "B" })
            {
                var dir = Path.Combine(root, p); Directory.CreateDirectory(dir);
                File.WriteAllBytes(Path.Combine(dir, p + ".vbp"), Vb1252.Encoding.GetBytes($"Type=Exe\r\nModule=modAudit; ..\\Shared\\modAudit.bas\r\nName=\"{p}\"\r\n"));
            }
            var tool = new WorkspaceTools(new NoIde(), [root]).All().First(t => t.Name == "symbol_refs");
            var r = (await tool.InvokeAsync(new JsonObject { ["name"] = "Add_To_Audit", ["declarations"] = false }, CancellationToken.None)).ToJson();
            var o = JsonNode.Parse((string)r["content"]![0]!["text"]!)!;
            Assert.Equal(2, (int)o["count"]!);                       // the shared file once, without the declaration
            Assert.Equal(1, o["files"]!.AsArray().Count);
            Assert.Equal("A,B", string.Join(",", o["files"]![0]!["projects"]!.AsArray().Select(n => (string)n!).Order()));
            Assert.Equal(4, (int)o["files"]![0]!["refs"]![0]!["line"]!);   // IDE numbering: the Attribute line isn't counted
        }
        finally { Directory.Delete(root, true); }
    }
}
