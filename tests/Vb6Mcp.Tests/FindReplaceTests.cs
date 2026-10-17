using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>find/replace in code_replace_lines / file_replace_lines (PLAN.md A6, improvement notes 5).</summary>
public static class FindReplaceTests
{
    const string Code =
        "Private Sub Send()\r\n" +
        "    m_comm.Output = \"m_comm.Output\"   ' m_comm.Output\r\n" +
        "    x = m_comm.OUTPUT & m_comm.OutputLen\r\n" +
        "End Sub";

    [Test]
    public static void OnlyCodeAndWholeWordsChangeByDefault()
    {
        var found = FindReplace.Compute(Code, "m_comm.Output", "m_comm.OutputText");
        Assert.Equal(2, found.Count);
        Assert.Equal("    m_comm.OutputText = \"m_comm.Output\"   ' m_comm.Output", found[0].New);
        Assert.Equal("    x = m_comm.OutputText & m_comm.OutputLen", found[1].New);   // any case; OutputLen isn't a whole-word match
        var all = FindReplace.Compute(Code, "m_comm.Output", "X", inStrings: true, inComments: true);
        Assert.Equal("    X = \"X\"   ' X", all[0].New);
        Assert.Equal(1, FindReplace.Compute(Code, "m_comm.Output", "X", matchCase: true).Count);
    }

    [Test]
    public static void BadArgumentsAreRefused()
    {
        foreach (var args in new[]
        {
            new JsonObject { ["find"] = "a", ["replace"] = "b", ["start_line"] = 1 },
            new JsonObject { ["find"] = "a" },
        })
            try { FindReplace.Args(args); Assert.True(false, args.ToJsonString()); } catch (ToolException) { }
        try { FindReplace.Compute(Code, "", "x"); Assert.True(false, "empty find"); } catch (ToolException) { }
    }

    sealed class NoIde : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => throw new InvalidOperationException();
        public AddinClient Client(InstanceInfo instance) => throw new InvalidOperationException();
    }

    [Test]
    public static async Task AClosedFileIsPreviewedThenChangedInOneCall()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-fr-").FullName;
        try
        {
            var path = Path.Combine(dir, "modComm.bas");
            File.WriteAllBytes(path, Vb1252.Encoding.GetBytes("Attribute VB_Name = \"modComm\"\r\n" + Code + "\r\n"));
            var before = File.ReadAllBytes(path);
            var tool = new WorkspaceTools(new NoIde(), [dir]).All().First(t => t.Name == "file_replace_lines");
            var hash = VbCode.Hash(VbSourceFile.Load(path).Code);
            JsonObject Args(bool apply) => new() { ["path"] = path, ["find"] = "m_comm.Output", ["replace"] = "m_comm.OutputText", ["expected_hash"] = hash, ["apply"] = apply };

            var preview = JsonNode.Parse((string)(await tool.InvokeAsync(Args(false), CancellationToken.None)).ToJson()["content"]![0]!["text"]!)!;
            Assert.Equal(2, (int)preview["preview"]!["lines"]!);
            Assert.True(File.ReadAllBytes(path).AsSpan().SequenceEqual(before), "a preview changes nothing");

            var r = (await tool.InvokeAsync(Args(true), CancellationToken.None)).ToJson();
            Assert.True(!(bool)r["isError"]!, r.ToJsonString());
            var text = Vb1252.Encoding.GetString(File.ReadAllBytes(path));
            Assert.True(text.Contains("m_comm.OutputText = \"m_comm.Output\"") && text.Contains("x = m_comm.OutputText & m_comm.OutputLen") && text.StartsWith("Attribute VB_Name"), text);
        }
        finally { Directory.Delete(dir, true); }
    }
}
