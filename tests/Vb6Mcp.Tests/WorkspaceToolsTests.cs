using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

public static class WorkspaceToolsTests
{
    sealed class NoIde : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => throw new InvalidOperationException();
        public AddinClient Client(InstanceInfo instance) => throw new InvalidOperationException();
    }

    static async Task<(bool IsError, string Text)> Call(IReadOnlyList<Tool> tools, string name, JsonObject args)
    {
        var r = (await tools.First(t => t.Name == name).InvokeAsync(args, CancellationToken.None)).ToJson();
        return ((bool)r["isError"]!, (string)r["content"]![0]!["text"]!);
    }

    static void Write(string path, string text)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, Vb1252.Encoding.GetBytes(text.Replace("\n", "\r\n")));
    }

    [Test]
    public static async Task ProjectsFileProjectsAndFileRead()
    {
        var root = Directory.CreateTempSubdirectory("vb6mcp-ws-").FullName;
        try
        {
            var shared = Path.Combine(root, "Shared", "modShared.bas");
            Write(shared, "Attribute VB_Name = \"modShared\"\nOption Explicit\n\nPublic Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n");
            Write(Path.Combine(root, "App1", "App1.vbp"), "Type=Exe\nModule=modShared; ..\\Shared\\modShared.bas\nName=\"App1\"\n");
            Write(Path.Combine(root, "Lib2", "Lib2.vbp"), "Type=OleDll\nModule=modShared; ..\\Shared\\modShared.bas\nName=\"Lib2\"\nCompatibleMode=\"2\"\n");
            var tools = new WorkspaceTools(new NoIde(), [root]).All();

            var p = await Call(tools, "ws_projects", new JsonObject { ["filter"] = "lib" });
            Assert.True(!p.IsError, p.Text);
            var pj = JsonNode.Parse(p.Text)!;
            Assert.Equal(2, (int)pj["total"]!);
            Assert.Equal(1, (int)pj["matched"]!);
            Assert.Equal("binary", (string)pj["projects"]![0]!["compatibility"]!);

            var fp = JsonNode.Parse((await Call(tools, "ws_file_projects", new JsonObject { ["path"] = shared })).Text)!;
            Assert.Equal(true, (bool)fp["shared"]!);
            Assert.Equal(2, fp["projects"]!.AsArray().Count);

            var fr = JsonNode.Parse((await Call(tools, "file_read", new JsonObject { ["path"] = shared, ["start_line"] = 3, ["count"] = 1 })).Text)!;
            Assert.Equal("disk", (string)fr["source"]!);
            Assert.Equal("modShared", (string)fr["component"]!);
            Assert.Equal(5, (int)fr["countOfLines"]!, "Attribute line not counted");
            Assert.Equal("Public Function Twice(ByVal x As Long) As Long", (string)fr["text"]!);
        }
        finally { Directory.Delete(root, true); }
    }

    [Test]
    public static async Task PathsOutsideRootsAndMissingRootsAreRefused()
    {
        var root = Directory.CreateTempSubdirectory("vb6mcp-ws-").FullName;
        try
        {
            var outside = await Call(new WorkspaceTools(new NoIde(), [root]).All(), "file_read", new JsonObject { ["path"] = @"C:\Windows\win.ini" });
            Assert.True(outside.IsError && outside.Text.Contains("outside the workspace roots"), outside.Text);
            var none = await Call(new WorkspaceTools(new NoIde(), []).All(), "ws_projects", new JsonObject());
            Assert.True(none.IsError && none.Text.Contains("VB6MCP_ROOTS"), none.Text);
        }
        finally { Directory.Delete(root, true); }
    }

    /// <summary>One IDE that has the given file loaded as Basic.modShared (clean), in design mode.</summary>
    sealed class OneIde(FakeAddin addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "Basic - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    [Test]
    public static async Task FileReplaceWritesClosedFilesAndRoutesOpenOnes()
    {
        var root = Directory.CreateTempSubdirectory("vb6mcp-ws-").FullName;
        try
        {
            var file = Path.Combine(root, "modShared.bas");
            Write(file, "Attribute VB_Name = \"modShared\"\nOption Explicit\n\nPublic Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n");
            var tools = new WorkspaceTools(new NoIde(), [root]).All();
            var read = JsonNode.Parse((await Call(tools, "file_read", new JsonObject { ["path"] = file })).Text)!;
            string hash = (string)read["hash"]!;

            var stale = await Call(tools, "file_replace_lines", new JsonObject { ["path"] = file, ["start_line"] = 4, ["count"] = 1, ["new_text"] = "x", ["expected_hash"] = "0000" });
            Assert.True(stale.IsError && stale.Text.Contains("hash"), stale.Text);

            var ok = await Call(tools, "file_replace_lines", new JsonObject { ["path"] = file, ["start_line"] = 4, ["count"] = 1, ["new_text"] = "    Twice = x + x", ["expected_hash"] = hash });
            Assert.True(!ok.IsError && ok.Text.Contains("\"source\":\"disk\""), ok.Text);
            Assert.Equal("Attribute VB_Name = \"modShared\"\r\nOption Explicit\r\n\r\nPublic Function Twice(ByVal x As Long) As Long\r\n    Twice = x + x\r\nEnd Function\r\n",
                         Vb1252.Encoding.GetString(File.ReadAllBytes(file)), "header kept, CR LF, final line break");

            // several edits in one call, on disk: all against the same text, written once
            hash = (string)JsonNode.Parse((await Call(tools, "file_read", new JsonObject { ["path"] = file })).Text)!["hash"]!;
            var many = await Call(tools, "file_replace_lines", new JsonObject
            {
                ["path"] = file, ["expected_hash"] = hash,
                ["edits"] = new JsonArray(
                    new JsonObject { ["start_line"] = 2, ["count"] = 0, ["new_text"] = "' shared helpers" },
                    new JsonObject { ["start_line"] = 4, ["count"] = 1, ["new_text"] = "    Twice = x * 2" }),
            });
            Assert.True(!many.IsError, many.Text);
            Assert.Equal("Attribute VB_Name = \"modShared\"\r\nOption Explicit\r\n' shared helpers\r\n\r\nPublic Function Twice(ByVal x As Long) As Long\r\n    Twice = x * 2\r\nEnd Function\r\n",
                         Vb1252.Encoding.GetString(File.ReadAllBytes(file)), "both edits, original numbering");
            // back to the first version for the routed test below
            hash = (string)JsonNode.Parse((await Call(tools, "file_read", new JsonObject { ["path"] = file })).Text)!["hash"]!;
            var back = await Call(tools, "file_replace_lines", new JsonObject
            {
                ["path"] = file, ["expected_hash"] = hash,
                ["edits"] = new JsonArray(
                    new JsonObject { ["start_line"] = 2, ["count"] = 1, ["new_text"] = "" },
                    new JsonObject { ["start_line"] = 5, ["count"] = 1, ["new_text"] = "    Twice = x + x" }),
            });
            Assert.True(!back.IsError, back.Text);

            // the same file open in an IDE: the edit goes there, the disk file is untouched
            var replaced = new List<string>();
            var ideText = "Option Explicit\r\n\r\nPublic Function Twice(ByVal x As Long) As Long\r\n    Twice = x + x\r\nEnd Function";
            using var fake = new FakeAddin((verb, a) => verb switch
            {
                "PROJECTS" => ("OK", $"1{Wire.US}Basic{Wire.US}0{Wire.US}{root}\\Basic.vbp{Wire.US}0"),
                "COMPONENTS" => ("OK", $"modShared{Wire.US}1{Wire.US}0{Wire.US}5{Wire.US}{file}"),
                "READ" => ("OK", $"5{Wire.US}{(replaced.Count > 0 ? 1 : 0)}{Wire.US}{(replaced.Count > 0 ? replaced[0] : ideText)}"),
                "REPLACE" => Replace(a),
                _ => ("ERR", $"400{Wire.US}unknown {verb}"),
            });
            (string, string) Replace(string[] a) { replaced.Add(ideText.Replace(a[4], a[5])); return ("OK", $"5{Wire.US}1"); }
            var before = File.ReadAllBytes(file);
            var routed = await Call(new WorkspaceTools(new OneIde(fake), [root]).All(), "file_replace_lines",
                new JsonObject { ["path"] = file, ["start_line"] = 4, ["count"] = 1, ["new_text"] = "    Twice = 2 * x", ["expected_hash"] = VbCode.Hash(ideText) });
            Assert.True(!routed.IsError && routed.Text.Contains("\"source\":\"ide\""), routed.Text);
            Assert.True(replaced.Count == 1 && replaced[0].Contains("Twice = 2 * x"), "edited in the IDE");
            Assert.True(before.AsSpan().SequenceEqual(File.ReadAllBytes(file)), "disk file untouched");
        }
        finally { Directory.Delete(root, true); }
    }
}