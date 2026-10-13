using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
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
}
