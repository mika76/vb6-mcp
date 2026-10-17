using System.Text.RegularExpressions;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>Documentation that must follow the code (AGENTS.md: stale docs count as a bug).</summary>
public static class DocsTests
{
    sealed class NoIde : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => throw new InvalidOperationException();
        public AddinClient Client(InstanceInfo instance) => throw new InvalidOperationException();
    }

    /// <summary>The tool index in docs/tools.md lists exactly the server's tools (regenerate with scripts/tool-index.js).</summary>
    [Test]
    public static void TheToolIndexListsEveryTool()
    {
        var docs = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "docs", "tools.md");
        var text = File.ReadAllText(docs);
        int start = text.IndexOf("## Tool index", StringComparison.Ordinal), end = text.IndexOf("## Common arguments", StringComparison.Ordinal);
        Assert.True(start >= 0 && end > start, "docs/tools.md has a Tool index before Common arguments");
        var listed = Regex.Matches(text[start..end], @"^\| [^|]+ \| `([a-z_]+)` \|", RegexOptions.Multiline).Select(m => m.Groups[1].Value).ToHashSet();
        var tools = new IdeTools(new NoIde()).All().Concat(new WorkspaceTools(new NoIde(), []).All()).Select(t => t.Name)
            .Append("batch_run").Append("server_restart").ToHashSet();   // both are added in Program.cs
        var missing = tools.Except(listed).Order().ToList();
        var extra = listed.Except(tools).Order().ToList();
        Assert.True(missing.Count == 0 && extra.Count == 0,
            $"docs/tools.md index is stale (run node scripts/tool-index.js <vb6mcp.exe> <out> and paste it): missing {string.Join(", ", missing)}; extra {string.Join(", ", extra)}");
    }
}
