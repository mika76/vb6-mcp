using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>Compile/run tools wait for reloads the watchdog has pending (Phase 7 B6, improvement notes 27).</summary>
public static class ReloadSettleTests
{
    sealed class Env : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "P [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance);
    }

    [Test]
    public static async Task APendingReloadIsWaitedForAndReported()
    {
        var path = Path.Combine(Path.GetTempPath(), $"vb6mcp-sync-{Guid.NewGuid():N}.json");
        try
        {
            var t0 = DateTime.Now;
            SyncState.Write(new SyncSnapshot([], [
                new SyncEvent(t0, 77, "pending", @"D:\p\modA.bas", "P", "modA", "will reload in design mode"),
                new SyncEvent(t0, 77, "pending", @"D:\p\modB.bas", "P", "modB", null),
                new SyncEvent(t0.AddSeconds(1), 77, "reloaded", @"D:\p\modB.bas", "P", "modB", null),
                new SyncEvent(t0, 88, "pending", @"D:\q\other.bas", "Q", "other", null),   // another IDE
            ], t0), path);
            var tools = new IdeTools(new Env()) { SyncStatePath = path, ReloadSettle = TimeSpan.FromMilliseconds(600) };
            var r = await tools.SettleReloadsAsync(77, CancellationToken.None);
            Assert.True(r is not null, "pending reload found");
            Assert.Equal(@"D:\p\modA.bas", (string)r!["stillPending"]![0]!);
            Assert.Equal(1, r["stillPending"]!.AsArray().Count);
            Assert.True(await new IdeTools(new Env()) { SyncStatePath = path }.SettleReloadsAsync(99, CancellationToken.None) is null, "nothing pending for another pid");
        }
        finally { File.Delete(path); }
    }

    [Test]
    public static void WithAddsAFieldToAJsonResult()
    {
        var r = ToolResult.Ok(new { status = "ok" }).With("reloads", new JsonObject { ["applied"] = 1 }).ToJson();
        Assert.Equal("{\"status\":\"ok\",\"reloads\":{\"applied\":1}}", (string)r["content"]![0]!["text"]!);
    }
}
