using Vb6Mcp.Core.Bridge;

namespace Vb6Mcp.Tests;

public static class DiscoveryTests
{
    [Test]
    public static void SplitsLiveStaleAndUnreadable()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-disc-").FullName;
        try
        {
            File.WriteAllText(Path.Combine(dir, "100.json"), """{"pid":100,"port":5000,"token":"aa","protocol":1,"addin":"0.1.0","started":"2026-10-13T14:00:00"}""");
            File.WriteAllText(Path.Combine(dir, "200.json"), """{"pid":200,"port":5001,"token":"bb","protocol":1}""");
            File.WriteAllText(Path.Combine(dir, "300.json"), """{"pid":300,"port":""");   // half-written
            var r = Discovery.Scan(dir, pid => pid == 100);

            Assert.Equal(1, r.Live.Count, "live");
            Assert.Equal(5000, r.Live[0].Port);
            Assert.Equal("0.1.0", r.Live[0].AddinVersion);
            Assert.Equal(1, r.Stale.Count, "stale");
            Assert.Equal(200, r.Stale[0].Pid);
            Assert.Equal(1, r.Unreadable.Count, "unreadable");

            Assert.Equal(1, Discovery.DeleteStale(r));
            Assert.True(!File.Exists(Path.Combine(dir, "200.json")), "stale file deleted");
            Assert.True(File.Exists(Path.Combine(dir, "100.json")), "live file kept");
        }
        finally { Directory.Delete(dir, true); }
    }

    [Test]
    public static void MissingFolderMeansNoInstances()
    {
        var r = Discovery.Scan(Path.Combine(Path.GetTempPath(), "vb6mcp-does-not-exist-" + Guid.NewGuid()));
        Assert.Equal(0, r.Live.Count + r.Stale.Count + r.Unreadable.Count);
    }

    [Test]
    public static void OwnProcessIsNotVb6()
        => Assert.True(!Discovery.IsVb6Process(Environment.ProcessId), "the test runner isn't VB6");
}
