namespace Vb6Mcp.Tests;

/// <summary>Self-checks for the assertion helpers the runner relies on.</summary>
public static class RunnerTests
{
    [Test]
    public static void AssertEqualPassesAndFails()
    {
        Assert.Equal(3, 1 + 2);
        Assert.Throws<AssertFailed>(() => Assert.Equal(3, 4));
    }

    /// <summary>The runner keeps test runs out of the real audit log (fake approvals don't belong in a security log).</summary>
    [Test]
    public static void TestsDontWriteTheRealAuditLog()
    {
        Assert.True(!string.Equals(Path.GetFullPath(Vb6Mcp.Tools.Audit.LogPath), Path.GetFullPath(Vb6Mcp.Tools.Audit.DefaultPath), StringComparison.OrdinalIgnoreCase),
            "Audit.LogPath is the real log: " + Vb6Mcp.Tools.Audit.LogPath);
    }
}
