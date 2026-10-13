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
}
