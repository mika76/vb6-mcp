using Vb6Mcp.Tools;

namespace Vb6Mcp.Tests;

/// <summary>code_read / file_read paging (improvement notes 13: a 185 KB class overflowed the client).</summary>
public static class PagingTests
{
    static string Lines(int n, int width = 10) => string.Join("\r\n", Enumerable.Range(1, n).Select(i => i.ToString().PadRight(width, 'x')));

    [Test]
    public static void ASmallModuleComesBackWhole()
    {
        var p = Paging.Page(Lines(10), null, null);
        Assert.Equal(10, p.Lines);
        Assert.True(!p.Truncated && p.NextStartLine is null, "not truncated");
    }

    [Test]
    public static void ALongModuleIsCutAtTheLineCapAndSaysWhereToReadOn()
    {
        var p = Paging.Page(Lines(2500), null, null);
        Assert.Equal(Paging.MaxLines, p.Lines);
        Assert.True(p.Truncated, "truncated");
        Assert.Equal(2001, p.NextStartLine ?? 0);
        var next = Paging.Page(Lines(2500), p.NextStartLine, null);
        Assert.Equal(500, next.Lines);
        Assert.True(!next.Truncated, "an explicit start without count reads to the end");
    }

    [Test]
    public static void WideLinesAreCutAtTheCharacterCap()
    {
        var p = Paging.Page(Lines(500, 1000), null, null);   // 500 KB
        Assert.True(p.Truncated && p.Lines < 100 && p.Text.Length <= Paging.MaxChars, $"{p.Lines} lines, {p.Text.Length} chars");
    }

    [Test]
    public static void AnExplicitRangeIsHonoured()
    {
        var p = Paging.Page(Lines(2500), 10, 5);
        Assert.Equal(5, p.Lines);
        Assert.True(p.Text.StartsWith("10xxxxxxxx") && !p.Truncated, p.Text);
    }
}
