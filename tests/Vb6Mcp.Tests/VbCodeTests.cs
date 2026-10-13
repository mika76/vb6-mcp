using Vb6Mcp.Core;

namespace Vb6Mcp.Tests;

public static class VbCodeTests
{
    [Test]
    public static void HashIsStableShortAndContentSensitive()
    {
        var h = VbCode.Hash("Option Explicit");
        Assert.Equal(16, h.Length);
        Assert.Equal(h, VbCode.Hash("Option Explicit"));
        Assert.True(h != VbCode.Hash("Option Explicit "), "one extra space changes the hash");
    }

    [Test]
    public static void PrepareNormalizesLineEndingsAndDropsOneTrailingBreak()
    {
        Assert.Equal("a\r\nb\r\nc", VbCode.PrepareForInsert("a\nb\r\nc\n"));
        Assert.Equal("a\r\n", VbCode.PrepareForInsert("a\r\n\r\n"), "only one trailing break is removed");
        Assert.Equal("", VbCode.PrepareForInsert(""));
        Assert.Equal("MsgBox \"café €\"", VbCode.PrepareForInsert("MsgBox \"café €\""));   // é and € exist in 1252
    }

    [Test]
    public static void PrepareEnforcesVb6LimitsAndEncoding()
    {
        Assert.Throws<ArgumentException>(() => VbCode.PrepareForInsert(new string('x', 1024)));
        VbCode.PrepareForInsert(new string('x', 1023));
        var ok = string.Join("\n", Enumerable.Repeat("x = 1 + _", 24)) + "\n1";
        VbCode.PrepareForInsert(ok);
        var tooMany = string.Join("\n", Enumerable.Repeat("x = 1 + _", 25)) + "\n1";
        Assert.Throws<ArgumentException>(() => VbCode.PrepareForInsert(tooMany));
        var ex = Assert.Throws<ArgumentException>(() => VbCode.PrepareForInsert("MsgBox \"日本\""));
        Assert.True(ex.Message.Contains("1252"), "explains the encoding");
        Assert.Throws<ArgumentException>(() => VbCode.PrepareForInsert("a\x1Fb"));
    }

    [Test]
    public static void SliceAndTrim()
    {
        var text = "l1\r\nl2\r\nl3\r\nl4";
        Assert.Equal("l2\r\nl3", VbCode.Slice(text, 2, 2));
        Assert.Equal("l4", VbCode.Slice(text, 4, 10));
        Assert.Equal("", VbCode.Slice(text, 9, 1));
        var (t, start, count) = VbCode.TrimLeadingBlankLines("\r\n  \r\n' doc\r\nSub X()\r\nEnd Sub", 10, 5);
        Assert.Equal("' doc\r\nSub X()\r\nEnd Sub", t);
        Assert.Equal(12, start);
        Assert.Equal(3, count);
    }
}
