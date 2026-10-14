using Vb6Mcp.Core;

namespace Vb6Mcp.Tests;

public static class LineDiffTests
{
    static string Apply(IReadOnlyList<string> a, IReadOnlyList<LineDiff.Edit> edits)
        => string.Join("\n", edits.Where(e => e.Op != '-').Select(e => e.Text));

    [Test]
    public static void EditScriptsRebuildTheTarget()
    {
        var cases = new (string A, string B)[]
        {
            ("", ""), ("a", ""), ("", "a"), ("a\nb\nc", "a\nb\nc"), ("a\nb\nc", "a\nx\nc"),
            ("a\nb\nc\nd", "b\nc\nd\ne"), ("x\ny", "a\nb\nc"), ("Option Explicit\n\nSub A()\nEnd Sub", "Option Explicit\n' new\nSub A()\n  x = 1\nEnd Sub"),
        };
        foreach (var (a, b) in cases)
        {
            var la = LineDiff.Lines(a); var lb = LineDiff.Lines(b);
            var edits = LineDiff.Diff(la, lb);
            Assert.Equal(string.Join("\n", lb), Apply(la, edits), $"'{a}' -> '{b}'");
            Assert.Equal(string.Join("\n", la), string.Join("\n", edits.Where(e => e.Op != '+').Select(e => e.Text)), "source side intact");
        }
    }

    [Test]
    public static void UnifiedOutput()
    {
        Assert.Equal("", LineDiff.Unified("a\r\nb", "a\r\nb", "ide", "disk"));
        var d = LineDiff.Unified("Option Explicit\r\nSub A()\r\n    x = 1\r\nEnd Sub", "Option Explicit\r\nSub A()\r\n    x = 2\r\nEnd Sub", "ide", "disk");
        Assert.Equal("--- ide\n+++ disk\n@@ -1,4 +1,4 @@\n Option Explicit\n Sub A()\n-    x = 1\n+    x = 2\n End Sub\n", d.Replace("\r\n", "\n"));
    }
}
