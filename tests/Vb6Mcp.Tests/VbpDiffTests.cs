using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

/// <summary>project_save's .vbp report (B8): what VB rewrote, and control/reference version bumps (a production project, 2026-10-17).</summary>
public static class VbpDiffTests
{
    const string Before =
        "Type=Exe\r\n" +
        "Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\\..\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation\r\n" +
        "Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0; mscomctl.OCX\r\n" +
        "Object={6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0; comctl32.Ocx\r\n" +
        "Form=frmMain.frm\r\n";

    [Test]
    public static void VersionBumpsAndRewrittenLinesAreReported()
    {
        string after = Before
            .Replace("#..\\..\\Windows\\SysWOW64\\stdole2.tlb#", "#C:\\Windows\\SysWOW64\\stdole2.tlb#")
            .Replace("#2.1#0; mscomctl.OCX", "#2.2#0; MSCOMCTL.OCX")
            .Replace("#1.4#0; comctl32.Ocx", "#1.5#0; COMCTL32.OCX");
        var c = VbpDiff.Compare(Before, after);
        Assert.Equal(3, c.Removed.Count);
        Assert.Equal(3, c.Added.Count);
        Assert.Equal("MSCOMCTL.OCX 2.1>2.2|COMCTL32.OCX 1.4>1.5",
                     string.Join("|", c.Versions.Select(v => $"{v.Name} {v.OldVersion}>{v.NewVersion}")));
        Assert.True(c.Versions.All(v => v.Kind == "control"), "both are controls");
    }

    [Test]
    public static void AnUnchangedVbpIsNoChange()
        => Assert.True(!VbpDiff.Compare(Before, Before).Any, "no change");

    [Test]
    public static void AReorderedFileIsNoChange()
    {
        var lines = Before.Split("\r\n", StringSplitOptions.RemoveEmptyEntries).Reverse();
        Assert.True(!VbpDiff.Compare(Before, string.Join("\r\n", lines) + "\r\n").Any, "same lines, other order");
    }
}
