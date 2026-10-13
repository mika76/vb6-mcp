using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

public static class VbSourceFileTests
{
    static string Fixture(string name)
    {
        // tests run from bin\...; walk up to the repo root
        var dir = AppContext.BaseDirectory;
        while (!File.Exists(Path.Combine(dir, "vb6-mcp.slnx"))) dir = Path.GetDirectoryName(dir)!;
        return Path.Combine(dir, "tests", "fixtures", "Basic", name);
    }

    [Test]
    public static void Every1252ByteRoundTrips()
    {
        // Hard rule 1: byte-exact, including the bytes 1252 leaves undefined (0x81, 0x8D, 0x8F, 0x90, 0x9D)
        var all = Enumerable.Range(0, 256).Select(b => (byte)b).ToArray();
        var back = Vb1252.Encoding.GetBytes(Vb1252.Encoding.GetString(all));
        Assert.SequenceEqual(all, back, "all 256 byte values");
    }

    [Test]
    public static void ModuleSplitsAttributeHeaderFromCode()
    {
        var f = VbSourceFile.Load(Fixture("modMath.bas"));
        Assert.Equal(VbFileKind.Module, f.Kind);
        Assert.Equal("modMath", f.VbName);
        Assert.Equal(1, f.Header.Count, "only the Attribute VB_Name line");
        Assert.True(f.Code.StartsWith("Option Explicit"), "code starts like in the IDE");
        Assert.Equal(13, f.CountOfLines, "same count as the IDE reported (IDE test: 13)");
    }

    [Test]
    public static void FormSkipsDesignerBlockAndAttributes()
    {
        var f = VbSourceFile.Load(Fixture("frmMain.frm"));
        Assert.Equal(VbFileKind.Form, f.Kind);
        Assert.Equal("frmMain", f.VbName);
        Assert.True(f.Header[0] == "VERSION 5.00", "header starts with VERSION");
        Assert.True(f.Header.Any(l => l.Contains("Begin VB.CommandButton cmdAdd")), "designer block is header");
        Assert.True(f.Code.StartsWith("Option Explicit"), "code starts after the designer block and attributes");
        Assert.True(!f.Code.Contains("Attribute"), "no Attribute lines in code");
    }

    [Test]
    public static void ClassSkipsBeginEndBlock()
    {
        var f = VbSourceFile.Load(Fixture("clsCounter.cls"));
        Assert.Equal(VbFileKind.Class, f.Kind);
        Assert.Equal("clsCounter", f.VbName);
        Assert.True(f.Code.StartsWith("Option Explicit"), "code after BEGIN..END and attributes");
        Assert.Equal(18, f.CountOfLines, "same count as the IDE reported (18)");
    }

    [Test]
    public static void InlineAttributesInsideProceduresAreSeparated()
    {
        var src = "Attribute VB_Name = \"C\"\r\nOption Explicit\r\nPublic Property Get Value() As Long\r\nAttribute Value.VB_UserMemId = 0\r\n    Value = 1\r\nEnd Property\r\n";
        var f = VbSourceFile.Parse("C.cls", Vb1252.Encoding.GetBytes(src));
        Assert.Equal("Option Explicit\r\nPublic Property Get Value() As Long\r\n    Value = 1\r\nEnd Property", f.Code);
        Assert.Equal(1, f.InlineAttributes.Count);
        Assert.Equal(2, f.InlineAttributes[0].AfterCodeLine, "follows code line 2 (the Property Get line)");
    }
}
