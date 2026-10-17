using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

public static class ReplaceCodeTests
{
    static VbSourceFile Parse(string name, string text) => VbSourceFile.Parse(name, Vb1252.Encoding.GetBytes(text));

    [Test]
    public static void HeaderAndOtherLinesStayByteExact()
    {
        var form = "VERSION 5.00\r\nBegin VB.Form F \r\n   Caption = \"x\"\r\nEnd\r\nAttribute VB_Name = \"F\"\r\nOption Explicit\r\nSub A()\r\n    x = 1\r\nEnd Sub\r\n";
        var f = Parse("F.frm", form);
        Assert.Equal("Option Explicit\r\nSub A()\r\n    x = 1\r\nEnd Sub", f.Code);
        var after = Vb1252.Encoding.GetString(f.ReplaceCode(3, 1, "    x = 2\r\n    y = 3"));
        Assert.Equal(form.Replace("    x = 1\r\n", "    x = 2\r\n    y = 3\r\n"), after);
        // insert at the end (count 0)
        var appended = Vb1252.Encoding.GetString(f.ReplaceCode(5, 0, "' end"));
        Assert.Equal(form + "' end\r\n", appended);
    }

    [Test]
    public static void InlineAttributesFollowTheirLine()
    {
        var cls = "Attribute VB_Name = \"C\"\r\nPublic Property Get Value() As Long\r\nAttribute Value.VB_UserMemId = 0\r\n    Value = 1\r\nEnd Property\r\n";
        var f = Parse("C.cls", cls);
        Assert.Equal("Public Property Get Value() As Long\r\n    Value = 1\r\nEnd Property", f.Code);
        // whole property replaced, signature kept: the attribute goes back after it
        var after = Vb1252.Encoding.GetString(f.ReplaceCode(1, 3, "Public Property Get Value() As Long\r\n    Value = 42\r\nEnd Property"));
        Assert.Equal(cls.Replace("Value = 1", "Value = 42"), after);
        // signature changed: refused, the attribute would be lost
        try { f.ReplaceCode(1, 3, "Public Property Get Other() As Long\r\nEnd Property"); throw new InvalidOperationException("should refuse"); }
        catch (ArgumentException ex) { Assert.True(ex.Message.Contains("VB_UserMemId"), ex.Message); }
        // body only: the attribute sits before the range and stays
        var body = Vb1252.Encoding.GetString(f.ReplaceCode(2, 1, "    Value = 7"));
        Assert.Equal(cls.Replace("Value = 1", "Value = 7"), body);
    }

    [Test]
    public static void RangesAreChecked()
    {
        var f = Parse("M.bas", "Attribute VB_Name = \"M\"\r\nOption Explicit\r\n");
        foreach (var (s, c) in new[] { (0, 1), (3, 0), (1, 2) })
            try { f.ReplaceCode(s, c, "x"); throw new InvalidOperationException($"{s},{c} should fail"); }
            catch (ArgumentException) { }
    }

    /// <summary>
    /// A WithEvents declaration's hidden VB_VarHelpID can be dropped on request (improvement notes 12): VB writes it again
    /// on save (spike 28b). Without the option the edit is refused, and the error names the option.
    /// </summary>
    [Test]
    public static void AWithEventsAttributeCanBeDroppedOnRequest()
    {
        var text = "Attribute VB_Name = \"clsT\"\r\nOption Explicit\r\nPrivate WithEvents m_comm As MSCommLib.MSComm\r\nAttribute m_comm.VB_VarHelpID = -1\r\n";
        var f = VbSourceFile.Parse("clsT.cls", Vb1252.Encoding.GetBytes(text));
        var kept = Vb1252.Encoding.GetString(f.ReplaceCode(2, 1, "' serial port\r\nPrivate WithEvents m_comm As MSCommLib.MSComm"));
        Assert.True(kept.Contains("MSComm\r\nAttribute m_comm.VB_VarHelpID = -1"), "an unchanged declaration keeps its attribute: " + kept);
        // the declaration changes (indented inside #If), so the attribute has no line to follow
        const string replacement = "#If NOTERM Then\r\n    Private m_comm As Object\r\n#Else\r\n    Private WithEvents m_comm As MSCommLib.MSComm\r\n#End If";
        try { f.ReplaceCode(2, 1, replacement); throw new InvalidOperationException("should be refused"); }
        catch (ArgumentException ex) { Assert.True(ex.Message.Contains("drop_var_help_ids"), ex.Message); }

        var after = Vb1252.Encoding.GetString(f.ReplaceCode(2, 1, replacement, dropVarHelpIds: true));
        Assert.True(!after.Contains("VB_VarHelpID") && after.Contains("#If NOTERM Then"), after);
    }
}
