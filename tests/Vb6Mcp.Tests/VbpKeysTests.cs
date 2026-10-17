using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

/// <summary>Setting .vbp keys byte-exactly (Phase 7 A3): CondComp, Command32 and friends; members and references refused.</summary>
public static class VbpKeysTests
{
    const string Vbp = "Type=Exe\r\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#stdole2.tlb#OLE Automation\r\nForm=frmMain.frm\r\n" +
                       "Command32=\"/nu /config\"\r\nName=\"App\"\r\nCondComp=\"nologin = 0\"\r\nMajorVer=1\r\n\r\n[MS Transaction Server]\r\nAutoRefresh=1\r\n";

    [Test]
    public static void AnExistingKeyChangesAndNothingElse()
    {
        var after = VbpKeys.Set(Vbp, "CondComp", "nologin = 1");
        Assert.Equal(Vbp.Replace("CondComp=\"nologin = 0\"", "CondComp=\"nologin = 1\""), after);
        Assert.Equal("/nu", VbpKeys.Get(VbpKeys.Set(Vbp, "command32", "/nu"), "Command32"));
    }

    [Test]
    public static void ANewKeyGoesIntoTheProjectSection()
    {
        var after = VbpKeys.Set(Vbp, "HelpContextID", "12");
        Assert.True(after.Contains("MajorVer=1\r\nHelpContextID=12\r\n\r\n[MS Transaction Server]"), after);
        Assert.True(VbpKeys.Set(Vbp, "Title", "App").Contains("Title=\"App\""), "string keys are quoted");
    }

    [Test]
    public static void MembersReferencesAndCompatibilityAreRefused()
    {
        foreach (var key in new[] { "Reference", "Object", "Form", "CompatibleMode", "Type" })
        {
            try { VbpKeys.Set(Vbp, key, "x"); throw new InvalidOperationException(key + " should be refused"); }
            catch (ArgumentException) { }
        }
        try { VbpKeys.Set(Vbp, "Title", "a\"b"); throw new InvalidOperationException("quote should be refused"); }
        catch (ArgumentException) { }
    }
}
