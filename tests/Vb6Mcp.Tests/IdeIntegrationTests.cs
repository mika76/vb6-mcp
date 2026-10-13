using Vb6Mcp.Core.Bridge;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>
/// End-to-end checks against a real IDE: VB6 running (elevated) with the VB6McpBridge add-in and the
/// tests/fixtures/Basic project open and unmodified. Run: dotnet run --project tests/Vb6Mcp.Tests -- --cat IDE
/// </summary>
public static class IdeIntegrationTests
{
    static async Task<(AddinClient Client, InstanceInfo Instance)> Connect()
    {
        var live = Discovery.Scan().Live;
        Assert.True(live.Count > 0, "start VB6 with the VB6McpBridge add-in");
        foreach (var inst in live)
        {
            var c = new AddinClient(inst);
            if ((await c.ProjectsAsync()).Any(p => p.Name == "Basic")) return (c, inst);
        }
        throw new AssertFailed("open tests/fixtures/Basic/Basic.vbp in VB6");
    }

    [Test, Category("IDE")]
    public static async Task PingAndModeFromTitle()
    {
        var (c, inst) = await Connect();
        var ping = await c.PingAsync();
        Assert.Equal(inst.Pid, ping.Pid);
        Assert.Equal(1, ping.Protocol);
        Assert.Equal(IdeMode.Design, IdeWindows.Read(inst.Pid).Mode);
    }

    [Test, Category("IDE")]
    public static async Task ComponentsOfTheFixture()
    {
        var (c, _) = await Connect();
        var comps = (await c.ComponentsAsync("Basic")).ToDictionary(x => x.Name);
        Assert.Equal(5, comps["frmMain"].Type, "frmMain is a form");
        Assert.Equal(1, comps["modMath"].Type, "modMath is a standard module");
        Assert.Equal(2, comps["clsCounter"].Type, "clsCounter is a class");
        Assert.True(comps["modMath"].FileName.EndsWith("modMath.bas", StringComparison.OrdinalIgnoreCase), "file name");
    }

    [Test, Category("IDE")]
    public static async Task ReadReturnsTheModuleWithoutHiddenAttributes()
    {
        var (c, _) = await Connect();
        var code = await c.ReadAsync("Basic", "modMath");
        Assert.True(code.Text.StartsWith("Option Explicit"), "code starts after the hidden Attribute lines");
        Assert.True(!code.Text.Contains("Attribute VB_Name"), "Attribute lines are not part of the code module");
        Assert.Equal(code.CountOfLines, code.Text.Split("\r\n").Length);
        var part = await c.ReadAsync("Basic", "modMath", 7, 1);
        Assert.Equal("Public Function AddNumbers(ByVal a As Long, ByVal b As Long) As Long", part.Text);
    }

    [Test, Category("IDE")]
    public static async Task MembersListEveryPropertyKind()
    {
        var (c, _) = await Connect();
        var members = await c.MembersAsync("Basic", "clsCounter");
        var valueKinds = members.Where(m => m.Name == "Value").Select(m => m.ProcKind).OrderBy(k => k).ToList();
        // Get = 3 and Let = 1: Members itself lists the property once; the add-in expands the kinds
        Assert.SequenceEqual(new int?[] { 1, 3 }, valueKinds, "Value property kinds");
        Assert.True(members.Any(m => m.Name == "Changed" && m.MemberType == 4), "event listed");
        Assert.True(members.Any(m => m.Name == "m_value" && m.MemberType == 3 && m.Scope == 1), "private variable listed");
    }

    [Test, Category("IDE")]
    public static async Task ProcIncludesLeadingComment()
    {
        var (c, _) = await Connect();
        var p = await c.ProcAsync("Basic", "modMath", "AddNumbers", 0);
        Assert.True(p.Text.Contains("' Adds two numbers"), "comment above the function is part of the procedure");
        Assert.True(p.Text.TrimEnd().EndsWith("End Function"), "ends with End Function");
        Assert.True(p.ProcBody > p.ProcStart, "body line comes after start line");
    }

    [Test, Category("IDE")]
    public static async Task UnknownComponentAndProcedureAre404()
    {
        var (c, _) = await Connect();
        try { await c.ReadAsync("Basic", "NoSuchThing"); throw new AssertFailed("expected 404"); }
        catch (AddinErrorException ex) { Assert.Equal(404, ex.Code); }
        try { await c.ProcAsync("Basic", "modMath", "Missing", 0); throw new AssertFailed("expected 404"); }
        catch (AddinErrorException ex) { Assert.Equal(404, ex.Code); }
    }
}
