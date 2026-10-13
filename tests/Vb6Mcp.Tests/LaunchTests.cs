using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

public static class LaunchTests
{
    [Test]
    public static void ExeFromAssociationCommand()
    {
        Assert.Equal(@"C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe",
            Vb6Install.ExeFromCommand("\"C:\\Program Files (x86)\\Microsoft Visual Studio\\VB98\\vb6.exe\" \"%1\""));
        Assert.Equal(@"C:\VB98\VB6.EXE", Vb6Install.ExeFromCommand(@"C:\VB98\VB6.EXE %1"));
    }

    [Test]
    public static void FindExeHonoursTheOverride()
    {
        var fake = Path.Combine(Path.GetTempPath(), $"vb6mcp-fake-{Guid.NewGuid():N}.exe");
        File.WriteAllText(fake, "");
        var old = Environment.GetEnvironmentVariable("VB6MCP_VB6_EXE");
        try
        {
            Environment.SetEnvironmentVariable("VB6MCP_VB6_EXE", fake);
            Assert.Equal(Path.GetFullPath(fake), Vb6Install.FindExe());
        }
        finally { Environment.SetEnvironmentVariable("VB6MCP_VB6_EXE", old); File.Delete(fake); }
    }

    [Test]
    public static void ElevationOfOwnProcessIsReadable()
        => Assert.True(Vb6Install.IsElevated(Environment.ProcessId) is not null, "integrity level readable");

    static async Task<(bool IsError, string Text)> Launch(IIdeEnvironment env, JsonObject args)
    {
        var r = (await new IdeTools(env).All().First(t => t.Name == "ide_launch").InvokeAsync(args, CancellationToken.None)).ToJson();
        return ((bool)r["isError"]!, (string)r["content"]![0]!["text"]!);
    }

    sealed class Env(FakeAddin? addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new(addin is null ? [] : [addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "Basic - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    [Test]
    public static async Task ArgumentsAreCheckedBeforeStartingAnything()
    {
        var bad = await Launch(new Env(null), new JsonObject { ["project"] = @"C:\x\notes.txt" });
        Assert.True(bad.IsError && bad.Text.Contains(".vbp"), bad.Text);
        var missing = await Launch(new Env(null), new JsonObject { ["project"] = @"C:\does\not\exist.vbp" });
        Assert.True(missing.IsError && missing.Text.Contains("not found"), missing.Text);
        var mode = await Launch(new Env(null), new JsonObject { ["elevation"] = "maybe" });
        Assert.True(mode.IsError && mode.Text.Contains("auto, elevated or normal"), mode.Text);
    }

    [Test]
    public static async Task RefusesAProjectAlreadyOpenInAnotherIde()
    {
        var vbp = Path.Combine(Path.GetTempPath(), $"vb6mcp-launch-{Guid.NewGuid():N}.vbp");
        File.WriteAllText(vbp, "Type=Exe");
        try
        {
            using var fake = new FakeAddin((verb, _) => verb == "PROJECTS"
                ? ("OK", $"1{Wire.US}Basic{Wire.US}0{Wire.US}{vbp}{Wire.US}0") : ("ERR", $"400{Wire.US}x"));
            var r = await Launch(new Env(fake), new JsonObject { ["project"] = vbp });
            Assert.True(r.IsError && r.Text.Contains("already open"), r.Text);
        }
        finally { File.Delete(vbp); }
    }
}
