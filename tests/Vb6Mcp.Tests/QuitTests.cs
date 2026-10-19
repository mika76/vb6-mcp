using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;
using Vb6Mcp.Watchdog;

namespace Vb6Mcp.Tests;

/// <summary>ide_quit (PLAN.md C7): only an IDE this session owns (or force), never with unsaved work unless discard, WM_CLOSE via the watchdog.</summary>
public static class QuitTests
{
    const char US = Wire.US;

    sealed class Actions : IWatchdogActions
    {
        public List<string> Calls = [];
        public bool Elevated => true;
        public int ParentPid => 1;
        public bool IsVb6(int pid) => true;
        public int WindowPid(long hwnd) => 1;
        public Task<string> BreakAsync(int pid, CancellationToken ct) => throw new NotSupportedException();
        public string Press(long h) => throw new NotSupportedException();
        public string Kill(int pid) => throw new NotSupportedException();
        public string Quit(int pid) { Calls.Add($"quit {pid}"); return "close-posted"; }
        public string Make(string vbp, string defines) => throw new NotSupportedException();
        public Task<string> SnapAsync(long h, string id, CancellationToken ct) => throw new NotSupportedException();
        public string Launch(string? project) => throw new NotSupportedException();
        public string Regsvr(string file, bool unregister) => throw new NotSupportedException();
        public string RegStatus(int pid) => throw new NotSupportedException();
    }

    sealed class Env(FakeAddin addin, InstanceInfo watchdog) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "P - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
        public InstanceInfo? FindWatchdog() => watchdog;
    }

    static async Task<JsonObject> Call(IdeTools tools, JsonObject args)
    {
        var r = (await tools.All().First(t => t.Name == "ide_quit").InvokeAsync(args, CancellationToken.None)).ToJson();
        var text = (string)r["content"]![0]!["text"]!;
        return (bool)r["isError"]! ? new JsonObject { ["error"] = text } : JsonNode.Parse(text)!.AsObject();
    }

    [Test]
    public static async Task QuitNeedsOwnershipOrForceAndNoUnsavedWork()
    {
        var actions = new Actions();
        var discovery = Path.Combine(Path.GetTempPath(), $"vb6mcp-wd-{Guid.NewGuid():N}.json");
        var owners = Directory.CreateTempSubdirectory("vb6mcp-owners-").FullName;
        using var stop = new CancellationTokenSource();
        _ = new WatchdogHost(actions, _ => { }, discovery).RunAsync(stop.Token);
        for (int i = 0; i < 50 && !File.Exists(discovery); i++) await Task.Delay(20);
        var wd = Discovery.TryRead(discovery)!;
        try
        {
            bool dirty = true;
            using var fake = new FakeAddin((verb, _) => verb switch
            {
                "PROJECTS" => ("OK", $"1{US}P{US}0{US}C:/x/P.vbp{US}0"),
                "COMPONENTS" => ("OK", $"modA{US}1{US}{(dirty ? 1 : 0)}{US}1{US}C:/x/modA.bas"),
                _ => ("ERR", $"400{US}unexpected {verb}"),
            });
            bool alive = true;
            var tools = new IdeTools(new Env(fake, wd)) { OwnersFolder = owners, ProcessAlive = _ => alive && actions.Calls.Count == 0 };

            var notMine = await Call(tools, new JsonObject());
            Assert.True(((string?)notMine["error"] ?? "").Contains("force=true"), notMine.ToJsonString());

            var unsaved = await Call(tools, new JsonObject { ["force"] = true });
            Assert.True(((string?)unsaved["error"] ?? "").Contains("P.modA"), unsaved.ToJsonString());
            Assert.Equal(0, actions.Calls.Count);

            dirty = false;
            var quit = await Call(tools, new JsonObject { ["force"] = true });
            Assert.Equal("quit", (string?)quit["status"] ?? quit.ToJsonString());
            Assert.Equal($"quit {fake.Instance.Pid}", string.Join("|", actions.Calls));
        }
        finally { stop.Cancel(); Directory.Delete(owners, true); }
    }
}
