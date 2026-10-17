using System.Text.Json.Nodes;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;

namespace Vb6Mcp.Tests;

/// <summary>
/// ide_launch's real start path, end to end: needs the elevated watchdog running (vb6mcp watchdog --stay, or one started by
/// an add-in), so VB6 starts without a UAC prompt. Launches a copy of the Basic fixture, checks it is started and owned,
/// then quits it. Run: dotnet run --project tests/Vb6Mcp.Tests -- --cat Launch
/// </summary>
public static class LaunchIntegrationTests
{
    static JsonObject Json(ToolResult r)
    {
        var j = r.ToJson();
        var text = (string)j["content"]![0]!["text"]!;
        Assert.True(!(bool)j["isError"]!, text);
        return JsonNode.Parse(text)!.AsObject();
    }

    [Test, Category("Launch")]
    public static async Task LaunchThroughTheWatchdogThenQuit()
    {
        Assert.True(WatchdogDiscovery.Find() is not null, "start the watchdog first (scripts\\build-addin.ps1 -WatchdogOnly -Stay)");
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "fixtures", "Basic"));
        var dir = Directory.CreateTempSubdirectory("vb6mcp-launch-").FullName;
        try
        {
            foreach (var f in new[] { "Basic.vbp", "frmMain.frm", "modMath.bas", "clsCounter.cls" }) File.Copy(Path.Combine(fixture, f), Path.Combine(dir, f));
            var tools = new IdeTools(new IdeEnvironment()).All().ToDictionary(t => t.Name);
            var started = Json(await tools["ide_launch"].InvokeAsync(new JsonObject { ["project"] = Path.Combine(dir, "Basic.vbp") }, CancellationToken.None));
            Assert.Equal("started", (string)started["status"]!);
            Assert.Equal("you", (string)started["owner"]!);
            Assert.True(((string)started["via"]!).StartsWith("watchdog"), "no UAC prompt: " + started["via"]);
            int pid = (int)started["pid"]!;
            // a copy outside source control gets "Add this project to SourceSafe?" a moment after it loads: answer it so the IDE
            // can close (ide_quit rightly reports it otherwise)
            for (var until = DateTime.UtcNow.AddSeconds(8); DateTime.UtcNow < until; await Task.Delay(500))
                if (UI.IdeWindows.Dialogs(pid).Any(d => d.Caption == "Source Code Control"))
                {
                    Json(await tools["ide_dialog_respond"].InvokeAsync(new JsonObject { ["pid"] = pid, ["button"] = "No" }, CancellationToken.None));
                    break;
                }
            var quit = Json(await tools["ide_quit"].InvokeAsync(new JsonObject { ["pid"] = pid }, CancellationToken.None));
            Assert.True((string?)quit["status"] == "quit", quit.ToJsonString());
            Assert.True(!Discovery.IsVb6Process(pid), "VB6 is gone");
        }
        // the elevated IDE writes files there (.vbw, source-control state) that a non-elevated process may not delete
        finally { try { Directory.Delete(dir, true); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { } }
    }
}
