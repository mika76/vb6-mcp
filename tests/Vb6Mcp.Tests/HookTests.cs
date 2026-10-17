using System.Text.Json.Nodes;

namespace Vb6Mcp.Tests;

public static class HookTests
{
    static (int Code, string Err) Run(string json)
    {
        var err = new StringWriter();
        return (Hook.Run(new StringReader(json), err), err.ToString());
    }

    [Test]
    public static void BlocksVb6FilesOnly()
    {
        var bas = Run("""{"tool_name":"Edit","tool_input":{"file_path":"D:\\p\\modMath.BAS","old_string":"a","new_string":"b"}}""");
        Assert.True(bas.Code == 2 && bas.Err.Contains("file_replace_lines"), bas.Err);
        Assert.Equal(2, Run("""{"tool_name":"Write","tool_input":{"file_path":"D:\\p\\App.vbp","content":"x"}}""").Code, ".vbp too");
        Assert.Equal(0, Run("""{"tool_name":"Edit","tool_input":{"file_path":"D:\\p\\Program.cs"}}""").Code, "other files allowed");
        var frx = Run("""{"tool_name":"Write","tool_input":{"file_path":"D:\\p\\frmMain.frx","content":"x"}}""");
        Assert.True(frx.Code == 2 && frx.Err.Contains("binary file VB6 writes itself"), "companions get their own reason: " + frx.Err);
        Assert.Equal(2, Run("""{"tool_name":"Edit","tool_input":{"file_path":"D:\\p\\App.vbw"}}""").Code, ".vbw too");
        Assert.Equal(0, Run("not json").Code, "malformed input never blocks");
        Assert.Equal(0, Run("""{"tool_name":"Bash","tool_input":{"command":"ls"}}""").Code, "no file_path");
    }

    [Test]
    public static void InstallMergesAndReplacesOurEntry()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-hook-").FullName;
        try
        {
            var settings = Path.Combine(dir, "settings.json");
            File.WriteAllText(settings, """{"model":"x","hooks":{"PreToolUse":[{"matcher":"Bash","hooks":[{"type":"command","command":"other.exe"}]}]}}""");
            Hook.Install(settings, @"C:\a\vb6mcp.exe", TextWriter.Null);
            Hook.Install(settings, @"C:\b\vb6mcp.exe", TextWriter.Null);      // a later install replaces ours
            var root = JsonNode.Parse(File.ReadAllText(settings))!;
            Assert.Equal("x", (string)root["model"]!, "other settings kept");
            var pre = root["hooks"]!["PreToolUse"]!.AsArray();
            Assert.Equal(2, pre.Count, "the other hook and one of ours");
            Assert.Equal("other.exe", (string)pre[0]!["hooks"]![0]!["command"]!);
            Assert.Equal("\"C:\\b\\vb6mcp.exe\" hook", (string)pre[1]!["hooks"]![0]!["command"]!);
            Assert.True(File.Exists(settings + ".vb6mcp-backup"), "backup written");

            Hook.Uninstall(settings, TextWriter.Null);
            var after = JsonNode.Parse(File.ReadAllText(settings))!;
            var left = after["hooks"]!["PreToolUse"]!.AsArray();
            Assert.Equal(1, left.Count, "only the other hook stays");
            Assert.Equal("other.exe", (string)left[0]!["hooks"]![0]!["command"]!);
            Assert.Equal("x", (string)after["model"]!, "other settings kept");
            Hook.Uninstall(settings, TextWriter.Null);                              // nothing left of ours: no change, no error
        }
        finally { Directory.Delete(dir, true); }
    }

    [Test]
    public static void UninstallRemovesEmptyHookSections()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-hook-").FullName;
        try
        {
            var settings = Path.Combine(dir, "settings.json");
            File.WriteAllText(settings, """{"model":"x"}""");
            Hook.Install(settings, @"C:\a\vb6mcp.exe", TextWriter.Null);
            Hook.Uninstall(settings, TextWriter.Null);
            var root = JsonNode.Parse(File.ReadAllText(settings))!.AsObject();
            Assert.True(!root.ContainsKey("hooks") && (string)root["model"]! == "x", root.ToJsonString());
        }
        finally { Directory.Delete(dir, true); }
    }
}
