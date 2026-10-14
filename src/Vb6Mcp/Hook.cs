using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Vb6Mcp;

/// <summary>
/// The Claude Code PreToolUse hook (PLAN.md §8): agents must not edit VB6 source files with the generic
/// Edit/Write tools. Those would write UTF-8, ignore the hidden header and Attribute lines, and bypass
/// a running IDE whose memory is the truth (hard rules 1, 2, 4). The hook blocks them (exit code 2,
/// reason on stderr, which Claude Code shows to the agent) and points at the vb6 MCP tools.
/// </summary>
public static class Hook
{
    /// <summary>Extensions of VB6 source and project files.</summary>
    public static readonly string[] Extensions = [".bas", ".cls", ".frm", ".ctl", ".pag", ".dsr", ".dob", ".vbp", ".vbg"];

    public const string Reason =
        "This is a VB6 source file (Windows-1252, hidden header and Attribute lines, possibly open in the VB6 IDE). " +
        "Don't use Edit/Write on it: use the vb6 MCP tools. Project open in the IDE: code_replace_lines, code_insert, proc_replace " +
        "(then project_save when the user agrees). Closed project: file_read + file_replace_lines (it edits the IDE instead if the file is open there).";

    /// <summary>Reads the hook input (JSON on stdin); returns 2 to block, 0 to allow. Malformed input is allowed (a hook must not break other work).</summary>
    public static int Run(TextReader stdin, TextWriter stderr)
    {
        string? path;
        try
        {
            var input = JsonNode.Parse(stdin.ReadToEnd());
            path = input?["tool_input"]?["file_path"]?.GetValue<string>();
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException) { return 0; }
        if (path is null || !Blocks(path)) return 0;
        stderr.WriteLine(Reason);
        return 2;
    }

    public static bool Blocks(string path) => Extensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    public static string UserSettingsPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".claude", "settings.json");

    /// <summary>
    /// Adds (or updates) the hook in a Claude Code settings file: hooks.PreToolUse gets one entry with the
    /// matcher "Edit|Write|MultiEdit" running "&lt;exe&gt; hook". Other settings and other hooks are kept; a
    /// previous vb6mcp entry is replaced. A backup of the file is written first.
    /// </summary>
    public static int Install(string settingsPath, string exe, TextWriter output)
    {
        JsonObject root;
        if (File.Exists(settingsPath))
        {
            File.Copy(settingsPath, settingsPath + ".vb6mcp-backup", overwrite: true);
            root = JsonNode.Parse(File.ReadAllText(settingsPath))?.AsObject() ?? [];
        }
        else root = [];
        var hooks = root["hooks"] as JsonObject ?? new JsonObject();
        root["hooks"] = hooks;
        var pre = hooks["PreToolUse"] as JsonArray ?? new JsonArray();
        hooks["PreToolUse"] = pre;
        // drop earlier vb6mcp entries (any install folder)
        foreach (var entry in pre.OfType<JsonObject>().Where(IsOurs).ToList()) pre.Remove(entry);
        pre.Add(new JsonObject
        {
            ["matcher"] = "Edit|Write|MultiEdit",
            ["hooks"] = new JsonArray(new JsonObject { ["type"] = "command", ["command"] = $"\"{exe}\" hook" }),
        });
        Directory.CreateDirectory(Path.GetDirectoryName(settingsPath)!);
        File.WriteAllText(settingsPath, root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }), new UTF8Encoding(false));
        output.WriteLine($"PreToolUse hook registered in {settingsPath}: \"{exe}\" hook");
        return 0;
    }

    private static bool IsOurs(JsonObject entry)
        => entry["hooks"] is JsonArray hs && hs.OfType<JsonObject>().Any(h => h["command"]?.GetValue<string>() is { } c && c.Contains("vb6mcp", StringComparison.OrdinalIgnoreCase) && c.EndsWith(" hook", StringComparison.Ordinal));
}
