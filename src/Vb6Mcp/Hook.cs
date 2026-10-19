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
    /// <summary>
    /// Extensions of VB6 source and project files, plus the binary companions VB writes itself (.frx/.ctx/.pgx/.dsx/.dox:
    /// a form's, control's, property page's, designer's or document's pictures and lists; .res: the resource file; .vbw: the
    /// window layout). Editing a companion corrupts it; there is no tool for that on purpose (hard rule 2).
    /// </summary>
    public static readonly string[] Extensions = [".bas", ".cls", ".frm", ".ctl", ".pag", ".dsr", ".dob", ".vbp", ".vbg",
                                                  ".frx", ".ctx", ".pgx", ".dsx", ".dox", ".res", ".vbw"];

    /// <summary>Why a companion file is refused (it isn't source: the tools change it through VB).</summary>
    public const string CompanionReason =
        "This is a binary file VB6 writes itself (.frx/.ctx/.pgx/.dsx/.dox: a form's or control's pictures and lists; .res: resources; .vbw: window layout). " +
        "Don't edit or replace it: change the form through the vb6 tools (form_set_property with @picture:<file> for pictures, List/ItemData for lists), and VB rewrites it on save.";

    public const string Reason =
        "This is a VB6 source file (Windows-1252, hidden header and Attribute lines, possibly open in the VB6 IDE). " +
        "Don't use Edit/Write on it: use the vb6 MCP tools. Project open in the IDE: code_replace_lines, code_insert, proc_replace " +
        "(then project_save when the user agrees). Closed project: file_read + file_replace_lines (it edits the IDE instead if the file is open there).";

    /// <summary>The MCP tool name of debug_eval as Claude Code passes it to hooks (server registered as "vb6").</summary>
    public const string EvalTool = "mcp__vb6__debug_eval";

    /// <summary>The tools the hook is registered for: file edits, shell commands (ShellPolicy) and debug_eval (EvalPolicy).</summary>
    public const string Matcher = "Edit|Write|MultiEdit|NotebookEdit|Bash|PowerShell|" + EvalTool;

    /// <summary>
    /// Reads the hook input (JSON on stdin); returns 2 to block, 0 otherwise. Malformed input is allowed (a hook must not break
    /// other work).
    /// - Edit/Write on VB6 source files: blocked (there is always a right tool for it).
    /// - Bash/PowerShell commands ShellPolicy flags, and debug_eval with a side effect and allow_side_effect=true: the user is
    ///   asked (owner, 2026-10-17: ask rather than refuse), by writing permissionDecision "ask" to stdout. Without a person to
    ///   answer (headless runs) Claude Code doesn't run the tool. VB6MCP_SHELL_GUARD=off skips the shell checks.
    /// </summary>
    public static int Run(TextReader stdin, TextWriter stderr, Func<string, string?>? env = null, TextWriter? stdout = null)
    {
        env ??= Environment.GetEnvironmentVariable;
        stdout ??= Console.Out;
        string? path, tool, command, expression; bool allow;
        try
        {
            var input = JsonNode.Parse(stdin.ReadToEnd());
            tool = input?["tool_name"]?.GetValue<string>();
            path = input?["tool_input"]?["file_path"]?.GetValue<string>();
            command = input?["tool_input"]?["command"]?.GetValue<string>();
            expression = input?["tool_input"]?["expression"]?.GetValue<string>();
            allow = input?["tool_input"]?["allow_side_effect"] is JsonValue v && v.TryGetValue(out bool b) && b;
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException) { return 0; }
        if (tool is "Bash" or "PowerShell")
        {
            if (string.Equals(env("VB6MCP_SHELL_GUARD"), "off", StringComparison.OrdinalIgnoreCase) || command is null) return 0;
            if (ShellPolicy.Check(command) is not { } why) return 0;
            return Ask(stdout, "vb6-mcp security policy: " + why + " Allow it only if you asked for this.");
        }
        if (tool == EvalTool)
        {
            // without allow_side_effect the server refuses on its own and says how to ask; nothing to decide here
            if (!allow || expression is null || Tools.EvalPolicy.SideEffect(expression) is not { } effect) return 0;
            return Ask(stdout, $"vb6-mcp: the agent wants to run an expression with a side effect in the VB6 IDE (with its administrator rights): {effect}. Expression: ? {expression}");
        }
        if (path is null || !Blocks(path)) return 0;
        stderr.WriteLine(Path.GetExtension(path).ToLowerInvariant() is ".frx" or ".ctx" or ".pgx" or ".dsx" or ".dox" or ".res" or ".vbw" ? CompanionReason : Reason);
        return 2;
    }

    /// <summary>Claude Code's PreToolUse "ask" answer: the user sees the reason and approves or declines.</summary>
    private static int Ask(TextWriter stdout, string reason)
    {
        stdout.WriteLine(new JsonObject
        {
            ["hookSpecificOutput"] = new JsonObject { ["hookEventName"] = "PreToolUse", ["permissionDecision"] = "ask", ["permissionDecisionReason"] = reason },
        }.ToJsonString());
        return 0;
    }

    /// <summary>
    /// Whether the hook is installed in the user's settings and covers debug_eval: only then can an approved
    /// allow_side_effect be trusted, because the user was asked (an agent could pass the flag itself otherwise).
    /// </summary>
    public static bool GuardsEval(string? settingsPath = null)
    {
        try
        {
            var path = settingsPath ?? UserSettingsPath;
            if (!File.Exists(path) || JsonNode.Parse(File.ReadAllText(path))?["hooks"]?["PreToolUse"] is not JsonArray pre) return false;
            return pre.OfType<JsonObject>().Any(e => IsOurs(e) && (e["matcher"]?.GetValue<string>() ?? "").Split('|').Contains(EvalTool));
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException or InvalidOperationException) { return false; }
    }

    public static bool Blocks(string path) => Extensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    public static string UserSettingsPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".claude", "settings.json");

    /// <summary>
    /// Adds (or updates) the hook in a Claude Code settings file: hooks.PreToolUse gets one entry with the
    /// matcher <see cref="Matcher"/> running "&lt;exe&gt; hook". Other settings and other hooks are kept; a
    /// previous vb6mcp entry is replaced. A backup of the file is written first.
    /// </summary>
    public static int Install(string settingsPath, string exe, TextWriter output, bool setIdleTimeout = true)
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
            ["matcher"] = Matcher,
            ["hooks"] = new JsonArray(new JsonObject { ["type"] = "command", ["command"] = $"\"{exe}\" hook" }),
        });
        string? idleNote = setIdleTimeout ? SetIdleTimeout(root) : null;
        Directory.CreateDirectory(Path.GetDirectoryName(settingsPath)!);
        File.WriteAllText(settingsPath, root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }), new UTF8Encoding(false));
        output.WriteLine($"PreToolUse hook registered in {settingsPath}: \"{exe}\" hook");
        if (idleNote is not null) output.WriteLine(idleNote);
        return 0;
    }

    /// <summary>The Claude Code setting for how long a tool call may stay silent (spike 42).</summary>
    public const string IdleTimeoutVariable = "CLAUDE_CODE_MCP_TOOL_IDLE_TIMEOUT";

    /// <summary>The value the installers set: 2 minutes, in milliseconds.</summary>
    public const string IdleTimeoutValue = "120000";

    /// <summary>
    /// Spike 42: Claude Code aborts a tool call whose server sends nothing (no result, no progress) for this long; the
    /// default is 30 minutes, so a dead server once kept a call waiting half an hour. The server sends a heartbeat every
    /// 10 s during slow calls, so 2 minutes is plenty for vb6 and catches a dead or frozen server quickly. The setting is
    /// global (every MCP server of the user), so a value the user set is kept.
    /// </summary>
    private static string SetIdleTimeout(JsonObject root)
    {
        var env = root["env"] as JsonObject ?? new JsonObject();
        root["env"] = env;
        if (env[IdleTimeoutVariable] is { } existing)
            return $"{IdleTimeoutVariable} is already set ({existing.ToJsonString()} ms); kept";
        env[IdleTimeoutVariable] = IdleTimeoutValue;
        return $"{IdleTimeoutVariable}={IdleTimeoutValue} set in env: Claude Code now aborts a tool call of any MCP server that sends nothing for 2 minutes (default 30 minutes); vb6 sends a heartbeat during slow calls";
    }

    /// <summary>Removes our PreToolUse entry (any install folder) from a settings file; everything else is kept, a backup written.</summary>
    public static int Uninstall(string settingsPath, TextWriter output)
    {
        if (!File.Exists(settingsPath)) { output.WriteLine($"no {settingsPath}; nothing to remove"); return 0; }
        var root = JsonNode.Parse(File.ReadAllText(settingsPath))?.AsObject() ?? [];
        if (root["hooks"]?["PreToolUse"] is not JsonArray pre) { output.WriteLine("no vb6mcp hook registered"); return 0; }
        var ours = pre.OfType<JsonObject>().Where(IsOurs).ToList();
        if (ours.Count == 0) { output.WriteLine("no vb6mcp hook registered"); return 0; }
        File.Copy(settingsPath, settingsPath + ".vb6mcp-backup", overwrite: true);
        foreach (var e in ours) pre.Remove(e);
        if (pre.Count == 0) root["hooks"]!.AsObject().Remove("PreToolUse");
        if (root["hooks"] is JsonObject h && h.Count == 0) root.Remove("hooks");
        // the idle timeout the installer set (only our exact value: one the user chose stays)
        if (root["env"] is JsonObject env && (string?)env[IdleTimeoutVariable] == IdleTimeoutValue)
        {
            env.Remove(IdleTimeoutVariable);
            if (env.Count == 0) root.Remove("env");
            output.WriteLine($"{IdleTimeoutVariable} removed");
        }
        File.WriteAllText(settingsPath, root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }), new UTF8Encoding(false));
        output.WriteLine($"PreToolUse hook removed from {settingsPath}");
        return 0;
    }

    private static bool IsOurs(JsonObject entry)
        => entry["hooks"] is JsonArray hs && hs.OfType<JsonObject>().Any(h => h["command"]?.GetValue<string>() is { } c && c.Contains("vb6mcp", StringComparison.OrdinalIgnoreCase) && c.EndsWith(" hook", StringComparison.Ordinal));
}
