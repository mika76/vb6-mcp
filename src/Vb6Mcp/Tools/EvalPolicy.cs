using System.Text.RegularExpressions;

namespace Vb6Mcp.Tools;

/// <summary>
/// What debug_eval refuses (PLAN.md E5). The Immediate window runs with the IDE's rights, which are usually an
/// administrator's (VB6 runs elevated here), so an expression is code execution, not just a look at a variable. Both T12
/// batch runs changed rows of the owner's test database from it (`? SqlCon.Execute("update orders set ...")`); the owner
/// asked for side effects of every kind to be refused, not only database writes ("security should be taken seriously").
/// Refused by default; VB6MCP_EVAL=open allows them, flagged in the result and the audit log.
/// It is a guard rail against the obvious, not a sandbox: a procedure of the project that does these things (UpdateCard 184)
/// passes, since only the expression's text is checked.
/// </summary>
public static class EvalPolicy
{
    /// <summary>SQL that changes data or schema, anywhere in a text (it is usually inside a string literal).</summary>
    public static readonly Regex Sql = new(
        @"\b(insert\s+into|update\s+\[?\w+\]?(\.\[?\w+\]?)*\s+set|delete\s+(from\s+)?\[?\w|drop\s+(table|view|procedure|proc|index|database|function|trigger|user|login)|alter\s+(table|view|procedure|proc|database|function|trigger|user|login)|truncate\s+table|create\s+(table|view|procedure|proc|index|database|function|trigger|user|login)|merge\s+into|exec(ute)?\s+\w|grant\s|revoke\s|deny\s|sp_configure|xp_cmdshell|bulk\s+insert|restore\s+database|backup\s+database|shutdown\b)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    static readonly (string Category, Regex Pattern)[] Rules =
    [
        ("database write", Sql),
        ("database write", R(@"\.(AddNew|Update|UpdateBatch|Delete|CommitTrans|Edit|BeginTrans)\b(?!\s*=)")),
        ("process start", R(@"\b(Shell|ShellExecute(Ex)?[AW]?|CreateProcess[AW]?|WinExec)\s*\(|\bShell\s+""|\.(Run|Exec|ShellExecute)\s*[\(""]|""(WScript\.Shell|Shell\.Application|WScript\.Network)""|winmgmts|Win32_Process")),
        ("file change", R(@"\b(Kill|FileCopy|MkDir|RmDir|SetAttr|SavePicture|SaveSetting|DeleteSetting)\b|\bName\s+.+\s+As\s+|\bOpen\s+.+\s+For\s+(Output|Append|Binary|Random)\b|\.(DeleteFile|DeleteFolder|CopyFile|CopyFolder|MoveFile|MoveFolder|CreateTextFile|CreateFolder|SaveToFile|WriteLine|WriteBlankLines)\b|\b(DeleteFile|MoveFile(Ex)?|CopyFile(Ex)?|CreateFile|WriteFile|RemoveDirectory|CreateDirectory|SHFileOperation)[AW]?\s*\(|\.OpenTextFile\s*\([^)]*,\s*(2|8|ForWriting|ForAppending)\b|""Scripting\.FileSystemObject""|""ADODB\.Stream""")),
        ("registry change", R(@"\.(RegWrite|RegDelete)\b|\bReg(SetValue|SetValueEx|CreateKey|CreateKeyEx|DeleteKey|DeleteValue)[AW]?\s*\(|StdRegProv")),
        ("network", R(@"""(MSXML2\.(Server)?XMLHTTP[\d.]*|Microsoft\.XMLHTTP|WinHttp\.WinHttpRequest[\d.]*|InetCtls\.Inet[\d.]*|MSWinsock\.Winsock[\d.]*)""|\bURLDownloadToFile[AW]?\s*\(|\bInternet(Open|Connect|OpenUrl)[AW]?\s*\(")),
        ("script or code", R(@"""(MSScriptControl\.ScriptControl|ScriptControl)""|\.(AddCode|ExecuteStatement|ExecuteGlobal)\b|\bCallByName\s*\(|""VBIDE\.|\.VBE\b|\.CodeModule\b")),
        ("keystrokes", R(@"\b(SendKeys|AppActivate|keybd_event|SendInput|mouse_event)\b|\.SendKeys\b")),
        ("system", R(@"\b(ExitWindowsEx|InitiateSystemShutdown(Ex)?[AW]?|TerminateProcess|SetSystemTime|AdjustTokenPrivileges)\s*\(")),
    ];

    static Regex R(string p) => new(p, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>The first kind of side effect the expression looks like it has ("process start: Shell("), or null.</summary>
    public static string? SideEffect(string expression)
    {
        foreach (var (category, pattern) in Rules)
            if (pattern.Match(expression) is { Success: true } m) return $"{category}: {Collapse(m.Value)}";
        return null;
    }

    /// <summary>The clientInfo.name Claude Code sends in initialize (2.1.300, from the server log). Only Claude Code runs our hook.</summary>
    public const string ClaudeCodeClient = "claude-code";

    /// <summary>
    /// How an allow_side_effect=true call can be approved by the user (spike 40). The Claude Code hook asks the user before
    /// the call reaches us, but only Claude Code runs it: on a machine with the hook installed, Copilot CLI passed the flag and
    /// its expression ran with no prompt. So the hook counts only for Claude Code; other clients are asked through MCP
    /// elicitation when they offer it; anything else is refused.
    /// </summary>
    public static EvalApproval Approval(string clientId, bool hookInstalled, bool clientElicits) =>
        clientId == ClaudeCodeClient && hookInstalled ? EvalApproval.Hook
        : clientElicits ? EvalApproval.Elicitation
        : EvalApproval.None;

    /// <summary>
    /// Whether an elicitation result approves: action "accept" with content.approve = true. The form has one required
    /// boolean, so an accepted form whose box was left unticked is not an approval.
    /// </summary>
    public static bool Approves(System.Text.Json.Nodes.JsonObject? result) =>
        (string?)result?["action"] == "accept"
        && result["content"]?["approve"] is System.Text.Json.Nodes.JsonValue v && v.TryGetValue<bool>(out var b) && b;

    /// <summary>Kept for callers that only care about databases.</summary>
    public static string? DatabaseWrite(string expression) => SideEffect(expression) is { } s && s.StartsWith("database write", StringComparison.Ordinal) ? s : null;

    private static string Collapse(string s) => Regex.Replace(s.Trim(), @"\s+", " ");
}

/// <summary>Who asks the user before debug_eval runs a side effect: the Claude Code hook, the client through MCP elicitation, or nobody (refused).</summary>
public enum EvalApproval { None, Hook, Elicitation }
