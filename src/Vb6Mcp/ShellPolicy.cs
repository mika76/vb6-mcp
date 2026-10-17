using System.Text.RegularExpressions;
using Vb6Mcp.Tools;

namespace Vb6Mcp;

/// <summary>
/// What the Claude Code hook refuses in the agent's own Bash/PowerShell commands (PLAN.md E5b). In the T12 rerun both agents
/// read the database login from the project's files and ran sqlcmd; one changed a row with an UPDATE. debug_eval's guard
/// couldn't see that, and the owner asked for shell commands to be covered too ("security should be taken seriously").
/// The vb6 tools already refuse or audit these actions; the shell must not be a way around them:
/// - database writes through command-line clients or scripts (reads stay allowed: agents look up test data);
/// - writes to VB6 source files (hard rules 1, 2, 4: 1252, hidden lines, the IDE's memory; benchmarks saw .frm files
///   rewritten through Python);
/// - COM registration and registry edits (hard rule 13: only through the elevated, audited tools);
/// - ending or building with VB6 directly (ide_quit / ide_kill / build_make keep unsaved work and the user's session safe).
/// Text matching, so a determined script can get past it; it catches what agents actually did. VB6MCP_SHELL_GUARD=off
/// (in Claude Code's settings env) turns it off.
/// </summary>
public static class ShellPolicy
{
    const string Vb6Ext = @"\.(bas|cls|frm|frx|ctl|ctx|pag|pgx|dsr|dsx|dob|dox|vbp|vbg|vbw)\b";
    static Regex R(string p) => new(p, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline);

    static readonly Regex DbClient = R(@"\b(sqlcmd|osql|isql|bcp|sqlplus|mysql|psql|sqlite3|Invoke-Sqlcmd|Invoke-DbaQuery)\b|SqlClient|SqlConnection|OleDbConnection|OdbcConnection|ADODB\.Connection|ADODB\.Recordset|pyodbc|pymssql|sqlalchemy|System\.Data\.");
    static readonly Regex BcpIn = R(@"\bbcp\b.*\s(in|format)\s");

    static readonly Regex[] Vb6Writes =
    [
        R(@">{1,2}\s*[""']?[^\s|&;<>]*" + Vb6Ext),                                              // redirection into a VB6 file
        R(@"\b(Set-Content|Add-Content|Out-File|Clear-Content|Tee-Object)\b.*" + Vb6Ext),
        R(@"\b(Remove-Item|Rename-Item|Move-Item|rm|del|erase|ren|rename|mv|move)\b[^|;&]*" + Vb6Ext),
        R(@"\b(Copy-Item|cp|copy|xcopy|robocopy)\b[^|;&]*" + Vb6Ext + @"[""']?\s*($|[|;&])"),     // copied onto a VB6 file (the destination comes last)
        R(@"\b(sed|perl)\s+(-\w*\s+)*-i\b.*" + Vb6Ext),
        R(@"(WriteAll(Text|Bytes|Lines)|AppendAll(Text|Lines)|StreamWriter|\.write_(text|bytes)\s*\(|\bopen\s*\([^)]*,\s*[""'](w|a|wb|ab|w\+|r\+|a\+)[""']|writeFileSync|appendFileSync)" + ".*" + Vb6Ext),
        R(Vb6Ext + @".*(WriteAll(Text|Bytes|Lines)|AppendAll(Text|Lines)|StreamWriter|\.write_(text|bytes)\s*\(|\bopen\s*\([^)]*,\s*[""'](w|a|wb|ab|w\+|r\+|a\+)[""']|writeFileSync|appendFileSync|\.write\s*\()"),
    ];

    static readonly Regex Registration = R(@"\b(regsvr32|RegAsm|regtlib|regtlibv12)\b|\breg(\.exe)?\s+(add|delete|import|copy|load|restore|unload)\b|\b(New-ItemProperty|Set-ItemProperty|Remove-ItemProperty|Rename-ItemProperty)\b|\b(Remove-Item|New-Item|Set-Item)\b[^|;&]*(HK(LM|CU|CR|U|CC):|Registry::)");
    static readonly Regex Vb6Process = R(@"\b(taskkill|Stop-Process|kill|pskill)\b[^|;&]*\bvb6\b|\bvb6(\.exe)?[""']?\s+/(make|m|run|r|d)\b");

    /// <summary>Why the command is refused, or null to allow it.</summary>
    public static string? Check(string command)
    {
        if (string.IsNullOrWhiteSpace(command)) return null;
        if ((DbClient.IsMatch(command) && EvalPolicy.Sql.IsMatch(command)) || BcpIn.IsMatch(command))
            return $"this command looks like it changes a database ({(BcpIn.IsMatch(command) ? "bcp in" : Regex.Replace(EvalPolicy.Sql.Match(command).Value, @"\s+", " "))}). " +
                   "Agents should set up test data through the program's own UI (app_click, app_set_text); reading is fine.";
        // per line: a script that reads .frm files and writes a .json elsewhere isn't a VB6 write (T12 round 8 false positive:
        // the .frm name and the open(..., 'w') were on different lines of one Python heredoc)
        if (command.Split('\n').Any(line => Vb6Writes.Any(r => r.IsMatch(line))))
            return "this command looks like it writes a VB6 source file (Windows-1252 with hidden header and Attribute lines, possibly open in the VB6 IDE). " +
                   "The vb6 tools do this safely: code_replace_lines / proc_replace (open project), file_replace_lines (closed), form_* for designers, project_checkpoint / project_rollback for copies.";
        if (Registration.IsMatch(command))
            return "this command registers COM components or changes the registry, which the vb6 tools do elevated and audited (com_register, com_unregister, refs_*). Reading the registry (reg query) is fine.";
        if (Vb6Process.IsMatch(command))
            return "this command ends VB6 or builds with it directly, which can lose unsaved work and needs elevation. The vb6 tools: ide_quit, ide_kill force=true (hung IDE), build_make.";
        return null;
    }
}
