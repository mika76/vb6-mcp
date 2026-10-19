namespace Vb6Mcp.Tests;

/// <summary>The hook's checks on the agent's own shell commands (PLAN.md E5b).</summary>
public static class ShellPolicyTests
{
    [Test]
    public static void DatabaseWritesAreBlockedAndReadsAllowed()
    {
        // the T12 rerun's commands, with the login replaced
        Assert.True(ShellPolicy.Check("sqlcmd -S \".\\SQLEXPRESS\" -U sa -P 'x' -d AppDb -Q \"select retries from orders where order_number=183; update orders set retries=3 where order_number=183\"") is not null, "update via sqlcmd");
        Assert.True(ShellPolicy.Check("sqlcmd -S . -d AppDb -h -1 -Q \"set nocount on; select top 3 order_number from orders\"") is null, "select via sqlcmd");
        Assert.True(ShellPolicy.Check("Invoke-Sqlcmd -Query 'DELETE FROM log'") is not null, "Invoke-Sqlcmd delete");
        Assert.True(ShellPolicy.Check("python -c \"import pyodbc; c=pyodbc.connect(s); c.execute('insert into t values (1)')\"") is not null, "pyodbc insert");
        Assert.True(ShellPolicy.Check("bcp AppDb.dbo.orders in data.txt -T") is not null, "bcp in");
        Assert.True(ShellPolicy.Check("git commit -m \"update readme set up\"") is null, "no database client: not a database write");
    }

    [Test]
    public static void Vb6SourceWritesAreBlockedAndReadsAllowed()
    {
        foreach (var bad in new[]
        {
            "echo x > frmMain.frm", "Set-Content -Path 'D:\\p\\modA.bas' -Value $t", "rm Common/frmMonitor.frm",
            "sed -i 's/a/b/' frmMain.frm", "python - <<'EOF'\nopen('D:/p/frmMain.frm','w').write(s)\nEOF",
            "cp backup/frmMain.frm 'APP/frmMain.frm'", "[IO.File]::WriteAllBytes('D:\\p\\App.vbp', $b)",
            // in-place line-ending conversion (notes from real use, 2026-10-19: a .vbp rewritten with LF)
            "unix2dos -q App.vbp", "dos2unix Common/modA.bas", "cd src && todos frmMain.frm", "flip -m App.vbp",
        })
            Assert.True(ShellPolicy.Check(bad) is not null, bad);
        foreach (var ok in new[]
        {
            "wc -l frmMain.frm frmAddCards.frm", "grep -n Caption frmMain.frm", "Get-Content frmMain.frm | Select-Object -First 5",
            "python - <<'EOF'\ntxt=open(f,encoding='cp1252').read()\nEOF", "cp frmMain.frm /tmp/backup/", "dotnet build src/Vb6Mcp",
            // T12 round 8: reads .frm files, writes a .json: different lines
            "python - <<'EOF'\nfiles={'a':'System/frmDenomBreakdown.frm'}\ntxt=open(f,encoding='cp1252').read()\njson.dump(out, open('changes.json','w'))\nEOF",
            // "flip" is an ordinary word: only a command position counts, and other files aren't ours
            "git commit -m \"flip the order of the fields in frmMain.frm\"", "unix2dos build/notes.txt",
        })
            Assert.True(ShellPolicy.Check(ok) is null, ok);
    }

    [Test]
    public static void RegistrationRegistryAndVb6ProcessesAreBlocked()
    {
        foreach (var bad in new[] { "regsvr32 /s x.ocx", "RegAsm.exe a.dll /tlb", "reg add HKCU\\Software\\X /v a /d b", "Set-ItemProperty HKLM:\\Software\\X -Name a -Value 1",
                                     "taskkill /f /im vb6.exe", "Stop-Process -Name VB6", "\"C:\\VB98\\VB6.EXE\" /make x.vbp",
                                     // what an unhooked agent ran (full T13-T18 run), and the other ways to the same end
                                     "Stop-Process -Name VB6 -Force -ErrorAction SilentlyContinue", "Get-Process VB6 | Stop-Process -Force",
                                     "wmic process where name='VB6.EXE' delete", "(Get-Process vb6).Kill()",
                                     "Get-CimInstance Win32_Process -Filter \"Name='VB6.EXE'\" | Invoke-CimMethod -MethodName Terminate",
                                     // a build where the exe is in a variable and /make in an array, so neither is next to the other
                                     "Start-Process -FilePath $exe -ArgumentList '/make','App.vbp','/out','build.log' -Wait",
                                     "& $exe /make App.vbp /out build.log", "cmd /c \"C:\\VB98\\VB6.EXE\" /make App.vbp",
                                     "Start-Process 'C:\\VB98\\VB6.EXE' -ArgumentList '/make', $proj -Wait" })
            Assert.True(ShellPolicy.Check(bad) is not null, bad);
        foreach (var ok in new[] { "reg query HKCR\\TypeLib", "Get-ItemProperty HKLM:\\Software\\X", "Get-Process vb6",
                                   // naming a project without building it, and paths that only contain "make"
                                   "Start-Process notepad App.vbp", "Start-Process -FilePath $exe -ArgumentList 'App.vbp'",
                                   "Get-ChildItem D:/make/App.vbp", "grep -n make build.log" })
            Assert.True(ShellPolicy.Check(ok) is null, ok);
    }

    static (int Code, string Out) RunHook(string json, string? guard = null)
    {
        var output = new StringWriter();
        return (Hook.Run(new StringReader(json), TextWriter.Null, k => k == "VB6MCP_SHELL_GUARD" ? guard : null, output), output.ToString());
    }

    /// <summary>The owner chose asking over refusing: the hook answers permissionDecision "ask" with the reason.</summary>
    [Test]
    public static void TheHookAsksAboutFlaggedShellCommandsUnlessTurnedOff()
    {
        var json = """{"tool_name":"Bash","tool_input":{"command":"sqlcmd -Q \"delete from orders\""}}""";
        var (code, output) = RunHook(json);
        Assert.Equal(0, code);
        var decision = System.Text.Json.Nodes.JsonNode.Parse(output)!["hookSpecificOutput"]!;
        Assert.Equal("ask", (string)decision["permissionDecision"]!);
        Assert.True(((string)decision["permissionDecisionReason"]!).Contains("database"), output);
        Assert.Equal("", RunHook(json, "off").Out);
        Assert.Equal("", RunHook("""{"tool_name":"PowerShell","tool_input":{"command":"Get-ChildItem"}}""").Out);
    }

    [Test]
    public static void TheHookAsksAboutDebugEvalSideEffectsOnlyWhenAllowIsRequested()
    {
        Assert.Equal("", RunHook("""{"tool_name":"mcp__vb6__debug_eval","tool_input":{"expression":"Shell(\"calc\")"}}""").Out);   // the server refuses on its own
        Assert.True(RunHook("""{"tool_name":"mcp__vb6__debug_eval","tool_input":{"expression":"Shell(\"calc\")","allow_side_effect":true}}""").Out.Contains("\"ask\""), "asks");
        Assert.Equal("", RunHook("""{"tool_name":"mcp__vb6__debug_eval","tool_input":{"expression":"x + 1","allow_side_effect":true}}""").Out);   // no side effect
    }

    [Test]
    public static void GuardsEvalNeedsOurHookWithTheDebugEvalMatcher()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-guard-").FullName;
        try
        {
            var settings = Path.Combine(dir, "settings.json");
            Assert.True(!Hook.GuardsEval(settings), "no file");
            File.WriteAllText(settings, """{"hooks":{"PreToolUse":[{"matcher":"Edit|Write|MultiEdit","hooks":[{"type":"command","command":"\"C:\\x\\vb6mcp.exe\" hook"}]}]}}""");
            Assert.True(!Hook.GuardsEval(settings), "an old matcher without debug_eval");
            Hook.Install(settings, @"C:\x\vb6mcp.exe", TextWriter.Null);
            Assert.True(Hook.GuardsEval(settings), "installed now");
        }
        finally { Directory.Delete(dir, true); }
    }
}
