using System.Diagnostics;
using System.Text;
using Vb6Mcp.Core;

namespace Vb6Mcp.Tests;

/// <summary>
/// Tests that need a fresh server process: type initializers have already run in the test runner's process, so a bug
/// in their order only shows when the server's very first call takes that path.
/// </summary>
public static class FreshProcessTests
{
    /// <summary>
    /// An edit as the first call of a new server. VbCode's strict 1252 encoding was a field initializer that ran before
    /// the static constructor registered the code-pages provider: TypeInitializationException on the first edit.
    /// </summary>
    [Test]
    public static void AnEditCanBeTheFirstCallOfANewServer()
    {
        var root = Directory.CreateTempSubdirectory("vb6mcp-fresh-").FullName;
        try
        {
            var file = Path.Combine(root, "modA.bas");
            File.WriteAllBytes(file, Vb1252.Encoding.GetBytes("Attribute VB_Name = \"modA\"\r\nOption Explicit\r\n"));
            string hash = VbCode.Hash("Option Explicit");   // the server's file_read hash of the code part
            // the server's own launcher: it is built x86 (like VB6), and dotnet on PATH is usually the x64 host
            var server = Path.Combine(AppContext.BaseDirectory, "vb6mcp.exe");
            var psi = new ProcessStartInfo(server)
            {
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false,
                StandardInputEncoding = new UTF8Encoding(false), StandardOutputEncoding = Encoding.UTF8,
            };
            psi.Environment["VB6MCP_ROOTS"] = root;
            using var p = Process.Start(psi)!;
            p.StandardInput.WriteLine("{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{\"protocolVersion\":\"2025-06-18\",\"capabilities\":{},\"clientInfo\":{\"name\":\"test\",\"version\":\"1\"}}}");
            p.StandardInput.WriteLine("{\"jsonrpc\":\"2.0\",\"id\":2,\"method\":\"tools/call\",\"params\":{\"name\":\"file_replace_lines\",\"arguments\":" +
                          System.Text.Json.JsonSerializer.Serialize(new { path = file, start_line = 2, count = 0, new_text = "' first", expected_hash = hash }) + "}}");
            p.StandardInput.Flush();
            // keep stdin open until the call's answer is in: the server stops reading (and exits) at end of input
            var output = new StringBuilder();
            var read = Task.Run(() => { string? l; while ((l = p.StandardOutput.ReadLine()) != null) { output.AppendLine(l); if (l.Contains("\"id\":2")) break; } });
            if (!read.Wait(30_000)) { p.Kill(); throw new InvalidOperationException("no answer from the server: " + output); }
            p.StandardInput.Close();
            p.WaitForExit(10_000);
            Assert.True(!output.ToString().Contains("TypeInitializationException") && !output.ToString().Contains("\"isError\":true"), output.ToString());
            Assert.True(Vb1252.Encoding.GetString(File.ReadAllBytes(file)).Contains("' first\r\n"), "the edit reached the file; server said: " + output);
        }
        finally { Directory.Delete(root, true); }
    }

    /// <summary>server_restart (PLAN.md E7) answers, then ends the process while stdin is still open.</summary>
    [Test]
    public static void ServerRestartAnswersAndThenExits()
    {
        var psi = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "vb6mcp.exe"))
        {
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false,
            StandardInputEncoding = new UTF8Encoding(false), StandardOutputEncoding = Encoding.UTF8,
        };
        using var p = Process.Start(psi)!;
        p.StandardInput.WriteLine("{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/call\",\"params\":{\"name\":\"server_restart\",\"arguments\":{\"confirm\":true}}}");
        p.StandardInput.Flush();
        var answer = Task.Run(() => p.StandardOutput.ReadLine());
        Assert.True(answer.Wait(20_000) && (answer.Result ?? "").Contains("exiting"), "answered first: " + answer.Result);
        Assert.True(p.WaitForExit(10_000), "the process ended with stdin still open");
        Assert.Equal(3, p.ExitCode);
    }
}
