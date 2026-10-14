using System.Runtime.InteropServices;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

public static class DebugOutputTests
{
    [DllImport("kernel32.dll", CharSet = CharSet.Ansi)] static extern void OutputDebugStringA(string text);

    [Test]
    public static void CapturesOutputDebugStringOfThisProcess()
    {
        using var mon = new DebugOutput();
        var why = mon.Start();
        if (why != null) { Console.WriteLine($"    (skipped: {why})"); return; }    // DebugView or a server owns the buffer
        OutputDebugStringA("vb6mcp test: café über\r\n");
        var deadline = DateTime.UtcNow.AddSeconds(3);
        IReadOnlyList<DebugMessage> got = [];
        while (DateTime.UtcNow < deadline && (got = mon.Read(0, m => m.Pid == Environment.ProcessId)).Count == 0) Thread.Sleep(50);
        Assert.Equal(1, got.Count, "one message from this process");
        Assert.Equal("vb6mcp test: café über", got[0].Text, "ANSI text decoded, line break trimmed");
        Assert.True(mon.Read(got[0].Id).All(m => m.Id > got[0].Id), "afterId filters");
    }
}
