// vb6mcp: MCP server for the VB6 IDE. See docs/architecture.md and docs/tools.md.
// Usage: vb6mcp [mcp]       MCP server on stdio (default; what Claude Code starts)
// Diagnostics go to stderr, or to the file named by VB6MCP_LOG (some clients hide stderr).
using System.Text;
using Vb6Mcp.Mcp;
using Vb6Mcp.Tools;

namespace Vb6Mcp;

public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        string mode = args.Length > 0 ? args[0].ToLowerInvariant() : "mcp";
        if (mode != "mcp")
        {
            Console.Error.WriteLine($"vb6mcp: unknown mode '{mode}'. Usage: vb6mcp [mcp]");
            return 2;
        }

        var logPath = Environment.GetEnvironmentVariable("VB6MCP_LOG");
        TextWriter log = logPath is { Length: > 0 }
            ? new StreamWriter(logPath, append: true, new UTF8Encoding(false)) { AutoFlush = true }
            : Console.Error;

        var stdin = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
        var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { NewLine = "\n" };
        var tools = new IdeTools(new IdeEnvironment()).All();
        log.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} vb6mcp {McpServer.ServerVersion} started, pid {Environment.ProcessId}, {tools.Count} tools");
        await new McpServer(stdin, stdout, log, tools).RunAsync();
        return 0;
    }
}
