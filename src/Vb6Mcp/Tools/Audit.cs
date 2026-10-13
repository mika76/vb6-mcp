namespace Vb6Mcp.Tools;

/// <summary>
/// Audit log of actions that send input or run code in the IDE (AGENTS.md hard rule 8): synthetic UI
/// input (app_click, app_set_text), Immediate evaluation, launches. One line per action in
/// %LOCALAPPDATA%\vb6mcp\audit.log, opened shared so several server instances can append.
/// </summary>
public static class Audit
{
    public static string DefaultPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp", "audit.log");

    /// <summary>Path of the audit log; tests point it elsewhere.</summary>
    public static string LogPath { get; set; } = DefaultPath;

    private static readonly object s_lock = new();

    public static void Write(string tool, int pid, string details)
    {
        try
        {
            lock (s_lock)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                using var fs = new FileStream(LogPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
                using var w = new StreamWriter(fs);
                w.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{Environment.ProcessId}] {tool} pid={pid} {details.Replace("\r", "\\r").Replace("\n", "\\n")}");
            }
        }
        catch (IOException) { }                 // auditing must never break the tool itself
        catch (UnauthorizedAccessException) { }
    }
}
