using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace Vb6Mcp.Core.Bridge;

/// <summary>
/// The watchdog's discovery file, %LOCALAPPDATA%\vb6mcp\watchdog.json. It has the same fields as an
/// add-in's discovery file (docs/protocol.md "Discovery"), so <see cref="Discovery.TryRead"/> and
/// <see cref="AddinClient"/> work for the watchdog too. The profile folder's ACL keeps it user-only.
/// </summary>
public static class WatchdogDiscovery
{
    public static string DefaultPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp", "watchdog.json");

    /// <summary>Writes the file atomically (temp + move), so a reader never sees half of it.</summary>
    public static void Write(string path, int pid, int port, string token, string version)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string json = string.Create(CultureInfo.InvariantCulture,
            $"{{\"pid\":{pid},\"port\":{port},\"token\":\"{token}\",\"protocol\":1,\"addin\":\"watchdog {version}\",\"started\":\"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\"}}");
        var tmp = path + ".tmp";
        File.WriteAllText(tmp, json, new UTF8Encoding(false));
        File.Move(tmp, path, overwrite: true);
    }

    /// <summary>The running watchdog, or null if there is none (no file, or its process is gone or isn't vb6mcp).</summary>
    public static InstanceInfo? Find(string? path = null)
    {
        path ??= DefaultPath;
        if (!File.Exists(path)) return null;
        var info = Discovery.TryRead(path);
        return info != null && IsWatchdogProcess(info.Pid) ? info : null;
    }

    private static bool IsWatchdogProcess(int pid)
    {
        try
        {
            using var p = Process.GetProcessById(pid);
            return !p.HasExited && p.ProcessName.Equals("vb6mcp", StringComparison.OrdinalIgnoreCase);
        }
        catch (ArgumentException) { return false; }
        catch (InvalidOperationException) { return false; }
    }
}
