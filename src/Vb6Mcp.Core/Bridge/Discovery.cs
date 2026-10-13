using System.Diagnostics;
using System.Text.Json;

namespace Vb6Mcp.Core.Bridge;

/// <summary>One add-in instance as described by its discovery file (docs/protocol.md "Discovery").</summary>
public sealed record InstanceInfo(int Pid, int Port, string Token, int Protocol, string AddinVersion, string Started, string FilePath);

/// <summary>Result of scanning the discovery folder.</summary>
public sealed record DiscoveryResult(IReadOnlyList<InstanceInfo> Live, IReadOnlyList<InstanceInfo> Stale, IReadOnlyList<string> Unreadable);

/// <summary>
/// Finds running add-in instances via %LOCALAPPDATA%\vb6mcp\instances\&lt;pid&gt;.json.
/// A clean unload deletes the file; a file whose process is gone (or whose PID was reused by
/// another program) belongs to a killed/crashed IDE (spike 12) and is "stale".
/// </summary>
public static class Discovery
{
    public static string DefaultFolder =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp", "instances");

    /// <summary>Scans the folder. <paramref name="isVb6Alive"/> is injectable for tests.</summary>
    public static DiscoveryResult Scan(string? folder = null, Func<int, bool>? isVb6Alive = null)
    {
        folder ??= DefaultFolder;
        isVb6Alive ??= IsVb6Process;
        var live = new List<InstanceInfo>();
        var stale = new List<InstanceInfo>();
        var bad = new List<string>();
        if (!Directory.Exists(folder)) return new DiscoveryResult(live, stale, bad);

        foreach (var file in Directory.GetFiles(folder, "*.json"))
        {
            InstanceInfo? info = TryRead(file);
            if (info is null) { bad.Add(file); continue; }
            (isVb6Alive(info.Pid) ? live : stale).Add(info);
        }
        return new DiscoveryResult(live, stale, bad);
    }

    /// <summary>Deletes the discovery files of stale instances. Returns how many were deleted.</summary>
    public static int DeleteStale(DiscoveryResult result)
    {
        int n = 0;
        foreach (var s in result.Stale)
        {
            try { File.Delete(s.FilePath); n++; } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
        return n;
    }

    /// <summary>Parses a discovery file; null if it is missing fields or isn't valid JSON (e.g. half-written).</summary>
    public static InstanceInfo? TryRead(string file)
    {
        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(file, Vb1252.Encoding));
            var r = doc.RootElement;
            return new InstanceInfo(
                r.GetProperty("pid").GetInt32(),
                r.GetProperty("port").GetInt32(),
                r.GetProperty("token").GetString() ?? throw new FormatException("token"),
                r.TryGetProperty("protocol", out var p) ? p.GetInt32() : 0,
                r.TryGetProperty("addin", out var a) ? a.GetString() ?? "" : "",
                r.TryGetProperty("started", out var s) ? s.GetString() ?? "" : "",
                file);
        }
        catch (Exception ex) when (ex is JsonException or KeyNotFoundException or FormatException or IOException or InvalidOperationException)
        {
            return null;
        }
    }

    /// <summary>True if <paramref name="pid"/> is a running VB6.EXE (PIDs get reused, so the name matters).</summary>
    public static bool IsVb6Process(int pid)
    {
        try
        {
            using var p = Process.GetProcessById(pid);
            return !p.HasExited && p.ProcessName.Equals("VB6", StringComparison.OrdinalIgnoreCase);
        }
        catch (ArgumentException) { return false; }
        catch (InvalidOperationException) { return false; }
    }
}
