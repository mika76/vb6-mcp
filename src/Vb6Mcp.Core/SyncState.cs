using System.Text.Json;
using System.Text.Json.Serialization;

namespace Vb6Mcp.Core;

/// <summary>A component whose file changed on disk while it had unsaved changes in an IDE.</summary>
public sealed record SyncConflict(int Pid, string Project, string Component, string File, DateTime Detected);

/// <summary>Something the sync engine did or saw: reloaded, conflict, pending, project-changed, deleted, resolved.</summary>
public sealed record SyncEvent(DateTime Time, int Pid, string Kind, string File, string? Project, string? Component, string? Note);

/// <summary>The sync engine's shared state.</summary>
public sealed record SyncSnapshot(IReadOnlyList<SyncConflict> Conflicts, IReadOnlyList<SyncEvent> Events, DateTime Updated);

/// <summary>
/// %LOCALAPPDATA%\vb6mcp\sync\state.json: written by the watchdog's sync poller, read by the server's
/// sync_* tools (PLAN.md §8). Atomic writes; a missing or unreadable file means "nothing known".
/// </summary>
public static class SyncState
{
    public static string DefaultPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp", "sync", "state.json");

    private static readonly JsonSerializerOptions s_json = new() { WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull };

    public static SyncSnapshot Read(string? path = null)
    {
        try { return JsonSerializer.Deserialize<SyncSnapshot>(File.ReadAllBytes(path ?? DefaultPath)) ?? Empty; }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { return Empty; }
    }

    public static void Write(SyncSnapshot s, string? path = null)
    {
        path ??= DefaultPath;
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var tmp = path + ".tmp";
        File.WriteAllBytes(tmp, JsonSerializer.SerializeToUtf8Bytes(s, s_json));
        File.Move(tmp, path, overwrite: true);
    }

    public static SyncSnapshot Empty => new([], [], DateTime.MinValue);
}
