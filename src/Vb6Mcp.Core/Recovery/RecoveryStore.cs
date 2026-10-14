using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Vb6Mcp.Core.Recovery;

/// <summary>One recovery copy: the IDE text of an unsaved component when it was taken.</summary>
public sealed record RecoveryCopy(string Project, string ProjectFile, string Component, int Type, string File, int Lines, string Hash, DateTime Taken);

/// <summary>
/// What is known about one IDE session (one VB6 process). <see cref="Note"/> marks a session of copies
/// kept on purpose (sync_take_disk); such sessions have their own id (<see cref="RecoveryStore.KeptId"/>)
/// and are kept like a crashed session, so the poller doesn't prune them as "saved".
/// </summary>
public sealed record RecoverySession(int Pid, string Started, DateTime Updated, bool Crashed, IReadOnlyList<string> ProjectFiles, string? Note = null);

/// <summary>
/// The watchdog's recovery store (PLAN.md §3.1): copies of unsaved IDE code, taken only while the
/// user is idle, so work survives a hang, crash or ide_kill. VB6 has no autosave. Layout:
/// %LOCALAPPDATA%\vb6mcp\recovery\&lt;pid&gt;\session.json plus &lt;project&gt;__&lt;component&gt;.txt (1252, the
/// code as the IDE shows it) and .json (metadata). The watchdog writes, the server reads. Copies of
/// components that were saved are removed; a cleanly closed IDE's folder is removed.
/// </summary>
public sealed class RecoveryStore(string? root = null)
{
    public string Root { get; } = root ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp", "recovery");

    private static readonly JsonSerializerOptions s_json = new() { WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull };

    /// <summary>Session id for copies kept on purpose from IDE <paramref name="pid"/> (not a process id).</summary>
    public static int KeptId(int pid) => 1_000_000_000 + pid;

    private string Dir(int pid) => Path.Combine(Root, pid.ToString(CultureInfo.InvariantCulture));
    private static string BaseName(string project, string component) => $"{project}__{component}";

    public void WriteSession(RecoverySession s)
    {
        Directory.CreateDirectory(Dir(s.Pid));
        WriteAtomic(Path.Combine(Dir(s.Pid), "session.json"), JsonSerializer.SerializeToUtf8Bytes(s, s_json));
    }

    public RecoverySession? ReadSession(int pid)
    {
        try { return JsonSerializer.Deserialize<RecoverySession>(File.ReadAllBytes(Path.Combine(Dir(pid), "session.json"))); }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { return null; }
    }

    /// <summary>All sessions with their folder (pids from folder names).</summary>
    public IReadOnlyList<RecoverySession> Sessions()
    {
        if (!Directory.Exists(Root)) return [];
        var list = new List<RecoverySession>();
        foreach (var d in Directory.GetDirectories(Root))
            if (int.TryParse(Path.GetFileName(d), NumberStyles.None, CultureInfo.InvariantCulture, out var pid) && ReadSession(pid) is { } s)
                list.Add(s);
        return list;
    }

    public void WriteCopy(int pid, RecoveryCopy meta, string text)
    {
        Directory.CreateDirectory(Dir(pid));
        var b = Path.Combine(Dir(pid), BaseName(meta.Project, meta.Component));
        WriteAtomic(b + ".txt", Vb1252.Encoding.GetBytes(text));
        WriteAtomic(b + ".json", JsonSerializer.SerializeToUtf8Bytes(meta, s_json));
    }

    public IReadOnlyList<RecoveryCopy> Copies(int pid)
    {
        if (!Directory.Exists(Dir(pid))) return [];
        var list = new List<RecoveryCopy>();
        foreach (var f in Directory.GetFiles(Dir(pid), "*__*.json"))
        {
            try { if (JsonSerializer.Deserialize<RecoveryCopy>(File.ReadAllBytes(f)) is { } c) list.Add(c); }
            catch (Exception ex) when (ex is IOException or JsonException) { }
        }
        return list.OrderBy(c => c.Project).ThenBy(c => c.Component).ToList();
    }

    /// <summary>The copied text, or null.</summary>
    public string? ReadText(int pid, string project, string component)
    {
        var f = Path.Combine(Dir(pid), BaseName(project, component) + ".txt");
        return File.Exists(f) ? Vb1252.Encoding.GetString(File.ReadAllBytes(f)) : null;
    }

    public void DeleteCopy(int pid, string project, string component)
    {
        var b = Path.Combine(Dir(pid), BaseName(project, component));
        TryDelete(b + ".txt"); TryDelete(b + ".json");
    }

    public void DeleteSession(int pid)
    {
        try { if (Directory.Exists(Dir(pid))) Directory.Delete(Dir(pid), recursive: true); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    private static void WriteAtomic(string path, byte[] bytes)
    {
        var tmp = path + ".tmp";
        File.WriteAllBytes(tmp, bytes);
        File.Move(tmp, path, overwrite: true);
    }

    private static void TryDelete(string f) { try { File.Delete(f); } catch (IOException) { } }
}
