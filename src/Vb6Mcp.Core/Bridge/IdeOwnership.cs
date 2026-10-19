using System.Diagnostics;
using System.Text.Json;

namespace Vb6Mcp.Core.Bridge;

/// <summary>The MCP server session that owns an IDE: its process (pid + start time, so a reused pid doesn't count) and client.</summary>
public sealed record IdeOwner(int ServerPid, long ServerStartTicks, string Client, DateTime Since);

/// <summary>
/// Which MCP server session an IDE belongs to (PLAN Phase 7 A9). Several IDEs always worked (a discovery file per pid,
/// pid on every tool), but nothing said whose an IDE was: on 2026-10-17 a test script of one Claude session saved
/// another session's project, because the tools default to the only running IDE.
/// <para>
/// An IDE is owned by the session that started it (ide_launch) or claimed it (ide_claim). Owner files live in
/// %LOCALAPPDATA%\vb6mcp\owners\&lt;ide pid&gt;.json, not next to the discovery files (Discovery reads every *.json there).
/// An owner whose server process is gone owns nothing, and its file is removed. This is coordination between the
/// user's own sessions, not a security boundary: every session runs as the same user.
/// </para>
/// </summary>
public static class IdeOwnership
{
    public static string DefaultFolder =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp", "owners");

    private static readonly Lazy<long> s_myStart = new(() => StartTicks(Environment.ProcessId) ?? 0);

    /// <summary>This server process as an owner.</summary>
    public static IdeOwner Me(string client) => new(Environment.ProcessId, s_myStart.Value, client, DateTime.UtcNow);

    /// <summary>True when <paramref name="owner"/> is this server process.</summary>
    public static bool IsMe(IdeOwner owner) => owner.ServerPid == Environment.ProcessId && owner.ServerStartTicks == s_myStart.Value;

    /// <summary>The live owner of an IDE, or null (no owner file, unreadable, or the owner process has gone).</summary>
    public static IdeOwner? Read(int idePid, string? folder = null, Func<IdeOwner, bool>? isAlive = null)
    {
        var file = FileOf(idePid, folder);
        IdeOwner? owner;
        try { owner = File.Exists(file) ? JsonSerializer.Deserialize<IdeOwner>(File.ReadAllBytes(file)) : null; }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { return null; }
        if (owner is null) return null;
        if (!(isAlive ?? Alive)(owner)) { TryDelete(file); return null; }
        return owner;
    }

    /// <summary>Records <paramref name="owner"/> as the owner of the IDE.</summary>
    public static void Claim(int idePid, IdeOwner owner, string? folder = null)
    {
        var file = FileOf(idePid, folder);
        Directory.CreateDirectory(Path.GetDirectoryName(file)!);
        var tmp = file + ".tmp";
        File.WriteAllBytes(tmp, JsonSerializer.SerializeToUtf8Bytes(owner));
        File.Move(tmp, file, overwrite: true);
    }

    /// <summary>Removes the IDE's owner file.</summary>
    public static void Release(int idePid, string? folder = null) => TryDelete(FileOf(idePid, folder));

    // ---------- handover across server_restart (improvement notes 53) ----------

    /// <summary>What a server that is about to end (server_restart) leaves for the server its Claude session starts next. Runs: IDEs whose current run it started (run_end ends those without force, notes 56).</summary>
    public sealed record Handover(string Session, int OldServerPid, int[] Ides, DateTime Written, int[]? Runs = null);

    /// <summary>How long a handover waits for the next server (the desktop app starts it on the session's next vb6 call).</summary>
    public static readonly TimeSpan HandoverMaxAge = TimeSpan.FromHours(12);

    /// <summary>
    /// The Claude session this server belongs to: CLAUDE_CODE_SESSION_ID, which Claude Code passes to the MCP servers it
    /// starts (checked live 2026-10-19). Null when it isn't set: then there is no handover, and a restarted server has to
    /// ide_claim its IDEs.
    /// </summary>
    public static string? SessionKey() => Environment.GetEnvironmentVariable("CLAUDE_CODE_SESSION_ID") is { Length: > 0 } s ? s : null;

    /// <summary>
    /// Before this server ends: records the IDEs it owns for the next server of the same session (server_restart dropped
    /// them: ownership is by server process, so the restarted server found its own IDE "not this session's"). Returns them.
    /// </summary>
    public static IReadOnlyList<int> WriteHandover(string? session, string? folder = null, IEnumerable<int>? runs = null)
    {
        if (session is null) return [];
        var dir = folder ?? DefaultFolder;
        if (!Directory.Exists(dir)) return [];
        var mine = new List<int>();
        foreach (var f in Directory.EnumerateFiles(dir, "*.json"))
            if (int.TryParse(Path.GetFileNameWithoutExtension(f), out var ide) && ReadRaw(f) is { } o && IsMe(o)) mine.Add(ide);
        if (mine.Count == 0) return [];
        var file = HandoverFile(session, dir);
        File.WriteAllBytes(file + ".tmp", JsonSerializer.SerializeToUtf8Bytes(new Handover(session, Environment.ProcessId, [.. mine], DateTime.UtcNow, [.. (runs ?? []).Where(mine.Contains)])));
        File.Move(file + ".tmp", file, overwrite: true);
        return mine;
    }

    /// <summary>
    /// At startup: claims the IDEs a previous server of the same session handed over, when they still run and nobody else
    /// claimed them since (owner file gone, or still naming the old server). The handover file is used once. runs is filled with the taken IDEs whose run the old server had started.
    /// </summary>
    public static IReadOnlyList<int> TakeOver(string? session, string? folder = null, Func<int, bool>? ideAlive = null, Func<IdeOwner, bool>? isAlive = null, List<int>? runs = null)
    {
        if (session is null) return [];
        var dir = folder ?? DefaultFolder;
        var file = HandoverFile(session, dir);
        Handover? h;
        try { h = File.Exists(file) ? JsonSerializer.Deserialize<Handover>(File.ReadAllBytes(file)) : null; }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { h = null; }
        TryDelete(file);
        if (h is null || h.Session != session || DateTime.UtcNow - h.Written > HandoverMaxAge) return [];
        var taken = new List<int>();
        foreach (var ide in h.Ides)
        {
            if (!(ideAlive ?? Discovery.IsVb6Process)(ide)) continue;
            var current = ReadRaw(FileOf(ide, dir));
            if (current is not null && current.ServerPid != h.OldServerPid && (isAlive ?? Alive)(current)) continue;   // someone else claimed it since
            Claim(ide, Me(current?.Client ?? "(restarted server)"), dir);
            taken.Add(ide);
            if (h.Runs?.Contains(ide) == true) runs?.Add(ide);
        }
        return taken;
    }

    private static string HandoverFile(string session, string dir) =>
        Path.Combine(dir, "handover-" + Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(session)))[..16] + ".handover");

    private static IdeOwner? ReadRaw(string file)
    {
        try { return File.Exists(file) ? JsonSerializer.Deserialize<IdeOwner>(File.ReadAllBytes(file)) : null; }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException) { return null; }
    }

    private static string FileOf(int idePid, string? folder) => Path.Combine(folder ?? DefaultFolder, idePid + ".json");

    private static bool Alive(IdeOwner o) => o.ServerPid > 0 && StartTicks(o.ServerPid) is long t && Math.Abs(t - o.ServerStartTicks) < TimeSpan.TicksPerSecond;

    private static long? StartTicks(int pid)
    {
        try { using var p = Process.GetProcessById(pid); return p.StartTime.ToUniversalTime().Ticks; }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception) { return null; }
    }

    private static void TryDelete(string file)
    {
        try { File.Delete(file); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }
}
