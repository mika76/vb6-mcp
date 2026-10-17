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
