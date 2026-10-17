using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Vb6Mcp.Core.Files;

/// <summary>One file in a checkpoint: where it lives, and its copy (null when it didn't exist at checkpoint time).</summary>
public sealed record CheckpointFile(string Path, string? Copy, string? Sha256);

/// <summary>A saved copy of a project's files on disk: the .vbp, every member, their binary companions and the resource file.</summary>
public sealed record Checkpoint(string Id, string Vbp, string? Label, DateTime Created, IReadOnlyList<CheckpointFile> Files);

/// <summary>What a rollback did: files written back, files left as they were, failures, and members added since (left on disk).</summary>
public sealed record RollbackResult(IReadOnlyList<string> Restored, int Unchanged, IReadOnlyList<string> Failed, IReadOnlyList<string> AddedSince);

/// <summary>
/// Project checkpoints (PLAN.md E2): copies of a project's files kept under %LOCALAPPDATA%\vb6mcp\checkpoints, one folder
/// per project (keyed by its normalized path, hard rule 3) and one per checkpoint, with a manifest. The T9–T12 redesigns
/// changed hundreds of lines across many forms with no way back but the owner restoring files; a checkpoint lets an
/// agent try a bold change and undo it in one call. Copies are byte-exact (hard rule 1: bytes, never text).
/// </summary>
public sealed class Checkpoints(string? root = null)
{
    /// <summary>Checkpoints kept per project; the oldest go first.</summary>
    public const int Keep = 20;

    public string Root { get; } = root ?? System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vb6mcp", "checkpoints");

    // Written along with a component: forms .frx, UserControls .ctx, property pages .pgx, designers .dsx, UserDocuments .dox
    static readonly string[] Companions = [".frx", ".ctx", ".pgx", ".dsx", ".dox"];

    /// <summary>The files that make up the project now: the .vbp, members, their companions, the resource file.</summary>
    public static IReadOnlyList<string> ProjectFiles(string vbp)
    {
        var p = VbpFile.Load(vbp);
        var files = new List<string> { p.Path };
        foreach (var m in p.Members)
        {
            files.Add(m.FullPath);
            files.AddRange(Companions.Select(x => System.IO.Path.ChangeExtension(m.FullPath, x)).Where(File.Exists));
        }
        if (p.Settings.GetValueOrDefault("ResFile32") is { Length: > 0 } res)
            files.Add(PathKey.Normalize(res, System.IO.Path.GetDirectoryName(p.Path)!));
        return files.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    public Checkpoint Create(string vbp, string? label)
    {
        vbp = PathKey.Normalize(vbp);
        var id = DateTime.Now.ToString("yyyyMMdd-HHmmss-fff", CultureInfo.InvariantCulture);
        var dir = System.IO.Path.Combine(ProjectFolder(vbp), id);
        Directory.CreateDirectory(dir);
        var files = new List<CheckpointFile>();
        int n = 0;
        foreach (var path in ProjectFiles(vbp))
        {
            if (!File.Exists(path)) { files.Add(new CheckpointFile(path, null, null)); continue; }
            var bytes = File.ReadAllBytes(path);
            var copy = $"{++n:D4}{System.IO.Path.GetExtension(path)}";
            File.WriteAllBytes(System.IO.Path.Combine(dir, copy), bytes);
            files.Add(new CheckpointFile(path, copy, Sha(bytes)));
        }
        var cp = new Checkpoint(id, vbp, string.IsNullOrWhiteSpace(label) ? null : label.Trim(), DateTime.Now, files);
        File.WriteAllText(System.IO.Path.Combine(dir, "manifest.json"), JsonSerializer.Serialize(cp));
        foreach (var old in List(vbp).Skip(Keep)) TryDelete(System.IO.Path.Combine(ProjectFolder(vbp), old.Id));
        return cp;
    }

    /// <summary>The project's checkpoints, newest first.</summary>
    public IReadOnlyList<Checkpoint> List(string vbp)
    {
        var folder = ProjectFolder(PathKey.Normalize(vbp));
        if (!Directory.Exists(folder)) return [];
        return Directory.EnumerateDirectories(folder).Select(d => System.IO.Path.Combine(d, "manifest.json")).Where(File.Exists)
            .Select(m => { try { return JsonSerializer.Deserialize<Checkpoint>(File.ReadAllText(m)); } catch (JsonException) { return null; } })
            .OfType<Checkpoint>().OrderByDescending(c => c.Id, StringComparer.Ordinal).ToList();
    }

    /// <summary>A checkpoint by id, by label, or "latest"; null when there's none.</summary>
    public Checkpoint? Find(string vbp, string which)
    {
        var all = List(vbp);
        return which.Equals("latest", StringComparison.OrdinalIgnoreCase) ? all.FirstOrDefault()
            : all.FirstOrDefault(c => c.Id == which) ?? all.FirstOrDefault(c => string.Equals(c.Label, which, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Files that differ now from the checkpoint (changed, deleted, or created where there was none).</summary>
    public static IReadOnlyList<string> ChangedSince(Checkpoint cp)
        => cp.Files.Where(f => f.Sha256 is null ? File.Exists(f.Path) : !File.Exists(f.Path) || Sha(File.ReadAllBytes(f.Path)) != f.Sha256).Select(f => f.Path).ToList();

    /// <summary>
    /// Writes the checkpoint's files back where they differ. Files that didn't exist at checkpoint time and members added
    /// since are left on disk and listed (the restored .vbp no longer names them); deleting is the agent's decision.
    /// </summary>
    public RollbackResult Restore(Checkpoint cp)
    {
        var dir = System.IO.Path.Combine(ProjectFolder(cp.Vbp), cp.Id);
        var restored = new List<string>(); var failed = new List<string>(); int unchanged = 0;
        var before = new HashSet<string>(cp.Files.Select(f => f.Path), StringComparer.OrdinalIgnoreCase);
        var added = File.Exists(cp.Vbp) ? ProjectFiles(cp.Vbp).Where(p => !before.Contains(p)).ToList() : [];
        added.AddRange(cp.Files.Where(f => f.Copy is null && File.Exists(f.Path)).Select(f => f.Path));
        foreach (var f in cp.Files.Where(f => f.Copy is not null))
        {
            try
            {
                var bytes = File.ReadAllBytes(System.IO.Path.Combine(dir, f.Copy!));
                if (File.Exists(f.Path) && File.ReadAllBytes(f.Path).AsSpan().SequenceEqual(bytes)) { unchanged++; continue; }
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(f.Path)!);
                var tmp = f.Path + ".vb6mcp.tmp";
                File.WriteAllBytes(tmp, bytes);
                File.Move(tmp, f.Path, overwrite: true);
                restored.Add(f.Path);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { failed.Add($"{f.Path}: {ex.Message}"); }
        }
        return new RollbackResult(restored, unchanged, failed, added.Distinct(StringComparer.OrdinalIgnoreCase).ToList());
    }

    private string ProjectFolder(string vbp)
    {
        var key = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(PathKey.Normalize(vbp).ToUpperInvariant())))[..12];
        return System.IO.Path.Combine(Root, System.IO.Path.GetFileNameWithoutExtension(vbp).Replace(' ', '_') + "-" + key);
    }

    private static string Sha(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));

    private static void TryDelete(string dir) { try { Directory.Delete(dir, true); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { } }
}
