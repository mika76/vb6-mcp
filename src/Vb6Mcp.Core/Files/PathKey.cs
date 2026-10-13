using System.Runtime.InteropServices;
using System.Text;

namespace Vb6Mcp.Core.Files;

/// <summary>
/// File identity = normalized absolute path (AGENTS.md hard rule 3): absolute, "..\" resolved,
/// either slash, subst drives mapped to their target, mapped network drives to their UNC path,
/// compared case-insensitively. Shared VB6 source files are reached through different relative
/// paths and drive letters, so anything keyed by a path must go through here.
/// </summary>
public static class PathKey
{
    /// <summary>Canonical full path (original casing kept). Relative paths resolve against <paramref name="baseDir"/>.</summary>
    public static string Normalize(string path, string? baseDir = null)
    {
        var p = path.Trim().Trim('"').Replace('/', '\\');
        if (!System.IO.Path.IsPathRooted(p) && baseDir != null) p = System.IO.Path.Combine(baseDir, p);
        p = System.IO.Path.GetFullPath(p);
        return ResolveDrive(p);
    }

    /// <summary>Comparison key: normalized and upper-cased (Windows file names are case-insensitive).</summary>
    public static string Key(string path, string? baseDir = null) => Normalize(path, baseDir).ToUpperInvariant();

    public static bool Same(string a, string b) => Key(a) == Key(b);

    /// <summary>Maps "X:\..." through subst (QueryDosDevice "\??\D:\real") or a network mapping (UNC).</summary>
    private static string ResolveDrive(string full)
    {
        if (full.Length < 2 || full[1] != ':') return full;
        string drive = full[..2];
        var target = new StringBuilder(1024);
        if (QueryDosDevice(drive, target, target.Capacity) > 0)
        {
            var t = target.ToString();
            if (t.StartsWith(@"\??\", StringComparison.Ordinal))         // subst drive
                return ResolveDrive(t[4..].TrimEnd('\\') + full[2..]);
        }
        var unc = new StringBuilder(1024);
        int len = unc.Capacity;
        if (WNetGetConnection(drive, unc, ref len) == 0)                   // mapped network drive
            return unc.ToString().TrimEnd('\\') + full[2..];
        return full;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int QueryDosDevice(string deviceName, StringBuilder targetPath, int max);

    [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
    private static extern int WNetGetConnection(string localName, StringBuilder remoteName, ref int length);
}
