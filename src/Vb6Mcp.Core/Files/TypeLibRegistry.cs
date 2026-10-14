using System.Globalization;
using Microsoft.Win32;

namespace Vb6Mcp.Core.Files;

/// <summary>A registered typelib version: HKCR\TypeLib\{LibId}\{maj.min} (hex) with its description and 32-bit file.</summary>
public sealed record RegisteredTypeLib(Guid LibId, int Major, int Minor, string Description, string? Win32Path)
{
    public bool FileExists => Win32Path is { Length: > 0 } p && ComRegistry.Exists(p);
}

/// <summary>
/// The registered typelibs in the 32-bit registry view (VB6 is 32-bit). This is where VB resolves a
/// "Reference=*\G{LIBID}#maj.min" line, so references are added from here: GUIDs come from the registry,
/// never typed (hard rule 13). Version keys are hex ("b.0" = 11.0, as seen on this machine).
/// </summary>
public static class TypeLibRegistry
{
    /// <summary>All registered typelib versions (any locale; the win32 path of the first locale that has one).</summary>
    public static IReadOnlyList<RegisteredTypeLib> All()
    {
        var list = new List<RegisteredTypeLib>();
        using var root = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32).OpenSubKey("TypeLib");
        if (root is null) return list;
        foreach (var g in root.GetSubKeyNames())
        {
            if (!Guid.TryParse(g, out var libId)) continue;
            using var gk = root.OpenSubKey(g);
            if (gk is null) continue;
            foreach (var v in gk.GetSubKeyNames())
            {
                if (!TryVersion(v, out int major, out int minor)) continue;
                using var vk = gk.OpenSubKey(v);
                if (vk is null) continue;
                list.Add(new RegisteredTypeLib(libId, major, minor, vk.GetValue(null) as string ?? "", Win32Path(vk)));
            }
        }
        return list;
    }

    /// <summary>One registered version, or null.</summary>
    public static RegisteredTypeLib? Find(Guid libId, int major, int minor)
        => All().FirstOrDefault(t => t.LibId == libId && t.Major == major && t.Minor == minor);

    /// <summary>Parses a hex "maj.min" version key.</summary>
    public static bool TryVersion(string key, out int major, out int minor)
    {
        major = minor = 0;
        var parts = key.Split('.');
        return parts.Length == 2
            && int.TryParse(parts[0], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out major)
            && int.TryParse(parts[1], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out minor);
    }

    private static string? Win32Path(RegistryKey version)
    {
        foreach (var lcid in version.GetSubKeyNames())
        {
            if (!int.TryParse(lcid, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)) continue;   // FLAGS, HELPDIR
            using var w = version.OpenSubKey(lcid + @"\win32");
            if (w?.GetValue(null) is string p && p.Length > 0) return Environment.ExpandEnvironmentVariables(p.Trim('"'));
        }
        return null;
    }
}
