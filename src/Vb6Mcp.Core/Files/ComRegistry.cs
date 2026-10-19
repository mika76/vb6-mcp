using Microsoft.Win32;

namespace Vb6Mcp.Core.Files;

/// <summary>A registered COM class (HKCR\CLSID\{clsid}, 32-bit view): its server file and typelib.</summary>
public sealed record ComClass(Guid Clsid, string? ProgId, string? InprocServer, string? LocalServer, string? ThreadingModel, Guid? TypeLib, string? TypeLibVersion)
{
    /// <summary>The server file (in-process first), environment variables expanded, quotes and arguments removed.</summary>
    public string? ServerFile => Clean(InprocServer) ?? Clean(LocalServer);

    public bool ServerExists => ServerFile is { } f && ComRegistry.Exists(f);

    private static string? Clean(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = Environment.ExpandEnvironmentVariables(s.Trim());
        if (s.StartsWith('"')) { int q = s.IndexOf('"', 1); return q > 0 ? s[1..q] : s.Trim('"'); }
        // LocalServer32 may carry arguments ("C:\x\App.exe /automation"); keep the part that ends in .exe
        int exe = s.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
        return exe > 0 && exe + 4 < s.Length ? s[..(exe + 4)] : s;
    }
}

/// <summary>
/// COM class registrations in the 32-bit registry view, where VB6 (32-bit) looks them up. Reading only;
/// (un)registration goes through the elevated add-in (hard rule 13, PLAN.md §9.2).
/// </summary>
public static class ComRegistry
{
    private static RegistryKey Root() => RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32);

    /// <summary>The class registered under a CLSID, or null.</summary>
    public static ComClass? Class(Guid clsid)
    {
        using var root = Root();
        using var k = root.OpenSubKey($@"CLSID\{clsid:B}");
        if (k is null) return null;
        string? Sub(string name) { using var s = k.OpenSubKey(name); return s?.GetValue(null) as string; }
        string? threading;
        using (var ip = k.OpenSubKey("InprocServer32")) threading = ip?.GetValue("ThreadingModel") as string;
        var tlb = Sub("TypeLib");
        return new ComClass(clsid, Sub("ProgID"), Sub("InprocServer32"), Sub("LocalServer32"), threading,
            Guid.TryParse(tlb, out var g) ? g : null, Sub("Version"));
    }

    /// <summary>
    /// Whether a registered file exists. Many system registrations use a bare name ("combase.dll") that Windows finds
    /// through the system folders: for those, the (32-bit) system directory, the Windows directory and PATH are searched,
    /// so they aren't reported as missing (com_orphans live test: 173 false hits before this).
    /// </summary>
    public static bool Exists(string file)
    {
        // a typelib path may end in "\<n>": the typelib resource index inside the file (MSVBVM60.DLL\3, vbscript.dll\2)
        var m = System.Text.RegularExpressions.Regex.Match(file, @"^(.+\.\w+)\\\d+$");
        if (m.Success && !File.Exists(file)) file = m.Groups[1].Value;
        if (Path.IsPathRooted(file)) return File.Exists(file);
        var dirs = new[] { Environment.SystemDirectory, Environment.GetFolderPath(Environment.SpecialFolder.Windows) }
            .Concat((Environment.GetEnvironmentVariable("PATH") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries));
        return dirs.Any(d => { try { return File.Exists(Path.Combine(d, file)); } catch (ArgumentException) { return false; } });
    }

    /// <summary>Every registered class (32-bit view). Thousands of keys: a second or two.</summary>
    public static IReadOnlyList<ComClass> AllClasses()
    {
        var list = new List<ComClass>();
        string[] names;
        using (var root = Root())
        using (var k = root.OpenSubKey("CLSID"))
            names = k?.GetSubKeyNames() ?? [];
        foreach (var n in names)
            if (Guid.TryParse(n, out var g) && Class(g) is { } c) list.Add(c);
        return list;
    }

    /// <summary>
    /// VB6's debugger stub. While an ActiveX project runs in the IDE, VB registers its classes to VB98\VB6DEBUG.DLL
    /// and its typelib to the .vbp; a normal End restores the previous registration, but a killed or crashed IDE
    /// leaves them behind (spike 24), and programs then can't create those classes.
    /// </summary>
    public static bool IsVbDebugServer(string? file) => file is not null && Path.GetFileName(file).Equals("VB6DEBUG.DLL", StringComparison.OrdinalIgnoreCase);

    /// <summary>A typelib registered to a project file (.vbp/.vbg): VB's registration while debugging (spike 24).</summary>
    public static bool IsProjectTypeLib(RegisteredTypeLib t)
        => t.Win32Path is { } p && (p.EndsWith(".vbp", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".vbg", StringComparison.OrdinalIgnoreCase));

    /// <summary>Machine type from a PE file's header: "x86", "x64", other hex, or null if unreadable.</summary>
    public static string? Bitness(string file)
    {
        try
        {
            using var fs = File.OpenRead(file);
            using var r = new BinaryReader(fs);
            if (r.ReadUInt16() != 0x5A4D) return null;            // MZ
            fs.Position = 0x3C;
            fs.Position = r.ReadInt32();
            if (r.ReadUInt32() != 0x4550) return null;            // PE\0\0
            return r.ReadUInt16() switch { 0x14C => "x86", 0x8664 => "x64", 0xAA64 => "arm64", var m => $"0x{m:X}" };
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or EndOfStreamException) { return null; }
    }

    /// <summary>The CLSID a ProgID (or version-independent ProgID through CurVer) points at, or null.</summary>
    public static Guid? ClsidOf(string progId)
    {
        using var root = Root();
        for (int hop = 0; hop < 3; hop++)
        {
            using var k = root.OpenSubKey(progId);
            if (k is null) return null;
            using (var c = k.OpenSubKey("CLSID"))
                if (c?.GetValue(null) is string s && Guid.TryParse(s, out var g)) return g;
            using var cur = k.OpenSubKey("CurVer");
            if (cur?.GetValue(null) is not string next || next.Equals(progId, StringComparison.OrdinalIgnoreCase)) return null;
            progId = next;
        }
        return null;
    }

    /// <summary>ProgIDs whose name contains <paramref name="text"/> (case-insensitive) and that have a CLSID, up to <paramref name="limit"/>.</summary>
    public static IReadOnlyList<(string ProgId, Guid Clsid)> FindProgIds(string text, int limit = 50)
        => FindProgIds(name => name.Contains(text, StringComparison.OrdinalIgnoreCase), limit);

    /// <summary>ProgIDs whose name <paramref name="match"/> accepts and that have a CLSID, up to <paramref name="limit"/>.</summary>
    public static IReadOnlyList<(string ProgId, Guid Clsid)> FindProgIds(Func<string, bool> match, int limit = 50)
    {
        var found = new List<(string, Guid)>();
        using var root = Root();
        foreach (var name in root.GetSubKeyNames())
        {
            if (name.StartsWith('.') || name.StartsWith('{') || !match(name)) continue;
            using var c = root.OpenSubKey(name + @"\CLSID");
            if (c?.GetValue(null) is string s && Guid.TryParse(s, out var g)) found.Add((name, g));
            if (found.Count >= limit) break;
        }
        return found;
    }
}
