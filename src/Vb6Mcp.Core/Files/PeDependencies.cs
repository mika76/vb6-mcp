using System.Text;

namespace Vb6Mcp.Core.Files;

/// <summary>A DLL a file needs that Windows wouldn't find.</summary>
/// <param name="Chain">From the checked file to the missing DLL, e.g. DPFPDevX.dll, DPFPApi.dll, DPSCApi.dll.</param>
/// <param name="Delay">A link of the chain is a delay import (loaded on first use: the program fails only then, with 0xc06d007e).</param>
/// <param name="Note">Null when the DLL is nowhere on the search path; otherwise why it still won't be found.</param>
public sealed record MissingDll(IReadOnlyList<string> Chain, bool Delay, string? Note);

/// <summary>
/// Which DLLs a 32-bit DLL/OCX needs that the loader wouldn't find (spike 39: DPFPDevX.dll delay-loads DPFPApi.dll,
/// whose DPSCApi.dll isn't installed; the IDE died with 0xc06d007e and the Application log named only KERNELBASE).
/// Reads the import and delay-import tables (no code of the file runs: LoadLibrary would run its DllMain in the
/// server) and resolves names the way the loader does for a COM server, which COM loads with
/// LOAD_WITH_ALTERED_SEARCH_PATH: its own folder, then the system folders and PATH. A delay import is loaded later
/// with a plain LoadLibrary, which looks in the program's folder instead of the DLL's.
/// Not covered: manifests other than the VC 8/9 runtimes' WinSxS lookup, DLLs loaded by name at run time.
/// </summary>
public static class PeDependencies
{
    /// <summary>The import and delay-import DLL names of a PE file, or null when it isn't one.</summary>
    public static (IReadOnlyList<string> Imports, IReadOnlyList<string> Delay)? Read(string path)
    {
        byte[] b;
        try { b = File.ReadAllBytes(path); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return null; }
        try { return Parse(b); } catch (Exception ex) when (ex is ArgumentException or IndexOutOfRangeException) { return null; }
    }

    private static (IReadOnlyList<string>, IReadOnlyList<string>)? Parse(byte[] b)
    {
        if (b.Length < 0x40 || b[0] != 'M' || b[1] != 'Z') return null;
        int pe = I32(b, 0x3c);
        if (pe <= 0 || pe + 24 > b.Length || I32(b, pe) != 0x4550) return null;   // "PE\0\0"
        int sections = U16(b, pe + 6), optSize = U16(b, pe + 20), opt = pe + 24;
        bool pe32 = U16(b, opt) == 0x10b;
        long imageBase = pe32 ? (uint)I32(b, opt + 28) : BitConverter.ToInt64(b, opt + 24);
        int dirs = opt + (pe32 ? 96 : 112);
        var secs = new List<(int Va, int Size, int Raw)>();
        for (int i = 0, s = opt + optSize; i < sections; i++, s += 40)
            secs.Add((I32(b, s + 12), Math.Max(I32(b, s + 8), I32(b, s + 16)), I32(b, s + 20)));
        int Off(long rva)
        {
            foreach (var (va, size, raw) in secs) if (rva >= va && rva < va + size) return (int)(rva - va + raw);
            return -1;
        }
        string Str(long rva)
        {
            int o = Off(rva), e = o;
            while (o >= 0 && e < b.Length && b[e] != 0) e++;
            return o < 0 ? "" : Encoding.Latin1.GetString(b, o, e - o);
        }
        List<string> Table(int dirIndex, int step, int nameAt)
        {
            var names = new List<string>();
            int rva = I32(b, dirs + dirIndex * 8);
            if (rva == 0) return names;
            for (int o = Off(rva); o >= 0 && o + step <= b.Length; o += step)
            {
                long name = (uint)I32(b, o + nameAt);
                if (name == 0) break;
                if (dirIndex == 13 && (I32(b, o) & 1) == 0) name -= imageBase;   // old delay descriptors hold VAs
                if (Str(name) is { Length: > 0 } n) names.Add(n);
            }
            return names;
        }
        return (Table(1, 20, 12), Table(13, 32, 4));
    }

    private static int I32(byte[] b, int o) => BitConverter.ToInt32(b, o);
    private static int U16(byte[] b, int o) => BitConverter.ToUInt16(b, o);

    /// <summary>
    /// Walks the dependencies of <paramref name="file"/> (and of every DLL found outside the Windows folder, a few levels
    /// deep) and returns the ones that wouldn't load. <paramref name="appFolders"/> are the folders of the programs that
    /// load it (VB98 for the IDE), which a delay import searches instead of the DLL's folder.
    /// </summary>
    public static IReadOnlyList<MissingDll> Missing(string file, IReadOnlyList<string>? appFolders = null, int maxDepth = 6)
    {
        var missing = new List<MissingDll>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var apps = appFolders ?? [];
        void Walk(string path, List<string> chain, bool viaDelay)
        {
            if (!seen.Add(path) || chain.Count > maxDepth || Read(path) is not { } t) return;
            var dir = Path.GetDirectoryName(path)!;
            foreach (var (name, delay) in t.Imports.Select(n => (n, false)).Concat(t.Delay.Select(n => (n, true))))
            {
                if (IsApiSet(name)) continue;
                var next = new List<string>(chain) { name };
                var found = Find(name, delay ? apps : [dir]);
                if (found is null)
                {
                    // a delay import that sits only next to the DLL: the program's LoadLibrary won't look there
                    string? note = delay && File.Exists(Path.Combine(dir, name)) ? $"it is next to {Path.GetFileName(path)}, but a delay import is loaded later from the program's folder and the system folders, not from there" : null;
                    missing.Add(new MissingDll(next, viaDelay || delay, note));
                    continue;
                }
                // Windows' own DLLs load (and walking them is slow); third-party ones in SysWOW64 are walked too
                // (spike 39: DPFPApi.dll in SysWOW64 was the one with the missing dependency)
                if (!IsWindowsOwn(found)) Walk(found, next, viaDelay || delay);
            }
        }
        Walk(Path.GetFullPath(file), [Path.GetFileName(file)], false);
        return missing;
    }

    private static readonly string WindowsDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows) + "\\";

    private static bool IsWindowsOwn(string path)
    {
        if (!path.StartsWith(WindowsDir, StringComparison.OrdinalIgnoreCase)) return false;
        try { return System.Diagnostics.FileVersionInfo.GetVersionInfo(path).CompanyName?.Contains("Microsoft", StringComparison.OrdinalIgnoreCase) == true; }
        catch (FileNotFoundException) { return false; }
    }

    // api-ms-win-* / ext-ms-* are API sets: the loader maps them to system DLLs, there is no such file
    private static bool IsApiSet(string name) =>
        name.StartsWith("api-ms-", StringComparison.OrdinalIgnoreCase) || name.StartsWith("ext-ms-", StringComparison.OrdinalIgnoreCase);

    /// <summary>Where the loader of a 32-bit process would find <paramref name="name"/>, or null.</summary>
    private static string? Find(string name, IReadOnlyList<string> first)
    {
        var win = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        // a 32-bit process's System32 is SysWOW64 (the server is x86, but say it explicitly)
        var system = Directory.Exists(Path.Combine(win, "SysWOW64")) ? Path.Combine(win, "SysWOW64") : Path.Combine(win, "System32");
        var path = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var d in first.Concat([system, Path.Combine(win, "System"), win]).Concat(path))
        {
            try { var p = Path.Combine(d, name); if (File.Exists(p)) return p; }
            catch (ArgumentException) { }   // a malformed PATH entry
        }
        return SideBySide(name, win);
    }

    // The VC 8/9 runtimes (msvcr80.dll, mfc90u.dll, ...) load from WinSxS through the DLL's manifest
    private static string? SideBySide(string name, string win)
    {
        if (!System.Text.RegularExpressions.Regex.IsMatch(name, @"^(msvc[rpm]|mfc|atl)(80|90)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) return null;
        var sxs = Path.Combine(win, "WinSxS");
        try
        {
            return Directory.EnumerateDirectories(sxs, "x86_microsoft.vc*").Select(d => Path.Combine(d, name)).FirstOrDefault(File.Exists);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return null; }
    }
}
