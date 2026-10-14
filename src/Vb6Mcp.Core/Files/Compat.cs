namespace Vb6Mcp.Core.Files;

/// <summary>One finding of <see cref="Compat.Check"/>: "error" or "warning", and the text.</summary>
public sealed record CompatFinding(string Level, string Text);

/// <summary>The compatibility state of one project (compat_status).</summary>
public sealed record CompatReport(
    VbpFile Project, string Mode, string? CompatibleFile, bool CompatibleFileExists, TypeLibInfo? CompatibleTypeLib,
    string? Output, bool OutputExists, TypeLibInfo? OutputTypeLib, IReadOnlyList<RegisteredTypeLib> Registered, IReadOnlyList<CompatFinding> Findings);

/// <summary>
/// Binary compatibility checks for ActiveX projects (PLAN.md §9.2.1). Only facts read from the .vbp,
/// the files' typelibs (LIBID + version) and the registry; nothing is changed. Never decides about
/// breaking compatibility (hard rule 13): it reports.
/// </summary>
public static class Compat
{
    /// <summary>Project types that build a COM server (VB writes these in .vbp Type=).</summary>
    public static bool IsComServer(VbpFile p) => p.Type is "OleDll" or "OleExe" or "Control";

    private static string U(Guid g) => g.ToString("B").ToUpperInvariant();

    public static string ModeName(int mode) => mode switch { 0 => "none", 1 => "project", 2 => "binary", _ => mode.ToString() };

    /// <param name="p">The project.</param>
    /// <param name="typeLib">Typelib reader (TypeLibInfo.Read; injected for tests).</param>
    /// <param name="registered">Registered typelib versions (TypeLibRegistry.All; injected for tests).</param>
    public static CompatReport Check(VbpFile p, Func<string, TypeLibInfo?> typeLib, IReadOnlyList<RegisteredTypeLib> registered)
    {
        var f = new List<CompatFinding>();
        bool compatExists = p.CompatibleFile is { } cf && File.Exists(cf);
        bool outExists = p.OutputPath is { } of && File.Exists(of);
        var compatLib = compatExists ? typeLib(p.CompatibleFile!) : null;
        var outLib = outExists ? typeLib(p.OutputPath!) : null;

        if (!IsComServer(p))
        {
            if (p.CompatibleMode != 0) f.Add(new("warning", $"a {p.Type} project with {ModeName(p.CompatibleMode)} compatibility: compatibility only matters for ActiveX DLL/EXE/OCX projects"));
        }
        else
        {
            if (p.CompatibleMode != 2)
                f.Add(new("warning", $"{ModeName(p.CompatibleMode)} compatibility, not binary: the GUIDs of the next build aren't tied to an existing binary, so compiled programs using {p.Name} can stop working after a rebuild"));
            if (p.CompatibleMode == 2 && p.CompatibleFile is null)
                f.Add(new("error", "binary compatibility is on but no compatibility file is set (CompatibleEXE32)"));
            if (p.CompatibleFile is not null && !compatExists)
                f.Add(new("error", $"the compatibility file doesn't exist: {p.CompatibleFile}"));
            if (compatExists && compatLib is null)
                f.Add(new("error", $"the compatibility file has no readable typelib: {p.CompatibleFile}"));
            if (p.CompatibleFile is not null && p.OutputPath is not null && PathKey.Same(p.CompatibleFile, p.OutputPath))
                f.Add(new("warning", "the compatibility file is the build output itself, so every build replaces the reference it is checked against; keep a separate reference copy"));
            if (compatLib is not null && outLib is not null)
            {
                if (compatLib.LibId != outLib.LibId)
                    f.Add(new("error", $"the build output's typelib {U(outLib.LibId)} differs from the compatibility file's {U(compatLib.LibId)}: the output wasn't built compatible with it (compatibility broken, or built without it)"));
                else if (compatLib.Major != outLib.Major)
                    f.Add(new("warning", $"the output's typelib version {outLib.Major}.{outLib.Minor} has another major version than the compatibility file's {compatLib.Major}.{compatLib.Minor}"));
            }
            if (!outExists) f.Add(new("warning", $"no build output yet ({p.OutputPath ?? "no ExeName32 set"})"));
        }

        // What's registered for the output's LIBID, and whether it points at this output
        var lib = outLib ?? compatLib;
        var reg = lib is null ? [] : registered.Where(r => r.LibId == lib.LibId).OrderBy(r => r.Major).ThenBy(r => r.Minor).ToList();
        if (lib is not null && IsComServer(p))
        {
            if (reg.Count == 0) f.Add(new("warning", $"typelib {U(lib.LibId)} isn't registered on this machine"));
            else if (outExists && !reg.Any(r => r.Win32Path is { } w && PathKey.Same(w, p.OutputPath!)))
                f.Add(new("warning", $"typelib {U(lib.LibId)} is registered, but not for the build output: {string.Join("; ", reg.Select(r => $"{r.Major}.{r.Minor} -> {r.Win32Path ?? "(no file)"}{(r.FileExists ? "" : " (missing)")}"))}"));
        }
        return new CompatReport(p, ModeName(p.CompatibleMode), p.CompatibleFile, compatExists, compatLib, p.OutputPath, outExists, outLib, reg, f);
    }
}
