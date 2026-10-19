namespace Vb6Mcp.Core.Files;

/// <summary>A control or reference of a .vbp that isn't registered on this machine (by typelib GUID + major version).</summary>
public sealed record MissingDependency(string Kind, string Name, Guid LibId, int Major, IReadOnlyList<string> OtherVersions, string Raw);

/// <summary>
/// What project_open checks before asking VB6 to load a project. A production project (improvement notes 8 and 17): two controls
/// weren't registered, VB6 asked "'cappro.dll' could not be loaded--Continue Loading Project?", the prompt resolved as No
/// through AddFromFile, and the IDE ended up with an empty project and no message. A missing control (Object= line)
/// therefore blocks the open unless forced. A missing reference only warns: VB6 loads the project and marks it MISSING.
/// Resolution follows VB: GUID + major version through HKCR\TypeLib (32-bit view); the hint path is only a fallback
/// (spike 25), so it isn't consulted here.
/// </summary>
public static class ProjectPreflight
{
    public static IReadOnlyList<MissingDependency> Missing(VbpFile vbp, IReadOnlyList<RegisteredTypeLib> registered)
    {
        var missing = new List<MissingDependency>();
        void Check(string raw, string kind)
        {
            if (ComDependency.Parse(raw) is not { } d) return;   // project references, unknown forms
            var sameLib = registered.Where(t => t.LibId == d.LibId).ToList();
            if (sameLib.Any(t => t.Major == d.Major)) return;
            missing.Add(new MissingDependency(kind, NameOf(raw, kind), d.LibId, d.Major,
                sameLib.Select(t => $"{t.Major}.{t.Minor}").Distinct().ToList(), raw));
        }
        foreach (var o in vbp.Objects) Check(o, "control");
        foreach (var r in vbp.References) Check(r, "reference");
        return missing;
    }

    /// <summary>
    /// The registered DLL/OCX files behind a .vbp's controls and references, by display name (not .tlb/.olb: those
    /// hold no code; not the project typelibs the IDE's debugger registers).
    /// </summary>
    public static IReadOnlyList<(string Name, string Path)> Binaries(VbpFile vbp, IReadOnlyList<RegisteredTypeLib> registered)
    {
        var list = new List<(string, string)>();
        void Add(string raw, string kind)
        {
            if (ComDependency.Parse(raw) is not { } d) return;
            var path = registered.Where(t => t.LibId == d.LibId && t.Major == d.Major && !ComRegistry.IsProjectTypeLib(t))
                                 .Select(t => t.Win32Path).OfType<string>().Where(File.Exists).FirstOrDefault();
            if (path is null || path.EndsWith(".tlb", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".olb", StringComparison.OrdinalIgnoreCase)) return;
            list.Add((NameOf(raw, kind), path));
        }
        foreach (var o in vbp.Objects) Add(o, "control");
        foreach (var r in vbp.References) Add(r, "reference");
        return list;
    }

    /// <summary>
    /// The DLLs that the 32-bit controls and references of a .vbp need and Windows wouldn't find (spike 39: a missing
    /// delay-loaded DLL ended the IDE with 0xc06d007e while the program ran, unsaved work included). run_start reports
    /// them; com_diagnose project=... does the same check with every other one.
    /// </summary>
    public static IReadOnlyList<(string Name, MissingDll Missing)> MissingDlls(VbpFile vbp, IReadOnlyList<RegisteredTypeLib> registered, IReadOnlyList<string> appFolders)
    {
        var list = new List<(string, MissingDll)>();
        foreach (var (name, path) in Binaries(vbp, registered))
        {
            if (ComRegistry.Bitness(path) is not "x86") continue;   // a 64-bit server can't load at all: com_diagnose says so
            foreach (var m in PeDependencies.Missing(path, appFolders)) list.Add((name, m));
        }
        return list;
    }

    // Object: {guid}#2.0#0; cappro.dll        Reference: *\G{guid}#1.0#0#C:\x\lib.dll#Description
    private static string NameOf(string raw, string kind)
        => kind == "control" ? (raw.Contains(';') ? raw[(raw.IndexOf(';') + 1)..].Trim() : raw)
                             : raw.Split('#') is { Length: >= 5 } p ? (p[4].Length > 0 ? p[4] : Path.GetFileName(p[3])) : raw;

    /// <summary>How many members of the .vbp are source components VB loads (Form, Module, Class, ...; not RelatedDoc).</summary>
    public static int ComponentCount(VbpFile vbp) => vbp.Members.Count(m => !m.Kind.Equals("RelatedDoc", StringComparison.OrdinalIgnoreCase));
}
