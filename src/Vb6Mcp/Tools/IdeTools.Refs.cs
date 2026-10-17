using System.Text.Json.Nodes;
using Microsoft.Win32;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 5 reference tools (PLAN.md §9.2). References are read and changed live in the IDE (VBIDE
/// References). GUIDs come only from the registry or a typelib (hard rule 13): refs_add resolves a
/// name or a file to what's registered, and file paths must be under the workspace roots or be the
/// registered file of a typelib (hard rule 8). Components (OCXs, the .vbp's Object= lines) can be added
/// through VBIDE but not listed, so the list comes from the .vbp on disk.
/// </summary>
public sealed partial class IdeTools
{
    private IReadOnlyList<Tool> RefTools() =>
    [
        new("refs_list",
            "A project's references as the IDE has them now (name, LIBID, version, file, broken?), plus its ActiveX control components " +
            "(Object= lines; read from the .vbp on disk because VBIDE can't list them) resolved through the registry.",
            Schema.Object(("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => RefsListAsync(a)),
        new("refs_add",
            "Add a reference to a project in the IDE, given exactly one of: name (part of a registered typelib's description, e.g. " +
            "'ActiveX Data Objects 2.8'; several matches are listed instead of added), libid + version (as listed by refs_add or " +
            "com_search), or file (a DLL/OCX/TLB under the workspace roots or a registered typelib's file; a .vbp adds a project " +
            "reference inside a group). Design mode. Unsaved; save with project_save.",
            Schema.Object(("name", "string", "Text to find in registered typelib descriptions.", false),
                          ("libid", "string", "Typelib GUID, from a refs_add/com_search listing.", false),
                          ("version", "string", "maj.min with libid, e.g. 2.8.", false),
                          ("file", "string", "Full path of a typelib file or .vbp.", false),
                          ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => RefsAddAsync(a)),
        new("refs_remove",
            "Remove a reference (by its name as refs_list shows it) from a project in the IDE. Or name=#<position> (1-based, in refs_list order) for a broken reference whose name can't be read; VB removes a broken reference only while its typelib is registered (refs_fix_broken says how). Built-in references can't be removed. " +
            "Design mode. Unsaved; save with project_save.",
            Schema.Object(("name", "string", "Reference name, e.g. ADODB.", true), ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => RefsRemoveAsync(a)),
        new("refs_add_component",
            "Add an ActiveX control's component (OCX) to a project's toolbox, so form_add_control can place it: give the ProgID of one of " +
            "its controls (e.g. MSComctlLib.ListViewCtrl.2). The ProgID must be registered. VBIDE can't remove components again. " +
            "Design mode. Unsaved; save with project_save.",
            Schema.Object(("prog_id", "string", "ProgID of a control in the OCX.", true), ("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => RefsAddComponentAsync(a)),
        new("refs_fix_broken",
            "Broken (MISSING:) references of a project open in the IDE, each with what the .vbp asked for (LIBID, version, file hint), " +
            "what is registered for that LIBID, and the calls that fix it: refs_remove name=#<position> then refs_add with a registered " +
            "version or the file. Report only; it changes nothing itself.",
            Schema.Object(("project", "string", ProjectDesc, false), ("pid", "integer", PidDesc, false)),
            (a, _) => RefsFixBrokenAsync(a)),
    ];

    private async Task<ToolResult> RefsFixBrokenAsync(JsonObject a)
    {
        var (client, project) = await ResolveProjectAsync(a);
        var refs = await client.RefsAsync(project);
        var file = (await client.ProjectsAsync()).FirstOrDefault(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase))?.FileName;
        // A broken reference may not even answer Name/Guid (spike 23: empty fields), so the .vbp says what was asked for
        var broken = refs.Select((r, i) => (Ref: r, Position: i + 1)).Where(x => x.Ref.IsBroken || x.Ref.Guid.Length == 0).ToList();
        if (broken.Count == 0) return ToolResult.Ok(new { source = "ide", project, broken = Array.Empty<object>(), note = "no broken references" });
        var okIds = refs.Where(r => !r.IsBroken && Guid.TryParse(r.Guid, out _)).Select(r => Guid.Parse(r.Guid)).ToHashSet();
        var registered = TypeLibRegistry.All();
        var asked = new List<(ComDependency Dep, string? Hint, string Desc)>();
        if (file is { Length: > 0 } && File.Exists(file))
        {
            string dir = Path.GetDirectoryName(file)!;
            foreach (var line in VbpFile.Load(file).References)
            {
                if (ComDependency.Parse(line) is not { } d || okIds.Contains(d.LibId)) continue;
                var parts = line.Split('#');                   // *\G{guid}#maj.min#lcid#path#description
                string? hint = parts.Length > 3 && parts[3].Length > 0 ? PathKey.Normalize(parts[3], dir) : null;
                asked.Add((d, hint, parts.Length > 4 ? parts[4] : ""));
            }
        }
        return ToolResult.Ok(new
        {
            source = "ide", project,
            broken = broken.Select(b => new { position = b.Position, name = b.Ref.Name.Length > 0 ? b.Ref.Name : null, libId = b.Ref.Guid.Length > 0 ? b.Ref.Guid : null, version = b.Ref.Guid.Length > 0 ? $"{b.Ref.Major}.{b.Ref.Minor}" : null }),
            vbpAskedFor = asked.Select(x =>
            {
                var versions = registered.Where(t => t.LibId == x.Dep.LibId && !ComRegistry.IsProjectTypeLib(t)).OrderByDescending(t => t.Major).ThenByDescending(t => t.Minor).ToList();
                // usable = the registered file exists AND contains a type library (a .NET COM DLL registered without /tlb doesn't)
                var noTypeLib = versions.Where(t => t.FileExists && TypeLibInfo.Read(t.Win32Path!) is null).ToList();
                var usable = versions.Where(t => t.FileExists && !noTypeLib.Contains(t)).ToList();
                var sameMajor = usable.FirstOrDefault(t => t.Major == x.Dep.Major);
                var sameMajorNoTlb = noTypeLib.FirstOrDefault(t => t.Major == x.Dep.Major);
                string id = x.Dep.LibId.ToString("B").ToUpperInvariant();
                int? pos = broken.FirstOrDefault(b => Guid.TryParse(b.Ref.Guid, out var g) && g == x.Dep.LibId).Position is int p && p > 0 ? p : null;
                string at = pos is null ? "#<position>" : $"#{pos}";
                // A file carrying exactly this typelib: the hint path (under the roots), or a workspace build or compatibility copy
                var carriers = new List<string>();
                if (x.Hint is not null && File.Exists(x.Hint) && ProjectMap.IsUnderRoots(x.Hint, Roots)) carriers.Add(x.Hint);
                if (MapProvider() is { } map)
                    foreach (var f in map.Projects.SelectMany(pr => new[] { pr.OutputPath, pr.CompatibleFile }).OfType<string>().Distinct(StringComparer.OrdinalIgnoreCase))
                        if (File.Exists(f) && !carriers.Contains(f, StringComparer.OrdinalIgnoreCase) && TypeLibInfo.Read(f) is { } tl && tl.LibId == x.Dep.LibId && tl.Major == x.Dep.Major)
                            carriers.Add(f);
                // VBIDE removes a broken reference only while its typelib is registered (spike 25: "Object library not registered" otherwise)
                string fix = sameMajor is not null
                        ? $"refs_remove name={at}, then refs_add libid={id} version={sameMajor.Major}.{sameMajor.Minor}"
                    : sameMajorNoTlb is not null
                        ? $"{id} {sameMajorNoTlb.Major}.{sameMajorNoTlb.Minor} is registered to {sameMajorNoTlb.Win32Path}, but that file contains no type library, so VB6 can't load it. Register one (for a .NET COM DLL, elevated: RegAsm \"{sameMajorNoTlb.Win32Path}\" /tlb) or provide the .tlb at the .vbp's hint path; VB then resolves the reference by itself"
                    : carriers.Count > 0
                        ? $"com_register file=\"{carriers[0]}\" confirm=true (VB re-resolves the reference as soon as it's registered, spike 25); then refs_fix_broken again: if it's still broken, refs_remove name={at} and refs_add libid={id} version={x.Dep.Major}.<minor as registered>"
                    : usable.Count > 0
                        ? $"only another major version is registered ({usable[0].Major}.{usable[0].Minor}): refs_add libid={id} version={usable[0].Major}.{usable[0].Minor} adds it, but the broken entry can't be removed through VBIDE until version {x.Dep.Major}.x is registered (or the user unchecks it in Project > References)"
                        : $"nothing carries {id} {x.Dep.Major}.x here: VBIDE can't remove the broken entry; install or build the component, or the user unchecks it in Project > References";
                return new
                {
                    position = pos, filesWithThisTypeLib = carriers,
                    description = x.Desc, libId = x.Dep.LibId.ToString("B").ToUpperInvariant(), major = x.Dep.Major, fileHint = x.Hint, fileHintExists = x.Hint is not null && File.Exists(x.Hint),
                    registered = versions.Select(t => new { version = $"{t.Major}.{t.Minor}", file = t.Win32Path, fileExists = t.FileExists, hasTypeLib = t.FileExists && !noTypeLib.Contains(t) }),
                    fix,
                };
            }),
            steps = "follow each fix; positions shift after a removal, so work from the highest position down; save with project_save",
        });
    }

    private async Task<ToolResult> RefsListAsync(JsonObject a)
    {
        var (client, project) = await ResolveProjectAsync(a);
        var refs = await client.RefsAsync(project);
        var file = (await client.ProjectsAsync()).FirstOrDefault(p => p.Name.Equals(project, StringComparison.OrdinalIgnoreCase))?.FileName;
        var registered = TypeLibRegistry.All();
        object? components = null;
        if (file is { Length: > 0 } && File.Exists(file))
        {
            components = VbpFile.Load(file).Objects.Select(o =>
            {
                var d = ComDependency.Parse(o);
                var r = d is null ? null : registered.Where(t => t.LibId == d.LibId && t.Major == d.Major).OrderByDescending(t => t.Minor).FirstOrDefault();
                return new
                {
                    line = o, libId = d?.LibId.ToString("B").ToUpperInvariant(), major = d?.Major,
                    registered = r is null ? null : new { version = $"{r.Major}.{r.Minor}", description = r.Description, file = r.Win32Path, fileExists = r.FileExists },
                };
            }).ToList();
        }
        return ToolResult.Ok(new
        {
            source = "ide", project,
            references = refs.Select(r => new
            {
                r.Name, libId = r.Guid, version = $"{r.Major}.{r.Minor}", file = r.FullPath, builtIn = r.BuiltIn, broken = r.IsBroken,
                kind = r.Type == 1 ? "project" : "typelib", description = r.Description,
            }),
            components,
            componentsNote = file is { Length: > 0 } ? "components come from the .vbp on disk; ones added since the last save aren't listed" : "the project has never been saved, so its components can't be listed",
        });
    }

    private async Task<ToolResult> RefsAddAsync(JsonObject a)
    {
        string? name = a.Str("name"), libid = a.Str("libid"), file = a.Str("file");
        if (new[] { name, libid, file }.Count(s => !string.IsNullOrEmpty(s)) != 1)
            throw new ToolException("give exactly one of name, libid (+ version) or file");
        var inst = ResolveInstance(a);
        RequireDesignMode(inst);
        var (client, project) = await ResolveProjectAsync(a);
        var registered = TypeLibRegistry.All();
        RefInfo added;
        try
        {
            if (!string.IsNullOrEmpty(file))
            {
                var path = PathKey.Normalize(file);
                bool underRoots = Roots.Count > 0 && ProjectMap.IsUnderRoots(path, Roots);
                bool isRegistered = registered.Any(t => t.Win32Path is { } p && PathKey.Same(p, path));
                if (!underRoots && !isRegistered)
                    throw new ToolException($"{path} is neither under the workspace roots nor the file of a registered typelib (refs_add by name or libid adds registered ones)");
                if (!File.Exists(path)) throw new ToolException("file not found: " + path);
                added = await client.RefAddFileAsync(project, path);
            }
            else if (!string.IsNullOrEmpty(libid))
            {
                if (!Guid.TryParse(libid, out var g)) throw new ToolException("libid isn't a GUID");
                if (a.Str("version") is not { } ver || !TryDecimalVersion(ver, out int ma, out int mi)) throw new ToolException("version (maj.min, e.g. 2.8) is required with libid");
                var r = registered.FirstOrDefault(t => t.LibId == g && t.Major == ma && t.Minor == mi)
                        ?? throw new ToolException($"{g.ToString("B").ToUpperInvariant()} {ver} isn't registered; registered versions: {string.Join(", ", registered.Where(t => t.LibId == g).Select(t => $"{t.Major}.{t.Minor}").DefaultIfEmpty("none"))}");
                added = await client.RefAddGuidAsync(project, r.LibId, r.Major, r.Minor);
            }
            else
            {
                var matches = registered.Where(t => t.Description.Contains(name!, StringComparison.OrdinalIgnoreCase) && t.FileExists)
                    .GroupBy(t => t.LibId).Select(grp => grp.OrderByDescending(t => t.Major).ThenByDescending(t => t.Minor).First())
                    .OrderBy(t => t.Description, StringComparer.OrdinalIgnoreCase).ToList();
                if (matches.Count == 0) throw new ToolException($"no registered typelib (with its file present) has '{name}' in its description");
                if (matches.Count > 1)
                    return ToolResult.Ok(new
                    {
                        added = (object?)null, note = "several typelibs match; call refs_add again with libid and version",
                        candidates = matches.Take(30).Select(t => new { libId = t.LibId.ToString("B").ToUpperInvariant(), version = $"{t.Major}.{t.Minor}", description = t.Description, file = t.Win32Path }),
                    });
                added = await client.RefAddGuidAsync(project, matches[0].LibId, matches[0].Major, matches[0].Minor);
            }
        }
        catch (AddinErrorException ex) { throw new ToolException("VB6 refused the reference: " + ex.AddinMessage); }
        return ToolResult.Ok(new
        {
            source = "ide", project,
            added = new { added.Name, libId = added.Guid, version = $"{added.Major}.{added.Minor}", file = added.FullPath, broken = added.IsBroken, kind = added.Type == 1 ? "project" : "typelib", description = added.Description },
            note = "unsaved in the IDE; save with project_save",
        });
    }

    /// <summary>Versions as people write them (2.8, decimal); the registry's keys are hex and parsed in TypeLibRegistry.</summary>
    private static bool TryDecimalVersion(string s, out int major, out int minor)
    {
        major = minor = 0;
        var p = s.Split('.');
        return p.Length == 2 && int.TryParse(p[0], out major) && int.TryParse(p[1], out minor);
    }

    private async Task<ToolResult> RefsRemoveAsync(JsonObject a)
    {
        RequireDesignMode(ResolveInstance(a));
        var (client, project) = await ResolveProjectAsync(a);
        var name = a.Req("name");
        try { await client.RefRemoveAsync(project, name); }
        catch (AddinErrorException ex) { throw new ToolException(ex.AddinMessage); }
        return ToolResult.Ok(new { source = "ide", project, removed = name, note = "unsaved in the IDE; save with project_save" });
    }

    private async Task<ToolResult> RefsAddComponentAsync(JsonObject a)
    {
        RequireDesignMode(ResolveInstance(a));
        var (client, project) = await ResolveProjectAsync(a);
        var progId = a.Req("prog_id");
        // The ProgID must resolve (32-bit view) to a CLSID with a server file: that's what VB loads
        using var hkcr = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32);
        var clsid = hkcr.OpenSubKey(progId + @"\CLSID")?.GetValue(null) as string
                    ?? throw new ToolException($"ProgID {progId} isn't registered (32-bit view)");
        var server = hkcr.OpenSubKey($@"CLSID\{clsid}\InprocServer32")?.GetValue(null) as string;
        try { await client.ToolboxAddAsync(project, progId, null); }
        catch (AddinErrorException ex) { throw new ToolException("VB6 refused the component: " + ex.AddinMessage); }
        return ToolResult.Ok(new { source = "ide", project, added = progId, clsid, file = server, note = "unsaved in the IDE; save with project_save" });
    }
}
