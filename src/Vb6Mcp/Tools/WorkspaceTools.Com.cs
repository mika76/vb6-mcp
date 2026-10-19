using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>
/// Phase 5 COM tools (PLAN.md §9.2): what's inside a component (its typelib), what's registered where,
/// binary compatibility state, and (un)registration. Reads run in the server (registry 32-bit view, our
/// own typelib reader); registration runs regsvr32 elevated through the add-in, is audited, and is
/// checked afterwards in the registry rather than trusted from an exit code (hard rule 13).
/// </summary>
public sealed partial class WorkspaceTools
{
    private const int DescribeLimit = 60_000;

    private IReadOnlyList<Tool> ComTools() =>
    [
        new("com_describe",
            "What a COM component offers: the classes, interfaces, methods, properties, events and enums in its type library, as VB " +
            "declarations (read from the file, nothing is registered). Give exactly one of file (DLL/OCX/TLB/OLB under the roots or a " +
            "registered one), prog_id, or libid + version. filter keeps types whose name contains it; large libraries without a filter " +
            "return the type list only.",
            Schema.Object(("file", "string", "Full path of the DLL/OCX/TLB.", false), ("prog_id", "string", "A registered ProgID, e.g. ADODB.Recordset.", false),
                          ("libid", "string", "Typelib GUID (with version).", false), ("version", "string", "maj.min with libid.", false),
                          ("filter", "string", "Only types whose name contains this.", false), ("include_hidden", "boolean", "Also hidden types and members.", false)),
            (a, _) => Task.FromResult(ComDescribe(a))),
        new("com_search",
            "Find registered COM components: type libraries whose description or name contains the text, and ProgIDs containing it " +
            "(32-bit registry, as VB6 sees it). Each with its file and whether that file exists.",
            Schema.Object(("text", "string", "Text to find, e.g. 'ADO', 'Scripting', a company prefix.", true)),
            (a, _) => Task.FromResult(ComSearch(a))),
        new("com_where",
            "Where a component is registered: give one of prog_id, clsid, libid or file. Lists the registered versions and server files, " +
            "whether each file exists and when it was written, and which workspace project builds it. For a file: which of its classes " +
            "and its typelib are registered, and to which file.",
            Schema.Object(("prog_id", "string", "ProgID.", false), ("clsid", "string", "Class GUID.", false), ("libid", "string", "Typelib GUID.", false),
                          ("file", "string", "Full path of a DLL/OCX/EXE.", false)),
            (a, _) => Task.FromResult(ComWhere(a))),
        new("compat_status",
            "Binary compatibility state of ActiveX projects under the roots: mode, compatibility file, the typelib IDs of the compatibility " +
            "file and the build output (they must match for binary compatibility), what's registered, and warnings (no binary compatibility, " +
            "missing reference file, reference = output, broken compatibility, registered elsewhere). project: a .vbp path or a name filter; " +
            "default: every ActiveX project.",
            Schema.Object(("project", "string", ".vbp path or project name filter.", false), ("refresh", "boolean", "Rescan the roots first.", false)),
            (a, _) => Task.FromResult(CompatStatus(a))),
        new("com_register",
            "Register a DLL/OCX under the workspace roots (regsvr32, elevated through the VB6 add-in, audited). Then checks the registry: " +
            "its typelib and classes must point at this file. Needs confirm=true.",
            Schema.Object(("file", "string", "Full path of the DLL/OCX.", true), ("confirm", "boolean", "Required: true.", true), ("pid", "integer", "IDE whose add-in runs it (default: the only one).", false)),
            (a, _) => RegisterAsync(a, unregister: false)),
        new("com_unregister",
            "Unregister a DLL/OCX under the workspace roots (regsvr32 /u, elevated through the VB6 add-in, audited), then check the " +
            "registry. Refused if its classes are registered to a different file (unregistering would remove those registrations too) " +
            "unless force=true. Needs confirm=true.",
            Schema.Object(("file", "string", "Full path of the DLL/OCX.", true), ("confirm", "boolean", "Required: true.", true),
                          ("force", "boolean", "Also when its classes point at another copy.", false), ("pid", "integer", "IDE whose add-in runs it.", false)),
            (a, _) => RegisterAsync(a, unregister: true)),
        new("com_orphans",
            "Broken COM registrations, report only (nothing is changed): classes whose server file is missing, typelibs whose file is " +
            "missing, and VB6 debugger registrations left behind by a killed or crashed IDE (classes pointing at VB6DEBUG.DLL, typelibs " +
            "pointing at a .vbp) with the call that repairs them. Optional text filter and under_roots (only files under the workspace roots).",
            Schema.Object(("text", "string", "Only entries whose ProgID, description or file contains this.", false),
                          ("under_roots", "boolean", "Only registrations of files under the workspace roots.", false)),
            (a, _) => ComOrphansAsync(a)),
        new("com_diagnose",
            "Checks behind a COM runtime error (429 can't create object, 430 class doesn't support the expected interface, 48 error " +
            "loading DLL, or any other) for a target: prog_id, clsid, file (DLL/OCX), or project (.vbp: all its references and " +
            "components). Each check says ok/warning/error with details: registered?, file present?, 32-bit?, typelib version, VB6 " +
            "debugger leftovers, binary compatibility of a workspace build.",
            Schema.Object(("error", "integer", "VB runtime error number, e.g. 429.", true), ("prog_id", "string", "ProgID the program creates.", false),
                          ("clsid", "string", "Class GUID.", false), ("file", "string", "Component file.", false), ("project", "string", ".vbp of the program.", false)),
            (a, _) => ComDiagnoseAsync(a)),
    ];

    // ---------- orphans and debug leftovers ----------

    /// <summary>.vbp files open in an IDE that is in run or break mode right now (their debug registration is live, not left over).</summary>
    private async Task<HashSet<string>> DebuggingProjectsAsync()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var inst in env.Scan().Live)
        {
            if (env.ReadWindows(inst.Pid).Mode is not (IdeMode.Run or IdeMode.Break)) continue;
            try { foreach (var p in await env.Client(inst).ProjectsAsync()) if (p.FileName.Length > 0) set.Add(PathKey.Normalize(p.FileName)); }
            catch (Exception ex) when (ex is AddinBusyException or AddinUnreachableException or AddinErrorException) { }
        }
        return set;
    }

    /// <summary>VB6 debugger registrations: per .vbp, whether it's live or left over, its classes, and the repair.</summary>
    private object DebugLeftovers(IReadOnlyList<ComClass> classes, IReadOnlyList<RegisteredTypeLib> libs, HashSet<string> debugging)
    {
        var debugClasses = classes.Where(c => ComRegistry.IsVbDebugServer(c.ServerFile)).ToList();
        var projectLibs = libs.Where(ComRegistry.IsProjectTypeLib).ToList();
        var byProject = projectLibs.Select(t =>
        {
            var vbp = PathKey.Normalize(t.Win32Path!);
            bool live = debugging.Contains(vbp);
            string? output = null;
            try { if (File.Exists(vbp)) output = VbpFile.Load(vbp).OutputPath; } catch (IOException) { }
            bool outputOk = output is not null && File.Exists(output);
            // Without its own build: another workspace build carrying the same typelib (e.g. the compatibility copy's project)
            if (!outputOk && roots.Count > 0)
            {
                var other = Map(false).Projects.Select(p => p.OutputPath).OfType<string>()
                    .FirstOrDefault(o => File.Exists(o) && TypeLibInfo.Read(o) is { } tl && tl.LibId == t.LibId && tl.Major == t.Major);
                if (other is not null) { output = other; outputOk = true; }
            }
            return new
            {
                project = vbp, state = live ? "debugging now (the IDE restores the registration when the run ends)" : "left over (killed or crashed IDE)",
                libId = t.LibId.ToString("B").ToUpperInvariant(), version = $"{t.Major}.{t.Minor}",
                classes = debugClasses.Where(c => c.TypeLib == t.LibId).Select(c => new { clsid = c.Clsid.ToString("B").ToUpperInvariant(), progId = c.ProgId }),
                repair = live ? null
                       : outputOk ? $"com_register file=\"{output}\" confirm=true (a build with this typelib; registering it again repairs the classes)"
                       : "build the project (build_make), which registers the new build; there is no build output to register now",
            };
        }).ToList();
        var stray = debugClasses.Where(c => c.TypeLib is null || !projectLibs.Any(t => t.LibId == c.TypeLib))
            .Select(c => new { clsid = c.Clsid.ToString("B").ToUpperInvariant(), progId = c.ProgId }).ToList();
        return new { projects = byProject, classesWithoutProject = stray };
    }

    private async Task<ToolResult> ComOrphansAsync(JsonObject a)
    {
        var text = a.Str("text");
        bool underRoots = a.Bool("under_roots") == true;
        bool Keep(params string?[] fields)
            => (text is null || fields.Any(f => f?.Contains(text, StringComparison.OrdinalIgnoreCase) == true))
               && (!underRoots || fields.Any(f => f is { Length: > 2 } && f[1] == ':' && ProjectMap.IsUnderRoots(f, roots)));
        var classes = ComRegistry.AllClasses();
        var libs = TypeLibRegistry.All();
        var missingClasses = classes.Where(c => c.ServerFile is not null && !c.ServerExists && !ComRegistry.IsVbDebugServer(c.ServerFile) && Keep(c.ProgId, c.ServerFile)).ToList();
        var missingLibs = libs.Where(t => t.Win32Path is not null && !ComRegistry.IsProjectTypeLib(t) && !t.FileExists && Keep(t.Description, t.Win32Path)).ToList();
        const int limit = 100;
        return ToolResult.Ok(new
        {
            vb6DebugRegistrations = DebugLeftovers(classes, libs, await DebuggingProjectsAsync()),
            classesWithMissingServer = new { total = missingClasses.Count, items = missingClasses.Take(limit).Select(c => new { clsid = c.Clsid.ToString("B").ToUpperInvariant(), progId = c.ProgId, server = c.ServerFile }) },
            typeLibsWithMissingFile = new { total = missingLibs.Count, items = missingLibs.Take(limit).Select(t => new { libId = t.LibId.ToString("B").ToUpperInvariant(), version = $"{t.Major}.{t.Minor}", description = t.Description, file = t.Win32Path }) },
            scanned = new { classes = classes.Count, typeLibs = libs.Count },
            note = "report only: nothing was changed. Registrations of files outside the workspace roots can't be cleaned up with these tools.",
        });
    }

    // ---------- diagnose ----------

    private sealed record Check(string Item, string Status, string Detail);

    private static readonly Dictionary<int, string> s_errorHints = new()
    {
        [429] = "ActiveX component can't create object. Usual causes: the class isn't registered, its server file is missing, can't load or is 64-bit, or it's registered to the VB6 debugger after an IDE was killed.",
        [430] = "Class doesn't support Automation or does not support expected interface. Usual cause: the registered build isn't the one the program was compiled against (rebuilt without binary compatibility, or another copy registered).",
        [48] = "Error in loading DLL. Usual causes: the file is missing, is 64-bit, or a DLL it depends on is missing (dependencies aren't checked here).",
        [53] = "File not found. For COM: the registered server file doesn't exist.",
    };

    private async Task<ToolResult> ComDiagnoseAsync(JsonObject a)
    {
        int error = a.Int("error") ?? throw new ToolException("missing required argument 'error'");
        string? progId = a.Str("prog_id"), clsidText = a.Str("clsid"), file = a.Str("file"), project = a.Str("project");
        if (new[] { progId, clsidText, file, project }.Count(s => !string.IsNullOrEmpty(s)) != 1) throw new ToolException("give exactly one of prog_id, clsid, file or project");
        var libs = TypeLibRegistry.All();
        var checks = new List<Check>();
        var debugging = await DebuggingProjectsAsync();
        ProjectMap? map = roots.Count > 0 ? Map(false) : null;
        // the folders of the programs that load these DLLs: a delay import is looked for there, not next to the DLL
        var appFolders = new List<string>();
        if (Vb6Install.FindExe() is { } vb6) appFolders.Add(Path.GetDirectoryName(vb6)!);
        if (!string.IsNullOrEmpty(project)) appFolders.Add(Path.GetDirectoryName(PathKey.Normalize(project))!);

        // spike 39: registered, present and 32-bit, yet it can't load: a DLL it needs (here delay-loaded) is missing
        void CheckDependencies(string label, string path)
        {
            var missing = PeDependencies.Missing(path, appFolders);
            if (missing.Count == 0) return;   // only problems: an ok row per component doubled the result of a project check
            foreach (var m in missing.Take(10))
                checks.Add(new(label + " dependencies", "error", string.Join(" -> ", m.Chain) + ": not found" +
                    (m.Note is { } n ? $" ({n})" : "") +
                    (m.Delay ? ". A DLL in this chain is delay-loaded: the program fails only when it first needs it, with exception 0xc06d007e, which ends the process (the IDE too, spike 39)"
                             : ". Loading the DLL fails (error 48 or 429)")));
            if (missing.Count > 10) checks.Add(new(label + " dependencies", "error", $"{missing.Count - 10} more missing"));
        }

        void CheckServerFile(string label, string? path)
        {
            if (path is null) { checks.Add(new(label, "error", "no server file registered")); return; }
            if (ComRegistry.IsVbDebugServer(path)) { checks.Add(new(label, "error", $"registered to the VB6 debugger ({path}); see com_orphans for the repair (left over unless that project is running in an IDE now)")); return; }
            if (!File.Exists(path)) { checks.Add(new(label, "error", $"server file missing: {path}")); return; }
            var bits = ComRegistry.Bitness(path);
            checks.Add(new(label, bits is "x86" or null ? "ok" : "error", $"{path} ({bits ?? "machine type unreadable"}, written {Stamp(path)})" + (bits is "x86" or null ? "" : "; VB6 programs are 32-bit and can't load it")));
            if (bits is "x86") CheckDependencies(label, path);
        }
        void CheckTypeLib(string label, Guid lib, int? major)
        {
            var reg = libs.Where(t => t.LibId == lib).ToList();
            if (reg.Count == 0) { checks.Add(new(label, "error", $"typelib {lib.ToString("B").ToUpperInvariant()} isn't registered")); return; }
            var exact = major is null ? reg : reg.Where(t => t.Major == major).ToList();
            if (exact.Count == 0) { checks.Add(new(label, "error", $"version {major}.x isn't registered; registered: {string.Join(", ", reg.Select(t => $"{t.Major}.{t.Minor}"))}")); return; }
            foreach (var t in exact)
            {
                if (ComRegistry.IsProjectTypeLib(t))
                    checks.Add(new(label, debugging.Contains(PathKey.Normalize(t.Win32Path!)) ? "warning" : "error", $"{t.Major}.{t.Minor} is registered to a project file ({t.Win32Path}): VB6 debugger registration" + (debugging.Contains(PathKey.Normalize(t.Win32Path!)) ? " (that project is running in an IDE now)" : ", left over; see com_orphans")));
                else if (!t.FileExists) checks.Add(new(label, "error", $"{t.Major}.{t.Minor} -> {t.Win32Path ?? "(no file)"} (missing)"));
                // Registered and present is not enough: the file must contain a type library. A .NET COM DLL registered
                // without /tlb points the registration at a DLL that has none, and VB6 marks the reference MISSING
                // (a production project, 2026-10-15: Vendor.License.COM.dll).
                else if (TypeLibInfo.Read(t.Win32Path!) is null)
                    checks.Add(new(label, "error", $"{t.Major}.{t.Minor} -> {t.Win32Path}: registered, but the file contains no type library, so VB6 shows the reference as MISSING (unless the .vbp hint path has one). For a .NET COM DLL: RegAsm <dll> /tlb (elevated) writes and registers one"));
                else checks.Add(new(label, "ok", $"{t.Major}.{t.Minor} -> {t.Win32Path}"));
            }
        }
        void CheckClass(Guid clsid, string label)
        {
            var c = ComRegistry.Class(clsid);
            if (c is null) { checks.Add(new(label, "error", $"class {clsid.ToString("B").ToUpperInvariant()} isn't registered")); return; }
            checks.Add(new(label, "ok", $"registered{(c.ProgId is null ? "" : " as " + c.ProgId)}"));
            CheckServerFile(label + " server", c.ServerFile);
            if (c.TypeLib is { } tl) CheckTypeLib(label + " typelib", tl, null);
            CompatOf(c.ServerFile);
        }
        void CompatOf(string? path)
        {
            if (map is null || path is null) return;
            foreach (var p in map.Projects.Where(p => p.OutputPath is { } o && PathKey.Same(o, path)))
                foreach (var f in Compat.Check(p, f => TypeLibInfo.Read(f), libs).Findings)
                    checks.Add(new($"{p.Name} (builds this file)", f.Level, f.Text));
        }

        if (!string.IsNullOrEmpty(progId))
        {
            var clsid = ComRegistry.ClsidOf(progId);
            checks.Add(new("ProgID " + progId, clsid is null ? "error" : "ok", clsid is null ? "not registered (32-bit view)" : "-> " + clsid.Value.ToString("B").ToUpperInvariant()));
            if (clsid is { } g) CheckClass(g, "class");
        }
        else if (!string.IsNullOrEmpty(clsidText))
        {
            if (!Guid.TryParse(clsidText, out var g)) throw new ToolException("clsid isn't a GUID");
            CheckClass(g, "class");
        }
        else if (!string.IsNullOrEmpty(file))
        {
            var path = PathKey.Normalize(file);
            if (!File.Exists(path)) checks.Add(new("file", "error", "not found: " + path));
            else
            {
                CheckServerFile("file", path);
                var reg = FileRegistration(path, libs);
                if (reg.TypeLib is null) checks.Add(new("typelib", "warning", reg.Note ?? "no type library"));
                else checks.Add(new("registration", reg.AllHere ? "ok" : reg.AnyElsewhere ? "warning" : "error",
                    reg.AllHere ? "the typelib and every class point at this file" : reg.AnyElsewhere ? "some classes are registered to another file (com_where file=...)" : "not (fully) registered: com_register"));
                CompatOf(path);
            }
        }
        else
        {
            var vbpPath = PathKey.Normalize(project!);
            if (!File.Exists(vbpPath)) throw new ToolException("project not found: " + vbpPath);
            var vbp = VbpFile.Load(vbpPath);
            foreach (var line in vbp.References.Concat(vbp.Objects))
            {
                var d = ComDependency.Parse(line);
                if (d is null) { checks.Add(new("reference " + line, "info", "not a typelib reference (project reference or unknown form)")); continue; }
                var name = line.Split('#').LastOrDefault() is { Length: > 0 } n && !line.TrimStart().StartsWith('{') ? n : line.Split(';').Last().Trim();
                CheckTypeLib(name, d.LibId, d.Major);
                var reg = libs.Where(t => t.LibId == d.LibId && t.Major == d.Major && !ComRegistry.IsProjectTypeLib(t)).Select(t => t.Win32Path).OfType<string>().Where(File.Exists).FirstOrDefault();
                bool binary = reg is not null && !reg.EndsWith(".tlb", StringComparison.OrdinalIgnoreCase) && !reg.EndsWith(".olb", StringComparison.OrdinalIgnoreCase);
                if (binary && ComRegistry.Bitness(reg!) is { } bits && bits != "x86")
                    checks.Add(new(name, "error", $"{reg} is {bits}; VB6 programs are 32-bit"));
                else if (binary) CheckDependencies(name, reg!);
            }
        }
        return ToolResult.Ok(new
        {
            error, meaning = s_errorHints.GetValueOrDefault(error, "no specific hints for this error; the checks below cover registration, files and versions"),
            problems = checks.Count(c => c.Status == "error"),
            checks = checks.Select(c => new { item = c.Item, status = c.Status, detail = c.Detail }),
        });
    }

    // ---------- describe / search / where ----------

    private ToolResult ComDescribe(JsonObject a)
    {
        string? file = a.Str("file"), progId = a.Str("prog_id"), libid = a.Str("libid");
        if (new[] { file, progId, libid }.Count(s => !string.IsNullOrEmpty(s)) != 1) throw new ToolException("give exactly one of file, prog_id or libid (+ version)");
        var registered = TypeLibRegistry.All();
        string path;
        if (!string.IsNullOrEmpty(file))
        {
            path = PathKey.Normalize(file);
            if (!ProjectMap.IsUnderRoots(path, roots) && !registered.Any(t => t.Win32Path is { } w && PathKey.Same(w, path)))
                throw new ToolException($"{path} is neither under the workspace roots nor a registered typelib's file");
        }
        else if (!string.IsNullOrEmpty(progId))
        {
            var clsid = ComRegistry.ClsidOf(progId) ?? throw new ToolException($"ProgID {progId} isn't registered (32-bit view)");
            var cls = ComRegistry.Class(clsid) ?? throw new ToolException($"{progId} points at {clsid.ToString("B").ToUpperInvariant()}, which isn't registered");
            var lib = cls.TypeLib is { } g ? registered.Where(t => t.LibId == g).OrderByDescending(t => t.Major).ThenByDescending(t => t.Minor).FirstOrDefault() : null;
            path = lib?.Win32Path ?? cls.ServerFile ?? throw new ToolException($"{progId} has neither a registered typelib nor a server file");
        }
        else
        {
            if (!Guid.TryParse(libid, out var g)) throw new ToolException("libid isn't a GUID");
            var versions = registered.Where(t => t.LibId == g).ToList();
            var pick = a.Str("version") is { } v ? versions.FirstOrDefault(t => $"{t.Major}.{t.Minor}" == v) : versions.OrderByDescending(t => t.Major).ThenByDescending(t => t.Minor).FirstOrDefault();
            path = pick?.Win32Path ?? throw new ToolException($"{g.ToString("B").ToUpperInvariant()} {a.Str("version")} isn't registered with a file");
        }
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);
        TlbDescription d;
        try { d = TypeLibDescriber.Describe(path); }
        catch (InvalidOperationException ex) { throw new ToolException(ex.Message); }
        var filter = a.Str("filter");
        var text = TypeLibDescriber.Render(d, filter, a.Bool("include_hidden") == true);
        var info = new { libId = d.Info.LibId.ToString("B").ToUpperInvariant(), version = $"{d.Info.Major}.{d.Info.Minor}", name = d.Info.Name, description = d.Doc, file = path };
        if (text.Length > DescribeLimit && filter is null)
            return ToolResult.Ok(new
            {
                typeLib = info, note = $"the full listing is {text.Length:N0} characters; call again with filter (a type name) to see members",
                types = d.Types.Where(t => !t.Hidden).Select(t => new { t.Name, t.Kind, members = t.Members.Count(m => !m.Hidden) }),
            });
        return ToolResult.Ok(new { typeLib = info, declarations = text.Length > DescribeLimit ? text[..DescribeLimit] + "\n' ... truncated: narrow the filter" : text });
    }

    private ToolResult ComSearch(JsonObject a)
    {
        var text = a.Req("text");
        var libs = TypeLibRegistry.All().Where(t => t.Description.Contains(text, StringComparison.OrdinalIgnoreCase) || (t.Win32Path ?? "").Contains(text, StringComparison.OrdinalIgnoreCase))
            .OrderBy(t => t.Description, StringComparer.OrdinalIgnoreCase).ThenBy(t => t.Major).ThenBy(t => t.Minor).Take(50)
            .Select(t => new { libId = t.LibId.ToString("B").ToUpperInvariant(), version = $"{t.Major}.{t.Minor}", description = t.Description, file = t.Win32Path, fileExists = t.FileExists }).ToList();
        var progIds = ComRegistry.FindProgIds(text).Select(p =>
        {
            var c = ComRegistry.Class(p.Clsid);
            return new { progId = p.ProgId, clsid = p.Clsid.ToString("B").ToUpperInvariant(), file = c?.ServerFile, fileExists = c?.ServerExists ?? false };
        }).ToList();
        return ToolResult.Ok(new { text, typeLibs = libs, progIds, note = "at most 50 of each; 32-bit registry view" });
    }

    private ToolResult ComWhere(JsonObject a)
    {
        string? progId = a.Str("prog_id"), clsidText = a.Str("clsid"), libid = a.Str("libid"), file = a.Str("file");
        if (new[] { progId, clsidText, libid, file }.Count(s => !string.IsNullOrEmpty(s)) != 1) throw new ToolException("give exactly one of prog_id, clsid, libid or file");
        var registered = TypeLibRegistry.All();
        ProjectMap? map = roots.Count > 0 ? Map(false) : null;

        if (!string.IsNullOrEmpty(file))
        {
            var path = PathKey.Normalize(file);
            if (!File.Exists(path)) throw new ToolException("file not found: " + path);
            var reg = FileRegistration(path, registered);
            return ToolResult.Ok(new { file = path, written = Stamp(path), builtBy = BuiltBy(map, path), reg.TypeLib, reg.Classes, registeredHere = reg.AllHere, note = reg.Note });
        }
        if (!string.IsNullOrEmpty(libid))
        {
            if (!Guid.TryParse(libid, out var g)) throw new ToolException("libid isn't a GUID");
            var versions = registered.Where(t => t.LibId == g).OrderBy(t => t.Major).ThenBy(t => t.Minor).ToList();
            return ToolResult.Ok(new
            {
                libId = g.ToString("B").ToUpperInvariant(),
                versions = versions.Select(t => new { version = $"{t.Major}.{t.Minor}", t.Description, file = t.Win32Path, fileExists = t.FileExists, written = Stamp(t.Win32Path), builtBy = BuiltBy(map, t.Win32Path) }),
                builtBy = map?.Projects.Where(p => Compat.IsComServer(p) && Affected.TypeLibOf(p) is { } tl && tl.LibId == g).Select(p => p.Path),
                note = versions.Count == 0 ? "not registered on this machine" : null,
            });
        }
        Guid clsid;
        if (!string.IsNullOrEmpty(progId)) clsid = ComRegistry.ClsidOf(progId) ?? throw new ToolException($"ProgID {progId} isn't registered (32-bit view)");
        else if (!Guid.TryParse(clsidText, out clsid)) throw new ToolException("clsid isn't a GUID");
        var c = ComRegistry.Class(clsid) ?? throw new ToolException($"{clsid.ToString("B").ToUpperInvariant()} isn't registered (32-bit view)");
        var lib = c.TypeLib is { } tg ? registered.Where(t => t.LibId == tg).ToList() : [];
        return ToolResult.Ok(new
        {
            clsid = clsid.ToString("B").ToUpperInvariant(), progId = c.ProgId, server = c.ServerFile, serverExists = c.ServerExists, written = Stamp(c.ServerFile),
            threadingModel = c.ThreadingModel, builtBy = BuiltBy(map, c.ServerFile),
            warning = ComRegistry.IsVbDebugServer(c.ServerFile) ? "registered to the VB6 debugger: the project is running in an IDE now, or this was left over by a killed IDE (com_orphans shows which, and the repair)" : null,
            typeLib = c.TypeLib is null ? null : new
            {
                libId = c.TypeLib.Value.ToString("B").ToUpperInvariant(), classVersion = c.TypeLibVersion,
                registered = lib.Select(t => new { version = $"{t.Major}.{t.Minor}", file = t.Win32Path, fileExists = t.FileExists }),
            },
        });
    }

    private static string? Stamp(string? path) => path is not null && File.Exists(path) ? File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss") : null;

    /// <summary>Workspace projects whose build output is this file.</summary>
    private static IEnumerable<string>? BuiltBy(ProjectMap? map, string? path)
        => map is null || path is null ? null : map.Projects.Where(p => p.OutputPath is { } o && PathKey.Same(o, path)).Select(p => p.Path);

    /// <summary>
    /// Whether com_(un)register did its job. Unregister: nothing points at the file any more. Register: everything points
    /// here, or regsvr32 succeeded, the typelib points here, no class points at another file and at least one class (if
    /// any) is registered: classes nothing registers are listed, not failed (improvement notes 18).
    /// </summary>
    public static bool Verified(bool unregister, int? exit, bool allHere, bool anyHere, bool libHere, int elsewhere, int missing, int classes)
        => unregister ? !anyHere : allHere || (exit == 0 && libHere && elsewhere == 0 && (classes == 0 || missing < classes));

    private sealed record Registration(object? TypeLib, List<object> Classes, bool AllHere, bool AnyElsewhere, bool AnyHere, string? Note)
    {
        public bool LibHere { get; init; }
        /// <summary>Creatable classes of the typelib that no file registers (often internal helpers the DLL never registers).</summary>
        public List<string> Missing { get; init; } = [];
        /// <summary>Classes registered to another file.</summary>
        public List<string> Elsewhere { get; init; } = [];
    }

    /// <summary>
    /// The registration state of a file: its typelib (read from the file) and its creatable classes, each with where
    /// the registry points. AllHere = everything registered to this file; AnyElsewhere = a class points at another file.
    /// </summary>
    private static Registration FileRegistration(string path, IReadOnlyList<RegisteredTypeLib> registered)
    {
        TlbDescription d;
        try { d = TypeLibDescriber.Describe(path); }
        catch (InvalidOperationException) { return new Registration(null, [], false, false, false, "the file has no type library, so its classes can't be listed"); }
        var tl = registered.FirstOrDefault(t => t.LibId == d.Info.LibId && t.Major == d.Info.Major && t.Minor == d.Info.Minor);
        bool libHere = tl?.Win32Path is { } w && PathKey.Same(w, path);
        var classes = new List<object>();
        var missing = new List<string>(); var elsewhereNames = new List<string>();
        bool allHere = libHere, anyElsewhere = false, anyHere = libHere;
        foreach (var t in d.Types.Where(t => t.Kind == "coclass" && (t.Creatable || t.Control)))
        {
            var c = ComRegistry.Class(t.Guid);
            bool here = c?.ServerFile is { } s && PathKey.Same(s, path);
            bool elsewhere = c?.ServerFile is { } s2 && !PathKey.Same(s2, path);
            allHere &= here; anyElsewhere |= elsewhere; anyHere |= here;
            if (c?.ServerFile is null) missing.Add(t.Name);
            if (elsewhere) elsewhereNames.Add(t.Name);
            classes.Add(new { name = t.Name, clsid = t.Guid.ToString("B").ToUpperInvariant(), progId = c?.ProgId, registeredTo = c?.ServerFile, here });
        }
        var lib = new
        {
            libId = d.Info.LibId.ToString("B").ToUpperInvariant(), version = $"{d.Info.Major}.{d.Info.Minor}", name = d.Info.Name,
            registeredTo = tl?.Win32Path, here = libHere,
        };
        return new Registration(lib, classes, allHere, anyElsewhere, anyHere, null) { LibHere = libHere, Missing = missing, Elsewhere = elsewhereNames };
    }

    // ---------- compatibility ----------

    private ToolResult CompatStatus(JsonObject a)
    {
        var map = Map(a.Bool("refresh") == true);
        var spec = a.Str("project");
        IEnumerable<VbpFile> projects = map.Projects.Where(Compat.IsComServer);
        if (spec is { Length: > 0 })
            projects = spec.EndsWith(".vbp", StringComparison.OrdinalIgnoreCase)
                ? map.Projects.Where(p => PathKey.Same(p.Path, PathKey.Normalize(spec)))
                : projects.Where(p => p.Name.Contains(spec, StringComparison.OrdinalIgnoreCase));
        var registered = TypeLibRegistry.All();
        var reports = projects.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).Take(100).Select(p => Compat.Check(p, f => TypeLibInfo.Read(f), registered)).ToList();
        if (reports.Count == 0) throw new ToolException(spec is null ? "no ActiveX projects under the roots" : $"no project matches '{spec}'");
        static object? Lib(TypeLibInfo? t) => t is null ? null : new { libId = t.LibId.ToString("B").ToUpperInvariant(), version = $"{t.Major}.{t.Minor}" };
        return ToolResult.Ok(new
        {
            projects = reports.Select(r => new
            {
                r.Project.Name, r.Project.Type, path = r.Project.Path, mode = r.Mode,
                version = $"{r.Project.Settings.GetValueOrDefault("MajorVer", "?")}.{r.Project.Settings.GetValueOrDefault("MinorVer", "?")}.{r.Project.Settings.GetValueOrDefault("RevisionVer", "?")}",
                compatibleFile = r.CompatibleFile, compatibleFileExists = r.CompatibleFileExists, compatibleTypeLib = Lib(r.CompatibleTypeLib),
                output = r.Output, outputExists = r.OutputExists, outputTypeLib = Lib(r.OutputTypeLib),
                registered = r.Registered.Select(t => new { version = $"{t.Major}.{t.Minor}", file = t.Win32Path, fileExists = t.FileExists }),
                findings = r.Findings.Select(f => new { level = f.Level, text = f.Text }),
            }),
            summary = new { projects = reports.Count, errors = reports.Sum(r => r.Findings.Count(f => f.Level == "error")), warnings = reports.Sum(r => r.Findings.Count(f => f.Level == "warning")) },
        });
    }

    // ---------- (un)registration ----------

    private async Task<ToolResult> RegisterAsync(JsonObject a, bool unregister)
    {
        string verb = unregister ? "com_unregister" : "com_register";
        if (a.Bool("confirm") != true) throw new ToolException($"{verb} changes the machine's COM registration; it needs confirm=true");
        var path = PathKey.Normalize(a.Req("file"));
        RequireUnderRoots(path);
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext is not (".dll" or ".ocx")) throw new ToolException("only .dll and .ocx files are registered this way");
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);

        var before = FileRegistration(path, TypeLibRegistry.All());
        if (unregister && before.AnyElsewhere && a.Bool("force") != true)
            throw new ToolException("some of this file's classes are registered to another file; unregistering this one removes those registrations too. " +
                                    "Check with com_where, then pass force=true if that's intended");

        var live = env.Scan().Live;
        var inst = a.Int("pid") is int pid ? live.FirstOrDefault(i => i.Pid == pid) ?? throw new ToolException($"no VB6 IDE with the add-in has pid {pid}")
                 : live.Count == 1 ? live[0]
                 : throw new ToolException(live.Count == 0 ? "registration runs elevated through the VB6 add-in: start VB6 (with the add-in) first" : "several VB6 IDEs are running; pass pid");
        var client = env.Client(inst);
        int child;
        try { child = await client.RegsvrAsync(path, unregister); }
        catch (AddinErrorException ex) when (ex.Code == 400 && ex.AddinMessage.StartsWith("unknown verb", StringComparison.Ordinal))
        { throw new ToolException("the add-in is older than this server (no REGSVR verb): rebuild and register it (scripts\\build-addin.ps1)"); }
        catch (AddinErrorException ex) { throw new ToolException(ex.AddinMessage); }
        Audit.Write(verb, inst.Pid, $"regsvr32{(unregister ? " /u" : "")} \"{path}\" child pid {child}");

        int? exit = null;
        var deadline = DateTime.UtcNow.AddSeconds(60);
        while (exit is null && DateTime.UtcNow < deadline)
        {
            await Task.Delay(200);
            exit = await client.RegStatusAsync(child);
        }
        var after = FileRegistration(path, TypeLibRegistry.All());
        // Registered = regsvr32 succeeded, the typelib points here and no class points at another file. A class of the
        // typelib that nothing registers isn't a failure (improvement notes 18: CapPRO.dll registered 5 of 6 classes;
        // the sixth, RivetTalk, is an internal helper the DLL never registers): listed in notRegistered instead.
        bool verified = Verified(unregister, exit, after.AllHere, after.AnyHere, after.LibHere, after.Elsewhere.Count, after.Missing.Count, after.Classes.Count);
        Audit.Write(verb, inst.Pid, $"\"{path}\" exit {exit?.ToString() ?? "timeout"} verified={verified}");
        return ToolResult.Ok(new
        {
            file = path, action = unregister ? "unregister" : "register", exitCode = exit, verified,
            notRegistered = !unregister && after.Missing.Count > 0 ? after.Missing : null,
            registeredElsewhere = !unregister && after.Elsewhere.Count > 0 ? after.Elsewhere : null,
            before = new { before.TypeLib, before.Classes }, after = new { after.TypeLib, after.Classes },
            note = exit is null ? "regsvr32 didn't finish within 60 s" :
                   verified && !unregister && after.Missing.Count > 0 ? $"registered; the file doesn't register {string.Join(", ", after.Missing)} (classes in its typelib that nothing registers are usually internal helpers, not a failure)" :
                   verified ? null :
                   !unregister && after.Elsewhere.Count > 0 ? $"regsvr32 finished, but {string.Join(", ", after.Elsewhere)} {(after.Elsewhere.Count == 1 ? "is" : "are")} registered to another file (com_where)" :
                   unregister ? "regsvr32 finished, but registry entries still point at this file" : "regsvr32 finished, but the registry doesn't point at this file for everything",
        });
    }
}
