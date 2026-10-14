namespace Vb6Mcp.Core.Files;

/// <summary>One project that needs a rebuild, and why.</summary>
/// <param name="Project">The project.</param>
/// <param name="Includes">Changed files the project includes itself (empty if it's affected only through a dependency).</param>
/// <param name="Uses">Affected projects whose DLL/OCX/EXE typelib this project references.</param>
/// <param name="TypeLib">The typelib of the project's output (or compatibility file), if it has one that could be read.</param>
public sealed record AffectedProject(VbpFile Project, IReadOnlyList<string> Includes, IReadOnlyList<VbpFile> Uses, TypeLibInfo? TypeLib);

/// <summary>Result of <see cref="Affected.Find"/>: projects in rebuild order, plus warnings.</summary>
public sealed record AffectedResult(IReadOnlyList<AffectedProject> Order, IReadOnlyList<string> Warnings);

/// <summary>
/// Which projects a change to shared source files affects, and the order to rebuild them in
/// (ws_affected, PLAN.md §9): the projects that include a file, then, transitively, every project that
/// references one of those projects' COM output (DLL, OCX, ActiveX EXE). Dependencies come first.
/// A reference matches a project by the LIBID and major version of the typelib in its built output
/// (or its binary compatibility file), never by file name (hard rule 13).
/// </summary>
public static class Affected
{
    /// <param name="map">The project map.</param>
    /// <param name="files">Changed source files (any spelling).</param>
    /// <param name="typeLibOf">Reads a project's typelib (e.g. <see cref="TypeLibOf"/>); injected for tests.</param>
    public static AffectedResult Find(ProjectMap map, IEnumerable<string> files, Func<VbpFile, TypeLibInfo?> typeLibOf)
    {
        var warnings = new List<string>();
        var includes = new Dictionary<VbpFile, List<string>>();
        foreach (var f in files.Select(f => PathKey.Normalize(f)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var users = map.ProjectsUsing(f);
            if (users.Count == 0) warnings.Add($"no project under the roots includes {f}");
            foreach (var p in users)
            {
                if (!includes.TryGetValue(p, out var l)) includes[p] = l = [];
                l.Add(f);
            }
        }

        // Each project's dependencies on others, by LIBID + major (computed lazily, typelibs read once)
        var libs = new Dictionary<VbpFile, TypeLibInfo?>();
        TypeLibInfo? Lib(VbpFile p)
        {
            if (!libs.TryGetValue(p, out var t)) libs[p] = t = typeLibOf(p);
            return t;
        }
        var deps = map.Projects.ToDictionary(p => p, p => p.References.Concat(p.Objects).Select(ComDependency.Parse).OfType<ComDependency>().ToList());

        // Closure: affected producers → the projects that reference them
        var uses = includes.Keys.ToDictionary(p => p, _ => new List<VbpFile>());
        var queue = new Queue<VbpFile>(includes.Keys);
        while (queue.Count > 0)
        {
            var producer = queue.Dequeue();
            var lib = Lib(producer);
            if (lib is null)
            {
                if (!producer.Type.Equals("Exe", StringComparison.OrdinalIgnoreCase))
                    warnings.Add($"{producer.Name}: no typelib found in {producer.OutputPath ?? "(no output file set)"}{(producer.CompatibleFile is null ? "" : " or " + producer.CompatibleFile)}; build it once to find the projects that use it");
                continue;
            }
            foreach (var consumer in map.Projects)
            {
                if (consumer == producer || !deps[consumer].Any(d => d.LibId == lib.LibId && d.Major == lib.Major)) continue;
                if (!uses.TryGetValue(consumer, out var u)) { uses[consumer] = u = []; queue.Enqueue(consumer); }
                if (!u.Contains(producer)) u.Add(producer);
            }
        }

        // Two projects building the same typelib can't be told apart: say so instead of guessing
        foreach (var g in uses.Keys.Select(p => (p, lib: Lib(p))).Where(x => x.lib is not null).GroupBy(x => (x.lib!.LibId, x.lib.Major)).Where(g => g.Count() > 1))
            warnings.Add($"{string.Join(", ", g.Select(x => x.p.Path))} build the same typelib {{{g.Key.LibId}}} v{g.Key.Major}; their users are listed under each");

        // Rebuild order: Kahn's algorithm, producers before consumers, ties by name
        var pending = uses.ToDictionary(kv => kv.Key, kv => kv.Value.Count(uses.ContainsKey));
        var order = new List<VbpFile>();
        while (pending.Count > 0)
        {
            var ready = pending.Where(kv => kv.Value == 0).Select(kv => kv.Key).OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).ThenBy(p => p.Path, StringComparer.OrdinalIgnoreCase).ToList();
            if (ready.Count == 0)
            {
                var cycle = pending.Keys.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).ToList();
                warnings.Add($"circular references between {string.Join(", ", cycle.Select(p => p.Name))}; they're listed in name order");
                order.AddRange(cycle);
                break;
            }
            foreach (var p in ready)
            {
                order.Add(p);
                pending.Remove(p);
                foreach (var (c, u) in uses) if (pending.ContainsKey(c) && u.Contains(p)) pending[c]--;
            }
        }
        return new AffectedResult(
            order.Select(p => new AffectedProject(p, includes.TryGetValue(p, out var i) ? i : [], uses[p], libs.GetValueOrDefault(p))).ToList(),
            warnings);
    }

    /// <summary>The typelib in a project's built output, else in its binary compatibility file; null for Standard EXEs and unbuilt projects.</summary>
    public static TypeLibInfo? TypeLibOf(VbpFile p)
    {
        foreach (var file in new[] { p.OutputPath, p.CompatibleFile })
            if (file is not null && File.Exists(file) && TypeLibInfo.Read(file) is { } t) return t;
        return null;
    }
}
