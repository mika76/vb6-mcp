namespace Vb6Mcp.Core.Files;

/// <summary>A member of a project: kind (the .vbp key), component name if the line has one, and paths.</summary>
public sealed record VbpMember(string Kind, string? Name, string RelativePath, string FullPath);

/// <summary>
/// A parsed .vbp (read-only view). Member paths are relative to the .vbp folder, may climb out of it
/// ("..\..\Shared\Utils.bas") and may use either slash; they're resolved with <see cref="PathKey"/>.
/// References and objects are kept as the raw text after "=": GUIDs are never parsed into new lines
/// or rewritten here (AGENTS.md hard rule 13).
/// </summary>
public sealed record VbpFile(
    string Path, string Name, string Type, IReadOnlyList<VbpMember> Members,
    IReadOnlyList<string> References, IReadOnlyList<string> Objects,
    string? ExeName, string? OutputPath, string? Startup, int CompatibleMode, string? CompatibleFile)
{
    /// <summary>Keys whose value names a source file of the project.</summary>
    public static readonly string[] MemberKeys = ["Form", "Module", "Class", "UserControl", "PropertyPage", "Designer", "UserDocument", "RelatedDoc"];

    public static VbpFile Load(string path) => Parse(path, Vb1252.Encoding.GetString(File.ReadAllBytes(path)));

    public static VbpFile Parse(string path, string text)
    {
        string dir = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path))!;
        var members = new List<VbpMember>();
        var refs = new List<string>();
        var objs = new List<string>();
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var raw in text.Replace("\r\n", "\n").Split('\n'))
        {
            var line = raw.TrimEnd();
            if (line.StartsWith('[')) break;                     // add-in sections like [MS Transaction Server]
            int eq = line.IndexOf('=');
            if (eq <= 0) continue;
            string key = line[..eq].Trim(), value = line[(eq + 1)..].Trim();

            if (key.Equals("Reference", StringComparison.OrdinalIgnoreCase)) { refs.Add(value); continue; }
            if (key.Equals("Object", StringComparison.OrdinalIgnoreCase)) { objs.Add(value); continue; }
            if (MemberKeys.Contains(key, StringComparer.OrdinalIgnoreCase))
            {
                // "Module=modMath; modMath.bas" has a name; "Form=frmMain.frm" doesn't
                string? name = null; string rel = value;
                int semi = value.IndexOf(';');
                if (semi >= 0) { name = value[..semi].Trim(); rel = value[(semi + 1)..].Trim(); }
                rel = rel.Trim('"');
                members.Add(new VbpMember(key, name, rel, PathKey.Normalize(rel, dir)));
                continue;
            }
            values.TryAdd(key, value.Trim('"'));
        }

        string? exe = values.GetValueOrDefault("ExeName32");
        string? outDir = values.GetValueOrDefault("Path32");
        string? output = exe is null ? null : PathKey.Normalize(System.IO.Path.Combine(string.IsNullOrEmpty(outDir) ? "." : outDir, exe), dir);
        string? compat = values.GetValueOrDefault("CompatibleEXE32") is { Length: > 0 } c ? PathKey.Normalize(c, dir) : null;
        int.TryParse(values.GetValueOrDefault("CompatibleMode"), out int compatMode);

        return new VbpFile(PathKey.Normalize(path), values.GetValueOrDefault("Name") ?? System.IO.Path.GetFileNameWithoutExtension(path),
            values.GetValueOrDefault("Type") ?? "", members, refs, objs, exe, output, values.GetValueOrDefault("Startup"), compatMode, compat);
    }
}

/// <summary>A parsed .vbg project group: the member .vbp paths (resolved) and the startup project.</summary>
public sealed record VbgFile(string Path, IReadOnlyList<string> Projects, string? StartupProject)
{
    public static VbgFile Load(string path)
    {
        string dir = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path))!;
        var projects = new List<string>();
        string? startup = null;
        foreach (var raw in Vb1252.Encoding.GetString(File.ReadAllBytes(path)).Replace("\r\n", "\n").Split('\n'))
        {
            int eq = raw.IndexOf('=');
            if (eq <= 0) continue;
            string key = raw[..eq].Trim(), value = raw[(eq + 1)..].Trim();
            if (key.Equals("Project", StringComparison.OrdinalIgnoreCase)) projects.Add(PathKey.Normalize(value, dir));
            else if (key.Equals("StartupProject", StringComparison.OrdinalIgnoreCase)) { startup = PathKey.Normalize(value, dir); projects.Add(startup); }
        }
        return new VbgFile(PathKey.Normalize(path), projects, startup);
    }
}
