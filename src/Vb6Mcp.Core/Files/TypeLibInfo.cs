using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Vb6Mcp.Core.Files;

/// <summary>
/// Identity of a type library: LIBID and version. VB6 references a DLL/OCX by exactly this
/// ("Reference=*\G{LIBID}#maj.min#lcid#path#desc", "Object={LIBID}#maj.min#lcid; file.ocx"), so this is
/// how a project's output is matched to the projects that use it (hard rule 13: by typelib GUID and
/// major version, never by file name).
/// </summary>
public sealed record TypeLibInfo(Guid LibId, int Major, int Minor, string Name)
{
    private const int REGKIND_NONE = 2;

    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern ITypeLib LoadTypeLibEx(string file, int regKind);

    /// <summary>
    /// Reads the typelib embedded in a DLL/OCX/EXE (or a .tlb/.olb) without registering it (REGKIND_NONE).
    /// Null if the file has no typelib or can't be loaded. Runs on its own STA thread (AGENTS.md: in-process
    /// COM never on thread-pool threads), with a timeout so a bad file can't hang a request.
    /// </summary>
    public static TypeLibInfo? Read(string path, TimeSpan? timeout = null)
    {
        TypeLibInfo? result = null;
        var t = new Thread(() =>
        {
            ITypeLib? lib = null;
            try
            {
                lib = LoadTypeLibEx(path, REGKIND_NONE);
                lib.GetLibAttr(out var p);
                try
                {
                    var attr = Marshal.PtrToStructure<TYPELIBATTR>(p);
                    lib.GetDocumentation(-1, out var name, out _, out _, out _);
                    result = new TypeLibInfo(attr.guid, attr.wMajorVerNum, attr.wMinorVerNum, name);
                }
                finally { lib.ReleaseTLibAttr(p); }
            }
            catch (Exception ex) when (ex is COMException or FileNotFoundException or ArgumentException or InvalidCastException) { }
            finally { if (lib is not null) Marshal.ReleaseComObject(lib); }
        }) { IsBackground = true, Name = "typelib" };
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        return t.Join(timeout ?? TimeSpan.FromSeconds(10)) ? result : null;
    }
}

/// <summary>
/// A COM dependency of a project, from a .vbp "Reference=" or "Object=" line: the LIBID and major version.
/// Project references ("*\A...vbp", used inside project groups) aren't parsed yet: not verified on this machine.
/// </summary>
public sealed record ComDependency(Guid LibId, int Major, string Raw)
{
    /// <summary>Parses "*\G{guid}#maj.min#lcid#path#desc" (Reference=) or "{guid}#maj.min#lcid; file" (Object=). Null otherwise.</summary>
    public static ComDependency? Parse(string raw)
    {
        var s = raw.Trim().Trim('"');
        if (s.StartsWith(@"*\G", StringComparison.OrdinalIgnoreCase)) s = s[3..];
        if (!s.StartsWith('{')) return null;
        int close = s.IndexOf('}');
        if (close < 0 || !Guid.TryParse(s[..(close + 1)], out var guid)) return null;
        var rest = s[(close + 1)..];
        if (!rest.StartsWith('#')) return null;
        var ver = rest[1..].Split('#', ';')[0].Trim();
        int dot = ver.IndexOf('.');
        // Parsed as hex like the registry's TypeLib version keys ("b.5", "a.0" on this machine), through which VB
        // resolves the line. Same as decimal below 10; what VB writes for 10+ isn't verified yet.
        if (!int.TryParse(dot < 0 ? ver : ver[..dot], System.Globalization.NumberStyles.HexNumber, null, out int major)) return null;
        return new ComDependency(guid, major, raw);
    }
}
