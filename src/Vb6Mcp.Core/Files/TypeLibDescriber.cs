using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Vb6Mcp.Core.Files;

/// <summary>A parameter of a typelib method, with its type in VB syntax.</summary>
public sealed record TlbParam(string Name, string Type, bool ByRef, bool Optional, bool ParamArray, string? Default);

/// <summary>A method, property accessor, event, enum constant or record field. Kind: "function", "sub", "get", "let", "set", "const" (enum/module constant) or "field".</summary>
public sealed record TlbMember(string Name, string Kind, int MemId, string? ReturnType, IReadOnlyList<TlbParam> Params, string? Value, string? Doc, bool Hidden);

/// <summary>A type in a typelib: coclass, interface, dispinterface, enum, record (UDT), module or alias. Implements: for coclasses, the implemented interfaces, "[default]"/"[source]" marked.</summary>
public sealed record TlbType(string Name, string Kind, Guid Guid, string? Doc, bool Hidden, bool Creatable, IReadOnlyList<string> Implements, IReadOnlyList<TlbMember> Members, string? AliasOf)
{
    /// <summary>An ActiveX control class (TYPEFLAG_FCONTROL): registered like a creatable class, though not flagged creatable (Phase 5: XmText.ocx).</summary>
    public bool Control { get; init; }
}

/// <summary>A whole typelib: identity, file, and every type.</summary>
public sealed record TlbDescription(TypeLibInfo Info, string? Doc, string? HelpFile, IReadOnlyList<TlbType> Types);

/// <summary>
/// Reads the full contents of a typelib (DLL/OCX/EXE/TLB/OLB) through ITypeLib/ITypeInfo, without registering it,
/// and renders types in VB syntax ("Long", "String", "Object", user types by name). Backs com_describe
/// (PLAN.md §9.1: the agent has never seen in-house components) and spikes on VBIDE (VB6EXT.OLB).
/// Runs on its own STA thread like <see cref="TypeLibInfo.Read"/>.
/// </summary>
public static class TypeLibDescriber
{
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern ITypeLib LoadTypeLibEx(string file, int regKind);

    /// <summary>Describes a typelib file; throws <see cref="InvalidOperationException"/> if it has none or times out.</summary>
    public static TlbDescription Describe(string path, TimeSpan? timeout = null)
    {
        TlbDescription? result = null;
        Exception? error = null;
        var t = new Thread(() =>
        {
            // one reader at a time, no ReleaseComObject: typelibs are cached and shared (TypeLibInfo.ComLock)
            lock (TypeLibInfo.ComLock)
            {
                try { result = Read(LoadTypeLibEx(path, 2 /* REGKIND_NONE */)); }
                catch (Exception ex) { error = ex; }
            }
        }) { IsBackground = true, Name = "typelib-describe" };
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(timeout ?? TimeSpan.FromSeconds(30))) throw new InvalidOperationException("reading the typelib timed out");
        return result ?? throw new InvalidOperationException($"no readable typelib in {path}: {error?.Message}");
    }

    /// <summary>
    /// Renders types as compact VB-like declarations, one member per line (what an agent reads best).
    /// <paramref name="filter"/> keeps types whose name contains it (case-insensitive); hidden members are left out unless asked.
    /// </summary>
    public static string Render(TlbDescription d, string? filter = null, bool includeHidden = false)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"' {d.Info.Name} {{{d.Info.LibId.ToString().ToUpperInvariant()}}} v{d.Info.Major}.{d.Info.Minor}{(d.Doc is null ? "" : " - " + d.Doc)}");
        foreach (var t in d.Types.Where(t => filter is null || t.Name.Contains(filter, StringComparison.OrdinalIgnoreCase)))
        {
            if (t.Hidden && !includeHidden && filter is null) continue;
            sb.AppendLine();
            var head = t.Kind switch
            {
                "enum" => $"Enum {t.Name}", "record" => $"Type {t.Name}", "alias" => $"' alias {t.Name} = {t.AliasOf}",
                "coclass" => $"Class {t.Name}{(t.Control ? " ' control" : t.Creatable ? "" : " ' not creatable")}", _ => $"{t.Kind} {t.Name}",
            };
            sb.AppendLine(head + (t.Doc is null ? "" : "  ' " + t.Doc) + (t.Hidden ? "  ' hidden" : ""));
            foreach (var i in t.Implements) sb.AppendLine("    Implements " + i);
            foreach (var m in t.Members.Where(m => includeHidden || !m.Hidden))
            {
                string ps = string.Join(", ", m.Params.Select(p =>
                    (p.Optional ? "Optional " : "") + (p.ParamArray ? "ParamArray " : "") + (p.ByRef ? "" : "ByVal ") + p.Name + " As " + p.Type + (p.Default is null ? "" : " = " + p.Default)));
                string line = m.Kind switch
                {
                    "const" => m.Value is null ? m.Name : $"{m.Name} = {m.Value}",
                    "field" => $"{m.Name} As {m.ReturnType}",
                    "sub" => $"Sub {m.Name}({ps})",
                    "function" => $"Function {m.Name}({ps}) As {m.ReturnType}",
                    "get" => $"Property Get {m.Name}({ps}) As {m.ReturnType}",
                    _ => $"Property {(m.Kind == "let" ? "Let" : "Set")} {m.Name}({ps})",
                };
                sb.AppendLine("    " + line + (m.Doc is null ? "" : "  ' " + m.Doc) + (m.Hidden ? "  ' hidden" : ""));
            }
            if (t.Kind is "enum" or "record") sb.AppendLine($"End {(t.Kind == "enum" ? "Enum" : "Type")}");
        }
        return sb.ToString();
    }

    private static TlbDescription Read(ITypeLib lib)
    {
        lib.GetLibAttr(out var pa);
        TYPELIBATTR la;
        try { la = Marshal.PtrToStructure<TYPELIBATTR>(pa); } finally { lib.ReleaseTLibAttr(pa); }
        lib.GetDocumentation(-1, out var name, out var doc, out _, out var help);
        var types = new List<TlbType>();
        int n = lib.GetTypeInfoCount();
        for (int i = 0; i < n; i++)
        {
            lib.GetTypeInfo(i, out var ti);
            types.Add(ReadType(ti));
        }
        return new TlbDescription(new TypeLibInfo(la.guid, la.wMajorVerNum, la.wMinorVerNum, name), Empty(doc), Empty(help), types);
    }

    private static string? Empty(string? s) => string.IsNullOrEmpty(s) ? null : s;

    private static TlbType ReadType(ITypeInfo ti)
    {
        ti.GetTypeAttr(out var pa);
        try
        {
            var a = Marshal.PtrToStructure<TYPEATTR>(pa);
            ti.GetDocumentation(-1, out var name, out var doc, out _, out _);
            string kind = a.typekind switch
            {
                TYPEKIND.TKIND_COCLASS => "coclass", TYPEKIND.TKIND_DISPATCH => "dispinterface", TYPEKIND.TKIND_INTERFACE => "interface",
                TYPEKIND.TKIND_ENUM => "enum", TYPEKIND.TKIND_RECORD => "record", TYPEKIND.TKIND_MODULE => "module",
                TYPEKIND.TKIND_ALIAS => "alias", TYPEKIND.TKIND_UNION => "union", _ => a.typekind.ToString(),
            };
            var implements = new List<string>();
            if (a.typekind == TYPEKIND.TKIND_COCLASS)
                for (int i = 0; i < a.cImplTypes; i++)
                {
                    ti.GetRefTypeOfImplType(i, out int href);
                    ti.GetImplTypeFlags(i, out var f);
                    var marks = (f.HasFlag(IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) ? "[default] " : "") + (f.HasFlag(IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) ? "[source] " : "");
                    implements.Add(marks + RefName(ti, href));
                }
            var members = new List<TlbMember>();
            for (int i = 0; i < a.cFuncs; i++) members.Add(ReadFunc(ti, i));
            for (int i = 0; i < a.cVars; i++) members.Add(ReadVar(ti, i, a.typekind));
            string? alias = a.typekind == TYPEKIND.TKIND_ALIAS ? TypeName(ti, a.tdescAlias) : null;
            var flags = (TYPEFLAGS)a.wTypeFlags;
            return new TlbType(name, kind, a.guid, Empty(doc), flags.HasFlag(TYPEFLAGS.TYPEFLAG_FHIDDEN),
                a.typekind == TYPEKIND.TKIND_COCLASS && flags.HasFlag(TYPEFLAGS.TYPEFLAG_FCANCREATE), implements, members, alias) { Control = flags.HasFlag(TYPEFLAGS.TYPEFLAG_FCONTROL) };
        }
        finally { ti.ReleaseTypeAttr(pa); }
    }

    private static TlbMember ReadFunc(ITypeInfo ti, int index)
    {
        ti.GetFuncDesc(index, out var pf);
        try
        {
            var f = Marshal.PtrToStructure<FUNCDESC>(pf);
            var names = new string[f.cParams + 1];
            ti.GetNames(f.memid, names, names.Length, out int got);
            ti.GetDocumentation(f.memid, out _, out var doc, out _, out _);
            var ps = new List<TlbParam>();
            int size = Marshal.SizeOf<ELEMDESC>();
            for (int i = 0; i < f.cParams; i++)
            {
                var e = Marshal.PtrToStructure<ELEMDESC>(f.lprgelemdescParam + i * size);
                var pflags = e.desc.paramdesc.wParamFlags;
                bool byRef = (VarEnum)e.tdesc.vt == VarEnum.VT_PTR;
                string? def = null;
                if (pflags.HasFlag(PARAMFLAG.PARAMFLAG_FHASDEFAULT) && e.desc.paramdesc.lpVarValue != IntPtr.Zero)
                {
                    // PARAMDESCEX: cBytes (ULONG, padded to 8) then the VARIANT
                    try { def = Marshal.GetObjectForNativeVariant(e.desc.paramdesc.lpVarValue + 8) is { } v ? Literal(v) : null; } catch (ArgumentException) { }
                }
                bool retval = pflags.HasFlag(PARAMFLAG.PARAMFLAG_FRETVAL);
                if (retval) continue;   // shown as the return type
                ps.Add(new TlbParam(i + 1 < got ? names[i + 1] : $"p{i + 1}", TypeName(ti, byRef ? Deref(e.tdesc) : e.tdesc), byRef,
                    pflags.HasFlag(PARAMFLAG.PARAMFLAG_FOPT), false, def));
            }
            if (f.cParamsOpt == -1 && ps.Count > 0) ps[^1] = ps[^1] with { ParamArray = true };
            // return type: [retval] parameter (interfaces) or the function's own type (dispinterfaces)
            string? ret = null;
            for (int i = 0; i < f.cParams; i++)
            {
                var e = Marshal.PtrToStructure<ELEMDESC>(f.lprgelemdescParam + i * size);
                if (e.desc.paramdesc.wParamFlags.HasFlag(PARAMFLAG.PARAMFLAG_FRETVAL)) ret = TypeName(ti, Deref(e.tdesc));
            }
            var own = (VarEnum)f.elemdescFunc.tdesc.vt;
            if (ret is null && own is not (VarEnum.VT_VOID or VarEnum.VT_HRESULT)) ret = TypeName(ti, f.elemdescFunc.tdesc);
            string kind = f.invkind switch
            {
                INVOKEKIND.INVOKE_PROPERTYGET => "get", INVOKEKIND.INVOKE_PROPERTYPUT => "let", INVOKEKIND.INVOKE_PROPERTYPUTREF => "set",
                _ => ret is null ? "sub" : "function",
            };
            bool hidden = ((FUNCFLAGS)f.wFuncFlags).HasFlag(FUNCFLAGS.FUNCFLAG_FHIDDEN) || ((FUNCFLAGS)f.wFuncFlags).HasFlag(FUNCFLAGS.FUNCFLAG_FRESTRICTED);
            return new TlbMember(names[0], kind, f.memid, ret, ps, null, Empty(doc), hidden);
        }
        finally { ti.ReleaseFuncDesc(pf); }
    }

    private static TlbMember ReadVar(ITypeInfo ti, int index, TYPEKIND owner)
    {
        ti.GetVarDesc(index, out var pv);
        try
        {
            var v = Marshal.PtrToStructure<VARDESC>(pv);
            var names = new string[1];
            ti.GetNames(v.memid, names, 1, out _);
            ti.GetDocumentation(v.memid, out _, out var doc, out _, out _);
            string? value = null;
            if (v.varkind == VARKIND.VAR_CONST && v.desc.lpvarValue != IntPtr.Zero)
                try { value = Marshal.GetObjectForNativeVariant(v.desc.lpvarValue) is { } o ? Literal(o) : null; } catch (ArgumentException) { }
            string kind = v.varkind == VARKIND.VAR_CONST ? "const" : "field";
            string? type = owner == TYPEKIND.TKIND_ENUM ? null : TypeName(ti, v.elemdescVar.tdesc);
            return new TlbMember(names[0], kind, v.memid, type, [], value, Empty(doc), ((VARFLAGS)v.wVarFlags).HasFlag(VARFLAGS.VARFLAG_FHIDDEN));
        }
        finally { ti.ReleaseVarDesc(pv); }
    }

    private static string Literal(object v) => v switch
    {
        string s => "\"" + s.Replace("\"", "\"\"") + "\"",
        bool b => b ? "True" : "False",
        IFormattable f => f.ToString(null, System.Globalization.CultureInfo.InvariantCulture),
        _ => v.ToString() ?? "",
    };

    private static TYPEDESC Deref(TYPEDESC t) => (VarEnum)t.vt == VarEnum.VT_PTR ? Marshal.PtrToStructure<TYPEDESC>(t.lpValue) : t;

    private static string RefName(ITypeInfo ti, int href)
    {
        try
        {
            ti.GetRefTypeInfo(href, out var r);
            r.GetDocumentation(-1, out var n, out _, out _, out _);
            return n;
        }
        catch (COMException) { return "?"; }   // type from a typelib that isn't registered
    }

    /// <summary>A TYPEDESC in VB syntax.</summary>
    private static string TypeName(ITypeInfo ti, TYPEDESC t) => (VarEnum)t.vt switch
    {
        VarEnum.VT_I2 => "Integer", VarEnum.VT_I4 or VarEnum.VT_INT => "Long", VarEnum.VT_R4 => "Single", VarEnum.VT_R8 => "Double",
        VarEnum.VT_CY => "Currency", VarEnum.VT_DATE => "Date", VarEnum.VT_BSTR or VarEnum.VT_LPSTR or VarEnum.VT_LPWSTR => "String",
        VarEnum.VT_DISPATCH => "Object", VarEnum.VT_ERROR => "SCODE", VarEnum.VT_BOOL => "Boolean", VarEnum.VT_VARIANT => "Variant",
        VarEnum.VT_UNKNOWN => "IUnknown", VarEnum.VT_DECIMAL => "Decimal", VarEnum.VT_UI1 => "Byte", VarEnum.VT_I1 => "SByte",
        VarEnum.VT_UI2 => "UInteger", VarEnum.VT_UI4 or VarEnum.VT_UINT => "ULong", VarEnum.VT_I8 => "LongLong", VarEnum.VT_UI8 => "ULongLong",
        VarEnum.VT_VOID => "Any", VarEnum.VT_HRESULT => "HRESULT",
        VarEnum.VT_PTR => TypeName(ti, Marshal.PtrToStructure<TYPEDESC>(t.lpValue)),
        VarEnum.VT_SAFEARRAY => TypeName(ti, Marshal.PtrToStructure<TYPEDESC>(t.lpValue)) + "()",
        VarEnum.VT_CARRAY => TypeName(ti, Marshal.PtrToStructure<TYPEDESC>(t.lpValue)) + "()",   // ARRAYDESC starts with the element TYPEDESC
        VarEnum.VT_USERDEFINED => RefName(ti, unchecked((int)(long)t.lpValue)),
        var other => other.ToString(),
    };
}
