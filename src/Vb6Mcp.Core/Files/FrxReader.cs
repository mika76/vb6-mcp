using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Vb6Mcp.Core.Files;

/// <summary>
/// A designer property stored in the companion binary file: which control, which property (with its property group, e.g.
/// Images.ListImage1.Picture for an ImageList image), where, and the group's Key when it has one (ImageList images do).
/// </summary>
public sealed record FrxEntry(string Control, string Property, string File, int Offset, bool LongString, string? Key = null);

/// <summary>What a block holds: a picture (original bytes and type), list items, a long string, or other binary data.</summary>
public sealed record FrxBlock(string Kind, byte[] Data, string? ImageType, IReadOnlyList<string>? Items, string? Text);

/// <summary>
/// Reads the .frx/.ctx/.pgx/.dsx/.dox companions of VB6 designer files (spike 34), read-only (hard rule 2: never edited).
/// The .frm points at byte offsets ("frmMain.frx":08CA); a picture block is u32 length, "lt\0\0", u32 image length and the
/// original image (a GIF stays a GIF, unlike VBIDE's SavePicture, which writes BMP); a long string ($"...") is u32 length
/// and the text; List/ItemData are u16 count, u16, then u16-length strings.
/// </summary>
public static class FrxReader
{
    static readonly Regex Ref = new(@"^\s*(\w+(?:\(\d+\))?)\s*=\s*(\$?)""([^""]+\.(?:frx|ctx|pgx|dsx|dox))"":([0-9A-Fa-f]+)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    static readonly Regex Begin = new(@"^\s*Begin\s+[\w.]+\s+(\w+)", RegexOptions.CultureInvariant);
    static readonly Regex Index = new(@"^\s*Index\s*=\s*(\d+)", RegexOptions.CultureInvariant);

    /// <summary>The companion-file properties of a designer file (.frm/.ctl/.pag/.dsr/.dob), in file order.</summary>
    public static IReadOnlyList<FrxEntry> Entries(string designerPath)
    {
        var text = Vb1252.Encoding.GetString(File.ReadAllBytes(designerPath));
        var dir = Path.GetDirectoryName(Path.GetFullPath(designerPath))!;
        var result = new List<(FrxEntry Entry, int Depth)>();
        var stack = new List<string>();
        var groups = new List<string>();   // BeginProperty <name> ... EndProperty (Font, ImageList Images/ListImageN) inside a control
        int propertyGroups = 0;
        var keyOfGroup = new Dictionary<string, string>();   // "control|group path" -> Key = "..." inside that group
        foreach (var raw in text.Replace("\r\n", "\n").Split('\n'))
        {
            var line = raw.TrimEnd();
            if (line.StartsWith("Attribute VB_Name", StringComparison.Ordinal)) break;   // end of the designer block
            var t = line.TrimStart();
            if (t.StartsWith("BeginProperty", StringComparison.Ordinal))
            {
                propertyGroups++;
                groups.Add(t.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries) is { Length: > 1 } p ? p[1] : "?");
                continue;
            }
            if (t.StartsWith("EndProperty", StringComparison.Ordinal)) { propertyGroups = Math.Max(0, propertyGroups - 1); if (groups.Count > 0) groups.RemoveAt(groups.Count - 1); continue; }
            if (groups.Count > 0 && stack.Count > 0 && KeyLine.Match(line) is { Success: true } k) { keyOfGroup[stack[^1] + "|" + string.Join(".", groups)] = k.Groups[1].Value; continue; }
            if (Begin.Match(line) is { Success: true } b) { stack.Add(b.Groups[1].Value); continue; }
            if (t == "End") { if (stack.Count > 0) stack.RemoveAt(stack.Count - 1); continue; }
            // a control array element: the Index line names which one (it comes after Begin, often after the .frx lines)
            if (propertyGroups == 0 && stack.Count > 0 && Index.Match(line) is { Success: true } ix)
            {
                for (int i = 0; i < result.Count; i++)
                    if (result[i].Depth == stack.Count && result[i].Entry.Control == stack[^1])
                        result[i] = (result[i].Entry with { Control = $"{stack[^1]}({ix.Groups[1].Value})" }, result[i].Depth);
                stack[^1] = $"{stack[^1]}({ix.Groups[1].Value})";
                continue;
            }
            if (Ref.Match(line) is { Success: true } m)
                result.Add((new FrxEntry(stack.Count > 0 ? stack[^1] : "", string.Join(".", groups.Append(m.Groups[1].Value)), Path.Combine(dir, m.Groups[3].Value),
                    int.Parse(m.Groups[4].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture), m.Groups[2].Value == "$"), stack.Count));
        }
        // the Key of the group a property sits in (ImageList images: ListImage3's Key = "save" comes after its Picture line)
        return result.Select(r =>
        {
            var e = r.Entry;
            int dot = e.Property.LastIndexOf('.');
            return dot > 0 && keyOfGroup.TryGetValue(e.Control + "|" + e.Property[..dot], out var key) && key.Length > 0 ? e with { Key = key } : e;
        }).ToList();
    }

    static readonly Regex KeyLine = new(@"^\s*Key\s*=\s*""([^""]*)""", RegexOptions.CultureInvariant);

    /// <summary>Reads one block. Throws InvalidDataException when the offset or a length doesn't fit the file.</summary>
    public static FrxBlock Read(FrxEntry e)
    {
        var bytes = File.ReadAllBytes(e.File);
        int Need(int at, int n) => at >= 0 && n >= 0 && at + n <= bytes.Length ? at : throw new InvalidDataException($"{Path.GetFileName(e.File)}: block at {e.Offset:X4} runs past the end of the file");
        int U32(int at) => BitConverter.ToInt32(bytes, Need(at, 4));
        int U16(int at) => BitConverter.ToUInt16(bytes, Need(at, 2));

        if (e.LongString)
        {
            int len = U32(e.Offset);
            return new FrxBlock("text", [], null, null, Vb1252.Encoding.GetString(bytes, Need(e.Offset + 4, len), len));
        }
        var own = e.Property[(e.Property.LastIndexOf('.') + 1)..];   // the property itself, without its group
        if (own.StartsWith("List", StringComparison.OrdinalIgnoreCase) || own.StartsWith("ItemData", StringComparison.OrdinalIgnoreCase))
        {
            int count = U16(e.Offset), at = e.Offset + 4;
            var items = new List<string>(count);
            for (int i = 0; i < count; i++)
            {
                int len = U16(at);
                items.Add(Vb1252.Encoding.GetString(bytes, Need(at + 2, len), len));
                at += 2 + len;
            }
            return new FrxBlock("list", [], null, items, null);
        }
        bool Lt(int at) => at + 4 <= bytes.Length && bytes[at] == (byte)'l' && bytes[at + 1] == (byte)'t' && bytes[at + 2] == 0 && bytes[at + 3] == 0;
        bool Clsid(int at, byte[] id) => at + 16 <= bytes.Length && bytes.AsSpan(at, 16).SequenceEqual(id);
        FrxBlock Picture(int lenAt)
        {
            int len = U32(lenAt);
            var image = bytes.AsSpan(Need(lenAt + 4, len), len).ToArray();
            return new FrxBlock(len == 0 ? "empty picture" : "picture", image, len == 0 ? null : ImageType(image), null, null);
        }
        // a bare picture: u32 block, "lt\0\0", u32 image length, image
        if (Lt(e.Offset + 4)) return Picture(e.Offset + 8);
        // a saved StdPicture object (MouseIcon, CellPicture, Image...): u32 block, its CLSID, then the same "lt" block
        if (Clsid(e.Offset + 4, StdPicture) && Lt(e.Offset + 20)) return Picture(e.Offset + 24);
        int block = U32(e.Offset);
        // a saved StdFont object (u32 block, CLSID, font data): kept as data; the font is better read through form_properties
        if (Clsid(e.Offset + 4, StdFont))
            return new FrxBlock("font", bytes.AsSpan(Need(e.Offset + 4, block), Math.Min(block, bytes.Length - e.Offset - 4)).ToArray(), null, null, FontName(bytes, e.Offset));
        // text: u32 length + text (multi-line Text, long Tag...), or a u8 length + text (short Text)
        if (block > 0 && e.Offset + 4 + block <= bytes.Length && Printable(bytes, e.Offset + 4, block))
            return new FrxBlock("text", [], null, null, Vb1252.Encoding.GetString(bytes, e.Offset + 4, block));
        int shortLen = bytes[Need(e.Offset, 1)];
        if (shortLen > 0 && e.Offset + 1 + shortLen <= bytes.Length && Printable(bytes, e.Offset + 1, shortLen))
            return new FrxBlock("text", [], null, null, Vb1252.Encoding.GetString(bytes, e.Offset + 1, shortLen));
        // anything else: an ActiveX control's own saved state (OleObjectBlob, MessagesAndKeys, PropertyPages...)
        int n = block > 0 && e.Offset + 4 + block <= bytes.Length ? block : bytes.Length - e.Offset - 4;
        return new FrxBlock("binary", bytes.AsSpan(Need(e.Offset + 4, n), n).ToArray(), null, null, null);
    }

    // CLSIDs as stored (little-endian GUID bytes): StdPicture {0BE35204-8F91-11CE-9DE3-00AA004BB851}, StdFont {0BE35203-...}
    static readonly byte[] StdPicture = new Guid("0BE35204-8F91-11CE-9DE3-00AA004BB851").ToByteArray();
    static readonly byte[] StdFont = new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851").ToByteArray();

    static bool Printable(byte[] b, int at, int n)
    {
        for (int i = at; i < at + n; i++)
            if (b[i] < 0x20 && b[i] is not ((byte)'\r' or (byte)'\n' or (byte)'\t')) return false;
        return true;
    }

    /// <summary>The face name in a saved StdFont (u8 length + name after a fixed header), or null.</summary>
    static string? FontName(byte[] b, int offset)
    {
        int at = offset + 4 + 16 + 10;   // block length, CLSID, version/charset/flags/weight/size
        return at < b.Length && b[at] > 0 && at + 1 + b[at] <= b.Length && Printable(b, at + 1, b[at]) ? Vb1252.Encoding.GetString(b, at + 1, b[at]) : null;
    }

    /// <summary>The image format from its first bytes: gif, jpg, bmp, ico, cur, wmf, emf, png, or null.</summary>
    public static string? ImageType(byte[] b)
    {
        bool At(int i, params byte[] sig) => b.Length >= i + sig.Length && sig.Select((x, k) => b[i + k] == x).All(x => x);
        if (At(0, (byte)'G', (byte)'I', (byte)'F', (byte)'8')) return "gif";
        if (At(0, 0xFF, 0xD8)) return "jpg";
        if (At(0, (byte)'B', (byte)'M')) return "bmp";
        if (At(0, 0, 0, 1, 0)) return "ico";
        if (At(0, 0, 0, 2, 0)) return "cur";
        if (At(0, 0xD7, 0xCD, 0xC6, 0x9A)) return "wmf";
        if (At(0, 0x89, (byte)'P', (byte)'N', (byte)'G')) return "png";
        if (At(40, (byte)' ', (byte)'E', (byte)'M', (byte)'F')) return "emf";
        return null;
    }

    /// <summary>Connection-string passwords masked (Password=, Pwd=), for showing long strings (spike 34).</summary>
    public static string MaskSecrets(string text)
        => Regex.Replace(text, @"(?i)\b(password|pwd)\s*=\s*(""[^""]*""|[^;]*)", m => m.Groups[1].Value + "=***");
}
