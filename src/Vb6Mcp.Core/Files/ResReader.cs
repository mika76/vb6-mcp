using System.Text;

namespace Vb6Mcp.Core.Files;

/// <summary>One resource of a .res file: type and name (a number or a string), language, and where its data is.</summary>
public sealed record ResEntry(string Type, int? TypeId, string Name, int? NameId, int Language, int Offset, int Size)
{
    /// <summary>"TYPE/name" as tools show and accept it, e.g. BITMAP/101, GROUP_ICON/MAINICON, CUSTOM/7.</summary>
    public string Key => $"{Type}/{Name}";
}

/// <summary>
/// Reads 32-bit .res files (VB6's ResFile32, made with the VB Resource Editor or rc.exe), read-only (spike 35). The layout was
/// checked on .res files on this machine (a production workspace, the VB98 samples): an empty first entry, then per resource
/// u32 data size, u32 header size, type and name (0xFFFF + u16 ordinal, or a null-terminated UTF-16 string), DWORD padding,
/// u32 data version, u16 memory flags, u16 language, u32 version, u32 characteristics, the data, DWORD padding.
/// </summary>
public static class ResReader
{
    static readonly Dictionary<int, string> TypeNames = new()
    {
        [1] = "CURSOR", [2] = "BITMAP", [3] = "ICON", [4] = "MENU", [5] = "DIALOG", [6] = "STRING", [7] = "FONTDIR", [8] = "FONT",
        [9] = "ACCELERATOR", [10] = "RCDATA", [11] = "MESSAGETABLE", [12] = "GROUP_CURSOR", [14] = "GROUP_ICON", [16] = "VERSION",
        [17] = "DLGINCLUDE", [19] = "PLUGPLAY", [20] = "VXD", [21] = "ANICURSOR", [22] = "ANIICON", [23] = "HTML", [24] = "MANIFEST",
    };

    public static IReadOnlyList<ResEntry> Entries(byte[] b)
    {
        // a 32-bit .res starts with an empty 32-byte entry (both files checked here do); 16-bit ones don't and aren't read
        if (b.Length < 32 || BitConverter.ToInt32(b, 0) != 0 || BitConverter.ToInt32(b, 4) != 32)
            throw new InvalidDataException("not a 32-bit .res file (it doesn't start with the empty 32-byte entry)");
        try { return ReadEntries(b); }
        catch (ArgumentException) { throw new InvalidDataException("the .res file is damaged (an entry runs past the end)"); }
    }

    static IReadOnlyList<ResEntry> ReadEntries(byte[] b)
    {
        var list = new List<ResEntry>();
        int at = 0;
        while (at + 8 <= b.Length)
        {
            int size = BitConverter.ToInt32(b, at), header = BitConverter.ToInt32(b, at + 4);
            if (size < 0 || header < 16 || at + header + size > b.Length) throw new InvalidDataException($"not a 32-bit .res file, or damaged at byte {at}");
            int p = at + 8;
            var (typeName, typeId) = IdOrName(b, ref p);
            var (name, nameId) = IdOrName(b, ref p);
            p = (p + 3) & ~3;
            int lang = BitConverter.ToUInt16(b, p + 6);
            if (!(size == 0 && typeId == 0 && nameId == 0))   // the empty first entry that marks a 32-bit .res
                list.Add(new ResEntry(typeId is int t ? TypeNames.GetValueOrDefault(t, "#" + t) : typeName!, typeId, nameId?.ToString() ?? name!, nameId, lang, at + header, size));
            at = (at + header + size + 3) & ~3;
        }
        return list;
    }

    static (string? Name, int? Id) IdOrName(byte[] b, ref int p)
    {
        if (BitConverter.ToUInt16(b, p) == 0xFFFF) { int id = BitConverter.ToUInt16(b, p + 2); p += 4; return (null, id); }
        var sb = new StringBuilder();
        while (p + 1 < b.Length && BitConverter.ToUInt16(b, p) != 0) { sb.Append((char)BitConverter.ToUInt16(b, p)); p += 2; }
        p += 2;
        return (sb.ToString(), null);
    }

    /// <summary>All strings of the string tables: block n holds ids (n-1)*16 .. (n-1)*16+15, each a u16 length and UTF-16 text.</summary>
    public static IReadOnlyList<(int Id, int Language, string Text)> Strings(byte[] b, IReadOnlyList<ResEntry> entries)
    {
        var result = new List<(int, int, string)>();
        foreach (var e in entries.Where(e => e.TypeId == 6 && e.NameId is int))
        {
            int p = e.Offset, end = e.Offset + e.Size;
            for (int i = 0; i < 16 && p + 2 <= end; i++)
            {
                int len = BitConverter.ToUInt16(b, p); p += 2;
                if (len > 0 && p + len * 2 <= end) result.Add(((e.NameId!.Value - 1) * 16 + i, e.Language, Encoding.Unicode.GetString(b, p, len * 2)));
                p += len * 2;
            }
        }
        return result.OrderBy(r => r.Item1).ToList();
    }

    /// <summary>A BITMAP resource as a .bmp file: the resource is a DIB without the 14-byte file header.</summary>
    public static byte[] Bitmap(byte[] b, ResEntry e)
    {
        int infoSize = BitConverter.ToInt32(b, e.Offset), bitCount = BitConverter.ToUInt16(b, e.Offset + 14);
        int compression = infoSize >= 40 ? BitConverter.ToInt32(b, e.Offset + 16) : 0, used = infoSize >= 40 ? BitConverter.ToInt32(b, e.Offset + 32) : 0;
        int palette = (used > 0 ? used : bitCount <= 8 ? 1 << bitCount : 0) * 4 + (compression == 3 && infoSize == 40 ? 12 : 0);
        var file = new byte[14 + e.Size];
        file[0] = (byte)'B'; file[1] = (byte)'M';
        BitConverter.GetBytes(file.Length).CopyTo(file, 2);
        BitConverter.GetBytes(14 + infoSize + palette).CopyTo(file, 10);
        Array.Copy(b, e.Offset, file, 14, e.Size);
        return file;
    }

    /// <summary>
    /// A GROUP_ICON (or GROUP_CURSOR) as an .ico (.cur) file: the group lists its images by resource id; the file lists them by
    /// offset. Cursor images start with their hotspot, which the .cur entry keeps instead of planes/bit count.
    /// </summary>
    public static byte[] IconGroup(byte[] b, ResEntry group, IReadOnlyList<ResEntry> entries)
    {
        bool cursor = group.TypeId == 12;
        int count = BitConverter.ToUInt16(b, group.Offset + 4);
        var images = new List<(byte[] Entry, byte[] Data)>();
        for (int i = 0; i < count; i++)
        {
            int p = group.Offset + 6 + i * 14;
            int id = BitConverter.ToUInt16(b, p + 12);
            var img = entries.FirstOrDefault(x => x.TypeId == (cursor ? 1 : 3) && x.NameId == id) ?? throw new InvalidDataException($"{group.Key} refers to a missing image {id}");
            var data = b.AsSpan(img.Offset + (cursor ? 4 : 0), img.Size - (cursor ? 4 : 0)).ToArray();
            var entry = new byte[16];
            if (cursor)
            {
                int w = BitConverter.ToUInt16(b, p), h = BitConverter.ToUInt16(b, p + 2) / 2;   // a cursor group gives width/height as words, height doubled
                entry[0] = (byte)w; entry[1] = (byte)h;
                BitConverter.GetBytes(BitConverter.ToUInt16(b, img.Offset)).CopyTo(entry, 4);       // hotspot x
                BitConverter.GetBytes(BitConverter.ToUInt16(b, img.Offset + 2)).CopyTo(entry, 6);   // hotspot y
            }
            else Array.Copy(b, p, entry, 0, 8);   // width, height, colours, reserved, planes, bit count
            BitConverter.GetBytes(data.Length).CopyTo(entry, 8);
            images.Add((entry, data));
        }
        using var ms = new MemoryStream();
        ms.Write([0, 0, (byte)(cursor ? 2 : 1), 0, (byte)count, (byte)(count >> 8)]);
        int offset = 6 + 16 * count;
        foreach (var (entry, data) in images) { BitConverter.GetBytes(offset).CopyTo(entry, 12); ms.Write(entry); offset += data.Length; }
        foreach (var (_, data) in images) ms.Write(data);
        return ms.ToArray();
    }

    /// <summary>The VERSION resource: file and product version numbers and the string values (CompanyName, FileDescription...).</summary>
    public static (string? FileVersion, string? ProductVersion, IReadOnlyDictionary<string, string> Strings) Version(byte[] b, ResEntry e)
    {
        var strings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string? fileVersion = null, productVersion = null;
        void Block(int at, int depth)
        {
            int len = BitConverter.ToUInt16(b, at), valueLen = BitConverter.ToUInt16(b, at + 2), type = BitConverter.ToUInt16(b, at + 4);
            if (len == 0 || at + len > e.Offset + e.Size) return;
            int p = at + 6; var key = new StringBuilder();
            while (BitConverter.ToUInt16(b, p) != 0) { key.Append((char)BitConverter.ToUInt16(b, p)); p += 2; }
            p = (p + 2 + 3) & ~3;
            int valueBytes = type == 1 ? valueLen * 2 : valueLen;
            if (depth == 0 && valueBytes >= 52 && BitConverter.ToUInt32(b, p) == 0xFEEF04BD)
            {
                string V(int o) => $"{BitConverter.ToUInt16(b, p + o + 2)}.{BitConverter.ToUInt16(b, p + o)}.{BitConverter.ToUInt16(b, p + o + 6)}.{BitConverter.ToUInt16(b, p + o + 4)}";
                fileVersion = V(8); productVersion = V(16);
            }
            else if (type == 1 && depth == 3 && valueLen > 0)
                strings[key.ToString()] = Encoding.Unicode.GetString(b, p, Math.Max(0, valueLen - 1) * 2);
            p = (p + valueBytes + 3) & ~3;
            while (p + 6 <= at + len) { int child = BitConverter.ToUInt16(b, p); if (child == 0) break; Block(p, depth + 1); p = (p + child + 3) & ~3; }
        }
        Block(e.Offset, 0);
        return (fileVersion, productVersion, strings);
    }
}
