using System.Text;
using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

/// <summary>Reading .res files (spike 35), built byte by byte in the layout seen on this machine.</summary>
public static class ResTests
{
    static byte[] Entry(object type, object name, byte[] data, int lang = 0x409)
    {
        static byte[] IdOrName(object v) => v is int id ? [0xFF, 0xFF, (byte)id, (byte)(id >> 8)] : [.. Encoding.Unicode.GetBytes((string)v), 0, 0];
        var head = new List<byte>();
        head.AddRange(IdOrName(type)); head.AddRange(IdOrName(name));
        while ((8 + head.Count) % 4 != 0) head.Add(0);
        head.AddRange(new byte[4]); head.AddRange(BitConverter.GetBytes((ushort)0x1030)); head.AddRange(BitConverter.GetBytes((ushort)lang)); head.AddRange(new byte[8]);
        var e = new List<byte>(); e.AddRange(BitConverter.GetBytes(data.Length)); e.AddRange(BitConverter.GetBytes(8 + head.Count)); e.AddRange(head); e.AddRange(data);
        while (e.Count % 4 != 0) e.Add(0);
        return e.ToArray();
    }

    [Test]
    public static void StringsBitmapsAndIconGroupsAreRead()
    {
        // string table block 1 = ids 0..15: id 1 "Hello", id 3 "World"
        var table = new List<byte>();
        for (int i = 0; i < 16; i++)
        {
            string s = i == 1 ? "Hello" : i == 3 ? "World" : "";
            table.AddRange(BitConverter.GetBytes((ushort)s.Length)); table.AddRange(Encoding.Unicode.GetBytes(s));
        }
        byte[] dib = [.. BitConverter.GetBytes(40), .. BitConverter.GetBytes(2), .. BitConverter.GetBytes(2), 1, 0, 24, 0, .. new byte[24], .. new byte[16]];
        byte[] iconImage = [.. BitConverter.GetBytes(40), .. new byte[36], 9, 9, 9, 9];
        byte[] group = [0, 0, 1, 0, 1, 0, 16, 16, 0, 0, 1, 0, 32, 0, .. BitConverter.GetBytes(iconImage.Length), 7, 0];
        var res = new List<byte>();
        res.AddRange(Entry(0, 0, []));                     // the empty first entry
        res.AddRange(Entry(6, 1, table.ToArray()));
        res.AddRange(Entry(2, 101, dib));
        res.AddRange(Entry(3, 7, iconImage));
        res.AddRange(Entry(14, "MAINICON", group));
        res.AddRange(Entry("CUSTOM", 5, [1, 2, 3]));
        var b = res.ToArray();

        var entries = ResReader.Entries(b);
        Assert.Equal("STRING/1|BITMAP/101|ICON/7|GROUP_ICON/MAINICON|CUSTOM/5", string.Join("|", entries.Select(e => e.Key)));
        Assert.Equal("1:Hello|3:World", string.Join("|", ResReader.Strings(b, entries).Select(s => $"{s.Id}:{s.Text}")));

        var bmp = ResReader.Bitmap(b, entries[1]);
        Assert.True(bmp[0] == 'B' && bmp[1] == 'M' && BitConverter.ToInt32(bmp, 10) == 54 && bmp.Length == 14 + dib.Length, "a .bmp with its file header (24-bit: no palette)");

        var ico = ResReader.IconGroup(b, entries[3], entries);
        Assert.True(ico[2] == 1 && ico[4] == 1 && ico[6] == 16, "an .ico header with one 16x16 entry");
        Assert.Equal(22, BitConverter.ToInt32(ico, 6 + 12));           // the image follows the 6-byte header and one 16-byte entry
        Assert.True(ico.AsSpan(22).SequenceEqual(iconImage), "the ICON resource's bytes");
    }

    [Test]
    public static void AFileThatIsntAResIsRefused()
    {
        try { ResReader.Entries(Encoding.ASCII.GetBytes("this is not a resource file at all")); Assert.True(false, "refused"); }
        catch (InvalidDataException) { }
    }
}
