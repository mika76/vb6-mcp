using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

/// <summary>Reading .frx blocks (spike 34): every kind seen in the owner's workspace, built byte by byte.</summary>
public static class FrxTests
{
    static byte[] U32(int n) => BitConverter.GetBytes(n);
    static byte[] U16(int n) => BitConverter.GetBytes((ushort)n);
    static byte[] A(string s) => Vb1252.Encoding.GetBytes(s);
    static byte[] Cat(params byte[][] parts) => parts.SelectMany(p => p).ToArray();

    [Test]
    public static void EveryKindOfBlockIsRead()
    {
        var dir = Directory.CreateTempSubdirectory("vb6mcp-frx-").FullName;
        try
        {
            byte[] gif = Cat(A("GIF89a"), [1, 2, 3, 4]);
            byte[] ico = [0, 0, 1, 0, 1, 0, 16, 16];
            var stdPicture = new Guid("0BE35204-8F91-11CE-9DE3-00AA004BB851").ToByteArray();
            var stdFont = new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851").ToByteArray();
            var blocks = new List<(string Line, byte[] Bytes)>();
            int at = 0;
            void Add(string line, byte[] b) { blocks.Add((line.Replace("@", at.ToString("X4")), b)); at += b.Length; }
            Add("   Icon            =   \"F.frx\":@", Cat(U32(8 + gif.Length), A("lt\0\0"), U32(gif.Length), gif));
            Add("      MouseIcon       =   \"F.frx\":@", Cat(U32(16 + 8 + ico.Length), stdPicture, A("lt\0\0"), U32(ico.Length), ico));
            Add("      Caption         =   $\"F.frx\":@", Cat(U32(31), A("Password=x;Pwd=y; a long caption")[..31]));
            Add("      Text            =   \"F.frx\":@", Cat([5], A("Text4")));
            Add("      Text            =   \"F.frx\":@", Cat(U32(9), A("a\r\nb\r\nc\r\n")));
            Add("      List            =   \"F.frx\":@", Cat(U16(2), U16(3), U16(3), A("One"), U16(3), A("Two")));
            Add("      Font            =   \"F.frx\":@", Cat(U32(16 + 10 + 7), stdFont, [1, 0, 0, 0, 0x90, 1, 0, 0, 2, 0], [6], A("Tahoma")));
            Add("      OleObjectBlob   =   \"F.frx\":@", Cat(U32(6), [0x4C, 0x42, 1, 0, 0x0C, 0]));
            // an ImageList image, as a production project's frmMain stores them: a saved StdPicture inside BeginProperty ListImageN
            Add("            Picture         =   \"F.frx\":@", Cat(U32(16 + 8 + gif.Length), stdPicture, A("lt\0\0"), U32(gif.Length), gif));
            File.WriteAllBytes(Path.Combine(dir, "F.frx"), blocks.SelectMany(b => b.Bytes).ToArray());
            var frm = "VERSION 5.00\r\nBegin VB.Form F\r\n" + blocks[0].Line + "\r\n   Begin VB.CommandButton cmd\r\n" +
                      string.Join("\r\n", blocks.Skip(1).Take(7).Select(b => b.Line)) + "\r\n      Index           =   2\r\n   End\r\n" +
                      "   Begin MSComctlLib.ImageList ImageList1\r\n      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628}\r\n         NumListImages   =   1\r\n" +
                      "         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628}\r\n" + blocks[8].Line + "\r\n            Key             =   \"save\"\r\n" +
                      "         EndProperty\r\n      EndProperty\r\n   End\r\nEnd\r\nAttribute VB_Name = \"F\"\r\n";
            File.WriteAllBytes(Path.Combine(dir, "F.frm"), A(frm));

            var entries = FrxReader.Entries(Path.Combine(dir, "F.frm"));
            Assert.Equal(9, entries.Count);
            Assert.Equal("F", entries[0].Control);               // the form itself
            Assert.Equal("cmd(2)", entries[1].Control);          // control array element: Index comes after the .frx lines
            Assert.True(entries[8].Control == "ImageList1" && entries[8].Property == "Images.ListImage1.Picture" && entries[8].Key == "save",
                $"{entries[8].Control}.{entries[8].Property} key {entries[8].Key}");
            Assert.True(FrxReader.Read(entries[8]) is { Kind: "picture", ImageType: "gif" }, "the ImageList image");
            var read = entries.Select(FrxReader.Read).ToList();
            Assert.True(read[0].Kind == "picture" && read[0].ImageType == "gif" && read[0].Data.SequenceEqual(gif), "bare picture, original bytes");
            Assert.True(read[1].Kind == "picture" && read[1].ImageType == "ico", "picture inside a saved StdPicture");
            Assert.Equal("Password=***;Pwd=***; a long captio", FrxReader.MaskSecrets(read[2].Text!));
            Assert.Equal("Text4", read[3].Text);
            Assert.Equal("a\r\nb\r\nc", read[4].Text!.TrimEnd());
            Assert.Equal("One|Two", string.Join("|", read[5].Items!));
            Assert.True(read[6].Kind == "font" && read[6].Text == "Tahoma", "font: " + read[6].Text);
            Assert.True(read[7].Kind == "binary" && read[7].Data.Length == 6, "an ActiveX control's state");
        }
        finally { Directory.Delete(dir, true); }
    }
}
