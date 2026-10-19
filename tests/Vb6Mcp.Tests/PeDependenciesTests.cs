using System.Text;
using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

/// <summary>Import and delay-import reading, and missing DLLs (spike 39).</summary>
public static class PeDependenciesTests
{
    /// <summary>A minimal 32-bit PE: one section with an import descriptor per <paramref name="imports"/> and a delay descriptor per <paramref name="delay"/>.</summary>
    static byte[] MakePe(string[] imports, string[] delay)
    {
        var b = new byte[0x600];
        void I32(int o, int v) => BitConverter.GetBytes(v).CopyTo(b, o);
        void U16(int o, int v) => BitConverter.GetBytes((ushort)v).CopyTo(b, o);
        b[0] = (byte)'M'; b[1] = (byte)'Z'; I32(0x3c, 0x40);
        I32(0x40, 0x4550); U16(0x44, 0x14c); U16(0x46, 1); U16(0x54, 224);
        int opt = 0x58;
        U16(opt, 0x10b); I32(opt + 28, 0x400000);
        I32(opt + 96 + 1 * 8, 0x1000);    // import directory
        I32(opt + 96 + 13 * 8, 0x1100);   // delay-import directory
        int sec = opt + 224;
        I32(sec + 8, 0x400); I32(sec + 12, 0x1000); I32(sec + 16, 0x400); I32(sec + 20, 0x200);
        int names = 0x1200;
        void Name(int descriptorRaw, string n)
        {
            I32(descriptorRaw, names);
            Encoding.ASCII.GetBytes(n).CopyTo(b, names - 0x1000 + 0x200);
            names += 0x40;
        }
        for (int i = 0; i < imports.Length; i++) Name(0x200 + i * 20 + 12, imports[i]);
        for (int i = 0; i < delay.Length; i++) { I32(0x300 + i * 32, 1); Name(0x300 + i * 32 + 4, delay[i]); }
        return b;
    }

    [Test]
    public static void MissingImportsAndDelayImportsAreNamed()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vb6mcp-pe-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var dll = Path.Combine(dir, "server.dll");
            File.WriteAllBytes(dll, MakePe(["kernel32.dll", "nosuch_vb6mcp_a.dll", "api-ms-win-crt-runtime-l1-1-0.dll"], ["nosuch_vb6mcp_b.dll", "nextdoor.dll"]));
            File.WriteAllBytes(Path.Combine(dir, "nextdoor.dll"), MakePe([], []));
            var t = PeDependencies.Read(dll);
            Assert.True(t is not null, "parsed");
            Assert.Equal("kernel32.dll,nosuch_vb6mcp_a.dll,api-ms-win-crt-runtime-l1-1-0.dll", string.Join(",", t!.Value.Imports));
            Assert.Equal("nosuch_vb6mcp_b.dll,nextdoor.dll", string.Join(",", t.Value.Delay));

            var missing = PeDependencies.Missing(dll, [Path.Combine(dir, "app")]);
            Assert.Equal("server.dll>nosuch_vb6mcp_a.dll|server.dll>nosuch_vb6mcp_b.dll|server.dll>nextdoor.dll",
                string.Join("|", missing.Select(m => string.Join(">", m.Chain))));
            Assert.True(!missing[0].Delay && missing[1].Delay, "delay flag");
            Assert.True(missing[2].Note?.Contains("next to server.dll") == true, "a delay import next to the DLL isn't found by the program: " + missing[2].Note);
            Assert.True(PeDependencies.Read(Path.Combine(dir, "nothing.dll")) is null, "no file, no result");
        }
        finally { try { Directory.Delete(dir, true); } catch (IOException) { } }
    }
}
