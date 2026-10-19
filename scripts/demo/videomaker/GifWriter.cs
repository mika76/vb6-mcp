using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace VideoMaker;

/// <summary>
/// Animated GIF that stores only what changed: each frame is the bounding box of the pixels that differ from the
/// previous one (identical frames just lengthen the previous delay), with its own palette, LZW, looping forever.
/// That keeps a smooth UI animation small, since most of the screen stays still.
/// </summary>
sealed class GifWriter : IDisposable
{
    readonly BinaryWriter _o;
    readonly int _w, _h;
    int[]? _prev;
    (int[] Px, Rectangle Box)? _pending;
    int _pendingDelay;

    public GifWriter(string path, int w, int h)
    {
        _w = w; _h = h;
        _o = new BinaryWriter(File.Create(path));
        _o.Write("GIF89a"u8.ToArray());
        _o.Write((ushort)w); _o.Write((ushort)h); _o.Write((byte)0x70); _o.Write((byte)0); _o.Write((byte)0);
        _o.Write(new byte[] { 0x21, 0xFF, 0x0B }); _o.Write("NETSCAPE2.0"u8.ToArray()); _o.Write(new byte[] { 3, 1, 0, 0, 0 });
    }

    /// <summary>Adds a frame shown for delayCs hundredths of a second.</summary>
    public void Add(Bitmap bmp, int delayCs)
    {
        var px = Pixels(bmp);
        var box = _prev is null ? new Rectangle(0, 0, _w, _h) : Diff(_prev, px);
        if (box.IsEmpty) { _pendingDelay += delayCs; return; }
        Flush();
        _pending = (px, box);
        _pendingDelay = delayCs;
        _prev = px;
    }

    void Flush()
    {
        if (_pending is not { } p) return;
        var (pal, idx) = Quantize(p.Px, p.Box);
        _o.Write(new byte[] { 0x21, 0xF9, 4, 0x04 });            // disposal: leave in place
        _o.Write((ushort)Math.Max(2, _pendingDelay)); _o.Write((byte)0); _o.Write((byte)0);
        _o.Write((byte)0x2C); _o.Write((ushort)p.Box.X); _o.Write((ushort)p.Box.Y); _o.Write((ushort)p.Box.Width); _o.Write((ushort)p.Box.Height);
        _o.Write((byte)(0x80 | 7));
        for (int i = 0; i < 256; i++) { var c = i < pal.Count ? pal[i] : 0; _o.Write((byte)(c >> 16)); _o.Write((byte)(c >> 8)); _o.Write((byte)c); }
        _o.Write((byte)8);
        Lzw(_o, idx);
        _pending = null;
    }

    public void Dispose() { Flush(); _o.Write((byte)0x3B); _o.Dispose(); }

    int[] Pixels(Bitmap bmp)
    {
        var d = bmp.LockBits(new Rectangle(0, 0, _w, _h), ImageLockMode.ReadOnly, PixelFormat.Format32bppRgb);
        var px = new int[_w * _h];
        for (int y = 0; y < _h; y++) Marshal.Copy(d.Scan0 + y * d.Stride, px, y * _w, _w);
        bmp.UnlockBits(d);
        for (int i = 0; i < px.Length; i++) px[i] &= 0xFFFFFF;
        return px;
    }

    Rectangle Diff(int[] a, int[] b)
    {
        int x0 = _w, y0 = _h, x1 = -1, y1 = -1;
        for (int y = 0; y < _h; y++)
            for (int x = 0, i = y * _w; x < _w; x++, i++)
                if (a[i] != b[i]) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
        return x1 < 0 ? Rectangle.Empty : new Rectangle(x0, y0, x1 - x0 + 1, y1 - y0 + 1);
    }

    (List<int> Pal, byte[] Idx) Quantize(int[] px, Rectangle box)
    {
        var count = new Dictionary<int, int>();
        for (int y = box.Top; y < box.Bottom; y++)
            for (int x = box.Left; x < box.Right; x++) { int c = px[y * _w + x]; count[c] = count.GetValueOrDefault(c) + 1; }
        var pal = count.OrderByDescending(kv => kv.Value).Take(256).Select(kv => kv.Key).ToList();
        var map = new Dictionary<int, byte>();
        for (int i = 0; i < pal.Count; i++) map[pal[i]] = (byte)i;
        var idx = new byte[box.Width * box.Height];
        int k = 0;
        for (int y = box.Top; y < box.Bottom; y++)
            for (int x = box.Left; x < box.Right; x++)
            {
                int c = px[y * _w + x];
                if (!map.TryGetValue(c, out var v)) { v = Nearest(pal, c); map[c] = v; }
                idx[k++] = v;
            }
        return (pal, idx);
    }

    static byte Nearest(List<int> pal, int c)
    {
        int best = 0, bd = int.MaxValue;
        for (int i = 0; i < pal.Count; i++)
        {
            int dr = ((pal[i] >> 16) & 255) - ((c >> 16) & 255), dg = ((pal[i] >> 8) & 255) - ((c >> 8) & 255), db = (pal[i] & 255) - (c & 255);
            int dd = 2 * dr * dr + 4 * dg * dg + db * db;
            if (dd < bd) { bd = dd; best = i; }
        }
        return (byte)best;
    }

    static void Lzw(BinaryWriter o, byte[] data)
    {
        const int clear = 256, end = 257;
        var block = new List<byte>(255);
        int bitBuf = 0, bitCount = 0, width = 9, next = 258;
        var dict = new Dictionary<int, int>();
        void Emit(int code)
        {
            bitBuf |= code << bitCount; bitCount += width;
            while (bitCount >= 8)
            {
                block.Add((byte)bitBuf); bitBuf >>= 8; bitCount -= 8;
                if (block.Count == 255) { o.Write((byte)255); o.Write(block.ToArray()); block.Clear(); }
            }
        }
        Emit(clear);
        int prefix = data[0];
        for (int i = 1; i < data.Length; i++)
        {
            int k = data[i], key = (prefix << 8) | k;
            if (dict.TryGetValue(key, out var code)) { prefix = code; continue; }
            Emit(prefix);
            if (next < 4096) { dict[key] = next++; if (next > (1 << width) && width < 12) width++; }
            else { Emit(clear); dict.Clear(); next = 258; width = 9; }
            prefix = k;
        }
        Emit(prefix);
        Emit(end);
        if (bitCount > 0) block.Add((byte)bitBuf);
        if (block.Count > 0) { o.Write((byte)block.Count); o.Write(block.ToArray()); }
        o.Write((byte)0);
    }
}
