using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace VideoMaker;

/// <summary>
/// H.264 MP4 through Media Foundation's sink writer (mfplat/mfreadwrite ship with Windows; no ffmpeg). Frames go
/// in as top-down RGB32; the sink writer converts them for the encoder. Only the COM slots this uses are typed:
/// the others are placeholders that keep the vtable order.
/// </summary>
sealed class Mp4Writer : IDisposable
{
    readonly IMFSinkWriter _writer;
    readonly int _stream, _w, _h;
    readonly long _frameDuration;   // 100 ns units
    long _time;

    public Mp4Writer(string path, int width, int height, int fps, int bitrate = 4_000_000)
    {
        _w = width; _h = height; _frameDuration = 10_000_000L / fps;
        Check(MFStartup(0x00020070, 0));
        if (File.Exists(path)) File.Delete(path);
        Check(MFCreateSinkWriterFromURL(path, IntPtr.Zero, IntPtr.Zero, out _writer));

        Check(MFCreateMediaType(out var outType));
        outType.SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video);
        outType.SetGUID(MF_MT_SUBTYPE, MFVideoFormat_H264);
        outType.SetUINT32(MF_MT_AVG_BITRATE, bitrate);
        outType.SetUINT32(MF_MT_INTERLACE_MODE, 2);   // progressive
        outType.SetUINT64(MF_MT_FRAME_SIZE, Pack(width, height));
        outType.SetUINT64(MF_MT_FRAME_RATE, Pack(fps, 1));
        outType.SetUINT64(MF_MT_PIXEL_ASPECT_RATIO, Pack(1, 1));
        Check(_writer.AddStream(outType, out _stream));

        Check(MFCreateMediaType(out var inType));
        inType.SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video);
        inType.SetGUID(MF_MT_SUBTYPE, MFVideoFormat_RGB32);
        inType.SetUINT32(MF_MT_INTERLACE_MODE, 2);
        inType.SetUINT64(MF_MT_FRAME_SIZE, Pack(width, height));
        inType.SetUINT64(MF_MT_FRAME_RATE, Pack(fps, 1));
        inType.SetUINT64(MF_MT_PIXEL_ASPECT_RATIO, Pack(1, 1));
        inType.SetUINT32(MF_MT_DEFAULT_STRIDE, width * 4);   // positive: top-down rows, as LockBits gives them
        Check(_writer.SetInputMediaType(_stream, inType, IntPtr.Zero));
        Check(_writer.BeginWriting());
    }

    public void Write(Bitmap frame)
    {
        int bytes = _w * _h * 4;
        Check(MFCreateMemoryBuffer(bytes, out var buffer));
        Check(buffer.Lock(out var dst, out _, out _));
        var d = frame.LockBits(new Rectangle(0, 0, _w, _h), ImageLockMode.ReadOnly, PixelFormat.Format32bppRgb);
        unsafe { for (int y = 0; y < _h; y++) Buffer.MemoryCopy((byte*)d.Scan0 + y * d.Stride, (byte*)dst + y * _w * 4, _w * 4, _w * 4); }
        frame.UnlockBits(d);
        Check(buffer.Unlock());
        Check(buffer.SetCurrentLength(bytes));
        Check(MFCreateSample(out var sample));
        Check(sample.AddBuffer(buffer));
        Check(sample.SetSampleTime(_time));
        Check(sample.SetSampleDuration(_frameDuration));
        Check(_writer.WriteSample(_stream, sample));
        _time += _frameDuration;
        Marshal.ReleaseComObject(sample);
        Marshal.ReleaseComObject(buffer);
    }

    public void Dispose()
    {
        Check(_writer.Finalize_());
        Marshal.ReleaseComObject(_writer);
        MFShutdown();
    }

    static long Pack(int hi, int lo) => ((long)hi << 32) | (uint)lo;
    static void Check(int hr) { if (hr < 0) Marshal.ThrowExceptionForHR(hr); }

    static readonly Guid MF_MT_MAJOR_TYPE = new("48eba18e-f8c9-4687-bf11-0a74c9f96a8f");
    static readonly Guid MF_MT_SUBTYPE = new("f7e34c9a-42e8-4714-b74b-cb29d72c35e5");
    static readonly Guid MF_MT_AVG_BITRATE = new("20332624-fb0d-4d9e-bd0d-cbf6786c102e");
    static readonly Guid MF_MT_INTERLACE_MODE = new("e2724bb8-e676-4806-b4b2-a8d6efb44ccd");
    static readonly Guid MF_MT_FRAME_SIZE = new("1652c33d-d6b2-4012-b834-72030849a37d");
    static readonly Guid MF_MT_FRAME_RATE = new("c459a2e8-3d2c-4e44-b132-fee5156c7bb0");
    static readonly Guid MF_MT_PIXEL_ASPECT_RATIO = new("c6376a1e-8d0a-4027-be45-6d9a0ad39bb6");
    static readonly Guid MF_MT_DEFAULT_STRIDE = new("644b4e48-1e02-4516-b0eb-c01ca9d49ac6");
    static readonly Guid MFMediaType_Video = new("73646976-0000-0010-8000-00AA00389B71");
    static readonly Guid MFVideoFormat_H264 = new("34363248-0000-0010-8000-00AA00389B71");
    static readonly Guid MFVideoFormat_RGB32 = new("00000016-0000-0010-8000-00AA00389B71");

    [DllImport("mfplat.dll")] static extern int MFStartup(int version, int flags);
    [DllImport("mfplat.dll")] static extern int MFShutdown();
    [DllImport("mfplat.dll")] static extern int MFCreateMediaType(out IMFMediaType type);
    [DllImport("mfplat.dll")] static extern int MFCreateMemoryBuffer(int max, out IMFMediaBuffer buffer);
    [DllImport("mfplat.dll")] static extern int MFCreateSample(out IMFSample sample);
    [DllImport("mfreadwrite.dll", CharSet = CharSet.Unicode)] static extern int MFCreateSinkWriterFromURL(string url, IntPtr byteStream, IntPtr attributes, out IMFSinkWriter writer);

    // IMFAttributes has 30 methods; SetUINT32 (19), SetUINT64 (20) and SetGUID (22) are the ones used.
    [ComImport, Guid("44ae0fa8-ea31-4109-8d2e-4cae4997c555"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMFMediaType
    {
        void a01(); void a02(); void a03(); void a04(); void a05(); void a06(); void a07(); void a08(); void a09(); void a10();
        void a11(); void a12(); void a13(); void a14(); void a15(); void a16(); void a17(); void a18();
        [PreserveSig] int SetUINT32([In] in Guid key, int value);
        [PreserveSig] int SetUINT64([In] in Guid key, long value);
        void a21();
        [PreserveSig] int SetGUID([In] in Guid key, [In] in Guid value);
    }

    [ComImport, Guid("c40a00f2-b93a-4d80-ae8c-5a1c634f58e4"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMFSample
    {
        void a01(); void a02(); void a03(); void a04(); void a05(); void a06(); void a07(); void a08(); void a09(); void a10();
        void a11(); void a12(); void a13(); void a14(); void a15(); void a16(); void a17(); void a18(); void a19(); void a20();
        void a21(); void a22(); void a23(); void a24(); void a25(); void a26(); void a27(); void a28(); void a29(); void a30();
        void s01(); void s02(); void s03();                      // GetSampleFlags, SetSampleFlags, GetSampleTime
        [PreserveSig] int SetSampleTime(long time);
        void s05();                                              // GetSampleDuration
        [PreserveSig] int SetSampleDuration(long duration);
        void s07(); void s08(); void s09();                      // GetBufferCount, GetBufferByIndex, ConvertToContiguousBuffer
        [PreserveSig] int AddBuffer(IMFMediaBuffer buffer);
    }

    [ComImport, Guid("045FA593-8799-42b8-BC8D-8968C6453507"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMFMediaBuffer
    {
        [PreserveSig] int Lock(out IntPtr buffer, out int max, out int current);
        [PreserveSig] int Unlock();
        void b03();                                              // GetCurrentLength
        [PreserveSig] int SetCurrentLength(int length);
    }

    [ComImport, Guid("3137f1cd-fe5e-4805-a5d8-fb477448cb3d"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMFSinkWriter
    {
        [PreserveSig] int AddStream(IMFMediaType type, out int stream);
        [PreserveSig] int SetInputMediaType(int stream, IMFMediaType type, IntPtr encodingParameters);
        [PreserveSig] int BeginWriting();
        [PreserveSig] int WriteSample(int stream, IMFSample sample);
        void w05(); void w06(); void w07(); void w08();          // SendStreamTick, PlaceMarker, NotifyEndOfSegment, Flush
        [PreserveSig] int Finalize_();
    }
}
