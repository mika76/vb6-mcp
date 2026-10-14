using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace Vb6Mcp.UI;

/// <summary>One OutputDebugString message.</summary>
public sealed record DebugMessage(long Id, DateTime Time, int Pid, string Process, string Text);

/// <summary>
/// Captures OutputDebugString output of programs in this Windows session through the DBWIN shared
/// buffer: the DBWIN_BUFFER mapping plus the DBWIN_BUFFER_READY / DBWIN_DATA_READY events, the
/// protocol DebugView uses. Compiled VB6 programs have no Debug.Print (it is compiled out), so
/// programs that log usually call OutputDebugString. Only one monitor can own the buffer at a time:
/// if the events already exist (DebugView, or another vb6mcp server), this one doesn't start.
/// The writer (kernel32) opens the objects itself, including from the elevated IDE (higher integrity
/// may write to objects a medium process created).
/// </summary>
public sealed class DebugOutput : IDisposable
{
    private const int BufferSize = 4096;
    private const int Keep = 2000;

    private readonly object _lock = new();
    private readonly LinkedList<DebugMessage> _messages = new();
    private readonly Dictionary<int, string> _names = [];
    private IntPtr _bufferReady, _dataReady, _mapping, _view;
    private Thread? _thread;
    private volatile bool _stop;
    private long _nextId = 1;

    /// <summary>True once capturing.</summary>
    public bool Running => _thread != null;

    /// <summary>When capturing started.</summary>
    public DateTime Started { get; private set; }

    /// <summary>Starts capturing. Returns null on success, or why it couldn't (e.g. another monitor owns the buffer).</summary>
    public string? Start()
    {
        lock (_lock)
        {
            if (_thread != null) return null;
            _bufferReady = CreateEvent(IntPtr.Zero, false, false, "DBWIN_BUFFER_READY");
            bool exists = Marshal.GetLastWin32Error() == ERROR_ALREADY_EXISTS;
            if (_bufferReady == IntPtr.Zero) return "couldn't create DBWIN_BUFFER_READY (error " + Marshal.GetLastWin32Error() + ")";
            if (exists)
            {
                CloseHandle(_bufferReady); _bufferReady = IntPtr.Zero;
                return "another debug-output monitor owns the DBWIN buffer (DebugView, a debugger, or another vb6mcp session); close it and try again";
            }
            _dataReady = CreateEvent(IntPtr.Zero, false, false, "DBWIN_DATA_READY");
            _mapping = CreateFileMapping(new IntPtr(-1), IntPtr.Zero, PAGE_READWRITE, 0, BufferSize, "DBWIN_BUFFER");
            _view = _mapping == IntPtr.Zero ? IntPtr.Zero : MapViewOfFile(_mapping, FILE_MAP_READ, 0, 0, (UIntPtr)BufferSize);
            if (_dataReady == IntPtr.Zero || _view == IntPtr.Zero) { Release(); return "couldn't create the DBWIN buffer"; }
            Started = DateTime.Now;
            _thread = new Thread(Loop) { IsBackground = true, Name = "DBWIN monitor" };
            _thread.Start();
            return null;
        }
    }

    private void Loop()
    {
        var bytes = new byte[BufferSize - 4];
        while (!_stop)
        {
            SetEvent(_bufferReady);                              // "the buffer is free": a writer may fill it
            if (WaitForSingleObject(_dataReady, 500) != 0) continue;
            int pid = Marshal.ReadInt32(_view);
            Marshal.Copy(_view + 4, bytes, 0, bytes.Length);
            int len = Array.IndexOf(bytes, (byte)0);
            // OutputDebugStringA writes ANSI (the system code page; Windows-1252 here)
            var text = Ansi.GetString(bytes, 0, len < 0 ? bytes.Length : len).TrimEnd('\r', '\n');
            lock (_lock)
            {
                if (!_names.TryGetValue(pid, out var name))
                {
                    try { using var p = Process.GetProcessById(pid); name = p.ProcessName; } catch (ArgumentException) { name = "?"; }
                    _names[pid] = name;
                }
                _messages.AddLast(new DebugMessage(_nextId++, DateTime.Now, pid, name, text));
                while (_messages.Count > Keep) _messages.RemoveFirst();
            }
        }
    }

    /// <summary>Messages with an id greater than <paramref name="afterId"/>, oldest first, filtered.</summary>
    public IReadOnlyList<DebugMessage> Read(long afterId, Func<DebugMessage, bool>? filter = null, int max = 200)
    {
        lock (_lock)
            return _messages.Where(m => m.Id > afterId && (filter is null || filter(m))).TakeLast(max).ToList();
    }

    /// <summary>Id of the newest message (0 if none).</summary>
    public long LastId { get { lock (_lock) return _messages.Last?.Value.Id ?? 0; } }

    private static readonly Lazy<Encoding> s_ansi = new(() =>
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding((int)GetACP());
    });
    private static Encoding Ansi => s_ansi.Value;

    private void Release()
    {
        if (_view != IntPtr.Zero) UnmapViewOfFile(_view);
        foreach (var h in new[] { _mapping, _dataReady, _bufferReady }) if (h != IntPtr.Zero) CloseHandle(h);
        _view = _mapping = _dataReady = _bufferReady = IntPtr.Zero;
    }

    public void Dispose()
    {
        _stop = true;
        _thread?.Join(1000);
        _thread = null;
        lock (_lock) Release();
    }

    private const int ERROR_ALREADY_EXISTS = 183;
    private const uint PAGE_READWRITE = 0x04, FILE_MAP_READ = 0x0004;

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)] private static extern IntPtr CreateEvent(IntPtr attrs, bool manualReset, bool initial, string name);
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)] private static extern IntPtr CreateFileMapping(IntPtr file, IntPtr attrs, uint protect, uint sizeHigh, uint sizeLow, string name);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr MapViewOfFile(IntPtr mapping, uint access, uint offHigh, uint offLow, UIntPtr bytes);
    [DllImport("kernel32.dll")] private static extern bool UnmapViewOfFile(IntPtr view);
    [DllImport("kernel32.dll")] private static extern bool SetEvent(IntPtr h);
    [DllImport("kernel32.dll")] private static extern uint WaitForSingleObject(IntPtr h, uint ms);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll")] private static extern uint GetACP();
}
