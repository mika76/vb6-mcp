using System.Text;

namespace Vb6Mcp.Core.Protocol;

/// <summary>
/// Wire protocol v1 between vb6mcp and the VB6 add-in. See docs/protocol.md.
/// Frame = "&lt;length&gt;\n&lt;body&gt;", body fields separated by US (0x1F), table records by RS (0x1E),
/// all text Windows-1252 (VB6 is ANSI).
/// </summary>
public static class Wire
{
    public const string Magic = "VB6MCP1";
    public const char US = '\x1F';
    public const char RS = '\x1E';
    public const int MaxBodyBytes = 16 * 1024 * 1024;

    /// <summary>Windows-1252, the encoding of everything on the wire.</summary>
    public static Encoding Ansi => Vb1252.Encoding;

    /// <summary>Encodes a body into a complete frame.</summary>
    public static byte[] EncodeFrame(string body)
    {
        var bytes = Ansi.GetBytes(body);
        if (bytes.Length > MaxBodyBytes) throw new ProtocolException($"frame body too large ({bytes.Length} bytes)");
        var head = Encoding.ASCII.GetBytes(bytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
        var frame = new byte[head.Length + bytes.Length];
        head.CopyTo(frame, 0);
        bytes.CopyTo(frame, head.Length);
        return frame;
    }

    /// <summary>Reads one frame body from the stream. Returns null on a clean end of stream before a frame starts.</summary>
    public static async Task<string?> ReadFrameAsync(Stream stream, CancellationToken ct)
    {
        // Length line: ASCII digits up to LF. 9 digits are enough for MaxBodyBytes.
        var digits = new StringBuilder();
        var one = new byte[1];
        while (true)
        {
            int n = await stream.ReadAsync(one, ct).ConfigureAwait(false);
            if (n == 0)
            {
                if (digits.Length == 0) return null;
                throw new ProtocolException("connection closed inside a frame header");
            }
            byte b = one[0];
            if (b == (byte)'\n') break;
            if (b < (byte)'0' || b > (byte)'9' || digits.Length >= 9)
                throw new ProtocolException($"invalid frame header byte 0x{b:X2}");
            digits.Append((char)b);
        }
        if (digits.Length == 0) throw new ProtocolException("empty frame length");
        int length = int.Parse(digits.ToString(), System.Globalization.CultureInfo.InvariantCulture);
        if (length > MaxBodyBytes) throw new ProtocolException($"frame body too large ({length} bytes)");

        var body = new byte[length];
        int read = 0;
        while (read < length)
        {
            int n = await stream.ReadAsync(body.AsMemory(read, length - read), ct).ConfigureAwait(false);
            if (n == 0) throw new ProtocolException($"connection closed after {read} of {length} body bytes");
            read += n;
        }
        return Ansi.GetString(body);
    }

    /// <summary>Builds a request body: Magic US token US id US verb [US arg]...</summary>
    public static string BuildRequest(string token, string id, string verb, params string[] args)
    {
        var parts = new List<string>(4 + args.Length) { Magic, token, id, verb };
        parts.AddRange(args);
        foreach (var p in parts)
            if (p.Contains(US) || p.Contains(RS))
                throw new ProtocolException("request fields must not contain US (0x1F) or RS (0x1E)");
        return string.Join(US, parts);
    }

    /// <summary>Parses a response body. The payload is everything after the third US.</summary>
    public static Response ParseResponse(string body)
    {
        var parts = body.Split(US, 4);
        if (parts.Length < 3 || parts[0] != Magic) throw new ProtocolException("not a VB6MCP1 response");
        var status = parts[2] switch
        {
            "OK" => ResponseStatus.Ok,
            "ERR" => ResponseStatus.Error,
            "QUEUED" => ResponseStatus.Queued,
            _ => throw new ProtocolException($"unknown response status '{parts[2]}'"),
        };
        return new Response(parts[1], status, parts.Length == 4 ? parts[3] : "");
    }

    /// <summary>Splits a table payload into records (RS) and fields (US). An empty payload has no records.</summary>
    public static IReadOnlyList<string[]> Records(string payload)
        => payload.Length == 0 ? [] : payload.Split(RS).Select(r => r.Split(US)).ToArray();

    /// <summary>Splits a payload into exactly <paramref name="count"/> fields; the last one keeps any further US.</summary>
    public static string[] Fields(string payload, int count)
    {
        var f = payload.Split(US, count);
        if (f.Length != count) throw new ProtocolException($"expected {count} fields, got {f.Length}");
        return f;
    }
}

public enum ResponseStatus { Ok, Error, Queued }

/// <summary>A parsed response. For errors, <see cref="ErrorCode"/>/<see cref="ErrorMessage"/> decode the payload.</summary>
public sealed record Response(string Id, ResponseStatus Status, string Payload)
{
    public int ErrorCode => Status == ResponseStatus.Error && int.TryParse(Payload.Split(Wire.US, 2)[0], out var c) ? c : 0;
    public string ErrorMessage => Status == ResponseStatus.Error ? (Payload.Split(Wire.US, 2) is [_, var m] ? m : Payload) : "";
}

/// <summary>Malformed frames, bodies or responses.</summary>
public sealed class ProtocolException(string message) : Exception(message);
