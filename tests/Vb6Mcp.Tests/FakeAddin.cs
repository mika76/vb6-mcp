using System.Net;
using System.Net.Sockets;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;

namespace Vb6Mcp.Tests;

/// <summary>
/// A loopback server that speaks wire protocol v1 like the VB6 add-in, for unit tests.
/// The handler gets (verb, args) and returns (status, payload); returning null means "never reply".
/// </summary>
public sealed class FakeAddin : IDisposable
{
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly CancellationTokenSource _stop = new();

    public string Token { get; } = "0123456789abcdef0123456789abcdef";
    public int Port { get; }

    public FakeAddin(Func<string, string[], (string Status, string Payload)?> handler)
    {
        _listener.Start();
        Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
        _ = Task.Run(async () =>
        {
            while (!_stop.IsCancellationRequested)
            {
                TcpClient c;
                try { c = await _listener.AcceptTcpClientAsync(_stop.Token); } catch { return; }
                _ = Task.Run(() => Serve(c, handler));
            }
        });
    }

    public InstanceInfo Instance => new(Environment.ProcessId, Port, Token, 1, "0.0.0-test", "", "");

    private async Task Serve(TcpClient c, Func<string, string[], (string, string)?> handler)
    {
        using (c)
        {
            var s = c.GetStream();
            while (await Wire.ReadFrameAsync(s, _stop.Token) is string body)
            {
                var f = body.Split(Wire.US);
                string id = f.Length > 2 ? f[2] : "?";
                (string Status, string Payload)? reply = f.Length < 4 || f[0] != Wire.Magic ? ("ERR", $"400{Wire.US}bad request")
                    : f[1] != Token ? ("ERR", $"401{Wire.US}bad token")
                    : handler(f[3], f[4..]);
                if (reply is null) { await Task.Delay(Timeout.Infinite, _stop.Token); return; }
                await s.WriteAsync(Wire.EncodeFrame($"{Wire.Magic}{Wire.US}{id}{Wire.US}{reply.Value.Status}{Wire.US}{reply.Value.Payload}"));
            }
        }
    }

    public void Dispose() { _stop.Cancel(); _listener.Stop(); }
}
