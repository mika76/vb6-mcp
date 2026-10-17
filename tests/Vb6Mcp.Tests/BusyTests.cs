using System.Net;
using System.Net.Sockets;
using Vb6Mcp.Core.Bridge;

namespace Vb6Mcp.Tests;

/// <summary>A refused connection while the IDE runs an operation we queued is "busy", not "unreachable" (Phase 7 B4).</summary>
public static class BusyTests
{
    [Test]
    public static async Task ARefusedConnectionDuringAQueuedMakeIsBusy()
    {
        // a port nobody listens on: take a free one and release it
        var probe = new TcpListener(IPAddress.Loopback, 0); probe.Start();
        int port = ((IPEndPoint)probe.LocalEndpoint).Port; probe.Stop();
        var inst = new InstanceInfo(Environment.ProcessId, port, "token", 1, "0.0.0-test", "", "");   // this process: alive
        var client = new AddinClient(inst, TimeSpan.FromSeconds(10));   // a refused loopback connect takes ~2 s on Windows

        try { await client.PingAsync(); throw new InvalidOperationException("should fail"); }
        catch (AddinUnreachableException) { }   // nothing queued: unreachable

        client.MarkBusy("compiling (File > Make)");
        try { await client.PingAsync(); throw new InvalidOperationException("should fail"); }
        catch (AddinBusyException ex) { Assert.True(ex.Message.Contains("busy compiling"), ex.Message); }
        finally { client.ClearBusy(); }   // the mark is per pid, and the fake add-ins of other tests use this process id
    }
}
