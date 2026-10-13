using System.Net;
using System.Net.Sockets;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;

namespace Vb6Mcp.Tests;

public static class AddinClientTests
{
    const char US = Wire.US, RS = Wire.RS;

    [Test]
    public static async Task TypedVerbsParsePayloads()
    {
        using var fake = new FakeAddin((verb, args) => verb switch
        {
            "PING" => ("OK", $"4242{US}0.1.0{US}1"),
            "PROJECTS" => ("OK", $"1{US}Project1{US}0{US}D:\\p\\Project1.vbp{US}1"),
            "COMPONENTS" => ("OK", $"Form1{US}5{US}1{US}5{US}D:\\p\\Form1.frm{RS}Module1{US}1{US}0{US}12{US}"),
            "READ" => ("OK", $"5{US}1{US}Option Explicit\r\n\r\nPrivate Sub Form_Load()"),
            "MEMBERS" => ("OK", $"Form_Load{US}1{US}1{US}0{US}3{US}0{US}2{US}3{US}3{RS}m_x{US}3{US}1{US}0{US}1{US}{US}{US}{US}"),
            "PROC" => ("OK", $"2{US}3{US}3{US}Private Sub Form_Load()\r\nEnd Sub"),
            _ => ("ERR", $"400{US}unknown verb {verb}"),
        });
        var client = new AddinClient(fake.Instance);

        var ping = await client.PingAsync();
        Assert.Equal(4242, ping.Pid); Assert.Equal("0.1.0", ping.AddinVersion); Assert.Equal(1, ping.Protocol);

        var prj = Assert1(await client.ProjectsAsync());
        Assert.Equal("Project1", prj.Name); Assert.True(prj.IsDirty);

        var cmps = await client.ComponentsAsync("Project1");
        Assert.Equal(2, cmps.Count); Assert.Equal("", cmps[1].FileName); Assert.Equal(12, cmps[1].CountOfLines);

        var code = await client.ReadAsync("Project1", "Form1");
        Assert.Equal(5, code.CountOfLines); Assert.True(code.IsDirty); Assert.True(code.Text.StartsWith("Option Explicit\r\n"));

        var mem = await client.MembersAsync("Project1", "Form1");
        Assert.Equal(0, mem[0].ProcKind); Assert.Equal(3, mem[0].ProcBody);
        Assert.Equal<int?>(null, mem[1].ProcKind, "variable has no proc kind");

        var proc = await client.ProcAsync("Project1", "Form1", "Form_Load", 0);
        Assert.Equal(3, proc.ProcCount); Assert.True(proc.Text.EndsWith("End Sub"));
    }

    [Test]
    public static async Task ErrorsBecomeAddinErrorException()
    {
        using var fake = new FakeAddin((_, _) => ("ERR", $"404{US}component not found: Project1.Form9"));
        var ex = await ThrowsAsync<AddinErrorException>(() => new AddinClient(fake.Instance).ReadAsync("Project1", "Form9"));
        Assert.Equal(404, ex.Code);
        Assert.Equal("component not found: Project1.Form9", ex.AddinMessage);
    }

    [Test]
    public static async Task WrongTokenIsRejected()
    {
        using var fake = new FakeAddin((_, _) => ("OK", ""));
        var bad = fake.Instance with { Token = "ffffffffffffffffffffffffffffffff" };
        var ex = await ThrowsAsync<AddinErrorException>(() => new AddinClient(bad).PingAsync());
        Assert.Equal(401, ex.Code);
    }

    [Test]
    public static async Task NoReplyBecomesBusyNotAHang()
    {
        using var fake = new FakeAddin((_, _) => null);   // like an IDE stuck in a busy loop
        var sw = System.Diagnostics.Stopwatch.StartNew();
        await ThrowsAsync<AddinBusyException>(() => new AddinClient(fake.Instance, TimeSpan.FromMilliseconds(300)).PingAsync());
        Assert.True(sw.ElapsedMilliseconds < 3000, $"took {sw.ElapsedMilliseconds} ms");
    }

    [Test]
    public static async Task ClosedPortBecomesUnreachable()
    {
        var l = new TcpListener(IPAddress.Loopback, 0); l.Start();
        int port = ((IPEndPoint)l.LocalEndpoint).Port; l.Stop();
        var info = new InstanceInfo(1, port, "t", 1, "", "", "");
        await ThrowsAsync<AddinUnreachableException>(() => new AddinClient(info).PingAsync());
    }

    static T Assert1<T>(IReadOnlyList<T> list) { Assert.Equal(1, list.Count, "count"); return list[0]; }

    static async Task<TEx> ThrowsAsync<TEx>(Func<Task> f) where TEx : Exception
    {
        try { await f(); }
        catch (TEx ex) { return ex; }
        catch (Exception ex) { throw new AssertFailed($"expected {typeof(TEx).Name}, got {ex.GetType().Name}: {ex.Message}"); }
        throw new AssertFailed($"expected {typeof(TEx).Name}, nothing thrown");
    }
}
