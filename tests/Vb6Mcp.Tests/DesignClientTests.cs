using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Core.Protocol;

namespace Vb6Mcp.Tests;

public static class DesignClientTests
{
    const char US = Wire.US, RS = Wire.RS;

    /// <summary>DESIGN is queued, RESULT is polled until the deferred operation ran (hard rule 14).</summary>
    [Test]
    public static void DesignQueuesThenPollsTheResult()
    {
        int polls = 0;
        string? queuedId = null;
        using var fake = new FakeAddin((verb, args) => verb switch
        {
            "DESIGN" => (queuedId = args[0]) is not null && args[1] == "CONTROLS" ? ("QUEUED", args[0]) : ("ERR", $"400{US}bad"),
            "RESULT" when args[0] == queuedId => ++polls < 3 ? ("OK", "pending") : ("OK", $"ok{US}1{US}Command1{US}-1{US}VB.CommandButton"),
            _ => ("ERR", $"400{US}unexpected {verb}"),
        });
        var (dirty, payload) = new AddinClient(fake.Instance).DesignAsync("CONTROLS", ["Project1", "Form1"]).GetAwaiter().GetResult();
        Assert.True(dirty, "dirty flag");
        Assert.Equal($"Command1{US}-1{US}VB.CommandButton", payload);
        Assert.Equal(3, polls);
    }

    [Test]
    public static void FailedDesignOperationThrowsWithTheVbError()
    {
        using var fake = new FakeAddin((verb, args) => verb == "DESIGN" ? ("QUEUED", args[0]) : ("OK", $"failed{US}380: Invalid property value"));
        try
        {
            new AddinClient(fake.Instance).DesignAsync("SET", ["P", "F", "Text1", "Left", "x"]).GetAwaiter().GetResult();
            throw new InvalidOperationException("should throw");
        }
        catch (AddinErrorException ex) { Assert.Equal("380: Invalid property value", ex.AddinMessage); }
    }

    [Test]
    public static void OldAddinGetsAClearMessage()
    {
        using var fake = new FakeAddin((verb, _) => ("ERR", $"400{US}unknown verb {verb}"));
        try
        {
            new AddinClient(fake.Instance).DesignAsync("CONTROLS", ["P", "F"]).GetAwaiter().GetResult();
            throw new InvalidOperationException("should throw");
        }
        catch (AddinErrorException ex) { Assert.True(ex.AddinMessage.Contains("rebuild"), ex.AddinMessage); }
    }

    [Test]
    public static void ReferencesParse()
    {
        using var fake = new FakeAddin((verb, _) => verb == "REFS"
            ? ("OK", $"VBA{US}{{000204EF-0000-0000-C000-000000000046}}{US}6{US}0{US}C:\\x\\VBE6.DLL{US}1{US}0{US}0{US}Visual Basic For Applications{RS}" +
                     $"Gone{US}{US}{US}{US}{US}0{US}1{US}{US}")
            : ("ERR", $"400{US}x"));
        var refs = new AddinClient(fake.Instance).RefsAsync("P").GetAwaiter().GetResult();
        Assert.Equal(2, refs.Count);
        Assert.True(refs[0].BuiltIn && !refs[0].IsBroken && refs[0].Major == 6, "built-in VBA");
        Assert.True(refs[1].IsBroken && refs[1].Name == "Gone", "broken reference with empty fields");
    }

    [Test]
    public static void RegistryVersionKeysAreHex()
    {
        Assert.True(TypeLibRegistry.TryVersion("b.5", out int ma, out int mi) && ma == 11 && mi == 5, "b.5");
        Assert.True(!TypeLibRegistry.TryVersion("FLAGS", out _, out _), "not a version");
        // stdole is registered on every Windows: the registry reader finds it with its file
        var stdole = TypeLibRegistry.All().Where(t => t.Description.Contains("OLE Automation", StringComparison.OrdinalIgnoreCase) && t.Major == 2).ToList();
        Assert.True(stdole.Any(t => t.FileExists), "OLE Automation 2.x registered with its file");
    }

    /// <summary>The typelib describer on stdole2.tlb: a known enum and a coclass, rendered in VB syntax.</summary>
    [Test]
    public static void DescribesATypeLibInVbSyntax()
    {
        var tlb = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.SystemX86), "stdole2.tlb");
        var d = TypeLibDescriber.Describe(tlb);
        Assert.Equal("stdole", d.Info.Name);
        var font = d.Types.FirstOrDefault(t => t.Name == "StdFont");
        Assert.True(font is not null && font.Kind == "coclass" && font.Implements.Any(i => i.Contains("[default]")), "StdFont coclass");
        var text = TypeLibDescriber.Render(d, "Font");
        Assert.True(text.Contains("Class StdFont") && text.Contains("Property Get Bold() As Boolean"), text[..Math.Min(400, text.Length)]);
    }
}
