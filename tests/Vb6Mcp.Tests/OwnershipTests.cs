using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.Tools;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tests;

/// <summary>IDE ownership between sessions (Phase 7 A9): another session's IDE refuses write tools, reads still work.</summary>
public static class OwnershipTests
{
    const char US = Wire.US;
    const string Module = "Option Explicit\r\n";

    sealed class Env(FakeAddin addin) : IIdeEnvironment
    {
        public DiscoveryResult Scan() => new([addin.Instance], [], []);
        public int DeleteStale(DiscoveryResult result) => 0;
        public IdeWindowState ReadWindows(int pid) => new(pid, true, "Basic - Microsoft Visual Basic [design]", IdeMode.Design, true, false, []);
        public AddinClient Client(InstanceInfo instance) => new(instance, TimeSpan.FromSeconds(2));
    }

    static async Task<(bool IsError, string Text)> Call(IdeTools tools, string name, JsonObject args)
    {
        var r = (await tools.All().First(t => t.Name == name).InvokeAsync(args, CancellationToken.None)).ToJson();
        return ((bool)r["isError"]!, (string)r["content"]![0]!["text"]!);
    }

    [Test]
    public static async Task AnotherSessionsIdeRefusesWritesButNotReads()
    {
        var owners = Directory.CreateTempSubdirectory("vb6mcp-owners-").FullName;
        try
        {
            var sent = new List<string>();
            using var fake = new FakeAddin((verb, a) =>
            {
                sent.Add(verb);
                return verb switch
                {
                    "PROJECTS" => ("OK", $"1{US}Basic{US}0{US}D:/b/Basic.vbp{US}0"),
                    "READ" => ("OK", $"1{US}0{US}{Module}"),
                    "REPLACE" => ("OK", $"2{US}1"),
                    _ => ("ERR", $"400{US}unexpected {verb}"),
                };
            });
            // the IDE belongs to another live session (any pid other than this test's)
            IdeOwnership.Claim(fake.Instance.Pid, new IdeOwner(424242, 1, "other-client 1.0", DateTime.UtcNow), owners);
            var tools = new IdeTools(new Env(fake)) { OwnersFolder = owners, OwnerAlive = _ => true };

            var write = await Call(tools, "code_insert", new JsonObject { ["component"] = "modA", ["line"] = 1, ["text"] = "' x", ["expected_hash"] = VbCode.Hash(Module) });
            Assert.True(write.IsError && write.Text.Contains("belongs to another session") && write.Text.Contains("other-client"), write.Text);
            Assert.True(!sent.Contains("REPLACE"), "nothing was changed");

            var read = await Call(tools, "code_read", new JsonObject { ["component"] = "modA" });
            Assert.True(!read.IsError, read.Text);

            // once the owner has gone, the IDE is shared again
            var after = new IdeTools(new Env(fake)) { OwnersFolder = owners, OwnerAlive = _ => false };
            var ok = await Call(after, "code_insert", new JsonObject { ["component"] = "modA", ["line"] = 1, ["text"] = "' x", ["expected_hash"] = VbCode.Hash(Module) });
            Assert.True(!ok.IsError, ok.Text);
        }
        finally { Directory.Delete(owners, true); }
    }

    [Test]
    public static async Task ClaimAndReleaseOwnIde()
    {
        var owners = Directory.CreateTempSubdirectory("vb6mcp-owners-").FullName;
        try
        {
            using var fake = new FakeAddin((verb, a) => ("ERR", $"400{US}unexpected {verb}"));
            var tools = new IdeTools(new Env(fake)) { OwnersFolder = owners };
            var claimed = await Call(tools, "ide_claim", new JsonObject { ["pid"] = fake.Instance.Pid });
            Assert.True(!claimed.IsError && claimed.Text.Contains("claimed"), claimed.Text);
            Assert.True(IdeOwnership.Read(fake.Instance.Pid, owners) is { } o && IdeOwnership.IsMe(o), "owned by this process");
            var released = await Call(tools, "ide_release", new JsonObject { ["pid"] = fake.Instance.Pid });
            Assert.True(!released.IsError && IdeOwnership.Read(fake.Instance.Pid, owners) is null, released.Text);
        }
        finally { Directory.Delete(owners, true); }
    }
}
