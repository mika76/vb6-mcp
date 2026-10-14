using Vb6Mcp.Core;
using Vb6Mcp.Core.Bridge;
using Vb6Mcp.Core.Protocol;
using Vb6Mcp.UI;
using Vb6Mcp.Watchdog;

namespace Vb6Mcp.Tests;

public static class SyncTests
{
    const char US = Wire.US;

    /// <summary>A fake IDE with one project and one module file in a temp folder.</summary>
    sealed class FakeIde : IDisposable
    {
        public readonly string Dir = Path.Combine(Path.GetTempPath(), $"vb6mcp-sync-{Guid.NewGuid():N}");
        public string Bas => Path.Combine(Dir, "modA.bas");
        public string Vbp => Path.Combine(Dir, "P.vbp");
        public bool Dirty;
        public List<string> OwnWrites = [];
        public List<string> Reloads = [];
        public FakeAddin Addin;

        public FakeIde()
        {
            Directory.CreateDirectory(Dir);
            File.WriteAllText(Vbp, "Type=Exe\r\nModule=modA; modA.bas\r\n");
            File.WriteAllText(Bas, "Attribute VB_Name = \"modA\"\r\nOption Explicit\r\n");
            Addin = new FakeAddin((verb, a) => verb switch
            {
                "WRITES" => ("OK", $"{OwnWrites.Count}{US}{string.Join(Wire.RS, OwnWrites.Skip(int.Parse(a[0])).Select(p => p + US + "0"))}"),
                "PROJECTS" => ("OK", $"1{US}P{US}0{US}{Vbp}{US}0"),
                "COMPONENTS" => ("OK", $"modA{US}1{US}{(Dirty ? 1 : 0)}{US}1{US}{Bas}"),
                "RELOAD" => Reload(a),
                _ => ("ERR", $"400{US}unknown {verb}"),
            });
        }

        (string, string) Reload(string[] a) { Reloads.Add($"{a[1]} {a[2]}"); return ("QUEUED", "reload"); }

        /// <summary>Changes the module on disk with a clearly newer time stamp.</summary>
        public void Touch(string text)
        {
            File.WriteAllText(Bas, text);
            File.SetLastWriteTimeUtc(Bas, File.GetLastWriteTimeUtc(Bas).AddSeconds(5 + Reloads.Count + OwnWrites.Count));
        }

        public void Dispose() { Addin.Dispose(); try { Directory.Delete(Dir, true); } catch (IOException) { } }
    }

    [Test]
    public static async Task OutsideChangesAreReloadedConflictsKept()
    {
        using var ide = new FakeIde();
        var mode = IdeMode.Design;
        var state = Path.Combine(ide.Dir, "state.json");
        var poller = new SyncPoller(() => new DiscoveryResult([ide.Addin.Instance], [], []), i => new AddinClient(i, TimeSpan.FromSeconds(2)), _ => mode, _ => { }, state);
        async Task Poll() => await poller.PollOnceAsync(default);

        await Poll();                                             // first sight: nothing
        ide.Touch("Attribute VB_Name = \"modA\"\r\nOption Explicit\r\n' changed by git\r\n");
        await Poll();                                             // seen once: wait for it to settle
        Assert.Equal(0, ide.Reloads.Count, "not before the change is stable");
        await Poll();
        Assert.SequenceEqual(new[] { "modA clean" }, ide.Reloads, "clean component reloaded");

        ide.OwnWrites.Add(ide.Bas);                               // the IDE saves the file itself
        ide.Touch("Attribute VB_Name = \"modA\"\r\nOption Explicit\r\n' saved in the IDE\r\n");
        await Poll(); await Poll();
        Assert.Equal(1, ide.Reloads.Count, "own save: no reload");

        ide.Dirty = true;                                         // unsaved IDE changes + outside change
        ide.Touch("' changed elsewhere\r\n");
        await Poll(); await Poll();
        Assert.Equal(1, ide.Reloads.Count, "dirty component never reloaded");
        var s = SyncState.Read(state);
        Assert.True(s.Conflicts.Count == 1 && s.Conflicts[0].Component == "modA", "conflict recorded");

        ide.Dirty = false;                                        // the user resolved it (e.g. sync_take_disk)
        await Poll();
        Assert.Equal(0, SyncState.Read(state).Conflicts.Count, "conflict resolved when no longer dirty");

        mode = IdeMode.Run;                                       // change during a run waits for design mode
        ide.Touch("' changed during a run\r\n");
        await Poll(); await Poll();
        Assert.Equal(1, ide.Reloads.Count, "no reload while running");
        Assert.True(SyncState.Read(state).Events.Any(e => e.Kind == "pending"), "pending event");
        mode = IdeMode.Design;
        await Poll();
        Assert.Equal(2, ide.Reloads.Count, "reloaded back in design mode");
        var kinds = SyncState.Read(state).Events.Select(e => e.Kind).ToList();
        Assert.SequenceEqual(new[] { "reloaded", "conflict", "resolved", "pending", "reloaded" }, kinds);
    }
}
