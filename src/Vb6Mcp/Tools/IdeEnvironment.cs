using Vb6Mcp.Core.Bridge;
using Vb6Mcp.UI;

namespace Vb6Mcp.Tools;

/// <summary>Everything the tools need from the machine; swapped out in tests.</summary>
public interface IIdeEnvironment
{
    DiscoveryResult Scan();
    int DeleteStale(DiscoveryResult result);
    IdeWindowState ReadWindows(int pid);
    AddinClient Client(InstanceInfo instance);
    /// <summary>Window handle of the IDE's Immediate window, 0 if not found.</summary>
    long FindImmediate(int pid) => IdeWindows.FindImmediate(pid);
    /// <summary>The running watchdog (Phase 2b), or null. Null by default so tests never see a real one.</summary>
    InstanceInfo? FindWatchdog() => null;
    /// <summary>The recovery store the watchdog writes.</summary>
    Core.Recovery.RecoveryStore Recovery => new(Path.Combine(Path.GetTempPath(), "vb6mcp-no-recovery"));
}

/// <summary>The real environment: discovery folder, window reads, TCP to the add-in.</summary>
public sealed class IdeEnvironment(string? discoveryFolder = null, TimeSpan? timeout = null) : IIdeEnvironment
{
    public DiscoveryResult Scan() => Discovery.Scan(discoveryFolder);
    public int DeleteStale(DiscoveryResult result) => Discovery.DeleteStale(result);
    public IdeWindowState ReadWindows(int pid) => IdeWindows.Read(pid);
    public AddinClient Client(InstanceInfo instance) => new(instance, timeout);
    public InstanceInfo? FindWatchdog() => WatchdogDiscovery.Find();
    public Core.Recovery.RecoveryStore Recovery { get; } = new();
}
