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
}

/// <summary>The real environment: discovery folder, window reads, TCP to the add-in.</summary>
public sealed class IdeEnvironment(string? discoveryFolder = null, TimeSpan? timeout = null) : IIdeEnvironment
{
    public DiscoveryResult Scan() => Discovery.Scan(discoveryFolder);
    public int DeleteStale(DiscoveryResult result) => Discovery.DeleteStale(result);
    public IdeWindowState ReadWindows(int pid) => IdeWindows.Read(pid);
    public AddinClient Client(InstanceInfo instance) => new(instance, timeout);
}
