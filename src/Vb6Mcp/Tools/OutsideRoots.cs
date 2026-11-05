using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>
/// The refusal for a path outside the workspace roots, shared by every tool that checks them. It names the supported way
/// to work there (widen the roots) because the bare refusal left an agent with no way in, and it copied new VB6 files in
/// through the shell, past the hook (notes from real use, 2026-11-05).
/// </summary>
internal static class OutsideRoots
{
    public static ToolException Error(string path, IReadOnlyList<string> roots) =>
        new($"{path} is outside the workspace roots ({(roots.Count == 0 ? "none are configured" : string.Join("; ", roots))}). " +
            "The roots are the folders vb6-mcp may read, write, build and register in. To work in another folder, add it to them: " +
            (Environment.GetEnvironmentVariable("VB6MCP_ROOTS") is { Length: > 0 }
                ? "scripts\\install-server.ps1 -Roots with the old and new folders (development install), then server_restart. "
                : "run the package's install.ps1 -Roots with the old and new folders (one UAC prompt), then restart the session. ") +
            "Ask the user first; don't copy VB6 files in or out with the shell to get around this (the hook can't check those writes)");
}
