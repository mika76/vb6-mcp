# vb6-mcp

An MCP server that lets AI agents (Claude Code) work **inside the Visual Basic 6 IDE** the way a developer does: read and edit live (unsaved) code, run and debug, drive the running app's UI, take screenshots, and build.

- **Design and decisions:** [PLAN.md](PLAN.md)
- **Rules for agents working on this repo:** [AGENTS.md](AGENTS.md)
- **Architecture:** [docs/architecture.md](docs/architecture.md)
- **Wire protocol (server ↔ add-in ↔ watchdog):** [docs/protocol.md](docs/protocol.md)
- **MCP tool reference:** [docs/tools.md](docs/tools.md)
- **Experiments that back the design:** [docs/spikes/](docs/spikes/)

> Status: **Phase 1 in progress** (MCP server + production add-in). The spike code under `spikes/` is throwaway and is never referenced by `src/`.

## Requirements

- Windows 10/11, **Visual Basic 6.0 (SP6)**, normally run **as Administrator**
- .NET 10 SDK
- Claude Code (desktop app or CLI)

## Repository layout

| Path | What |
|------|------|
| `src/Vb6Mcp.Core` | Wire protocol, discovery, add-in client, VB6 file formats |
| `src/Vb6Mcp.UI` | Window enumeration, IDE title/mode, dialogs (later: screenshots) |
| `src/Vb6Mcp` | `vb6mcp.exe`: the MCP server (stdio) |
| `addin/VB6McpBridge` | The VB6 add-in that runs inside the IDE |
| `tests/Vb6Mcp.Tests` | Tests and our own minimal test runner |
| `tests/fixtures` | Small VB6 projects used by tests |
| `spikes/` | Phase 0 experiments (throwaway) |
| `docs/` | Architecture, protocol, tools, spike write-ups |

## Build and test

```powershell
dotnet build vb6-mcp.slnx
dotnet run --project tests/Vb6Mcp.Tests                # unit tests (no VB6 needed)
dotnet run --project tests/Vb6Mcp.Tests -- --cat IDE   # needs a running IDE with the add-in
dotnet run --project tests/Vb6Mcp.Tests -- --all       # everything
```

The test runner's exit code is the number of failed tests.

## Setup

### 1. Build and register the VB6 add-in (elevated)

From an **Administrator** PowerShell (building an ActiveX DLL registers it machine-wide, and starting VB6 from a non-elevated shell triggers a UAC prompt):

```powershell
.\scripts\build-addin.ps1
```

- Unload the add-in in VB6 first (Add-Ins → Add-In Manager), or close VB6: a loaded DLL can't be overwritten.
- The first build ever creates the binary-compatibility reference `addin\compat\VB6McpBridge.dll` (committed), so the add-in's ProgID/CLSIDs stay stable.
- The script then registers the add-in with VB6 (`scripts\register-addin.ps1`, HKCU). It must run after the build: a VB6 that starts while the Addins key names an unregistered ProgID deletes the key.
- It also publishes the **watchdog** (`vb6mcp.exe`) into `addin\VB6McpBridge\build\watchdog`. The add-in starts it when it loads, and it inherits elevation with no UAC prompt. It handles a stuck IDE (Ctrl+Break, dialogs, kill, screenshots) and exits 2 minutes after the last VB6 closes. A running watchdog is stopped first, because it locks its exe. Its log is `%LOCALAPPDATA%\vb6mcp\watchdog.log`.

**Restart VB6** (as Administrator): VB6 reads its add-in list only at startup. You should see a **"VB6 MCP" toolbar** showing `VB6 MCP <version> · built <time> · :<port> · <n> req`.

### 2. Install the MCP server for Claude Code

```powershell
.\scripts\install-server.ps1 -Roots 'D:\src\VB6Apps','D:\src\Shared'
```

`-Roots` sets the **workspace roots** (`VB6MCP_ROOTS`) for the disk tools (`ws_projects`, `ws_file_projects`, `file_read`): the folders that contain your `.vbp` files and any shared source folders. Paths outside them are refused.

Publishes `vb6mcp.exe` to a new folder `build\server\<timestamp>\` (a running Claude session locks the installed exe, so each install gets its own folder; old unused ones are removed) and registers it at **local scope** for this folder (`claude mcp add --scope local vb6 …`). Restart the Claude session to load it; check with `claude mcp get vb6` (run from this folder). The server's log goes to `build\vb6mcp.log`.

### 3. Use it

Open a project in VB6 (or let the agent call `project_open`), then ask Claude to use the `vb6` tools: `ide_list`, `ide_status`, `project_components`, `code_read`, `symbol_outline`, `proc_get`, `project_open`. See [docs/tools.md](docs/tools.md).

## Elevated or not?

Whether VB6 runs as administrator is your machine's configuration (a "Run as administrator" shortcut or compatibility setting); a default install runs it non-elevated (spike 16). vb6-mcp works either way (`ide_list` shows `elevated` per IDE):

| | VB6 elevated | VB6 not elevated |
|---|---|---|
| Read/edit live code, open/save projects, dialogs | ✓ | ✓ |
| Build ActiveX DLL/OCX (registers the component) | ✓ | ✗ registration needs admin rights |
| Input into the IDE / running app (`app_click`, `app_set_text`) | through the add-in only | through the add-in (Claude's computer use also works) |
| Screenshots (`app_screenshot`) | captured by the add-in | captured by the server |
| Starting VB6 from Claude (`ide_launch`) | UAC prompt | no prompt |

`ide_launch` takes `elevation: auto | elevated | normal`. Machine-specific values come from configuration: `VB6MCP_VB6_EXE` (VB6.EXE; otherwise found via the registry) and `VB6MCP_ROOTS` (workspace roots).

## Troubleshooting

| Symptom | Check |
|---------|-------|
| No "VB6 MCP" toolbar | Add-In Manager: is "VB6 MCP Bridge" loaded? `LoadBehavior` gets rewritten by Add-In Manager. Log: `%LOCALAPPDATA%\vb6mcp\addin.log` |
| "VB6 MCP Bridge" missing from Add-In Manager | Restart VB6 (list read at startup); re-run `scripts\register-addin.ps1` |
| Tools say "no VB6 IDE with the vb6-mcp add-in is running" | Discovery files in `%LOCALAPPDATA%\vb6mcp\instances\`; is the add-in loaded? |
| Tools say the IDE is busy | The debugged program is running a loop, or a dialog is open: `ide_status` |
| Build fails with "Permission denied" | The add-in DLL is loaded in VB6: unload it first |
| `vb6` isn't in the app's Connectors list | Expected: Connectors shows claude.ai account connectors and desktop extensions. Claude Code's own MCP servers (like `vb6`) are listed under `/mcp` in the session |
| New tools don't show up after `install-server.ps1` | The session keeps running the previously started server, even across a session restart. Quit and reopen the Claude app (the desktop app's `/mcp` list didn't offer a Reconnect button for `vb6`) |
| `ide_launch` with `normal` says elevation is required | VB6.EXE has a "Run as administrator"/compatibility setting on that path (spike 16); use `auto`/`elevated` or change the setting |
| Opening a project shows "Add this project to SourceSafe?" | Your source-control add-in; `project_open` reports it and leaves the answer to you |
