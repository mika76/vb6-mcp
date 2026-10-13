# Architecture

This describes **what exists in `src/` and `addin/` now**, and why. The full design, including parts not built yet, is in [PLAN.md](../PLAN.md). Each "why" links the spike that proved it.

## Processes

```
 Claude Code (medium integrity)
     │ stdio, MCP (JSON-RPC 2.0, newline-delimited)
     ▼
 vb6mcp.exe  (src/Vb6Mcp, medium integrity, x86)
     │  ├─ discovery: %LOCALAPPDATA%\vb6mcp\instances\<pid>.json
     │  ├─ window reads of the IDE (title/mode, responsiveness, dialogs)   ── src/Vb6Mcp.UI
     │  └─ TCP 127.0.0.1:<port>, wire protocol v1 (docs/protocol.md)      ── src/Vb6Mcp.Core
     ▼
 VB6.EXE (elevated) + VB6McpBridge add-in (addin/VB6McpBridge)
        VBIDE object model, CommandBars, in-process input
```

| Process | Integrity | Role |
|---------|-----------|------|
| Claude Code | medium | MCP client; starts `vb6mcp.exe` |
| `vb6mcp.exe` | medium | All logic. Talks MCP to Claude, TCP to add-ins, reads IDE windows |
| `VB6.EXE` + add-in | **high** (VB6 runs as Administrator) | Thin endpoint for VBIDE access |
| watchdog (Phase 2b) | high | Hangs, crashes, screenshots, restart (spikes 12, 12b, 15) |

### Why this split

- **Medium → high communication:** window messages and synthetic input are blocked (silently dropped, spike [10](spikes/10-computer-use-elevation.md)); **loopback TCP works** (spike [11](spikes/11-tcp-elevation.md)). So the add-in, already elevated inside VB6, listens on TCP.
- **Reads cross the boundary:** window titles, dialog text and window enumeration work from medium (spike 10). So IDE state (mode, dialogs, responsiveness) is read directly by `vb6mcp.exe`, even while the IDE thread is busy.
- **VBIDE is only reachable in-process:** VB6 doesn't register its `VBE` object in the ROT.

## Components in `src/`

### `Vb6Mcp.Core`

| Type | What | Notes |
|------|------|-------|
| `Vb1252` | Windows-1252 encoding | VB6 is ANSI (hard rule 1) |
| `VbCode` | module hash, validation of text to insert (CR LF, 1023 chars/line, 24 continuations, strict 1252), slicing, procedure-range trimming | hard rules 1, 10 |
| `Protocol.Wire` | frames, request building, response/record parsing | [protocol.md](protocol.md) |
| `Bridge.Discovery` | scans instance files; live vs stale (crashed/killed) | stale = process gone or PID reused (spike 12) |
| `Bridge.AddinClient` | one request per connection, timeout | timeout → `AddinBusyException` (hard rule 6) |

### `Vb6Mcp.UI`

| Type | What | Notes |
|------|------|-------|
| `IdeWindows` | IDE main window title → mode, `WM_NULL` probe, `IsHungAppWindow`, `#32770` dialogs | title tag not always last (spike 12); cached caption readable while busy |

### `Vb6Mcp` (`vb6mcp.exe`)

| Type | What | Notes |
|------|------|-------|
| `Mcp.McpServer` | JSON-RPC loop, initialize/ping/tools | hand-written, no SDK (spike [1](spikes/01-mcp-handshake.md)); unknown methods → -32601 (the desktop client probes `server/discover`); requests handled concurrently |
| `Mcp.Tool`, `ToolResult`, `Schema` | tool definitions and results | expected failures → `isError` results |
| `Tools.IdeTools` | read tools (`IdeTools.cs`) and editing tools (`IdeTools.Edit.cs`) | [tools.md](tools.md); edits: module hash + old-text guard, design mode only |
| `Tools.IIdeEnvironment` | discovery + window reads + clients | swapped for fakes in tests |

## The add-in (`addin/VB6McpBridge`)

VB6 ActiveX DLL, ProgID `VB6McpBridge.Connect`, version 0.1.0, binary-compatible with `addin/compat/VB6McpBridge.dll`. Built and registered by `scripts/build-addin.ps1` (elevated); registered with VB6 by `scripts/register-addin.ps1` (HKCU, called by the build after the DLL's ProgID exists: a VB6 that starts while the Addins key names an unregistered ProgID deletes the key).

| Module | What |
|--------|------|
| `Connect.cls` | `IDTExtensibility`: starts/stops listener and status toolbar |
| `Listener.bas` | Winsock listener, protocol v1 framing, buffered output (`FD_WRITE`), discovery file (temp + rename; deleted on clean unload), token check |
| `Verbs.bas` | `PING`, `PROJECTS`, `COMPONENTS`, `READ`, `MEMBERS`, `PROC`, `OPEN`, `REPLACE`, `SAVE`, `CLICK`; errors → 400/404/409/500 |
| `Defer.bas` | queue + `WM_DEFER` message: runs deferred actions (`OPEN`, `SAVE`) from a fresh dispatch |
| `Input.bas` | input into the IDE's own windows: `BM_CLICK` to buttons of this process only |
| `Status.bas` | "VB6 MCP" toolbar: version, DLL build time, port, request count |
| `Log.bas` | shared state, `%LOCALAPPDATA%\vb6mcp\addin.log` |

Rules it follows, all from spikes:
- Raw Winsock with `WSAAsyncSelect` on a hidden message-only window: socket events arrive on the IDE thread; no polling, no blocking calls (spike 2).
- **Indexed loops only** over VBIDE collections; `For Each` fails with error 70 in run/break mode (spike 3).
- **No VBIDE access in build-event handlers:** it ended the debugged program (spike 3).
- **Deferred execution** for anything that can run user code (IDE commands): queue via `PostMessage` to its own window, run from a fresh dispatch; otherwise a break under the handler deadlocks the IDE (spikes 3, 6/8; hard rule 14).
- Starts the watchdog with `CreateProcess`, never `Shell` (blocked the IDE ~9 s, spike 12).
- Status toolbar: version, DLL build time, port, request count (spike 3).

## Testing

`tests/Vb6Mcp.Tests` with our own runner (no framework package). Unit tests use `FakeAddin` (a loopback server speaking protocol v1) and fake environments; `IDE`-category tests need a running VB6 with the add-in.
