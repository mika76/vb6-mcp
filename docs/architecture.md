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
     │ CreateProcess (inherits elevation, no UAC)
     ▼
 vb6mcp.exe watchdog (elevated, one per user session; outlives VB6)
        TCP 127.0.0.1:<port>, same protocol; discovery %LOCALAPPDATA%\vb6mcp\watchdog.json
        acts when an IDE thread is stuck: Ctrl+Break, dialog buttons, kill, screenshots, launch
```

| Process | Integrity | Role |
|---------|-----------|------|
| Claude Code | medium | MCP client; starts `vb6mcp.exe` |
| `vb6mcp.exe` | medium | All logic. Talks MCP to Claude, TCP to add-ins, reads IDE windows |
| `VB6.EXE` + add-in | **high** (VB6 runs as Administrator) | Thin endpoint for VBIDE access |
| `vb6mcp.exe watchdog` | high (child of VB6) | Sidecar for stuck IDEs: Ctrl+Break, dialog buttons, kill, screenshots of a hung IDE, launching VB6 without UAC (spikes 12, 12b, 15). Not a proxy: everything else works without it |

### Why this split

- **Medium → high communication:** window messages and synthetic input are blocked (silently dropped, spike [10](spikes/10-computer-use-elevation.md)); **loopback TCP works** (spike [11](spikes/11-tcp-elevation.md)). So the add-in, already elevated inside VB6, listens on TCP.
- **Reads cross the boundary:** window titles, dialog text and window enumeration work from medium (spike 10). So IDE state (mode, dialogs, responsiveness) is read directly by `vb6mcp.exe`, even while the IDE thread is busy.
- **VBIDE is only reachable in-process:** VB6 doesn't register its `VBE` object in the ROT.
- **The watchdog** exists because the add-in shares the IDE thread: when the debugged program loops or the IDE hangs, the add-in can't answer, and the medium server can only look. Started by the add-in, it inherits elevation (spike 12) and serves from its own threads.

### Security

The add-in and the watchdog run elevated and serve loopback TCP; both check a 128-bit token from a discovery file in the user's profile, accept a fixed list of verbs, and act only on VB6 processes (hard rule 8). An elevated process that loads or starts files from a user-writable folder can be abused by other programs running as the same user. In the development layout the add-in DLL (`addin\VB6McpBridge\build`) and the watchdog (`build\watchdog` next to it) are in the repository, which the user can write. Starting the watchdog from the DLL's own folder adds no new exposure, since the DLL itself is already loaded from there. Installing both to a folder only administrators can write (Program Files) closes this, and belongs to packaging (Phase 6).

## Components in `src/`

### `Vb6Mcp.Core`

| Type | What | Notes |
|------|------|-------|
| `Vb1252` | Windows-1252 encoding | VB6 is ANSI (hard rule 1) |
| `Files.VbSourceFile` | splits `.bas/.cls/.frm/.ctl/.pag/.dsr` into header (VERSION, Object=, designer/class block, leading Attributes), code as the IDE shows it, and inline `Attribute` lines inside procedures; keeps the raw bytes | line counts match the IDE's for the Basic fixture; all 256 bytes round-trip through 1252 (hard rules 1, 2) |
| `Files.PathKey` | file identity: absolute, `..\` resolved, either slash, subst → target, mapped drive → UNC, case-insensitive key | hard rule 3 (shared source) |
| `Files.VbpFile`, `Files.VbgFile` | read-only `.vbp`/`.vbg` view: name, type, members with resolved paths, raw `Reference=`/`Object=` values, output file, compatibility mode/file | GUIDs kept raw, never rewritten (hard rule 13) |
| `Files.ProjectMap` | in-memory index of all `.vbp` under the roots: file → projects, shared files, roots check | 43 projects / 1,498 files in 4.9 s cold, 0.34 s warm on the owner's workspace: no cache file needed (spike [9](spikes/09-project-map-scan.md)) |
| `Vb6Install` | finds VB6.EXE (`VB6MCP_VB6_EXE`, registry `ProductDir`, `.vbp` association, default folder); reads a process's integrity level | no machine-specific paths in code (PLAN §0.2, spike 16) |
| `VbLocals` | names declared in one procedure (signature, return value, Dim/Static/ReDim) for `debug_locals` | the Locals window is custom-drawn (spike 18); declarations only, no expression parsing |
| `Recovery.RecoveryStore` | recovery copies of unsaved code: `%LOCALAPPDATA%\vb6mcp\recovery\<pid>\` (session + per component text and metadata) | written by the watchdog, read by the server; atomic writes |
| `Bridge.WatchdogDiscovery` | `watchdog.json` (same fields as an add-in discovery file) | so `AddinClient` talks to the watchdog unchanged |
| `VbCode` | module hash, validation of text to insert (CR LF, 1023 chars/line, 24 continuations, strict 1252), slicing, procedure-range trimming | hard rules 1, 10 |
| `Protocol.Wire` | frames, request building, response/record parsing | [protocol.md](protocol.md) |
| `Bridge.Discovery` | scans instance files; live vs stale (crashed/killed) | stale = process gone or PID reused (spike 12) |
| `Bridge.AddinClient` | one request per connection, timeout | timeout → `AddinBusyException` (hard rule 6) |

### `Vb6Mcp.UI`

| Type | What | Notes |
|------|------|-------|
| `IdeWindows` | IDE main window title → mode, `WM_NULL` probe, `IsHungAppWindow`, `#32770` dialogs, Immediate window, main window handle | title tag not always last (spike 12); cached caption readable while busy |
| `AppWindows` | the running program's windows (top-level windows of the VB6 process minus the IDE's own classes) and their controls (child walk), `Thunder*`/`ThunderRT6*` class → kind, `WM_GETTEXT` with a timeout | reads cross elevation (spike 10); Labels/Images have no hwnd |
| `Screenshots` | `PrintWindow` capture (non-elevated VB6 only), BMP load, PNG with max width (`System.Drawing`) | elevated VB6 is captured by the add-in instead (spike 15) |

### `Vb6Mcp` (`vb6mcp.exe`)

| Type | What | Notes |
|------|------|-------|
| `Mcp.McpServer` | JSON-RPC loop, initialize/ping/tools | hand-written, no SDK (spike [1](spikes/01-mcp-handshake.md)); unknown methods → -32601 (the desktop client probes `server/discover`); requests handled concurrently |
| `Mcp.Tool`, `ToolResult`, `Schema` | tool definitions and results | expected failures → `isError` results |
| `Tools.IdeTools` | read tools (`IdeTools.cs`), editing tools (`IdeTools.Edit.cs`), `ide_launch` (`IdeTools.Launch.cs`: ShellExecute for auto/elevated, RunAsInvoker for normal), run/debug (`IdeTools.Debug.cs`: command, then observe mode/dialogs; tracks runs it started and breakpoints it toggled), inspection (`IdeTools.Inspect.cs`: locals via declarations + Immediate `TypeName`/`CStr`, call stack via the dialog, runtime errors), the running program (`IdeTools.App.cs`: reads by the server, input and elevated screenshots through the add-in, effect verified afterwards) | [tools.md](tools.md); edits: module hash + old-text guard, design mode only |
| `Tools.WorkspaceTools` | disk tools over `VB6MCP_ROOTS`: project map, file → projects, `file_read` (with "open in an IDE?" check) | paths outside the roots refused (hard rule 8) |
| `Tools.Audit` | append-only `%LOCALAPPDATA%\vb6mcp\audit.log`: app input, `debug_eval`, `ide_launch` | hard rule 8; opened shared, never fails the tool |
| `Tools.IdeTools` (Watchdog, Recovery) | `ide_kill`, the watchdog fallbacks of `run_break`, `ide_dialog_respond`, `app_screenshot`, `ide_launch`; `recovery_*` | the watchdog is optional: tools say when they need it |
| `Watchdog.WatchdogHost` / `WatchdogActions` | the watchdog mode: TCP endpoint (own threads), BREAK/PRESS/KILL/SNAP/LAUNCH limited to VB6 processes | spike 12; Ctrl+Break only with VB6 in front |
| `Watchdog.RecoveryPoller` | polls add-ins every 10 s for unsaved components, only while the user is idle and the IDE isn't running the program | hard rule 12; polling instead of add-in heartbeats keeps the add-in thin |
| `Tools.IIdeEnvironment` | discovery + window reads + clients | swapped for fakes in tests |

## The add-in (`addin/VB6McpBridge`)

VB6 ActiveX DLL, ProgID `VB6McpBridge.Connect`, version 0.1.0, binary-compatible with `addin/compat/VB6McpBridge.dll`. Built and registered by `scripts/build-addin.ps1` (elevated); registered with VB6 by `scripts/register-addin.ps1` (HKCU, called by the build after the DLL's ProgID exists: a VB6 that starts while the Addins key names an unregistered ProgID deletes the key).

| Module | What |
|--------|------|
| `Connect.cls` | `IDTExtensibility`: starts/stops listener and status toolbar |
| `Listener.bas` | Winsock listener, protocol v1 framing, buffered output (`FD_WRITE`), discovery file (temp + rename; deleted on clean unload), token check |
| `Verbs.bas` | `PING`, `PROJECTS`, `COMPONENTS`, `READ`, `MEMBERS`, `PROC`, `OPEN`, `REPLACE`, `SAVE`, `CLICK`, `CMD`, `SETSEL`, `WHERE`, `IMM`, `IMMSELECT`/`IMMCOPY`/`IMMPEEK`/`IMMRESTORE`, `CALLSTACK`, `LISTITEMS`, `APPCLICK`, `APPTEXT`, `SNAP`, `MOUSE`, `SCROLL`, `SCROLLPOS`, `MOUSEREAL`, `MOUSERESTORE`; errors → 400/404/409/500 |
| `Defer.bas` | queue + `WM_DEFER` message: runs deferred actions (`OPEN`, `SAVE`, `CMD`, menu items by path (full compile, Call Stack), `IMM`, `APPTEXT`, `SNAP`) from a fresh dispatch |
| `Input.bas` | input into the IDE's own process only: `BM_CLICK` to dialog and program buttons, `WM_SETTEXT` to program text boxes, posted mouse and scroll messages, real input (`SendInput`) for windowless controls with cursor/foreground restore (spike 17), `WM_CHAR`/Enter to the Immediate window, clipboard save/restore |
| `Sidecar.bas` | starts the watchdog (`App.Path\watchdog\vb6mcp.exe watchdog`) with `CreateProcess`, no wait |
| `Snap.bas` | `PrintWindow` of an own-process window into a 32 bpp BMP under `%LOCALAPPDATA%\vb6mcp\shots\` (temp + rename) |
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
