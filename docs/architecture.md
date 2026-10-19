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

The add-in and the watchdog run elevated and serve loopback TCP; both check a 128-bit token from a discovery file in the user's profile, accept a fixed list of verbs, and act only on VB6 processes (hard rule 8). An elevated process that loads or starts files from a user-writable folder can be abused by other programs running as the same user. In the development layout the add-in DLL (`addin\VB6McpBridge\build`) and the watchdog (`build\watchdog` next to it) are in the repository, which the user can write. Starting the watchdog from the DLL's own folder adds no new exposure, since the DLL itself is already loaded from there.

**Installed layout (Phase 6, spike [26](spikes/26-packaging.md))** closes this. `install.ps1` puts `VB6McpBridge.dll`, `vb6mcp.exe` (server and watchdog) and `roots.txt` in `%ProgramFiles(x86)%\vb6mcp`, where only administrators can write:
- **The add-in** reads `roots.txt` next to its DLL (`Verbs.RootsFile`) and starts the watchdog from `vb6mcp.exe` next to it (`Sidecar.StartWatchdog`).
- **The server** reads `roots.txt` next to its exe when `VB6MCP_ROOTS` isn't set.

The roots decide what the elevated side may build, register and load. With a user-writable roots file, any program of the user could widen them and get files registered elevated: a UAC bypass.

**Verified:**
- An ordinary process can't change `roots.txt`, add a file or rename the exe.
- With `C:\Windows` added to the user-writable `%LOCALAPPDATA%` roots file, the installed add-in still refused a path under `C:\Windows`.

A development build (no `roots.txt` next to the DLL) keeps reading `%LOCALAPPDATA%\vb6mcp\roots.txt`.

**What the token and the approvals don't do.** The token keeps other users and other machines out. It doesn't keep out processes of the same user: they can all read the discovery file, so any of them can talk to the elevated add-in.

The verbs' path checks (roots) still apply to such a process. The Immediate verb (`IMM`), however, takes any single line, so the side-effect check of `debug_eval` and its approval (hook or MCP elicitation, by client name, which the client reports itself) only constrain agents that go through the server. On top of that, the project's own code runs elevated whenever the program is started from the IDE.

So vb6-mcp guards against mistakes by agents using the tools, not against code running as the user. That limit comes from VB6 running as administrator (spike 40, docs/tools.md `debug_eval`).

## Components in `src/`

### `Vb6Mcp.Core`

| Type | What | Notes |
|------|------|-------|
| `Vb1252` | Windows-1252 encoding | VB6 is ANSI (hard rule 1) |
| `Files.VbSourceFile` | splits `.bas/.cls/.frm/.ctl/.pag/.dsr` into header (VERSION, Object=, designer/class block, leading Attributes), code as the IDE shows it, and inline `Attribute` lines inside procedures; keeps the raw bytes | line counts match the IDE's for the Basic fixture; all 256 bytes round-trip through 1252 (hard rules 1, 2) |
| `Files.PathKey` | file identity: absolute, `..\` resolved, either slash, subst → target, mapped drive → UNC, case-insensitive key | hard rule 3 (shared source) |
| `Files.VbpFile`, `Files.VbgFile` | read-only `.vbp`/`.vbg` view: name, type, members with resolved paths, raw `Reference=`/`Object=` values, output file, compatibility mode/file | GUIDs kept raw, never rewritten (hard rule 13) |
| `Files.ProjectMap` | in-memory index of all `.vbp` under the roots: file → projects, shared files, roots check | 43 projects / 1,498 files in 4.9 s cold, 0.34 s warm on the owner's workspace: no cache file needed (spike [9](spikes/09-project-map-scan.md)) |
| `Files.TypeLibInfo`, `Files.Affected` | typelib identity (LIBID, version) read with `LoadTypeLibEx(REGKIND_NONE)` on its own STA thread; `.vbp` `Reference=`/`Object=` parsing; affected projects and rebuild order (`ws_affected`) | references matched by LIBID + major, never by file name (hard rule 13); the reader returns the LIBID that VB resolves through `HKCR\TypeLib` (`AffectedTests`), and the LIBID VB itself wrote for a VB-built DLL (spike [23](spikes/23-designer-references.md)) |
| `Files.TypeLibDescriber`, `Files.TypeLibRegistry` | full typelib contents (types, members, parameters in VB syntax) for `vb6mcp describe` and `com_describe`; registered typelibs from `HKCR\TypeLib` (32-bit view, hex version keys) for `refs_add` and the COM tools | read without registering; one read at a time and never `ReleaseComObject` (OLE shares cached typelibs between threads, spike [24](spikes/24-com-registration.md)); used to read VB6EXT.OLB before writing designer code (spike [23](spikes/23-designer-references.md)) |
| `Files.ComRegistry`, `Files.Compat` | COM class registrations (CLSID → server file, ProgID → CLSID, ProgID search, all classes for `com_orphans`, VB6 debugger registrations, PE bitness, existence of bare system names and resource-suffixed typelib paths), 32-bit view; binary compatibility findings per project (`compat_status`) | read only; registration goes through the elevated add-in (hard rule 13); debugger registrations and their restore measured in spike [25](spikes/25-debug-registration-broken-refs.md) |
| `Vb6Install` | finds VB6.EXE (`VB6MCP_VB6_EXE`, registry `ProductDir`, `.vbp` association, default folder); reads a process's integrity level | no machine-specific paths in code (PLAN §0.2, spike 16) |
| `VbLocals` | names declared in one procedure (signature, return value, Dim/Static/ReDim) for `debug_locals` | the Locals window is custom-drawn (spike 18); declarations only, no expression parsing |
| `Recovery.RecoveryStore` | recovery copies of unsaved code: `%LOCALAPPDATA%\vb6mcp\recovery\<pid>\` (session + per component text and metadata) | written by the watchdog, read by the server; atomic writes |
| `Bridge.WatchdogDiscovery` | `watchdog.json` (same fields as an add-in discovery file) | so `AddinClient` talks to the watchdog unchanged |
| `LineDiff` | Myers line diff with unified output | for `sync_diff`; no third-party diff library (principle 1) |
| `SyncState` | `%LOCALAPPDATA%\vb6mcp\sync\state.json`: conflicts and recent sync events | written by the watchdog's sync poller, read by `sync_status` |
| `BuildLog` | parses `/make` logs: status, compile errors (0-based IDE lines → 1-based), compatibility errors | formats from spike 21 |
| `VbCode` | module hash, validation of text to insert (CR LF, 1023 chars/line, 24 continuations, strict 1252), slicing, procedure-range trimming | hard rules 1, 10 |
| `Protocol.Wire` | frames, request building, response/record parsing | [protocol.md](protocol.md) |
| `Bridge.Discovery` | scans instance files; live vs stale (crashed/killed) | stale = process gone or PID reused (spike 12) |
| `Bridge.AddinClient` | one request per connection, timeout | timeout → `AddinBusyException` (hard rule 6) |

### `Vb6Mcp.UI`

| Type | What | Notes |
|------|------|-------|
| `IdeWindows` | IDE main window title → mode, `WM_NULL` probe, `IsHungAppWindow`, `#32770` dialogs, Immediate window, main window handle, all child windows with rectangles (`Descendants`: designer windows for `form_screenshot`, `vb6mcp windows <pid>`), current window and client rectangles (stitching) | title tag not always last (spike 12); cached caption readable while busy |
| `AppWindows` | the running program's windows (top-level windows of the VB6 process minus the IDE's own classes) and their controls (child walk), `Thunder*`/`ThunderRT6*` class → kind, `WM_GETTEXT` with a timeout | reads cross elevation (spike 10); Labels/Images have no hwnd |
| `LocalInput` | input from the server into compiled programs started with `exe_run` (same integrity): posted click/text/mouse/scroll/keys/select, and `SendInput` real input with foreground and cursor restore | mirrors the add-in's `Input.bas`; only for processes the server started |
| `DebugOutput` | OutputDebugString capture (DBWIN_BUFFER mapping + DBWIN_BUFFER_READY/DATA_READY events) on a background thread, last 2,000 messages | one owner per session (DebugView competes); started on first `debug_output` |
| `Screenshots` | `PrintWindow` capture (non-elevated VB6 only), BMP load, PNG with max width (`System.Drawing`) | elevated VB6 is captured by the add-in instead (spike 15) |

### `Vb6Mcp` (`vb6mcp.exe`)

| Type | What | Notes |
|------|------|-------|
| `Mcp.McpServer` | JSON-RPC loop, initialize/ping/tools | hand-written, no SDK (spike [1](spikes/01-mcp-handshake.md)); unknown methods → -32601 (the desktop client probes `server/discover`); requests handled concurrently. Sends one request of its own, `elicitation/create` (`debug_eval` side effects for clients other than Claude Code, spike [40](spikes/40-other-harnesses.md)), and matches the responses by id. Keeps `clientInfo.name` and the elicitation capability from initialize. Sends `notifications/progress` every 10 s during calls longer than 5 s (when the client gave a `progressToken`), so a client's idle timeout can be short; after `server_restart`, refuses new calls at once ("restarting") instead of leaving them to the exiting process (spike [42](spikes/42-liveness.md)). Log lines carry the pid: servers of several sessions share the log |
| `Mcp.Tool`, `ToolResult`, `Schema` | tool definitions and results | expected failures → `isError` results |
| `Tools.IdeTools` | read tools (`IdeTools.cs`), editing tools (`IdeTools.Edit.cs`), `ide_launch` (`IdeTools.Launch.cs`: ShellExecute for auto/elevated, RunAsInvoker for normal), run/debug (`IdeTools.Debug.cs`: command, then observe mode/dialogs; tracks runs it started and breakpoints it toggled), inspection (`IdeTools.Inspect.cs`: locals via declarations + Immediate `TypeName`/`CStr`, call stack via the dialog, runtime errors), compiled programs (`IdeTools.Exe.cs`: `exe_run`/`exe_end`), the running program (`IdeTools.App.cs`: a `Target` is the IDE's program (input via the add-in) or an `exe_run` program (input via `LocalInput`); reads by the server, input and elevated screenshots through the add-in, effect verified afterwards) | [tools.md](tools.md); edits: module hash + old-text guard, design mode only |
| `Tools.WorkspaceTools` | disk tools over `VB6MCP_ROOTS`: project map, file → projects, `file_read` (with "open in an IDE?" check), `file_replace_lines` (disk edit, header byte-exact via `VbSourceFile.ReplaceCode`; routed to the IDE when the file is open), `ws_affected` (rebuild order, with where each project is open); `WorkspaceTools.Com`: `com_describe/search/where`, `compat_status`, `com_register/unregister` (through the add-in's `REGSVR`, audited, verified in the registry) | paths outside the roots refused (hard rule 8) |
| `Tools.Audit` | append-only `%LOCALAPPDATA%\vb6mcp\audit.log`: app input, `debug_eval`, `ide_launch` | hard rule 8; opened shared, never fails the tool |
| `Tools.IdeTools` (Watchdog, Recovery) | `ide_kill`, the watchdog fallbacks of `run_break`, `ide_dialog_respond`, `app_screenshot`, `ide_launch`; `recovery_*` | the watchdog is optional: tools say when they need it |
| `Watchdog.WatchdogHost` / `WatchdogActions` | the watchdog mode: TCP endpoint (own threads), BREAK/PRESS/KILL/SNAP/LAUNCH limited to VB6 processes | spike 12; Ctrl+Break only with VB6 in front |
| `Watchdog.SyncPoller` | every 2 s: own writes (add-in `WRITES`) vs file time stamps; reloads clean components (`RELOAD clean`), records conflicts, waits for design mode | PLAN §8; in the watchdog because outside changes happen without an agent session |
| `Watchdog.RecoveryPoller` | polls add-ins every 10 s for unsaved components, only while the user is idle and the IDE isn't running the program | hard rule 12; polling instead of add-in heartbeats keeps the add-in thin |
| `Hook` | `vb6mcp hook`: Claude Code PreToolUse hook blocking Edit/Write on VB6 files (exit 2 + reason); `install-hook` merges it into `~/.claude/settings.json`, `uninstall-hook` removes only our entry | PLAN §8; generic tools would write UTF-8 and bypass the IDE |
| `Tools.IIdeEnvironment` | discovery + window reads + clients | swapped for fakes in tests |

### Edits no tool made (`ToolEditLedger`)

One `ToolEditLedger` per server process, shared by `IdeTools` and `WorkspaceTools`, remembers the code each IDE component held right after the last tool edit (`code_replace_lines`, `code_insert`, `proc_replace`, `file_replace_lines` routed to the IDE, `recovery_restore`). Reads never update it.
- `project_save` (whole project and `components`) compares each dirty component's IDE text with it, or with the file's code when no tool edited it, and refuses a difference other than case unless `outside_edits: true`. `syntax_check` reports the same differences without refusing. The point: keystrokes that reach a code pane (a selection left on a compile error) or the user's typing must not be saved unseen (improvement notes 50).
- The other safety checks of the running-program tools sit next to their tools in `IdeTools.App.cs`: `app_keys` without `control` never targets the IDE's own windows (`ProgramFocusInIde`, note 51), and `app_click` refuses controls `AppWindows.Controls` marks `Clipped` (a pure `ClipState` over the ancestors' client areas, notes 33 and 52).

### Watchdog started by the user (`--stay`)

The add-in starts the watchdog (`vb6mcp watchdog`), which exits 2 minutes after the last VB6 closes, or 30 minutes after a crash (a discovery file left behind by a gone process), so the agent can still relaunch through it without a UAC prompt once it notices (PLAN.md D2). The server notices a crash before running the next tool and adds an `ideCrashed` notice to that tool's result (`IdeTools.Crash.cs`: only IDEs it saw running, each once; `ide_kill` and `ide_quit` don't count), with the Application-log entry read through `wevtutil`. `vb6mcp watchdog --stay`, started elevated by the user once, keeps running without any VB6.
- Because it is elevated, `ide_launch` can start elevated IDEs through it (`LAUNCH`, spike 12b) without a UAC prompt, so each Claude session can have its own IDE (PLAN Phase 7 A9).
- Its capabilities are unchanged: `LAUNCH` still starts only `VB6.EXE`, with project paths checked against the roots. It is only around for longer.

### IDE ownership between sessions (Phase 7 A9)

`IdeOwnership` (Core) records which MCP server process owns an IDE, in `%LOCALAPPDATA%\vb6mcp\owners\<ide pid>.json`. It is kept apart from `instances\`, because `Discovery` reads every `*.json` there. The record holds the server's pid and start time and the client name from `initialize`; a dead owner owns nothing.
- `IdeTools.All()` wraps every write tool (`IdeTools.WriteTools`) with a guard. The guard resolves the target IDE as the tool would and refuses one owned by another live session.
- `ResolveInstance` prefers the session's own IDE when several run.
- `ide_launch` claims what it starts. `ide_claim`/`ide_release` handle the rest.
- This is coordination between the user's own sessions, not a security boundary.

### Checkpoints

`Checkpoints` (Core, `Files/Checkpoints.cs`) copies a project's files (from its parsed `.vbp`) byte-exact to `%LOCALAPPDATA%\vb6mcp\checkpoints`, one folder per project keyed by its normalized path (hard rule 3), with a manifest of SHA-256 hashes. It only reads the project. `project_rollback` writes the files back; for an open project it reuses the `project_close` discard machinery (snapshot what VB writes on close, restore it) and then opens the project again, because VBIDE can't drop unsaved changes or reload a whole project (spike 31). No add-in change.

### `form_layout`

`FormLayout` (server, `Tools/FormLayout.cs`) turns a grid of control names into rectangles, from the controls' current sizes (one `DESIGN CONTROLS` read) and the container's inside size (`DESIGN PROPS` of the form or PictureBox; a Frame's own size). It is pure arithmetic, unit-tested without an IDE. The tool then sets only the changed `Left`/`Top`/`Width`/`Height` through the same `DESIGN SET` loop as `form_set_properties` (`ApplyChangesAsync`), so there is no new add-in verb and the usual overlap review runs afterwards.

### `batch_run`

`BatchTool` (built in `Program.cs` over every other tool) runs a list of tool calls in order through each tool's own `InvokeAsync`, so guards and checks apply per call. `ToolResult.Batch` concatenates their content. It saves model turns, not work in the IDE. It needs no add-in verb, so the add-in's verb list (hard rule 8) is unchanged.

### Busy, not unreachable

While the IDE thread runs a long operation the server queued (`MakeCompiledFile` for `build_in_ide` took ~30 s in a production project, or a full compile for `syntax_check`), the add-in can't accept connections and Windows refuses them. `AddinClient` remembers such operations per IDE pid (`MarkBusy`, at most 10 minutes, cleared by the next successful reply). A refused connection in that window, with the VB6 process alive, raises `AddinBusyException` ("VB6 is busy compiling (started 12 s ago)…"), which tools and polling loops already treat as busy, instead of "unreachable" (hard rule 6, improvement notes 21).

## The add-in (`addin/VB6McpBridge`)

VB6 ActiveX DLL, ProgID `VB6McpBridge.Connect`, version 0.1.0, binary-compatible with `addin/compat/VB6McpBridge.dll`. Built and registered by `scripts/build-addin.ps1` (elevated); registered with VB6 by `scripts/register-addin.ps1` (HKCU, called by the build after the DLL's ProgID exists: a VB6 that starts while the Addins key names an unregistered ProgID deletes the key).

| Module | What |
|--------|------|
| `Connect.cls` | `IDTExtensibility`: starts/stops listener and status toolbar |
| `Listener.bas` | Winsock listener, protocol v1 framing, buffered output (`FD_WRITE`), discovery file (temp + rename; deleted on clean unload), token check |
| `Verbs.bas` | `PING`, `PROJECTS`, `COMPONENTS`, `READ`, `MEMBERS`, `PROC`, `OPEN`, `REPLACE`, `SAVE`, `CLICK`, `CMD`, `SETSEL`, `WHERE`, `IMM`, `IMMSELECT`/`IMMCOPY`/`IMMPEEK`/`IMMRESTORE`, `CALLSTACK`, `LISTITEMS`, `APPCLICK`, `APPTEXT`, `SNAP`, `MOUSE`, `SCROLL`, `SCROLLPOS`, `MOUSEREAL`, `MOUSERESTORE`; errors → 400/404/409/500 |
| `Defer.bas` | queue + `WM_DEFER` message: runs deferred actions (`OPEN`, `SAVE`, `CMD`, menu items by path (full compile, Call Stack), `IMM`, `APPTEXT`, `SNAP`) from a fresh dispatch |
| `Input.bas` | input into the IDE's own process only: `BM_CLICK` to dialog and program buttons, `WM_SETTEXT` to program text boxes, posted mouse and scroll messages, real input (`SendInput`) for windowless controls with cursor/foreground restore (spike 17), `WM_CHAR`/Enter to the Immediate window, clipboard save/restore |
| `Sync.bas` | records the IDE's own file writes (`FileControlEvents.AfterWriteFile`, hooked in `Connect.cls`) for `WRITES` | event handlers only record |
| `Design.bas` | form designer verbs `DESIGN`/`RESULT`: controls, properties (object properties via their sub-`Properties` collection), add, remove, show (open the designer for `form_screenshot`) | deferred with request ids: a UserControl's own code runs when its properties are touched (hard rule 14, spike [23](spikes/23-designer-references.md)) |
| `Sidecar.bas` (registration) | `REGSVR`: 32-bit `regsvr32 /s [/u]` as an elevated child, handle kept for `REGSTATUS` (non-blocking exit code) | path checked on the elevated side too: roots, absolute, no `..`, .dll/.ocx (spike [24](spikes/24-com-registration.md)) |
| `Refs.bas` | `REFS`, `REFADD` (`AddFromFile`/`AddFromGuid`), `REFDEL`, `TOOLBOX` (`AddToolboxProgID`) | synchronous, design mode for changes; GUIDs only from the registry or a typelib (hard rule 13) |
| `Sidecar.bas` | starts the watchdog (`App.Path\watchdog\vb6mcp.exe watchdog`) with `CreateProcess`, no wait |
| `ProjectVerbs.bas` | `CLOSE` (with discard: `IsDirty` cleared, then `VBProjects.Remove`), `COMPADD` (`VBComponents.Add`/`AddFile`), `SAVECOMP` (`VBComponent.SaveAs`), all deferred; `PRJGET`/`PRJSET` (project properties, startup object), synchronous | Phase 7; paths checked against the roots here too |
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

## Packaging and installation (Phase 6)

| Piece | What | Why |
|-------|------|-----|
| `scripts\package.ps1` | `dist\vb6mcp-<version>\` + `.zip`: `vb6mcp.exe` (self-contained, single-file, compressed, win-x86, 48 MB), `VB6McpBridge.dll`, `install.ps1`, `uninstall.ps1`, `INSTALL.txt`; refuses an add-in build older than its source | no .NET needed on the target; the add-in can only be built elevated (spike [26](spikes/26-packaging.md)) |
| `installer\install.ps1` | runs as the user. One elevated pass: copies to `%ProgramFiles(x86)%\vb6mcp` (in-use files moved aside to `*.old-<time>`), writes `roots.txt`, stops a running watchdog, `regsvr32` (32-bit) and checks `InprocServer32` in the registry. Then as the user: HKCU add-in entry, `claude mcp add --scope user`, `vb6mcp install-hook` | admin-only folder for the elevated pieces (Security above); HKCU and Claude settings belong to the user, so they're written unelevated; Windows PowerShell 5.1 compatible |
| `installer\uninstall.ps1` | the reverse: hook (`uninstall-hook`), MCP registration and HKCU entry as the user; elevated: stop the watchdog, `regsvr32 /u`, delete the folder (files in use reported). It runs its elevated part from a temp copy, since it lives in the folder it deletes | refuses while VB6 runs unless `-Force`; `%LOCALAPPDATA%\vb6mcp` is kept unless `-RemoveData` |

Development keeps its own flow (`build-addin.ps1`, `install-server.ps1`); whichever ran last decides where the add-in's ProgID points.

## Testing

`tests/Vb6Mcp.Tests` with our own runner (no framework package). Unit tests use `FakeAddin` (a loopback server speaking protocol v1) and fake environments; `IDE`-category tests need a running VB6 with the add-in.
