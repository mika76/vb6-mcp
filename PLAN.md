# VB6 MCP Server: Analysis & Plan

Status: design; spikes 10 + 11 done. Last updated 2026-10-13 (rev 8: + COM/GUID handling, §9.2).

## 0. Principles

1. **No external dependencies unless they're necessary.** Microsoft's own packages (e.g. `System.Drawing.Common`, the Windows Desktop SDK) are **not** external and may be used (owner, 2026-10-13). Only the .NET 10 BCL and Win32/COM that ship with Windows and VB6. That covers UI Automation, GDI screenshots and `LoadTypeLibEx`. No NuGet packages: no MCP SDK, SQLite, parser, diff or test libraries. A dependency gets in only when writing it ourselves would clearly cost more than it saves, and that has to be argued in the change.
2. **Test outcomes rather than trusting outside knowledge.** Anything about VB6/IDE/Windows behaviour that isn't directly checked is a **hypothesis**, marked **(H)** in this plan, until a spike shows it on a real machine. Record results in `docs/spikes/` and update this plan when reality differs.
3. **Build the smallest thing that solves a real problem**, and add more only when real use shows it's missing.
4. **The IDE is the primary workspace.** A developer works in the IDE: unsaved edits, a running program, break mode. The agent has to see and act on *that* state the way a developer sitting at the IDE would. Disk and the command line are the secondary channel, for closed projects, batch builds and CI-like work.

## 0.2 Other users (owner, 2026-10-13)

The server should work for **other people's VB6 setups**, not just this machine. Spike 16 showed the forced elevation here is this machine's configuration; a default VB6 install starts non-elevated. So: **support elevated and non-elevated VB6 equally** (the TCP + add-in design works at any integrity level), report each IDE's elevation, and offer `ide_launch` with `auto | elevated | normal`. Machine-specific values (VB6 path, workspace roots) come from configuration, never from code.

## 0.1 Context (from the owner)

- **Many VB6 apps** of every kind: Standard EXE, ActiveX DLL, OCX, small to large.
- **Source is shared between projects**, including files outside the project folder (`..\..\Shared\Utils.bas`).
- **VB6 runs as Administrator.** **Codepage is always 1252.**
- Decided: reload outside changes automatically when that's safe (§8).
- Code server: evaluated, reduced to an in-memory project map (§9).

## 1. The problem

| # | Problem | Why it hurts |
|---|---------|--------------|
| P1 | **The IDE's memory and the files on disk drift apart, in both directions.** Unsaved IDE edits are invisible on disk. Edits on disk are invisible to the IDE, which later overwrites them. | An agent that reads files sees stale code. An agent that writes files fights the IDE. |
| P2 | **Run/debug exists only inside the IDE.** The program runs *in-process*, on the IDE's UI thread. There's no debugger API. | Nothing outside the IDE can start, step, set breakpoints, inspect variables or see `Debug.Print`. |
| P3 | **The running app is a GUI.** VB6 apps are forms. Testing means clicking through them. | An agent can't check its change actually works without driving the UI. |
| P4 | **Compiling** shows errors as modal message boxes. Building a DLL/OCX registers it, which needs admin rights. | No structured errors. The dialogs freeze the IDE. |
| P5 | **References and components are COM registry entries.** Projects depend on each other's typelibs. | Easy to break; build order is invisible. |
| P6 | **The file formats are full of traps**: 1252 encoding, CRLF line endings, hidden `Attribute` lines, the designer block, binary `.frx` files, the line limit. | A naive edit corrupts files. |
| P7 | **Shared source** across many projects. | One edit affects N apps. |

## 2. Four channels to VB6

No single mechanism covers everything, so we combine four. Each MCP tool uses whichever channel works best for it.

| Channel | What it reaches | How | Status |
|---------|-----------------|-----|--------|
| **A. VBIDE object model** | Live code (including **unsaved** edits), components, dirty state, references, form designers, project properties, build events | In-process add-in → loopback TCP → MCP server | Members confirmed in `VB6EXT.OLB` (listed below); behaviour **(H)** |
| **B. IDE commands** | Run, Start With Full Compile, Break, End, Continue, Step Into/Over/Out, Run To Cursor, Toggle Breakpoint, Clear All Breakpoints, Save, Make | `VBE.CommandBars` → `FindControl(Id).Execute`, called from the add-in | **(H)**: control IDs and behaviour need a spike |
| **C. UI automation** | Everything else on screen: Immediate/Locals/Watch windows, Call Stack, modal dialogs (compile errors, runtime errors, MsgBox), **the running app's forms**, screenshots | Win32 (`EnumWindows`, `WM_GETTEXT`, `BM_CLICK`, `PrintWindow`) + UI Automation (in-box COM). **Reads** from the medium MCP server, **writes** (input) through the elevated add-in (spike 10) | Read/write split confirmed in spikes 10 + 11; details **(H)** |
| **D. Disk & command line** | Closed projects, batch builds (`VB6.EXE /make`), `.vbp` parsing, typelibs, registry | MCP server, headless (builds with registration started elevated by the add-in) | Mostly standard; log format **(H)** |

VBIDE members confirmed in `VB6EXT.OLB`:
- **Code**: `CodeModule.Lines / CountOfLines / InsertLines / DeleteLines / ReplaceLine / AddFromString / Find / ProcStartLine / ProcCountLines / ProcBodyLine / ProcOfLine / CountOfDeclarationLines / Members`
- **Code panes**: `CodePane.GetSelection / SetSelection / TopLine / Show`, `VBE.ActiveCodePane`, `SelectedVBComponent`
- **Sync**: `VBComponent.IsDirty / FileNames / SaveAs / Reload()`, events `ItemReloaded`, `FileControlEvents.*`
- **Project**: `VBProjects.AddFromFile/Remove`, `VBProject.SaveAs / IsDirty / MakeCompiledFile / BuildFileName / References`
- **Forms**: `VBComponent.Designer` → `VBForm.VBControls / SelectedVBControls`, `VBControl.Properties`
- **Mode**: `VBBuildEvents.BeginCompile / EnterRunMode / EnterDesignMode`. There's **no break event and no Mode property**.
- **UI**: `VBE.CommandBars`, `VBE.Windows` (each with `.HWnd`), `MainWindow.HWnd`

## 3. Elevation and transport (decided by spikes 10 + 11)

VB6 runs elevated. Claude Code, and the stdio MCP server it spawns, runs at medium integrity. Measured:

- **Spike 10**: medium → high **reads work** (screenshots, `WM_GETTEXT`). **Writes are blocked**: posted messages are refused, and synthetic clicks are silently dropped. This includes Claude's own computer use.
- **Spike 11**: medium → high over **TCP on 127.0.0.1 works**.

**Chosen: the add-in is the elevated endpoint and listens on loopback TCP.** The MCP server connects to it directly. No separate elevated daemon, no Scheduled Task, no named pipe. (Rev 4's daemon + pipe design is superseded.)

```
 Claude Code (medium) ── stdio MCP ──► vb6mcp (MCP server, medium, win-x86, .NET 10)
                                         │  all logic: routing, live-state tracker, diff/merge,
                                         │  project map, .vbp/1252 I/O, /make log parsing, typelibs,
                                         │  channel C reads (EnumWindows, WM_GETTEXT, screen capture)
                                         │
                                         │ TCP 127.0.0.1:<random port>, token per instance
                                         ▼
        VB6.EXE (elevated)  +  VB6McpBridge add-in    × N instances
          · Winsock listener (WSAAsyncSelect → hidden window → IDE thread, event-driven)
          · A: VBIDE object model   · B: CommandBars   · C writes: input to IDE + debugged app
          · spawns elevated children when needed (VB6.EXE /make, watchdog) — inherit elevation, no UAC
                │ starts if not running                    ▲ heartbeat: mode, dirty components,
                ▼                                          │ recovery copies (when idle)
        vb6mcp watchdog (elevated sidecar, one per user session, outlives VB6) ─┘
          · hang/crash detection for every VB6 · Ctrl+Break · dialog response · kill (force only)
          · recovery store for unsaved code · TCP 127.0.0.1 + token, own discovery file
          ◄── MCP server talks to it only when an add-in stops answering, or for recovery
```

- **Discovery**: on connect, each add-in writes `%LOCALAPPDATA%\vb6mcp\instances\<pid>.json` with `{pid, port, token, projects}`. The directory's ACL allows the current user only. The file is deleted on disconnect. The MCP server lists the folder to find IDEs, and ignores files whose PID is dead.
- **Wire protocol**: length-prefixed messages; every request carries the token. The encoding should be easy for VB6 to parse, so the add-in stays thin (decide in spike 2: minimal JSON vs. a simple field format).
- **Channel C is split by direction.** *Reads* (window lists, text, dialog detection, screenshots) are done by the medium MCP server directly, and keep working when the IDE thread is busy. *Writes* (clicks, keys, setting text) go through the add-in, inside the elevated process.
- **What still needs elevation outside the add-in**:
  - *Builds with registration*: the add-in starts `VB6.EXE /make` as a child, which inherits elevation. When no IDE is open: either one UAC prompt per build, or the user starts an elevated build. Rare in practice; decide later.
  - *IDE thread stuck* (busy loop in the debugged program, or a real hang): the add-in can't answer, and the medium server can only *read*. That's the **watchdog**'s job (§3.1).
- **Security**: bind `127.0.0.1` only (never `0.0.0.0`), random port, 128-bit random token per instance, discovery file readable only by the user. The add-in executes a **fixed list of verbs**, with no "run arbitrary command". Input, build/registration and launch actions are logged. The exposure exists only while VB6 is running.
- One binary, two modes: `vb6mcp` (the MCP server, default) and `vb6mcp watchdog` (§3.1). Hand-written JSON-RPC on `System.Text.Json`. **win-x86** for the same registry view as VB6.

### 3.1 Watchdog (elevated sidecar)

Elevation leaves one gap: when the IDE thread is blocked, the add-in can't act, and the medium MCP server can only look. The watchdog fills that gap.

- **Start and lifetime**: the add-in starts `vb6mcp watchdog` on connect if it isn't already running (a named mutex makes it a single instance per user session). As a child of the elevated VB6 it **inherits elevation, with no UAC prompt** **(H)**. It **keeps running if VB6 hangs, crashes or is killed** **(H)**. It exits a few minutes after the last VB6 is gone.
- **Sidecar, not proxy**: the MCP server talks to add-ins directly. It turns to the watchdog only when an add-in stops answering, or for recovery. **If the watchdog isn't running, everything else still works**, just without hang handling.
- **Polling, not heartbeats** (decided in Phase 2b): the watchdog polls each add-in (projects, components, dirty code) the way the MCP server does, instead of the add-in pushing heartbeats. A push would need a non-blocking Winsock client in VB6, and the add-in stays thin. Mode comes from the IDE title, which is readable even while the IDE is busy (spike 2).
- **Health checks**: `SendMessageTimeout(WM_NULL)` / `IsHungAppWindow` on each VB6 main window, plus heartbeat age and dialog detection (`EnumWindows`). States: `responsive`, `busy-running` (run mode, not pumping), `blocked-dialog`, `hung`, `crashed`.
- **Actions** (fixed verbs, audited):
  - `ide_break(pid)`: drop a busy loop into break mode from outside the IDE thread: the IDE's Break command (toolbar button), or Ctrl+Break input as a fallback **(H)**. Many keyboards (e.g. Mac) have no Break key, so don't depend on the keystroke.
  - `ide_dialog_respond(pid, button)`: works even while the IDE thread is blocked.
  - `ide_kill(pid, force: true)`: last resort. Never automatic, and never without `force`. Reports which unsaved components have recovery copies, and when they were taken.
- **Screenshots** (spike 15): elevated `PrintWindow` captures the IDE, dialogs and running forms **even when covered**, without changing focus, including controls that have no hwnd (Label). The non-elevated side can't (`PrintWindow` blocked; a screen copy only sees what's on top). So the watchdog serves `screenshot(pid, target, maxWidth)`: IDE panes are cropped from one capture of the main window; `PrintWindow` runs with a timeout, falling back to a screen copy if VB6 hangs. Results go back as MCP image content.
- **Escalation ladder** (the agent drives it; the watchdog never escalates on its own): report state → Ctrl+Break → End (through the add-in, once it responds) → kill with `force`.
- **Recovery store**: when the IDE is idle, the add-in sends the text of changed **unsaved** components (code, plus designer properties where it can get them) to the watchdog. The watchdog keeps them under `%LOCALAPPDATA%\vb6mcp\recovery\<pid>\` (user-only ACL), and prunes them when the component is saved or the IDE closes cleanly. After a crash or kill: `recovery_list`, `recovery_diff` (against disk), `recovery_restore` (into a reopened IDE, through the add-in). VB6 has no autosave, so this is useful even for users not working with an agent.
- **Crash detection**: a discovery file whose process is gone, without a clean unload (a clean `OnDisconnection` deletes the file), is reported as `crashed`. A user closing VB6 normally is never treated as a crash.
- **Guided restart + recovery (owner, 2026-10-13)**: the watchdog keeps the **last session** per IDE (pid, open `.vbp`/`.vbg` paths, dirty components, recovery copies).
  - `ide_launch(project)`: the watchdog starts `VB6.EXE "<project>"`. It's elevated, so the new IDE inherits elevation with **no UAC prompt** (**confirmed, spike 12b**: High integrity, window in 1.5 s). It watches for startup dialogs and reports them, and refuses when the project is already open in a live IDE (VB6 shows "already open").
  - `ide_recover`: lists what was unsaved at the time of the crash or kill, diffs it against disk, and on request writes it back into the reopened IDE through the add-in (components end up dirty; the user decides whether to save).
  - **No automatic restart by default** (rule 12; avoids crash loops and restarts after a deliberate close; startup dialogs such as missing references or source-control prompts would be unattended). Optional later: `autoRestart: crash-only`, rate-limited, never after a clean close.

## 4. Live IDE state: what the agent sees

**Rule: if a project is loaded in an IDE, the agent reads and writes the IDE's memory, never the file.** Unsaved edits are the truth.

- **Reading**: `code_read`, `proc_get` and `code_search` go to `CodeModule` via channel A, so they return what's in the IDE, saved or not. Each result says `source: ide|disk` and `dirty: true|false`.
- **Run/break mode (spikes 2, 3)**: VBIDE is **not** locked. Only `For Each` over `VBProjects`/`VBE.Windows` raises error 70. **Indexed access works in every mode**, so live reads work while running and debugging. Designer/property access fails at runtime (error 393) and must be avoided in run mode. Writes in break mode: allowed by the API, but structural edits there contributed to a deadlock (spike 3), so the server refuses edits outside design mode by default.
- **Never execute user code under the add-in (spike 3)**: anything that can run user code (IDE commands, form/property access) must not run synchronously inside the socket handler. Commands are queued with `PostMessage` to the add-in's own window and run from a fresh dispatch. Otherwise a break inside that code leaves the IDE in break mode with every command disabled, which needed a Task Manager kill.
- **Writing**: edits go to `CodeModule` and appear in the IDE immediately, unsaved, like a human edit (undo history is **(H)**). Saving is a separate, explicit step (`project_save`, or `save: true` on the write).
- **What the user changed**: VBIDE has no text-changed event. The MCP server keeps a **last-seen snapshot** of each component the agent has read. `ide_changes` returns a diff of IDE memory against that snapshot ("the user edited `Form1.cmdOK_Click` since you last looked"), using hashes and our own line diff. It's cheap enough to poll on request.
- **Unsaved vs disk**: `sync_diff(component)` shows IDE memory against the file on disk, and `ide_status` lists every dirty component.
- **The user's context**: `ide_context` returns the active project and component, the cursor position and selection (`CodePane.GetSelection`), the procedure at the cursor (`ProcOfLine`), the selected controls in an open designer, and the current mode. That makes "fix the function I'm in" possible.
- **Forms with unsaved designer changes**: read and change them through `VBForm`/`VBControl.Properties` (channel A), not the `.frm` file.

### Building and running *unsaved* code

`VB6.EXE /make` only sees disk, so for a dirty project:

| Goal | How |
|------|-----|
| Find compile errors in current memory | **Start With Full Compile** (channel B), then the dialog watchdog (C) reads the compile-error box and `GetSelection` gives the highlighted line. Then End and return the error. **(H)** |
| Run or debug what's in memory | Run in the IDE (channel B). That's what a developer does. |
| Produce a binary from memory | `MakeCompiledFile` (A) with the watchdog active **(H)**, or save first and then `/make` (D) |
| Build without touching the user's files | Shadow copy: write the IDE's memory into a temp copy of the project, then `/make` there. `.frm` designer changes need care (spike: is `SaveAs` safe or does it rebind the component's path? **(H)**) |

## 5. Run & debug

The IDE has no debugger API, so the debugging tools combine channels B and C. Every item here is **(H)** until spiked. Spike 5 is the most important one in the project.

| Capability | Mechanism |
|------------|-----------|
| Start / Full-compile start / Break / Continue / End / Restart | B: command bars |
| Step Into / Over / Out, Run To Cursor | B, plus `SetSelection` for Run To Cursor |
| Breakpoints: set / clear / clear all | `SetSelection(line)` then **Toggle Breakpoint** (B). VBIDE can't list breakpoints, so the MCP server keeps its own list, which can drift if the user toggles by hand. Clear All resets it. |
| Mode (design / run / break) | **IDE main window title** `… [design]` / `[run]` / `[break]` (class `wndclass_desked_gsk`), readable by the non-elevated server without the add-in (**confirmed, spike 2**), plus `EnterRunMode`/`EnterDesignMode` events |
| Current line when stopped | Active code pane + `GetSelection` in break mode |
| **Evaluate expression / run statement** | C: type `? expr` + Enter into the **Immediate window** (handle from `VBE.Windows`), read the new text back |
| `Debug.Print` output | C: read the Immediate window text. For compiled EXEs, capture `OutputDebugString` (DBWIN). |
| Locals / Watches | C: read the Locals/Watch windows (UIA or Win32). Fallback: evaluate each variable through the Immediate window. |
| Call stack | C: open Call Stack (Ctrl+L), read the list, close it |
| Runtime error ("Run-time error '91'… End / Debug") | Watchdog (C) catches the dialog. The tool returns the error number and text. `debug_runtime_error(action: debug|end)` presses the button; Debug lands in break mode at the faulting line. |
| User's program shows a MsgBox | Watchdog reports it (caption, text, buttons) and the agent picks a button |

**Threading caveat**: the program runs on the IDE's thread. While it's busy (no message pumping), channel A calls hang. Channel B and channel C writes go through the add-in, so they're blocked too. Channel C **reads** from the MCP server (window list, dialog text, screenshots) keep working, so the server can always report what's on screen. Every call has a timeout and reports "running, busy". Modal dialogs pump messages, so the add-in probably still answers while one is open **(H)**. Measure it.

## 6. Driving the running app (UI testing)

When a VB6 app runs from the IDE, its forms are windows of `VB6.EXE`. A compiled EXE is its own process. Either way the MCP server reads them directly (window tree, text, screenshots), and input goes through the elevated add-in. For a compiled EXE started non-elevated, the server could send input itself; for one started elevated by the add-in, input goes through the add-in:

- `app_windows`: top-level forms of the target process (caption, class, hwnd, rect)
- `app_tree(window)`: control tree. Intrinsic controls are real windows with readable classes (`ThunderTextBox`, `ThunderCommandButton`, `ThunderRT6*` when compiled) and text. UIA fills in what Win32 can't. Lightweight controls (Label, Image, Shape, Line) **have no hwnd**, so they're visible only in screenshots. **(H)**
- `app_click(control)`, `app_mouse(action, x, y)`, `app_scroll`, `app_set_text(control, text)`, `app_keys(keys)`, `app_select(combo/list, item)`. Mouse input is posted (doesn't disturb the user); Labels/Images react only to real input, so `app_mouse` has an explicit `real_input` (spike 17)
- `app_screenshot(window)`: `PrintWindow`/GDI → PNG, returned as **MCP image content**
- `app_wait(condition)`: a window appears, text matches, or a timeout passes
- `exe_run(path, args)`: launch a compiled build to test it the same way

## 7. Computer use

Claude can already drive the desktop with screenshots and clicks through computer use in the desktop app. How it fits:

| | Generic computer use (Claude) | Our structured UI tools (§5–6) |
|---|---|---|
| Works on | Anything visible | VB6 IDE + VB6 apps |
| Precision | Pixel coordinates from screenshots | Exact controls by hwnd/name; exact text |
| Cost/speed | A screenshot per step: slow, token-heavy | Small JSON, fast |
| Elevated VB6 | **Screenshots work; clicks are silently dropped** (UIPI). **Confirmed in spike 10.** | Works: input goes through the elevated add-in (spike 11 transport) |

**Plan (decided by spike 10)**: structured tools are the main path. `app_screenshot` gives the agent eyes for visual checks: layout, labels, owner-drawn controls. The add-in also provides low-level `ui_click(x,y)` / `ui_keys`: computer-use-style control routed through the elevated process, for anything we haven't modelled (third-party OCX dialogs, odd wizards). Generic computer use is useful only for **non-elevated** targets, such as a compiled EXE started normally. Every input tool must **verify its effect**, because a dropped input looks like success (see `docs/spikes/10-computer-use-elevation.md`).

## 8. Sync policy (disk ↔ IDE)

This matters only when files change *outside* the IDE (git checkout, another editor, another IDE instance with a shared file). Rule: reload automatically when it's safe, and never lose unsaved IDE work.

| Situation | Action |
|-----------|--------|
| Disk changed; component **clean**; design mode | Auto `Reload()` after a ~500 ms debounce (batches mass changes); notify |
| Disk changed; run/break mode | Queue the reload and apply it on `EnterDesignMode` |
| Disk changed; component **dirty** | **Conflict.** Touch nothing. Offer `sync_diff` / 3-way `sync_merge` (base snapshot from last load or save) / `sync_take_disk` / `sync_take_ide` |
| IDE about to save over a newer disk file | Veto through `RequestWriteFile` if possible **(H)**; otherwise snapshot the overwritten file |
| `.vbp` changed | Notify; reload the project only on request (it's disruptive) |
| File deleted or renamed | Notify only |

**Built (Phase 4):** the sync engine runs in the watchdog (`SyncPoller`, every 2 s). The IDE's own saves come from the add-in (`FileControlEvents.AfterWriteFile` → `WRITES`); a change must be stable for one more poll; reloads are deferred `RELOAD clean`, rechecked just before; conflicts are resolved with `sync_diff`, `sync_take_disk` (recovery copy first) and `sync_take_ide`. 3-way merge and the `RequestWriteFile` veto aren't built yet: a conflict is reported, and the user chooses.

A Claude Code `PreToolUse` hook denies direct `Edit`/`Write` on VB6 source files, so the agent uses the MCP tools instead. Those tools route to the IDE when it's open and handle the encoding.

## 9. Project map (what's left of the "code server")

The MCP server keeps in memory: file (normalized absolute path) → all `.vbp`s that include it, and per project: references, output binary and its typelib. That's enough for shared-source routing and rebuild order (`ws_file_projects`, `ws_affected`). No LSP. Symbols come from the IDE itself (§9.1), not from our own parser.

### 9.1 Symbol cache from the IDE (rev 7)

Rev 3 rejected a symbol index because it needed our own VB6 parser. **The IDE has already parsed the code**, so the add-in can just ask it. Plan, smallest first:

1. **Live outline** (planned): `CodeModule.Members` (`Name`, `Type`, `Scope`, `Static`, `CodeLocation`, `Description`, `Hidden`, confirmed in `VB6EXT.OLB`) + `ProcStartLine/ProcBodyLine/ProcCountLines` + the declaration line text for signatures + designer controls. Cached in the MCP server, keyed by component + hash of its code, and refreshed only when the hash changes. Works for **unsaved** code. Tools: `symbol_outline(component|project)`, `symbol_find(name)` across all open projects.
2. **Typelib cache** (planned): interfaces of referenced and in-house ActiveX DLLs/OCXs, read with our own `ITypeLib` reader and cached by path + timestamp. Backs `com_describe` and `com_search`. Probably the biggest win: the agent has never seen in-house components.
3. **Closed projects** (only if needed): a small declaration-only parser (module-level `Sub/Function/Property/Declare/Type/Enum/Const/Event/Dim`, line continuations). Cached as JSON per file, keyed by hash. **Tested against the IDE's `Members` output as an oracle.**
4. **Usages / call graph**: not planned. Text search plus the outline is good enough.

## 9.2 COM components and GUIDs (rev 8)

Most VB6 pain across many DLL/OCX projects comes from GUIDs and registration. Specific handling:

1. **Binary compatibility** (`CompatibleMode`, `CompatibleEXE32` in `.vbp`):
   - `compat_status(project)` reports the mode and the compat file. It warns on: an ActiveX DLL/OCX with no or project compatibility (new GUIDs every build); a missing compat file; the compat file being the build output itself instead of a reference copy; and the compat file's version vs the project's version.
   - **Never accept "break compatibility" automatically.** When the IDE shows the break-compatibility dialog, the dialog watchdog returns it as a decision for the agent and the user.
   - What `/make` does on a break (log error vs dialog) is **(H)**: spike 14.
2. **References resolve through the registry**, not the `.vbp` path: `*\G{GUID}#maj.min#lcid#path#name` → `HKCR\TypeLib\{GUID}\{ver}\{lcid}\win32` (32-bit view). `refs_add` takes a name or file and resolves the GUID and version. `refs_fix_broken` offers other registered versions of the same typelib GUID.
3. **OCX lines live in two places**: `Object=` in `.vbp` *and* `Object = "{GUID}#…"; "X.OCX"` in each `.frm`/`.ctl` header. Preserve them byte-exact; any change updates both consistently. Diagnostics flag stale `.oca` files.
4. **Registration state**:
   - `com_where(progid|clsid|typelib)`: every registered path and version, file exists?, timestamp, and whether it matches a known project's build output.
   - `com_diagnose(error)`: runtime errors 429/430/48 mapped to concrete checks (registered? file present? bitness? version mismatch?).
   - `com_orphans`: registrations pointing at missing files. **Report only**; cleanup is an explicit, confirmed action.
5. **Debugging ActiveX projects in the IDE**: while running, VB6 temporarily registers the project's classes against the IDE. If VB6 crashes or is killed mid-run, that registration may be left behind and break clients **(H)**. After a crash or kill, the watchdog checks for this and **offers** (never performs automatically) to re-register the compiled binary.
6. **All (un)registration is elevated**: `/make` (registers on success), `regsvr32`, `/unregserver`. It runs through the add-in or watchdog as a child process, never from the medium MCP server, and is audited. When a build's output path changes, unregister the old path first.
7. **Identity**: a component is identified by **typelib GUID + major version** (and CLSID for classes), never by file name. GUIDs come only from the registry or a typelib, never typed or generated by the agent.

## 10. MCP tool surface (draft)

- **IDE & context**: `ide_list`, `ide_status`, `ide_context`, `ide_changes`, `ide_dialogs`, `ide_dialog_respond(button)`
- **Health & recovery** (watchdog): `ide_health`, `ide_break`, `ide_kill(force)`, `ide_launch(project)`, `ide_recover`, `recovery_list`, `recovery_diff`, `recovery_restore`
- **Code** (IDE-first, disk fallback): `code_read`, `code_search`, `proc_list/get/replace/add/delete`, `code_replace_lines`, `code_declarations_get/set` (writes take `expected_hash`)
- **Project**: `project_info`, `project_open`, `project_save`, `component_add/remove/import/export/rename`
- **Forms (designer)**: `form_controls`, `form_set_property`, `form_add_control`, `form_remove_control`
- **Debug**: `run_start(fullCompile)`, `run_break`, `run_continue`, `run_end`, `step_into/over/out`, `run_to_cursor`, `bp_set/clear/clear_all/list`, `debug_eval(expr)`, `debug_immediate_read`, `debug_locals`, `debug_callstack`, `debug_where`, `debug_runtime_error(action)`
- **App UI**: `app_windows`, `app_tree`, `app_click`, `app_set_text`, `app_keys`, `app_select`, `app_screenshot`, `app_wait`, `exe_run`
- **Build**: `build` (disk `/make`, elevated), `build_in_ide`, `syntax_check` (memory, via full compile)
- **Sync**: `sync_status`, `sync_reload`, `sync_diff`, `sync_merge`, `sync_take_disk/ide`
- **Refs & COM**: `refs_list/add/remove/fix_broken`, `com_search`, `com_describe`, `com_where`, `com_diagnose`, `com_orphans`, `compat_status`, `com_register/unregister` (§9.2)
- **Workspace**: `ws_projects`, `ws_file_projects`, `ws_affected` (§9)
- **Symbols**: `symbol_outline`, `symbol_find` (§9.1)
- **Low-level UI**: `ui_click(x,y)`, `ui_keys` (needed: spike 10 showed computer use can't send input to elevated VB6)

## 11. Phases

### Phase 0: Spikes (throwaway code, findings in `docs/spikes/`)
1. ~~MCP handshake~~: **done**, works with Claude Code 2.1.295 in the desktop app. The client probes `server/discover` (protocol `2026-07-28`) and falls back to `initialize` on -32601; it sends `progressToken` and supports elicitation (`docs/spikes/01-mcp-handshake.md`).
2. ~~Bridge~~: **done**. Unsaved code read over TCP; the add-in answers in run/break mode and behind a modal MsgBox; the IDE title shows `[design]`/`[run]`/`[break]` (`docs/spikes/02-addin-bridge.md`). Its "VBIDE locked in run/break mode" conclusion was corrected by spike 3.
3. ~~Live state~~: **done**. Indexed VBIDE access works in all modes (only `For Each` fails); writes join undo; build-event handlers must not touch VBIDE; never execute user code under the add-in's socket handler (deadlock) (`docs/spikes/03-05-ide-probes.md`).
4. **Bridge under stress**: run mode (idle vs busy loop), break mode, modal dialogs. Timeouts for channels A, B and C.
5. ~~Debugging~~: **mostly done** (`docs/spikes/03-05-ide-probes.md`). Command Ids found (Start 186, Break 189, End 228, Step 188/194/2559, Toggle BP 51, …); Start/Break work; **Immediate evaluation works** (WM_CHAR + Enter + Select All/Copy + clipboard), `Debug.Print` readable the same way. Open: End failing in break mode, SetSelection-based commands, Locals/Call Stack. Original scope: command-bar IDs (Start, Full Compile, Break, End, Continue, Step ×3, Run To Cursor, Toggle/Clear Breakpoints); break-mode detection; current line; Immediate window write and read (`? x`); reading Locals and Call Stack; the runtime-error dialog's Debug button.
6. ~~Compile errors from memory~~: **done**. Deferred "Start With Full Compile" → watchdog reads `Compile error: Variable not defined` → `PRESS OK` → `GetSelection` = line/col of the error (`docs/spikes/06-08-compile-reload.md`).
7. **Running app UI**: `app_tree` of a running fixture form in the IDE and as an EXE; click, set text, screenshot via `PrintWindow`; which controls have no hwnd.
8. ~~Reload~~: **done**. `Reload()` refreshes code + designer of a clean component; **on a dirty component it silently discards the unsaved IDE edits**; refused in break mode (error 5). Deferred execution retest: Start→Stop→normal break with Continue/End enabled, deferred End works (rule 14 confirmed). Still open: `RequestWriteFile` veto, `SaveAs` rebinding. Original scope: `Reload()` on `.bas`/`.cls`/`.frm` (code and designer), dirty, break mode; `RequestWriteFile` veto; `SaveAs` rebinding.
9. **Disk builds**: `/make /out` log formats; elevated DLL/OCX registration. **Done so far:** 1252 round-trip is byte-exact for all 256 bytes; project-map scan of the real workspace (43 projects, 1,498 files, 139 shared) takes 4.9 s cold / 0.34 s warm, so the map stays in memory (`docs/spikes/09-project-map-scan.md`).
10. ~~Computer use vs elevation~~: **done**, input blocked (`docs/spikes/10-computer-use-elevation.md`).
11. ~~TCP across elevation~~: **done**, works (`docs/spikes/11-tcp-elevation.md`).
12. ~~Watchdog~~: **done**. Started by the add-in elevated with no UAC; survives VB6 being killed; busy-running vs hung classified from title + WM_NULL probe + `IsHungAppWindow`; **injected Ctrl+Break breaks a tight loop**; screenshot falls back to a screen copy while hung. Fixes: title regex (mode tag not always last), `CreateProcess` instead of `Shell` (blocked the IDE ~9 s), delete stale discovery files (`docs/spikes/12-watchdog.md`). Open: Ctrl+Break while VB6 is in the background; dialog response while blocked.
13. **Symbols from the IDE**: what `CodeModule.Members` returns (form members, events, controls, declarations), whether `CodeLocation` matches `ProcBodyLine`, speed on the largest real modules, and behaviour in break mode.
14. **COM edge cases** (binary compatibility part answered by spike 21: `/make` fails with a log error, no dialog; the IDE asks with a dialog): what `/make` does on a binary-compatibility break (log vs dialog); the break dialog in the IDE (`MakeCompiledFile`/full compile); what registry state a running ActiveX DLL project leaves after VB6 is killed, and whether re-registering the compiled DLL fixes it.
15. ~~Screenshots~~: **done**. Elevated `PrintWindow` captures the covered IDE, dialogs and running forms (incl. Labels); non-elevated can't. The watchdog hosts screenshots (`docs/spikes/15-screenshots.md`). Phase 3 moved elevated screenshots into the add-in (`SNAP`), so they work without the watchdog.
17. ~~Clicking windowless controls~~: **done**. Posted/sent mouse messages work for windowed controls but never raise `Click` on a Label/Image, not with the cursor moved, the form in front or a faked button state; only real input (`SendInput`) does. UIPI also blocks `SBM_*` reads of scroll bar controls from the server (`docs/spikes/17-lightweight-click.md`).
18. ~~Locals, Call Stack, runtime errors~~: **done**. Locals can't be read as text (custom-drawn); `TypeName` never errors, `CStr` of Null raises a dialog; the Call Stack is a dialog with a ListBox (read in-process, UIPI); the runtime error dialog's Debug stops at the failing line (`docs/spikes/18-locals-callstack-errors.md`).
19. ~~Watchdog live~~: **done**. Don't attach to the busy VB6 thread; `SwitchToThisWindow`/Alt bring an unresponsive VB6 forward; other add-ins' startup dialogs delay ours (`docs/spikes/19-watchdog-live.md`).
20. ~~Drag-and-drop, debug output~~: **done**. DBWIN capture works from the elevated IDE and compiled programs; classic drag needs real input; OLE drag didn't start from synthetic input here (`docs/spikes/20-drag-drop-debug-output.md`).
21. ~~Builds~~: **done**. `/make` refuses projects open in any IDE; its error lines are 0-based IDE lines; a compatibility break fails `/make` without a dialog but is a dialog in the IDE; `MakeCompiledFile` fails generically on compile errors (`docs/spikes/21-builds.md`).
22. ~~Sync engine live~~: **done**. Outside changes reloaded, conflicts kept, own saves recognised; fixes: the 4-parameter `AfterWriteFile` (`_dispVBFileControlEvents`), no project group on open, kept recovery copies in their own session (`docs/spikes/22-sync.md`).

### Phase 1: Skeleton + live code — **done** (2026-10-13)
MCP server, production add-in (TCP listener, discovery, token, verbs, deferred execution, status toolbar), read tools, editing with hash guards, `project_open`/`project_save`, `ide_dialog_respond`, `symbol_outline/find`, disk tools and the project map, `ide_launch` (auto/elevated/normal). See `docs/tools.md`. Not built: `ide_context`, `ide_changes` (fold into later phases when needed).

### Phase 2: Run & debug — **first round done** (2026-10-13)
Built and verified live: `run_start/break/continue/end`, `step_into/over/out`, `run_to_cursor`, `bp_toggle/clear_all`, `debug_where`, `debug_eval`, `debug_immediate`, `syntax_check`. Reading the Immediate window needed several fixes (commands only update when idle, Copy is asynchronous, Select All must be undone); see `docs/protocol.md`. Second round (2026-10-13): `debug_locals` (declarations + Immediate `TypeName`/`CStr`; the Locals window is custom-drawn), `debug_callstack` (Call Stack dialog read by the add-in), `debug_error` (runtime error dialog → Debug/End) (spike 18). DBWIN capture: `debug_output` (spike 20).

### Phase 2b: Watchdog
Watchdog mode, heartbeat, health states, `ide_break`, dialog response while blocked, `ide_kill(force)`, recovery store + `recovery_*`. Comes right after debugging, because debugging is what produces hangs.

### Phase 3: App UI automation — **first round built** (2026-10-13)
Built: `app_windows`, `app_tree`, `app_click`, `app_set_text`, `app_mouse` (posted by default, `real_input` for windowless controls), `app_scroll`, `app_screenshot` (MCP image; add-in capture for an elevated IDE), audit log. Verified live on `tests/fixtures/Ui`: text + `Change`, check box, ListBox double-click and scroll, scroll bar, PictureBox drag with exact coordinates, screenshots. Second round (2026-10-14): `app_select`, `app_keys` (posted, with deferred modifier state and continuation after each key down; or real input), `app_wait`, `exe_run`/`exe_end` with `exe` on every `app_*` tool (compiled programs, input from the server). Leftovers (2026-10-14, spike 20): `debug_output` (DBWIN) verified for the IDE and compiled programs; classic drag via `real_input` verified in the IDE; OLE drag not achievable with synthetic input on the test machine (a VM over RDP): documented, no tool.

### Phase 4: Builds, sync, shared source
Elevated `/make` + log parsing, `build_in_ide`, shadow builds; sync engine + reload policy + conflicts/merge; `ws_affected` rebuild order; Claude Code hook.

**Built (2026-10-14):** builds (spike 21), the sync engine (spike 22), `file_replace_lines` with the PreToolUse hook, and `ws_affected`, which matches references to projects by the LIBID and major version of the typelib in each DLL/OCX's build (`LoadTypeLibEx`, no registration). Not built: shadow builds, 3-way merge, and following group project references (`*\A`).

### Phase 5: Forms, references, COM
`form_*`, `refs_*`, `com_*`.

### Phase 6: Packaging
Installer: add-in registration (regsvr32 + `HKCU\...\Addins` + `vbaddin.ini`), config, Claude Code MCP config + hook. No service or Scheduled Task.

The phase order follows the principle that the IDE is the main workspace: live code first, then debugging, then UI testing. The disk-centric pieces come after.

## 12. Risks

| Risk | Likelihood | Mitigation |
|------|-----------|------------|
| Debug-window scraping (Immediate/Locals/Call Stack) proves unreliable | **High** | Spike 5 first. Fall back to Immediate-only evaluation, then to `Debug.Print`/DBWIN instrumentation |
| Calls hang during run mode or behind dialogs | High | Timeouts everywhere; watchdog; `ide_status` that never needs the bridge |
| Breakpoint list drifts from the IDE | Medium | Server-owned list + Clear All to reset; document it |
| Elevated add-in listening on TCP = attack surface (any local process can connect) | Medium | Loopback only, random port, per-instance token in a user-only file, fixed verbs, audit log; only while VB6 runs |
| IDE thread busy, so the add-in can't answer | Medium | Watchdog (§3.1): health states, Ctrl+Break, dialog response, kill with `force`; medium-side reads still report state |
| Watchdog becomes a second fragile component | Low–Med | Sidecar, not proxy: everything works without it; small fixed verb set; one per session |
| Killing VB6 loses unsaved work | Medium | Recovery store; kill only with `force`; result lists what's recoverable |
| Command-bar IDs differ by SP/locale | Low–Med | Look up by ID with a caption fallback; check at connect time |
| Add-in crash kills the IDE (and unsaved work) | Medium | Thin add-in, `On Error` everywhere, nothing long on the UI thread |
| Encoding corruption on the disk path | Medium | One 1252 I/O module, byte-exact tests, deny hook |

## 13. Open questions
- Workspace roots (where the `.vbp` files and shared folders live)? Network/mapped drives?
- Source control (git? SourceSafe?).
- 2–3 real projects as fixtures: ideally one form-heavy EXE, one ActiveX DLL with binary compatibility, one that uses shared source.
- Which third-party OCXs do the apps use most? They decide how much UI automation needs UIA versus screenshots.
