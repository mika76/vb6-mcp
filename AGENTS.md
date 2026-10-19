# AGENTS.md

Guidance for AI coding agents working in this repo. Read `PLAN.md` for the architecture and why it was chosen.

## Principles (read first)

1. **No external dependencies unless they're necessary.** Use only the .NET 10 BCL, **Microsoft's own packages** (e.g. `System.Drawing.Common`, Windows Desktop SDK: the owner doesn't count them as external), Win32/COM APIs that ship with Windows, and VB6 itself. No third-party packages: no MCP SDK, JSON-RPC library, SQLite, diff library, parser generator or test framework. If you believe a dependency is truly necessary, stop and argue it in the change (what it saves versus writing it ourselves). Don't just add it.
2. **Test outcomes; don't trust outside knowledge.** Anything about VB6, VBIDE, COM, UAC or MCP behaviour that you "know" from docs, forums or training is a hypothesis. Prove it with a spike or test on this machine before code relies on it, and write the result up in `docs/spikes/`. If an experiment contradicts PLAN.md or this file, the experiment wins: update the docs.
3. **Build the smallest thing that solves a real problem.** No speculative features (LSP, usage/call-graph tracking, extra abstraction layers) until real use shows they're needed. Symbols come from the IDE's own parse (`CodeModule.Members`), not from a parser we write. A declaration parser for closed projects gets built only if needed, and must be tested against the IDE's output.
4. **The IDE is the primary workspace.** Unsaved edits, the running program and break mode are what the developer actually sees, so the agent works on that live state. Disk and the command line are the secondary channel, for closed projects and batch builds. If a design reads "file first, IDE as an afterthought", it's wrong.

## What this is

An MCP server that lets AI agents work inside the VB6 IDE the way a developer does: read and edit live (unsaved) code, run, debug, drive the running app's UI, and build. It works on many projects (EXE, ActiveX DLL, OCX), including source files shared between projects.

It reaches VB6 through **four channels** (PLAN.md §2):
- **A**: the VBIDE object model, through the add-in: live code, dirty state, designers, references.
- **B**: IDE commands through `CommandBars`: run, break, step, breakpoints, compile.
- **C**: UI automation (Win32 + UI Automation, screenshots): debug windows, dialogs, the running app's forms.
- **D**: disk and command line: closed projects, `/make`, `.vbp` parsing, typelibs.

It has three parts:

- **`vb6mcp`, the MCP server** (C#, .NET 10, win-x86, **medium integrity**: spawned by Claude Code over stdio, with hand-written JSON-RPC 2.0 on `System.Text.Json`). It holds all the logic: routing, live-state tracker, sync engine, diff/merge, project map, `.vbp`/1252 I/O, build-log parsing, typelibs, and channel C **reads** (window lists, text, dialog detection, screenshots).
- **`addin/VB6McpBridge`** (VB6 ActiveX DLL, runs inside **elevated** `VB6.EXE`). A thin endpoint: a Winsock listener on `127.0.0.1:<random port>` (`WSAAsyncSelect`, so events arrive on the IDE thread), a discovery file, token checks, and a fixed list of verbs for channel A (VBIDE), channel B (CommandBars) and channel C **writes** (input into the IDE and the app being debugged). When something needs elevation, it starts elevated child processes (`VB6.EXE /make`, an optional helper); they inherit elevation, so there's no UAC prompt.

Why it's split this way (measured, not assumed):
- VB6 runs as Administrator. A medium process can *read* elevated windows but can't *send* messages or input to them (spike 10).
- Loopback TCP from medium to high works (spike 11).

- **`vb6mcp watchdog`, the elevated sidecar** (same exe, different mode). The add-in starts it as a child, so it inherits elevation with no UAC prompt. There's one per user session, and it outlives VB6. It detects hung or crashed IDEs, sends Ctrl+Break, answers dialogs while the IDE thread is blocked, kills VB6 only with `force`, and keeps recovery copies of unsaved code. **It's a sidecar, not a proxy:** the MCP server talks to add-ins directly, and everything must keep working when the watchdog isn't running. See PLAN.md §3.1.

So the elevated add-in serves TCP, the elevated watchdog covers what happens when the IDE thread is stuck, and everything else stays unelevated. See PLAN.md §3.

Planned layout:

```
src/Vb6Mcp.Core/        file I/O (1252/CRLF), vbp/vbg + source parsers, path normalization,
                        project map, line diff/merge, build-log parser,
                        typelib reader (LoadTypeLibEx/ITypeLib), Win32 helpers
src/Vb6Mcp.UI/          channel C reads: window enumeration, control text, UIA, screen capture,
                        dialog watchdog, Immediate/Locals/Call Stack readers
src/Vb6Mcp/             the exe. Mode `mcp` (default): hand-written MCP (JSON-RPC over stdio),
                        tool handlers, routing, live-state tracker, sync engine, add-in client.
                        Mode `watchdog`: health checks, Ctrl+Break, dialog response, kill, recovery store
addin/VB6McpBridge/     VB6 add-in (.vbp, Connect.Dsr, Bridge/Listener/Verbs modules)
tests/Vb6Mcp.Tests/     tests + our own minimal runner (console exe)
tests/fixtures/         small VB6 projects: a form-heavy EXE (for debug/UI tests), a shared-source
                        layout, an ActiveX DLL w/ binary compat
spikes/                 Phase 0 throwaway experiments. Never referenced by src/
docs/spikes/            one write-up per spike
scripts/                development: build-addin, install-server; package.ps1 builds the user package
installer/              install.ps1 / uninstall.ps1 / INSTALL.txt, shipped in the package (Program Files install)
```

## Environment facts (this machine)

- VB6: `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`. Extensibility typelib: `VB6EXT.OLB` in the same folder.
- Add-ins are registered under `HKCU\Software\Microsoft\Visual Basic\6.0\Addins\<ProgID>`. That's enough on its own: `C:\Windows\vbaddin.ini` `[Add-Ins32]` was never needed in any phase.
- **Installed vs development** (spike 26):
  - The package installs to `%ProgramFiles(x86)%\vb6mcp`: `VB6McpBridge.dll`, `vb6mcp.exe` (server and watchdog), `roots.txt` (admin-only machine policy, read by the add-in and the server), `uninstall.ps1`.
  - Development builds (`build-addin.ps1`, `install-server.ps1`) use the repo and `%LOCALAPPDATA%\vb6mcp\roots.txt`.
  - Whichever ran last decides where the ProgID points. A local-scope `vb6` MCP registration in a folder and the installer's user-scope one are separate entries: remove one when testing the other.
- .NET SDK 10.x, Node 24. Shell: PowerShell 7 or Git Bash.
- VB6 is 32-bit and runs elevated (UAC consent prompt, same user account). The MCP server runs non-elevated and as x86 (32-bit registry view, same as VB6).

## Hard rules

1. **Encoding: Windows-1252, CRLF, no BOM. Always.** All `.bas/.cls/.frm/.ctl/.dsr/.vbp/.vbg` I/O goes through the one encoding module in `Vb6Mcp.Core`. Round-trips must be **byte-exact**, including bytes undefined in 1252. Remember `CodePagesEncodingProvider` in .NET. Add a round-trip test whenever you touch it.
2. **Preserve what the IDE hides.** Keep `Attribute VB_*` lines, the `.frm`/`.ctl` designer block and the `Object=`/`Reference=` lines in `.vbp` byte-exact unless the tool exists to change them. Never edit `.frx`/`.ctx`.
3. **File identity is the normalized absolute path** (case-insensitive, `..\` resolved, mapped drives and `subst` mapped to one canonical form). Never key anything by component name alone: the same name means different files in different projects, and one file can belong to many projects.
4. **The IDE's memory is the truth when a project is loaded.** Reads and writes go through the bridge and return `source: ide|disk` and `dirty`. Never read the file instead of the loaded component. Never write a file on disk that is dirty in any open IDE: raise a conflict. Never discard unsaved IDE work automatically. Saving is always an explicit action. Follow the sync policy in PLAN.md §8.
5. **Never block the IDE's UI thread.** Add-in methods are short: no `DoEvents` loops, no sleeps, no modal UI, no blocking socket calls (use `WSAAsyncSelect` only). Builds, scanning, diffing and merging belong in the MCP server or in child processes.
6. **Every bridge call has a timeout.** On timeout, return a structured "IDE busy" error with the reason (run mode, modal dialog). Never hang an MCP request.
7. **Add-in code: `On Error` in every public method.** Return structured errors. An add-in crash kills the user's IDE and loses their unsaved work.
8. **The add-in is an elevated TCP endpoint, so treat it as a security boundary.** Bind `127.0.0.1` only (never `0.0.0.0`), random port, a 128-bit random token per instance, and the discovery file in `%LOCALAPPDATA%\vb6mcp\instances\` with a user-only ACL. Reject any request without the token. Verbs are a fixed list, with no generic "run command" or "run process". Every path is checked against the configured workspace roots. Registration, builds, process launches and **all synthetic UI input** (`app_*` actions, `ui_*`, `debug_eval`) go to the audit log. UI input goes only to windows of the IDE's own process, of processes the add-in launched, or of compiled programs the server started with `exe_run` (those run at the server's own integrity, so no boundary is crossed), unless a tool explicitly says otherwise. Don't add verbs that weaken this without saying so in the change. `debug_eval` refuses expressions with side effects unless `VB6MCP_EVAL=open` (PLAN.md E5), and the Claude Code hook refuses the agent's shell commands that would get around the tools: database writes, VB6 source writes, registration, ending VB6 (E5b).
9. **Destructive actions need explicit parameters**, never defaults. That covers dismissing or answering dialogs, `com_register`, removing components, `sync_take_*`, project reloads and `run_end` while the user is debugging.
10. **Line limits**: at most 1023 chars per line, at most 24 `_` continuations. Validate before inserting code.
11. **Don't hijack the user's session.** Debug and UI tools act on what the user already has open; they don't silently start, stop or reset a session the user is in. Report the current mode before acting, and require `force` to end a run the agent didn't start.
12. **The watchdog never escalates on its own.** It reports state; the agent decides. `ide_kill` requires `force: true`, and its result must list which unsaved components have recovery copies and how old they are. Recovery copies are written only when the IDE is idle, never on a timer that interrupts typing or debugging.
13. **GUIDs are never hand-written.** Never create, type or edit a GUID in `Reference=`/`Object=` lines, `.frm`/`.ctl` headers or compatibility settings. GUIDs come only from the registry or a typelib. Identify components by typelib GUID + major version, never by file name. Never accept a "break binary compatibility" prompt automatically: surface it as a decision. (Un)registration only through the elevated add-in or watchdog, audited. See PLAN.md §9.2.
14. **Never run user code under the add-in's request handler** (spike 3). IDE commands (Start, Continue, Step, Break, End, Run To Cursor), Immediate input, and anything that can instantiate a form or run user code must be **deferred**: validate, reply "queued", `PostMessage` the add-in's own window, and execute from a fresh dispatch. The MCP server then observes the result (title, dialogs, Immediate text). A break inside code started from the handler left the IDE in break mode with every command disabled, and it had to be killed.

## VBIDE quick reference (checked against VB6EXT.OLB)

- Code: `CodeModule.Lines(start,count)`, `CountOfLines`, `InsertLines`, `DeleteLines`, `ReplaceLine`, `AddFromString`, `Find`, `ProcStartLine/ProcBodyLine/ProcCountLines(name, kind)`, `ProcOfLine`, `CountOfDeclarationLines`, `Members`
- Sync: `VBComponent.IsDirty`, `FileNames(i)`, `FileCount`, `SaveAs`, `Reload()`. Events: `VBComponentsEvents.ItemReloaded`, `FileControlEvents.AfterWriteFile/RequestWriteFile/BeforeLoadFile`
- Project: `VBProjects.AddFromFile/Remove`, `VBProject.SaveAs`, `MakeCompiledFile`, `BuildFileName`, `References.AddFromGuid/AddFromFile/Remove`, `Reference.IsBroken`
- Forms: `VBComponent.Designer` → `VBForm.VBControls`, `VBControl.Properties`
- Code panes / context: `VBE.ActiveCodePane`, `CodePane.GetSelection/SetSelection/TopLine/Show`, `SelectedVBComponent`, `VBForm.SelectedVBControls`
- Windows: `VBE.Windows` items have `.HWnd`. Use them to find the Immediate, Locals and Watch windows for channel C.
- Mode: `VBBuildEvents.BeginCompile/EnterRunMode/EnterDesignMode`. **There is no break-mode event, no Run/Stop method and no debugger API** (no breakpoint list, no evaluate). Debugging = channel B commands + channel C window reading.
- The `VBE` object is **not** in the ROT. It's only reachable from inside the add-in.

## Build & test

```powershell
dotnet build src/Vb6Mcp -c Debug
dotnet run --project tests/Vb6Mcp.Tests              # unit tests
dotnet run --project tests/Vb6Mcp.Tests -- --cat Build,IDE,Launch
```

- **Running in development**: the MCP server runs **non-elevated** (`dotnet run --project src/Vb6Mcp`, or registered in Claude Code). VB6 runs elevated with the add-in loaded. They find each other through the discovery files.
- **Add-in**: `VB6.EXE /make addin\VB6McpBridge\VB6McpBridge.vbp /out build\addin.log` (elevated), then `scripts\register-addin.ps1`. Keep **Binary Compatibility** pointed at `addin/compat/` so the ProgID and CLSIDs stay stable.
- **Test runner**: our own, roughly 100 lines. It finds `[Test]`/`[Category("...")]` methods by reflection, runs them, prints pass/fail and returns a non-zero exit code on failure. Keep it that small.
- **Test categories**:
  - Unit tests (no category) need neither VB6 nor elevation.
  - `Build` tests shell out to `VB6.EXE`.
  - `IDE` tests need a running elevated IDE with the fixture loaded and the add-in connected.
  - `Launch` tests need the elevated watchdog running (`scripts\build-addin.ps1 -WatchdogOnly -Stay`): they start their own IDE through it (no UAC prompt), and quit it.
  - The last two run only when requested.
- **Prefer outcome tests**: build a fixture and check the binary and log; edit through the bridge and check the IDE's text and the bytes on disk. Don't assert on internals.

## Conventions

- C#: nullable enabled, file-scoped namespaces, `async` I/O to the add-in with a timeout on every request. Any in-process COM (typelibs, UIA) runs on a dedicated STA thread; never call COM from thread-pool threads.
- The add-in stays thin: verbs, input and VBIDE access only. Logic lives in the MCP server.
- MCP tools: `area_verb` names, small JSON results, 1-based line numbers. Write tools take `expected_hash` and return the new hash.
- VB6 add-in: `Option Explicit` everywhere. Late binding outside VBIDE. Public surface only on `Bridge.cls`, using strings, Longs and Variant arrays. No UDTs across COM. If you're about to write parsing, diffing or formatting in VB6, put it in C# instead.

## Gotchas

- The program being debugged runs **on the IDE's UI thread**. While it's busy, the bridge can't answer. In break mode the IDE pumps messages; in a busy loop it doesn't.
- **Never `For Each` over VBIDE collections** (spike 3): in run/break mode `For Each` over `VBProjects` and `VBE.Windows` raises error 70 "Permission denied". Indexed loops (`For i = 1 To .Count`) work in every mode, so live reads work while running and debugging. The add-in answers socket requests in all modes, including behind a modal MsgBox.
- **Build-event handlers must not touch VBIDE** (spike 3): reading code inside `EnterRunMode` ended the debugged program as soon as `Form_Load` returned. Log or set a flag only. Entering break mode also fires `EnterDesignMode`.
- **No designer/property access in run mode** (error 393 "Get not supported at runtime", and it can instantiate the form).
- **Designer through VBIDE** (spike 23):
  - `VBComponent.Designer` works with the designer window closed.
  - An object property's `Value` (Font, icons) is a VBIDE **`Properties` collection** of sub-properties, not a StdFont.
  - Call `ContainedVBControls.Add` **early-bound**: late-bound with the typed optional omitted gave error 13.
  - `Add` takes the **`.frm` class name** (`MSComctlLib.ListView`) and returns **Nothing** (no error) for anything else, including the registered ProgID. `AddToolboxProgID` wants the registered ProgID (`MSComctlLib.ListViewCtrl.2`). VBIDE can't list or remove toolbox components.
  - A form isn't a `VBControl`: its properties are `VBComponent.Properties`.
  - Designer access is deferred, because a UserControl's own code runs when its properties are read or set (hard rule 14).
- **Designer windows** (spike 28): an open designer is `MDIClient` > `DesignerWindow "frmMain (Form)"` (no project prefix, unlike code windows) > `ThunderForm`. Take the caption from `DesignerWindow.Caption`; don't build it (it also changes with the window state: maximized drops the project prefix). A form bigger than the designer: post `WM_VSCROLL`/`WM_HSCROLL` to the `DesignerWindow` and stitch captures; `PrintWindow` of the `ThunderForm` alone is misaligned. Opening a designer the first time runs its UserControls' code, which can show a modal box (a production project: "Login Error") that holds the deferred op until answered.
- **`References.AddFromFile` on an unregistered DLL** adds it (empty `FullPath`/`Description`) without registering it. The saved line carries the file's typelib LIBID and a path relative to the `.vbp` (spike 23).
- **COM** (spike 24):
  - OLE caches loaded typelibs, so concurrent reads share one COM object. Serialize typelib reads (`TypeLibInfo.ComLock`) and never `Marshal.ReleaseComObject` them, or another thread's read fails ("separated from its underlying RCW").
  - ActiveX control classes have `TYPEFLAG_FCONTROL`, not `FCANCREATE`, but are registered like creatable ones.
  - Without binary compatibility each build gets new LIBID/CLSIDs, and `/make` unregisters the old ones.
  - A medium process can open an elevated one (query/synchronize), but a quick child may be gone first. The add-in keeps the handle (`REGSTATUS`).
- **PowerShell traps in the scripts** (spike 26):
  - `-File` doesn't parse arrays (`-Roots a,b` is one string): split it yourself.
  - Inside `catch`, `$_` is the error record, not the pipeline item: copy it to a variable first.
  - `claude mcp remove` writes to stderr when there's nothing to remove, and with `$ErrorActionPreference='Stop'` Windows PowerShell 5.1 turns that into a terminating error.
  - A running exe or loaded DLL can't be overwritten but can be renamed: move it aside, copy the new one. So "can I rename it?" does **not** tell whether it's running; an exclusive open (`FileShare.None`) does, where you may write. An elevated process's path is hidden from medium processes, so the watchdog logs its own `exe=` in `watchdog.log`.
  - Check arguments **before** elevating, so a mistake costs no UAC prompt.
  - The installer scripts must stay Windows PowerShell 5.1 compatible.
- **Debugger registrations and broken references** (spike 25):
  - While an ActiveX project runs in the IDE, its classes are registered to `VB98\VB6DEBUG.DLL` and its typelib to the `.vbp`. End restores the previous registration.
  - A killed IDE leaves them until VB6.EXE next starts.
  - A reference to an **unregistered** typelib still works if the `.vbp` hint path exists: the hint is a fallback.
  - A broken reference has no readable `Name`, though Guid and version are readable. `References.Remove` on it fails ("Object library not registered") until its typelib is registered. Registering a file carrying it heals the reference by itself.
- **Designer changes VB doesn't count** (spike 29): setting only an object member (`Font.Bold`) through VBIDE leaves `IsDirty` False, and Save Project skips the form: the change is lost. Setting a scalar (the form's `Tag` to itself) makes VB count it. A form's `Font` doesn't change existing controls (through VBIDE). A CommandButton shows `BackColor` only with `Style = 1` and has no `ForeColor` (438).
- **`IsDirty = False` saves** (spike 31): setting `VBComponent.IsDirty` (or `VBProject.IsDirty`) to False writes the component to disk; it is not a way to drop changes. `VBProjects.Remove` also saves changed components, silently: VBIDE can't close without saving. Discard = snapshot the files, close, restore them. `VBComponent.SaveAs <own path>` saves one component without the `.vbp`. A component that was never saved makes File > Save Project stop at a Save File As dialog.
- **A run marks forms dirty** (spike 32): starting and ending the program in the IDE left the forms it loaded `IsDirty = True` with nothing changed; saving them wrote byte-identical files. A build from disk then refuses ("unsaved changes"). The server reports them as `markedChangedByRun`; it doesn't save them itself.
- **`WithEvents` attributes** (spike 30): VB writes `Attribute x.VB_VarHelpID = -1` itself when it saves a `WithEvents` declaration, so a disk edit may drop it. Other hidden attributes (`VB_UserMemId`, `VB_Description`) can't be recreated and must stay after their line.
- **Designer lists and pictures** (spike 25):
  - `Property.IndexedValue(j)` past the end returns ""/0, no error. The end of a List is its last non-empty item.
  - `Set Property.Object = LoadPicture(file)` sets Picture/Icon; VB writes the `.frx`.
- **Screenshots need elevation** (spike 15): non-elevated `PrintWindow` on VB6 windows fails, and a screen copy only sees what's on top. Elevated `PrintWindow` (watchdog) captures covered windows, dialogs and running forms, including Labels (no hwnd). For IDE panes, capture the main IDE window and crop; `PrintWindow` on the child panes is misaligned. Use a timeout: `PrintWindow` blocks on a hung window.
- `VBE.Windows(i).HWnd` is always 0. Find IDE windows by enumerating child windows of the main window (e.g. Immediate = `VbaWindow` captioned "Immediate").
- Debug command Ids: Start/Continue 186, Break 189, End 228, Restart 515, Step Into/Over/Out 188/194/2559, Run To Cursor 1811, Toggle Breakpoint 51, Clear All Breakpoints 579, Copy 19, Paste 22, Select All 756, Undo 128. Many controls share the generic Id 746: find those by menu path + caption.
- Immediate evaluation: WM_CHAR the text + VK_RETURN to the Immediate `VbaWindow`, then Select All + Copy, then read the clipboard (save and restore the user's clipboard). `Debug.Print` output is read the same way.
- **Closing VB6** (2026-10-17, `ide_quit`): VB6 has two top-level windows, the MDI frame `wndclass_desked_gsk` and `IDEOwner`, which owns it and has the same caption. `WM_CLOSE` to the frame destroys only the frame; the process lives on without UI. Close `IDEOwner`.
- **Mode = IDE window title** (spike 2): class `wndclass_desked_gsk`, caption ends in `[design]`, `[run]` or `[break]`. The non-elevated server can read it, no add-in needed.
- **Starting `VB6.EXE` from the non-elevated side triggers a UAC prompt** (spike 2). Never build from the MCP server; builds go through the elevated add-in.
- Modal dialogs don't stop the add-in from answering. The MCP server's dialog watchdog (`EnumWindows` on the VB6 PID) sees dialogs regardless, because reads cross the elevation boundary.
- VBIDE can't list breakpoints. The MCP server's list is ours and drifts if the user toggles breakpoints by hand. Say so in results; "clear all" resets it.
- Lightweight VB6 controls (Label, Image, Shape, Line) have **no hwnd**. They're visible only in screenshots. VB raises their `Click` **only for real mouse input** (`SendInput`): posted/sent messages, a moved cursor, the foreground and a faked button state all failed (spike 17). Windowed controls work with posted messages. Intrinsic controls are windows (`ThunderTextBox`… in the IDE, `ThunderRT6TextBox`… when compiled).
- **Elevation splits channel C by direction** (spikes 10/11): a medium process can *read* elevated windows (screenshots, `WM_GETTEXT`) but its posted messages are refused and synthetic input is **silently dropped**. This includes Claude's own computer use. All input goes through the elevated add-in, and every input tool must verify its effect, because a dropped input looks like success.
- **Termination crosses elevation** (spike 38): a non-elevated process of the user can **end** the elevated VB6 and the watchdog (`PROCESS_TERMINATE` is granted), though it can't read or write their memory or start threads. An unhooked benchmark agent ran `Stop-Process -Name VB6 -Force` and ended the IDE. The hook asks before shell commands that end VB6 (incl. pipelines, `wmic`, `.Kill()`); the tools end it (`ide_quit`, `ide_kill force=true`).
- `/make` can't see unsaved code. For "does what's in memory compile?", use full-compile start in the IDE and read the compile-error dialog.
- **`VBComponent.Reload()` on a dirty component silently discards the unsaved IDE edits** (spike 8): no prompt, no error. Never reload a dirty component; snapshot the IDE text before any reload. In break mode `Reload()` fails with error 5; only reload in design mode. On a clean component it refreshes code **and** designer.
- Compile errors from memory (spike 6): deferred "Start With Full Compile" (menu path, Id 746 is generic) → the dialog `Compile error:\r\n\r\n<message>` (read by the watchdog) → watchdog `PRESS OK` → `GetSelection` gives line/col. A dialog-open IDE still *responds*, so report dialogs separately from responsiveness.
- `VB6.EXE /make` reads **disk**, never IDE memory, and registers DLLs/OCXs on success. For DLL/OCX it must run elevated: the add-in starts it as a child process.
- **`/make` refuses a project that is open in any IDE** ("The file '...' is already open.", spike 21). Build open projects in their IDE (`MakeCompiledFile`). `/make` error lines are the IDE's code lines **counted from 0**. A binary compatibility break makes `/make` fail with a log line, no dialog; in the IDE it is a decision dialog (never answer it, hard rule 13).
- **`MakeCompiledFile` shows no dialog for a compile error:** it fails with a generic error (E_FAIL, "Path/File access error"). Find the real error with a full compile (spike 6).
- **Opening projects** (spike 22): `VBProjects.AddFromFile(path, False)` makes the IDE a project group even with one project (Save then asks for a group file); `AddFromFile(path, True)` with projects still open writes a `Group1.vbg`. Remove the open projects, then use `True`.
- **Event signatures: dump the coclass's source interface** (spike 22). `FileControlEvents` sources `_dispVBFileControlEvents` (`AfterWriteFile` has 4 parameters); the similarly named `_FileControlEvents` has 2 and VB rejects a handler written to it.
- OCX references live in two places: `Object=` in the `.vbp` **and** `Object = "{GUID}#ver#lcid"; "X.OCX"` in every `.frm`/`.ctl` header. Keep them consistent.
- `.vbp` reference paths are hints: VB resolves `{GUID}#maj.min` through `HKCR\TypeLib` (32-bit view) first, and falls back to the hint path when it isn't registered (spike 25). Resolve the same way.
- **Real-project setup** (spike 27):
  - A registered type library whose file contains **no** type library (a .NET COM DLL registered without `/tlb`) beats the hint path. VB binds to it, finds no types, and the code fails with "Can't find project or library" or "User-defined type not defined". `RegAsm <dll> /tlb` fixes it.
  - VB binds references when it **loads** a project: after a registration change, reload the project.
  - `.vbp` versions are **hex** (`1.a` = 1.10).
  - Licensed controls (Desaware SpyWorks) need the **developer licence** (`HKCR\Licenses`) to load in the IDE or compile; the registered OCX is only the runtime part. VB writes the reason to `<form>.log`.
- `MISSING:` references show as `IsBroken = True`. Command-line build errors for them don't name the reference.
- `.vbp` paths are relative to the `.vbp` folder, may climb out of it (`..\..\Shared\x.bas`), and may use either slash. Typelib refs look like `*\G{GUID}#maj.min#lcid#path#desc`.
- **Mode tag isn't always last in the IDE title** (spike 12): a maximized MDI child appends ` - [Form1 (Code)]`. Match `\[(design|run|break)\]`. The title stays readable while the IDE is busy (cached caption).
- **Don't use VB `Shell` to start the watchdog** (spike 12): it blocked the IDE thread ~9 s. Use `CreateProcess` without waiting.
- **Stale discovery files**: a killed VB6 leaves `instances\<pid>.json` behind. Ignore and delete entries whose PID is dead or isn't VB6.
- A busy loop in the debugged program is rescued by **injected Ctrl+Break** from the elevated watchdog (spike 12); then verify `[break]` in the title.
- **Never `AttachThreadInput` to a VB6 thread** when VB6 is busy (spike 19): the debugged program's form is often the foreground window, and `SetForegroundWindow` while attached to its busy thread blocks the caller for good. An unresponsive window doesn't take the foreground from `SetForegroundWindow`; `SwitchToThisWindow` or an Alt press first does. Ctrl+Break is injected only while a VB6 window (or its `Ghost`) is in front: in a console it would kill the program running there.
- **Drag-and-drop** (spike 20): VB's classic drag needs real input (`SendInput`); posted drags do nothing. OLE drag didn't start from any synthetic input on the test machine (a VM used over RDP: the RDP client keeps resetting the cursor), only by hand.
- **Other add-ins can block VB6 startup** with a dialog before ours loads (spike 19: Visual Component Manager, "Method '~' of object '~' failed"). `ide_launch` reports it; the watchdog clicks it.
- The watchdog is a child of whichever VB6 started it. Don't put it in a job object that kills children; it has to survive VB6 crashing.
- **Restarting the server inside a session** (PLAN.md E7): after `server_restart` ends the process, the desktop app starts it again on the next tool call, from the command it read at session start (a new registration is ignored; the MCP dialog shows no Reconnect for it). That's why `install-server.ps1` installs into the fixed `build\server\current\` and renames files in use aside. The tool **list** isn't fetched again: changed tools work at once, new tools only appear after a session restart.
- **Claude Code loads MCP tools lazily** (`ToolSearch`), so an agent doesn't know the vb6 tools until it looks for them. The server's `instructions` (initialize result) tell it up front.
  - Without them, the benchmark agent tried Edit on VB6 files first (hook-blocked) and once rewrote a `.frm` through Python in Bash.
  - Keep the instructions short: they cost tokens on every request (docs/benchmarks).
- MCP (spike 1): answer unknown methods with -32601. The desktop client first probes `server/discover` and falls back to `initialize` only on that error. Negotiate the protocol version from our supported list. Register servers with `claude mcp add --scope local <name> -e K=V -- <exe>` (name before `-e`, since `-e` is variadic). A project `.mcp.json` didn't load in the desktop app. Local-scope servers are per project folder and don't show in the Connectors UI; check with `claude mcp get <name>`. They load on the next session restart.
- If the MCP server can't find an IDE: check that the add-in is loaded, that a discovery file exists for that PID, and that the token matches. Loopback TCP itself crosses elevation fine (spike 11).

## Workflow for agents

- **Keep everything documented in detail.** Documentation is part of every change, never a follow-up. In the same commit as the code:
  - **`docs/architecture.md`**: components, processes, data flow, and *why* (link the spike that proved it).
  - **`docs/protocol.md`**: the exact wire format between the MCP server, the add-in and the watchdog. Every verb, its arguments, results and errors.
  - **`docs/tools.md`**: every MCP tool. Purpose, parameters, result shape, errors, which channel/process serves it, and an example. Sections are by topic; its Tool index is generated (`node scripts/tool-index.js <vb6mcp.exe> <out.md>`), and a unit test fails when it misses a tool.
  - **`README.md`**: setup, build, registration, running, troubleshooting.
  - **`docs/spikes/NN-*.md`**: every experiment, including failures and surprises.
  - **Code comments explain *why***, especially VB6/Win32 quirks and the spike that found them. Public types and members in C# get XML doc comments.
  - **When behaviour changes, update the docs that describe it**, including PLAN.md and this file. Stale docs count as a bug.
- **Commit often.** Small, working commits are preferred over large ones. Each commit does one thing and leaves the repo in a working state: it builds and the unit tests pass. A spike plus its write-up, or a single tool plus its tests, is a good commit size. Don't bundle unrelated changes.
- Phase 0 work goes in `spikes/`, with a write-up in `docs/spikes/NN-name.md` (what you tried, what happened, the decision) before production code depends on it.
- When a finding contradicts PLAN.md, update PLAN.md in the same change.
- Don't change `addin/compat/`, the add-in ProgID/CLSIDs, the discovery-file format or the wire protocol without calling it out. These break installed machines.
