# MCP tool reference

Tools exposed by `vb6mcp` to Claude Code. Every result is a single JSON text item unless noted. Failures are MCP tool results with `isError: true` and a message meant for the agent (e.g. *"several VB6 IDEs are running (8684, 15652); pass pid"*).

**Implementation:** `src/Vb6Mcp/Tools/IdeTools.cs`. **Wire verbs used:** [protocol.md](protocol.md).

**Server instructions:** the `initialize` answer carries a short `instructions` text (`McpServer.Instructions`), which Claude Code shows the model up front. It says VB6 files must not be changed with Edit/Write or the shell, and which tools to use for what.
- Without it, Claude Code's lazily loaded tools meant the agent tried Edit first and found these tools only after the hook blocked it.
- In a benchmark on a private project (report not published) the instructions took hook blocks from 12 to 0 and shell writes of VB6 files from 1 to 0.
- They also cut the edit tasks' tokens by up to a third.

## Tool index

All 95 tools, generated from the server's tool list (`node scripts/tool-index.js`). Details below, by topic.

| Area | Tool | What it does |
|---|---|---|
| IDEs | `ide_claim` | Make an IDE yours, so other sessions' write tools leave it alone (ide_launch claims the IDEs it starts) |
| IDEs | `ide_dialog_respond` | Click a button in a dialog the IDE is showing (see ide_status) |
| IDEs | `ide_kill` | Last resort: end a VB6 process that is hung or stuck (needs force=true; never done automatically) |
| IDEs | `ide_launch` | Start a new VB6 IDE, optionally opening a .vbp/.vbg (which skips VB6's New Project chooser) |
| IDEs | `ide_list` | List running VB6 IDEs that have the vb6-mcp add-in: pid, mode (design/run/break), responsiveness, open dialogs, projects |
| IDEs | `ide_quit` | Close a VB6 IDE normally (File > Exit), e.g. one this session started with ide_launch and no longer needs |
| IDEs | `ide_release` | Give up ownership of an IDE you claimed or started, so other sessions can change it again |
| IDEs | `ide_status` | State of one IDE without needing its UI thread: title, mode, responds, hung flag, open dialogs (text and buttons), and whether the add-in answers |
| IDEs | `server_restart` | End this vb6 MCP server process (after answering), e.g. after scripts\install-server.ps1 installed a newer build |
| Projects | `project_add_component` | Add a component to a project open in the IDE: a new one (kind: module, class, form, mdiform, usercontrol, propertypage, with a name) or an existing… |
| Projects | `project_checkpoint` | Save a copy of a project's files (the .vbp, every component, .frx and resource files) as they are on disk, to go back to with project_rollback if a… |
| Projects | `project_close` | Close a project in the IDE. Refused while it has unsaved changes, unless discard=true, which throws those changes away (recovery copies the watchdo… |
| Projects | `project_components` | Components of a project in the IDE: name, type, unsaved (dirty) flag, line count, file name |
| Projects | `project_new` | Create a new project in the IDE and save it: kind exe (Standard EXE), axexe (ActiveX EXE), dll (ActiveX DLL) or ocx (ActiveX control); path: the ne… |
| Projects | `project_open` | Open a .vbp or .vbg in a running IDE. If the IDE already has projects, pass mode 'add' (adds to a group) or 'replace' (closes them; refused while a… |
| Projects | `project_rollback` | Put a project's files back as they were at a checkpoint (project_checkpoint) |
| Projects | `project_save` | Save a project like File > Save Project (all unsaved components and the .vbp) |
| Projects | `project_settings` | Read or change project settings |
| Code | `code_insert` | Insert lines into a component before line `line` (use line = countOfLines + 1 to append) |
| Code | `code_read` | Read a component's code as the IDE holds it right now, including unsaved edits |
| Code | `code_replace_lines` | Replace lines of a component's code in the IDE (unsaved, undoable in VB6) |
| Code | `proc_get` | Read one procedure or property (with the comment lines above it) from the IDE, including unsaved edits |
| Code | `proc_replace` | Replace a whole procedure or property (from its first comment line to End Sub/Function/Property) with new_text, exactly the range proc_get returns |
| Code | `symbol_find` | Find where a name is declared in the projects open in running IDEs (procedures, properties, variables, constants, events), from the IDE's own parser |
| Code | `symbol_outline` | Outline of a component from the IDE's own parser (CodeModule.Members): procedures, properties, events, variables, constants with scope and line spans |
| Code | `symbol_refs` | Every statement that uses a name (a procedure, variable, constant...) across the projects under the workspace roots: file, line, the full statement… |
| Run and debug | `bp_clear_all` | Clear all breakpoints (Ctrl+Shift+F9), including ones the user set |
| Run and debug | `bp_toggle` | Toggle a breakpoint on a line (F9) |
| Run and debug | `debug_callstack` | The call stack in break mode, most recent call first (from View > Call Stack, which is opened, read and closed again) |
| Run and debug | `debug_error` | The runtime error dialog ('Run-time error 11: Division by zero'), if one is open: error number and message |
| Run and debug | `debug_eval` | Evaluate an expression in the Immediate window ('? expr') and return the result |
| Run and debug | `debug_immediate` | Read the Immediate window (Debug.Print output and earlier evaluations), optionally only the last N lines |
| Run and debug | `debug_locals` | Variables in scope at the current statement (break mode): the procedure's parameters, return value and Dim/Static variables, plus the module's vari… |
| Run and debug | `debug_output` | OutputDebugString output of the programs being worked on (compiled VB6 programs have no Debug.Print; programs that log usually call OutputDebugString) |
| Run and debug | `debug_where` | Where the IDE is: the active code pane's component, line and procedure (in break mode, the current statement), plus that line's text |
| Run and debug | `run_break` | Break into the running program |
| Run and debug | `run_continue` | Continue the program from break mode (F5) |
| Run and debug | `run_end` | End the running program and return to design mode |
| Run and debug | `run_start` | Start the program in the IDE (F5), or with full_compile=true as 'Start With Full Compile' (Ctrl+F5, compiles everything first) |
| Run and debug | `run_to_cursor` | Run until the given line (Ctrl+F8) |
| Run and debug | `step_into` | Step into the next statement (F8) |
| Run and debug | `step_out` | Step out of the current procedure (Ctrl+Shift+F8) |
| Run and debug | `step_over` | Step over the next statement (Shift+F8) |
| Run and debug | `syntax_check` | Compile the code as it is in the IDE (unsaved edits included) via 'Start With Full Compile' |
| Running program | `app_click` | Click a command button, check box or option button of the running program |
| Running program | `app_keys` | Type text or press keys in the running program |
| Running program | `app_mouse` | Mouse input at a position in a form of the running program: click, double_click, right_click, or drag (from x,y to to_x,to_y) |
| Running program | `app_screenshot` | Screenshot of a window as an image: a form or dialog of the running program (default), or the whole IDE (target 'ide') |
| Running program | `app_scroll` | Scroll a control of the running program (list box, multi-line text box, grid, scroll bar control, or a scrollable form) |
| Running program | `app_select` | Select an item in a list box or combo box of the running program, by text (exact, case-insensitive) or index (0-based) |
| Running program | `app_set_text` | Set the text of a text box in the running program (fires its Change event) |
| Running program | `app_tree` | Controls of a form of the running program: window handle, kind (CommandButton, TextBox, ...), text, position (relative to the form's client area), … |
| Running program | `app_wait` | Wait until something happens in the running program or the IDE, up to timeout seconds: a window with a caption appears (window) or disappears (wind… |
| Running program | `app_windows` | List the windows (forms, dialogs) of the running program (in the IDE, or a compiled EXE with exe), with caption, class, position and size |
| Running program | `exe_end` | Close a program started with exe_run: asks its windows to close (WM_CLOSE), waits 5 s; with force=true ends it if it is still running |
| Running program | `exe_run` | Start a compiled program (.exe under the workspace roots) so the app_* tools can drive it: pass its pid as exe |
| Forms | `form_add_control` | Add a control to a form/UserControl in the designer |
| Forms | `form_controls` | List the controls of a form, UserControl or property page as the designer holds them now (unsaved changes included): name, index (control arrays), … |
| Forms | `form_layout` | Arrange controls of one container (the form, or a Frame/PictureBox) as a grid, instead of working out Left/Top by hand |
| Forms | `form_properties` | All design-time properties of one control (or of the form itself: pass the form's name) with their current values |
| Forms | `form_remove_control` | Remove a control (and any controls inside it) from the designer |
| Forms | `form_resources` | What a form, UserControl or property page keeps in its binary companion file (.frx/.ctx/.pgx): pictures, list items, long or multi-line texts, font… |
| Forms | `form_screenshot` | Screenshot of a form, UserControl or property page as it looks in the designer now (unsaved changes included): layout, fonts, colours, without runn… |
| Forms | `form_set_properties` | Set several design-time properties in one call, on one form/UserControl: changes is a list of {control, property, value}, values as for form_set_pr… |
| Forms | `form_set_property` | Set a design-time property of a control (or of the form: pass its name) |
| References | `refs_add` | Add a reference to a project in the IDE, given exactly one of: name (part of a registered typelib's description, e.g. 'ActiveX Data Objects 2.8'; s… |
| References | `refs_add_component` | Add an ActiveX control's component (OCX) to a project's toolbox, so form_add_control can place it: give the ProgID of one of its controls (e.g. MSC… |
| References | `refs_fix_broken` | Broken (MISSING:) references of a project open in the IDE, each with what the .vbp asked for (LIBID, version, file hint), what is registered for th… |
| References | `refs_list` | A project's references as the IDE has them now (name, LIBID, version, file, broken?), plus its ActiveX control components (Object= lines; read from… |
| References | `refs_remove` | Remove a reference (by its name as refs_list shows it) from a project in the IDE. Or name=#<position> (1-based, in refs_list order) for a broken re… |
| COM | `com_describe` | What a COM component offers: the classes, interfaces, methods, properties, events and enums in its type library, as VB declarations (read from the … |
| COM | `com_diagnose` | Checks behind a COM runtime error (429 can't create object, 430 class doesn't support the expected interface, 48 error loading DLL, or any other) f… |
| COM | `com_orphans` | Broken COM registrations, report only (nothing is changed): classes whose server file is missing, typelibs whose file is missing, and VB6 debugger … |
| COM | `com_register` | Register a DLL/OCX under the workspace roots (regsvr32, audited) |
| COM | `com_search` | Find registered COM components: type libraries whose description or file name, and ProgIDs whose name, has a word starting with the text (32-bit re… |
| COM | `com_unregister` | Unregister a DLL/OCX under the workspace roots (regsvr32 /u, audited; elevated like com_register), then check the registry |
| COM | `com_where` | Where a component is registered: give one of prog_id, clsid, libid or file |
| COM | `compat_status` | Binary compatibility state of ActiveX projects under the roots: mode, compatibility file, the typelib IDs of the compatibility file and the build o… |
| Builds | `build_in_ide` | Compile an open project from the IDE's memory (unsaved edits included) with File > Make, to its output file |
| Builds | `build_make` | Build a project the way the command line does |
| Sync | `sync_diff` | Differences between a component's code in the IDE (including unsaved changes) and its file on disk, as a unified diff (--- ide, +++ disk) |
| Sync | `sync_status` | Files changed outside the IDE: conflicts (a component's file changed on disk while it has unsaved changes in the IDE; nothing was touched) and rece… |
| Sync | `sync_take_disk` | Resolve a conflict in favour of the file on disk: reloads the component, DISCARDING its unsaved IDE changes |
| Sync | `sync_take_ide` | Resolve a conflict in favour of the IDE: saves the project (File > Save Project), writing the IDE's version over the changed file |
| Workspace (disk) | `file_read` | Read a VB6 source file from disk the way the IDE shows it (code only: no designer block, no hidden Attribute lines) |
| Workspace (disk) | `file_replace_lines` | Edit a VB6 source file of a closed project on disk: replace count lines from start_line (1-based, the code as the IDE shows it, as file_read return… |
| Workspace (disk) | `res_read` | A VB6 resource file (.res; or pass the .vbp and its ResFile32 is used), read-only: every resource (type, name or id, language, size), the string ta… |
| Workspace (disk) | `ws_affected` | Which projects a change to (shared) source files affects, in rebuild order: the projects that include the files, then every project that references… |
| Workspace (disk) | `ws_file_projects` | Which projects include a source file (.bas/.cls/.frm/...)? Paths are compared normalized (relative paths, drive mappings, case), so shared files re… |
| Workspace (disk) | `ws_projects` | List the VB6 projects (.vbp) under the workspace roots (VB6MCP_ROOTS): name, type, path, member count, output file, binary compatibility |
| Recovery | `recovery_discard` | Delete the recovery copies of one session (e.g. after restoring them) |
| Recovery | `recovery_list` | Recovery copies of unsaved IDE code, taken by the watchdog while the user is idle: per VB6 session (pid) whether it is still running or crashed/was… |
| Recovery | `recovery_read` | The text of one recovery copy, and whether it differs from the file on disk |
| Recovery | `recovery_restore` | Put a recovery copy back into a running IDE: replaces the code of the component with the same file in an open project (design mode) |
| Other | `batch_run` | Run several vb6 tool calls in one call, in order, e.g. app_click then app_wait then app_screenshot, or run_start, clicks and run_end |

## Common arguments

| Argument | Type | Meaning |
|----------|------|---------|
| `pid` | integer | VB6 process id. **Optional** when exactly one IDE with the add-in is running. |
| `project` | string | `VBProject.Name`. **Optional** when that IDE has exactly one project. |
| `component` | string | `VBComponent.Name` (e.g. `Form1`, `Module1`). |

**Source of truth** (AGENTS.md hard rule 4): code tools read the IDE's memory, never the file on disk. Results carry `source: "ide"` and `dirty` (unsaved changes).

**Modes:** reads work in **design, run and break** mode (spike 3: indexed VBIDE access). While the debugged program runs a loop, the add-in can't answer; tools then fail with *"VB6 (pid …) didn't answer … the IDE is busy or blocked"* after 5 s (hard rule 6). Use `ide_status` to see why.

**Slow calls and a dead server** (spike 42):
- **Heartbeat:** a call that runs longer than 5 s (`ide_launch`, builds, `app_wait`) sends MCP progress notifications every 10 s (`<tool>: still working (N s)`) when the client gave a `progressToken`. Claude Code and Copilot CLI do.
- **Idle timeout:** Claude Code resets its idle timeout on each notification. The installers set that timeout to 2 minutes (`CLAUDE_CODE_MCP_TOOL_IDLE_TIMEOUT=120000`; default 30 minutes). A server that died or froze then shows up as "sent no response or progress for 120s" within 2 minutes, while slow calls keep running.
- **Recovering:** after such an error, reconnect the vb6 server or restart the session.

---

## IDEs and sessions

Finding, starting, owning and recovering VB6 IDEs. `ide_quit` and `ide_kill` are under [Watchdog and recovery](#watchdog-and-recovery).

### `ide_list`

List running IDEs that have the add-in, and clean up after crashed ones.

**Arguments:** none.

**Served by:** discovery files + window reads (non-elevated) + add-in `PROJECTS` (only if the IDE responds).

**Result:**
```json
{
  "ides": [
    { "pid": 8684, "elevated": true, "mode": "design", "responds": true,
      "title": "Project1 - Microsoft Visual Basic [design]",
      "dialogs": [], "addin": "0.1.0",
      "projects": [ { "Name": "Project1", "type": "StandardExe", "file": "D:\\...\\Project1.vbp", "dirty": false } ],
      "addinError": null }
  ],
  "crashed": [ { "pid": 16640, "lastStarted": "2026-10-13T13:26:31" } ],
  "cleanedUpDiscoveryFiles": 1
}
```
- `crashed`: discovery files whose process is gone or isn't VB6 (left by a killed or crashed IDE; a clean close deletes its file). They're deleted after being reported.
- `projects` is null and `addinError` explains why when the IDE doesn't respond.

### `ide_status`

State of one IDE, readable even when its UI thread is busy.

**Arguments:** `pid?`.

**Served by:** window reads (title, `WM_NULL` probe, `IsHungAppWindow`, dialogs) + add-in `PING`.

**Result:**
```json
{ "pid": 8684, "title": "Project1 - Microsoft Visual Basic [break]", "mode": "break",
  "responds": true, "hung": false,
  "dialogs": [ { "caption": "Microsoft Visual Basic", "text": "Compile error:\r\n\r\nVariable not defined", "buttons": ["OK", "Help"] } ],
  "addin": "answers (v0.1.0, protocol 1)" }
```
- `mode`: `design`, `run`, `break` or `unknown`, parsed from the window title (spikes 2, 12).
- `addin`: `answers (…)`, `busy: …` (timeout) or `unreachable: …`.
- A dialog-open IDE still *responds* (modal loops pump messages), so check `dialogs`.

### `ide_launch`

Start a new VB6 IDE.

**Arguments:** `project?` (full path of a `.vbp`/`.vbg`), `elevation?` (`auto` default, `elevated`, `normal`).

**Elevation** (spike 16: whether VB6 runs as administrator is per-machine configuration):

| `elevation` | How | Notes |
|---|---|---|
| `auto` | ShellExecute: Windows applies the machine's configuration | UAC prompt if VB6 is configured to elevate |
| `elevated` | ShellExecute with verb `runas` | always a UAC prompt (until the watchdog can start it without one, Phase 2b) |
| `normal` | CreateProcess with `__COMPAT_LAYER=RunAsInvoker` | fails with an explanation where the machine forces elevation |

**Behaviour:**
- With `project`, VB6 opens it directly (no New Project chooser). Refused if that project is already open in a running IDE (VB6 would show "already open").
- Waits up to 90 s (time to approve UAC) for the new IDE's add-in to answer, then reports mode, projects and dialogs. Dialogs (the New Project chooser when no project was given, source-control prompts, missing references) are **reported, not answered**: use `ide_dialog_respond`.
- Then waits (at most 6 s) until a dialog shows or the IDE has been in design mode with no dialog for 1.5 s, because VB6 shows follow-up dialogs (the New Project chooser) a moment later (improvement notes 31).
- A declined UAC prompt is an error; VB6 isn't started.
- VB6.EXE is found via `VB6MCP_VB6_EXE`, the installer's `ProductDir` in the registry, the `.vbp` file association, or the default folder.

**Result:**
```json
{ "status": "started", "pid": 16724, "elevated": true, "elevation": "auto",
  "exe": "C:\\Program Files (x86)\\Microsoft Visual Studio\\VB98\\VB6.EXE",
  "mode": "design", "title": "Basic - Microsoft Visual Basic [design]", "projects": ["Basic"], "dialogs": [], "note": null }
```
`status: "no-addin"` means VB6 started but the add-in didn't connect (not loaded on startup).

`startSeconds` is how long the shell start took: with a UAC prompt it returns only once the prompt is answered, so a large value means someone had to approve it (null when the watchdog started VB6). `seconds` is the time until the add-in answered. A declined or timed-out prompt fails with how long it waited (improvement notes 16).

### IDE ownership between sessions

Several IDEs always worked (a discovery file per pid, `pid` on every tool), but nothing said whose an IDE was. On 2026-10-17 a test script of one Claude session saved another session's project, because the tools default to the only running IDE.

- **Ownership:** an IDE belongs to the MCP server session that started it (`ide_launch`) or claimed it (`ide_claim`).
  - The owner is recorded in `%LOCALAPPDATA%\vb6mcp\owners\<pid>.json`: the server's pid and start time (so a reused pid doesn't count), the client's name, and when.
  - An owner whose server process has ended owns nothing, and its file is removed.
- **Write tools refuse another session's IDE:** editing, saving, opening and closing projects, run and debug commands, Immediate input, input into the running program, designer changes, references, builds in the IDE, dialog answers, `ide_kill`, `ide_quit`, `form_screenshot` (it opens and scrolls a designer).
  - The error names the owner and points to `ide_launch`.
  - Reading tools work on every IDE.
  - An unowned IDE (one the user opened) is usable by all sessions, as before.
- **Defaulting:** with several IDEs running and no `pid`, a session that owns exactly one of them gets that one.
- **`ide_list`** shows `owner`: `"you"`, `{client, serverPid, since}`, or null.
- **One IDE per session:** with the watchdog started by the user (`watchdog --stay`), `ide_launch` starts elevated IDEs with no UAC prompt.
- **Not a security boundary:** all sessions run as the same user. It keeps the user's own sessions from stepping on each other.

#### `ide_claim`

**Arguments:** `pid`, `force?`. Makes an unowned IDE (or one whose owner has ended) yours. An IDE another live session owns needs `force: true`: ask the user (audited).

#### `ide_release`

**Arguments:** `pid`. Gives up an IDE you own.

### `server_restart` (PLAN.md E7)

**Arguments:** `confirm` (must be true).

Answers, then ends the server process a second later (exit code 3), so a newly installed build can take over without restarting the Claude session. Calls still running then are finished first (at most 2 minutes): an `ide_launch` the client sent right after it was once cut off ("Connection closed") while VB6 had already started. IDEs, the watchdog and unsaved code are not touched. Not allowed in `batch_run`. Served by the server itself (`Program.cs`).

**Wait for its answer before the next vb6 call** (spike 42): once `server_restart` has run, every new call is answered at once with "the vb6 server is restarting … call again in a few seconds", instead of going to the exiting process. In the incident that led to this, an `ide_launch` batched with the restart ran in the dying server; the call after it was never answered, and the session's vb6 connection didn't recover.

**Later in the desktop app (2026-10-19): the server didn't come back.** Twice, after `server_restart`, the app started no new server. It kept showing `vb6` as connected (`session_connectors_status`, 94 tools) while no process ran, and every vb6 call waited until the idle timeout (spike 42). A session restart didn't help, and `/mcp` offered no Reconnect; only restarting the app did. **So prefer restarting the session or the app after installing a new build.** `server_restart` is kept for clients that bring the server back.

What the desktop app did (tested 2026-10-17):
- The MCP servers dialog never showed Reconnect: it kept saying "Connected · 90 tools" after the process ended.
- On the **next vb6 tool call** the app started the server again, by itself, **from the command it read at session start**. A re-registration to a new path during the session is ignored.
- So `scripts\install-server.ps1` now installs into one fixed folder, `build\server\current\`, moving files in use aside (`*.old-<stamp>`, spike 26: renamed, not overwritten). After an install: `server_restart`, then any vb6 call runs the new build. The first session after switching to the fixed folder still needs one restart, to pick up the new path.
- The app keeps the **tool list** it fetched at session start (seen 2026-10-17: after `server_restart`, the changed tools ran the new code, but two tools added in that build weren't offered). New tools need a session restart.
- **IDE ownership survives the restart** (improvement notes 53): ownership is per server process, so before ending, `server_restart` writes the IDEs it owns to `owners\handover-<hash of CLAUDE_CODE_SESSION_ID>.handover` (result: `handedOver`). The next server of the same Claude session claims them at startup, unless another live session claimed one in between; the file is used once and expires after 12 h. Without `CLAUDE_CODE_SESSION_ID` there is no handover and the IDEs need `ide_claim`. The runs this server started (`run_start`) go along too, so `run_end` doesn't ask for `force` afterwards (improvement notes 56); a run counts only if the IDE is still in run or break mode when the next server starts.

### Crash notices (`ideCrashed`, PLAN.md D2)

When an IDE this server saw running has crashed (its discovery file is left behind and the process is gone, spike 12b), the **next result of any IDE tool** gets `ideCrashed`, once per IDE (improvement notes 25: the agent only found out when its next call failed):
- `pid`, `projects` (from the recovery store), `recovery` (the copies the watchdog kept, with their age);
- `applicationLog`: the Windows "Application Error" event of that process (`time`, `faultingModule`, `exceptionCode`, `faultOffset`, `modulePath`), read with `wevtutil` (matched by "Faulting process ID"); null when there is none, e.g. it was killed;
- `next`: relaunch with `ide_launch reopen: true` (the watchdog stays up 30 minutes after a crash, so no UAC prompt while it runs), then `recovery_restore`.

`ide_kill` and `ide_quit` end IDEs on purpose and aren't reported. An IDE that closed normally removes its discovery file and isn't reported either. On a failed result the notice is appended to the error text.

`ide_launch reopen: true` (instead of `project`) opens the first project of the IDE that crashed last; any others are listed in `alsoOpen` (open them with `project_open`).

## Projects

### `project_components`

**Arguments:** `pid?`, `project?`.

**Served by:** add-in `COMPONENTS`.

**`markedChangedByRun`** (spike 32): running the program in the IDE marks the forms it loaded as changed, though nothing changed (saving writes identical bytes). When a run was started through this server (`run_start`, a step or run-to-cursor from design mode, `syntax_check`) and the IDE is back in design mode, `project_components`, `run_end` and `syntax_check` list the components that became dirty only during that run and whose code equals the file: `{components: ["P.frmMain", ...], note}` with the `project_save components=[...]` call that clears them (only those files, not the `.vbp`). Never saved automatically (hard rule 4). `build_make`'s "unsaved changes" refusal names them too. Both T12 rerun agents saved, tested the running program and left exactly such forms behind, so the build from disk refused.
- **`designerOpen`** (improvement notes 46): the listed components whose designer window is open, found by its caption: `frmSlide (Form)` when maximized (spike 28), `Ui - frmSlide (Form)` when not (live re-test 2026-10-19). A form the run didn't load was once listed, and its designer was open: VB probably re-initialises the UserControls of open designers when the mode changes. That cause isn't confirmed live yet, so the result only names them and the note says why they may be there.

**Result:**
```json
{ "project": "Project1",
  "components": [ { "Name": "Form1", "type": "Form", "dirty": true, "lines": 5, "file": "D:\\...\\Form1.frm" } ] }
```
- `type`: `StdModule`, `ClassModule`, `Form`, `MDIForm`, `UserControl`, `PropertyPage`, `ResFile`, … (VBIDE `vbext_ComponentType`).
- `lines` is `-1` for components without a code module; `file` is empty if never saved.

### `project_open`

Open a `.vbp` or `.vbg` in a **running** IDE.

**Arguments:** `path` (full path), `mode?` (`add` or `replace`), `pid?`.

**Rules:**
- If the IDE has no project, `mode` isn't needed.
- If it has projects, `mode` is **required** (hard rule 11, don't hijack the user's session). `add` adds to a group; `replace` closes the open projects and is **refused while any of them has unsaved changes**, because VB6 would ask to save and nobody would answer. The add-in first removes the open projects, then opens the new one with `AddFromFile(path, True)` into the empty IDE.
  - With projects still open, `AddFromFile(path, True)` wrote a `Group1.vbg` of the closed projects into VB6's current folder.
  - `AddFromFile(path, False)` makes even a single project a project group, and then Save asks for a group file name.
  - So only `add` next to other projects uses `False`.
- If the path is already open, the result is `already-open` and nothing happens.
- After the project appears it waits like `ide_launch` (dialog, or 1.5 s quiet, at most 6 s). A dialog shown by then (SourceSafe's "Add this project?") comes back as `dialogs` with a note to answer it with `ide_dialog_respond` before the next call (improvement notes 31).
- Starting VB6 itself isn't possible from here: a non-elevated process starting VB6 triggers a UAC prompt (spike 2). That's the watchdog's `ide_launch` (Phase 2b).

**Served by:** add-in `OPEN` (deferred), then polling `PROJECTS` and window dialogs for up to 20 s.

**Result** (`status`):
```json
{ "status": "opened", "path": "D:\\...\\Basic.vbp", "project": "Basic", "type": "StandardExe", "projects": ["Basic"] }
{ "status": "dialog", "path": "...", "note": "VB6 is showing a dialog while opening; it was not answered",
  "dialogs": [ { "caption": "Microsoft Visual Basic", "text": "...", "buttons": ["OK"] } ] }
{ "status": "already-open", "path": "...", "projects": ["Basic"] }
{ "status": "pending", "path": "...", "note": "not visible after 20 s; check ide_status" }
```


**Preflight and load check** (improvement notes 8 and 17: a production project loaded 0 of 172 components without a word, because two OCXs weren't registered):
- **Before opening a `.vbp`**, every `Object=` control and `Reference=` is resolved like VB does (typelib GUID + major version, 32-bit registry).
  - An unregistered control **refuses the open**. VB6 would load the project without its components: through `AddFromFile` its "Continue Loading Project?" prompt resolves as No. The error names each control and any other registered versions. `force: true` opens anyway.
  - Unregistered references only produce `warnings`: VB6 loads the project and shows them as MISSING.
  - **Hints** (PLAN.md C3, improvement notes 19), in brackets after each missing item:
    - files **on disk** that carry the wanted library but aren't registered: searched by the file name the `.vbp` gives (the `Object=` file, or the reference's hint path) in the project folder (recursively), `SysWOW64`, the hint folder and the roots (recursively), confirmed by reading their typelib (LIBID); capped at 8 s. `com_register` them.
    - **registered libraries with a similar name** (description or file name, ignoring versions and words like Lib/Control): another version or edition, a *different* library (`CapturePRO 3.0` for a project that wants `CAPTUREPROLib` 2.0). They say what's installed; they don't replace the wanted one (hard rule 13).
- **After opening**, the loaded components are compared with the `.vbp`'s members. Fewer gives `status: "opened-incomplete"` with `incomplete: {loaded, expected, notLoaded, loadLogs}`. `loadLogs` holds the `.log` files VB wrote next to the forms during the load. Don't save such a project: the `.vbp` would lose the missing members.
- **On a failed open**, the add-in's error is returned (`status: "failed"`, e.g. `60082: A project with the name 'AppA' is already loaded`) instead of a bare `pending`.

### `project_save`

**Arguments:** `project?`, `components?`, `path?`, `outside_edits?`, `pid?`.

**Served by:** `SAVE` (deferred File > Save Project), then polling `COMPONENTS` and dialogs for up to 20 s.

**Result** (`status`): `saved` (with the components that were unsaved), `nothing-to-save`, `dialog` (reported, not answered, e.g. a source-control checkout prompt) or `pending`.

**A never-saved project** (File > New Project, the New Project chooser at startup) needs `path`: the new `.vbp`, under the workspace roots, not existing yet (a missing folder is created). Without `path` it is refused: Save would open a Save As dialog, and an unchanged new project isn't even dirty.
- Served by `PRJSAVEAS`: each never-saved component is saved next to the `.vbp` as `<name>.<ext>`, then `VBProject.SaveAs` writes the `.vbp` (spike 43). A file already there refuses the call before anything is written.
- The result is `project_new`'s (below), with `status: "saved"`. The project keeps its name (rename it with `project_settings`).
- `path` for a project that already has a `.vbp` is refused (no Save As to another file).


**`components: [...]`** saves only those components (`SAVECOMP`), and the `.vbp` isn't written. Components that were never saved are refused, since they have no file yet.

**The `.vbp`:** saving a project VB counts as changed also rewrites the `.vbp`. In a production project (2026-10-17) that rewrote 17 `Reference=`/`Object=` lines: hint paths made absolute, OCX names re-cased, and two control versions bumped to the ones registered (mscomctl 2.1 to 2.2, comctl32 1.4 to 1.5).
- The result's `vbp` lists the lines removed and added, and `versionChanges`: `{kind, guid, name, oldVersion, newVersion}` for references and controls.
- A version change also gets a `note`, because committing it changes what the project needs on other machines.
- `vbp` is absent when the `.vbp` didn't change.

**Edits no tool made** (improvement notes 50): before saving (the whole project or `components`), each dirty component's IDE text is compared with what the tools last left in it (`code_replace_lines`, `code_insert`, `proc_replace`, `file_replace_lines` routed to the IDE, `recovery_restore`), or with the file on disk when no tool edited it since the server started.
- A difference other than the case of identifiers is **refused**, listing each component's diff (at most 40 lines, `sync_diff` for the rest). Stray keystrokes that reach a code pane, or the user's own typing, are caught this way.
- `outside_edits: true` saves anyway; the result then has `outsideEdits`: `{component, since, changedLines, diff, moreDiffLines}` (`since` is `last vb6-mcp edit` or `disk`).
- Reads don't update the record, so a `code_read` after a stray keystroke doesn't hide it. Never-saved components and unreadable ones are skipped.

### `project_new`

**Arguments:** `kind` (`exe`, `axexe`, `dll`, `ocx`), `path` (the new `.vbp`), `name?` (default: the file name), `mode?` (`add` or `replace`; required when the IDE has a project), `pid?`. Design mode. Served by `PRJNEW`.

Creates a project in the IDE and saves it, so an agent no longer writes a `.vbp` by script (notes from real use).
- **Checks first, nothing written on refusal:** the path must be a `.vbp` under the workspace roots that doesn't exist; the name a valid VB name not already open (unless `replace`); the folder must not hold the default component's file. A missing folder is created (and removed again if the add-in refuses).
- **What VB creates** (`VBProjects.Add`, spike 43): `exe` a `Form1`, `axexe` and `dll` a `Class1`, `ocx` a `UserControl1`. It's saved next to the `.vbp` as `Form1.frm`, `Class1.cls` or `UserControl1.ctl`, then the `.vbp` with `VBProject.SaveAs`.
- **`mode: add`** puts the new project next to the open ones: the IDE becomes a project group, and VB asks about saving the group (`.vbg`) when the IDE closes. **`mode: replace`** closes the open projects first; refused while any has unsaved changes (closing would save them, spike 31).
- **Result:** `status` (`created`; `incomplete` when the `.vbp` doesn't list a component or its file is missing, with `notInVbp`), `project`, `path`, `type`, `dirty`, `components` (`name`, `type`, `file`), `projects` when the IDE holds a group.
- **SourceSafe:** on a machine with Visual SourceSafe, VB's source-control add-in asks "Add this project to SourceSafe?" after the files are written. The result is then `status: "dialog"` with a note; answer it with `ide_dialog_respond` (No keeps it out; it still leaves a read-only `MSSCCPRJ.SCC` in the folder) and check with `project_components`.

```json
{"kind": "dll", "path": "D:\\src\\Lib\\MyLib.vbp", "mode": "add"}
-> {"status": "created", "project": "MyLib", "path": "D:\\src\\Lib\\MyLib.vbp", "type": "ActiveXDll", "dirty": false,
    "components": [{"name": "Class1", "type": "ClassModule", "file": "D:\\src\\Lib\\Class1.cls"}], "projects": ["App", "MyLib"], ...}
```

### Closing, adding, settings, checkpoints

#### `project_close`

**Arguments:** `project?`, `discard?`, `pid?`. Design mode. Served by `CLOSE`.

Closes a project in the IDE.
- A project with unsaved changes is refused, with the changed components named, unless `discard: true`. That throws the changes away (audited); the watchdog's recovery copies stay.
  - VBIDE can't close without saving: `VBProjects.Remove` saves changed components silently, and so does setting `IsDirty = False` (spike 31).
  - So a discard snapshots the `.vbp` and each changed component's files first, lets VB close, writes the snapshots back, and deletes source files VB created for never-saved components.
  - The result's `restored` lists `files`, `deleted` and any `failed`.
  - Refused if another IDE has one of those files loaded (hard rule 4) or the project was never saved.
- Result: `status: "closed"`, `discarded` (the components whose changes were thrown away), `projects` (still open). A dialog is reported, not answered.

#### `project_add_component`

**Arguments:** `kind` + `name` (new: `module`, `class`, `form`, `mdiform`, `usercontrol`, `propertypage`), or `file` (an existing source file under the roots), `project?`, `pid?`. Design mode. Served by `COMPADD`.

Adds the component in the IDE. This replaces writing files and `.vbp` lines by script (improvement notes 1).
- A **new** component is saved straight away to `<project folder>\<name>.<ext>` (`SAVECOMP` with a path). The file must not exist yet: nothing is overwritten. Otherwise File > Save Project would stop at a Save File As dialog for it (spike 31).
- `project_save` then adds it to the `.vbp`.
- Result: `status: "added"`, `component`, `type`, `file`, `dirty`.
- A name collision or a bad kind is refused before anything happens.

#### `project_settings`

**Arguments:** `set?` (object), `reopen?`, `path?` (a closed project's `.vbp`), `project?`, `pid?`.

Reads and changes project settings (improvement notes 2, 7, 22).
- **Open project, IDE settings:** `name`, `description`, `helpfile`, `helpcontextid`, `startmode` (0 standalone, 1 ActiveX component), `buildfilename` and `startup` (a form, `Sub Main` or `(None)`) change live through `PRJSET`, unsaved until `project_save`.
- **Open project, `.vbp`-only settings:** `CondComp` (conditional compilation constants), `Command32` (command-line arguments) and other project-section keys exist only in the `.vbp`. VB rewrites the `.vbp` from memory when it saves. They need `reopen: true`: the project must have no unsaved changes, it is closed, the `.vbp` is edited byte-exactly, and it is opened again. Without `reopen` the call is refused before anything changes.
- **Closed project (`path`):** keys are `.vbp` names (`CondComp`, `Command32`, `Startup`, `Title`, ...), edited byte-exactly (`VbpKeys`). Refused while any IDE has the project open (hard rule 4).
- **Never changed here:** members, `Reference=`, `Object=`, `ResFile32`, `Type` and binary compatibility (`CompatibleMode`, `CompatibleEXE32`, `VersionCompatible32`). Those have their own tools, or carry GUIDs (hard rules 2 and 13).
- **Result:** `settings` (from `PRJGET`), `vbp` (the `.vbp`'s project-section keys), `reopened` when the project was closed and opened.

#### `project_checkpoint`

**Arguments:** `label?`, `list?`, `path?` (a `.vbp`, e.g. of a closed project), `project?`, `pid?`.

Copies a project's files as they are on disk: the `.vbp`, every member (`Form=`, `Module=`, `Class=`, `UserControl=`, `PropertyPage=`, `Designer=`, `UserDocument=`, `RelatedDoc=`), their `.frx`/`.ctx`/`.pgx`/`.dsx`/`.dox` and the `ResFile32` file, byte-exact, to `%LOCALAPPDATA%\vb6mcp\checkpoints\<project>-<path hash>\<id>\` with a `manifest.json` (path, copy, SHA-256 per file).
- **Why:** the T9–T12 redesigns changed hundreds of lines across many forms with no way back but the owner restoring files (PLAN.md E2).
- **Open project:** refused while it has unsaved changes. Unsaved IDE changes can't be copied without saving them (VBIDE only saves, spike 31), so save first.
- **`list: true`:** the checkpoints, newest first: `id`, `label`, `created`, `files`, `changedSince` (files that differ now).
- **Kept:** the last 20 per project. Nothing under the project folder is written.
- **Result:** `id`, `label`, `files` (count), `missing` (members named in the `.vbp` that weren't on disk), and how to roll back.

#### `project_rollback`

**Arguments:** `checkpoint` (id, label or `latest`), `confirm` (must be true), `discard_unsaved?`, `path?`, `project?`, `pid?`.

Writes a checkpoint's files back where they differ (temp file, then move).
- **Closed project:** written directly. Refused if any IDE has one of the files loaded (hard rule 4), including a shared module loaded by another project.
- **Open project:** design mode, and not in an IDE another session owns. Unsaved changes need `discard_unsaved: true` (hard rule 9). The project is closed (the files VB writes on close are snapshotted and put back, and source files it creates for never-saved components are deleted, as in `project_close discard`), the checkpoint is restored, and the project is opened again in the same IDE.
- **Paths** must all be under the workspace roots (hard rule 8). Audited.
- **Files added since** the checkpoint (new members, or files that didn't exist then) are left on disk and listed in `addedSince`; the restored `.vbp` doesn't name them.
- **Result:** `status` (`restored` or `partly restored`), `restored` (paths written), `unchanged` (count), `failed`, `addedSince`, `reopened` (`{pid, project}`, or `{pid, error}`).

Verified live on a copy of the Basic fixture: a saved edit plus the `.vbp` rewrite VB made on save were rolled back and the project reopened with the original hash; with an unsaved edit the call was refused until `discard_unsaved: true`. The files ended byte-identical to the originals.

## Reading code

### `code_read`

Read code as the IDE holds it right now, including unsaved edits.

**Arguments:** `component`, `project?`, `pid?`, `start_line?`, `count?` (1-based; both or neither).

**Served by:** add-in `READ`.

**Result:**
```json
{ "source": "ide", "project": "Project1", "component": "Form1", "dirty": true,
  "countOfLines": 5, "hash": "3f2a9c0d1e4b5a67", "startLine": 1,
  "text": "Option Explicit\r\n\r\nPrivate Sub Form_Load()\r\n    MsgBox \"hi\"\r\nEnd Sub" }
```
- Lines are CR LF separated. The hidden `Attribute` lines and the `.frm` designer block aren't part of the code module, so they're not included.


**Large modules:** without `start_line`/`count`, at most 2,000 lines and 100,000 characters come back (a 185 KB class overflowed the client, improvement notes item 13).
- A cut result has `truncated: true`, `nextStartLine` and a `note`. Read on with `start_line`.
- `hash` always covers the whole text, so edits need no full read.
- An explicit range is returned as asked, up to the character cap.

### `symbol_outline`

Outline from the IDE's own parser (`CodeModule.Members`): a cheap way to navigate large modules.

**Arguments:** `component`, `project?`, `pid?`.

**Served by:** add-in `MEMBERS`.

**Result:**
```json
{ "source": "ide", "project": "Project1", "component": "Form1",
  "members": [
    { "Name": "Form_Load", "kind": "method", "procKind": "proc", "scope": "private", "static": false,
      "line": 3, "startLine": 2, "lines": 3 },
    { "Name": "m_count", "kind": "variable", "procKind": null, "scope": "private", "static": false,
      "line": 1, "startLine": null, "lines": null }
  ] }
```
- `kind`: `method`, `property`, `variable`, `event`, `const`. `procKind`: `proc`, `get`, `let`, `set` (methods/properties only).
- `line` = the `Sub`/`Function`/`Property` line (or the declaration line); `startLine`/`lines` follow VB6's definition and include the blank and comment lines above the procedure.
- A property appears once **per kind** (`get`, `let`, `set`), even though VBIDE's `Members` lists it only once.

### `symbol_find`

Find where a name is declared in the projects open in running IDEs.

**Arguments:** `name`, `contains?` (substring match instead of exact), `project?`, `pid?` (narrow the search).

**Served by:** add-in `PROJECTS` + `COMPONENTS` + `MEMBERS` for every component with a code module, i.e. the IDE's own parser (principle 3: no parser of ours). Matching is case-insensitive.

**Result:**
```json
{ "name": "AddNumbers", "contains": false, "truncated": false,
  "matches": [ { "pid": 10856, "project": "Basic", "component": "modMath", "componentType": "StdModule",
                 "Name": "AddNumbers", "kind": "method", "procKind": "proc", "scope": "public", "line": 7 } ],
  "skipped": [] }
```
- Stops after 200 matches (`truncated`). IDEs that don't answer (busy, unreachable) are listed in `skipped`.
- Finds declarations only, not usages; text search over the files covers usages (principle 3).
- Projects that aren't open in an IDE aren't searched; use `ws_projects` / `file_read` for those.

### `symbol_refs`

Every statement that uses a name, across the projects under the workspace roots (PLAN.md C5, improvement notes 4: the agent needed every caller of `Add_To_Audit` in six projects and which optional arguments they pass; grep missed the calls continued over several lines).

**Arguments:** `name` (one identifier), `project?` (projects whose name contains it, or one `.vbp` path), `declarations?` (default true), `refresh?`.

- **Served by** the server (`VbRefs`, `WorkspaceTools`): the code part of every member of the selected projects (`VbSourceFile`, IDE line numbering), each shared file once. A component that is **dirty in a running IDE** is read from the IDE instead (hard rule 4); `source` says which.
- **Matching:** whole word, case-insensitive (VB names aren't case-sensitive), also when qualified (`modAudit.Add_To_Audit`). Statements are joined across ` _` continuations first. String literals and comments (`'`, `Rem`) don't count. No parser beyond that (principle 3): a local variable with the same name in another module is listed too.
- **Declarations** (`Sub`/`Function`/`Property`/`Declare`/`Event` of that name) are marked `declaration: true`; `declarations: false` leaves them out.
- **Limit:** 300 statements (`truncated`, narrow with `project`).
- **Result:** `count`, `filesScanned`, `files: [{path, source, hash, projects, refs: [{line, endLine?, declaration?, statement}]}]`, `warnings` (an IDE that didn't answer: its unsaved changes aren't included). `hash` covers the text that was scanned (the IDE's for a file open there), so it is the `expected_hash` for `file_replace_lines`: edit the found lines without reading each file first.

```json
{ "name": "Add_To_Audit", "projects": 61, "filesScanned": 1830, "count": 2,
  "files": [ { "path": "D:\\src\\Shared\\modAudit.bas", "source": "disk", "projects": ["AppA", "AppB"],
               "refs": [ { "line": 4, "endLine": 6, "statement": "Add_To_Audit \"Pay\", \"amount \" & CStr(x)" } ] } ] }
```

### `proc_get`

Read one procedure or property, with the comment lines above it.

**Arguments:** `component`, `name`, `kind?` (`proc` default, `get`, `let`, `set`), `project?`, `pid?`.

**Served by:** add-in `PROC`.

**Result:**
```json
{ "source": "ide", "project": "Project1", "component": "Form1", "name": "Form_Load", "kind": "proc",
  "startLine": 2, "bodyLine": 3, "lines": 3, "text": "Private Sub Form_Load()\r\n    MsgBox \"hi\"\r\nEnd Sub" }
```

---

## Dialogs

### `ide_dialog_respond`

**Arguments:** `button` (caption, e.g. `OK`, `No`), `dialog?` (caption; needed when several dialogs are open), `repeat?`, `pid?`.

**Served by:** window reads (finds the dialog and the button's window handle) + add-in `CLICK`, then a re-read to verify. Before the add-in has loaded (another add-in's startup error), or when it doesn't answer, the watchdog clicks. When VB6 isn't elevated and no watchdog is running yet, the server clicks the button itself, since it runs at the same level (`via: server (…; VB6 isn't elevated)`, spike 41).

**`repeat: true`** (the owner: a missing reference or control makes VB "ask the same message many many times" in a project with many forms; "click OK until it stops", spike 36): after the first click, keeps pressing the same button while the next dialog is the same kind of message (its text with quoted paths blanked out), waiting up to 5 s for each next form; at most 500 clicks. Then it reads the `.log` files the messages named and sums them up. Result: `repeated` (clicks), `stoppedBecause` (`no more dialogs`, `a different dialog`, the limit), `distinctMessages`, `missing` (per class or licence: `class`, `reason`, `controls`, `forms`, `formNames`), `dialogsNow`. Verified on a 12-form fixture: 12 clicks, stopped at the SourceSafe question, `MSComctlLib.ProgressBar` on 12 forms. Saving such a project would drop the controls that didn't load; fix the component and reopen instead.

**Result:**
```json
{ "clicked": "No", "dialog": "Source Code Control", "text": "Add this project to SourceSafe?",
  "dialogClosed": true, "dialogsNow": [] }
```
The button must always be named explicitly (hard rule 9). An unknown button lists the available ones.

**A click that doesn't take** (improvement notes 11: the first click on "Continue Loading Project?" left it open, the same call a minute later worked). When the dialog is still open half a second after the click, the tool waits up to 2 s more, then presses the same button of the same window once more and waits again. The result then has `retried: true`. If the dialog is still there after that, `dialogClosed: false` comes with a `note`: it may be waiting behind another IDE window (a dialog the user opened), or the IDE is busy.

---

## Editing

Rules for all editing tools:
- **Optimistic concurrency:** `code_read` and `proc_get` return `hash`, a hash of the **whole module** (first 16 hex characters of SHA-256 over its 1252 bytes). Edits require it as `expected_hash`. The server re-reads the module and refuses on mismatch (*"conflict: … changed since you read it"*); the add-in additionally checks that the target lines are still exactly what was read (protocol `REPLACE`, 409). So a user typing in the IDE is never overwritten: re-read and redo.
- **Design mode only**, and no dialog open. Edits in break mode contributed to the spike 3 deadlock.
- **Validated before sending:** line endings normalized to CR LF, one trailing line break dropped, at most 1023 characters per line, at most 24 line continuations, only characters that exist in Windows-1252 (hard rules 1, 10).
- Edits are **unsaved** in the IDE and **undoable** there (Edit > Undo). Nothing is saved automatically (hard rule 4).
- Every edit returns the new `hash`, so the next edit can follow without a re-read.

### `code_replace_lines`

**Arguments:** `component`, `start_line`, `count` (0 = insert), `new_text` (empty = delete), or `edits`, plus `expected_hash`, `project?`, `pid?`.

**Several ranges in one call:** `edits: [{start_line, count, new_text}, …]`, at most 500 (improvement notes item 5: a 50-line rename took 50 calls, each needing the previous hash).
- All ranges are numbered against the text `expected_hash` covers, and must not overlap or share a start line. They are checked before anything is sent.
- They are applied bottom-up, so line numbers above each edit stay valid.
- Each is a separate VBIDE edit with its old text as guard. If one fails, the error says how many later edits were already applied (undo in VB6).
- **After a failure** other than the guard (409), the server re-reads the module and the error says which state it is in: nothing changed, the later edits applied but not this one, this one applied too (don't redo it), or neither old nor new (half-applied: undo in VB6 or repair; don't retry blindly). This holds for a single range too (improvement notes item 30).
- **Once applied, an edit is never reported as an error:** the follow-up checks (`vbKeptSpelling`, `codeHints`) are dropped if they fail, and if the final re-read fails the result is still a success with a note to get the new hash with `code_read` (item 29: an empty literal `""` used to throw here).
- A fresh module reads as `Option Explicit\r\n` with `countOfLines: 2`: VB counts the empty line after the last line break, and it can be replaced like any other (`start_line: 1, count: 2`).
- The result lists them as `edits` (in file order) instead of `replaced`.

**Find/replace** (PLAN.md A6), instead of ranges: `find`, `replace`, `whole_word?` (default true), `match_case?` (default false, as VB names ignore case), `in_strings?`, `in_comments?` (default false: only code changes, so a rename doesn't touch messages or comments), `apply?`.
- Without `apply: true` it is a **preview**: `preview: {lines, changes: [{line, old, new}] (at most 200), note}` and the unchanged `hash`. Nothing changes.
- With `apply: true` the changed lines become an `edits` list against `expected_hash` (at most 500 lines), applied as above. No matches: an error, nothing changed.
- Within a line only (find and replace can't contain line breaks). Also on `file_replace_lines` (disk, or routed to the IDE that has the file loaded).

**Line breaks and escapes** (owner, 2026-10-17: CR LF and backslashes had to be fixed by hand): any mix of CR, LF and CR LF in the submitted text becomes CR LF. The code parameters' descriptions and the server instructions say VB6 is Windows-only and that VB strings have no escapes. When a string literal in the submitted code contains `\n`, `\r`, `\t` or a doubled backslash (not a leading `\\` of a UNC path), the result has `codeHints` explaining it (`vbCrLf` / single backslashes). Applies to `code_replace_lines`, `code_insert`, `proc_replace` and `file_replace_lines`.

**Served by:** `READ` (hash check) + one `REPLACE` per range + `READ`.

**Re-cased names:** VB keeps one spelling per name for the whole project and re-cases new text to it. Retyping `Err.number` as `Err.Number` can therefore leave the module, and its hash, unchanged (improvement notes item 24). The result's `vbKeptSpelling` lists each typed name VB spells differently (`"Number -> number"`), and the note says when the module didn't change at all. Strings and comments are ignored. Applies to `code_replace_lines`, `code_insert` and `proc_replace`.

**Result:**
```json
{ "source": "ide", "project": "Basic", "component": "modMath", "dirty": true, "countOfLines": 14,
  "hash": "9c1e0f6a2b7d4e11", "replaced": { "startLine": 7, "oldLines": 1, "newLines": 2 },
  "note": "unsaved in the IDE; undo with Edit > Undo in VB6, save with project_save" }
```

### `code_insert`

**Arguments:** `component`, `line` (insert before; `countOfLines + 1` appends), `text`, `expected_hash`, `project?`, `pid?`. Same result as `code_replace_lines`.

### `proc_replace`

**Arguments:** `component`, `name`, `new_text`, `expected_hash`, `kind?` (`proc`, `get`, `let`, `set`), `project?`, `pid?`.

Replaces exactly the range `proc_get` returns: from the first comment line above the declaration to `End Sub`/`Function`/`Property` (VB6's own procedure start includes the blank lines above it; those are left alone). `new_text` is the complete new procedure including its comments.

## Run and debug

VB6 has no debugger API. These tools run IDE commands (deferred in the add-in, protocol `CMD`) and then **observe the outcome**: the mode from the window title, dialogs, the current line. A sent command isn't proof that it worked (spike 10). Each result carries `status`, `mode`, `responds`, `title`, `dialogs` and a `note` when something needs attention (e.g. a dialog to answer with `ide_dialog_respond`, or an IDE that stopped responding because the program is in a loop).

| Tool | Arguments | Needs mode | What it does |
|------|-----------|------------|--------------|
| `run_start` | `full_compile?`, `pid?` | design | F5, or Start With Full Compile (Ctrl+F5). Warns about DLLs the project's controls and references need that aren't installed (`missingDlls`, below) |
| `run_break` | `pid?` | run | Break. If the program is in a busy loop the IDE can't take the command (the add-in shares its thread); the watchdog will inject Ctrl+Break (Phase 2b) |
| `run_continue` | `pid?` | break | Continue |
| `run_end` | `force?`, `pid?` | run, break | End. **Needs `force: true` if the agent didn't start the run** (hard rule 11) |
| `step_into` / `step_over` / `step_out` | `pid?` | break (`step_into` also design) | F8 / Shift+F8 / Ctrl+Shift+F8. In break mode the result has `at`: `component`, `line`, `procedure`, `lineText` where it stopped |
| `run_to_cursor` | `component`, `line`, `project?`, `pid?` | break, design | sets the cursor, then Ctrl+F8 |
| `bp_toggle` | `component`, `line`, `project?`, `pid?` | design, break, run (spike 37) | sets the cursor, then F9. Returns the breakpoints this server knows; VBIDE can't list breakpoints, so ones set by hand are unknown |
| `bp_clear_all` | `pid?` | design, break | Ctrl+Shift+F9 (all breakpoints, including the user's). Not while the program runs: VB ignores it then (spike 37) |
| `debug_where` | `pid?` | any | active code pane: component, line, column, procedure, line text (in break mode: the current statement) |
| `debug_eval` | `expression`, `pid?` | break, design | types `? expression` into the Immediate window and returns the new output. An expression that doesn't compile is a tool error with VB's message; the compile-error dialog it caused is closed (a run-time error dialog is left open and reported) |
| `debug_immediate` | `last?`, `pid?` | any | the Immediate window's text (`Debug.Print` output) |
| `syntax_check` | `pid?` | design | compiles the code **as it is in the IDE** (unsaved edits included) |

**Verified live** (Basic fixture, 2026-10-13): `syntax_check` (compiles, start + end), `bp_toggle` in `AddNumbers`, `run_start`, the user clicked Add, the IDE stopped at the breakpoint, `debug_where` (line 8, the current statement), `debug_eval` `a` = 1, `b` = 0, `g_callCount` = 0, `step_over` (line 9), `bp_clear_all`, `run_continue`, `run_end`; in design mode `debug_eval` of `PI_APPROX * 2` (6.28318) and `AddNumbers(40, 2)` (42), with the Immediate window's existing content and the user's clipboard text preserved.

**`debug_eval`** result: `{ "expression": "x", "status": "ok", "result": "42" }`. In design mode only module-level and global names are in scope, and calling a function runs it. A dialog that appears while evaluating (e.g. a runtime error) is reported (`status: "dialog"`), not answered.

**Side effects are refused** (PLAN.md E5). The Immediate window runs with the IDE's rights (usually an administrator's), so an expression is code execution. Both T12 batch runs changed rows of the owner's test database from it (`SqlCon.Execute("update orders set ...")`), and the owner asked for every kind of side effect to be covered. `EvalPolicy` refuses, before anything is typed:
- **database writes:** SQL that changes data, schema or rights (`INSERT INTO`, `UPDATE x SET`, `DELETE FROM`, `DROP`/`ALTER`/`CREATE`/`TRUNCATE`, `MERGE`, `EXEC`, `GRANT`/`REVOKE`/`DENY`, `xp_cmdshell`, `BULK INSERT`, `BACKUP`/`RESTORE`, `SHUTDOWN`) and recordset/connection calls (`.AddNew`, `.Update`, `.Delete`, `.Edit`, `.BeginTrans`, `.CommitTrans`);
- **process starts:** `Shell`, `ShellExecute`, `CreateProcess`, `WinExec`, `WScript.Shell` `.Run`/`.Exec`, `Shell.Application`, WMI (`winmgmts`, `Win32_Process`);
- **file changes:** `Kill`, `FileCopy`, `Name … As`, `MkDir`, `RmDir`, `SetAttr`, `Open … For Output/Append/Binary/Random`, `SavePicture`, `FileSystemObject` and `ADODB.Stream` writers, the Win32 file APIs;
- **registry:** `SaveSetting`, `DeleteSetting`, `.RegWrite`, `.RegDelete`, `Reg*` APIs, `StdRegProv`;
- **network:** XMLHTTP, WinHttp, Inet, Winsock, `URLDownloadToFile`, WinInet;
- **scripts and code:** `ScriptControl`, `.AddCode`/`.ExecuteStatement`, `CallByName`, the VBIDE object model;
- **keystrokes and system:** `SendKeys`, `AppActivate`, `keybd_event`, `SendInput`, shutdown and privilege APIs.

The refusal is audited and tells the agent to test through the program's UI. **Asking the user** (spikes 33, 40): the agent can call again with `allow_side_effect: true`, and the user is asked by the client. Who asks depends on the client that connected (`clientInfo.name` and `capabilities` from `initialize`):
- **Claude Code** (`claude-code`) with the vb6-mcp hook installed (its matcher includes `mcp__vb6__debug_eval`): the hook shows a permission prompt with the expression and what it does, and only an approved call reaches the server, which runs it. Audit: `approved by the user (Claude Code hook prompt)`.
- **Any other client that offers form elicitation** (Copilot CLI, OpenCode, Codex; also Claude Code without the hook, if it ever offers it): the server sends `elicitation/create` with the expression, the matched category, and a form with one required boolean `approve`. It runs only on `accept` with `approve: true`; decline, cancel, an unticked box, an error, or no answer within 5 minutes refuse it ("refused: the side effect was not approved ..."). The agent never sees the form. Audit: `approved by the user (MCP elicitation in <client>)`, or `refused, not approved (MCP elicitation in <client>: <action>)`.
- **Otherwise refused**, since an agent could pass the flag itself. Before spike 40 the hook check was about the machine, not the client: Copilot CLI on a machine with the hook installed ran a `CreateTextFile` with no prompt.

**What the approval protects against, and what it doesn't:**
- **The client's name is self-reported.** `clientInfo.name` is whatever the client sends. A client that calls itself `claude-code` gets the hook route, though only Claude Code runs the hook.
  - Checked live (2026-10-19): on this machine, with the hook installed, a scripted client named `claude-code` ran a side effect with no prompt. The audit line claimed the hook's prompt.
  - Real harnesses send their own name (`copilot-cli`, OpenCode, Codex), and the agent (the model) can't change it.
- **What it stops:** an agent working through the vb6 tools running a side effect the user didn't approve. That covers accidents and over-eager agents, which is what the benchmarks showed: database writes through `debug_eval`.
- **What it doesn't stop:** code running as the user. Two reasons:
  - **Direct access to the add-in.** The elevated add-in serves any process that holds the token from the discovery file, and every process of the user can read that file. The side-effect check runs in the server, not in the add-in. So a program, or an agent with a shell in a harness without the Claude Code hook, can send Immediate input to the add-in directly.
  - **Editing and running the project.** An agent can edit the project's code and start it (`code_replace_lines`, `run_start`). That code runs with the IDE's administrator rights, and no approval is asked.

  Both follow from VB6 running as administrator (README, "Elevated or not?"). Run VB6 without elevation where you can.

In headless runs nobody can approve, so it doesn't run: `claude -p` counts the hook's "ask" as denied, and `copilot -p` (even with `--allow-all-tools`) declines elicitation by itself. **`VB6MCP_EVAL=open`** (installers: `-EvalOpen`) lets such expressions run; the result then has `sideEffect` and a note to undo it, and `audit.log` gets a line before it. A guard rail, not a sandbox: a procedure of the project that writes (`UpdateCard(184)`) isn't recognised, since only the expression's text is checked. The agent's own shell commands are covered by the Claude Code hook (`ShellPolicy`, README). Reading the Immediate window uses the clipboard; the user's clipboard **text** is restored.

**`syntax_check`** results:
```json
{ "compiles": false, "message": "Variable not defined",
  "location": { "Project": "Basic", "Component": "modMath", "line": 8, "column": 5, "endColumn": 6, "procedure": "AddNumbers", "lineText": "    y = 1" } }
{ "compiles": true, "status": "ended", "mode": "design" }
{ "compiles": true, "status": "running-with-dialog", "note": "...", "dialogs": [ ... ] }
```
`outsideEdits` (info only, same shape as in `project_save`) lists unsaved components whose code differs from the last tool edit or from disk: the compile used that text too.

It runs Start With Full Compile, so on success the program **starts** and is ended again at once; startup code (e.g. `Form_Load`) runs briefly. If that code shows a dialog, the program is left running with the dialog reported.
### Debug output (`debug_output`)

Compiled VB6 programs have no `Debug.Print`, because it is compiled out. Programs that log usually call `OutputDebugString` (`Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA"`). `debug_output` captures that output through the DBWIN shared buffer, the protocol DebugView uses:
- **Starting:** the first call starts capturing. Output written before then is lost.
- **Reading:** later calls return the messages after `after_id`. Each call returns `last_id` to pass next time.
- **Filtering:** by default only `VB6.EXE` (the program running in the IDE) and programs started with `exe_run`; `all: true` shows every process.
- **Buffer:** only one monitor can own the DBWIN buffer. If DebugView, a debugger or another vb6mcp session owns it, the tool says so. `stop: true` releases it again.
- **Storage:** messages are kept in memory (the last 2,000). The text is ANSI, decoded with the system code page.

```json
{ "capturing": true, "since": "2026-10-14 00:50:12", "last_id": 3,
  "messages": [ { "id": 3, "time": "00:50:31.204", "pid": 17264, "process": "VB6", "text": "UI fixture: text = Hello" } ] }
```

### Inspecting a stopped program

| Tool | Arguments | Needs mode | What it does |
|------|-----------|------------|--------------|
| `debug_locals` | `module?` (true), `pid?` | break | variables in scope with their runtime type and, for simple types, their value |
| `debug_callstack` | `pid?` | break | the call stack, most recent call first |
| `debug_error` | `action?` (debug, end), `force?`, `pid?` | a runtime error dialog is open | error number and message; `debug` stops at the failing line, `end` ends the program |

**`debug_locals`.** The Locals window is custom-drawn and can't be read as text (spike 18), so the tool builds the list itself:
- **Names:** the current procedure's parameters, return value and `Dim`/`Static`/`ReDim` variables (scanned from its code). With `module`, also the module's variables (from the IDE's `Members`). A local hides a module variable of the same name.
- **Types:** one Immediate line of `TypeName(...)` for all names. `TypeName` raises no run-time error, but it doesn't **compile** for a variable of a user-defined type declared in a standard module: VB shows "Compile error: Only user-defined types defined in public object modules can be coerced to or from a variant" (benchmark T17).
  - In that case the tool closes the dialog, which its own Immediate line caused, and evaluates that chunk one name at a time.
  - Such a variable is reported with type `(user-defined type)`; look at its fields with `debug_eval` (`CStr(products(1).Price)`).
  - Before this fix, one such module variable left every value empty and the dialog open.
- **Values:** a second line, only for Integer, Long, Single, Double, Currency, Decimal, Date, Boolean, Byte and String. `CStr` of Null raises a runtime error dialog, so arrays, objects and Null show their type only; look inside with `debug_eval`. Strings are cut at 200 characters (`length`, `truncated`); `\r` and `\n` stand for line breaks.

No user code runs. Example (DebugFixture, stopped in `Inner`):
```json
{ "Project": "DebugFixture", "Component": "modCalc", "procedure": "Inner", "line": 23,
  "locals": [ { "name": "n", "kind": "parameter", "type": "Long", "value": "4", "declared": "Long" },
              { "name": "tag", "kind": "parameter", "type": "String", "value": "outer/inner", "length": 11 },
              { "name": "Inner", "kind": "return", "type": "Long", "value": "0" },
              { "name": "items", "kind": "local", "type": "Variant()", "value": null, "declared": "Variant" } ],
  "module": [ { "name": "g_runs", "kind": "public", "type": "Long", "value": "1" } ] }
```

**`debug_callstack`.**
- It opens View > Call Stack (a modal dialog, deferred in the add-in), reads the list in the add-in (UIPI blocks reading it from here), and closes it with its Close button.
- An already open Call Stack dialog is read and left open.
- Result: `frames: [ { depth, project, component, procedure } ]`. An entry that isn't `Project.Module.Procedure` (e.g. `[<Non-Basic Code>]`) comes back as `{ depth, text }`.

**`debug_error`.** Every tool that reports dialogs adds `runtimeError: { number, message, next: "debug_error" }` to a VB runtime error dialog.
- Without `action`, `debug_error` reports the number, message and buttons.
- `action: "debug"` presses Debug: the IDE stops at the failing statement, and the result has its `location` (component, line, procedure, line text).
- `action: "end"` presses End, and needs `force: true` when the agent didn't start the run (hard rule 11).
- Run results (`run_start`, `run_continue`, ...) whose dialogs include a runtime error say so in `note` and point to `debug_error`.

**Verified live** (DebugFixture, 2026-10-13), with no help from the user:
1. `run_start`, then `app_click` Compute stopped at the breakpoint in `Inner`.
2. `debug_locals` matched the Locals window (n 4, tag "outer/inner", i 5, acc 10, items `Variant()`, flag True, when 02/01/2003, g_runs 1).
3. `debug_callstack` gave `Inner` ← `Outer` ← `cmdCompute_Click`, and the dialog was closed again.
4. `run_continue`, then `app_click` Error showed the dialog with `runtimeError` 11 "Division by zero".
5. `debug_error` with `debug` stopped at line 27 in `Divide`; `debug_locals` showed b = 0; the call stack was `Divide` ← `cmdError_Click`.
6. `run_continue` raised the error again, and `debug_error` with `end` returned to design mode.

---

### Pending reloads before compiling or running

`syntax_check`, `build_in_ide` and `run_start` first wait (up to 8 s) for reloads the watchdog has pending for that IDE: files changed on disk during a run are reloaded only in design mode, a few seconds after `run_end` (improvement notes 27).
- When there were any, the result has `reloads: {applied, stillPending?, note}`.
- `stillPending` means the IDE still has the previous text of those files.

### Missing DLLs before a run

`run_start` checks the 32-bit DLLs and OCXs behind the open projects' `Object=` and `Reference=` lines for DLLs they need that Windows wouldn't find, imports and delay imports (spike 39, improvement notes 55; the same check as `com_diagnose project=...`). A missing delay-loaded DLL ends the IDE process (exception `0xc06d007e`) the moment the program first uses that component, unsaved work included.
- The run still starts: the code that uses the component may never run.
- When something is missing, the result has `missingDlls: {dlls: [{project, component, chain, delayLoaded}], more?, hint}`. The hint says to `project_save` first when the IDE has unsaved changes.
- The result is kept per `.vbp` (path and write time) for 10 minutes, so repeated runs don't walk the DLLs again.

## The running program

Drive the program being debugged, the way a user would. The program runs **inside `VB6.EXE`**, so its windows belong to the IDE process.

- **Reads** (windows, controls, text) are done by the server with Win32 calls; they work across the elevation boundary (spike 10).
- **Input** goes through the add-in, because a non-elevated process's input to an elevated window is silently dropped (spike 10). It is limited to windows of the IDE's own process, and every input is written to the audit log `%LOCALAPPDATA%\vb6mcp\audit.log` (hard rule 8). So are `debug_eval` and `ide_launch`.
- Every input tool **verifies afterwards**: the IDE mode, dialogs and windows, and for text the value read back.

| Tool | Arguments | Needs mode | What it does |
|------|-----------|------------|--------------|
| `app_windows` | `pid?` | any | the program's top-level windows: `hwnd` (hex), `kind` (Form, MDIForm, Dialog), caption, class, position, size, enabled |
| `app_tree` | `window?`, `pid?` | run, break | the controls of a form: `hwnd`, `name` and, for UserControls and other ActiveX controls, `type` (`XmTextCtrl.XmText`) where the form's `.frm` names them (`namesFrom`: which file, how many; spike 44), `kind` (CommandButton, TextBox, CheckBox, ...), text, position relative to the form's client area; for check boxes and option buttons `value` (0 unchecked, 1 checked, 2 grayed; for a program in the IDE it needs add-in 1.3); for list and combo boxes `itemCount`, the first 20 `items` and `selected` (read through the add-in for a program in the IDE, since UIPI blocks `LB_*` from the server); for text boxes their `index` for `app_set_text` (reading order) and `readOnly` for a Locked one; `depth`, `hidden`, `clipped` (`partly` or `hidden`: outside an ancestor's visible area, e.g. buttons parked below a panel's edge or inside a 2 x 2 PictureBox; less than 4 pixels showing across or down counts as `hidden`), `disabled` only when they apply (compact, see the token breakdown) |
| `app_click` | `control?` or `text?`, `window?`, `force?`, `pid?` | run | clicks a command button, check box or option button. `text` matches the caption without `&`, case-insensitive; when it matches several, the only one a user can see wins. A control the user can't see (window hidden, or `clipped: "hidden"`) is refused unless `force: true` |
| `app_set_text` | `text`, `control?` or `index?`, `window?`, `pid?` | run, no dialog | sets a text box's text (fires `Change`); returns `textNow` and `applied`. `index` counts the form's text boxes in reading order (top to bottom, left to right). A read-only (Locked) box is set but flagged (`readOnly`, with a note): a user couldn't type there (benchmark T18) |
| `app_mouse` | `action` (click, double_click, right_click, drag), `x`, `y`, `to_x?`, `to_y?`, `real_input?`, `window?`, `pid?` | run, no dialog | mouse input at a point relative to the form's client area. By default it posts messages to the deepest window under the point without touching the user's mouse, and a drag goes entirely to the window it starts in. **Labels and Images need `real_input: true`** (spike 17) |
| `app_scroll` | `direction` (up, down, left, right), `amount?` (line, page, end), `count?` (1), `method?` (scrollbar, wheel), `control?`, `window?`, `pid?` | run, no dialog | scrolls a list box, text box, grid, scroll bar control or the form; returns `before`/`after` scroll positions and `moved` |
| `app_select` | `control?`, `text?` or `index?`, `pid?` | run, no dialog | selects an item in a list or combo box and raises its `Click`; returns `selectedNow`/`applied`. Without `control`: the form's only list or combo box, or the only one that has the item `text` |
| `app_keys` | `keys`, `control?`, `real_input?`, `pid?` | run, no dialog | types text and keys (syntax below) into the control or the program's focused control. Without `control` for a program running in the IDE, the focus must be in one of the program's windows: otherwise `real_input` goes to the program's form and posted input is refused (it would reach VB6 itself) |
| `app_wait` | `window`+`gone?` / `control`+`text`/`contains` / `mode` / `dialog`, `timeout?` (10, max 120), `pid?` | any | polls (every 200 ms) until the condition holds; returns `met`, `seen`, `waitedMs` |
| `app_screenshot` | `window?` (or `ide`), `max_width?` (1024), `region?` (`{x, y, width, height}`, client coordinates), `control?` (hex handle), `pid?` | any | an image of the window, plus `{ window, hwnd, width, height, method }`. With `region` or `control` only that part (a control with 4 pixels around it), cropped at full resolution before any scaling, plus `crop` (`x`, `y`, `width`, `height` in client coordinates): `app_mouse` x = `crop.x` + image x / `imageScale` |

`window` is a hex handle from `app_windows` or a caption. Without it, the tools use the program's first form. `control` is the control's hex `hwnd` from `app_tree` (`hwnd` is accepted as another name for the argument). Control names such as `txtPrice` can't be used: they aren't visible from outside the running program.

**Compiled programs:** every `app_*` tool also takes `exe`, the pid returned by `exe_run`, to act on a compiled EXE instead of the program in the IDE.
- Such a program runs at the server's integrity, so the server reads it and sends input itself: posted messages, or `SendInput` for `real_input`. Screenshots are taken by the server.
- Input only goes to programs the agent started with `exe_run` (hard rule 8).
- Posted keys can't be deferred inside the program's thread from outside. The server attaches to the program's input thread to set modifiers, and pauses 60 ms after each key down so the generated character arrives first.
- `mode` in `app_wait` is IDE-only.
- `app_wait` returns early (`met: false`, note `stopped waiting: dialog ... appeared` or `... is open`) while a dialog is open, including one that was already open when the wait began (a MsgBox blocks the program, so the wait could never succeed). Waiting for a dialog (`dialog`) is exempt (improvement notes 32; the live re-test found the first version only stopped for new dialogs, after `app_click` had returned with the MsgBox open).
- `app_set_text` and `app_keys` are refused while a dialog is open, unless `window` names that dialog (caption or hex handle) or `control` is a child of it: that is how a program's own InputBox is answered (improvement notes 36).

| Tool | Arguments | What it does |
|------|-----------|--------------|
| `exe_run` | `path` (an `.exe` under `VB6MCP_ROOTS`), `args?` | starts it non-elevated; waits up to 15 s for its first window; returns `exe` (pid), windows, dialogs. A program that requires elevation is refused with an explanation |
| `exe_end` | `exe`, `force?` | posts `WM_CLOSE` to its forms (so `QueryUnload`/`Unload` run) and waits 5 s; if it's still running, reports its dialogs, or ends it with `force: true` |

**Verified live** (2026-10-14):
- **In the IDE:**
  - `app_select` on the list and the combo box raised `Click`.
  - `app_keys` "Hello{BACKSPACE}{BACKSPACE}p!" gave "Help!", once the continuation after each key down was in place.
  - `{CTRL+LEFT}{SHIFT+END}X` gave "X".
  - `KeyDown` saw Shift.
  - `app_wait` caught each text immediately.
- **On the compiled `Ui.exe` started with `exe_run`:**
  - Every `app_*` tool worked: text with `Change`, check box, combo and list selection, posted keys including Backspace and Shift, scroll bar position (read directly), wheel scrolling, drag with exact coordinates.
  - The Label click and keys worked with `real_input`.
  - Screenshots were taken by the server.
  - `exe_end` closed it with `WM_CLOSE`.
- **Restoring the foreground:**
  - During one real-input call the user was typing in another window, and Windows kept the program in front.
  - The restore now checks the result and falls back to `SwitchToThisWindow`.
  - If the previous window still can't come back, the result carries `warning`.

**Lightweight controls** (Label, Image, Shape, Line) have no window. They are missing from `app_tree` and visible only in `app_screenshot`. To check a Label's text, take a screenshot. To click one, use `app_mouse` with `real_input: true`: VB raises their `Click` only for real mouse input (spike 17).

**`app_click` result:**
```json
{ "action": "clicked CommandButton 'Add'", "mode": "run", "responds": true, "dialogs": [],
  "windows": [ { "hwnd": "1A0C4E", "kind": "Form", "caption": "Basic" } ], "note": null }
```
If the click hits a breakpoint or a runtime error, `mode` becomes `break` and the `note` says so. A `MsgBox` shown by the click is reported in `dialogs`; answer it with `ide_dialog_respond`. `app_click` may target a button of a dialog the program shows, but `app_set_text` refuses while a dialog is open.
`app_click` compares a fingerprint of the program's UI before and after (windows, captions, control texts, enabled/visible, positions, list item counts and selection, read through the add-in for a program in the IDE because UIPI blocks `LB_*` from the server; check-box states for compiled programs only). When nothing changed, the result adds `changed: false` and a note to verify the click took effect (improvement notes 34).

**`app_mouse`:**
- It can't tell whether the input did anything, so check with `app_screenshot` or `app_tree`.
- The result names the window that got the input (`target`), the point in that window's coordinates, and the state afterwards.
- `real_input: true` (add-in `MOUSEREAL`):
  - It brings the form to the front, moves the real cursor and clicks with `SendInput`.
  - About 400 ms later it puts the cursor and the previously active window back.
  - It refuses if another program's window covers the point.
  - Use it for Labels and Images, or wherever posted input has no effect. It briefly takes the user's mouse and focus, so it isn't the default.
- Verified live (UI fixture, 2026-10-13):
  - drag in a PictureBox, with exact coordinates;
  - double-click in a ListBox;
  - Label click: no effect without `real_input`; with `real_input: true` the label's `Click` fired ("Label clicked 1"). Run with Claude in front: afterwards the same window was in front again and the cursor was back at the same pixel.
- **Drag-and-drop** (spike 20):
  - VB's classic drag (`DragMode`, `DragDrop` event) needs `real_input: true`. It worked in the IDE with exact drop coordinates, and posted drags do nothing.
  - **OLE drag (`OLEDragMode`/`OLEDropMode`) could not be started with synthetic input** on the test machine, though it works by hand. The machine is a VM used over RDP, and the RDP client pulls the cursor back during slower drags.
  - Use the program's non-drag path for that, or ask the user.

**`app_keys` syntax:**
- Plain text is typed as characters (Windows-1252; a line break means Enter).
- Braces name keys: `{ENTER}` `{TAB}` `{ESC}` `{BACKSPACE}` `{DELETE}` `{INSERT}` `{SPACE}` `{UP}` `{DOWN}` `{LEFT}` `{RIGHT}` `{HOME}` `{END}` `{PGUP}` `{PGDN}` `{F1}`..`{F12}`, letters and digits.
- Modifiers: `{CTRL+S}` `{SHIFT+TAB}` `{ALT+F}` `{CTRL+SHIFT+END}`.
- `{LBRACE}`/`{RBRACE}` type literal braces.

By default the keys are posted, and modifiers are set in the program's key state between the messages (protocol `KEYS`). `real_input: true` sends real keystrokes instead: it brings the form to the front for a moment and restores the window that was in front. With real input, each character the keyboard layout has a key for is typed as that key (Shift added; `VkKeyScan`), so `KeyDown` sees its real `KeyCode` (66 for `b`), not 231 (VK_PACKET); characters without a plain key (AltGr) go as Unicode. The result has `textNow` for text and combo boxes.

**`app_scroll`:**
- `moved` compares the positions before and after.
  - A window's own scroll bars (list box, text box, form) are read by the server with `GetScrollInfo`.
  - A scroll bar control (`VScrollBar`/`HScrollBar`) is read by the add-in (`SCROLLPOS`), because UIPI blocks those messages from the server (spike 17).
  - The numbers are the raw Win32 positions, not the control's `Value`. VB keeps its scroll bar controls on an internal 0-32767 range: with `Max = 100`, `Value` 5 read as 1622 (live test).
- Windows without a standard scroll bar, such as some third-party grids, return `before: null`; check those with `app_screenshot`.
- `amount: "end"` goes to the top or bottom.
- `method: "wheel"` sends `count` notches at the control's centre.

**`app_screenshot`:**
- When VB6 is not elevated, the server captures the window itself (`method: "server PrintWindow"`).
- When VB6 is elevated, the add-in captures it in-process (protocol `SNAP`), and the server waits up to 5 s for the file.
- Covered windows are captured too.
- The result adds `clientOrigin {x, y}`, the offset of the window's client area inside the image, and `imageScale` when the image was downscaled (`max_width`). `app_mouse` takes client coordinates: `(image x - clientOrigin.x) / imageScale` (improvement notes 35).
- With `window: "ide"` it captures the whole IDE main window. Crop from that for IDE panes: `PrintWindow` on child panes is misaligned (spike 15).

Errors: `the IDE is in design mode; the program must be running` (input tools), `no window 'X'. Open: ...`, `no button ... captioned 'X'`, `N controls are captioned 'X'; pass control`, `'X' is disabled`, `control 1A2B is a CommandButton, not a text box`, and the add-in's 400 when the handle isn't a window of the IDE process.

---

## Form designer

The designer of a form, UserControl or property page as the IDE holds it now, unsaved changes included.

**Served by:** the add-in's deferred `DESIGN` verb (docs/protocol.md). Reading or setting a UserControl's properties runs its own code, so it never runs under the request handler (hard rule 14). The server queues the operation and polls `RESULT`, so each tool is still one call.

**Common rules:**
- Design mode only, with no dialog open.
- The designer window doesn't need to be open (spike 23).
- `control` is `Name`, `Name(index)` for a control-array element, or the component's own name for the form itself.
- Every change is unsaved (`dirty: true`); save it with `project_save`.

### `form_controls`

**Arguments:** `component`, `project?`, `pid?`.

Every control, walking containers from the form down. Each has:
- `name`, `index` (null outside control arrays), `progId` (the `.frm` class, e.g. `VB.CommandButton`, `MSComctlLib.ListView`) and `className`;
- `kind`: `lightweight`, `standard` or `container`. VB reports Label as standard;
- `container`: null on the form;
- `left`, `top`, `width`, `height`: in the container's units, usually twips.

```json
{ "source": "ide", "project": "Basic", "component": "frmMain", "dirty": false, "count": 2,
  "controls": [ { "name": "fraOptions", "index": null, "progId": "VB.Frame", "className": "Frame", "kind": "standard",
                  "container": null, "left": 240, "top": 1800, "width": 3000, "height": 1200 },
                { "name": "chkFast", "index": 0, "progId": "VB.CheckBox", "className": "CheckBox", "kind": "standard",
                  "container": "fraOptions", "left": 120, "top": 300, "width": 2000, "height": 480 } ] }
```

### `form_properties`

**Arguments:** `component`, `control`, `names?`, `project?`, `pid?`.

`names` limits the list to those properties: exact names, or prefixes ending in `.` (`["Caption", "Left", "Font."]`). A control has 50–100 properties, and full lists were about 5% of a redesign run's re-read tokens.

All design-time properties as `{ name, type, value }`:
- `type` is `string`, `number`, `boolean`, `object`, `indexed` or `error`.
- `indexed` properties (List, ItemData) come with `items` instead of `value`. VB returns "" past the end instead of an error, so List ends at its last non-empty item and ItemData has the same count (spike 25).
- An object property's value is a collection of sub-properties (spike 23), so its members are listed too: `Font.Name`, `Font.Size`, `Font.Bold`, `Font.Weight`, ..., and for pictures `Icon.Handle` (0 = none), `Icon.Type`, `Icon.Width`, `Icon.Height`.
- Numbers use `.` (`8.25`).
- Colours are plain numbers: `-2147483633` is the system colour `&H8000000F&`.

### `form_set_property`

**Arguments:** `component`, `control`, `property`, `value` (text), `project?`, `pid?`.

The text is converted to the property's type:
- **numbers:** `1200`, `-3.5`, `&H0080FFFF&`. Anything else is refused; VB's `Val` would have turned `abc` into 0.
- **booleans:** `True`/`False`/`1`/`0`.
- **strings:** as given, in Windows-1252.
- **object members:** set as `Parent.Member` (`Font.Bold`, `Font.Size`, `Font.Name`).
- **pictures** (Picture, Icon, MouseIcon, DragIcon):
  - `@picture:<full path>` loads a .bmp/.ico/.jpg/.gif under the workspace roots. The server and the add-in both check the path.
  - `@picture:` clears it.
  - VB writes the picture into the `.frx` itself, so we never touch `.frx` files. A jpg is stored as an uncompressed bitmap: a 37 KB jpg became a 6.9 MB `.frx`.
- **lists** (List, ItemData): one item per line, set with `IndexedValue(0, 1, ...)`.

Renaming (`Name`) doesn't rename the control's event procedures, just as in the IDE.

Result: `now`, the value read back. For `Font.Bold` that's all the `Font.*` members.

### `form_set_properties`

**Arguments:** `component`, `changes` (list of `{control, property, value}`, at most 500), `verbose?`, `project?`, `pid?`. Design mode.

Several `form_set_property` changes on one form or UserControl in a single call. Values and rules are the same: text converted to the property's type, `Font.*` members, `@picture:<path>` under the workspace roots, lists one item per line.
- **Why:** restyling six real forms took 107–160 separate `form_set_property` calls (benchmark T8 on a private project, report not published), and each MCP round trip costs the agent tokens. The server runs the same deferred `DESIGN SET` per change, so there's no new add-in verb.
- **Checked first:** a malformed item (missing control, property or value) fails the call before anything changes.
- **In order, no rollback:** a value VB refuses (`not a number: x`, error 380) is reported on its item and the rest still run.
- **Stops early** when the IDE gets busy or a dialog opens (a UserControl's code, spike 28). The items after it are counted in `notAttempted`.
- **Fonts:** setting the form's `Font` doesn't restyle the controls already on it (T8); set each control's `Font.*` too.

Result: `applied`, `failed`, `notAttempted`, `dirty`, `stopped` (the reason, or null) and `results`: by default only the **failed** changes, `{control, property, ok: false, error}`. Echoing every change was about 15% of a redesign run's re-read tokens ([token breakdown](benchmarks/2026-10-17-token-breakdown.md)). With `verbose: true`, one per change in order: `{control, property, ok, now}` (the value read back); for a list property, or when VB returns more than the one property, `nowAll` has every record.

Example:

```json
{"component": "frmAddCards", "changes": [
  {"control": "frmAddCards", "property": "BackColor", "value": "&H00FAF7F5&"},
  {"control": "Label1", "property": "Font.Name", "value": "Segoe UI"},
  {"control": "cmdWrite", "property": "Caption", "value": "Write orders"}]}
```

### `form_layout`

**Arguments:** `component`, `rows`, `container?`, `left?`, `top?`, `gap_x?` (120), `gap_y?` (90), `margin?` (120), `column_widths?`, `row_heights?`, `fill?` (true), `valign?` (`center`), `align?` (`left`), `anchor?` (`top`), `size_container?`, `dry_run?`, `project?`, `pid?`. Design mode.

Arranges the controls of one container (the form, a Frame or a PictureBox) as a grid. The agent names the arrangement; the server works out Left/Top/Width/Height and sets them through the same `DESIGN SET` loop as `form_set_properties` (no add-in verb).
- **Why:** the T9–T12 redesigns placed 39–60 controls one at a time with twips worked out by hand, which is where overlapping controls came from (PLAN.md E1).
- **`rows`:** a list of rows, each a list of control names (`Name` or `Name(index)`). `""` is an empty cell. The same name in adjacent cells of a row spans those columns. One row is a row of buttons; one name per row is a column.
- **Sizes:** a column is as wide as its widest single-column control, a row as high as its tallest control, unless `column_widths`/`row_heights` give a number (`null` = automatic). With `fill` every control takes its cell's width (all text boxes of a column the same width); heights are kept. Controls are centred vertically in their row, so a label sits level with the text box beside it.
- **Placement:** from `left`/`top` if given; otherwise `margin` from the container's left (`align: left`), centred, or `margin` from the right (`align: right`), and `margin` from the top or bottom (`anchor`). A Frame's caption takes 180 twips at the top (an 8 pt caption), so its first row starts at 300 by default.
- **Container size:** the form's and a PictureBox's inside size is `ScaleWidth` x `ScaleHeight`; a Frame's is its `Width` x `Height` (its children are placed from its outer corner). `size_container: true` sets the container's `Width`/`Height` so the block plus `margin` fits; for a form or PictureBox the border and title bar (`Width - ScaleWidth`) are added. It needs a block placed from the top-left. A container whose `ScaleMode` isn't twips gets a warning, and isn't resized.
- **Checked first:** names not in the container (with the container they are in), a name used twice, an empty grid, bad options: nothing changes.
- **Only what changed** is set (positions already right are skipped). The result lists `positions`, the `block` rectangle, `containerSize`, `applied`/`failed`/`notAttempted`, `warnings` (the block runs past the container's edge) and the usual `overlaps` with controls outside the grid.
- **`dry_run: true`** returns the positions and how many properties would be set, without changing anything.
- Its own review doesn't add the "You placed N controls by hand ... form_layout" hint; that hint follows `form_set_properties` and `form_add_control` only (improvement notes 38).

Verified live on a production project's `frmFixAccount` (a frame of label pairs fitted with `size_container`, two buttons with `align: right, anchor: bottom`), checked with `form_screenshot`.

Example:

```json
{"component": "frmCustomer", "container": "fraDetails", "size_container": true,
 "rows": [["lblName", "txtName"], ["lblCity", "txtCity"], ["txtNotes", "txtNotes"]]}
{"component": "frmCustomer", "rows": [["cmdOK", "cmdCancel"]], "align": "right", "anchor": "bottom"}
```

### `form_add_control`

**Arguments:** `component`, `prog_id`, `name?`, `container?`, `left?`, `top?`, `width?`, `height?`, `properties?`, `project?`, `pid?`.

`properties` sets properties on the new control in the same call: `{"Caption": "OK", "Font.Bold": "True"}` (values as text, numbers or true/false), with the same conversion, checks and hints as `form_set_properties`.
- A malformed `properties` argument adds nothing.
- A value VB refuses is reported in `properties` (`{property, ok, error}`), and the control stays.

`prog_id` is the class name **as a `.frm` writes it**: `VB.CommandButton`, `VB.Frame`, `MSComctlLib.ListView`.
- It is not the registered ProgID (`MSComctlLib.ListViewCtrl.2`): with that, VB's `Add` returns nothing (spike 23).
- An ActiveX control's component must be in the project first (`refs_add_component`).
- `container` puts it inside a Frame or PictureBox.

Result: `added`, the new control as in `form_controls`. When saved, VB writes the OCX's `Object =` line into the `.frm` header and the `.vbp` itself, consistently (verified).

### Designer rules, hints and overlaps (`form_set_property`, `form_set_properties`, `form_layout`, `form_add_control`)

Verified on this machine (spike [29](spikes/29-designer-rules.md)). The tools act on them or say them in the result:
- **Controls placed by hand** (benchmark T12 round 6: agents computed twips in Python although `form_layout` was loaded): when one call moved at least 4 controls of the same container (`Left`/`Top`), a hint gives the `form_layout` call that arranges them, rows taken from where they ended up (controls within 120 twips of a row's top form a row, left to right), e.g. `form_layout component=frmX container=fraInfo rows=[["lblName","txtName"],["lblCity","txtCity"]]`. The server instructions and `form_add_control` point at `form_layout` too.
- **Changes only to `Font.*` members aren't counted by VB.** `IsDirty` stays False and Save Project skips the form, so the change would be lost. When that happens, the tool sets the form's `Tag` to its own value so VB counts it. If it can't, a hint says so.
- **A form's `Font` doesn't change the controls already on it.** Hint after a form `Font.*` change.
- **A CommandButton shows `BackColor` only with `Style = 1` (Graphical).** Hint when a button's `Style` isn't 1 and the same call doesn't set it.
- **A CommandButton has no `ForeColor`.** VB's error 438 comes back with an explanation.
- **Overlaps:** after a change to `Left`, `Top`, `Width`, `Height`, `Visible` or `Container`, or an added control, `overlaps` lists sibling controls (same container) that overlap a changed control: `{controls: [a, b], container, overlap: {width, height}, sameRectangle}`.
  - A Label overlapping a windowed control (frame, button, text box) is reported, with a note: VB draws labels behind windowed controls, so the overlapped part is hidden. A Label that wholly contains the other control is a background (a coloured header band or status strip with buttons on it) and isn't reported.
  - Not checked: label on label; Images, Shapes, Lines and Timers; members of one control array (often stacked on purpose); overlaps under 30 twips (2 pixels).
  - Controls that are never visible at the same time may overlap on purpose. The hint says so.
  - `form_add_control` without `left`/`top` (and without `Left`/`Top` in `properties`) puts the control at the default position, among others still waiting to be arranged. Its overlaps are only counted in a hint that points to `form_layout`, not listed (improvement notes 43: adding 30 controls that way made every result longer than the last).
  - This costs one extra `CONTROLS` read, only after such changes.
- After any move or added control, a hint says to check the layout with `form_screenshot`.
- `hints` and `overlaps` are left out when empty. A check that couldn't run (a dialog, a busy IDE) is reported in `hints`.

### `form_screenshot`

**Arguments:** `component`, `maximize?` (default true), `max_width?` (default 1024), `project?`, `pid?`. Design mode.

An image of a form, UserControl or property page as the designer shows it now, unsaved changes included. Use it to check layout, fonts and colours without running the program (many real forms need a database or login to run at all).
- The add-in opens the designer (`DESIGN SHOW`, deferred) and maximizes it inside the IDE.
- The server finds the designer's MDI child by the exact caption the add-in reports (VBIDE `DesignerWindow.Caption`, e.g. `frmMain (Form)`, with no project prefix; spike [28](spikes/28-designer-screenshot.md)). If two loaded projects have a component of that name, the designer just activated is on top of the MDI z-order and is the one taken. Inside it, the form surface is the largest visible child window that isn't a scroll bar (`ThunderForm` for a form).
- It captures the whole IDE window, through the add-in when the IDE is elevated (spike 15: child panes captured on their own come out misaligned), and crops it to the form inside the designer's client area.
- **A form bigger than the designer** is scrolled through a page at a time (`WM_VSCROLL`/`WM_HSCROLL` to the designer through the add-in, audited) and the pieces are stitched into one image; then the designer is scrolled back to the top left.

Result: the PNG image plus `component`, `designer` (caption), `formSize`, `stitched` (number of captures, when more than one), `clipped` and `dirty`. `dirty: true` can come from opening the designer itself: UserControls on the form run their code then, and in a production project that marked `frmMain` changed (spike 28), just as opening it by hand does. `clipped: true` means part of the form couldn't be captured (the designer didn't scroll that far); those parts are blank.

The designer stays open afterwards, as if the user had opened it.

**Dialogs:** opening a designer the first time runs the code of UserControls on it. In a production project that showed a database "Login Error" box. This tool and the other designer tools then fail after about a second with the dialog's caption, text and buttons. Answer it with `ide_dialog_respond`; the request finishes by itself, so call the tool again.

**Verified live** (2026-10-15, a production project, spike 28): six forms captured exactly; two larger than the IDE's 763x436 MDI area reported `clipped` (first version). With stitching (T9 round): `frmMain` 688x548 and `frmFixAccount` complete from 2 captures each.

### `form_remove_control`

**Arguments:** `component`, `control`, `confirm: true` (required), `project?`, `pid?`.

Removes the control and any controls inside it. Its event procedures stay in the code (verified), as when deleting by hand; Edit > Undo in VB6 brings it back.

**Verified live** (2026-10-14, scratch copy of Basic, spike 23):
- **Listing:** controls and properties listed with the designer window closed.
- **Setting:** a caption with `–` and `é` (saved as 1252 bytes 0x96/0xE9), a hex colour, a position and `Font.Bold`. `abc` for `Left` was refused.
- **Adding:** a Frame; a CheckBox inside it; a ListView (after `refs_add_component`).
- **Control arrays:** `Index` set to 0, then addressed as `chkFast(0)`.
- **Removing:** a button, with its `Click` code kept.
- **After saving:** the `.frm` had the nested Frame/CheckBox, `Index = 0`, a `BeginProperty Font` block with `Weight = 700`, the ListView and its `Object =` header line, all in CR LF.
- **Second round:** the form's own Caption was set (the component's name as `control`). `MSComctlLib.ListViewCtrl.2` was refused with the class-name hint. Picture properties list as `Icon.Handle`, `Icon.Type`, `Icon.Width`, `Icon.Height` (a non-zero handle means a picture is set). Their members are informational: pictures can't be set.

---

### `form_resources`

What a form, UserControl or property page keeps in its binary companion file (`.frx`/`.ctx`/`.pgx`), read from disk (spike 34). The owner's case: taking an image out of a form; VBIDE's `SavePicture` would turn it into a BMP, while the `.frx` keeps the original (a GIF stays a GIF), and it works for closed projects.

**Arguments:** `path` (the `.frm`/`.ctl`/`.pag`/`.dsr`/`.dob` under the roots), `extract?` (`["imgLogo.Picture", "ImageList1.Images.ListImage3.Picture"]` or `["all"]` for every picture), `out_dir?` (a folder under the roots).

- **Listing** (no `extract`): every property stored in the companion file: `name` (`Control.Property`, with its property group, e.g. an ImageList image `ImageList1.Images.ListImage1.Picture`), `kind` (`picture`, `empty picture`, `list`, `text`, `font`, `binary`), `imageType` (`gif`, `jpg`, `bmp`, `ico`, `cur`, `wmf`, `emf`, `png`), `bytes`, `key` (an ImageList image's Key), `items` (List/ItemData), `text` (long and multi-line texts, the font's face name).
- **Extracting:** with `out_dir`, pictures are written in their original format and other blocks (an ActiveX control's saved state, a font) as `.bin`, named `<form>_<control>_<property>.<ext>`; existing files are never overwritten (`-2`, `-3`… are added). Without `out_dir`, the first requested picture is shown as an image (PNG).
- **Secrets:** `Password=`/`Pwd=` values in texts (connection strings in `DesignConnection` and the like) are masked.
- **Disk only:** an open project's unsaved designer changes aren't in the `.frx` until VB saves; `warning` says when a running IDE has the form dirty.
- **Read-only** (hard rule 2): change pictures and lists with `form_set_property` (`@picture:`, `List`); VB rewrites the `.frx` on save. Audited.

Verified on a production project's `frmMain` (an ICO, a GIF89a, ICOs and a BMP extracted byte-exact; two masked connection strings; SpyWorks state as binary) and another production project's `frmMain` (4 ImageList GIFs, SSTab empty pictures).

## References and components

Live from the IDE (VBIDE `References`).

**Served by:** the add-in's `REFS`/`REFADD`/`REFDEL`/`TOOLBOX` verbs (synchronous: no user code runs).

**Rules:**
- **Mode:** changes need design mode.
- **GUIDs** come only from the registry or a typelib (hard rule 13). Names and GUIDs are resolved against `HKCR\TypeLib` in the 32-bit view, where VB resolves them too.
- **File paths** must be under the workspace roots, or be the registered file of a typelib (hard rule 8).

### `refs_list`

**Arguments:** `project?`, `pid?`.

- **`references`:** `{ Name, libId, version, file, builtIn, broken, kind (typelib|project), description }`, live from the IDE.
- **`components`:** the `.vbp`'s `Object=` lines, read **from disk** because VBIDE can't list them, each with its registered version, description and file.

### `refs_add`

Exactly one of:
- **`name`:** part of a registered description. Registered versions whose file exists are searched, keeping the newest version per LIBID. One match is added. Several are returned as `candidates` (libId, version, description, file), and nothing is added.
- **`libid` + `version`:** e.g. `2.8`. It must be registered; the registered versions are listed otherwise.
- **`file`:** a DLL/OCX/TLB, or a `.vbp` for a project reference inside a group. It must be under the roots or a registered typelib's file. VB adds an unregistered DLL without registering it; the reference then has an empty `file` and `description`.

VB's own errors pass through, e.g. "Name conflicts with existing module, project, or object library" for a duplicate.

```json
{ "source": "ide", "project": "Basic",
  "added": { "Name": "ADODB", "libId": "{2A75196C-D9EB-4129-B803-931327F72D5C}", "version": "2.8",
             "file": "C:\\Program Files (x86)\\Common Files\\System\\ado\\msado28.tlb", "broken": false, "kind": "typelib",
             "description": "Microsoft ActiveX Data Objects 2.8 Library" }, "note": "unsaved in the IDE; save with project_save" }
```

### `refs_remove`

**Arguments:** `name` (as `refs_list` shows it), `project?`, `pid?`. Built-in references (VBA, VBRUN, VB) are refused.

`name` may also be `#<position>`, 1-based in `refs_list` order. A broken reference has no readable name. VB removes a broken reference **only while its typelib is registered**; otherwise "Object library not registered" (spike 25). `refs_fix_broken` gives the steps.

### `refs_fix_broken`

**Arguments:** `project?`, `pid?`. Report only.

**For each broken (MISSING:) reference:**
- `position`; the name is usually unreadable.
- What the `.vbp` asked for: LIBID, major version, the hint path and whether it exists.
- What's registered for that LIBID.
- `filesWithThisTypeLib`: the hint path under the roots, plus every workspace build output and compatibility copy carrying exactly that LIBID and major version.
- `fix`.

**The fix, in order:**
1. **Same major version registered:** `refs_remove name=#<position>`, then `refs_add libid … version …`.
2. **A file carrying it exists:** `com_register` it. VB re-resolves the reference as soon as the typelib is registered: verified, the reference healed with no further step. If it's still broken, remove and re-add.
3. **Only another major version is registered:** `refs_add` it. The broken entry can't be removed through VBIDE until its own version is registered.
4. **Nothing carries it:** VBIDE can't remove it. Install or build the component, or the user unchecks it in Project > References.

**Not broken, though unregistered:** when the typelib isn't registered but the `.vbp` hint path points at an existing file, VB loads it from there (spike 25).

### `refs_add_component`

**Arguments:** `prog_id`: the **registered** ProgID of one of the OCX's controls (`MSComctlLib.ListViewCtrl.2`), checked in the registry (32-bit view). This calls `VBProject.AddToolboxProgID`. Result: the ProgID, its CLSID and server file. VBIDE can't remove components.

**Verified live** (2026-10-14, spike 23):
- **By name:** "Scripting Runtime" added `Scripting`; "ActiveX Data Objects" listed 9 candidates.
- **By libid:** ADODB 2.8 added.
- **By file:** the unregistered ComFixture.dll was added. The saved `.vbp` line `Reference=*\G{9FAE05B6-…}#1.0#0#..\ComCompat\ComFixture.dll#ComFixture` carries exactly the LIBID our typelib reader reads from the file, which closes `ws_affected`'s open question.
- **Refused:** notepad.exe (outside the roots and not a typelib), an unregistered GUID, a duplicate (VB's error 32813), and removing VBA.
- **Component:** MSCOMCTL added (`Object={831FDD16-…}#2.2#0; mscomctl.OCX` after save).

---

## COM components

What components offer, what's registered where, binary compatibility, and (un)registration.

**Served by:**
- **Reads:** `vb6mcp.exe` itself (`Vb6Mcp.Core.Files`): the registry in the 32-bit view (where VB6 looks), and our own typelib reader (`LoadTypeLibEx(REGKIND_NONE)`, nothing registered, one read at a time).
- **(Un)registration:** the add-in's `REGSVR` verb, elevated (spike 24).

GUIDs are only ever read, never typed in (hard rule 13).

### `com_describe`

**Arguments:** exactly one of `file` (DLL/OCX/TLB/OLB under the roots, or a registered typelib's file), `prog_id`, or `libid` + `version?` (newest registered if omitted). Optionally `filter` (type names containing it) and `include_hidden`.

The type library as VB-like declarations:
- `Class X` with its `Implements` lines (`[default]`, `[source]` = events). Marked `' control` for ActiveX controls, `' not creatable` for PublicNotCreatable classes.
- Interfaces with `Sub`/`Function`/`Property Get|Let|Set` and parameters: `ByVal`, `Optional`, defaults, `ParamArray`, VB types.
- Hidden types are left out (`include_hidden` shows them), except interfaces a visible class implements. VB marks every class's own interface hidden (`_clsTax`, events `__clsTax`), so a VB-built DLL used to show only `Implements _clsTax` (benchmark T15). Hidden and restricted members stay out unless `include_hidden` is set. The seven restricted `IUnknown`/`IDispatch` methods (`QueryInterface` … `Invoke`) stay out even then, since VB code can't call them (improvement notes 45).
- `Enum`s with values, and `Type`s.
- Help strings as comments.
- **Default members** (improvement notes 14, PLAN.md C1): the member with `DISPID_VALUE` (memid 0) is marked `' default member`, and listed even when hidden. A class line names its default interface's default member: `Class MSComm ' control ... ' default member: CommPort (as hidden _CommPort)`. That's what VB uses when code names the object itself (`x = MSComm1`, `If MSComm1 Then`). A hidden `_X` alias of a visible `X` is reported as `X`.

Without a filter, a listing over 60,000 characters returns the type list (name, kind, member count) instead, so the agent can then ask for one type.

```json
{ "typeLib": { "libId": "{420B2830-E718-11CF-893D-00A0C9054228}", "version": "1.0", "name": "Scripting", "description": "Microsoft Scripting Runtime", "file": "C:\\Windows\\SysWOW64\\scrrun.dll" },
  "declarations": "' Scripting {420B2830-…} v1.0 - Microsoft Scripting Runtime\r\n\r\ndispinterface IDictionary ...\r\n    Sub Add(Key As Variant, Item As Variant)  ' Add a new key and item to the dictionary.\r\n..." }
```

The same text is available on the command line: `vb6mcp describe <file> [filter]`.

### `com_search`

**Arguments:** `text`, `match?` (`word`, the default, or `substring`).
- **Typelibs:** registered ones whose description or file name matches (libId, version, description, file, fileExists).
- **ProgIDs:** registered ones whose name matches (progId, clsid, server file, fileExists).
- **`word`** (improvement notes 44): the text must start a word. Words split at anything that isn't a letter or digit, and where a lower-case letter is followed by an upper-case one: `Ml` matches `MlText.ocx` and `Ml Controls`, not `MSXML6`. A text with spaces or dots counts when it starts at a word (`Common Cont`, `ADODB.Rec`). Whole words come first. When nothing matches that way, the result falls back to substrings and says so (`match: "substring"`, `note`).
- **`.oca` files:** VB writes an `.oca` next to an `.ocx` it loads (its cache of the control's type information), and may register it as a typelib. Those entries come after the real files, with a `note`.
- Live on this machine: a company's two-letter prefix gave 21 typelibs (its controls, then their `.oca` files) and 10 ProgIDs. As a substring it hit the 50-ProgID limit with MSXML.

At most 50 of each.

### `com_where`

**Arguments:** exactly one of `prog_id`, `clsid`, `libid`, `file`.
- **`prog_id`/`clsid`:** the server file, whether it exists and when it was written, the threading model, `builtBy` (workspace projects whose build output is that file), and the class's typelib with its registered versions.
- **`libid`:** every registered version with its file, `builtBy`, and the projects whose current build has that LIBID.
- **`file`:** its typelib (read from the file) and each creatable or control class, each with `registeredTo` and `here` (whether the registry points at this file), plus `registeredHere` (everything points here).

```json
{ "file": "D:\\src\\Controls\\XmText.ocx", "written": "2025-07-22 08:30:57", "builtBy": ["D:\\src\\ControlsSrc\\XmText\\XmText.vbp"],
  "TypeLib": { "libId": "{ABFCD1E9-…}", "version": "1.0", "name": "XmTextCtrl", "registeredTo": "C:\\Program Files (x86)\\Common Files\\Acme\\Controls\\XmText.ocx", "here": false },
  "Classes": [ { "name": "XmText", "clsid": "{FCBBA601-…}", "progId": "XmTextCtrl.XmText", "registeredTo": "C:\\Program Files (x86)\\Common Files\\Acme\\Controls\\XmText.ocx", "here": false } ],
  "registeredHere": false }
```

### `compat_status`

**Arguments:** `project?` (a `.vbp` path, or a name filter; default every ActiveX DLL/EXE/OCX project under the roots), `refresh?`.

Per project:
- `mode` (none/project/binary), `version` (MajorVer.MinorVer.RevisionVer);
- the compatibility file and the build output, each with whether it exists and its typelib (LIBID, version);
- what's registered for that LIBID;
- `findings`.

| Level | Finding |
|-------|---------|
| warning | not binary compatibility. Verified: without it every build gets a new LIBID and new CLSIDs, and the old ones are unregistered (spike 24) |
| error | binary compatibility without a compatibility file, the file missing, or no readable typelib in it |
| warning | the compatibility file is the build output itself (each build replaces its own reference) |
| error | the output's LIBID differs from the compatibility file's (built without, or broke, compatibility) |
| warning | the output's major version differs from the compatibility file's |
| warning | no build output; typelib not registered; or registered, but not for the build output (e.g. an installed copy) |

It only reports; it never changes compatibility settings (hard rule 13).

### `com_register` / `com_unregister`

**Arguments:** `file` (a `.dll`/`.ocx` under the workspace roots), `confirm: true` (required), `force?` (unregister only), `elevate?` (allow a UAC prompt), `pid?` (which IDE's add-in runs it).

**How it runs:** `regsvr32 /s` (32-bit). Registering needs administrator rights (spike 41), so it runs through the first of these that is elevated:
1. **The add-in** of the IDE named by `pid`, or of an elevated IDE. No UAC prompt. The add-in checks the path again on its side. (With `pid` naming an IDE that isn't elevated, regsvr32 runs anyway and fails with exit 5; the `note` says why.)
2. **The elevated watchdog** (`REGSVR`, docs/protocol.md): no UAC prompt, and no IDE needs to be open. It checks the path against its own roots file.
3. **A UAC prompt**, only with `elevate: true`: the server starts regsvr32 with ShellExecute `runas`, and the user sees Windows' prompt naming regsvr32. Declining is an error ("the user declined the UAC prompt: nothing was registered"); the 60 s wait starts after the answer.
4. Otherwise it's **refused** before anything runs, and the error says to pass `elevate: true` or to start VB6 or the watchdog as administrator.

`via` in the result says which one ran it. Each call, including a declined prompt, goes to the audit log.

**Result:**
- `exitCode`.
- `before` and `after`: the file's typelib and classes, with where the registry points.
- `verified`: after unregistering, nothing points at this file. After registering: regsvr32 exited 0, the typelib points here and no class points at another file (or simply everything points here).
  - Creatable classes of the typelib that **nothing** registers are listed in `notRegistered` with a note, not counted as a failure (improvement notes 18: `CapPRO.dll` registers 5 of its 6 classes; the sixth, `RivetTalk`, is an internal helper). If no class at all got registered, it isn't verified.
  - Classes registered to **another** file are in `registeredElsewhere` and fail the check (`com_where` shows which file).

`com_unregister` is refused when some of the file's classes are registered to another file, because VB's unregistration removes the class keys whoever they point at. Pass `force: true` if that's intended.

**Verified live** (2026-10-15, spike 24):
- **Registering:** scratch `ComFixture.dll` registered, then unregistered. Both `verified` and exit code 0.
- **Refused:** unregistering a second copy with the same CLSIDs (`force` missing), `confirm: false`, and a system DLL outside the roots. Raw `REGSVR` sent to the add-in, bypassing the server, was refused by the add-in itself for a system DLL, a `..` path and a `.cls` file.
- **Rebuilding without compatibility:** two `build_make` runs of the same source without compatibility gave different LIBIDs and CLSIDs, and the first build's registrations were gone.
- **Read tools** on the owner's 12 ActiveX projects: the findings above, and `com_where` traced XmText.ocx to an installed copy.

### `com_orphans`

**Arguments:** `text?` (only entries whose ProgID, description or file contains it), `under_roots?` (only files under the workspace roots). Report only.

**`vb6DebugRegistrations`** (spike 25). While an ActiveX project runs in the IDE, VB registers its classes to `VB98\VB6DEBUG.DLL` and its typelib to the `.vbp`. End puts the previous registration back, and so does the next start of VB6.EXE after a crash. Until then a killed IDE leaves them behind, and programs can't create those classes. Per `.vbp`:
- `state`: debugging now, or left over;
- `libId`, `classes`;
- `repair`: `com_register` of a build carrying that typelib. That's the project's own output, or another workspace build or compatibility copy with the same LIBID; otherwise "build it".

**`classesWithMissingServer`, `typeLibsWithMissingFile`:** registrations whose file doesn't exist as a 32-bit program sees it, at most 100 each with the total. These aren't counted as missing:
- bare system names (`combase.dll`, found through the system folders);
- typelib resource suffixes (`vbscript.dll\2`);
- typelibs without a 32-bit file.

On the owner's machine: 137 classes and 15 typelibs, e.g. Codejock OCXs registered from a deleted installer folder.

The scan covers ~9,100 classes and ~800 typelibs and takes a few seconds.

### `com_diagnose`

**Arguments:** `error` (VB runtime error number) and exactly one of `prog_id`, `clsid`, `file`, `project` (a `.vbp`).

**Result:** `meaning`, VB's message and its usual causes for 429, 430, 48 and 53, plus `checks` (`item`, `status` ok/warning/error, `detail`) and the number of `problems`.

**Checks:**
- **For a class:** ProgID → CLSID; class registered; server file present and **x86** (VB6 programs are 32-bit; a PE header check); server = VB6DEBUG.DLL (debugger leftover); typelib registered, with its file, or registered to a `.vbp`.
- **For a file:** whether it's registered to this file.
- **For a project:** each `Reference=`/`Object=` line is checked the same way.
- **Dependencies** (spike 39): for every 32-bit DLL/OCX checked, the DLLs it imports and delay-loads, and theirs, are looked up the way the loader does (the DLL's folder for imports; the VB98 and project folders for delay-loads, which a plain `LoadLibrary` searches; then SysWOW64, Windows, PATH; WinSxS for the VC 8/9 runtimes). Only problems are listed: `X.dll -> Y.dll -> Z.dll: not found`, saying when a delay-load is in the chain (the program then dies with 0xc06d007e on first use, the IDE too). Windows' own DLLs aren't walked. No code of the files runs: the tables are read from the file.
- **When a workspace project builds the file:** its `compat_status` findings, which matter for 430.

**Verified live:** 429 on `ComFixture.clsGreeter` after an IDE kill reported "registered to the VB6 debugger" and "typelib registered to a project file, left over". On Form1x it found its unregistered ComFixture reference.

---

## Builds

VB6 `/make` refuses a project that is open in any IDE (spike 21). So the tools split by whether the project is open:

| Tool | Arguments | What it does |
|------|-----------|--------------|
| `build_make` | `path?` (`.vbp`), `project?`, `pid?`, `defines?` | **Closed project** (`path` under `VB6MCP_ROOTS`): `VB6.EXE /make` from the files on disk, started elevated by an add-in, or by the watchdog when no IDE runs (PLAN.md D1: `via` says which; with neither, the error says how to start one), so an ActiveX DLL/OCX is registered on success. When the running VB6 isn't elevated but an elevated watchdog runs, the watchdog starts it, so the build still registers. Started by a VB6 that isn't elevated, `/make` writes the DLL/OCX but can't register it ("Error accessing the system registry"): then `registered: false` and a `note` with the cause and the way out (`com_register elevate=true`, spike 41). Returns `status` (`succeeded`, `failed`, `dialog`, `timeout`), `errors` (file, line, message), the raw `log`, and the output file before/after. **Open project:** if saved, built in its IDE (File > Make; `via` says so); with unsaved changes, refused |
| `build_in_ide` | `project?`, `pid?` | File > Make from IDE memory (`MakeCompiledFile`), unsaved edits included; design mode. `status`: `succeeded`, `failed` (with `compileCheck`: message and location from a full compile, because `MakeCompiledFile` itself only fails generically), `make-failed` (the full compile passed, so the make step itself failed: improvement notes 20; `causes` lists what the server can see: output folder missing or not writable, output file read-only or in use), `dialog`, `timeout`. A project with no components loaded is refused up front (VB would open Project Properties, improvement notes 9). An existing output file is replaced without a prompt |

Notes from spike 21:
- **`defines`:** conditional compilation constants for this build, e.g. `{"noterm": 1}`, passed as `/make /d noterm=1` (improvement notes 7). Integers only (VB's `True` is -1). For closed projects only: an open project builds through VB's Make, and VBIDE can't set its constants. Use `project_settings` for a lasting change.
- **Load errors:** when VB can't load a form or control ("Errors during load. Refer to '…\frmMain.log' for details", e.g. a licensed control without its licence, spike 27), the error has `kind: "load"`, `file` is that `.log`, and `details` holds its lines (up to 50), e.g. `Line 12: Cannot load control WinHook1; license not found`.
- **Error lines:** `/make` counts the IDE's code lines from 0. `errors[].line` is the 1-based IDE line (as in `code_read`) and `fileLine` the line in the file on disk.
- **Binary compatibility break:** `/make` fails with a `kind: "compatibility"` error naming the class and member. Nothing is written or registered.
- **In the IDE:** the same break is a decision dialog. `build_in_ide` returns `status: "compatibility-decision"` and leaves it for the user (hard rule 13).

**Verified live** (2026-10-14): closed EXE and ActiveX DLL builds (the DLL registered and callable), compile errors from `/make` and from memory with location, and a binary compatibility break both ways.

---
## Sync: files changed outside the IDE

The watchdog's sync engine checks the files of every loaded project every 2 seconds (PLAN.md §8):
- **How it knows:** a file whose time stamp or size changed without the IDE writing it (the add-in records the IDE's own saves) was changed outside: git, another editor, or another IDE that has the same shared file. The change must look the same on the next check before anything happens, so half-written files are left alone.
- **Clean component, design mode:** reloaded automatically (the add-in checks again right before).
- **Run or break mode:** reloaded once the IDE is back in design mode.
- **Component with unsaved IDE changes:** a **conflict**. Nothing is touched; it is resolved with the tools below.
- **`.vbp` changed, file deleted:** reported only. Reloading a project is disruptive and left to the user.

| Tool | Arguments | What it does |
|------|-----------|--------------|
| `sync_status` | `pid?` | conflicts (`pid`, `project`, `component`, `file`, `detected`) and the last 30 events (`reloaded`, `pending`, `conflict`, `resolved`, `not-reloaded`, `project-changed`, `deleted`) |
| `sync_diff` | `component?`, `project?`, `pid?` | unified diff of the IDE text (unsaved changes included) against the file on disk (`--- ide`, `+++ disk`); `same` when equal. **Without `component`:** a project summary (improvement notes 23): counts `same`/`caseOnly`/`differs`, and for each component that differs its `state`: `case-only` with `respelled` (`"Number -> number"`, disk -> IDE: VB keeps one spelling per name project-wide, so new code re-cases other modules), `differs` with `changedLines`, `never saved`, `file missing` |
| `sync_take_disk` | `component`, `confirm` (required true), `project?`, `pid?` | keeps a recovery copy of the IDE text (if unsaved) in its own recovery session (`kept`; listed by `recovery_list` with a note, kept 7 days), then reloads the component from disk, **discarding the IDE changes**; design mode. Verifies the IDE text equals the file afterwards |
| `sync_take_ide` | `component`, `confirm` (required true), `project?`, `pid?` | saves the project, writing the IDE version over the changed file. **All** unsaved components of the project are saved (`alsoSaved` lists them) |

A conflict ends by itself when the component has no unsaved changes any more: after either tool, or after the user saves or reloads in VB6.

**Verified live** (2026-10-14, scratch copy of Basic):
- An outside change to an unmodified module was reloaded 4 s later, also while the "Add to SourceSafe?" prompt was open.
- An outside change to a module with an unsaved edit became a conflict, the edit untouched; `sync_diff` showed both sides.
- `sync_take_disk` reloaded from disk, and the conflict resolved itself on the next poll.
- **Found in that test:** the recovery poller deleted the kept copy, because the component was no longer unsaved. Kept copies now live in their own session; verified in the second round.
- `sync_take_ide` saved without a group dialog, and the conflict resolved as "the IDE saved its version". The save was recognised as the IDE's own, with no reload (spike 22).

---

## Workspace (disk)

Disk is the secondary channel (principle 4). These tools work on the **workspace roots** configured in `VB6MCP_ROOTS` (semicolon-separated folders; set by `scripts\install-server.ps1 -Roots`). Paths outside the roots are refused (hard rule 8); without roots the tools explain how to configure them.

**Served by:** `vb6mcp.exe` itself (`Vb6Mcp.Core.Files`), no add-in needed, except `file_read`'s "is it open in an IDE?" check.

### `ws_projects`

**Arguments:** `filter?` (substring of name or path), `refresh?` (rescan).

The **project map** is built on first use by scanning the roots for `.vbp` files (skipping `.git`, `.vs`, `bin`, `obj`), cached, and rebuilt when older than 5 minutes or on `refresh`. On the owner's workspace: 43 projects, 1,498 files, 139 shared files; 4.9 s cold, 0.34 s warm (spike 9).

**Result:**
```json
{ "roots": ["D:\\src"], "scannedMs": 338, "total": 43, "matched": 1,
  "projects": [ { "Name": "Reports", "Type": "Exe", "path": "D:\\...\\Reports.vbp", "members": 57,
                  "output": "D:\\...\\Reports.exe", "compatibility": "none", "compatibleFile": null } ],
  "truncated": false, "errors": [] }
```
At most 200 projects are listed (`truncated` says if there were more).

### `ws_file_projects`

**Arguments:** `path`, `refresh?`.

Which projects include a source file. Paths are compared normalized (hard rule 3): `..\..\Shared\x.bas` from one project and `D:\src\Shared\X.BAS` are the same file.

**Result:**
```json
{ "path": "D:\\src\\Shared\\modUtils.bas", "shared": true,
  "projects": [ { "Name": "App1", "Type": "Exe", "path": "D:\\src\\App1\\App1.vbp", "asMember": "..\\Shared\\modUtils.bas" } ] }
```

### `ws_affected`

**Arguments:** `paths` (changed source files; a JSON array, or one string separated by `;`), `refresh?`.

Which projects a change affects, **in rebuild order** (`Vb6Mcp.Core.Files.Affected`):
1. **Direct:** the projects that include one of the files (`includes`).
2. **Dependents, transitively:** every project that references an affected project's COM output (`uses`).
   - A `Reference=*\G{LIBID}#maj.min#…` or `Object={LIBID}#maj.min#…` line matches a project when the typelib in that project's **built output** (else its binary compatibility file) has that LIBID and major version.
   - The typelib is read with `LoadTypeLibEx(REGKIND_NONE)`: nothing is registered, and it runs on its own STA thread with a 10 s timeout.
   - Never matched by file name (hard rule 13).
3. **Order:** dependencies come before the projects that use them, ties by name.

`openIn` lists the IDEs that have the project open. Build those with `build_in_ide` (`/make` refuses them, spike 21) and closed ones with `build_make`.

**Warnings** (the result still lists what it found):
- A file no project includes.
- A DLL/OCX with no readable typelib (never built): its users can't be found until it's built once.
- Two projects that build the same typelib.
- Circular references (listed in name order).
- An IDE that didn't answer.

**Not covered yet:**
- Project references inside groups (`*\A…vbp`), until verified.
- `.vbp` versions of 10 and above: parsed as hex like the registry's TypeLib keys, but not verified.

**Verified:**
- For a VB-built DLL, VB itself wrote `Reference=*\G{9FAE05B6-…}#1.0#…` (added with `refs_add`, spike 23), the same LIBID the reader reads from the DLL.
- The LIBID the reader returns for `stdole2.tlb` is the `HKCR\TypeLib` key that VB resolves `Reference=*\G{00020430-…}#2.0` (a line VB wrote itself, in a scratch project) through.
- Order, transitive `Object=` users, other major versions, unbuilt DLLs and cycles are covered by unit tests (`AffectedTests`).
- **Live, on the owner's workspace** (43 projects, 2026-10-14): `Common\modBinding.bas` is included by 7 UserControl projects (binary compatibility).
  - The typelib of each project's built OCX was read, and 23 applications were found through their VB-written `Object=` lines, ordered after the OCXs.
  - Cross-checked by hand: exactly 18 `.vbp` lines name XmText's LIBID, and 18 projects were listed as using it.
  - No warnings.

**Result:**
```json
{ "files": ["D:\\src\\Shared\\modUtils.bas"],
  "order": [
    { "step": 1, "Name": "Lib", "Type": "OleDll", "path": "D:\\src\\Lib\\Lib.vbp", "output": "D:\\src\\Lib\\Lib.dll", "compatibility": "binary",
      "includes": ["D:\\src\\Shared\\modUtils.bas"], "uses": [],
      "typeLib": { "libId": "{…}", "version": "1.0", "Name": "Lib" }, "openIn": [] },
    { "step": 2, "Name": "App", "Type": "Exe", "path": "D:\\src\\App\\App.vbp", "output": "D:\\src\\App\\App.exe", "compatibility": "none",
      "includes": [], "uses": ["Lib"], "typeLib": null, "openIn": [ { "pid": 4242, "project": "App", "dirty": false } ] } ],
  "warnings": [], "note": "…" }
```

### `file_replace_lines`

Takes `edits` like `code_replace_lines` (several ranges, one hash, applied bottom-up). On disk all edits are applied in memory and the file is written once, so a failing edit changes nothing. When the file is open in an IDE, the edits go there one by one, as with `code_replace_lines`.

Edits the code of a VB6 source file **of a closed project** on disk. Arguments: `path`, `start_line` (1-based, the code as the IDE shows it and as `file_read` returns it), `count` (0 inserts before `start_line`), `new_text`, `expected_hash` (from `file_read`), and `pid?` when the file is open in several IDEs.

- **On disk** (the file isn't open in any IDE):
  - Only the code lines change; header, designer block and everything else stay byte-exact (hard rule 2), and the file stays Windows-1252 with CR LF.
  - The new text is validated: 1252 only, 1023 characters per line, 24 continuations.
  - Hidden inline `Attribute` lines (e.g. `VB_UserMemId` after a default property's signature, `VB_VarHelpID` after a `WithEvents` declaration) are put back after the same line when the new text keeps it; otherwise the edit is refused. This includes attributes right after the range's last line.
  - `drop_var_help_ids: true` lets an edit drop a `WithEvents` variable's `VB_VarHelpID` instead, because VB writes it again on save (spike [30](spikes/30-withevents-attribute.md)).
- **Open in an IDE:** the edit goes to the IDE instead, because its memory is the truth (hard rule 4). It stays unsaved there; `routedTo` names the component.
  - Refused if the IDE's text doesn't match `expected_hash`: it has unsaved changes, so use `code_read`.
  - Refused if the IDE isn't in design mode.
- **An IDE didn't answer:** refused, because a busy IDE could have the file open with unsaved changes.

Result: `{ "source": "disk", "path": "...", "component": "modShared", "countOfLines": 5, "hash": "..." }`, or `source: "ide"` with `routedTo`.

**Verified live** (2026-10-14):
- **Closed file:** a scratch copy of Basic's `clsCounter.cls` (class header, no final line break), with one line replaced by two, one of them containing `é`.
  - The result matched the expected bytes exactly: header untouched, CR LF everywhere, `é` stored as 0xE9, and still no final line break.
  - A second edit with the old hash was refused.
- **Open file:** an insert into `Sync2\modMath.bas`, open and clean in VB6.
  - It went to the IDE (`source: "ide"`, `dirty: true`, and `code_read` showed the line).
  - The disk file's SHA-1 didn't change, and the sync engine logged no new event.

### Claude Code hook

`install-server.ps1` registers `vb6mcp.exe hook` as a `PreToolUse` hook (matcher `Edit|Write|MultiEdit`) in `~/.claude/settings.json`, keeping a backup; `-NoHook` skips it.
- **Blocked:** edits of VB6 source and project files (`.bas .cls .frm .ctl .pag .dsr .dob .vbp .vbg`). The hook exits with 2 and tells the agent to use `code_replace_lines`/`proc_replace` (open in the IDE) or `file_replace_lines` (closed project).
- **Why:** the generic tools would write UTF-8, break the hidden header and Attribute lines, and bypass a running IDE.
- **Allowed:** reading those files and every other file type.
- **Verified live** (2026-10-14): right after `install-server.ps1`, with no restart needed, an Edit of a scratch `.bas` was blocked and the agent got the reason text.
### `file_read`

**Arguments:** `path`, `start_line?`, `count?`.

Reads a `.bas/.cls/.frm/.ctl/.pag/.dsr` from disk **the way the IDE shows it**: code only, without the designer block or the hidden `Attribute` lines (the line numbers match the IDE's).

**Result:**
```json
{ "source": "disk", "path": "D:\\...\\modMath.bas", "component": "modMath", "kind": "Module",
  "countOfLines": 13, "startLine": 1,
  "openInIde": [ { "pid": 10856, "project": "Basic", "component": "modMath", "dirty": true } ],
  "warning": "a running IDE has unsaved changes to this file; its memory is the truth (use code_read)",
  "text": "Option Explicit\r\n..." }
```
`openInIde` lists every running IDE that has the file loaded (hard rule 4: when it's loaded, the IDE's memory is the truth).

### `res_read`

A project's resource file (`.res`, the `.vbp`'s `ResFile32`), read-only (spike 35). VB code loads resources by id (`LoadResString`, `LoadResPicture`, `LoadResData`), so changing UI text or icons means knowing which ids exist and what they hold.

**Arguments:** `path` (a `.res`, or a `.vbp` whose `ResFile32` is used; under the roots), `extract?` (`TYPE/name` keys such as `BITMAP/101`, `GROUP_ICON/MAINICON`, `CUSTOM/7`, or `["all"]`), `out_dir?` (a folder under the roots).

- **Result:** `resources` (`key`, `type` such as STRING, BITMAP, GROUP_ICON, VERSION or a custom name like CUSTOM, `name` or id, `language`, `bytes`), `strings` (`id` as `LoadResString` takes it, `language`, `text`; at most 500), `version` (`fileVersion`, `productVersion`, `strings` such as CompanyName, FileDescription).
- **Extracting:** bitmaps as `.bmp` (the file header added), icon and cursor groups as `.ico`/`.cur` (built from their images), anything else as raw data (its picture type if it is one, else `.bin`). Written to `out_dir`, never overwriting; without `out_dir` the first picture is shown. `all` takes everything except strings, version and the single icon/cursor images (their groups make the files).
- Only 32-bit `.res` files (they start with an empty 32-byte entry). Changing resources stays with VB's Resource Editor; the hook refuses Edit/Write on `.res`.

Verified on `AceScrollBarControl.res` (icon group and bitmap extracted as valid files), `VersionRipper.res` (version 1.9.0.204 and its strings) and the VB98 sample `AECLIENT.RES` (52 strings).

## `batch_run`: several calls in one

**Arguments:** `calls` (list of `{tool, args}`, at most 50), `continue_on_error?`.

Runs vb6 tool calls in order, in the server, through the same tool code as single calls, so every check still applies per call: ownership, design mode, roots, audit, `expected_hash`. It adds no add-in verb.
- **Why:** each MCP call is a model turn, and turns are what cost tokens and time. The free-hand Opus runs (benchmark T12) made 285–326 calls, many of them small chains like `app_click`, `app_wait`, `app_screenshot`.
- **Errors:** it stops at the first error unless `continue_on_error: true`. The whole list is checked before anything runs: unknown tool names, malformed entries and excluded tools refuse the batch.
- **A blocking dialog stops it anyway:** a call that fails with "answer it first (ide_dialog_respond)" ends the batch even with `continue_on_error`, since every later call would fail on the same dialog; the summary note says so (improvement notes 37).
- **Not allowed in a batch:** `ide_launch`, `ide_kill`, `build_make`, `build_in_ide` (long or disruptive: they stay single, visible calls), and `batch_run` itself.
- **Values from earlier calls** (PLAN.md E4): an argument whose whole value is `$<n>.<path>` is replaced by that field of call n's result (its first text item, parsed as JSON), keeping its JSON type: `"$1.hash"`, `"$2.added.name"`, `"$1.positions[0].left"`. Keys match exactly, then case-insensitively. A path starts with a letter, so text such as `"Price = $1.00"` is never taken for one, and only whole values are replaced, never part of a string.
  - A reference to the same or a later call refuses the batch before anything runs.
  - A reference to a failed call, a missing field (the error lists the result's fields) or a non-JSON result (an image) fails that call.
- **`compact: true`:** successful calls other than the last keep only the first 160 characters of their text; errors, the last call and images come in full.
- **Result:** each call's content in order, text items labelled `[n] tool:` (with `(error)` for a failed call), images unchanged, then a summary `{ran, failed, notRun, note}`. `isError` when any call failed.

Example with a reference: `{"calls": [{"tool": "code_read", "args": {"component": "modMath"}}, {"tool": "code_replace_lines", "args": {"component": "modMath", "start_line": 5, "count": 1, "new_text": "...", "expected_hash": "$1.hash"}}, {"tool": "project_save"}]}`

Example: `{"calls": [{"tool": "app_click", "args": {"text": "Apply"}}, {"tool": "app_wait", "args": {"window": "Results"}}, {"tool": "app_screenshot", "args": {"window": "Results"}}]}`

## Watchdog and recovery

The elevated watchdog (`vb6mcp watchdog`, started by the add-in; PLAN.md §3.1) covers the times when an IDE's add-in can't answer, because the debugged program loops or the IDE hangs. It is a sidecar. Without it everything else still works, and the tools below say when they need it. `ide_list` and `ide_status` report it as `watchdog: { pid, version, elevated }`, or `null` when it isn't running.

**Fallbacks in existing tools:**
- **`run_break`:** uses the IDE's Break command when the IDE responds. When the program keeps the IDE busy, the watchdog injects Ctrl+Break (spike 12) instead. It does that only with VB6, or Windows' stand-in for its unresponsive window, in front, and puts the previous window back afterwards. `method` says which was used.
- **`ide_dialog_respond`:** if the add-in doesn't answer, the watchdog clicks the button (`via: "watchdog"`).
- **`app_screenshot`:** if an elevated IDE can't run the add-in's capture (busy or hung), the watchdog captures it. `PrintWindow` gets a 3 s timeout, then falls back to a copy of the screen area, which shows only what is visible.
- **`ide_launch`:** with an elevated watchdog running, VB6 is started by the watchdog: elevated, with **no UAC prompt** (`via: "watchdog (no UAC prompt)"`, spike 12b). If the new VB6 shows a dialog before our add-in has loaded (another add-in's error at startup, e.g. Visual Component Manager), the result is `status: "dialog-before-addin"` with the new `pid` and the dialog.
- **`ide_dialog_respond` with the pid of a VB6 that has no add-in yet** (still starting) clicks through the watchdog (`via: "watchdog (no add-in yet)"`).

| Tool | Arguments | What it does |
|------|-----------|--------------|
| `ide_kill` | `force` (required true), `pid?` | ends a hung VB6 through the watchdog. The result lists the recovery copies of that IDE with their age. Killed in run or break mode, it also lists `comLeftovers`: `.vbp` files whose classes stay registered to the VB6 debugger until VB6 next starts (spike 25; repair with `com_orphans`) |
| `ide_quit` | `discard?`, `force?`, `pid?` | closes VB6 normally (PLAN.md C7): `WM_CLOSE` through the watchdog (`QUIT`; a non-elevated post is refused, spike 10). Only an IDE this session owns (`ide_launch` / `ide_claim`), otherwise `force: true`. Design mode. Refused while anything is unsaved, listing it, unless `discard: true`: each project with unsaved changes is first closed through `project_close discard` (files snapshotted and restored, spike 31). Waits up to 15 s for the process to exit, then drops the owner file. `status`: `quit`, `dialog` (another add-in or a prompt; not answered), `still-running`, `not-quit` (a discard didn't finish) |
| `recovery_list` | `from_pid?` | recovery sessions: pid, `running`, `crashed`, projects, and per component when it was copied |
| `recovery_read` | `from_pid`, `component`, `project?` | the copied text, and whether it differs from the file on disk |
| `recovery_restore` | `from_pid`, `component`, `project?`, `pid?` (target) | replaces the code of the component **with the same file** in an open project of a running IDE (design mode). The component is left unsaved |
| `recovery_discard` | `from_pid`, `confirm` (required true) | deletes a finished session's copies (refused while that VB6 is still running) |

**How recovery copies are taken:**
- Every 10 s the watchdog asks each IDE's add-in for its projects and components. It reads the code of the **unsaved** ones and writes copies to `%LOCALAPPDATA%\vb6mcp\recovery\<pid>\` when the text changed.
- It does this only while the user has been idle for 5 s (`GetLastInputInfo`) and the IDE is in design or break mode, so it never interrupts typing or a running program (hard rule 12).
- Copies of components that were saved are removed. A cleanly closed IDE's folder is removed.
- An IDE whose discovery file stayed behind (crash, kill) is marked `crashed`, and its copies are kept for 7 days.
- Only the code is copied, not designer properties: a form's layout comes from its file on disk.

After a crash: `ide_launch` the project, then `recovery_list`, `recovery_read` and `recovery_restore`. The user reviews the result and saves (`project_save`).

**Verified live** (spike 19), with no help from the user:
1. A busy loop was broken with `run_break` through the watchdog while Claude was in front: VB6 came forward, stopped at the `Loop` line, and Claude was restored.
2. An unsaved edit was copied about 10 s after it was made.
3. `ide_kill` listed the copy.
4. `ide_launch` started VB6 elevated without a UAC prompt. A Visual Component Manager startup box was answered through the watchdog.
5. `recovery_restore` brought back the identical module (same hash), unsaved.

---

## Change log

- 2026-10-19: registration without an elevated VB6: `com_register`/`com_unregister` run through the elevated watchdog, or show a UAC prompt with `elevate: true`, and refuse otherwise; `build_make` uses an elevated watchdog when the running VB6 isn't elevated, so the build registers.
- 2026-10-19: non-elevated VB6 tested (spike 41): `ide_dialog_respond` clicks itself when VB6 isn't elevated and no watchdog runs; `build_make` (`registered: false`) and `com_register` explain failed registration.
- 2026-10-19: heartbeat during slow calls; calls refused while `server_restart` ends the server (spike 42).
- 2026-10-19: `debug_eval allow_side_effect` trusts the hook only for Claude Code; other clients are asked through MCP elicitation, or refused (spike 40).
- 2026-10-18: IDE edits with `""` no longer fail after being applied; a failed edit says what state the module is in (improvement notes 29, 30).
- 2026-10-17: `ide_dialog_respond repeat` (click through one load error per form, then sum up the load logs).
- 2026-10-17: `ide_launch` reports load dialogs instead of failing while VB loads the project (spike 36).
- 2026-10-17: `res_read` (resource files: strings, version, bitmaps, icons).
- 2026-10-17: `form_resources` (read and extract what a .frx holds).
- 2026-10-17: the hook also refuses Edit/Write on binary companions (.frx .ctx .pgx .dsx .dox .res) and .vbw; server instructions name the companion files.
- 2026-10-17: `codeHints` for C-style escapes in VB string literals; Windows/CR LF rule in code parameters and server instructions.
- 2026-10-17: `form_layout` suggested after controls are moved by hand.
- 2026-10-17: the hook asks the user (spike 33) for flagged shell commands and for `debug_eval` with `allow_side_effect`.
- 2026-10-17: `debug_eval` refuses side effects by default (`VB6MCP_EVAL=open` allows them); the hook checks Bash/PowerShell commands.
- 2026-10-17: `markedChangedByRun` (spike 32) in `run_end`, `syntax_check`, `project_components`, `build_make`.
- 2026-10-17: crash notices (`ideCrashed`) with the next result; `ide_launch reopen`; the watchdog stays 30 minutes after a crash.
- 2026-10-17: `build_make` without an IDE, through the watchdog (verb `MAKE`).
- 2026-10-17: `ide_launch` reports `startSeconds` / `seconds` (time the UAC prompt was up).
- 2026-10-17: find/replace with preview in `code_replace_lines` and `file_replace_lines`.
- 2026-10-17: `project_open` preflight hints: unregistered files with the library, similar registered libraries.
- 2026-10-17: `ide_quit` (watchdog verb `QUIT`).
- 2026-10-17: `sync_diff` project summary (case-only re-casing).
- 2026-10-17: `build_in_ide`: `make-failed` with `causes` when the code compiles; empty projects refused.
- 2026-10-17: `symbol_refs`.
- 2026-10-17: `com_register` verdict: `notRegistered` / `registeredElsewhere`.
- 2026-10-17: `com_describe` marks default members.
- 2026-10-17: `server_restart` (spike E7).
- 2026-10-17: shorter results: `form_set_properties` lists failed changes only (`verbose`), `form_properties` `names`, `batch_run` `compact`, compact `app_tree`.
- 2026-10-18: `app_set_text` `index` counts text boxes in reading order; `app_tree` shows each box's `index` and `readOnly`; setting a Locked box is flagged (benchmark T18).
- 2026-10-18: step results carry `at` (where the step stopped); `bp_toggle` works while the program runs (spike 37).
- 2026-10-18: `app_select` without `control` picks the form's only list or combo box, or the only one with the item (benchmark T17).
- 2026-10-18: `debug_locals` reports variables of a private user-defined type as `(user-defined type)` instead of losing every value; `debug_locals` and `debug_eval` close the compile-error dialog their own Immediate line caused (benchmark T17).
- 2026-10-18: every tool names all missing required arguments at once, the names it doesn't take and the ones it does (e.g. `missing required arguments: component, new_text, expected_hash. Not arguments of proc_replace: code, file. proc_replace takes: ...`); `hwnd` counts as `control`; a control name in `control` gets an error that says where the handle comes from (benchmark T16).
- 2026-10-18: `com_describe` lists the methods of VB-built classes (their hidden default interfaces) without `include_hidden`.
- 2026-10-17: `symbol_refs` returns each file's `hash` (the `expected_hash` for `file_replace_lines`; benchmark T13).
- 2026-10-17: `batch_run` references to earlier results (`"$1.hash"`).
- 2026-10-19: `run_start` warns about missing DLL dependencies (`missingDlls`).
- 2026-10-19: live re-test fixes: `app_wait` stops at a dialog already open; a control with less than 4 pixels showing is `clipped: hidden`.
- 2026-10-19: `markedChangedByRun` names forms whose designer is open (`designerOpen`).
- 2026-10-19: `com_search` matches at word starts (`match`), `.oca` entries marked.
- 2026-10-19: `project_new`; `project_save` takes `path` for a never-saved project (add-in verbs `PRJNEW`, `PRJSAVEAS`, spike 43).
- 2026-10-19: `form_add_control` without a position counts overlaps instead of listing them.
- 2026-10-19: `com_describe include_hidden` leaves out the IUnknown/IDispatch methods.
- 2026-10-19: `ide_dialog_respond` presses once more when the first click didn't close the dialog (`retried`).
- 2026-10-17: `project_checkpoint`, `project_rollback`.
- 2026-10-17: `form_layout` (arrange controls as a grid).
- 2026-10-17: `batch_run` (several tool calls in one).
- 2026-10-17: IDE ownership: `ide_claim`, `ide_release`, `owner` in `ide_list`; write tools refuse another session's IDE; `ide_launch` claims what it starts.
- 2026-10-17: `project_close` discards through VB's save prompt (spike 31); `project_add_component` saves new components to their file; `project_save` takes `components`.
- 2026-10-17: `project_close`, `project_add_component`, `project_settings`; `build_make` takes `defines` (add-in verbs `CLOSE`, `COMPADD`, `PRJGET`, `PRJSET`, `SAVECOMP`).
- 2026-10-17: `file_replace_lines` `drop_var_help_ids`; attributes right after the last replaced line are kept with it.
- 2026-10-17: `syntax_check`, `build_in_ide` and `run_start` wait for pending sync reloads (`reloads`).
- 2026-10-17: `project_open` preflight (unregistered controls refuse, `force`), `opened-incomplete` with load logs, the add-in's error on a failed open.
- 2026-10-17: `build_make` load errors include the form's `.log` lines (`details`).
- 2026-10-17: IDE edits report names VB re-cased (`vbKeptSpelling`).
- 2026-10-17: `project_save` reports `.vbp` rewrites and reference/control version changes (`vbp`).
- 2026-10-17: `code_read` and `file_read` return at most 2,000 lines / 100,000 characters without a range (`truncated`, `nextStartLine`).
- 2026-10-17: `form_add_control` takes `properties`.
- 2026-10-17: `edits` (several ranges in one call) for `code_replace_lines` and `file_replace_lines`.
- 2026-10-16: designer results carry `hints` and `overlaps`; `Font.*`-only changes are made to count for Save (spike 29); `form_screenshot` stitches forms bigger than the designer.
- 2026-10-15: `form_set_properties` (several designer properties in one call).
- 2026-10-15: `form_screenshot` (designer image, add-in `DESIGN SHOW`).
- 2026-10-15: `com_orphans`, `com_diagnose`, `refs_fix_broken`; pictures and lists in `form_set_property`/`form_properties`; `refs_remove name=#<position>`; `ide_kill` warns about debugger registrations; `com_where` flags them.
- 2026-10-15: COM tools (`com_describe`, `com_search`, `com_where`, `compat_status`, `com_register`, `com_unregister`).
- 2026-10-14: form designer tools (`form_controls`, `form_properties`, `form_set_property`, `form_add_control`, `form_remove_control`) and references (`refs_list`, `refs_add`, `refs_remove`, `refs_add_component`).
- 2026-10-14: `ws_affected` (rebuild order for shared source, by typelib identity).
- 2026-10-14: `file_replace_lines` (closed projects; routed to the IDE when open); `hash` in `file_read`; Claude Code PreToolUse hook.
- 2026-10-14: `sync_status`, `sync_diff`, `sync_take_disk`, `sync_take_ide` (sync engine in the watchdog).
- 2026-10-14: `build_make`, `build_in_ide`.
- 2026-10-13: watchdog: `ide_kill`, `recovery_list/read/restore/discard`; watchdog fallbacks in `run_break`, `ide_dialog_respond`, `app_screenshot`, `ide_launch`; `watchdog` in `ide_list`/`ide_status`.
- 2026-10-13: `debug_locals`, `debug_callstack`, `debug_error`; `runtimeError` on reported dialogs.
- 2026-10-14: `debug_output` (OutputDebugString capture through DBWIN).
- 2026-10-14: `exe_run`, `exe_end`; `exe` argument on all `app_*` tools (compiled programs, input from the server).
- 2026-10-14: `app_select`, `app_keys`, `app_wait`; `app_mouse` real input deferred in the add-in (hard rule 14).
- 2026-10-13: running-program tools (`app_windows`, `app_tree`, `app_click`, `app_set_text`, `app_mouse`, `app_scroll`, `app_screenshot`); audit log for input, `debug_eval` and `ide_launch`.
- 2026-10-13: run and debug tools (`run_*`, `step_*`, `run_to_cursor`, `bp_*`, `debug_where`, `debug_eval`, `debug_immediate`, `syntax_check`).
- 2026-10-13: `ide_launch` (auto/elevated/normal); `elevated` in `ide_list` and `ide_status`.
- 2026-10-13: `symbol_find` (declarations across open IDE projects).
- 2026-10-13: workspace tools (`ws_projects`, `ws_file_projects`, `file_read`) over `VB6MCP_ROOTS`.

- 2026-10-13: editing tools (`code_replace_lines`, `code_insert`, `proc_replace`, `project_save`, `ide_dialog_respond`); `hash` in `code_read`/`proc_get`; `proc_get` no longer includes the blank lines above a procedure.

- 2026-10-13: `project_open` (deferred open, mode rules, dialog reporting).
- 2026-10-13: Phase 1 read tools (`ide_list`, `ide_status`, `project_components`, `code_read`, `symbol_outline`, `proc_get`).

**Large modules:** without `start_line`/`count`, at most 2,000 lines and 100,000 characters come back (a 185 KB class overflowed the client, improvement notes item 13).
- A cut result has `truncated: true`, `nextStartLine` and a `note`. Read on with `start_line`.
- `hash` always covers the whole text, so edits need no full read.
- An explicit range is returned as asked, up to the character cap.

