# Changelog

## 1.2.0 (2026-10-20)

Two additions to the running-program tools. Only `vb6mcp.exe` changed: the add-in is the same build as in 1.1.0 and still reports 1.1.0.

**Upgrading:** run `install.ps1` from the new package. Restart a running `--stay` watchdog so it runs the new build too.

- `app_tree` lists the items of list and combo boxes: `itemCount`, the first 20 `items`, and `selected`. For a program in the IDE they're read through the add-in, because Windows blocks list messages from the non-elevated server. An event log kept in a ListBox used to show as `text: ""`.
- `app_screenshot` takes `region` (client coordinates) or `control` (a handle from `app_tree`) and returns only that part, at full resolution, with `crop` for mapping image positions back to `app_mouse` coordinates. Checking a few pixels no longer means reading the whole form.

## 1.1.2 (2026-10-20)

A fix for the shell guard. Only `vb6mcp.exe` changed: the add-in is the same build as in 1.1.0 and still reports 1.1.0.

**Upgrading:** run `install.ps1` from the new package. Restart a running `--stay` watchdog to get the fix. Claude Code runs the hook binary afresh on every call, so the new rules apply as soon as it is installed — but hook *configuration* is read when a session starts, so a session that opened before the hook was registered stays unguarded until it is restarted.

- Shell guard: two shapes an agent reaches for naturally got past `ShellPolicy`, both found by replaying a real session's commands through `vb6mcp.exe hook`.
  - A build whose executable sits in a variable and whose switch is an array element, as PowerShell usually writes it (`Start-Process -FilePath $exe -ArgumentList '/make','App.vbp'`, or `& $exe /make App.vbp`): the old rule wanted `vb6` and `/make` next to each other. A standalone `/make` now counts when the command also names a project file or VB6, so a `Start-Process` that merely mentions a `.vbp`, and a path that happens to contain "make", stay allowed.
  - `unix2dos` / `dos2unix` / `todos` / `fromdos` / `flip` on a VB6 source, which rewrite the file in place. A `.vbp` rewritten with LF leaves VB6 unable to parse it: the project loses its references and the build fails with a misleading "User-defined type not defined" in an unrelated `.bas`. Only at a command position, because "flip" is an ordinary word.

## 1.1.1 (2026-10-19)

A fix for the watchdog tray. Only `vb6mcp.exe` changed: the add-in is the same build as in 1.1.0 and still reports 1.1.0.

**Upgrading:** run `install.ps1` from the new package. Restart a running `--stay` watchdog to get the fix.

- Watchdog tray: Restart as administrator no longer leaves a console window open, and the log no longer calls the restart "stopped from the tray". The restart is now verified live.

## 1.1.0 (2026-10-19)

New projects, VB6 without administrator rights, and a watchdog you can see. One new tool (`project_new`, 95 in all). Details per tool: [docs/tools.md](docs/tools.md).

**Upgrading:** run `install.ps1` from the new package (roots are kept), then restart VB6 and Claude Code. The add-in reports version 1.1.0 and has two new verbs (`PRJNEW`, `PRJSAVEAS`); its ProgID, CLSIDs and the existing wire format are unchanged. A running `--stay` watchdog keeps the old version until you restart it.

### New
- **`project_new`** creates an EXE, ActiveX EXE, ActiveX DLL or ActiveX control project at a new `.vbp` path under the roots, also in an IDE with no project open. `project_save` takes `path` for a project that was never saved, such as the New Project chooser's `Project1` (spike 43). VB writes a new project's `.vbp` without its never-saved components, and still counts it as saved. So each component is saved first, and the server checks that the `.vbp` lists them all.
- **VB6 without elevation is supported and tested** (spike 41): everything works except registering components, and the README's "Elevated or not?" table says what changes.
- **Registration without an elevated VB6:** `com_register` / `com_unregister` run through an elevated VB6's add-in, else the elevated watchdog (no IDE needed), else with `elevate: true` through a UAC prompt you answer. Otherwise they refuse before running anything. `build_make` hands a build to an elevated watchdog when the running VB6 isn't elevated, so ActiveX builds still register.
- **Watchdog tray icon:** a "VB" tile in the notification area, with a shield when the watchdog runs as administrator. Its menu has Open log, Restart as administrator and Stop watchdog; `--no-tray` leaves the icon out.

### Fewer surprises
- **Slow vs dead:** during calls longer than 5 s the server sends a progress heartbeat every 10 s, and the installers set Claude Code's idle timeout to 2 minutes. A frozen or dead server now shows up as an error within 2 minutes instead of a 30-minute wait (spike 42). After `server_restart`, new calls are refused at once with "restarting".
- **A non-elevated VB6's startup dialogs** (another add-in's error before ours loads) are answered by the server itself.

### Other agents
- Tested in GitHub Copilot CLI and OpenCode (spike 40): all tools load. In clients other than Claude Code, `debug_eval`'s approval for side effects now asks through MCP elicitation, and is refused where the client can't ask. Before, such a client could run a side-effect expression with no prompt. The README lists what works where.

### Docs
- README: requirements as needed, tested on, and known untested limits (VB6 in another language, DBCS code pages); a screenshot of the tray icon.

## 1.0.0 (2026-10-19)

The first public release. Nothing new to learn compared with 0.2.0: the same 94 tools, with the bugs and surprises found in six more benchmark tasks (T13-T18: a change shared by a DLL and two EXEs, debugging a reported wrong total, binary compatibility, a cause hidden in a compiled DLL, breakpoints, stepping) and in real use fixed. Every fix of the last round was re-tested in a live IDE. Details per tool: [docs/tools.md](docs/tools.md); what's next: [docs/roadmap.md](docs/roadmap.md).

**Upgrading:** run `install.ps1` from the new package (roots are kept), then restart VB6 and Claude Code. The add-in reports version 1.0.0; its ProgID, CLSIDs and the wire protocol are unchanged.

### Safety
- `project_save` refuses code that no tool edited (stray keystrokes in a code window, someone else's edit) and shows the diff; `outside_edits: true` saves it anyway. `syntax_check` reports such edits too.
- `app_keys` never types into the IDE: without `control`, keys go to the program's focused control or its form, not to the Immediate window or a code pane.
- `app_tree` marks controls the user can't see (`clipped`: outside a container's visible area, or less than 4 pixels showing), and `app_click` refuses them unless `force: true`.
- `run_start` warns about DLLs the project's controls and references need that aren't installed (`missingDlls`). A missing delay-loaded DLL ends the IDE process the moment the program uses it, unsaved work included (spike 39). `com_diagnose` names the missing chain too.
- The hook also catches the agent ending VB6 through pipelines, `wmic`, CIM and `.Kill()`: a non-elevated process can end the elevated IDE (spike 38).

### Fewer surprises
- `server_restart` hands the session's IDEs and the runs it started to the next server, so `ide_quit` and `run_end` don't ask for `ide_claim` or `force` afterwards.
- Dialogs: `ide_launch` and `project_open` wait for the follow-up dialogs VB shows after loading and report them; `app_wait` stops as soon as a dialog blocks the program; `batch_run` stops at a dialog even with `continue_on_error`; `ide_dialog_respond` presses once more when the first click didn't take (`retried`).
- The program's own InputBox can be typed into (`app_set_text` / `app_keys` with `window` naming it).
- `app_click` says when nothing changed after the click (`changed: false`); lists of a program in the IDE are read through the add-in.
- `app_screenshot` returns `clientOrigin` and `imageScale`, so image positions convert to `app_mouse` coordinates.
- `app_keys`: a form with only Labels takes the keys itself; `control` may be a form; `real_input` types characters as their keys, so `KeyDown` sees the real `KeyCode`.
- Editing: an edit with `""` no longer reports an error after it was applied; a failed edit says what state the module is in; a fresh module's trailing line is addressable.
- `debug_locals` survives private user-defined types; `debug_locals` and `debug_eval` close the compile error their own line caused.
- Step results say where they stopped (`at`); `bp_toggle` works while the program runs.
- `app_set_text` `index` follows reading order; `app_tree` shows each text box's `index` and flags a Locked one (`readOnly`).
- `app_select` without `control` picks the only list or combo box (or the only one with the item).
- Missing required arguments are named all at once, with the right names.
- `markedChangedByRun` names listed forms whose designer is open (`designerOpen`).

### Shorter results
- `symbol_refs` returns each file's hash, so edits need no `file_read` first.
- `form_add_control` without a position counts overlaps instead of listing them.
- `com_search` matches at word starts (`match: substring` for the old way) and marks VB's `.oca` cache files.
- `com_describe` lists the methods of VB-built classes without `include_hidden`, and leaves out the IUnknown/IDispatch methods even with it.

### Development
- Benchmarks T13-T18 with their fixtures, generators, replays and scoring (`bench/`, `docs/benchmarks`).

## 0.2.0 (2026-10-17)

Everything learned from real use on two production projects and eight benchmark rounds of an agent redesigning a whole application (docs/benchmarks). 94 tools (0.1.0 had 78). Details per tool: [docs/tools.md](docs/tools.md).

**Upgrading:** run `install.ps1` from the new package (roots are kept), then restart VB6 and Claude Code. The add-in has new commands (closing projects, adding components, project settings, saving single components) and the watchdog too (quit, builds without an IDE); the installer replaces both, and they take effect once VB6 and the watchdog restart.

### Security
- `debug_eval` refuses expressions with side effects (database writes, `Shell` and other process starts, file, registry and network access, scripts, `SendKeys`). With the Claude Code hook installed, the agent can ask: `allow_side_effect: true` shows **you** a permission prompt. `VB6MCP_EVAL=open` (`-EvalOpen`) allows them without asking.
- The hook also checks the agent's own Bash/PowerShell commands and **asks you** before database writes (`sqlcmd`, `bcp in`, `Invoke-Sqlcmd`, ADO scripts), writes to VB6 files, COM registration or registry edits, and ending or building with VB6 directly. Headless runs can't answer, so those commands don't run. In benchmarks this stopped agents from changing a test database.
- The hook refuses Edit/Write on binary companions (`.frx`, `.ctx`, `.pgx`, `.dsx`, `.dox`, `.res`) and `.vbw`.
- IDE ownership: write tools refuse an IDE another Claude session started or claimed (`ide_claim`, `ide_release`).

### New tools
- **Forms:** `form_set_properties` (many properties in one call), `form_layout` (rows and columns, no twips by hand), `form_screenshot` (the designer as you see it, large forms stitched), `form_resources` (pictures in their original format, lists, texts and ActiveX state from a `.frx`).
- **Projects:** `project_close` (with `discard`), `project_add_component`, `project_settings` (incl. `CondComp`), `project_checkpoint` / `project_rollback`.
- **Code:** `symbol_refs` (every use of a name across the workspace, continued lines joined), find/replace with preview in `code_replace_lines` and `file_replace_lines`, several ranges per call (`edits`).
- **IDEs:** `ide_quit`, `ide_claim`, `ide_release`, `server_restart` (load a new server build without restarting Claude).
- **Files:** `res_read` (a project's resources: string table, version info, icons and bitmaps).
- **Batching:** `batch_run` (several tool calls in one, with references to earlier results).

### Fewer surprises
- Load errors: `ide_dialog_respond repeat: true` clicks through "Errors during load" form after form and says which controls are missing; `ide_launch` and `project_open` report load dialogs instead of failing; `project_open` refuses a project whose controls aren't registered (unless forced) and points at unregistered files or similar libraries.
- Running the program in the IDE marks the forms it loaded as changed although nothing changed; `run_end`, `syntax_check`, `project_components` and `build_make` now name them (`markedChangedByRun`) with the save that clears them.
- `project_save` reports `.vbp` rewrites (reference versions bumped by VB) and can save single components without the `.vbp`.
- Designer hints: overlapping controls, `Font.*` changes VB wouldn't save, button colours that don't show, and a `form_layout` suggestion after hand placement.
- Edits report names VB re-cased; `sync_diff` summarises a whole project (case-only differences); code hints for C-style escapes in VB strings.
- Crashes are reported with the next tool result, with the Windows error-log entry; the watchdog stays 30 minutes after a crash; `ide_launch reopen: true`.
- `build_make` works with no IDE open (through the watchdog), takes `defines`, includes the form `.log` on load errors; `build_in_ide` explains a failed make when the code compiles.
- Shorter results where they cost the most tokens (measured): failed changes only, `names` filters, compact `app_tree`, `batch_run compact`.
- `com_describe` marks default members; `com_register` no longer fails for classes a DLL never registers.

### Setup
- `install-server.ps1` installs into a fixed folder, so `server_restart` picks up new builds; it keeps your roots.
- `install.ps1` logs the server to `%LOCALAPPDATA%\vb6mcp\server.log`; README lists every log and every VB6 file type.
- The watchdog can be started by the user (`vb6mcp watchdog --stay`): sessions then launch VB6 without UAC prompts.
- The watchdog and the add-in report the release version (both said 0.1.0 before).

## 0.1.0 (2026-10-15)

First package: live code reading and editing in the IDE, run and debug, driving the running program, builds, sync with disk, form designer, references and COM, the elevated watchdog, the installer.
