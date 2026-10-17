# Changelog

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
