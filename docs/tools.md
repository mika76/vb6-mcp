# MCP tool reference

Tools exposed by `vb6mcp` to Claude Code. Every result is a single JSON text item unless noted. Failures are MCP tool results with `isError: true` and a message meant for the agent (e.g. *"several VB6 IDEs are running (8684, 15652); pass pid"*).

**Implementation:** `src/Vb6Mcp/Tools/IdeTools.cs`. **Wire verbs used:** [protocol.md](protocol.md).

## Common arguments

| Argument | Type | Meaning |
|----------|------|---------|
| `pid` | integer | VB6 process id. **Optional** when exactly one IDE with the add-in is running. |
| `project` | string | `VBProject.Name`. **Optional** when that IDE has exactly one project. |
| `component` | string | `VBComponent.Name` (e.g. `Form1`, `Module1`). |

**Source of truth** (AGENTS.md hard rule 4): code tools read the IDE's memory, never the file on disk. Results carry `source: "ide"` and `dirty` (unsaved changes).

**Modes:** reads work in **design, run and break** mode (spike 3: indexed VBIDE access). While the debugged program runs a loop, the add-in can't answer; tools then fail with *"VB6 (pid …) didn't answer … the IDE is busy or blocked"* after 5 s (hard rule 6). Use `ide_status` to see why.

---

## `ide_list`

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

## `ide_status`

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

## `project_components`

**Arguments:** `pid?`, `project?`.

**Served by:** add-in `COMPONENTS`.

**Result:**
```json
{ "project": "Project1",
  "components": [ { "Name": "Form1", "type": "Form", "dirty": true, "lines": 5, "file": "D:\\...\\Form1.frm" } ] }
```
- `type`: `StdModule`, `ClassModule`, `Form`, `MDIForm`, `UserControl`, `PropertyPage`, `ResFile`, … (VBIDE `vbext_ComponentType`).
- `lines` is `-1` for components without a code module; `file` is empty if never saved.

## `code_read`

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

## `symbol_outline`

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

## `symbol_find`

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

## `ide_launch`

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
- A declined UAC prompt is an error; VB6 isn't started.
- VB6.EXE is found via `VB6MCP_VB6_EXE`, the installer's `ProductDir` in the registry, the `.vbp` file association, or the default folder.

**Result:**
```json
{ "status": "started", "pid": 16724, "elevated": true, "elevation": "auto",
  "exe": "C:\\Program Files (x86)\\Microsoft Visual Studio\\VB98\\VB6.EXE",
  "mode": "design", "title": "Basic - Microsoft Visual Basic [design]", "projects": ["Basic"], "dialogs": [], "note": null }
```
`status: "no-addin"` means VB6 started but the add-in didn't connect (not loaded on startup).

## `project_open`

Open a `.vbp` or `.vbg` in a **running** IDE.

**Arguments:** `path` (full path), `mode?` (`add` or `replace`), `pid?`.

**Rules:**
- If the IDE has no project, `mode` isn't needed.
- If it has projects, `mode` is **required** (hard rule 11, don't hijack the user's session). `add` adds to a group; `replace` closes the open projects and is **refused while any of them has unsaved changes**, because VB6 would ask to save and nobody would answer. The add-in first removes the open projects, then opens the new one with `AddFromFile(path, True)` into the empty IDE.
  - With projects still open, `AddFromFile(path, True)` wrote a `Group1.vbg` of the closed projects into VB6's current folder.
  - `AddFromFile(path, False)` makes even a single project a project group, and then Save asks for a group file name.
  - So only `add` next to other projects uses `False`.
- If the path is already open, the result is `already-open` and nothing happens.
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

## `proc_get`

Read one procedure or property, with the comment lines above it.

**Arguments:** `component`, `name`, `kind?` (`proc` default, `get`, `let`, `set`), `project?`, `pid?`.

**Served by:** add-in `PROC`.

**Result:**
```json
{ "source": "ide", "project": "Project1", "component": "Form1", "name": "Form_Load", "kind": "proc",
  "startLine": 2, "bodyLine": 3, "lines": 3, "text": "Private Sub Form_Load()\r\n    MsgBox \"hi\"\r\nEnd Sub" }
```

---

## Editing

Rules for all editing tools:
- **Optimistic concurrency:** `code_read` and `proc_get` return `hash`, a hash of the **whole module** (first 16 hex characters of SHA-256 over its 1252 bytes). Edits require it as `expected_hash`. The server re-reads the module and refuses on mismatch (*"conflict: … changed since you read it"*); the add-in additionally checks that the target lines are still exactly what was read (protocol `REPLACE`, 409). So a user typing in the IDE is never overwritten: re-read and redo.
- **Design mode only**, and no dialog open. Edits in break mode contributed to the spike 3 deadlock.
- **Validated before sending:** line endings normalized to CR LF, one trailing line break dropped, at most 1023 characters per line, at most 24 line continuations, only characters that exist in Windows-1252 (hard rules 1, 10).
- Edits are **unsaved** in the IDE and **undoable** there (Edit > Undo). Nothing is saved automatically (hard rule 4).
- Every edit returns the new `hash`, so the next edit can follow without a re-read.

### `code_replace_lines`

**Arguments:** `component`, `start_line`, `count` (0 = insert), `new_text` (empty = delete), `expected_hash`, `project?`, `pid?`.

**Served by:** `READ` (hash check) + `REPLACE` + `READ`.

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

### `project_save`

**Arguments:** `project?`, `pid?`.

**Served by:** `SAVE` (deferred File > Save Project), then polling `COMPONENTS` and dialogs for up to 20 s.

**Result** (`status`): `saved` (with the components that were unsaved), `nothing-to-save`, `dialog` (reported, not answered, e.g. a source-control checkout prompt) or `pending`. Refused for a never-saved project (Save would open a Save As dialog).

### `ide_dialog_respond`

**Arguments:** `button` (caption, e.g. `OK`, `No`), `dialog?` (caption; needed when several dialogs are open), `pid?`.

**Served by:** window reads (finds the dialog and the button's window handle) + add-in `CLICK`, then a re-read to verify.

**Result:**
```json
{ "clicked": "No", "dialog": "Source Code Control", "text": "Add this project to SourceSafe?",
  "dialogClosed": true, "dialogsNow": [] }
```
The button must always be named explicitly (hard rule 9). An unknown button lists the available ones.

---

## Run and debug

VB6 has no debugger API. These tools run IDE commands (deferred in the add-in, protocol `CMD`) and then **observe the outcome**: the mode from the window title, dialogs, the current line. A sent command isn't proof that it worked (spike 10). Each result carries `status`, `mode`, `responds`, `title`, `dialogs` and a `note` when something needs attention (e.g. a dialog to answer with `ide_dialog_respond`, or an IDE that stopped responding because the program is in a loop).

| Tool | Arguments | Needs mode | What it does |
|------|-----------|------------|--------------|
| `run_start` | `full_compile?`, `pid?` | design | F5, or Start With Full Compile (Ctrl+F5) |
| `run_break` | `pid?` | run | Break. If the program is in a busy loop the IDE can't take the command (the add-in shares its thread); the watchdog will inject Ctrl+Break (Phase 2b) |
| `run_continue` | `pid?` | break | Continue |
| `run_end` | `force?`, `pid?` | run, break | End. **Needs `force: true` if the agent didn't start the run** (hard rule 11) |
| `step_into` / `step_over` / `step_out` | `pid?` | break (`step_into` also design) | F8 / Shift+F8 / Ctrl+Shift+F8 |
| `run_to_cursor` | `component`, `line`, `project?`, `pid?` | break, design | sets the cursor, then Ctrl+F8 |
| `bp_toggle` | `component`, `line`, `project?`, `pid?` | design, break | sets the cursor, then F9. Returns the breakpoints this server knows; VBIDE can't list breakpoints, so ones set by hand are unknown |
| `bp_clear_all` | `pid?` | design, break | Ctrl+Shift+F9 (all breakpoints, including the user's) |
| `debug_where` | `pid?` | any | active code pane: component, line, column, procedure, line text (in break mode: the current statement) |
| `debug_eval` | `expression`, `pid?` | break, design | types `? expression` into the Immediate window and returns the new output |
| `debug_immediate` | `last?`, `pid?` | any | the Immediate window's text (`Debug.Print` output) |
| `syntax_check` | `pid?` | design | compiles the code **as it is in the IDE** (unsaved edits included) |

**Verified live** (Basic fixture, 2026-10-13): `syntax_check` (compiles, start + end), `bp_toggle` in `AddNumbers`, `run_start`, the user clicked Add, the IDE stopped at the breakpoint, `debug_where` (line 8, the current statement), `debug_eval` `a` = 1, `b` = 0, `g_callCount` = 0, `step_over` (line 9), `bp_clear_all`, `run_continue`, `run_end`; in design mode `debug_eval` of `PI_APPROX * 2` (6.28318) and `AddNumbers(40, 2)` (42), with the Immediate window's existing content and the user's clipboard text preserved.

**`debug_eval`** result: `{ "expression": "x", "status": "ok", "result": "42" }`. In design mode only module-level and global names are in scope, and calling a function runs it. A dialog that appears while evaluating (e.g. a runtime error) is reported (`status: "dialog"`), not answered. Reading the Immediate window uses the clipboard; the user's clipboard **text** is restored.

**`syntax_check`** results:
```json
{ "compiles": false, "message": "Variable not defined",
  "location": { "Project": "Basic", "Component": "modMath", "line": 8, "column": 5, "endColumn": 6, "procedure": "AddNumbers", "lineText": "    y = 1" } }
{ "compiles": true, "status": "ended", "mode": "design" }
{ "compiles": true, "status": "running-with-dialog", "note": "...", "dialogs": [ ... ] }
```
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
- **Types:** one Immediate line of `TypeName(...)` for all names. `TypeName` never raises an error.
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

## The running program (Phase 3)

Drive the program being debugged, the way a user would. The program runs **inside `VB6.EXE`**, so its windows belong to the IDE process.

- **Reads** (windows, controls, text) are done by the server with Win32 calls; they work across the elevation boundary (spike 10).
- **Input** goes through the add-in, because a non-elevated process's input to an elevated window is silently dropped (spike 10). It is limited to windows of the IDE's own process, and every input is written to the audit log `%LOCALAPPDATA%\vb6mcp\audit.log` (hard rule 8). So are `debug_eval` and `ide_launch`.
- Every input tool **verifies afterwards**: the IDE mode, dialogs and windows, and for text the value read back.

| Tool | Arguments | Needs mode | What it does |
|------|-----------|------------|--------------|
| `app_windows` | `pid?` | any | the program's top-level windows: `hwnd` (hex), `kind` (Form, MDIForm, Dialog), caption, class, position, size, enabled |
| `app_tree` | `window?`, `pid?` | run, break | the controls of a form: `hwnd`, `kind` (CommandButton, TextBox, CheckBox, ...), text, position relative to the form's client area, visible, enabled, depth |
| `app_click` | `control?` or `text?`, `window?`, `pid?` | run | clicks a command button, check box or option button. `text` matches the caption without `&`, case-insensitive |
| `app_set_text` | `text`, `control?` or `index?`, `window?`, `pid?` | run, no dialog | sets a text box's text (fires `Change`); returns `textNow` and `applied` |
| `app_mouse` | `action` (click, double_click, right_click, drag), `x`, `y`, `to_x?`, `to_y?`, `real_input?`, `window?`, `pid?` | run, no dialog | mouse input at a point relative to the form's client area. By default it posts messages to the deepest window under the point without touching the user's mouse, and a drag goes entirely to the window it starts in. **Labels and Images need `real_input: true`** (spike 17) |
| `app_scroll` | `direction` (up, down, left, right), `amount?` (line, page, end), `count?` (1), `method?` (scrollbar, wheel), `control?`, `window?`, `pid?` | run, no dialog | scrolls a list box, text box, grid, scroll bar control or the form; returns `before`/`after` scroll positions and `moved` |
| `app_select` | `control`, `text?` or `index?`, `pid?` | run, no dialog | selects an item in a list or combo box and raises its `Click`; returns `selectedNow`/`applied` |
| `app_keys` | `keys`, `control?`, `real_input?`, `pid?` | run, no dialog | types text and keys (syntax below) into the control or the focused control |
| `app_wait` | `window`+`gone?` / `control`+`text`/`contains` / `mode` / `dialog`, `timeout?` (10, max 120), `pid?` | any | polls (every 200 ms) until the condition holds; returns `met`, `seen`, `waitedMs` |
| `app_screenshot` | `window?` (or `ide`), `max_width?` (1024), `pid?` | any | an image of the window, plus `{ window, hwnd, width, height, method }` |

`window` is a hex handle from `app_windows` or a caption. Without it, the tools use the program's first form.

**Compiled programs:** every `app_*` tool also takes `exe`, the pid returned by `exe_run`, to act on a compiled EXE instead of the program in the IDE.
- Such a program runs at the server's integrity, so the server reads it and sends input itself: posted messages, or `SendInput` for `real_input`. Screenshots are taken by the server.
- Input only goes to programs the agent started with `exe_run` (hard rule 8).
- Posted keys can't be deferred inside the program's thread from outside. The server attaches to the program's input thread to set modifiers, and pauses 60 ms after each key down so the generated character arrives first.
- `mode` in `app_wait` is IDE-only.

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

By default the keys are posted, and modifiers are set in the program's key state between the messages (protocol `KEYS`). `real_input: true` sends real keystrokes instead: it brings the form to the front for a moment and restores the window that was in front. The result has `textNow` for text and combo boxes.

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
- With `window: "ide"` it captures the whole IDE main window. Crop from that for IDE panes: `PrintWindow` on child panes is misaligned (spike 15).

Errors: `the IDE is in design mode; the program must be running` (input tools), `no window 'X'. Open: ...`, `no button ... captioned 'X'`, `N controls are captioned 'X'; pass control`, `'X' is disabled`, `control 1A2B is a CommandButton, not a text box`, and the add-in's 400 when the handle isn't a window of the IDE process.

---

## Watchdog and recovery (Phase 2b)

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

## Builds (Phase 4)

VB6 `/make` refuses a project that is open in any IDE (spike 21). So the tools split by whether the project is open:

| Tool | Arguments | What it does |
|------|-----------|--------------|
| `build_make` | `path?` (`.vbp`), `project?`, `pid?` | **Closed project** (`path` under `VB6MCP_ROOTS`): `VB6.EXE /make` from the files on disk, started elevated by an add-in, so an ActiveX DLL/OCX is registered on success. Returns `status` (`succeeded`, `failed`, `dialog`, `timeout`), `errors` (file, line, message), the raw `log`, and the output file before/after. **Open project:** if saved, built in its IDE (File > Make; `via` says so); with unsaved changes, refused |
| `build_in_ide` | `project?`, `pid?` | File > Make from IDE memory (`MakeCompiledFile`), unsaved edits included; design mode. `status`: `succeeded`, `failed` (with `compileCheck`: message and location from a full compile, because `MakeCompiledFile` itself only fails generically), `dialog`, `timeout`. An existing output file is replaced without a prompt |

Notes from spike 21:
- **Error lines:** `/make` counts the IDE's code lines from 0. `errors[].line` is the 1-based IDE line (as in `code_read`) and `fileLine` the line in the file on disk.
- **Binary compatibility break:** `/make` fails with a `kind: "compatibility"` error naming the class and member. Nothing is written or registered.
- **In the IDE:** the same break is a decision dialog. `build_in_ide` returns `status: "compatibility-decision"` and leaves it for the user (hard rule 13).

**Verified live** (2026-10-14): closed EXE and ActiveX DLL builds (the DLL registered and callable), compile errors from `/make` and from memory with location, and a binary compatibility break both ways.

---
## Sync: files changed outside the IDE (Phase 4)

The watchdog's sync engine checks the files of every loaded project every 2 seconds (PLAN.md §8):
- **How it knows:** a file whose time stamp or size changed without the IDE writing it (the add-in records the IDE's own saves) was changed outside: git, another editor, or another IDE that has the same shared file. The change must look the same on the next check before anything happens, so half-written files are left alone.
- **Clean component, design mode:** reloaded automatically (the add-in checks again right before).
- **Run or break mode:** reloaded once the IDE is back in design mode.
- **Component with unsaved IDE changes:** a **conflict**. Nothing is touched; it is resolved with the tools below.
- **`.vbp` changed, file deleted:** reported only. Reloading a project is disruptive and left to the user.

| Tool | Arguments | What it does |
|------|-----------|--------------|
| `sync_status` | `pid?` | conflicts (`pid`, `project`, `component`, `file`, `detected`) and the last 30 events (`reloaded`, `pending`, `conflict`, `resolved`, `not-reloaded`, `project-changed`, `deleted`) |
| `sync_diff` | `component`, `project?`, `pid?` | unified diff of the IDE text (unsaved changes included) against the file on disk (`--- ide`, `+++ disk`); `same` when equal |
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

## Form designer (Phase 5)

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

**Arguments:** `component`, `control`, `project?`, `pid?`.

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

### `form_add_control`

**Arguments:** `component`, `prog_id`, `name?`, `container?`, `left?`, `top?`, `width?`, `height?`, `project?`, `pid?`.

`prog_id` is the class name **as a `.frm` writes it**: `VB.CommandButton`, `VB.Frame`, `MSComctlLib.ListView`.
- It is not the registered ProgID (`MSComctlLib.ListViewCtrl.2`): with that, VB's `Add` returns nothing (spike 23).
- An ActiveX control's component must be in the project first (`refs_add_component`).
- `container` puts it inside a Frame or PictureBox.

Result: `added`, the new control as in `form_controls`. When saved, VB writes the OCX's `Object =` line into the `.frm` header and the `.vbp` itself, consistently (verified).

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

## References and components (Phase 5)

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

## COM components (Phase 5)

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
- `Enum`s with values, and `Type`s.
- Help strings as comments.

Without a filter, a listing over 60,000 characters returns the type list (name, kind, member count) instead, so the agent can then ask for one type.

```json
{ "typeLib": { "libId": "{420B2830-E718-11CF-893D-00A0C9054228}", "version": "1.0", "name": "Scripting", "description": "Microsoft Scripting Runtime", "file": "C:\\Windows\\SysWOW64\\scrrun.dll" },
  "declarations": "' Scripting {420B2830-…} v1.0 - Microsoft Scripting Runtime\r\n\r\ndispinterface IDictionary ...\r\n    Sub Add(Key As Variant, Item As Variant)  ' Add a new key and item to the dictionary.\r\n..." }
```

The same text is available on the command line: `vb6mcp describe <file> [filter]`.

### `com_search`

**Arguments:** `text`.
- **Typelibs:** registered ones whose description or file path contains the text (libId, version, description, file, fileExists).
- **ProgIDs:** registered ones containing it (progId, clsid, server file, fileExists).

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

**Arguments:** `file` (a `.dll`/`.ocx` under the workspace roots), `confirm: true` (required), `force?` (unregister only), `pid?` (which IDE's add-in runs it).

**How it runs:**
- `regsvr32 /s` (32-bit) runs elevated through the add-in; no UAC prompt.
- The add-in checks the path again on its side.
- Each call goes to the audit log.

**Result:**
- `exitCode`.
- `before` and `after`: the file's typelib and classes, with where the registry points.
- `verified`: after registering, everything points at this file; after unregistering, nothing does.

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
- **When a workspace project builds the file:** its `compat_status` findings, which matter for 430.

**Verified live:** 429 on `ComFixture.clsGreeter` after an IDE kill reported "registered to the VB6 debugger" and "typelib registered to a project file, left over". On Form1x it found its unregistered ComFixture reference.

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

Edits the code of a VB6 source file **of a closed project** on disk. Arguments: `path`, `start_line` (1-based, the code as the IDE shows it and as `file_read` returns it), `count` (0 inserts before `start_line`), `new_text`, `expected_hash` (from `file_read`), and `pid?` when the file is open in several IDEs.

- **On disk** (the file isn't open in any IDE):
  - Only the code lines change; header, designer block and everything else stay byte-exact (hard rule 2), and the file stays Windows-1252 with CR LF.
  - The new text is validated: 1252 only, 1023 characters per line, 24 continuations.
  - Hidden inline `Attribute` lines (e.g. `VB_UserMemId` after a default property's signature) are put back after the same line when the new text keeps it; otherwise the edit is refused.
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

## Change log

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
