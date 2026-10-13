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
- If it has projects, `mode` is **required** (hard rule 11, don't hijack the user's session). `add` adds to a group; `replace` closes the open projects and is **refused while any of them has unsaved changes**, because VB6 would ask to save and nobody would answer. The add-in removes the open projects and then adds the new one; it doesn't use `AddFromFile(path, True)`, which wrote a `Group1.vbg` into VB6's current folder.
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
| `app_screenshot` | `window?` (or `ide`), `max_width?` (1024), `pid?` | any | an image of the window, plus `{ window, hwnd, width, height, method }` |

`window` is a hex handle from `app_windows` or a caption. Without it, the tools use the program's first form.

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
- Not yet verified: VB drag-and-drop (`DragMode` automatic, OLE drag).

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
| `ide_kill` | `force` (required true), `pid?` | ends a hung VB6 through the watchdog. The result lists the recovery copies of that IDE with their age |
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

- 2026-10-13: watchdog: `ide_kill`, `recovery_list/read/restore/discard`; watchdog fallbacks in `run_break`, `ide_dialog_respond`, `app_screenshot`, `ide_launch`; `watchdog` in `ide_list`/`ide_status`.
- 2026-10-13: `debug_locals`, `debug_callstack`, `debug_error`; `runtimeError` on reported dialogs.
- 2026-10-13: running-program tools (`app_windows`, `app_tree`, `app_click`, `app_set_text`, `app_mouse`, `app_scroll`, `app_screenshot`); audit log for input, `debug_eval` and `ide_launch`.
- 2026-10-13: run and debug tools (`run_*`, `step_*`, `run_to_cursor`, `bp_*`, `debug_where`, `debug_eval`, `debug_immediate`, `syntax_check`).
- 2026-10-13: `ide_launch` (auto/elevated/normal); `elevated` in `ide_list` and `ide_status`.
- 2026-10-13: `symbol_find` (declarations across open IDE projects).
- 2026-10-13: workspace tools (`ws_projects`, `ws_file_projects`, `file_read`) over `VB6MCP_ROOTS`.

- 2026-10-13: editing tools (`code_replace_lines`, `code_insert`, `proc_replace`, `project_save`, `ide_dialog_respond`); `hash` in `code_read`/`proc_get`; `proc_get` no longer includes the blank lines above a procedure.

- 2026-10-13: `project_open` (deferred open, mode rules, dialog reporting).
- 2026-10-13: Phase 1 read tools (`ide_list`, `ide_status`, `project_components`, `code_read`, `symbol_outline`, `proc_get`).
