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
    { "pid": 8684, "mode": "design", "responds": true,
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

## `project_open`

Open a `.vbp` or `.vbg` in a **running** IDE.

**Arguments:** `path` (full path), `mode?` (`add` or `replace`), `pid?`.

**Rules:**
- If the IDE has no project, `mode` isn't needed.
- If it has projects, `mode` is **required** (hard rule 11, don't hijack the user's session). `add` adds to a group; `replace` closes the open projects and is **refused while any of them has unsaved changes**, because VB6 would ask to save and nobody would answer.
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

## Change log

- 2026-10-13: editing tools (`code_replace_lines`, `code_insert`, `proc_replace`, `project_save`, `ide_dialog_respond`); `hash` in `code_read`/`proc_get`; `proc_get` no longer includes the blank lines above a procedure.

- 2026-10-13: `project_open` (deferred open, mode rules, dialog reporting).
- 2026-10-13: Phase 1 read tools (`ide_list`, `ide_status`, `project_components`, `code_read`, `symbol_outline`, `proc_get`).
