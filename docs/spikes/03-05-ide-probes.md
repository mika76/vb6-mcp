# Spikes 3 + 5: Live state, run/break access, IDE commands, Immediate window

Date: 2026-10-13. Code: `spikes/02-addin-bridge/` (`Probe.bas`, `Status.bas`, build-event hooks in `Connect.cls`; helpers `client.ps1`, `vbwin.ps1`). Scratch project: `tests/fixtures/scratch/` (Standard EXE, Form1 with a MsgBox in `Form_Load`).

## Questions

- Spike 3: which VBIDE members work in design, run and break mode? Do add-in writes join the IDE's undo history? Can code be read inside the build events?
- Spike 5: command-bar IDs for run/debug; break-mode detection; can the Immediate window be used to evaluate expressions?

## Results: VBIDE access by mode

| Access | Design | Run | Break |
|--------|--------|-----|-------|
| `VBProjects.Count`, `VBProjects(i)`, `VBComponents(i)` (indexed) | ✓ | ✓ | ✓ |
| `CodeModule.Lines/CountOfLines`, `Members`, `ProcOfLine`, `ProcStartLine/BodyLine/CountLines` | ✓ | ✓ | ✓ |
| `ActiveCodePane`, `GetSelection`, `SelectedVBComponent`, `IsDirty` | ✓ | ✓ | ✓ |
| **`For Each` over `VBProjects`** | ✓ | ✗ error 70 | ✗ error 70 |
| **`For Each` over `VBE.Windows`** | ✓ | – | ✗ error 70 |
| `For Each` over `CodeModule.Members` | ✓ | ✓ | ✓ |
| Designer: `VBComponent.Properties("Caption")`, `Designer.VBControls` | ✓ | ✗ error 393 "Get not supported at runtime" | ✓ |
| `VBE.Windows(i).HWnd` | **0 for every window** | | |

**Correction to spike 2:** VBIDE is **not** locked in run/break mode. Only **`For Each` enumeration** of some collections is refused. Indexed loops work.

## Results: writes and undo

- `CodeModule.InsertLines` works. Edit → Undo changes from "Can't Undo" to **"&Undo Insert Text"**, so **add-in writes join the IDE's undo history**.
- `LIST` shows the component as dirty afterwards.

## Results: build events

- `EnterRunMode` / `EnterDesignMode` fire. `BeginCompile` didn't fire on a plain Start.
- **Entering break mode fires `EnterDesignMode`.** Events alone can't tell break from design. The window title can (`[break]`, spike 2).
- **Reading code inside `EnterRunMode` made the debugged program end as soon as `Form_Load` returned.** With the handlers changed to log only, the program kept running (confirmed by hand). → **Build-event handlers must not touch VBIDE.**

## Results: command bars (`CMDBARS` dump: `spikes/02-addin-bridge/cmdbars-dump.txt`)

| Command | Id | | Command | Id |
|---|---|---|---|---|
| Start / Continue (caption changes) | 186 | | Step Into / Over / Out | 188 / 194 / 2559 |
| Break | 189 | | Run To Cursor | 1811 |
| End | 228 | | Toggle Breakpoint | 51 |
| Restart | 515 | | Clear All Breakpoints | 579 |
| Immediate Window | 2554 | | Set / Show Next Statement | 1812 / 1813 |
| Undo / Redo | 128 / 129 | | Copy / Paste / Select All | 19 / 22 / 756 |

- Many controls share the **generic Id 746** (e.g. "Start With Full Compile", "Make Project1.exe…"). Those must be found by menu path + caption.
- `Enabled` and `Caption` reflect state ("&Start" becomes "&Continue" in break mode).
- `EXEC 186` (Start) and `EXEC 189` (Break) worked. `EXEC 228` (End) failed with `Method '~' of object '~' failed` in break mode. The cause wasn't isolated; see the deadlock below.

## Results: Immediate window

- It's a custom-drawn `VbaWindow` captioned "Immediate", so `WM_GETTEXT` returns only the caption. Find it by enumerating child windows of the IDE (the non-elevated side can do this).
- **Evaluation works** in break mode, entirely through the add-in:
  1. `WM_CHAR` the text (`? 1+1`) to the Immediate `VbaWindow`
  2. `WM_KEYDOWN` VK_RETURN
  3. Select All (756) + Copy (19)
  4. read the clipboard
  - The clipboard then contained `? 1+1` / ` 2`.
- `Debug.Print` output appears in the same window (`dbg x=42`), so it's readable the same way.
- Caveat: this overwrites the user's clipboard. Production must save and restore it.

## Results: input from the add-in

- A space key posted to the MsgBox's OK button (from inside the IDE process) **closed the MsgBox**. In-process input works, as PLAN §3 assumed.

## Critical finding: deadlock when break mode happens under the add-in

Sequence:
1. IDE in break mode.
2. Code inserted (including a `Dim`).
3. `EXEC 186`.
4. A `STATE` request whose designer probe reads `VBComponent.Properties("Caption")`.

The request never returned (client timeout). The IDE showed `[break]` with the yellow arrow on `Stop` in `Form_Load`, and `dbg x=42` in the Immediate window. But **every** debug command was disabled (Continue, Break, End, Restart, also from the menu). Typing `End` in the Immediate window didn't help. **VB6 had to be killed in Task Manager.**

Explanation (consistent with all observations, not proven step by step): VB6 runs break mode as a nested message loop **on top of the running code's stack**. A request handler in the add-in (most likely the designer probe touching the form's properties at runtime, or a command executed inside the handler) caused user code (`Form_Load`) to run **inside the add-in's call stack**. It broke at `Stop` there, and VB6 refuses to continue or end while foreign code sits below the break point.

## Decisions

1. **Indexed loops only** (`For i = 1 To .Count`) for every VBIDE collection in the add-in. Never `For Each` over VBIDE collections.
2. **Build-event handlers: log or flag only, never touch VBIDE.**
3. **Never run anything that can execute user code synchronously inside the socket handler.** IDE commands (Start, Continue, Step, Break, End, Run To Cursor, Immediate input) are **deferred**: the handler validates, replies "queued", and `PostMessage`s its own hidden window. The action runs from a fresh top-level dispatch with no add-in frames below it. The MCP server then observes the outcome (title, dialogs, Immediate text).
4. **No designer/property access in run mode** (error 393 at best, form instantiation at worst).
5. **Mode** = window title (`[design]`/`[run]`/`[break]`). The events are only hints.
6. Debug commands by Id (table above). "Start With Full Compile" and "Make" by caption path.
7. **Immediate evaluation and `Debug.Print` capture** go through WM_CHAR + Enter + Select All/Copy + clipboard, with the clipboard saved and restored.
8. The hard-kill outcome confirms the **watchdog + recovery store** (PLAN §3.1). Unsaved edits would have been lost if this hadn't been a scratch project.
9. The **status toolbar** (version, DLL build time, port, request count) goes into the production add-in. It immediately revealed which build was loaded.

## Open

- Why `EXEC 228` (End) failed in a plain break.
- Whether deferred execution (decision 3) is enough to keep Continue/End available. Re-test in Phase 2.
- Run To Cursor / Toggle Breakpoint need `CodePane.SetSelection` first; not probed yet.
- Clipboard save/restore with non-text formats.
