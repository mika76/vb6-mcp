# Spikes 6 + 8: Compile errors from memory, Reload from disk, deferred execution retest

Date: 2026-10-13. Code: `spikes/02-addin-bridge/Probe.bas` (`QEXEC`, `QEXECPATH`, `FINDPATH`, `INFO`, `RELOAD`, `SETSEL`), `Net.bas` (`WM_DEFER`, `CreateProcess` launch), `spikes/12-watchdog` (`PRESS`). Scratch project: `tests/fixtures/scratch/` (Form1.frm backed up and restored).

## Deferred execution (hard rule 14) retest

- `QEXEC <id>` / `QEXECPATH "Menu|Item"` store the command and `PostMessage(WM_USER+101)` the add-in's own window. `WndProc` runs it from a fresh dispatch.
- **Start (deferred) → `Form_Load` hits `Stop` → normal break mode: Continue and End enabled** (`186 en=True &Continue`, `228 en=True &End`). **Deferred End → design mode.** The spike 3 deadlock didn't recur. → **Rule 14 confirmed as the fix.**
- Log order: `deferred: executing` → `Execute returned` → `EnterRunMode`. `Execute` returns immediately and the run starts afterwards.
- `FINDPATH "Run|Start With Full Compile"` → `id=746` (generic). Lookup by menu path + caption works for the generic-Id commands.

## Watchdog start without blocking

- `CreateProcess` (`CREATE_NO_WINDOW`) instead of VB `Shell`: **98 ms** (was ~9 s). The watchdog came up elevated, parent = VB6.

## Spike 6: compile errors from unsaved code

1. Insert `y = 1` (undeclared, `Option Explicit`) through the add-in. The code isn't saved.
2. Deferred `Start With Full Compile`.
3. The watchdog `STATUS` shows the dialog: **`'Microsoft Visual Basic'` · Static `"Compile error:\r\n\r\nVariable not defined"` · Buttons `OK`, `Help`**. Screenshot 206×156 via `PrintWindow` in 15 ms.
4. The add-in kept answering while the dialog was up.
5. Watchdog **`PRESS 7352 OK`** (posted `BM_CLICK`) closed it.
6. `GetSelection` = **`6,5 - 6,8`** (the offending `y`), `ProcOfLine` = `Form_Load`.

→ A structured compile result from **memory** is possible: `{component: Form1, line: 6, col: 5-8, proc: Form_Load, message: "Variable not defined"}`. The IDE stays in design mode.

Note: with the dialog open the IDE still **responds** (modal loop), so `STATUS` needs an explicit `dialog-open` state alongside responsiveness.

## Spike 8: `VBComponent.Reload()`

| Case | Result |
|------|--------|
| Clean component; **code** changed on disk | ✓ new line visible after `RELOAD`; `dirty=False` |
| Clean component; **designer** property changed on disk (`Form1.Caption`) | ✓ `caption=Changed on disk` after reload: designer + code both reloaded |
| **Dirty** component (unsaved IDE edit) + disk change | ⚠ **Reload silently discards the unsaved IDE edit**: no prompt, no error, `dirty` True→False, inserted line gone |
| Break mode | ✗ `error 5: Invalid procedure call or argument`: refused, nothing changed |

## Decisions

1. **Never call `Reload()` on a dirty component.** Before *any* reload, snapshot the IDE text (recovery store). The sync policy (PLAN §8) is mandatory, not a nicety.
2. Reload only in design mode (break mode refuses; queue until `[design]`).
3. `syntax_check` / compile-from-memory = deferred `Start With Full Compile` + watchdog dialog read + `PRESS OK` + `GetSelection`, then report. If the compile **succeeds**, the program starts running. The tool must then End it (deferred), unless the user asked to run.
4. Deferred execution is the standard path for every IDE command (rule 14).
5. Watchdog `PRESS` (posted `BM_CLICK`) answers IDE dialogs from outside.

## Open

- What a successful full compile looks like (run starts; Form_Load's MsgBox then appears) → the syntax-check tool needs a "compile only" strategy, e.g. End right after `EnterRunMode`.
- Compile errors in components that aren't open in a code pane: does VB open and select them? (Expected yes; verify.)
- `RequestWriteFile` veto (save-over-newer-disk-file) not tested.
