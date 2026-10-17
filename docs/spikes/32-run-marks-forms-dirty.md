# Spike 32: running the program marks forms as changed

**Date:** 2026-10-17. **Found by:** the T12 rerun (benchmark round 6). Both runs ended with four forms unsaved although the agents had saved, so the harness's `build_make` refused ("open in VB6 with unsaved changes") and the runs counted as failed. The second run of the earlier $20 round failed the same way.

## What happened in the runs

The agent's last `project_save` listed `frmDbSelect`, `frmLogin`, `frmTransferHistory`, `frmMonitor`, `frmHistory`, `frmMain` as saved. After it, the agent only ran the program and clicked through it (`run_start`, `app_*`, F12-Exit). Its final `project_components` showed `frmMain` `dirty: true`; it still reported "all edits are saved".

## Reproduction (VB6 pid 9832, the run's benchmark copy of a production project)

1. Nothing dirty (`project_components`). Copied the four forms' `.frm` and `.frx`.
2. `run_start`, `app_wait` for the main window, `run_end`.
3. `project_components`: exactly `frmDbSelect`, `frmLogin`, `frmMonitor`, `frmMain` are `dirty: true`, the forms the program loads at startup. No designer access is possible in run mode (error 393), and nothing was edited.
4. `project_save components=[those four]`, then compared with the copies: **all eight files byte-identical**.

So the flag is spurious: VB marks forms that were loaded while the program ran as changed, and saving them writes the same bytes. The runs that passed hadn't changed these shared forms or saved again after their last run; which property of a form triggers it wasn't narrowed down (not needed for the fix).

Also seen: `sync_diff` without a component (PLAN.md C4) listed 23 components that differ from disk only by VB's re-casing while VB doesn't count them as changed (`dirty: false`). Those are the identifiers VB re-cased project-wide (improvement notes 23); a save of any of them writes the new spelling.

## Decision

- Saving stays explicit (hard rule 4): the server doesn't clear the flag (which would save, spike 31).
- `run_start` remembers which components were dirty before the run. `run_end` reports the components the run marked changed whose code still equals disk (`markedChangedByRun`), explains it, and names the call that clears them: `project_save components=[...]` (writes only those files, identical bytes, not the `.vbp`).
- A build from disk (`build_make`) refusing because of them is correct behaviour, not a bug: the refusal now also points at this when it applies.
