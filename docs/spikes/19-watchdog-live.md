# Spike 19: the watchdog, live (Phase 2b)

Date: 2026-10-13. Elevated VB6 SP6, `tests/fixtures/Debug` (Loop button: `Do: i = i + 1: Loop` with no `DoEvents`), production watchdog (`vb6mcp watchdog`) started by the add-in. Claude Code was the foreground app unless noted.

## Results

| Step | Result |
|------|--------|
| Add-in starts the watchdog on load | ✓ elevated, parent = VB6, `ide_list` shows `watchdog: { elevated: true }` |
| Loop clicked (`app_click`) | ✓ IDE `responds: false`, add-in busy, `ide_status` says to use `run_break` |
| `run_break`, attempt 1: the program's form was in front | ✗ **watchdog blocked**: VB6 was already in front, but the code attached to the foreground thread (VB6's busy thread) and called `SetForegroundWindow`, which waited on that thread forever. The request timed out; the log ended at `BREAK 16724` |
| attempt 2 (fix: VB6 in front → change nothing; never attach to a VB6 thread), Claude in front | ✗ refused correctly: `SetForegroundWindow` didn't move the unresponsive VB6 forward, so Ctrl+Break wasn't sent (it would have gone to Claude) |
| attempt 3 (fix: also `SwitchToThisWindow`, then Alt + `SetForegroundWindow`) | ✓ VB6 came to the front, Ctrl+Break was injected, the IDE stopped at the `Loop` line in `cmdLoop_Click`, and Claude was back in front afterwards |
| Unsaved edit (`proc_replace` in modCalc) | ✓ copied by the watchdog about 10 s later (idle user, design mode) |
| `recovery_list` / `recovery_read` | ✓ session running, 1 copy; text differs from disk (30 vs 28 lines) |
| `ide_kill force` | ✓ VB6 killed; the result listed the copy and its age |
| `ide_launch` (through the watchdog) | ✓ VB6 started **elevated, no UAC prompt**, but it reported `no-addin`: a **Visual Component Manager** box ("Method '~' of object '~' failed", another add-in, a known VB6 quirk) held up startup, and add-ins load one after the other |
| that box's OK pressed through the watchdog (`PRESS`) | ✓ startup finished, our add-in loaded, DebugFixture open and unchanged |
| `recovery_restore` from the killed session | ✓ modCalc restored; module hash identical to before the kill (`475f8f90712ba1b6`), left unsaved |

## Findings

1. **Never attach input to the busy VB6 thread.** `AttachThreadInput` to a thread that doesn't pump messages, followed by `SetForegroundWindow`, blocks the caller. When the debugged program's form is in front, that thread *is* the foreground thread. Foreground changes now skip VB6 threads and run on a separate thread with a timeout.
2. **An unresponsive window doesn't take the foreground from `SetForegroundWindow`**, even with the input attached to the foreground thread. `SwitchToThisWindow` or a synthetic Alt press followed by `SetForegroundWindow` works. That answers spike 12's open question: VB6 doesn't need to be in front already, because the watchdog can bring it there.
3. **Startup dialogs from other add-ins delay ours.** `ide_launch` now watches the new VB6 process and reports such a dialog with its pid (`status: "dialog-before-addin"`). `ide_dialog_respond` accepts that pid and clicks through the watchdog, because there is no add-in to ask yet.
4. The full crash path works with no help from the user: kill, relaunch without UAC, restore unsaved code by file identity.
