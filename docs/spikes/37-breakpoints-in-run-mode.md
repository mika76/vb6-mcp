# Spike 37: breakpoints while the program runs (2026-10-18)

## Question

`bp_toggle` and `bp_clear_all` refused run mode ("needs design or break mode"). While preparing benchmark T18, a breakpoint was needed after the program had already started. The detour was `run_break`, `bp_toggle`, `run_continue`. Does VB6 take the commands while the program runs?

## What was tried

The Order fixture (`tests\fixtures\Order`), running in the IDE (pid 16004) with the add-in loaded, using a server build that allowed run mode:

1. `run_start`, then `bp_toggle modOrder:24` (the line in `OrderTotal` that calls `LineAmount`) while in **run** mode.
   - That's the same mechanism as in break mode: `SetSelection` on the code pane, then the Toggle Breakpoint command (Id 51).
2. `app_set_text` 2 and 10, `app_click` Add.
   - **The program stopped at line 24** (`mode: break`).
3. `run_continue`, then `bp_toggle modOrder:24` again in run mode, then `app_click` Add.
   - **No stop**: the breakpoint was cleared.
4. `run_start`, `bp_toggle` (set) in run mode, `bp_clear_all` in run mode (the Clear All Breakpoints command, Id 579), then `app_click` Add.
   - **The program stopped at line 24.** Clear All did nothing, yet the tool answered `cleared: true`.

## Decision

- **`bp_toggle` now works in design, break and run mode.**
- **`bp_clear_all` stays design and break only.** In run mode it would report success without clearing anything.
- To clear all breakpoints while the program runs, break first.

## Also from this session

- Step results carry `at`: the component, line, procedure and line text where the step stopped. Before, every `step_*` needed a `debug_where` after it.
- Verified: `step_into` from the breakpoint landed on `LineAmount`'s first line with `at` filled in.
