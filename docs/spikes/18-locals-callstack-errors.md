# Spike 18: Locals, Call Stack and runtime errors

**Question.** Phase 2 still needed the current procedure's variables, the call stack and a runtime-error flow. VBIDE has no debugger API. What can be read, and how?

**Setup.** `tests/fixtures/Debug` (project `DebugFixture`). `modCalc.Inner` has locals of several types (Long, String, a Variant array holding Null, Boolean, Date) and is called from `Outer`. `Divide(1, 0)` raises runtime error 11. Breakpoint on `Inner = acc`, program started with `run_start`, Compute clicked with `app_click`: the IDE stopped at line 23 with no help from the user. 2026-10-13, elevated IDE.

**Findings.**

1. **The Locals window can't be read as text.**
   - It is a custom-drawn `VbaWindow` captioned "Locals". Its only child windows are an `Edit` (the context, e.g. `DebugFixture.modCalc.Inner`) and a `...` button that opens the call stack.
   - The grid (Expression / Value / Type, with expandable nodes for arrays and the module) has no window text. A screenshot shows it correctly.
   - Reading it as text would mean hooking GDI text calls inside VB6.EXE. A bug there would crash the user's IDE and lose unsaved work, so that was ruled out.
2. **Immediate evaluation can stand in for it.**
   - `TypeName(x)` never raises an error. It gave `Long|String|Variant()|Null|Date|Boolean|Long` for `n, tag, items, items(2), when, flag, Inner`, including the function's return value. Many names fit in one line joined with `&`.
   - `CStr(x)` works for numbers, strings, dates and booleans (`02/01/2003`, `True`).
   - **`CStr` of Null raised a modal "Run-time error '94': Invalid use of Null" dialog.** Values therefore need a second pass that only covers types `CStr` can print.
3. **The Call Stack is a standard dialog.**
   - It is a `#32770` captioned "Call Stack" holding a `ListBox`, the buttons Show and Close, and a static header "Project.Module.Function".
   - From the non-elevated server, `LB_GETCOUNT` failed (UIPI, as with `SBM_*` in spike 17), so the add-in reads the list in-process.
   - It is modal, so opening it (View > Call Stack, found by menu path) is deferred (hard rule 14).
4. **The runtime error dialog is readable.**
   - "Microsoft Visual Basic" / `Run-time error '11':` / `Division by zero`, with buttons Continue (disabled), End, Debug and Help.
   - Pressing Debug through the add-in left the IDE in break mode at the failing statement: modCalc line 27, `Divide`.
5. **Side finding: reserved project names.** A project named `Debug` was renamed by VB to `vb4projectDebug` in memory, because `Debug` is reserved. The fixture is named `DebugFixture`.

**Decision.**
- **`debug_locals`:**
  - Names come from the procedure's own declarations: the signature, the return value, `Dim`, `Static` and `ReDim`, scanned by `Vb6Mcp.Core.VbLocals`, which has unit tests. Module variables come from `Members` (the IDE's parse).
  - Pass 1 gets the types with `TypeName`. Pass 2 gets the values of scalar types only: `CStr`, or `Len` + `Left$` + escaped line breaks for strings.
  - Arrays, objects and Null show their type only. No user code runs: only the current function's return value and variables are referenced.
- **`debug_callstack`:** add-in verbs `CALLSTACK` (deferred: View > Call Stack) and `LISTITEMS` (in-process `LB_GETTEXT`), then the dialog's Close button.
- **`debug_error`:** parses the dialog. `action: debug` presses Debug and returns the failing line; `action: end` ends the program, with `force` when the agent didn't start the run. Every reported dialog that is a runtime error now carries `runtimeError: { number, message }`.
