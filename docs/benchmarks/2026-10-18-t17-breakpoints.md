# T17: breakpoint debugging (2026-10-18, Sonnet, config B)

## The task

`tests\fixtures\Shop` (written by `bench\make-shop-fixture.cs`) is a shop form. Its price list is a binary `products.dat`: a `Long` count, then records as VB's `Put` writes `Name As String * 12`, `Price As Currency`, `Active As Boolean`, with no padding.

**The planted bug** (marked only in the generator):
- Tea is in the file twice: an old record (3.20, inactive) before the current one (2.80).
- `FindPrice` returns the first name match and ignores `Active`.
- The code looks right. The duplicate is only in the binary data, which the running program shows at a breakpoint.

**The prompt** asks for the debugger: breakpoints, run to them, step, inspect variables. So T17 runs in config B only; A has no debugger.

**Scoring:**
- The fix is saved and `products.dat` is untouched.
- The project builds.
- `Invoke-ShopReplay` expects Tea 1 → 2.80, Coffee 2 → 7.80, Juice 1 → 10.80.
- Recorded alongside: whether a breakpoint was used, and which debugger tools.

Checked first:
- The fixture build replays 3.20/8.20/11.20, and a copy with `And Products(i).Active` replays the expected totals.
- VB's `Get` reads the generated layout as intended: Coffee 2.50 and Juice 3.00 come out right.

## Runs

| Round | Runs correct | Breakpoint used | Cost (avg) | Time (avg) | Tool errors | Debugger tools |
|---|---|---|---|---|---|---|
| 1 | 2/2 | 2/2 | $0.20 | 60 s | 4, 5 | `bp_toggle`, `debug_locals`, `debug_eval` (5×) |
| 2 (after the fixes below) | 2/2 | 2/2 | $0.18 | 52 s | 4, 3 | `bp_toggle`, `debug_locals`, `debug_eval` |

A typical session (round 1, run 2):
1. `bp_toggle modProducts:30` (the `If` inside `FindPrice`), then `run_start`.
2. `app_tree`, `app_select` Tea, `app_click` Add. The program stopped at the breakpoint.
3. `debug_locals`, then `debug_eval i` = 2, `CStr(Products(i).Price)` = 3.2, `CStr(Products(i).Active)` = False, `UBound(Products)` = 5.
4. `run_end`, clear the breakpoint, add `And Products(i).Active`.
5. Run again and see Tea at 2.80, then `project_save`.

Neither run stepped (`step_over` and the others): one breakpoint plus inspection was enough.

## What the runs exposed, and the fixes

- **`debug_locals` lost every value at a breakpoint in this program, and left a dialog open.**
  - The cause: it evaluates `TypeName(...)` of all variables in one Immediate line, including the module's `Products`, an array of a user-defined type from a standard module.
  - VB doesn't compile that ("Only user-defined types defined in public object modules can be coerced to or from a variant"). The compile error broke the whole line, and the agent had to close VB's dialog itself.
  - Now the tool closes the compile error its own line caused and evaluates that chunk one name at a time. Such a variable is reported as `(user-defined type)`.
  - Verified live at the same breakpoint: `productName = "Tea"`, `i = 2`, `FindPrice = 3.2`, `Products` reported as a user-defined type, no dialog left.
- **`debug_eval` of a whole record (`Products(i)`) left a compile-error dialog too.**
  - Now it is a tool error with VB's message ("Type mismatch"), the dialog is closed, and the IDE stays in break mode.
  - The error suggests evaluating the record's fields.
  - Run-time error dialogs in the program are still left open and reported (hard rule 9).
- **`app_select` with only `text`** (3 calls across the runs) now picks the form's only list or combo box, or the only one that has that item. Here that's the product combo, not the list of order lines.
- **Remaining errors** are deliberate or outside the tools:
  - `Read` refuses the binary `.dat`.
  - `code_replace_lines` without `expected_hash`: the hash guards against stale edits.
