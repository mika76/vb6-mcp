# T16: a bug the code doesn't show (2026-10-18, Sonnet, dry run)

A harder T14. In T14 both configs could read the cause off the code, so neither needed the debugger. Here the cause is only visible at run time.

## The task

`tests\fixtures\CartTax` (written by `bench\make-carttax-fixture.cs`) is the order form with a tax category (Food, Standard, Books).
- **Where the rates come from:** `lib\RateLib.dll`, an ActiveX DLL shipped **compiled only**. Its source (`tests\fixtures\RateLibSrc`) is never copied into a run.
- **The planted bug** (marked only in the generator): the form passes `UCase$(Left$(cboCategory.Text, 4))` as the category code. RateLib knows `FOOD`, `STD` and `BOOKS`, and gives unknown codes the standard 20%.
  - So Books (`BOOK`) is taxed.
  - Standard (`STAN`) looks right only by accident.
- **What the code doesn't tell you:** the codes are in the DLL. `com_describe` shows `Codes()` and `RateFor(code)`, and calling them in the running program shows the codes and rates.

Prompt: books should be tax-free, but 10 × 1 Books shows 12.00 instead of 10.00; Food and Standard look right. The fix has to be in Cart.

**Scoring:**
- The fix is saved, and RateLib is untouched.
- The project builds.
- The replay (`bench\replay.ps1`, `Invoke-CartTaxReplay`) enters Books 10×1 (expects 10.00), Food 10×2 (31.00) and Standard 10×1 (43.00).
- Recorded alongside: the lines changed, whether the agent ran the program, and which debugger tools it used.
- Checked both ways first: the shipped build replays 12.00/33.00/45.00, and a copy with the codes mapped replays 10.00/31.00/43.00.

## Dry run (1 run each)

| | Result | Tokens | Cost | Time | Tool calls (errors) | How it found the cause |
|---|---|---|---|---|---|---|
| A (no tools) | **failed** | 282,243 | $0.09 | 35 s | 7 (1) | Called RateLib from a script and found the codes. Rewrote `frmCart.frm` with a script: LF line endings and 5 `Attribute` lines lost, so VB can no longer load the form ("could not be loaded"). Couldn't build or run (no elevation). |
| B (vb6 tools) | correct, unaided | 1,213,279 | $0.25 | 78 s | 43 (12) | `com_describe` → `run_start` → reproduced 12.00 with `app_select`, `app_set_text` and `app_click` → `run_break` → `debug_eval m_rates.Codes()` → fixed `CategoryCode` → ran again and saw 10.00 → saved |

- B used the debugger the way the task intends, with `debug_eval` on the running program's own objects.
- B's 12 errors:
  - 6 were app-tool arguments. It passed `hwnd` instead of `control` (twice), then control names such as `txtPrice` (3 times), and `app_select` without `control` (once).
  - 3 were `proc_replace` argument names (`code`, `file`), each fixed after one error message.
  - 1 was `debug_eval` on `m_rates` while the program wasn't in break mode (run-time error 424).
  - 1 was a PowerShell COM call: 64-bit PowerShell can't create a 32-bit VB class.
  - 1 was `com_describe` called without arguments.

## Changes from this run

- App tools accept `hwnd` as another name for `control`, the name `app_tree` uses for the handle.
- A control name in `control` gets an error that says names can't be seen from outside the program, and that the handle comes from `app_tree`.

## What it shows

- With the cause hidden in a binary, only B could close the loop: reproduce, inspect the live objects, fix, re-run.
- A reasoned as well, and probed the DLL from a script. But it can't compile or run VB6 without elevation, and its script edit broke the form file.
- B's cost is mostly the debugging loop (44 turns), and the argument errors added to it. Those are the cheapest part to remove.

## Follow-up runs

**Fewer argument errors.** After the argument fixes, two more B runs were correct with 3 and 1 tool errors, against 12 in the first run.
- The fix: every tool now names all missing required arguments at once, with the names it doesn't take and the ones it does.
- **But neither run used the debugger.** Both read the codes out of `RateLib.dll` itself: VB stores string literals as plain UTF-16, so a strings dump shows `FOOD,STD,BOOKS`.

**The fixture now hides the codes.** RateLib keeps them shifted by one letter (`CPPLT` for `BOOKS`) and decodes them at run time. The DLL no longer contains the codes as text in either UTF-16 or ASCII.

**Round 4 (2 runs each):**

| Run | Result | Tokens | Cost | Time | Tool calls (errors) | How it found the codes |
|---|---|---|---|---|---|---|
| B-1 | correct, unaided | 618,200 | $0.16 | 50 s | 25 (3) | a VBScript run by the 32-bit `cscript`, calling `Codes()` and `RateFor` |
| B-2 | correct, unaided | 623,545 | $0.16 | 60 s | 30 (2) | `run_break` + `debug_eval` in the running program |
| A-1 | right fix, **not checked** | 287,724 | $0.09 | 34 s | 8 (1) | a script calling the DLL. It couldn't build or run Cart (no elevation) and said so |
| A-2 | failed | 286,312 | $0.10 | 40 s | 7 (0) | the same. Its script rewrote `frmCart.frm` with broken line endings, so the build fails |

- **The cause now takes run-time probing.** Every run found it by calling the DLL, through the debugger or through a script. Only B can then check the fix in the program itself.
- **Scoring change.** T14 and T16 count a fix as not done when the answer says it couldn't run, build or start the program (`bench\scoring.ps1`), as T4 and T5 already did. Rescored with that rule, A-1 above is "not done": the prompt asked to check the fix by running the program.
