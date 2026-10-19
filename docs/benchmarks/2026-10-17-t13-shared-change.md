# T13: a signature change in shared source (2026-10-17, Sonnet, dry run)

## The task

`tests\fixtures\Shared` (written by `bench\make-shared-fixture.cs`) is a small workspace built around one shared module:

| Project | Type | Uses `Common\modFormat.bas` | Notes |
|---|---|---|---|
| SharedLib | ActiveX DLL, binary compatibility | yes | `clsPricing` calls `FormatAmount` on a continued line |
| AppA | EXE | yes | references SharedLib (so it builds after it); calls in a form and a module, one split over a `_` continuation; a string `"FormatAmount("` that isn't a call |
| AppB | EXE | yes | |
| AppC | EXE | no | its own `Private Function FormatAmount` in a different file (hard rule 3): must stay byte-identical |

Prompt: add a required `ByVal decimals As Integer` parameter to `FormatAmount`. Existing callers pass 2. Save, rebuild what has to be rebuilt in the right order, and say which projects were rebuilt.

**Scoring** (bench\run-bench.ps1, `Verify-Task T13`). The run passes only if all of these hold:
- the new signature is in place;
- AppC is untouched;
- the string is kept;
- SharedLib, AppA and AppB all build afterwards (so every caller was fixed);
- the agent itself rebuilt those three;
- SharedLib.dll was written before AppA.exe (the outputs' write times show this).

Rebuilding AppC is recorded as unnecessary work, not as a failure. In config B, AppA is open in the IDE, so the shared module is live there.

## Dry run (1 run each)

| | Result | Tokens | Cost | Time | Tool calls | UAC prompts | How |
|---|---|---|---|---|---|---|---|
| A (no tools) | correct, not unaided | 239,015 | $0.23 | 78 s | 6 | **3** | Grep, then edits as bytes in Python, then `VB6.EXE /make` per project (a UAC prompt each, approved by hand) |
| B (vb6 tools) | correct, unaided | 409,962 | $0.27 | 81 s | 19 | 0 | `symbol_refs`, `ws_affected`, 5 × `file_read`, 5 × `file_replace_lines`, `project_save`, 3 × `build_make` |

Both runs found every caller, including the call split over a continuation. Both left AppC and the string alone, and both built the DLL before AppA. B took its order from `ws_affected` (AppB, SharedLib, AppA). A worked out the order by reading the `.vbp` files.

## What it shows

- The fixture is small enough to grep. A's text approach works here, and the only thing it can't do without help is build: each DLL or EXE `/make` from a non-elevated shell is a UAC prompt. A larger workspace (T6's 43 projects) would separate A and B much more on finding callers and order.
- **B read every file only to get its hash:**
  - `symbol_refs` had already returned every line that needs changing, but no hash, and `file_replace_lines` needs `expected_hash`.
  - So the 5 `file_read` calls were spent on the hash alone, and they make up most of B's extra tokens.
  - Fixed: `symbol_refs` now returns each file's hash, the `expected_hash` for `file_replace_lines`.

## Follow-up runs (B only, 2 each)

| Build / harness | Runs correct | Tokens (avg) | Cost (avg) | Time (avg) | `file_read` calls |
|---|---|---|---|---|---|
| dry run (above) | 1/1 | 409,962 | $0.27 | 81 s | 5 |
| `symbol_refs` with `hash`, roots = all of `D:\bench` | 2/2 | 451,016 | $0.18 | 46 s | 4, 4 |
| the same, roots = the run's own folder | 2/2 | 404,044 | $0.15 | 36 s | 5, 3 |

- **The harness gave B too wide a workspace.**
  - Its MCP config set the roots to all of `D:\bench`, which holds every earlier run's copy.
  - So `symbol_refs` scanned 201 projects (2,744 files) and listed 73 uses, mostly in other runs' folders, with the agent's own files among them.
  - Each B run now gets its own config with the roots set to its run folder (the read-only workspace for T6/T7).
  - Rounds T8–T12 had the same wide roots, so workspace-wide searches in them may have scanned earlier copies too.
- **The hash alone didn't stop the reads.**
  - The agent reads each file to see the original lines (the call split over `_` continuations) before it rewrites them; `symbol_refs` shows the joined statement.
  - With files this small that is a sensible choice, so nothing more was changed.
  - One run edited a call with `file_replace_lines` `find`/`replace` instead, which needs no read.
- Every B run used `ws_affected` for the order (AppB, SharedLib, AppA), built AppA in its IDE (it is open there) and the others with `/make`, and needed no UAC prompt.
