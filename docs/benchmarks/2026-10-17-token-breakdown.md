# Where a run's tokens go (2026-10-17)

PLAN.md E3 said: measure before making screenshots smaller or results shorter. This looks at the four T12 Opus runs at $20 (two without `batch_run`, `bench/results/t12-opus-b20`, two with it, `bench/results/t12-opus-batch`) with `bench/token-breakdown.js`.

## Method

- Every tool result stays in the conversation and is read again (from the prompt cache) on every later turn. Its cost is about **size × turns left**. The script adds that up per tool, separately for result text, images and call arguments.
- Sizes are estimated: text at 4 characters a token, images at width × height / 750 (after scaling to 1568 px / 1.15 MP).
- **Calibration:** a least-squares fit of each turn's context growth (from the API's `usage`) against that turn's text characters and image pixels, over 725 turns, gave **2.3 characters per token** for text (JSON results: punctuation, numbers, hex handles) and **0.76 ×** the image estimate. So text weighs about 1.75 × the script's numbers, images a bit less. The shares below are corrected.
- Context at the first turn: 34.5k tokens (system prompt and tool definitions). At the last: 340–380k.

## Results

Share of the re-read tool traffic (calibrated, average of the four runs):

| Source | Share | Notes |
|---|---|---|
| `Read` of source files (`.vbp`, `.frm`) | ~37% | 7–10 reads of 4–5k tokens each, early in the run, so re-read ~180 times |
| `form_set_properties` results | ~15% | every change echoed back: `{control, property, ok, now, nowAll: null}` |
| `form_set_properties` call arguments | ~8% | the agent's change lists |
| `batch_run` results | 5–15% | the labelled results of 80–246 inner calls |
| `app_tree` | 2–8% | ~880 characters per form, `Class`, `Visible`, `Enabled` on every control |
| `form_properties` | ~5% | 50–100 properties per control, all of them |
| **Screenshots** (all tools) | **~3–5%** | 32–38 per run, ~300–420 tokens each |
| everything else | rest | |

## What it means

- **Screenshots aren't worth shrinking.** They are small (≈ 700 × 500) and few.
- **The JSON results are.** They are verbose and stay in the context for the rest of the run.
- **Reading whole files** is the largest single cost, but it's Claude Code's own `Read`, used deliberately (our instructions allow it). Not changed.

## Changes made (PLAN.md E3)

- `form_set_properties` lists only **failed** changes by default; `verbose: true` gives every change with the value read back.
- `form_properties` takes `names` (exact names, or prefixes ending in `.` such as `Font.`).
- `batch_run` takes `compact: true`: successful calls other than the last keep only their first 160 characters.
- `app_tree` drops the window class and writes `hidden`/`disabled`/`depth` only when they say something.

To be measured with the next T12 rerun, together with `form_layout` (E1), which should also shrink the `form_set_properties` traffic.
