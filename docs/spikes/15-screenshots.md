# Spike 15: Screenshots of the IDE, dialogs and running forms

Date: 2026-10-13. Code: `spikes/15-screenshots/` (C# console, .NET 10, win-x86, `UseWindowsForms` for `System.Drawing`). Images: `spikes/15-screenshots/out/` (gitignored).

## Question

Can we capture images of the elevated VB6 IDE, its designer and code panes, dialogs and the running program's forms, including controls without a window handle (Label)? From which process: the non-elevated MCP server or an elevated one (watchdog)? (Owner's idea: let the watchdog take screenshots.)

## Method

For every visible top-level window of the VB6 process, plus the IDE's designer, code, project and properties panes, two captures:
- **`PrintWindow(hwnd, PW_RENDERFULLCONTENT)`**: the window renders itself into our bitmap.
- **Screen copy** (`Graphics.CopyFromScreen` = BitBlt from the screen DC) of the window rect.

A "blankness" check (share of the most common colour in a sample grid) flags failed captures. Run once **non-elevated** (from Claude Code) and once **elevated** (owner's admin PowerShell). VB6 was **covered** by another window each time.

## Results

| Scenario | Non-elevated | Elevated |
|----------|--------------|----------|
| `PrintWindow` on any VB6 window | ✗ returns false, blank image (blocked by UIPI) | ✓ ~25–100 ms |
| Screen copy, VB6 covered | Captures **whatever is on top** (the Claude window) | same |
| Main IDE window (design + run mode), covered | – | ✓ **perfect**: toolbox, designer, code, Project, Properties, Immediate, title `[design]`/`[run]` |
| MsgBox (`#32770`) while running, covered | – | ✓ pixel-perfect (screen copy was blank) |
| **Running Form1** (`ThunderFormDC`), covered by PowerShell | – | ✓ **perfect, incl. `Label1`** (no hwnd) and `Command1` |
| `PrintWindow` on IDE **child panes** (DesignerWindow, VbaWindow) | – | ⚠ misaligned/partial content. Use a main-window capture **cropped** to the pane's rect instead. |

## Findings

1. **Screenshots are possible and very useful.** Form layouts, lightweight controls (Label/Image/Shape/Line), dialogs, the IDE state, and before/after checks for input.
2. **They must come from an elevated process.** Non-elevated `PrintWindow` is blocked, and a non-elevated screen copy only works when VB6 is in front (bringing it forward steals the user's focus).
3. **Elevated `PrintWindow` captures covered windows without changing focus.** That's exactly what the agent needs while the user keeps working.
4. IDE panes: capture the top-level IDE window once and crop by child rect.
5. `PrintWindow` sends a message to the target window, so a **hung** VB6 would block it. Run it on a worker thread with a timeout; fall back to a screen copy (visible area only) when it times out.
6. Sizes are small (Form1 320×240, IDE 1152×874; PNG). For MCP image results, crop to the window of interest and downscale when large.

## Decision

- **The watchdog hosts screenshots** (owner's suggestion, now confirmed): it's elevated, a separate process, survives a hung IDE, and can apply a timeout + screen-copy fallback. The add-in could also capture in-process, but would have to encode images inside VB6 and is unavailable while the IDE thread is busy.
- MCP tools: `app_screenshot(window|form)`, `ide_screenshot([pane])` return **MCP image content**. Watchdog verb: `screenshot(pid, hwnd|"ide"|pane, maxWidth)`. Without the watchdog: screen copy from the MCP server, only if the window is visible, marked as such.
- `System.Drawing` (Microsoft package / Windows Desktop SDK) is allowed (owner: Microsoft NuGets aren't external dependencies).
