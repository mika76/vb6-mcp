# Spike 10: Computer use vs elevated VB6

Date: 2026-10-13. Machine: dev workstation, Windows 11 Pro 22000, VB6 SP6 running **as Administrator**.

## Question

Can Claude's computer use, and in general a process running without admin rights, click and type into the elevated VB6 IDE? If not, UI input has to go through our elevated daemon (PLAN.md §3, §7).

## Setup (measured)

| Process | Integrity level |
|---------|-----------------|
| Claude desktop app | Medium (0x2000) |
| Claude Code shell (spawned by the app) | Medium (0x2000) |
| VB6.EXE | **High (0x3000)**, elevated |

Measured with `OpenProcessToken` + `TokenIntegrityLevel` (`spikes/10-computer-use-elevation/uipi.ps1`).

## Tests and results

### A. Win32 from a medium-integrity shell (`uipi*.ps1`)

| # | Test | Result |
|---|------|--------|
| A1 | `SendMessageTimeout(WM_GETTEXT)` to the VB6 main window | **Works**; returned "Microsoft Visual Basic" |
| A2 | `PostMessage(WM_MOUSEMOVE)` to the VB6 main window | **Blocked**: returned false, `GetLastError` = 5 (access denied) |
| A3 | `SetForegroundWindow(VB6)` | Plain call refused (foreground lock). Works with `AttachThreadInput` to the current foreground thread. |
| A4 | Synthetic Alt+F (`keybd_event`) with VB6 confirmed foreground | **No effect**: no new VB6 window of any class appeared (checked all visible top-level windows of the VB6 PID before and after) |

Caveat on A4: there was no positive control. It wasn't confirmed that a *real* Alt+F creates a detectable new top-level window in VB6's command-bar menus.

### B. Claude computer use (`mcp__computer-use__*`)

| # | Test | Result |
|---|------|--------|
| B1 | `request_access("Microsoft Visual Basic 6.0")` | Granted, tier **full** (typing allowed by policy) |
| B2 | Screenshot | **Works**: the elevated VB6 window is fully visible, unmasked |
| B3 | `left_click` on the "Existing" tab of the New Project dialog | Tool reported "Clicked", but **nothing happened**: the dialog stayed on the "New" tab. The input was dropped without any error. |
| B4 | `key` / `type` into the same dialog | Refused by the computer-use tool **itself**: it treats the dialog as a file dialog and allows typing only in a "File name" box. That's a tool safety rule, not UIPI. Its message says not to work around this with scripting, so keyboard input through computer use wasn't tested further. |

The computer-use tool's own description says the same: elevated processes "cannot be controlled even when granted: Windows UIPI blocks input from lower-integrity processes". B3 confirms it in practice.

## Conclusions

1. **Claude's computer use can SEE elevated VB6 (screenshots) but can't CLICK it.** Clicks are dropped silently: the tool reports success while nothing happens. Keyboard input travels the same `SendInput` path and is blocked by the same mechanism (consistent with A2/A4). It wasn't tested directly through computer use because of B4.
2. Medium → high **reads work** (screenshots, `WM_GETTEXT`). **Writes and input are blocked** (posted messages, synthetic input).
3. Practical consequence: generic computer use is **not** a fallback for an elevated VB6 IDE or for apps running inside it. It does work for non-elevated targets, e.g. a compiled EXE launched normally.
4. Even where input does get through, computer use has its own app-specific safeguards (B4) that can block legitimate IDE interactions.

## Decision

- **Keep `ui_click` / `ui_keys` / `app_*` input in the elevated daemon** (PLAN.md §6–7). That's no longer conditional on this spike.
- Screenshots can be taken either by computer use or by the daemon (`app_screenshot`). The daemon's version is preferred: it captures a specific window, works when that window is covered, and needs no app grant.
- A silent drop is the worst failure mode. Every daemon input tool must **verify its effect** (window state/text changed, expected window appeared) rather than just report "sent".

## Open follow-ups

- Positive control for A4: confirm what a real Alt+F produces in VB6 (needed anyway for spike 5's command-bar work).
- Alternative worth weighing: run VB6 **non-elevated** day to day. Then computer use, the ROT and window messages would all work from medium. Builds and registration would still need the elevated daemon. That depends on why the team runs VB6 as admin: usually COM registration on make, writes to Program Files, and debugging ActiveX DLLs/OCXs that must register.
