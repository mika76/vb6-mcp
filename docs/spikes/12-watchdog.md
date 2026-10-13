# Spike 12: Watchdog sidecar

Date: 2026-10-13. Code: `spikes/12-watchdog/` (C# WinExe, .NET 10, win-x86; client `wd.ps1`). The add-in starts it: `Net.bas` → `StartWatchdog`.

## Questions (PLAN §3.1)

1. Does a watchdog started by the add-in inherit elevation without a UAC prompt?
2. Does it survive VB6 being killed?
3. Can it tell "program busy in run mode" from a real hang?
4. Can it break a busy loop from outside the IDE thread?
5. Can it screenshot a busy or hung IDE?

## Setup

- Watchdog:
  - single instance (named mutex `Local\VB6McpWatchdog`)
  - loopback TCP + 128-bit token, discovery file `%LOCALAPPDATA%\vb6mcp\watchdog.json`, log `watchdog.log`
  - exits 2 minutes after the last VB6 is gone
  - verbs: `PING`, `STATUS`, `SCREENSHOT`, `BREAK`, `KILL … force`
- Add-in: `Shell(exe, vbHide)` on connect.
- Busy loop: `Command1_Click` with `Do: i = i + 1: If i > 2000000000 Then i = 0: Loop`, no `DoEvents`.

## Results

| Test | Result |
|------|--------|
| Started by the add-in, elevation | ✓ `elevated=True parent=<VB6 pid>`, **no UAC prompt** |
| `STATUS`, design mode | ✓ responsive (WM_NULL probe 0 ms) |
| Busy loop running: `STATUS` | ✓ `responds=False`, `IsHungAppWindow=True`, **title still readable: `[run]`** (cached caption, no message needed) → classify as **busy-running** |
| Busy loop: add-in `PING` | ✗ timeout after 5 s (expected: same thread) |
| Busy loop: `SCREENSHOT` | ✓ `PrintWindow` timed out after 3 s → **fell back to a screen copy**. VB6 was covered, so the image shows the window on top (inherent limit). |
| **`BREAK`: injected Ctrl+Break (`keybd_event` VK_CONTROL + VK_CANCEL)** | ✓ **title `[run]` → `[break]`**, yellow arrow on the loop line; IDE responsive again, add-in answers again; debug commands enabled (normal break, unlike the spike 3 deadlock) |
| `SCREENSHOT` after break | ✓ `PrintWindow` 57 ms, shows the current statement highlighted |
| `KILL <pid> force` | ✓ VB6 killed |
| **Watchdog after VB6 was killed** | ✓ **survives**: `parentAlive=False`, still elevated, still serving |

## Findings

1. Elevation by inheritance works: an elevated helper with no UAC prompt and no Scheduled Task (confirms PLAN §3.1).
2. The watchdog outlives its parent. VB6 doesn't kill children on exit, and there was no job object.
3. Health classification works from data readable without the IDE thread: window title (mode), `SendMessageTimeout(WM_NULL)`, `IsHungAppWindow`, dialog enumeration.
4. **Injected Ctrl+Break rescues a busy loop**, even on a keyboard without a Break key (the key is synthesized). Caveat: VB6 was probably the foreground window (the user had just clicked the button); behaviour while VB6 is in the background is still **(H)**.
5. **Title parsing:** the mode tag isn't always last. A maximized MDI child gives `Project1 - Microsoft Visual Basic [run] - [Form1 (Code)]`. Match `\[(design|run|break)\]` (fixed in the spike code).
6. **`Shell` blocked the IDE thread for ~9 s** while starting the watchdog (watchdog logged its start at :33, `Shell` returned at :42). Production must start the watchdog with `CreateProcess` (no wait), or from a deferred message, never blocking add-in load.
7. **Stale discovery files**: after a kill, the dead IDE's `instances\<pid>.json` stays behind (two found). The watchdog/MCP server must ignore and delete files whose PID is gone (and check the PID really is VB6, since PIDs get reused).
8. Screenshot fallback for a hung IDE only shows what's visible on screen. Acceptable: the title + `STATUS` already explain the state; the picture becomes useful again after Break.

## Decisions

- Watchdog design in PLAN §3.1 **confirmed**; add: title regex, stale-file cleanup, `CreateProcess` launch.
- The `ide_break` tool order: try the IDE Break command via the add-in when it responds; otherwise inject Ctrl+Break from the watchdog; then verify the title shows `[break]`.
- Kill stays `force`-only (rule 12). The result must list recoverable unsaved work (recovery store: next watchdog step).

## Spike 12b: `LAUNCH`, crash detection, `PRESS` (same day)

| Test | Result |
|------|--------|
| Add-in starts watchdog via `CreateProcess` | ✓ 98 ms (spike 6/8 build), later 0–16 ms |
| **Clean VB6 close** | ✓ add-in logs `OnDisconnection`, **its discovery file is deleted** |
| **Crash detection** (`STATUS`) | ✓ the two killed sessions are reported `state=crashed (discovery file left behind)`; the cleanly closed one isn't |
| Watchdog whose parent VB6 closed while another VB6 runs | ✓ keeps serving (`parentAlive=False`) |
| **`LAUNCH <vbp>`**: watchdog starts `VB6.EXE "<project>"` | ✓ new IDE **High integrity**, window after 1.5 s, **no UAC prompt** (owner confirmed: the only prompt of the session was opening the first VB6 manually) |
| Normal close of the launched IDE | ✓ clean `OnDisconnection`, its discovery file deleted |
| Add-in in the launched IDE | ✓ loads (`OnConnection mode=1`); its watchdog start exits on the mutex (`another watchdog is running`) |
| Same project already open in another IDE | VB6 shows **"The file '…Project1.vbp' is already open."** (VB6 locks open projects). Answered with `PRESS 15652 OK`. |
| `LoadBehavior` | VB6 **rewrote it to 0** after Add-In Manager unload/reload cycles. The user's Add-In Manager state overrides what the installer writes. |

Implications for `ide_launch` / `ide_recover`:
- After a real crash the original IDE is gone, so the "already open" dialog won't occur. But launches **must watch for startup dialogs** (missing references, source control, …) and report them, never answer them blindly.
- `ide_launch` should refuse (or warn) when the project is already open in a live IDE.

## Open

- Ctrl+Break injection while VB6 is not the foreground window.
- Idle exit after 2 minutes (check `watchdog.log`).
- Heartbeat/recovery store not built yet.
