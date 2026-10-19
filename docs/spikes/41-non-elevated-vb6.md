# Spike 41: vb6-mcp with VB6 not elevated and without compatibility mode (2026-10-19)

## Why

vb6-mcp was built and tested only against an elevated VB6. On this machine a per-path compatibility entry (`~ WINXPSP3`) and the owner's taskbar shortcut force elevation (spike 16). The README claimed "works either way" from the design (spike 11: loopback TCP works at any integrity level), not from a test. Many developers run VB6 neither as administrator nor in XP compatibility mode.

## Setup

- The owner removed the compatibility entry on `VB98\VB6.EXE` (HKCU `AppCompatFlags\Layers` now has no VB6 entry) and the "Run as administrator" setting of the exe and its shortcut.
- The elevated `--stay` watchdog was ended, so the non-elevated IDE would start its own.
- The add-in registration (HKCU entry plus the DLL from `build-addin.ps1`) stayed as it was.
- Server: the development build, with the liveness changes of spike 42.

## Results

| Check | Result |
|---|---|
| `ide_launch elevation=normal` | Started through the shell, **no UAC prompt**, add-in answering after ~5 s. `ide_list` shows `elevated: false`. With a watchdog already running, `ide_launch` (auto) went through it, also without a prompt. |
| Add-in, discovery, token | Loads and answers (1.0.0). |
| Watchdog | Started by the add-in, **not elevated** (`watchdog.log`: `elevated=False`). |
| Startup error of another add-in ("Data View: Method '~' of object '~' failed") | Appears on some launches, with or without compatibility mode, so it isn't caused by it. Before the add-in has loaded there is no watchdog yet, so `ide_dialog_respond` could only fail ("no add-in loaded yet and the watchdog isn't running"). **Changed:** when VB6 isn't elevated, the server now clicks the button itself (same integrity level, `LocalInput.Click`). Not yet seen live: the box didn't come up on the launches after the change. |
| `code_read`, `code_replace_lines` | Work. |
| `bp_toggle`, `run_start`, `app_click`, break at the breakpoint, `debug_locals`, `debug_eval`, `step_over`, `run_end`, `bp_clear_all` | Work. |
| `app_tree`, `app_set_text` | Work. |
| `app_screenshot` | Works through the server's own capture (`method: server PrintWindow`), Labels included. Without XP compatibility mode the form has the Windows 10 frame: `clientOrigin` is (8, 31) instead of (4, 27). The result reports it, so image-to-mouse coordinates stay right. |
| `form_screenshot` | Works. |
| Dialog answered through the add-in (SourceSafe question) | Works. |
| `build_in_ide` of a standard EXE | Succeeded (1.2 s). |
| `build_make` of an ActiveX DLL (closed project) | `/make` writes the DLL, then can't register it. The log has `Error accessing the system registry` next to `Build of 'ComFixture.dll' succeeded.` Nothing was registered: no ProgID or CLSID under HKCR, HKCU or HKLM. **Changed:** the result now has `registered: false` and a note naming the cause (administrator rights; the VB6 that ran `/make` isn't elevated) and the way out (an elevated `regsvr32`, or VB6 as administrator). Before, it said only "failed" with the bare registry message. |
| `com_register` | regsvr32 runs at VB6's level and fails with exit 5 (DllRegisterServer failed). **Changed:** the note now says why and what to do; before it said "the registry doesn't point at this file for everything". |
| `ide_quit` | Works; the non-elevated watchdog closes the IDE. |
| `--cat IDE,Launch` | 8/8 passed against the non-elevated IDE (Launch: launch through the non-elevated watchdog, then quit). |
| Demo scenario (`scripts/demo/video.js`) | Complete: reproduce, breakpoint, `same object: True`, fix, total 25.00. |
| VB6 without XP compatibility mode | No IDE problems seen in these runs. |

## Decisions

- **Supported, measured.** Everything the tools do works against a non-elevated VB6 except registering components, which needs administrator rights. The README's "Elevated or not?" section now says so from these results.
- **Registration failures say why.** `build_make` and `com_register` name the missing rights and the way out instead of a bare error.
- **Registration got elevated routes** (follow-up, same day): `com_register` runs through an elevated VB6's add-in, else the elevated watchdog (new verb `REGSVR`), else a UAC prompt with `elevate: true`, else it refuses before running a regsvr32 that can only fail. `build_make` uses an elevated watchdog when the running VB6 isn't elevated.
  - **Verified live** (2026-10-19, `scratch\comdll-ne\ComFixture.dll`, unregistered at the start): with a non-elevated VB6 and a non-elevated watchdog, `com_register` refused and named `elevate=true`. With `elevate: true` Windows showed the UAC prompt, the user approved it, and the result was exit 0, `verified`, via "regsvr32 elevated through a UAC prompt". Then, with an elevated `--stay` watchdog and no IDE open, unregister, register and unregister again all went through the watchdog with no prompt: exit 0 and `verified` each time.
  - Not run live: `build_make` choosing the elevated watchdog while a non-elevated VB6 is open (unit-tested routing only).
- **Dialogs before the add-in loads are answered by the server itself** when VB6 isn't elevated; the watchdog route stays for elevated VB6.
- **Security is better non-elevated.** Code the agent runs (Immediate, the project started from the IDE) runs with the user's rights, not the administrator's. The limits in docs/tools.md (`debug_eval`) matter far less there.

## Not covered

- **The server's own click** for a startup dialog, live: the "Data View" box didn't appear on the launches after the change.
- **OCX projects and binary compatibility builds** without elevation. They're expected to behave like the DLL (written, not registered); not run.
