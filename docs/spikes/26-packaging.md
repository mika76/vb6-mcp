# Spike 26: packaging and installing (Phase 6)

**Date:** 2026-10-15. **Machine:** the owner's VM. The owner approved installing on it.

**Questions:**
1. Can the server ship as one exe that needs no .NET on the target?
2. Does an install to Program Files with an admin-only roots file close the "widen the roots" hole?
3. Does a user-scope Claude Code registration work in the desktop app?
4. Do upgrade and uninstall work while VB6 and Claude Code hold the files?

## 1. One self-contained exe

| Publish | Size | Notes |
|---------|------|-------|
| framework-dependent (dev) | small, needs the .NET 10 x86 runtime | x86 runtimes are rare on user machines |
| self-contained single file | 110 MB | includes the Windows Desktop runtime (Vb6Mcp.UI) |
| + `EnableCompressionInSingleFile`, no pdb/xml | **48 MB** (41.6 MB zipped) | chosen |

A stdio smoke test of the compressed build passed:
- `initialize`, and `tools/list` (78 tools);
- `com_describe` (COM typelib reading, STA thread);
- `ide_list` (discovery and add-in TCP);
- `app_screenshot` of the IDE (through the add-in).

Startup is immediate. Watchdog mode is the same exe.

## 2. The roots hole, and the fix

**The hole:**
- In the development layout the elevated add-in read its workspace roots from `%LOCALAPPDATA%\vb6mcp\roots.txt`, which the user can write. Its DLL and the watchdog exe sit in the repo's build folder, also user-writable.
- Any ordinary process of the user can reach the add-in, because the token file is the user's. By widening the roots it could get the elevated add-in to `/make`, `regsvr32` or load files anywhere: a UAC bypass.

**The fix:**
- The installer puts the add-in, `vb6mcp.exe` and `roots.txt` in `%ProgramFiles(x86)%\vb6mcp`.
- The add-in reads `roots.txt` next to its DLL when it exists (`RootsFile()`); otherwise it reads the `%LOCALAPPDATA%` one (development).
- The server reads `roots.txt` next to its exe when `VB6MCP_ROOTS` isn't set.

**Tested as the ordinary user after install:**
- Appending to `roots.txt`, creating `evil.dll` and renaming `vb6mcp.exe` all failed with `UnauthorizedAccessException`.
- I added `C:\Windows` to the user-writable `%LOCALAPPDATA%\vb6mcp\roots.txt`, then sent raw `REGSVR C:\Windows\SysWOW64\<missing>.dll` to the installed add-in. It answered **400 "isn't under the workspace roots (C:\Program Files (x86)\vb6mcp\roots.txt)"**. A non-existent file was used, so nothing could have been registered either way.
- `build_make` inside the roots worked, and so did `com_unregister`.

## 3. Registration and processes after restart

- **MCP server:** after `claude mcp add --scope user vb6 -- "C:\Program Files (x86)\vb6mcp\vb6mcp.exe"` and a session restart, the desktop app ran `C:\Program Files (x86)\vb6mcp\vb6mcp.exe`. User scope works. The earlier local-scope registration in this folder was removed first, so the two didn't compete.
- **Add-in:** VB6 loaded the add-in from Program Files (the ProgID's `InprocServer32`) and started the watchdog.
- **Watchdog location: not proven (correction).**
  - An elevated process's path is hidden from an ordinary one, so I checked indirectly: the old development watchdog exe could be renamed, and I took that to mean it wasn't running.
  - **That test is invalid.** A running exe *can* be renamed; the installer itself depends on that. Back on the development setup, the same rename succeeded while that exe was running. An exclusive open (`File.Open(..., FileShare.None)`) does show it, but not in Program Files, where an ordinary user can't open files for writing at all.
  - So the installed watchdog's location follows from `Sidecar.bas` (it looks for `vb6mcp.exe` next to the DLL first) but wasn't verified.
  - **Fix:** the watchdog now logs `exe=<path>` in its start line in `%LOCALAPPDATA%\vb6mcp\watchdog.log`, which an ordinary process can read. Check it on the next install.

## 4. Upgrade and uninstall with files in use

- **Uninstall with `-Force`** (VB6 open, Claude Code running the server):
  - It removed the hook, the user-scope MCP registration and the HKCU add-in entry.
  - It stopped the watchdog and unregistered the DLL (regsvr32 exit 0, ProgID gone).
  - `VB6McpBridge.dll` (loaded in VB6) and `vb6mcp.exe` (this session's server) were reported as in use and stayed.
- **Reinstall over the in-use files:**
  - Each was moved aside (`*.old-<time>`), the new copy installed and registered, and all user settings restored.
  - Leftover `.old-*` files are deleted by the next install once they're free.

**Found along the way:**
- **UAC prompt:** cancelled twice (the prompt wasn't seen); approved the third time. The error used to say "cancelled" for any failure; it now includes the real exception message.
- **Uninstall messages:** they printed an empty file name, because inside `catch`, `$_` is the error record (the same trap as in `install-server.ps1`).
- **Reinstall without `-Roots` after an uninstall:** correctly refused, but only after copying files and after a UAC prompt. Both scripts now check before elevating and before copying.
- **PowerShell `-File` doesn't parse arrays:** `-Roots a,b` arrives as one string. The installer splits on `,` and `;`, and passes roots to its elevated part as one `;`-joined string.

## Decisions

- **Distribution:** a zip with `vb6mcp.exe` (self-contained), `VB6McpBridge.dll`, `install.ps1`, `uninstall.ps1` and `INSTALL.txt`, built by `scripts\package.ps1`. It refuses an add-in build older than its source.
- **The installer:**
  - runs as the user and elevates once, only for Program Files and `regsvr32`;
  - verifies the registration in the registry;
  - registers the server in user scope and installs the hook;
  - works in Windows PowerShell 5.1.
- **Roots** are machine policy: changing them means running the installer again (one UAC prompt).
- **Development** stays as it was (`build-addin.ps1`, `install-server.ps1`). A development build of the add-in has no `roots.txt` next to it, so it reads `%LOCALAPPDATA%`.

## Release check: 1.0.0 (2026-10-19)

The published `vb6mcp-1.0.0.zip`, unpacked and installed with Windows PowerShell 5.1 on the development machine:

- **Install** (one UAC prompt): the four files in `C:\Program Files (x86)\vb6mcp`, `roots.txt` writable only by SYSTEM and Administrators, the ProgID's `InprocServer32` pointing at the installed DLL, the HKCU add-in entry, the user-scope `vb6` server and the hook.
- **A short session through the installed `vb6mcp.exe`** (stdio, driven by a script): `initialize` (1.0.0, 94 tools), then the steps below. The add-in answered as v1.0.0 (protocol 1) and started the installed 1.0.0 watchdog.
  1. `ide_launch` of a scratch copy of the Basic fixture: it went through the shell with a UAC prompt, because no elevated watchdog was running.
  2. The SourceSafe question was reported and answered with `ide_dialog_respond`.
  3. `code_read`, an edit, `symbol_refs`, `syntax_check`.
  4. `run_start`, `app_tree`, `app_screenshot`, `run_end`.
  5. `ide_quit discard: true`, which restored the files.
- **Uninstall** (one UAC prompt): it stopped the installed watchdog, unregistered the DLL, and deleted the Program Files folder, the ProgID and CLSID keys, the HKCU add-in entry, the user-scope server and the hook. `%LOCALAPPDATA%\vb6mcp` stays (logs, recovery copies), as documented; `-RemoveData` deletes it.
- **Afterwards on a development machine:** the installed hook and ProgID replace the development ones, and uninstalling removes both. Restore them with `install-server.ps1` (hook, no elevation) and `build-addin.ps1` (add-in, elevated).
