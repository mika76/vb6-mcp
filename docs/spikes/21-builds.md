# Spike 21: builds (Phase 4)

Date: 2026-10-14. Elevated IDE with the Ui and Basic fixtures; `build_make` / `build_in_ide` through the add-in.

## Round 1

| Test | Result |
|------|--------|
| `build_make` on Ui with unsaved project settings | ✓ refused by the server (`/make` reads disk) |
| `build_in_ide` on Ui (`MakeCompiledFile`), `Ui.exe` already exists | ✓ written in 1.6 s, **no "replace file?" prompt** |
| `build_make` on Basic, **open in the IDE** | ✗ the log says `The file '...\Basic.vbp' is already open.` (twice), no build. **VB6 /make refuses a project that is open in any IDE** |
| `build_in_ide` with a compile error that exists only in memory | ✗ **no dialog**: `MakeCompiledFile` raised E_FAIL (`Method '~' of object '~' failed`) right away. The server had nothing to see and waited 5 minutes |

## Consequences

- The first design ("the add-in only builds projects loaded in this IDE") allowed exactly the projects `/make` can't build. Now:
  - **Open projects** are built in their IDE with `MakeCompiledFile`. If the project is saved, memory equals disk, so `build_make` passes it there and says so; with unsaved changes it is refused.
  - **Closed projects** use `/make` with a path. The add-in checks the path against `%LOCALAPPDATA%\vb6mcp\roots.txt` (written by `install-server.ps1` from the same roots as `VB6MCP_ROOTS`), requires an absolute path without `..`, and refuses projects open in its own IDE.
- The token is the real security boundary: anything that can read it can already drive the IDE. The roots check is defense in depth, because an elevated `/make` registers components machine-wide.
- `MakeCompiledFile`'s outcome is kept by the add-in (`MAKESTATUS`). On failure, the server finds the message and location with a full compile (`syntax_check`, spike 6).

## Round 2 (after the add-in rebuild)

| Test | Result |
|------|--------|
| `build_make` path, closed Basic | ✓ 1.6 s, log `Build of 'Basic.exe' succeeded.`, output written |
| Headless `/make` VB6 and our add-in | it **didn't load the add-in at all** (no `OnConnection`, no discovery file). The `ConnectMode` 3 guard stays as protection |
| Compile error in a module (scratch copy) | `Compile Error in File '...\modMath.bas', Line 8 : Variable not defined`, then `Build of 'Basic.exe' failed.` |
| Same in a module, another line; same in a form | **The line is the IDE's code line counted from 0**: IDE line 9 → 8, IDE 2 → 1, form IDE 7 → 6 (file line 40). The tools report the 1-based IDE line (`line`) and the file line (`fileLine`) |
| Compile error in memory, `build_in_ide` | `MakeCompiledFile` failed with a generic message ("Path/File access error" this time, E_FAIL before); the follow-up full compile gave "Variable not defined", modMath line 9, column 22 |
| ActiveX DLL (`ComFixture`), closed, `build_make` | ✓ built and **registered** at the build path (32-bit view: ProgID → CLSID → `InprocServer32`); `New-Object -ComObject` from 32-bit PowerShell returned "Hello, from 32-bit PowerShell" |
| Binary compatibility (`CompatibleMode` 2, reference copy), unchanged rebuild | ✓ succeeded |
| **Incompatible change** (extra parameter), `/make` | ✗ **no dialog**: the log says `There is a compatibility error between the current project and the version-compatible component: 'Hello' in the 'clsGreeter' class module has arguments and/or a return type that is incompatible ...`; the build fails and the DLL is untouched |
| Same change, in the IDE (`build_in_ide`) | `MakeCompiledFile` failed generically; the full compile showed the **decision dialog** (OK, Cancel, Break compatibility, Preserve compatibility (Advanced), a "treat others the same" check box). It was left for the user; the test pressed Cancel |

## Decision (round 2)

- **Log parsing:** status lines, compile errors with 1-based IDE lines plus the file line, compatibility errors (`kind: "compatibility"` with class and member), anything else as `other`.
- **Binary compatibility:** `/make` is safe by itself: it fails and changes nothing. In the IDE the break is a dialog; `build_in_ide` reports it as `status: "compatibility-decision"` and never answers it (hard rule 13). This answers **spike 14**.
- Test registrations to remove afterwards: `ComFixture.dll` (fixture and scratch paths, same CLSID).
