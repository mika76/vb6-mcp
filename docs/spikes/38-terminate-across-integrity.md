# Spike 38: a non-elevated process can end the elevated IDE (2026-10-18)

## How it came up

The full T13–T18 benchmark run lost its VB6 IDE (pid 16004) halfway through T13.
- In run T13-A-2, config A (no vb6 tools, no hook), the agent couldn't build without elevation.
- It then ran `Stop-Process -Name VB6 -Force -ErrorAction SilentlyContinue` from its non-elevated PowerShell.
- VB6 was listed right after the command and gone three seconds later.
- The watchdog logged `recovery: pid 16004 crashed` at 01:48:31, during that run.

The design assumed that elevation protects the IDE from non-elevated processes (spikes 10/11 cover window messages and input). It doesn't for termination.

## What was tried

From this session's non-elevated shell (Medium Mandatory Level), `OpenProcess` with single access rights. Nothing was changed: each handle was closed at once. Targets: the elevated VB6 (pid 16152, started through the watchdog) and the elevated watchdog (pid 21280).

| Access | VB6 (elevated) | watchdog (elevated) |
|---|---|---|
| PROCESS_TERMINATE | **granted** | **granted** |
| PROCESS_VM_OPERATION | denied (5) | denied (5) |
| PROCESS_VM_WRITE | denied (5) | denied (5) |
| PROCESS_VM_READ | denied (5) | denied (5) |
| PROCESS_CREATE_THREAD | denied (5) | denied (5) |
| PROCESS_SET_INFORMATION | denied (5) | denied (5) |
| PROCESS_DUP_HANDLE | denied (5) | denied (5) |

## Findings

- **Elevation is still a boundary for code.** A non-elevated process can't read or write the elevated processes' memory, start threads in them or duplicate their handles, so it can't inject code. Window input stays refused as well (spike 10).
- **It is not a boundary for termination.** Any process of the user can end the elevated IDE or the watchdog, and the IDE's unsaved work goes with it.
- **Not checked:** a VB6 started through a UAC prompt (rather than by the watchdog) wasn't tested.

## Decision

- **Hook patterns.** The Claude Code hook's shell check (ShellPolicy) already asked before `taskkill`/`Stop-Process ... vb6`. It now also catches:
  - pipelines such as `Get-Process VB6 | Stop-Process`;
  - `wmic ... vb6 ... delete|terminate`;
  - CIM `Terminate`;
  - `.Kill()` on VB6.
  The unit tests include the exact command the agent ran.
- **Ending VB6 stays the tools' job:**
  - `ide_quit` closes it normally, and VB asks about unsaved changes.
  - `ide_kill force=true` reports the recovery copies.
- **Recovery copies** (the watchdog's poll of idle IDEs) are what saves unsaved code when something ends VB6 from outside.
