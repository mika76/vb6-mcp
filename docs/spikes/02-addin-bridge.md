# Spike 2: Add-in bridge over loopback TCP

Date: 2026-10-13. Code: `spikes/02-addin-bridge/` (VB6 ActiveX DLL `VB6McpSpike`, test client `client.ps1`).

## Question

Can a VB6 add-in inside the **elevated** IDE serve requests over loopback TCP, so the **non-elevated** side can read live, **unsaved** code? Does it keep answering in run mode, break mode and behind modal dialogs?

## What was built

- `Connect.cls`: `Implements IDTExtensibility` (plain class, no designer). `OnConnection` starts the server, `OnDisconnection` stops it.
- `Net.bas`:
  - Raw Winsock (`ws2_32`): bind `127.0.0.1:0`, so Windows picks the port.
  - `WSAAsyncSelect` onto a hidden **message-only** window (`STATIC` subclassed with `AddressOf`), so socket events arrive as window messages on the IDE thread. No polling, no blocking calls.
  - Token: 128 bits from `CoCreateGuid`. Discovery file: `%LOCALAPPDATA%\vb6mcp\instances\<pid>.json`. Log: `%LOCALAPPDATA%\vb6mcp\addin-spike.log`.
  - Verbs: `PING`, `LIST` (project, component, type, `IsDirty`, line count), `READ project<TAB>component` (`CodeModule.Lines`).
  - Wire format (spike only): request `"<token> <VERB> args\n"`, response `"OK|ERR <bytes>\r\n<payload>"` (1252).
- The VBIDE reference GUID was read from `VB6EXT.OLB` (`{EF404E00-EDA6-101A-8DAF-00DD010F7EBB}` 5.3). The GUID written from memory would have been wrong (`…8DAA…`): rule 13 earned its keep.
- Add-in registration: `HKCU\Software\Microsoft\Visual Basic\6.0\Addins\VB6McpSpike.Connect` with `FriendlyName`, `Description`, `LoadBehavior=3`, `CommandLineSafe=0` (same as the existing add-ins on this machine). Loaded through Add-In Manager.

## Results

| Test | Result |
|------|--------|
| Compiles first try (`VB6.EXE /make`, 40 KB DLL) | ✓ |
| Add-in loads, listens, writes discovery file | ✓ `listening on 127.0.0.1:56296 pid=3500` |
| `PING` from non-elevated shell | ✓ ~15 ms per request (new connection each) |
| Unknown verb | ✓ `ERR unknown verb` |
| Wrong / missing token | ✓ `ERR bad token`, no data |
| `READ` of an **unsaved, never-saved** Form1 in design mode | ✓ **exact live code returned**: text that exists nowhere on disk |
| Run mode **with a modal MsgBox open**: `PING` | ✓ answers: socket messages are dispatched inside the modal loop |
| Run mode: `LIST` / `READ` | ✗ **`error 70: Permission denied`** |
| Break mode (IDE pause button): `PING` | ✓ |
| Break mode: `LIST` / `READ` | ✗ **`error 70: Permission denied`** |
| Back in design mode: `LIST` / `READ` | ✓ works again |

## Findings

1. **The bridge works end to end** across the elevation boundary, including the security token.
2. **The add-in stays responsive in run mode, in break mode and behind modal dialogs.** The (H) from PLAN §5 is confirmed for the modal-MsgBox case. Not tested: a busy loop that never pumps messages.
3. **New, important: VBIDE locks the object model in run and break mode.** Enumerating `VBProjects` raises error 70. So "read live code" works only in **design mode**. It's still unknown whether *every* VBIDE call is locked or only some (e.g. `ActiveCodePane`, `SelectedVBComponent`) → spike 3.
4. **The mode is visible in the IDE's main window title**: `Project1 - Microsoft Visual Basic [design]` / `[run]` / `[break]` (window class `wndclass_desked_gsk`). The **non-elevated** MCP server can read it directly, with no add-in and no event needed. That closes the "no break-mode event" gap.
5. The MsgBox is visible from the medium side as `#32770 'Project1'` → the dialog watchdog can be driven from window enumeration.
6. **Building from the non-elevated side triggers a UAC prompt.** `VB6.EXE` (compat layer `WINXPSP3` on this machine) asks for elevation when launched. The first build was approved and registered the DLL in `HKLM\SOFTWARE\Classes` + `WOW6432Node\CLSID`; the second launch was cancelled. → Builds must be started by the elevated add-in (PLAN §3), never from the MCP server.
7. On keyboards without a Break key (Mac), the IDE's **Break toolbar button** works. The watchdog/debug tools should run the IDE's Break command rather than rely on Ctrl+Break keystrokes.

## Decisions

- Transport, discovery, token and `WSAAsyncSelect` design: **confirmed** for production.
- **Live-state model** (PLAN §4) must handle the run/break lock:
  - The MCP server keeps a **snapshot** of each component's code, refreshed in design mode, and serves it in run/break mode marked `stale: true, reason: "ide in run|break mode"`.
  - Writes in run/break mode are refused with a clear error (or queued) until spike 3 shows what's possible.
  - Mode detection = window title (primary, works without the add-in) + `EnterRunMode`/`EnterDesignMode` events.
- Spike 3 adds: which VBIDE members work in run/break mode; snapshot on `BeginCompile`/`EnterRunMode` (is access still allowed inside the event?).
