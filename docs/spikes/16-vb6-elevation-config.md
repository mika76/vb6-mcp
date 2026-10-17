# Spike 16: Why VB6 starts elevated here, and what that means for other users

Date: 2026-10-13. Goal (owner): the MCP server should work for **other people** too, not only for this machine's setup.

## Question

Starting `VB6.EXE` from a non-elevated process always triggered UAC (spike 2). Is VB6 inherently elevated, or is that this machine's configuration? Can an agent start VB6 non-elevated?

## Findings

| Check | Result |
|-------|--------|
| HKCU `AppCompatFlags\Layers` for `…\VB98\VB6.EXE` | `~ WINXPSP3` (no explicit `RUNASADMIN`) |
| HKLM / WOW6432Node layers | none for VB6 |
| Manifest `requestedExecutionLevel` in VB6.EXE | none |
| Start-menu shortcut | "Run as administrator" **off** |
| Pinned taskbar shortcut | "Run as administrator" **on** (how the owner normally starts VB6) |
| Start the real exe with `__COMPAT_LAYER=RunAsInvoker` | **fails before starting: "The requested operation requires elevation"** |
| Start a **copy** of VB6.EXE (other path), with and without `RunAsInvoker` | **starts non-elevated** (then exits with 0xC0000135 "DLL not found", as expected without its DLLs) |

## Conclusions

1. **VB6.EXE itself doesn't require elevation**, and neither does Windows' compatibility database (it would have matched the copy too).
2. The forced elevation here comes from **path-specific configuration** of the installed exe (the compatibility entry on that path; `RunAsInvoker` can't override it).
3. So **other machines vary**: a default install starts non-elevated; many developers run it as administrator anyway (building ActiveX DLLs/OCXs registers them, which needs admin rights).

## Decisions

- The server supports **both elevated and non-elevated VB6**. Nothing in the architecture depends on elevation: loopback TCP and the add-in work at any integrity level (spike 11). Elevation only matters for builds/registration and for who can send input.
- `ide_list` / `ide_status` report whether each IDE is elevated (readable from the non-elevated side).
- **`ide_launch(project?, elevation = auto | elevated | normal)`**:
  - `auto`: start VB6 the way it's configured on that machine (UAC prompt if it's configured to elevate).
  - `elevated`: through the watchdog (no prompt) when one runs, otherwise with a UAC prompt.
  - `normal`: `RunAsInvoker`; if the machine forces elevation (as here), fail with this explanation.
  - With a project path, VB6 skips the New Project chooser; without one, the chooser is reported.
- README gets an "Elevated or not?" section: what works in each mode (non-elevated: builds of ActiveX projects can't register; elevated: input only through the add-in, Claude's computer use can't drive VB6).

## Addendum (2026-10-17, from real use: improvement notes 15, 16)

- **A copy of the whole VB98 folder doesn't give a usable non-elevated VB6 either.** The copy starts non-elevated (the path-specific setting doesn't apply to it), but stops at startup with "Missing or not registered VB6TMPL.TLB": VB6's registration points at the installed folder. So "run a non-elevated copy" isn't a way out without registry changes, which this project doesn't make.
- **A UAC prompt nobody answers times out**, and `ide_launch` couldn't say why it took so long. The result now has `startSeconds` (how long the start took, including the time the prompt was up) and `seconds` (until the add-in answered); a declined or timed-out prompt names how long it waited and points at the user-started watchdog (`vb6mcp watchdog --stay`, PLAN.md D6), which launches VB6 without prompts.
