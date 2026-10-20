# Roadmap

Version 1.0.0 (2026-10-19): 94 tools, everything in [PLAN.md](../PLAN.md) phases 0-7, the fixes from the last round of real use re-tested in a live IDE. What's open, roughly by value:

## Soon

- Nothing queued. The last item, creating a project, is done since 1.0.0: `project_new`, and `project_save path` for a never-saved project (spike 43).

## v2

### Remote use (Claude on another machine, VB6 on its own box)

- **`vb6mcp listen`:** runs in the user's desktop session, started at logon (not by the elevated add-in), and listens on loopback or a named pipe. For each connection it starts an ordinary `vb6mcp` child on its stdio, so per-session state (ownership, runs, breakpoints) works unchanged.
- **`vb6mcp connect`:** a stdio relay to the listener, for `claude mcp add vb6 -- ssh -T vb6box vb6mcp connect`. SSH handles authentication and encryption, and no new port opens.
- **Optional streamable-HTTP transport** (`HttpListener`) for LAN/VPN without SSH: TLS and a bearer token, off by default. It's a documented exception to hard rule 8.
- **Server instructions for remote use:** the agent can't Read or Grep the sources, so it works through `file_read`, `code_read` and `symbol_refs`. Adapt the hook for an agent machine that has no VB6.
- **Spikes first:**
  - Does an SSH session see the desktop session's windows? It probably doesn't, which is why the listener exists.
  - Does a self-contained .NET 10 win-x86 build start on Windows 7 SP1? If not: Windows 10 x86 on the VB6 box, or a .NET 6 multi-target build (a trial build needed about 10 small shims).

### Agent abilities

- UserControls by their VB name and type in `app_tree`, instead of `ThunderRT6UserControlDC`.
- Designer vs running: the same form side by side, or a diff image, to find properties lost on load.
- A warning before an edit about identifiers VB would re-case project-wide (VB keeps one spelling per name).
- In-IDE compiles with other conditional constants (decided against for now; `project_settings` with `reopen` and `build_make defines` cover it).
- Replayable UI scripts (PLAN.md E6): a named, saved list of `app_*` steps with expected results, replayed in one call, so the same clicks can be re-run after every change and serve as regression tests for forms.

### Reach

- **Other MCP clients** (Claude Desktop, VS Code, Cursor): the server is plain MCP, but the instructions and the hook are written for Claude Code. Document what works and what's lost without the hook.
- **Unattended use:** the watchdog started at logon, so sessions never hit a UAC prompt.
- **CI:** a GitHub Actions workflow that builds and runs the unit tests (they need no VB6).
