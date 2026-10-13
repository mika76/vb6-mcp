# Spike 1: Hand-written MCP server (no SDK)

Date: 2026-10-13. Code: `spikes/01-mcp-handshake/` (C#, .NET 10, win-x86, BCL only).

## Question

Can we skip the MCP SDK (principle 1) and write the protocol ourselves with `System.Text.Json`, and does that work with the real client, Claude Code in the desktop app?

## What was built

Newline-delimited JSON-RPC 2.0 over stdio:
- `initialize` (version negotiation from a supported list)
- `notifications/initialized`
- `ping`, `tools/list`, `tools/call`
- error handling for an unknown method (-32601), malformed JSON (-32700), and an unknown tool (`isError: true` result)

stdout carries protocol messages only. Diagnostics go to stderr, or to a file when `VB6MCP_LOG` is set. One tool, `vb6_processes`, lists VB6 instances with title and elevation. About 150 lines.

## Results

| Test | Result |
|------|--------|
| Raw protocol script (all methods + error cases) | All correct |
| `claude mcp list` / `get` health check (CLI 2.1.90) | ✓ Connected |
| Tool call from inside a desktop-app session (Claude Code 2.1.295) | **Works**: returned `{"count":1,"processes":[{"pid":3500,"title":"Microsoft Visual Basic","elevated":true}]}` |

## What the real client sends (from the server log)

1. **`server/discover` probe first** (desktop app, Claude Code 2.1.295). Request id `server-discover-probe-1`, with `_meta` carrying protocol version **`2026-07-28`**, clientInfo and clientCapabilities. We answered `-32601 Method not found`, and the client **fell back to `initialize`**. No problem, but a newer protocol generation exists. Revisit once its spec is checked by testing.
2. **`initialize` asks for protocol `2025-11-25`**, which wasn't in our supported list. We offered `2025-06-18`, and the client accepted. Negotiation works.
3. **Client capabilities**: `roots: {listChanged: true}`, `elicitation: {form: {}, url: {}}`. Elicitation could carry user confirmations, e.g. "break binary compatibility?".
4. **`tools/call` includes `_meta.progressToken`** (and `claudecode/toolUseId`), so long operations such as builds can send progress notifications.
5. Fixed: the default `System.Text.Json` escaping turned every `"` in tool text into `"`. Use `JavaScriptEncoder.UnsafeRelaxedJsonEscaping` for output.

## Registration findings (desktop app)

- A project **`.mcp.json` did not load** in the desktop app across two restarts, even with `enabledMcpjsonServers: ["vb6spike"]` in `.claude/settings.local.json`. Possibly only read when a session is created; not pursued.
- **`claude mcp add --scope local <name> -e KEY=VAL -- <exe>` works.** It's stored in `~/.claude.json` under this project's path, and the server loaded after a restart. Gotcha: `-e` is variadic, so put the server name *before* it.
- Local-scope servers are **per project folder**: `claude mcp get` from another directory says "not found".
- Claude Code MCP servers **don't appear in the app's Connectors UI** (+ menu or Settings → Connectors). Check with `claude mcp get <name>` from the project folder.

## Decision

- **Hand-written MCP confirmed.** No SDK. Production code starts from this spike.
- Answer unknown methods with -32601 (the client relies on it to fall back). Keep a supported-version list and fall back to the newest one we support.
- Registration for development and installation: `claude mcp add --scope local` (later user scope, from the installer script). Drop the project `.mcp.json`.
- Use `progressToken` for builds; consider elicitation for confirmations.
