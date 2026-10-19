# Spike 42: telling a slow vb6 call from a dead server (2026-10-19)

## Why

During the non-elevated test (spike 41) a vb6 call sat for 30 minutes, until the owner stopped it. VB6 was fine: it had started, the add-in had loaded, and the owner had closed a startup dialog. What happened:

1. `server_restart` answered, and the agent sent `ide_launch` in the same batch.
2. The old server, still alive for its last second, ran `ide_launch` (VB6 started) and then exited (`server_restart: exiting pid … now`).
3. The next call, `ide_dialog_respond`, went to the process that had just exited. Nothing answered, and Claude Code kept waiting.
4. Every later vb6 call in that session hung the same way: the client's connection never recovered.

The requests the owner saw on the VB6 toolbar came from other sessions' servers, which poll every IDE they discover. Nothing in the setup could tell "slow" from "dead". The owner: "there needs to be a way to figure out that things are not responding or frozen".

## What Claude Code does with a silent server (measured)

Setup:
- `spikes/42-liveness/fake-server.js`: an MCP server whose one tool is silent, answers late, or sends `notifications/progress` every 3 s before answering.
- Claude Code 2.1.300 driving it headless: `claude -p … --mcp-config <file> --strict-mcp-config`.

| Setting | Call | Result |
|---|---|---|
| per-server `"timeout": 10000` | silent | aborted: "timed out after 10s" |
| per-server `"timeout": 10000` | progress every 3 s, answer at 21 s | aborted after 10 s: **a total limit; progress doesn't extend it** |
| `CLAUDE_CODE_MCP_TOOL_IDLE_TIMEOUT=10000` | silent | aborted after **30 s**: "sent no response or progress for 30s". 30 s is the minimum |
| same | answer after 21 s, silent | answered (under the 30 s minimum) |
| same | answer after 45 s, silent | aborted after 30 s |
| same | progress every 3 s, answer at 45 s | **answered at 45 s: every progress notification resets the idle timer** |
| neither (defaults) | silent | the incident: aborted only after 1800 s |

So the per-server `timeout` is no use for slow vb6 calls (`ide_launch` up to 90 s, builds, `app_wait` up to 120 s). The idle timeout plus progress notifications is the right pair.

## What changed

- **Heartbeat** (`McpServer.HeartbeatAsync`): a tool call that brings a `progressToken` (Claude Code and Copilot CLI send one with every call) gets `notifications/progress` after 5 s, then every 10 s, until its result is out. The message is `<tool>: still working (N s)`, and `progress` counts up.
- **Idle timeout:** the installers set `CLAUDE_CODE_MCP_TOOL_IDLE_TIMEOUT=120000` (2 minutes) in Claude Code's `settings.json` `env`, through `vb6mcp install-hook`.
  - The setting is global, so a value the user already set is kept.
  - `-NoIdleTimeout` (both installers) leaves it alone.
  - The uninstall removes it only if it still has our value.
  - A dead or frozen server now shows as an error after 2 minutes instead of 30.
- **Restart:** `server_restart` sets `McpServer.Restarting`. Every call that arrives after that is answered at once with "the vb6 server is restarting … call again in a few seconds", so none goes to the exiting process unanswered.
  - A call that arrives in the same instant may still run; it finishes before the exit, since the server waits for running calls.
  - The tool description now says to wait for its answer.
- **Log:** every line in the server log carries the server's pid. Several servers (one per Claude session) append to the same file, which is how another session's requests were taken for this one's.

## Checked

- **Unit tests:**
  - `McpServerTests.SlowCallsSendAHeartbeat`: notifications carry the token, are numbered upwards, none comes after the result, and a call without a token gets none.
  - `McpServerTests.CallsWhileRestartingAreRefusedAtOnce`.
  - `HookTests.InstallSetsTheIdleTimeoutUnlessTheUserDid`.
- **End to end:** headless Claude Code with the idle timeout at its 30 s minimum, `app_wait` for 45 s on a real (non-elevated) IDE.
  - The 1.0.0 server was aborted after 30 s.
  - The new server answered after 45 s with `not seen within 45 s`.
- **Restart, scripted client:**
  - A call sent 200 ms after `server_restart`'s answer got "restarting" back in 0.1 s; the process exited 1 s later.
  - A call sent in the same instant ran and finished first.

## Not covered

- The **client side** of the original hang: why Claude Code's connection stayed dead for every later call, instead of starting the server again as it did after earlier restarts (PLAN.md E7). The idle timeout now turns that into an error within 2 minutes, but the session still needs a reconnect or restart. That's a Claude Code behaviour, not ours.
- **Other clients' idle handling:** Copilot CLI sends a `progressToken` (spike 40), so it gets heartbeats too. Its own timeout behaviour wasn't measured.
