# Spike 40: vb6-mcp in agent harnesses other than Claude Code (2026-10-19)

## Why

vb6-mcp 1.0.0 was built and tested only with Claude Code. The server itself is plain MCP over stdio, but several parts assume Claude Code:
- the installer (`claude mcp add`);
- the PreToolUse hook (blocks Edit/Write on VB6 files, guards shell commands, asks the user before `debug_eval` side effects);
- the server `instructions`;
- `CLAUDE_CODE_SESSION_ID` for the IDE handover after `server_restart`.

The owner asked how much work supporting other harnesses would be. This spike measures what actually breaks before any code is written (AGENTS.md principle 2).

## Setup

- The development server `build\server\current\vb6mcp.exe` (1.0.0), `VB6MCP_ROOTS=D:\_other\vb6-mcp\tests\fixtures`. `VB6MCP_LOG` was pointed at a separate file for each run, so the server's own log shows exactly what each client sent: initialize, capabilities, `tools/list`, every `tools/call`.
- The elevated watchdog was already running, so `ide_launch` started VB6 without a UAC prompt.
- Each harness was registered for one run only, through its command-line overrides. No user config was changed.
- Two tasks:
  - **Capability check:** how many vb6 tools does the model see, does it have the server's instructions, call `ide_list`.
  - **Cart task:** open `tests/fixtures/Cart`, reproduce the wrong total (2 x 10.00 + 1 x 5.00 shows 10.00), screenshot the form and describe it, look at the values in break mode, try `debug_eval` with `allow_side_effect=true`, fix the bug unsaved, end the run, quit the IDE (discard).
- Extra probes, each a separate run:
  - a plain "fix the bug" prompt on a **copy** of Cart in a folder without AGENTS.md, project not open, to see whether the agent edits `.frm` files with its own tools;
  - `app_wait` for 120 s, to measure the client's tool-call timeout;
  - `server_restart` followed by `ide_list`;
  - a `debug_eval` side effect aimed at a throwaway file in a scratch folder.
- **Test artefact:** the harnesses were started from a Claude Code session, so the first Copilot CLI runs inherited `CLAUDE_CODE_SESSION_ID` (the server log says `session set`). In a real Copilot session it isn't set. Later runs unset it.

What was available on this machine:

| Harness | Version | Tested? |
|---|---|---|
| GitHub Copilot CLI | 1.0.98 (model: Claude Sonnet 5.5 via Copilot) | **fully** |
| OpenCode (desktop app's CLI, `opencode-cli.exe run --standalone`) | 2.0.29 (free model `opencode/big-pickle`) | **fully** |
| Codex CLI | 0.140.0 | **handshake only**: the account's model (`gpt-6-astra`) needs a newer Codex; every older model was refused for a ChatGPT account. Upgrading means installing, which needs the owner's OK. |
| VS Code + Copilot agent mode | 1.142.0 | **not measured**: driving its chat UI needs computer use, which wasn't granted for this run; `code chat -m agent` from the command line never started the server (no log written), probably waiting on a trust/start prompt in the UI. |
| Cursor | none | not installed (asking the owner first) |
| Cline (VS Code extension) | 4.1.24 | not tested: GUI only, same as VS Code |

## Results

### Server side (any client)

- `tools/list`: **94 tools, 76.9 KB of JSON** (about 19k tokens on every request that carries the tool list). Largest: `form_layout` 3.0 KB, `file_replace_lines` 2.8 KB, `code_replace_lines` 2.6 KB.
- Schemas use only `object/string/integer/number/boolean/array`: no `anyOf`/`oneOf`, no `annotations` (`readOnlyHint` etc.), no `outputSchema`. Names are at most 21 characters, so even prefixed (`mcp__vb6__…`, `vb6-…`) they stay under the 64-character limits some clients have.
- Protocol: newer clients ask for `2025-11-25` and get `2025-06-18` (our newest), which they all accepted.

### GitHub Copilot CLI 1.0.98

| Check | Result |
|---|---|
| Connects | Yes. It first probes `server/discover` (with `_meta` protocol `2026-07-28`), gets -32601, then sends `initialize` with `2025-11-25` and accepts `2025-06-18`. Client capabilities: `sampling`, `elicitation` (`form`, `url`). |
| Tools loaded | **All 94**, named `vb6-<tool>`. No limit hit. |
| Instructions | **Not used by default.** The model saw no vb6 text. Copilot puts a server's instructions in the system prompt only for allowlisted servers, or with `--allow-all-mcp-server-instructions`: with that flag the model quoted "Tools for the Visual Basic 6 IDE and VB6 projects." It does read `AGENTS.md` from the working folder as custom instructions. |
| Images | **Passed to the model.** The tool result carries the PNG in `binaryResultsForLlm`. The agent described the Label captions "Price" and "Quantity", which are lightweight controls with no hwnd and exist only in the image (spike 17). It also used the image to notice it had typed into the wrong boxes. |
| Tool-call timeout | **At least 120 s.** `app_wait` for 120 s came back normally after 120014 ms. Copilot sends a `progressToken` with each call. |
| Own edit tools on VB6 files | **Yes, in both runs.** On the Cart copy, Copilot's built-in `edit` changed `frmCart.frm`. That happened without the instructions, and **also with them** (`--allow-all-mcp-server-instructions`). In the second run it even read the files with `vb6-file_read` first, then edited with `edit`. Its `edit` kept CRLF and added no BOM. The fixture is pure ASCII, so damage to 1252 bytes or hidden lines wasn't tested. One of the two fixes named a variable `line`. That isn't a valid variable name in a VB6 form, because `Line` is the form's graphics method. Nothing compiled it, since the agent never used the IDE. |
| Edits through the tools (Cart task) | When asked to work in the IDE, it used only vb6 tools: 42 calls with `ide_launch`, `bp_toggle`, `run_start`, `app_set_text`, `app_click`, `app_screenshot`, `debug_locals`, `code_replace_lines` (unsaved) and `ide_quit discard=true`. It found the real bug (one `As New clsLine` object reused for every line). |
| `debug_eval allow_side_effect` | **Runs with no approval (security gap, see below).** |
| `server_restart` | **No reconnect.** The next call failed with "MCP tool catalog changed", and then the `vb6-*` tools were gone from the session. The `-p` session can't get them back. Interactive Copilot has `/mcp`, which wasn't tested. |

### OpenCode 2.0.29

| Check | Result |
|---|---|
| Connects | Yes: `initialize` with `2025-11-25`, accepts `2025-06-18`. Capabilities: `elicitation` (`form`, `applyDefaults`). Log: `mcp connected server=vb6 tools=94`. Config: `opencode.json` in the project folder (`"mcp": {"vb6": {"type": "local", "command": [...], "environment": {...}}}`). |
| Tools loaded | **All 94, but not as model tools.** OpenCode 2 runs MCP tools in "Code Mode": the model gets one `execute` tool, which runs JavaScript, and calls `tools.vb6.ide_list(...)` from inside it. The vb6 tools are reachable only through a `search()` function in that runtime. `Object.keys(tools)` lists only `opencode` and `browser`. In the capability check, the model looked at its catalog, saw no vb6 tools and said none existed. Told "they are under tools.vb6, find them with search()", it searched and then drove the whole Cart task: about 100 calls, a lot of fix-and-rerun looping with a free model, ending with `ide_quit discard=true`. |
| Instructions | **Not seen.** The model found no server instructions in its context. |
| Images | **Dropped.** A tool result reaches the model as the JSON that the `execute` script returns. `app_screenshot` came back as `{"window":"Order","width":400,…}` with no image, so the agent never saw the form. |
| Tool-call timeout | Not measured. |
| Own edit tools | Not run separately. In the Cart task it edited only with `code_replace_lines`. |
| `debug_eval allow_side_effect` | It passed `allow_side_effect=true` on its own initiative, on `CartTotal(m_lines)` and `total = 0`. Neither expression matches the side-effect patterns, so the guard wasn't exercised. The guard runs in the server and doesn't depend on the client, so the Copilot result applies here too. |
| `server_restart` | Not measured. |

### Codex CLI 0.140.0 (handshake only)

- Connects: `initialize` with `2025-06-18`, client `codex-mcp-client` ("Codex"), capabilities `elicitation`. `tools/list` returned all 94 tools.
- The model run failed before any tool call: `gpt-6-astra` (the configured model) "requires a newer version of Codex". `gpt-5.5`, `gpt-5.4` and `gpt-5.3-codex` are "not supported when using Codex with a ChatGPT account".
- Registration without touching the user's config works with `-c mcp_servers.vb6.command='…' -c mcp_servers.vb6.env={…}`. Inside TOML, use single-quoted literal strings for Windows paths. The installer equivalent is `codex mcp add`.
- Still to measure after an upgrade: instructions, images, the timeout (Codex has a per-server `tool_timeout_sec`; its default wasn't checked), edits, `server_restart`.

### VS Code (Copilot agent mode)

Nothing measured. A workspace `.vscode/mcp.json` (`"servers": {"vb6": {"type": "stdio", …}}`) was written and `code chat -m agent` sent a prompt, but the server was never started. To finish this by hand, open the scratch workspace and check:
- whether a trust prompt appears;
- the tool count shown in the tools picker (VS Code has a per-request tool limit);
- whether a screenshot reaches the model;
- whether `edit` touches a `.frm`;
- what the vb6 server log shows for `initialize`.

## The security gap: `allow_side_effect` is trusted for any client

`debug_eval` runs an expression with a side effect only when `allow_side_effect=true` **and** `Hook.GuardsEval()` is true. `GuardsEval()` checks whether the vb6-mcp hook is installed in `~/.claude/settings.json` (spike 33). The idea was that Claude Code then shows the user a permission prompt before the call reaches the server.

That check is about the machine, not the client. On a machine where Claude Code's hook is installed, every other client gets the trust as well, and no other client runs Claude Code's hooks. In Copilot CLI:

1. `debug_eval` with `Kill "<scratch>\evaltest.txt"` and no flag: refused, "call again with allow_side_effect=true: the user is asked to approve it".
2. The same call with `allow_side_effect=true`: **sent to the IDE with no prompt.** It failed only because `? Kill …` doesn't compile (Kill is a statement). The audit log recorded `side effect approved by the user (hook prompt)`, which is false.
3. `CreateObject("Scripting.FileSystemObject").CreateTextFile("<scratch>\evalproof.txt") is Nothing` with `allow_side_effect=true`: `status: ok`, `sideEffect: file change: "Scripting.FileSystemObject"`, and **the file was created** by the elevated IDE. The agent said it saw no approval prompt.

So in 1.0.0, any non-Claude-Code client on a machine with the hook installed can run side-effect expressions with the IDE's administrator rights by passing one flag. The README's "you are asked to approve it" holds only for Claude Code. The hook's shell checks and its Edit/Write block don't exist in other clients either. That's expected, but it should be documented.

### Fixed (2026-10-19)

Recommendation 1 below is done:
- `EvalPolicy.Approval` trusts the hook only for `clientInfo.name` `claude-code` (the exact name Claude Code 2.1.300 sends, from the server log). Claude Code 2.1.300 advertises only `roots`, no elicitation.
- Other clients that offer form elicitation get an `elicitation/create` request: the expression, the category, and one required boolean `approve`. Only `accept` with `approve: true` runs it. The server now sends requests of its own and matches the client's responses by id (`McpServer.RequestAsync`).
- Clients with neither are refused.
- The audit line says `Claude Code hook prompt`, `MCP elicitation in <client>`, or `VB6MCP_EVAL=open`.

Live check with Copilot CLI 1.0.98 (`copilot -p … --additional-mcp-config @file --allow-all-tools`, the same `CreateTextFile` expression as step 3 above, `allow_side_effect=true`):
- Copilot sends `clientInfo.name` `copilot-cli`, with capabilities `elicitation: {form, url}`.
- The server sent `elicitation/create`. Copilot answered `{"action":"decline"}` 1 ms later, since `-p` has nobody to ask. `--allow-all-tools` does **not** accept elicitation for the user.
- `debug_eval` returned "refused: the side effect was not approved (file change: …; answer: decline)". The file was **not** created. The audit log reads `refused, not approved (MCP elicitation in copilot-cli 1.0.98: decline)`.
- Interactive Copilot (`copilot -i`) first shows its own folder-trust prompt, which needs the user, so the elicitation form itself wasn't seen on screen. The accept path is covered by unit tests only.

## Findings

1. **Connecting is not the problem.** Every client that got as far as a handshake (Copilot CLI, OpenCode, Codex) negotiated `2025-06-18` and listed all 94 tools. No client hit a tool-count limit. Copilot CLI's `server/discover` probe already gets the -32601 that spike 1 added for the desktop app.
2. **The instructions don't reach the model** in Copilot CLI (without a flag or allowlist) or OpenCode. Copilot CLI with the instructions still edited the `.frm` with its own `edit`, so the instructions alone don't keep agents on the tools. The Claude Code hook was what did that. Every harness tested reads `AGENTS.md` from the working folder.
3. **Elicitation is everywhere.** Copilot CLI, OpenCode and Codex all advertise `elicitation` (Copilot and OpenCode the 2025-11 `form` mode). It is the portable way to ask the user before a side effect.
4. **Images depend on the client.** Copilot CLI passes them. OpenCode's Code Mode drops them. VS Code and Codex are unmeasured. Several tools are worth much less without the image: `app_screenshot`, `form_screenshot`, and anything about lightweight controls.
5. **`server_restart` is Claude-Code-only.** In Copilot CLI the tools disappear for the rest of the session. The handover also depends on `CLAUDE_CODE_SESSION_ID`.
6. **Timeouts are no problem in Copilot CLI** (≥ 120 s). The longest built-in waits are `ide_launch` (90 s), `app_wait` (≤ 120 s) and `build_make`. Other clients are unmeasured.
7. **OpenCode needs a hint to find the tools at all.** In OpenCode 2 an agent sees the vb6 tools only if it calls `search()` in Code Mode. Without being told, it concluded there were none.
8. Side observations from the Copilot Cart run, not about harnesses:
   - right after `ide_dialog_respond`, `ide_status` still listed the dialog it had closed;
   - one `app_click` by handle on "Add" had no effect, while the next one by text worked.

## Recommended changes, ranked

1. **Close the `allow_side_effect` gap (security, before anything else).**
   - Trust the flag only when the client is Claude Code (`clientInfo.name` from `initialize` is `claude-code`) **and** the hook is installed.
   - For any other client, ask through **MCP elicitation** when the client advertises it: show the expression and the matched category, run it only on `accept`.
   - With no elicitation, refuse.
   - The audit line must say how it was approved (hook, elicitation, `VB6MCP_EVAL=open`).
   - Update README/tools.md: "asked" means asked by the client.
2. **Rules file in the user's project.** A short `AGENTS.md` section, which every harness tested reads: use the vb6 tools for `.bas/.cls/.frm/.ctl/.vbp`, never the built-in edit tools. Offer it as an install option, or as an `ide_*`/`ws_*` tool that writes it on request, since the user's repo isn't ours to change silently. Copilot CLI ignored the server instructions even when it had them, so this is a hint, not a guard. Write the risk down for clients without a hook.
3. **Installer support per client.** `install.ps1 -Client claude|copilot|codex|vscode|opencode`:
   - Copilot CLI: `~/.copilot/mcp-config.json` (`mcpServers`, `"type": "local"`), and say how to allow the instructions;
   - Codex: `codex mcp add vb6 --env … -- <exe>`;
   - VS Code: `code --add-mcp '{…}'`;
   - OpenCode: an `mcp` entry in its config.
   The Claude Code hook stays Claude-only.
4. **Image fallback.** Cheap and useful wherever images are dropped: when a screenshot is taken, also return a short text description of what the image adds (lightweight controls with their captions and positions from the form designer, which `form_controls` already knows), plus an option to save the PNG to a file. Measure Codex and VS Code first.
5. **Tool profiles: not needed now.** No client capped the 94 tools. The cost is about 19k tokens per request, which hurts small-context or local models. If wanted later: `VB6MCP_TOOLS=core` (edit/run/debug/app, about 40 tools) as an environment switch. Measure VS Code's limit before building it.
6. **`server_restart` outside Claude Code.** Say in the result that only Claude Code restarts the server by itself; elsewhere the user reconnects (Copilot `/mcp`) or starts a new session. Base the handover on something other than `CLAUDE_CODE_SESSION_ID`, or document `ide_claim` as the way back.
7. **Timeouts.** Nothing to do for Copilot CLI. Before supporting Codex, measure its default `tool_timeout_sec` against `ide_launch` (90 s) and `build_make`. If needed, the installer sets it, or the long tools report progress (`notifications/progress`: clients already send a `progressToken`).

## Not covered

- Codex beyond the handshake (needs a CLI upgrade); VS Code Copilot agent mode, Cline and Cursor (GUI, or not installed).
- 1252 and hidden-line damage from another client's edit tool (the fixture is ASCII).
- Interactive Copilot CLI (`/mcp` reconnect, tool approval prompts) and its instructions allowlist setting.
