# Spike 33: the hook asks instead of refusing

**Date:** 2026-10-17. **Why:** the owner chose "ask for permission" over "block" for side effects (`debug_eval`, and the agent's shell commands that would get around the vb6 tools). Claude Code's PreToolUse hook can answer `permissionDecision: "ask"`; what that does in the desktop app and in headless runs had to be seen, not assumed.

## How the hook answers

Exit code 0 and, on stdout:

```json
{"hookSpecificOutput":{"hookEventName":"PreToolUse","permissionDecision":"ask","permissionDecisionReason":"vb6-mcp security policy: ..."}}
```

## Results

1. **Desktop app (interactive, auto mode on):** a harmless `echo "sqlcmd -Q \"delete from orders where 1=0\""` brought up a permission prompt with the reason. Allowed, it ran. A second command whose own text contained the same pattern got the prompt too; the owner chose Deny, and the tool call was rejected ("The user doesn't want to proceed with this tool use").
2. **Headless (`claude -p --allowedTools Bash`, as the benchmark harness runs):** the same command didn't run. It is listed in `permission_denials`, and the agent reported the reason and didn't try to work around it. So "ask" means "no" when nobody is there to answer, even for a tool in `--allowedTools`.
3. The hook settings were picked up by the running session without a restart (the shell checks worked right after `install-server.ps1`).

## Decision

- Shell commands `ShellPolicy` flags, and `debug_eval` calls with a side effect and `allow_side_effect: true`, get "ask". Edit/Write on VB6 source files stay refused (there is always a right tool for that).
- The server runs a side-effect expression only with `allow_side_effect: true` **and** the hook installed in the user's settings with `mcp__vb6__debug_eval` in its matcher (`Hook.GuardsEval`). Without the hook an agent could pass the flag itself, so it is refused. `VB6MCP_EVAL=open` still allows everything without asking.
- The reason texts describe what the command does, since the user decides; they no longer say "blocked".
- Text matching flags commands that only mention a pattern (an `echo` of SQL); with "ask" that costs one click, not a refusal.
