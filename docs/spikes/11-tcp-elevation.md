# Spike 11: TCP loopback across the elevation boundary

Date: 2026-10-13. Same machine as spike 10.

## Question

Spike 10 showed that a medium-integrity process (Claude Code, the MCP server) can't send window messages or input to elevated VB6, and PLAN rev 4 assumed it can't use the ROT either. Can a medium process talk to an elevated one over **TCP on 127.0.0.1**? If so, the in-IDE add-in can be the elevated endpoint, and we don't need a separate always-running elevated daemon.

## Test

- `spikes/11-tcp-elevation/listener.ps1`, run by the owner in an **elevated** PowerShell. It plays the add-in: listens on `127.0.0.1:<random port>` and writes `{pid, port, token, elevated}` to `%LOCALAPPDATA%\vb6mcp-spike\listener.json`.
- `spikes/11-tcp-elevation/client.ps1`, run from the Claude Code shell (**not elevated**). It reads that file, connects, and sends `AUTH <token>`: first a wrong token, then the right one.

## Result

```
client pid=10920 elevated=False -> listener pid=12240 port=55624 elevated=True
  sent token 'wrong-to...' -> reply: DENIED
  sent token 'fe16479a...' -> reply: OK pong from pid=12240 elevated=True
```

1. **Medium → high over loopback TCP works.** No integrity-level or UIPI restriction applies.
2. The token check works.
3. The elevated process runs under the **same user profile** (UAC consent elevation, not a separate admin account): both sides resolved the same `%LOCALAPPDATA%`. So a per-user discovery file works.
4. **No Windows Firewall prompt** for the loopback-only listener (owner confirmed). The only prompt was the UAC consent when opening the elevated PowerShell. The add-in won't trigger even that, because it runs inside an already-elevated VB6.

## Decision

- **Transport between the MCP server and VB6 = loopback TCP, served by the add-in itself.** The always-running elevated daemon, its Scheduled Task and the named pipe are dropped (PLAN rev 5 §3).
- The add-in is in-process with the IDE *and* with the program being debugged from the IDE, so it can also deliver UI input (channel C writes) without UIPI problems.
- Security: bind `127.0.0.1` only, random port, per-instance random token, discovery file in a per-user directory. Every request must carry the token.

## Follow-ups

- Spike 2 must build the real thing inside VB6: a raw Winsock listener using `WSAAsyncSelect` on a hidden window, so socket events arrive as messages on the IDE thread.
- Check: does the add-in still answer while a modal dialog (MsgBox, compile error) is up? Modal loops should keep dispatching messages to our hidden window **(H)**.
- Restrict the discovery directory's ACL to the current user, then confirm another user can't read it.
