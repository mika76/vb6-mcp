# Wire protocol (v1)

How `vb6mcp` (the MCP server, non-elevated) talks to the **VB6 add-in** (`VB6McpBridge`, inside the elevated IDE). The watchdog will use the same framing (Phase 2b).

Why TCP on loopback: medium → high integrity works over loopback TCP, while window messages, synthetic input and the ROT don't (spikes [10](spikes/10-computer-use-elevation.md), [11](spikes/11-tcp-elevation.md)). Why this format: the add-in is VB6, so the format must be trivial to produce and parse there with `Split`, `InStr` and `Mid$`. No JSON parser in VB6.

## Discovery

Each add-in instance writes **`%LOCALAPPDATA%\vb6mcp\instances\<pid>.json`** when it starts listening, and deletes it on a clean unload (`OnDisconnection`).

```json
{"pid":8684,"port":57392,"token":"3f9c...e1","protocol":1,"addin":"0.1.0","started":"2026-10-13T14:18:13"}
```

| Field | Meaning |
|-------|---------|
| `pid` | VB6.EXE process id (also the file name) |
| `port` | TCP port on `127.0.0.1` |
| `token` | 32 hex chars (128 bits, `CoCreateGuid`), required in every request |
| `protocol` | protocol version (this document: `1`) |
| `addin` | add-in version (`App.Major.Minor.Revision`) |
| `started` | local time the listener started |

Rules for the server:
- **Stale files**: a file whose `pid` is not running, or isn't `VB6.EXE` (PIDs are reused), belongs to a killed or crashed IDE. Ignore it, report it as `crashed` if asked, and delete it (spike [12](spikes/12-watchdog.md)).
- `%LOCALAPPDATA%` is per-user and only readable by the user, SYSTEM and Administrators (default ACL), so the token isn't readable by other users.

## Transport

- TCP, `127.0.0.1` only (the add-in never binds `0.0.0.0`). Random port chosen by Windows.
- A connection can carry **any number of requests**, strictly one at a time: send a request, read its response, repeat. The server may also open a new connection per request.
- **Encoding: Windows-1252** for all text (VB6 is ANSI). Characters outside 1252 can't be represented anyway.
- The add-in processes requests on the **IDE thread** (`WSAAsyncSelect` → window message). A request can only be answered when the IDE pumps messages. The server **must** use a timeout on every request (default 5 s) and report "IDE busy" on timeout (spike [2](spikes/02-addin-bridge.md), [12](spikes/12-watchdog.md)).

## Framing

Every message (request and response) is one **frame**:

```
<decimal byte length of body> LF <body bytes>
```

- The length is ASCII digits, no sign, no padding, followed by a single LF (0x0A). No CR.
- Maximum body length: **16 MiB** (16777216). Larger frames are a protocol error; the receiver closes the connection.

## Separators

| Name | Byte | Use |
|------|------|-----|
| US (unit separator) | `0x1F` | between fields |
| RS (record separator) | `0x1E` | between records in a table payload |

VB6 source code can't contain these bytes, so code text never needs escaping. Arguments and payload values **must not** contain US or RS. The add-in rejects arguments that do.

## Request

```
VB6MCP1 US <token> US <id> US <verb> [US <arg>]...
```

- `VB6MCP1`: magic + protocol version.
- `<id>`: a client-chosen request id (decimal or any short string without separators); echoed in the response.
- `<verb>`: upper-case verb name (see below).
- `<arg>`: positional arguments as strings. Text arguments (code) may contain CR LF.

## Response

```
VB6MCP1 US <id> US <status> US <payload>
```

The **first three US** delimit the header. The payload is everything after the third US (it can itself contain US and RS).

| Status | Payload |
|--------|---------|
| `OK` | verb-specific result |
| `QUEUED` | for deferred verbs: a short description; the action runs after the response was sent (hard rule 14) |
| `ERR` | `<code> US <message>` |

### Error codes

| Code | Meaning |
|------|---------|
| `400` | bad request: malformed frame/body, unknown verb, wrong argument count, separator inside an argument |
| `401` | bad or missing token; the add-in closes the connection after replying |
| `404` | project / component / member not found |
| `409` | not allowed in the current IDE state (e.g. edit outside design mode) |
| `500` | VB error while executing; message = `Err.Number: Err.Description` |

## Verbs (v1)

Enum values below come from `VB6EXT.OLB` (component type: StdModule 1, ClassModule 2, VBForm 5, VBMDIForm 6, PropPage 7, UserControl 8, …; member type: Method 1, Property 2, Variable 3, Event 4, Const 5; scope: Private 1, Public 2, Friend 3; proc kind: Proc 0, Let 1, Set 2, Get 3; project type: StandardExe 0, ActiveXExe 1, ActiveXDll 2, ActiveXControl 3).

All VBIDE access uses **indexed loops**; `For Each` over VBIDE collections fails with error 70 in run/break mode (spike [3](spikes/03-05-ide-probes.md)).

### `PING`

No arguments. Payload: `pid US addinVersion US protocol`.

### `PROJECTS`

No arguments. Payload: one record per project, fields:
`index US name US type US fileName US isDirty`
- `fileName` is empty for a never-saved project.
- `isDirty` is `1` or `0`.

### `COMPONENTS <project>`

Payload: one record per component, fields:
`name US type US isDirty US countOfLines US fileName`
- `countOfLines` is `-1` if the component has no code module (e.g. a resource file).
- `fileName` is the first file (`FileNames(1)`), empty if never saved.

### `READ <project> <component> [<startLine> <count>]`

Payload: `countOfLines US isDirty US <text>`
- Without the optional arguments: the whole module.
- `startLine` is 1-based; `count` is clamped to the end of the module.
- `<text>` is the code as the IDE holds it, **including unsaved edits**, lines joined by CR LF. It doesn't include the hidden `Attribute` lines or the `.frm` designer block (VBIDE never shows those).

### `MEMBERS <project> <component>`

Payload: one record per member, fields:
`name US memberType US scope US static US codeLocation US procKind US procStart US procBody US procCount`
- `codeLocation` is the line VBIDE reports for the member.
- For methods and properties, `procKind`, `procStart` (first line incl. comments above), `procBody` (the `Sub`/`Function`/`Property` line) and `procCount` (lines incl. comments) come from `ProcOfLine(codeLocation)` and `ProcStartLine/ProcBodyLine/ProcCountLines`. For other members those fields are empty.
- **Properties:** VBIDE's `Members` lists a property **once**, under its first kind (verified against the Basic fixture: `Value` with Get and Let appeared only as Get). The add-in therefore emits **one record per existing kind** (Get, Let, Set), found with `ProcStartLine(name, kind)`.
- `procStart` is VB6's own start of the procedure: the line after the previous procedure ends, so it includes blank lines and comments above the `Sub`/`Function`/`Property` line.

### `PROC <project> <component> <name> <procKind>`

Payload: `procStart US procBody US procCount US <text>`. `<text>` is `Lines(procStart, procCount)`.

### `OPEN <path> <mode>` (deferred)

Opens a project or group: `VBProjects.AddFromFile(path, exclusive)`.
- `mode`: `add` (keep open projects; a group results) or `replace` (`exclusive`: close them first).
- Validated immediately: extension `.vbp`/`.vbg` (400), file exists (404), and for `replace` no open project is dirty (409; VB6 would otherwise ask to save).
- Then answers **`QUEUED`**. The open runs from a fresh dispatch of the add-in's own window message, because it can show modal dialogs (missing references, source control, "already open"), which must not run inside the request handler (hard rule 14, spike 3).
- Outcome: the caller watches `PROJECTS` (the new project's `fileName`) and the IDE's dialogs. Failures are only in the add-in log.

### `REPLACE <project> <component> <startLine> <count> <expectedOld> <newText>`

Replaces `count` lines at `startLine` with `newText` (CR LF separated): `DeleteLines` + `InsertLines`.
- `count` = 0 inserts before `startLine` (`startLine` may be `countOfLines + 1` to append); empty `newText` deletes.
- **Guard:** the write happens only if `Lines(startLine, count)` still equals `expectedOld` (empty when `count` is 0). The check and the write run in one go on the IDE thread, so a user edit in between gives **409**, never an overwrite.
- 400 for a range outside the module or non-numeric arguments.
- Payload: `countOfLines US isDirty` after the edit.
- The add-in doesn't check the IDE mode or the VB6 line limits; the server does that before sending (design mode only; at most 1023 characters per line, at most 24 continuations; only characters that exist in Windows-1252).
- VBIDE writes join the IDE's undo history (spike 3: "Undo Insert Text").

### `SAVE <project>` (deferred)

Makes the project active (`VBE.ActiveVBProject`) and runs **File > Save Project** (command bar Id 3), which saves the dirty components and the `.vbp` exactly like Ctrl+S.
- 409 if the project was never saved (it has no file name, so Save would open a Save As dialog).
- Answers **`QUEUED`**: saving can show dialogs (source-control checkout, read-only files). The caller watches `COMPONENTS` (all clean) and the dialogs.

### `CLICK <hwnd>`

Posts `BM_CLICK` to a dialog button, given as a hex window handle (the server finds it by reading the IDE's windows).
- 400 unless the window exists, belongs to **this IDE process**, and has class `Button` (hard rule 8: input only to the IDE's own windows).
- Posted, not sent: the dialog reacts in its own message loop after the handler has returned.
- Payload: `clicked <hwnd>`. Whether the dialog really closed is for the caller to verify (spike 10: dropped input looks like success).

## Versioning

Breaking changes bump the magic (`VB6MCP2`) and the discovery `protocol` field. A server talking to an add-in with a different protocol version reports it and doesn't send requests.

## Change log

- **v1** (2026-10-13): initial version (Phase 1).
