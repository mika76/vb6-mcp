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

### `CMD <id|fullcompile>` (deferred)

Runs a debug command from the command bars. **Allow-list only** (hard rule 8): `186` Start/Continue, `189` Break, `228` End, `515` Restart, `188` Step Into, `194` Step Over, `2559` Step Out, `1811` Run To Cursor, `51` Toggle Breakpoint, `579` Clear All Breakpoints (ids from the spike 5 dump). `fullcompile` runs **Run > Start With Full Compile**, found by menu path because its id is the generic 746.
- 400 for any other id.
- Answers **`QUEUED`**: these commands run user code, and a break under the request handler leaves the IDE in break mode with every command disabled (spike 3; hard rule 14). `Execute` returns immediately; the run or break follows.
- A disabled command (e.g. Break while in design mode) is skipped and logged; the caller sees that the mode didn't change.

### `SETSEL <project> <component> <line>`

Shows the component's code pane and puts the cursor at the start of `line` (`CodePane.Show` + `SetSelection`). Needed before Toggle Breakpoint and Run To Cursor, which act on the cursor line. Payload: `selected <component>:<line>`.

### `WHERE`

The active code pane: `project US component US startLine US startColumn US endLine US endColumn US procedure`. In break mode this is the **current statement** (verified live: a breakpoint in `AddNumbers` gave line 8, then line 9 after a step); after a compile error it's the highlighted error (spike 6).

### `IMM <hwnd> <text>` (deferred)

Types one line into the Immediate window and presses Enter (focus via `VBE.Windows` type 5, then `WM_CHAR` per character and `VK_RETURN`; spike 5). `hwnd` is the Immediate window (a `VbaWindow` of this process, found by the server by enumeration since `VBE.Windows(i).HWnd` is always 0). Deferred: Immediate input runs code (`? f()`, `Call Foo`). 400 for multi-line text or a window that isn't a `VbaWindow` of this IDE.

### `IMMSELECT <hwnd>`, `IMMCOPY`, `IMMPEEK`, `IMMRESTORE`

Reading the Immediate window is a **sequence of requests driven by the server**, with the IDE idle in between:

| Verb | Does | Payload |
|------|------|---------|
| `IMMSELECT <hwnd>` | saves and clears the user's clipboard text, focuses the window (Win32 `SetFocus` on the handle; the add-in runs on the IDE thread), runs Select All (756) | `selected`; 409 if Select All is disabled at that moment (the server retries) |
| `IMMCOPY` | runs Copy (19) | `copied`, or `empty` if Copy is disabled (nothing to select: the window is empty) |
| `IMMPEEK` | reads the clipboard text, without restoring | the text (may be empty while the copy is still in progress) |
| `IMMRESTORE` | **collapses the selection** (Ctrl+End in the window), then puts the user's clipboard text back (non-text formats can't be preserved) | `restored` |

The server: `IMMSELECT`, ~150 ms pause, `IMMCOPY`, then `IMMPEEK` every 100 ms (up to ~2 s) until text arrives, and **always** `IMMRESTORE`.

Why so many steps (Phase 2 live tests):
- The window is custom-drawn: `WM_GETTEXT` returns only its caption (spike 5), so the text has to go through the clipboard.
- VB6 updates a command's `Enabled` state only when idle: Copy run in the same handler right after Select All was still disabled (empty result in design mode; `Method '~' of object '~' failed` in break mode when executing it anyway).
- **Copy completes asynchronously:** reading the clipboard right after `Execute` found it empty, and the late copy then overwrote the restored user text.
- **Select All must be undone:** with everything still selected, the next typed line **replaced the whole Immediate content**, erasing the user's `Debug.Print` output (seen in a screenshot). `IMMRESTORE` and `IMM` therefore send Ctrl+End first. Key messages carry no modifier state, so the add-in sets Ctrl in its own thread's keyboard state (`SetKeyboardState`; it runs on the IDE thread) and restores it afterwards.

No user code runs, so none of these is deferred.

### `CALLSTACK` (deferred), `LISTITEMS <hwnd>`

`CALLSTACK` runs View > Call Stack, found by menu path because its command id isn't known. It is deferred because the dialog is modal, and a modal dialog under the request handler would block the add-in. Payload: `QUEUED`, `call stack`. The server waits for the `#32770` dialog captioned "Call Stack" (spike 18).

`LISTITEMS <hwnd>` reads a list box of the IDE process (`LB_GETCOUNT`, `LB_GETTEXTLEN`, `LB_GETTEXT`). It's read-only, and it is done in-process because UIPI blocks `LB_*` messages from the non-elevated server. Payload: the items separated by RS (empty payload = no items). Errors (400): not a window of this process, not a list box. The server then closes the dialog by clicking its Close button with `CLICK`.

### `MAKE <vbp> [<defines>]`, `MAKEIDE <project>` (deferred), `MAKESTATUS`

`MAKE <vbp>` starts `"<VB6.EXE>" /make "<vbp>" /out "<log>"` as a child of the IDE with `CreateProcess`, without waiting. The child inherits VB6's elevation, so an ActiveX DLL/OCX is registered on success without a UAC prompt. The log goes to `%LOCALAPPDATA%\vb6mcp\builds\<time>-<name>.log`.
- **For closed projects only.** VB6 `/make` refuses a project that is open in any IDE ("The file '...' is already open.", spike 21), and the add-in refuses one open in its own IDE (409).
- **Path check:** the path must be absolute, without `..`, and under one of the roots in `%LOCALAPPDATA%\vb6mcp\roots.txt`, which `install-server.ps1` writes from `-Roots`. The token is the real boundary; this is defense in depth, because an elevated `/make` registers components machine-wide (hard rule 8).
- **`defines`** (optional, Phase 7 A4) becomes `/d <defines>`: `name=value` pairs joined by `:`, e.g. `noterm=1:debug=0`. It goes on an elevated command line, so the add-in accepts only VB identifiers as names and integers as values, at most 200 characters (400 otherwise).
- Payload: `pid US logPath US vbp`. Errors: 400 bad or unlisted path, or bad defines; 404 file missing; 409 open in this IDE, or the process couldn't be started.

`MAKEIDE <project>` queues `VBProject.MakeCompiledFile`, which compiles **from IDE memory** to the project's `BuildFileName`. It is deferred because compiling may show modal dialogs. Payload: `QUEUED US outputFile`; 409 while the program runs. It sets the outcome to `queued`.

`MAKESTATUS` returns the outcome of the last `MAKEIDE`: `queued`, `running`, `ok`, or `failed US message`. A compile error shows **no dialog**: `MakeCompiledFile` raises E_FAIL (`Method '~' of object '~' failed`, spike 21). So the server polls this, and on `failed` finds the message and location with a full compile (`syntax_check`).
**Command-line builds load add-ins too.** `VB6.EXE /make` loads startup add-ins with `ConnectMode` 3 (`vbext_cm_CommandLine`). The add-in then starts neither the listener nor the watchdog, and writes no discovery file.

### `WRITES <since>`, `RELOAD <project> <component> <clean|discard>` (deferred)

For the sync engine (PLAN.md §8; the watchdog's `SyncPoller`).
- **`WRITES <since>`** returns the files the IDE wrote itself, recorded from `FileControlEvents.AfterWriteFile` for all projects (`Events.FileControlEvents(Nothing)`; signature `VBProject As VBProject, FileType As vbext_FileType, FileName As String, Result As Integer`, from the coclass's source interface `_dispVBFileControlEvents` in VB6EXT.OLB; the similarly named `_FileControlEvents` has a 2-parameter version that VB rejects). It keeps the last 200 writes. Payload: `next US (path US tick) RS ...`; pass `next` as `since` on the following call. A file whose time stamp changed without appearing here was changed outside this IDE.
- **`RELOAD <project> <component> <mode>`** queues `VBComponent.Reload` (spike 8), which refreshes code and designer.
  - `clean`: the automatic reload of an unmodified component. Refused (409) if the component has unsaved changes, and skipped if it has them by the time the action runs.
  - `discard`: throws away unsaved IDE changes. Only `sync_take_disk` sends it, with explicit confirmation, after keeping a recovery copy.
  - It is deferred because reloading a form can create design-time instances of user controls, which run their code.
  - Design mode only: 409 while running; `Reload` fails in break mode (spike 8).
  - Payload: `QUEUED`, `reload <component>`.

### `APPCLICK <hwnd>`

Clicks a button of the program running in the IDE: `<hwnd>` in hex. Only windows of the IDE's own process (the program being debugged runs inside `VB6.EXE`), and only the classes `ThunderCommandButton`, `ThunderCheckBox`, `ThunderOptionButton` (and their `ThunderRT6` forms) and `Button`. Anything else is 400. The add-in **posts** `BM_CLICK` to the control, so the click and its `_Click` event code run later from the message loop, not under the request handler (hard rule 14). Payload: `clicked <hwnd> (<class>)`. The server verifies the effect afterwards.

### `APPTEXT <hwnd> <text>` (deferred)

Sets the text of a text box (`ThunderTextBox`, `ThunderRT6TextBox`, `Edit`) of the IDE's own process with `WM_SETTEXT`. Deferred because `WM_SETTEXT` is sent synchronously and fires the control's `Change` event (user code). Payload: `QUEUED`, `set text`. The server reads the text back with `WM_GETTEXT`.

### `SNAP <hwnd> <id>` (deferred)

Captures a window of the IDE's own process (a form, a dialog or the IDE main window) with `PrintWindow` into `%LOCALAPPDATA%\vb6mcp\shots\<id>.bmp` (32 bpp BMP, written as `.tmp` and then renamed, so a file that exists is complete). `<id>` is 1-40 characters of `[A-Za-z0-9_-]`. Payload: `QUEUED`, then the full path of the file. The server polls for the file, converts it to PNG and deletes it.

Why the add-in captures: a non-elevated `PrintWindow` on an elevated VB6 fails, and a screen copy only sees what's on top; an elevated `PrintWindow` captures covered windows and windowless Labels (spike 15). It's deferred because painting a form of the running program can run user code (`Paint` events). A non-elevated VB6 is captured by the server directly, without this verb.

### `MOUSE <hwnd> <action> <x> <y> [<x2> [<y2>]]`

Posts mouse messages to a window of the IDE's own process. `x`/`y` are **client coordinates of `<hwnd>`**, signed 16-bit. The server picks the deepest child window under the point and converts the coordinates.

| action | Messages posted | Extra arguments |
|--------|-----------------|-----------------|
| `click` | `WM_MOUSEMOVE`, `WM_LBUTTONDOWN`, `WM_LBUTTONUP` | |
| `dblclick` | move, down, up, `WM_LBUTTONDBLCLK`, up (Windows' own sequence) | |
| `rclick` | move, `WM_RBUTTONDOWN`, `WM_RBUTTONUP` | |
| `drag` | move, down, 10 × `WM_MOUSEMOVE` with `MK_LBUTTON` along the line, up at the end point | `x2 y2`: end point, in the same window's coordinates (may lie outside it) |
| `wheel` | `WM_MOUSEWHEEL` × \|n\|, delta ±120, screen coordinates in `lParam` | `x2`: notches, positive = up |

The messages are posted, so event code runs later from the message loop (hard rule 14). A point over a Label or Image is posted to the window underneath, and VB's runtime routes it to the windowless control. Payload: `mouse <action> <hwnd>`. Errors (400): not a window of this process, an unknown action, a coordinate that isn't a number or is out of range, `drag`/`wheel` without the extra argument.

Verified live: a drag in a PictureBox (exact `MouseDown`/`MouseUp` coordinates), a double-click in a ListBox, and scrolling. **A Label or Image never raises `Click` from posted (or sent) messages** (spike 17); use `MOUSEREAL`. Not yet verified: VB's drag-and-drop (`DragMode`, OLE drag).

### `MOUSEREAL <hwnd> <action> <x> <y> [<x2> <y2>]` (deferred), `MOUSERESTORE`

Real mouse input with `SendInput`, for windowless controls (spike 17). Actions: `click`, `dblclick`, `rclick`, `drag` (with `x2 y2`). Coordinates are client coordinates of `<hwnd>`, which must belong to the IDE's own process.

`MOUSEREAL`:
1. Saves the cursor position and the foreground window, unless a restore is still pending.
2. Brings `<hwnd>`'s top-level window to the front: it attaches to the foreground thread's input, then calls `SetForegroundWindow` and `BringWindowToTop`.
3. Refuses (409) if the form couldn't be brought to the front, or if `WindowFromPoint` at the point isn't a window of this process (hard rule 8).
4. Sends one `SendInput` batch with absolute moves on the virtual desktop and the button events.

`MOUSEREAL` is **deferred**: bringing the form to the front activates it, which raises its `Activate` event, and that is user code (hard rule 14). The request validates its arguments and answers `QUEUED`, `queued`. The steps above run from `WM_DEFER`.

`MOUSERESTORE` puts the cursor and the saved foreground window back. Payload: `restored US outcome`. The outcome is that of the last deferred `MOUSEREAL`/`KEYSREAL`: `sent`, the reason it refused (front window, covered point, short `SendInput` count), or `not run`. The server always calls it about 500 ms after `MOUSEREAL`/`KEYSREAL`: VB processes the input only after the request returns, and the add-in never waits (hard rule 5).

### `LISTSTATE <hwnd>`, `SELECT <hwnd> <index>` (deferred)

`LISTSTATE` reads a list box or combo box of the IDE process in-process (UIPI blocks `LB_*`/`CB_*` from the server). Payload: `kind US count US selected US items`, where `kind` is `list` or `combo`, `selected` is -1 for none, and the items are separated by RS.

`SELECT` sets the selection (`LB_SETCURSEL` / `CB_SETCURSEL`). It then sends the parent `WM_COMMAND` with `LBN_SELCHANGE` / `CBN_SELCHANGE`, which is what a user's click produces and where VB raises `Click`. It is deferred because that is user code. Payload: `QUEUED`, `select <index>`; 400 for an index out of range. The server reads `LISTSTATE` again to verify. Multi-select list boxes aren't handled.

### `KEYS <hwnd> <events>`, `KEYSREAL <hwnd> <events>` (deferred)

Key events for a window of the IDE process; `hwnd` 0 means the program's focused control (`GetFocus`, which works because the program runs on the IDE thread). `events` is a list of tokens separated by `;` (at most 500), built by the server from the `app_keys` syntax (`KeySyntax`):

| Token | Meaning |
|-------|---------|
| `S<0-7>` | modifier state: 1 Shift, 2 Ctrl, 4 Alt |
| `D<vk>` / `U<vk>` | key down / up (virtual key 1-254) |
| `C<1-255>` | a Windows-1252 character |

- **`KEYS` posts:** `C` becomes `WM_CHAR`; `D`/`U` become `WM_KEYDOWN`/`WM_KEYUP` (`WM_SYSKEY*` while Alt is held), with scan code and extended-key flag in `lParam`. VB's own `TranslateMessage` produces the `WM_CHAR` of keys like Enter. Posted messages don't change the key state, so each `S` is **queued as a deferred `KEYSTATE` action** between the posted messages. It sets Shift/Ctrl/Alt in the IDE thread's key state with `SetKeyboardState`, so the program sees the modifiers (`GetKeyState`, the `Shift` argument of `KeyDown`). This needs the queue to run **one action per `WM_DEFER`**, so actions keep their place among posted messages.
  - **After each key down, the rest of the sequence is queued as a deferred `KEYSCONT`.** VB's loop runs `TranslateMessage` on the key down, and the `WM_CHAR` it makes (Backspace, Enter, Ctrl+letter) goes to the end of the queue, behind characters posted earlier. In the live test, "ab{BACKSPACE}c" gave "ab": the late Backspace removed the "c". The continuation runs after the key down was processed, so it posts behind the generated character.
  - Payload: `keys <n> US hwnd`.
- **`KEYSREAL`** is deferred, because focusing the control raises `GotFocus`/`LostFocus`. It brings the form to the front, focuses the control, and sends one `SendInput` batch: modifiers as key presses, characters as Unicode. It never leaves a modifier held down. Payload: `QUEUED`, `queued US hwnd`; the outcome comes back through `MOUSERESTORE`.

Errors (400): a bad token, not a window of this process. 409: no focused control.
### `SCROLLPOS <hwnd>`

Range and position of a scroll bar control of the IDE process: `min US max US pos` (`GetScrollRange(SB_CTL)`, `SBM_GETPOS`). It is read in-process because UIPI blocks `SBM_*` messages from the non-elevated server (spike 17).

### `SCROLL <hwnd> <v|h> <code> <count>`

Posts `WM_VSCROLL`/`WM_HSCROLL` with `code` (`SB_LINEUP` 0, `SB_LINEDOWN` 1, `SB_PAGEUP` 2, `SB_PAGEDOWN` 3, `SB_TOP` 6, `SB_BOTTOM` 7) `count` times (1-100), followed by `SB_ENDSCROLL`.
- For a scroll bar control (class contains `ScrollBar`), the messages go to its **parent** with the control's handle in `lParam`; that is how Windows notifies scroll bars.
- Any other window gets them itself, for its own scroll bars.

Payload: `scrolled <hwnd> (<class>)`. The server verifies the effect with `GetScrollInfo` before and after.

### `DESIGN <id> <op> <args...>` (deferred), `RESULT <id>`

Form designer operations (Phase 5, spike 23), on the designer of a form, UserControl or property page (`VBComponent.Designer`, which works with the designer window closed).
- **Why deferred:** reading or setting a UserControl's properties runs that UserControl's code at design time, and adding one runs its `InitProperties`. User code never runs under the request handler (hard rule 14). So `DESIGN` validates, queues and replies `QUEUED US id`.
- **Getting the outcome:** the operation runs from a fresh dispatch and stores its outcome under `id`, a server-chosen string of up to 40 characters.
- **Mode:** design mode only. Checked when queued and again when run, via command 228 (End), which is enabled only in run and break mode.

`RESULT <id>` returns one of:
- `pending`
- `ok US dirty(0/1) US payload`
- `failed US <number>: <message>`
- `unknown`

A finished result is removed once read. At most 100 uncollected results are kept.

| op | args | payload |
|----|------|---------|
| `CONTROLS` | `project component` | records `name US index US progId US className US controlType US container US left US top US width US height`, walking `ContainedVBControls` from the form down. `index` is `-1` outside control arrays; `container` is empty for the form; controlType 1 light, 2 standard, 3 container (VB reports Label as 2). |
| `PROPS` | `project component control` | records `name US kind US value`; kind `s`/`n`/`b`, `o` object, `i` indexed (not read), `e` read error. An object property's `Value` is a VBIDE `Properties` collection (spike 23), so its scalar members follow as `Parent.Member` records (`Font.Bold`). |
| `SET` | `project component control property value` | the property read back as `PROPS` records (all members for `Parent.Member`). The text is converted to the property's current type. Numbers must be `-12`, `3.5` or `&H8000000F&`; anything else is 400. |
| `ADD` | `project component class name container geometry` | the new control's `CONTROLS` record. `class` as a `.frm` writes it (`VB.CommandButton`, `MSComctlLib.ListView`, not the registered ProgID). `container` is empty for the form. `geometry` is `left,top,width,height`, any part empty. |
| `REMOVE` | `project component control` | `removed <control>` |
| `SHOW` | `project component [max]` | the designer window's caption (`DesignerWindow.Caption`). Opens the designer window (`VBComponent.Activate`), maximized inside the IDE with `max` (`DesignerWindow.WindowState`). Deferred like the others, because a form hosting a UserControl runs its code when the designer opens. Used by `form_screenshot`. The caption has no project prefix (`frmMain (Form)`, spike 28). |

- `control` is `Name` or `Name(index)`, or the component's own name for the form (`VBComponent.Properties`).
- **Indexed properties** (List, ItemData): `PROPS` gives their items LF-separated, read with `IndexedValue(0..1999)`. Past the end VB returns ""/0, not an error, so the server trims (spike 25). `SET` takes LF- or CRLF-separated items and sets `IndexedValue(j)`.
- **Pictures:** `SET … <prop> @picture:<path>` loads a picture file into an object property (`Set Property.Object = LoadPicture(path)`). The path is checked against `roots.txt` here too: absolute, no `..`. `@picture:` with no path clears it. VB writes it to the `.frx`.
- Values are formatted locale-independently (`Str$`, so `8.25` never `8,25`).

### `REFS <project>`, `REFADD ...`, `REFDEL <project> <name>`, `TOOLBOX <project> <progId> [<file>]`

References and toolbox components (Phase 5). These are synchronous, because no user code runs.
- **Mode:** `REFS` works in every mode; the others need design mode (409).
- **Reference record:** `name US guid US major US minor US fullPath US builtIn US isBroken US type US description`, where type 0 is a typelib and 1 a project reference. A broken reference has empty fields.
- **`REFADD <project> file <path>`** calls `References.AddFromFile`. For a DLL that isn't registered, VB adds the reference with an empty `fullPath` and description and doesn't register it (spike 23).
- **`REFADD <project> guid <{libid}> <major> <minor>`** calls `AddFromGuid`. The GUID comes from the registry or a typelib (hard rule 13); the server checks it is registered.
- Both return the new record. A duplicate gives VB's own error 32813, "Name conflicts with existing module, project, or object library".
- **`REFDEL`** returns `removed <name>`. Built-in references are refused (409).
  - The name may be `#<n>`: the n-th reference in `REFS` order, for a broken reference whose name can't be read.
  - VB removes a broken reference only while its typelib is registered; otherwise it raises "Object library not registered" (spike 25).
- **`TOOLBOX`** calls `VBProject.AddToolboxProgID` with the **registered** ProgID (`MSComctlLib.ListViewCtrl.2`), which adds the OCX (an `Object=` line). It returns `added <progId>`. VBIDE can neither list nor remove toolbox components.

### `REGSVR <path> <r|u>`, `REGSTATUS <pid>`

COM (un)registration (Phase 5, spike 24).
- **`REGSVR`** starts the 32-bit `regsvr32 /s [/u] "<path>"` as a child. It inherits VB6's elevation, so there's no UAC prompt. `GetSystemDirectory` in the 32-bit VB6 is redirected to SysWOW64.
  - The add-in checks the path itself (hard rule 8): absolute, no `..` or quotes, `.dll`/`.ocx` only, under `roots.txt`, file exists.
  - Payload: the child's pid. The add-in keeps the process handle.
- **`REGSTATUS <pid>`** returns `running`, `exit US <code>` (then the handle is closed), or `unknown` for a pid it didn't start. It is non-blocking (`GetExitCodeProcess`).
- **Errors:** 400 for a refused path, 404 file not found, 409 if regsvr32 couldn't start.
- **Server side:** the server audits the call, polls `REGSTATUS`, and verifies the result in the registry.

### Projects: `CLOSE`, `COMPADD`, `SAVECOMP` (deferred), `PRJGET`, `PRJSET` (Phase 7, `ProjectVerbs.bas`)

All need design mode (409), except `PRJGET`.
- **`CLOSE <project> <0|1>`** queues `VBProjects.Remove`.
  - With `0`, a project with unsaved changes is refused (409, the changed components named).
  - With `1` (discard; the server sends it only for an explicit `discard: true`, hard rule 9) the project is removed as it is.
  - `VBProjects.Remove` saves changed components without asking, and setting `IsDirty` to False saves too (spike 31).
  - So the server snapshots the affected files first and restores them after the close.
  - The add-in logs `closed <project>` or `close failed <project>: <error>`; the server watches `PROJECTS`.
- **`COMPADD <project> new <kind> <name>`** queues `VBComponents.Add` and sets the name. kind: `module`, `class`, `form`, `mdiform`, `usercontrol`, `propertypage`. The name must be a valid, unused VB name (400/409).
- **`COMPADD <project> file <path>`** queues `VBComponents.AddFile`. The path must be absolute, without `..`, under the roots, and a `.bas/.cls/.frm/.ctl/.pag/.dsr`.
  - Deferred because adding a form runs the code of UserControls on it (hard rule 14).
  - The add-in logs `added ...` or `add component failed ...`.
- **`SAVECOMP <project> <component> [<path>]`** queues `VBComponent.SaveAs`, which saves one component without File > Save Project, so the `.vbp` isn't rewritten (PLAN B8; quiet, verified in spike 31).
  - Without a path it saves to the component's own file.
  - With a path it saves to that **new** file, for a component that was never saved. The file must not exist (409); the path must be absolute, without `..`, and under the roots.
- **`PRJGET <project>`** returns `name US description US helpFile US helpContextId US startMode US buildFileName US startup US type`. `startup` is a component name, `Sub Main` or `(None)`.
- **`PRJSET <project> <property> <value>`** sets one of `name`, `description`, `helpfile`, `helpcontextid`, `startmode` (0/1), `buildfilename`, `startup`, and returns the `PRJGET` record.
  - Synchronous: no user code runs.
  - `startup` sets `VBComponents.StartUpObject`: `vbext_so_SubMain`, `vbext_so_None`, or the component (falling back to its name). Which form VB accepts is checked live.
- **Not here: conditional compilation constants and command-line arguments.** VBIDE has no property for them; the server edits them in the `.vbp` of a closed project.

## Watchdog (`vb6mcp watchdog`)

The elevated sidecar (PLAN.md §3.1, spike 12) speaks the **same protocol**: framing, `VB6MCP1` magic, a 128-bit token, `OK`/`ERR`. Its discovery file is `%LOCALAPPDATA%\vb6mcp\watchdog.json`, with the same fields as an add-in's (`addin` is `"watchdog <version>"`). It is written atomically and deleted on a clean exit.
- It listens on 127.0.0.1 only, on a random port.
- There is one watchdog per user session (mutex `Local\VB6McpWatchdog`). It exits 2 minutes after the last VB6 is gone.
- Unlike the add-in it serves from its own threads, so it answers while an IDE thread is busy or hung.
- Every action is limited to VB6 processes: a pid must be a running `VB6.EXE`, and a window handle must belong to one. Actions are logged to `%LOCALAPPDATA%\vb6mcp\watchdog.log`.

| Verb | Arguments | Does | Payload |
|------|-----------|------|---------|
| `PING` | | | `pid US version US elevated(0/1) US parentPid US uptimeSeconds` |
| `BREAK` | `pid` | drops a busy loop into break mode with injected Ctrl+Break (spike 12) | `break US front` (`vb6` or `ghost`); 409 if the IDE isn't in run mode, VB6 can't be brought to the front, or no break followed within 3 s |
| `PRESS` | `hwnd` | posts `BM_CLICK` to a dialog button | `pressed`; 409 if not a `Button` |
| `KILL` | `pid force` | ends VB6; the literal `force` is required (hard rules 9, 12) | `killed` or `kill-sent` |
| `QUIT` | `pid` | posts `WM_CLOSE` to the IDE's `IDEOwner` window (the main window `wndclass_desked_gsk` only if there is none), as File > Exit would. Posted to the main frame, it destroyed only the frame and the process stayed alive without UI (tested 2026-10-17); VB still asks about unsaved changes (the server only sends it when nothing is unsaved). Added 2026-10-17 (`ide_quit`, PLAN.md C7); an older watchdog answers 400 unknown verb | `close-posted`; 409 when the process has no main window |
| `MAKE` | `vbp [defines]` | starts `VB6.EXE /make "vbp" /out "log" [/d defines]` (no window) for `build_make` when no IDE runs (PLAN.md D1). Checked like the add-in's MAKE (hard rule 8): absolute `.vbp` without `..`, under the roots of `roots.txt` next to the watchdog exe (installed, admin-only) or `%LOCALAPPDATA%\vb6mcp\roots.txt` (development), defines `name=integer` pairs joined by `:`. Log in `%LOCALAPPDATA%\vb6mcp\builds\`. Added 2026-10-17 | `pid US log US vbp`; 400 bad arguments, 409 outside the roots or not found |
| `SNAP` | `hwnd id` | PNG to `%LOCALAPPDATA%\vb6mcp\shots\<id>.png`. `PrintWindow` with a 3 s timeout, then a copy of the screen area | `path US method` |
| `LAUNCH` | `[project]` | starts `VB6.EXE` (path from the registry), optionally with a full `.vbp`/`.vbg` path. Elevated like the watchdog, so no UAC prompt (spike 12b) | the new pid |

`BREAK` sends keyboard input, which goes to the window in front. Ctrl+Break in a console window would kill the program running there. So the watchdog injects only while a VB6 window is in front. Windows' `Ghost` stand-in for an unresponsive window also counts.
- If VB6 is already in front, typically the debugged program's form, nothing is changed.
- Otherwise the watchdog brings the IDE forward, attaching to the foreground thread's input as in spike 17. It **never attaches to a VB6 thread**. In the first live test the program's form was in front; attaching to its busy thread and calling `SetForegroundWindow` blocked the watchdog. The foreground change also runs on its own thread with a 1 s timeout.
- Afterwards the previous foreground window is restored, unless it was VB6.

## MCP side: requests the server sends to the client

The MCP connection (stdio, JSON-RPC 2.0, one message per line; spike [1](spikes/01-mcp-handshake.md)) is client to server, with one exception. `debug_eval` asks the user through **elicitation** before it runs a side effect for a client other than Claude Code (spike [40](spikes/40-other-harnesses.md)):

- From `initialize`, the server keeps `clientInfo.name` and whether `capabilities.elicitation` offers form mode: `{}` (2025-06-18, and "form" in 2025-11-25), or an object with a `form` key. A client offering only `url` counts as no elicitation.
- Request, sent while the `tools/call` is still open. Ids are strings `vb6mcp-<n>`, so they can't clash with the client's ids:
  ```json
  {"jsonrpc":"2.0","id":"vb6mcp-1","method":"elicitation/create","params":{
    "message":"An agent wants to run this in the VB6 IDE's Immediate window. ...\n\n? <expression>\n\nIt looks like it has a side effect (<category>). Run it?",
    "requestedSchema":{"type":"object","properties":{"approve":{"type":"boolean","title":"Run it","description":"..."}},"required":["approve"]}}}
  ```
- Response: `{"jsonrpc":"2.0","id":"vb6mcp-1","result":{"action":"accept","content":{"approve":true}}}`. Any other `action` (`decline`, `cancel`), `approve` not `true`, a JSON-RPC error, or no response within 5 minutes refuses the expression.
- A message without `method` is a response. It is matched to the waiting request by `id`; a response with an unknown id is logged and dropped. The read loop keeps reading while a tool waits, so `ping` and other calls are still answered.

## Versioning

Breaking changes bump the magic (`VB6MCP2`) and the discovery `protocol` field. A server talking to an add-in with a different protocol version reports it and doesn't send requests.

## Change log

- **v1** (2026-10-13): initial version (Phase 1).
- v1, 2026-10-13: added `APPCLICK`, `APPTEXT`, `SNAP`, `MOUSE`, `SCROLL`, `SCROLLPOS`, `MOUSEREAL`, `MOUSERESTORE` (Phase 3), `CALLSTACK`, `LISTITEMS` (Phase 2), `LISTSTATE`, `SELECT`, `KEYS`, `KEYSREAL` (Phase 3), `MAKE`, `MAKEIDE`, `MAKESTATUS`, `WRITES`, `RELOAD` (Phase 4; `MAKE` takes a `.vbp` path, for closed projects). `MOUSEREAL` became deferred (`QUEUED`) and `MOUSERESTORE` returns the outcome. Deferred actions run one per `WM_DEFER`. New verbs only: older servers never send them, and a new server talking to an older add-in gets 400 "unknown verb".
- v1, 2026-10-14: added `DESIGN` (deferred, with request ids) and `RESULT`, `REFS`, `REFADD`, `REFDEL`, `TOOLBOX` (Phase 5). New verbs only; against an older add-in the server says to rebuild it.
- v1, 2026-10-15: added `REGSVR`, `REGSTATUS` (Phase 5 round 2). New verbs only.
- v1, 2026-10-17: added `CLOSE`, `COMPADD`, `SAVECOMP`, `PRJGET`, `PRJSET`; `MAKE` takes an optional `defines` (Phase 7). New verbs and an optional argument only; against an older add-in the server says to rebuild it.
- v1, 2026-10-15: `REFDEL` accepts `#<n>`; `DESIGN SET` accepts `@picture:<path>` and indexed properties; `DESIGN PROPS` returns indexed items (spike 25). Older servers never send these forms.
