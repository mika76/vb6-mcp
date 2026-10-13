# Spike 17: clicking windowless controls (Label, Image)

**Question.** `app_mouse` posts mouse messages to the window under a point. That works for windowed controls: a drag in a PictureBox fires `MouseDown`/`MouseMove`/`MouseUp` with the exact coordinates, and a double-click in a ListBox fires `DblClick` (Phase 3 live test). A click over a **Label** reaches the form, but VB never raises `lblClick_Click`. What does VB need?

**Setup.** `tests/fixtures/Ui` running in the elevated IDE. `lblClick_Click` writes its click count into the form caption, so the effect can be read with `GetWindowText`, because a Label has no window. `spikes/17-lightweight-click` is a console program, run **elevated** like VB6, that tries one message sequence per variant and compares the caption before and after. Coordinates: (60, 228) in the form's client area, inside the label.

**Result** (2026-10-13, Windows 11, VB6 SP6, elevated IDE):

| Variant | Click raised? |
|---------|---------------|
| A: posted `WM_MOUSEMOVE`, `WM_LBUTTONDOWN`, `WM_LBUTTONUP` (what `MOUSE` does) | no |
| B: posted `WM_SETCURSOR` before the move and the button-down | no |
| C: sent `WM_NCHITTEST` + `WM_SETCURSOR`, then posted move/down/up | no |
| D: sent move/down/up (`SendMessage`) | no |
| E: real cursor moved onto the label (`SetCursorPos`), then posted move/down/up | no |
| G: form brought to the front (verified with `GetForegroundWindow`), then posted move/down/up | no |
| H: like G, mouse capture checked after the posted button-down | no (capture = 0) |
| I: attached to the form's input thread, `VK_LBUTTON` set down/up in the shared key state around posted down/up | no |
| F: `SendInput` click (real input), form in front, cursor on the label | **yes** |
| a person clicking the label | yes |

**Conclusions.**
- VB6 raises `Click` on a windowless control **only for real mouse input**. This is not about the coordinates (E), the foreground (G), the capture (H) or the button key state (I). Whatever VB checks, only input from the system's input queue has it.
- Windowed controls don't need real input. Their events fire from posted messages, and posted messages don't disturb the user.
- With the form in the background, F's first run had no effect: a background console isn't allowed to take the foreground, so the click went to whatever window was on top. Attaching to the foreground thread's input (`AttachThreadInput`) and then calling `SetForegroundWindow` works without synthetic keys. An Alt keypress also works, but it sends input to the user's current application.

**Decision.**
- `app_mouse` keeps posting by default.
- `real_input: true` uses new add-in verbs:
  - `MOUSEREAL` brings the form to the front, refuses if the point isn't on a window of the IDE process, and sends the input with `SendInput`.
  - `MOUSERESTORE` puts the cursor and the previously active window back. The server calls it about 400 ms later, after VB has processed the input. The add-in can't wait itself, because it must not block the IDE thread (hard rule 5).
- Real input is explicit, because it takes the foreground and moves the user's mouse for a moment.

**Side finding (scroll bars).** From the medium-integrity server, `SBM_GETPOS` to an elevated VB scroll bar control fails, and `GetScrollInfo(SB_CTL)` returns success with nothing filled in. UIPI blocks the `SBM_*` messages those calls rely on. A window's own scroll bars (`SB_VERT`/`SB_HORZ`, e.g. a ListBox) read fine. So scroll bar controls are read by the add-in (`SCROLLPOS`).
