# Spike 20: drag-and-drop and debug output (Phase 3 leftovers)

Date: 2026-10-14. `tests/fixtures/Ui` gained an OLE drag source (`lstItems`, `OLEDragMode` automatic), an OLE drop target (`txtDrop`, `OLEDropMode` automatic), a classic drag source (`cmdDrag`, `DragMode` automatic) dropped on `picDraw` (`DragDrop` event), and `OutputDebugString` in `txtName_Change`. Tested in the elevated IDE and as compiled `Ui.exe` (`exe_run`).

## Debug output

| Test | Result |
|------|--------|
| `debug_output` starts the DBWIN monitor (unit test: own process, ANSI text with accents) | ✓ |
| Program in the **elevated IDE** writes `OutputDebugString` | ✓ captured as process `VB6` (the elevated writer opens the objects the medium server created) |
| Compiled `Ui.exe` writes | ✓ captured as process `Ui` |
| `stop: true` | ✓ buffer released |

## Drag-and-drop

| Kind / where | Posted (`app_mouse drag`) | Real input, one batch | Real input, paced |
|--------------|---------------------------|-----------------------|-------------------|
| Classic (`DragMode`), IDE | ✗ nothing | ✓ "Dropped cmdDrag at (70,58)", exact | not needed |
| Classic, compiled EXE | | ✗ | ✗ (15 ms steps) |
| OLE (`OLEDragMode` → `OLEDropMode`), IDE | ✗ | ✗ | not testable from the add-in (no waiting on the IDE thread) |
| OLE, compiled EXE | | ✗ | ✗ (`SetCursorPos` steps, absolute moves, 1-pixel jiggles, 80 steps over 1.5 s; conditions verified: form in front, button down, cursor over the target at release) |
| **By hand** (the owner) | | | ✓ classic and OLE, EXE |

Findings:
- **The owner saw the cursor barely move**, far from the form, while the program's own `GetCursorPos` reported the target position. That is what a virtual machine with mouse integration does: the host keeps putting the guest's cursor where the real mouse is. Under that hypothesis:
  - A click or a single fast batch completes before the host interferes. That's why clicks, the Label click and the IDE's batch drag worked.
  - A drag spread over time is pulled back partway through, and OLE's `DoDragDrop` loop needs exactly such gradual movement. **Confirmed by the owner: the test machine is a VM, used over RDP.** The RDP client keeps sending the real mouse position into the session. On a physical machine, or at its local console, the paced variants may well work; that is untested.
- **Bringing a compiled program to the front:** `SetForegroundWindow` alone was refused while the user was active in Claude. `SwitchToThisWindow`, retried, got there. `LocalInput` now does that, the same as the watchdog (spike 19).
- Absolute `mouse_event` coordinates from a test process landed elsewhere (813,965 instead of 210,351). That's the same symptom.

## Decision

- **Classic drag:** `app_mouse` `drag` with `real_input: true`. Verified in the IDE; documented as environment-dependent.
- **OLE drag:** no dedicated tool. Synthetic input couldn't start it here, and an untested pacing mechanism in the add-in would need timers on the IDE thread. Agents should use the program's non-drag path, or ask the user. Revisit on a physical machine if a real program needs it.
- `debug_output` as built.
