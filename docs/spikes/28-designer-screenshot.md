# Spike 28: screenshots of form designers

**Question:** can the server show an agent what a form looks like in the designer (unsaved changes included) without running the program? Many real forms can't run at all without a database or login (a production project, spike 27).

**Setup:** A production project loaded in the elevated IDE (design mode), with the add-in's new deferred `DESIGN SHOW` op (`VBComponent.Activate`, then `DesignerWindow.WindowState = vbext_ws_Maximize`). The window tree was dumped with `vb6mcp windows <pid>`.

## What happened

- **Window structure** of an open designer, under the IDE main window:
  ```
  MDIClient ""
    DesignerWindow "frmRepairCard (Form)"      <- the MDI child
      ThunderForm "Repair Order"                <- the form surface, its caption is the form's Caption
        ThunderFrame / ThunderTextBox / ThunderCommandButton ...
        ThunderRT6UserControlDC ""             <- a compiled UserControl on the form
  ```
  A code window next to it is `VbaWindow "App - modMain (Code)"`.
- **The designer caption has no project prefix** (`frmRepairCard (Form)`), unlike the code window. Guessing it from `Project - Component (Form)` would have failed. `SHOW` therefore returns `DesignerWindow.Caption` from VBIDE and the server matches it exactly.
  - With two loaded projects that have a component of the same name, both designers would have the same caption. Activate puts ours at the top of the MDI z-order and `EnumChildWindows` walks in z-order, so the server takes the first match.
- **Capture and crop:** the add-in's `SNAP` of the whole IDE main window (spike 15), cropped to the `ThunderForm` rectangle relative to the main window. The images matched the designer exactly (title bar, grid dots, controls, pictures on buttons).
- **Small IDE window:** the MDI area here was 763x436. `frmMain` (666x478) and `frmFixAccount` (398x444) were cut off at the bottom in the first version (see the second round below).
- **UserControl code runs when a designer opens:** opening `frmMain` for the first time showed a production project's "Login Error" message box (a UserControl connecting to the database). The deferred `SHOW` waited behind it, and the request timed out after 20 s. Answering the box (`ide_dialog_respond`) let `SHOW` finish. Reading the same UserControls' properties afterwards didn't show it again, so it happens when the form is first loaded.
- **Opening that designer marked the form changed.** `frmMain` was clean before `SHOW` and dirty (`IsDirty`) once the box was answered, with nothing set by us; the other five forms stayed clean. Presumably a UserControl calls `PropertyChanged` while loading, which a developer opening the form by hand gets too. A changed component can't be replaced by `project_open` without a save or discard, so `form_screenshot` reports `dirty` and says where it may come from.

## Second round: forms bigger than the designer

In benchmark T9 an agent redesigned `frmMain` and said it couldn't see the bottom of the form in `form_screenshot`, so it left that part unchecked. Two ways to get the whole form were tried (`spikes/28-designer-snap/`):

- **Capturing the form window on its own** (the add-in's `PrintWindow` on the `ThunderForm`): the bitmap had the form's full size (688x548), but its content was a misaligned piece of whatever was on top (another designer), as spike 15 found for IDE panes. Rejected.
- **Scrolling the designer:** `WM_VSCROLL`/`WM_HSCROLL` posted to the `DesignerWindow` (the add-in's `SCROLL` verb: a window of the IDE's own process) moved the form inside it: `SB_BOTTOM` moved the `ThunderForm` from y=153 to y=-409, `SB_TOP` brought it back. Capturing the IDE after each page and pasting the visible pieces at the form's offset gave a complete, seamless 688x548 image of `frmMain` from two captures.

Also found:
- **The designer caption depends on the window state:** maximized `frmMain (Form)`, not maximized `App - frmMain (Form)`. Another reason to take it from VBIDE.
- **Crop to the client area:** the designer's window rectangle includes its scroll bars. The first version cropped to it, and a cut-off image showed a strip of scroll bar.

## Decision

- `form_screenshot`: `SHOW` (maximized by default), find the `DesignerWindow` by the caption VBIDE reports, take its largest visible non-scroll-bar child (the `ThunderForm`), and crop to the designer's client area inside the MDI area. A form bigger than that is scrolled through a page at a time and stitched (at most 20x20 captures), then the designer is scrolled back to the top left. `clipped` now only means the designer couldn't scroll far enough.
- All designer ops (`AddinClient.DesignAsync`) now check for a dialog in the VB6 process after about 1 s of waiting, and fail at once with its caption, text and buttons instead of waiting out the timeout. The op finishes by itself once the dialog is answered.
- Not verified: UserControl, property page and MDI form designers (expected to work the same; the caption pattern and surface class are taken from VBIDE and the window tree, not guessed).
