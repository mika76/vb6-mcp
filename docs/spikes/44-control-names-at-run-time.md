# Spike 44: naming a running form's controls

**Question.** `app_tree` lists the windows of a running form, but Windows knows them only by class. Every UserControl is `ThunderRT6UserControlDC` (`ThunderUserControlDC` before compile), so an agent driving a form full of third-party controls had to match them by position by hand (notes from real use). Can the server give each window its VB name and type?

**What doesn't work.**
- VBIDE: designer and property access in run mode fails with error 393 and can instantiate the form (AGENTS.md).
- The Immediate window (`?Forms(0).Controls(i).Name`) works only in break mode, and a UserControl's extender has no `hWnd`.

**What I tried.**
- I compared a real form's designer block (in the `.frm` on disk) with the window rectangles `app_tree` reported for the running form, at 96 DPI.
- Every control matched exactly once its twips were divided by 15 (1440 / 96). For example, 120,540 6000 x 700 twips is 8,36 400 x 47 pixels, and a 1540,1360 1500 x 315 text control is 103,91 100 x 21.

**Decision.**
- `FrmDesign` (Vb6Mcp.Core) reads the designer block: class, name, `Index`, position, size and `ScaleMode`. Nested `Begin` blocks give containers; `BeginProperty` blocks are skipped.
- `FrmDesign.Match` names windows top-down. A window whose parent is already matched gets the child of that container with the same size (within 1 pixel) and the nearest position (within 3 pixels: PictureBox and Frame borders).
- Units:
  - twips at the system DPI (`GetDpiForSystem`), or pixels when the container's `ScaleMode` is 3;
  - a Frame has no `ScaleMode`, so its children use its container's units;
  - other `ScaleMode`s aren't matched.
- The form is the `.frm` of the IDE's projects that names the most windows. The add-in lists the components; the files are cached by write time.
- **Verified live** on a form with 20 named windows: ActiveX text, check, combo and option controls (three of them inside a Frame), a data list, a command panel, and intrinsic list boxes. The inner windows of each UserControl (its own TextBox or OptionButton) correctly stay unnamed.

**Limits, said in the result** (`namesFrom`):
- Names come from disk: unsaved designer changes aren't there yet.
- Controls that code moves or resizes at run time (Form_Resize layouts) don't match and stay unnamed.
- Compiled programs (`exe_run`) aren't named: the server doesn't know which sources built them.
