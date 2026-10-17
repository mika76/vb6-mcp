# Spike 29: VB6 designer rules for agents, and a lost-save bug

**Question:** benchmarks T8 and T9 showed agents misjudging what a designer change does. Which VB6 designer rules are true on this machine, so the tools can tell agents at the moment they matter? And can a change that VB misses be saved?

**Setup:**
- A production project (the T9 run 1 copy) open in the elevated IDE, changed through `form_set_property` and `form_set_properties`.
- Results read back with `form_properties`, `project_components` (`IsDirty`), `form_screenshot` and the saved `.frm`.
- The copy's files were backed up first and restored afterwards.

## Results

1. **A form's `Font` doesn't change the controls already on it (through VBIDE).**
   - Setup: `frmRepairCard` used Trebuchet MS. `lblstatus`, `txtStart` and `cmdTest` had no font block of their own, `lblCardNumber` had an explicit Trebuchet block and `cmdserial` had MS Sans Serif.
   - Setting the form's `Font.Name` to Segoe UI left all five unchanged.
   - In T8, VB then wrote the unchanged controls' Arial explicitly when saving (33 new font blocks).
   - Not tested: changing the font by hand in the Properties window's font dialog, which the owner remembers carrying over to controls that still use the form's font. The rule is stated for what our tools do.
2. **A change made only to an object property's members isn't counted as a change. Save Project then skips the form and the change is lost.**
   - Setting `lblStart`'s `Font.Bold` on a clean `frmAddCards` left `IsDirty` False, though `form_properties` read back `True`.
   - **File > Save Project** (what `project_save` runs) saved the other changed forms but not `frmAddCards`, and the `.frm` on disk was unchanged.
   - Setting `lblStart`'s `Tag` to its own value ("") made `IsDirty` True. The next save wrote `Weight = 700` for `lblStart`.
   - A scalar change (a `Caption`) is counted as usual.
   - In T8 and T9 this went unnoticed because every form also had scalar changes.
3. **A CommandButton shows `BackColor` only with `Style = 1` (Graphical).**
   - Red `BackColor` on `frmFilter`'s Clear button with `Style = 0`: the screenshot showed it grey.
   - With `Style = 1`: red.
4. **A CommandButton has no `ForeColor`:** setting it fails with "438: Object doesn't support this property or method". In T8 an agent concluded the same itself, after trying.
5. **Overlaps:**
   - T9 run 1's `cmdFixAccount` and `cmdRefund` were at the same rectangle in `Frame3` (`8250, 300, 1260x975`). `form_controls` gives each control's container and rectangle, enough to find sibling overlaps.
   - The same form also has a control array whose two members (`imgStatusIcon(0)` and `(1)`) sit on the same spot on purpose.
6. **Reloading a designer re-runs its UserControls:** after its files were restored on disk, the IDE reloaded `frmMain`, which showed the Login Error box again. A designer call made right then failed on the dialog.

7. **`CONTROLS` fields:** the `.frm` class is the progId field (`VB.Label`); the class field is the short name (`Label`). The first overlap filter compared the wrong field and reported labels (benchmark T10).
8. **Labels under windowed controls are hidden:** in T10, header labels placed over the area of frames were partly covered. Label on label and 15-twip edge contacts were noise.

## Decision

- **The tools repair the lost-save bug:**
  - When a change touched only object members (`Font.*`) and VB still reports the form unchanged, `form_set_property` and `form_set_properties` set the form's `Tag` to its own value, so VB counts the change.
  - If that isn't possible, the result says so.
  - No add-in change was needed, so it works with installed add-ins.
- **Hints in results, only for the rules verified here:**
  - after a form `Font.*` change (rule 1);
  - for `BackColor` on a CommandButton whose `Style` isn't 1 (rule 3);
  - after any move or added control, "Check the layout with form_screenshot".
- **Error text:** VB's 438 on a button's `ForeColor` gets an explanation (rule 4).
- **Overlaps:** after a geometry change (`Left`, `Top`, `Width`, `Height`, `Visible`, `Container`) or `form_add_control`, the result lists sibling controls in the same container that overlap a changed control. Not checked:
  - label on label, Images, Shapes and Lines, which are meant to sit on others (a Label on a windowed control is reported: it's hidden);
  - overlaps under 30 twips (2 pixels);
  - a Label that wholly contains the other control (a background band, T12).
  - timers;
  - members of the same control array.
- **Checks that can't run** (a dialog, a busy IDE) say so in the hints instead of disappearing.
- **The same rules go in the tool descriptions, in one sentence:** that's where an agent reads them before acting. A broader guide or skill is a separate decision. These rules are the start of its verified content.
