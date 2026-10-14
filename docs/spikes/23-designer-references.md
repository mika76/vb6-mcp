# Spike 23: form designer and references through VBIDE (Phase 5)

**Date:** 2026-10-14. **Machine:** the owner's VM (VB6 SP6, elevated), add-in `Design.bas`/`Refs.bas`, scratch copies of `tests/fixtures/Basic`.

**Question:** can the add-in read and change forms, references and toolbox components through VBIDE, as the IDE holds them, safely? Which VBIDE members actually behave as their typelib suggests?

## Reading the API first

`vb6mcp describe "<VB98>\VB6EXT.OLB" <filter>` (our own typelib reader, `TypeLibDescriber`) printed the exact members before any code was written. These were the relevant ones:
- **Designer objects:** `VBForm.ContainedVBControls`/`VBControls`, with `.Add(ProgId, Optional RelativeVBControl As VBControl = 0, Optional before As Boolean = False)` and `.Remove(VBControl)`.
- **Controls and properties:** `VBControl.Properties`, `.ContainedVBControls`, `.ProgId`, `.ControlType`, `.Container` (settable). `Property.Value`, `.IndexedValue`, `.NumIndices`.
- **References:** `References.AddFromGuid(Guid, Major, Minor)`, `.AddFromFile`, `.Remove`. `Reference.Type` (`vbext_rk_TypeLib` 0, `vbext_rk_Project` 1), `.IsBroken`, `.BuiltIn`.
- **Components:** `VBProject.AddToolboxProgID(ProgId, [FileName])`. There is no member that lists or removes toolbox components.

## Design decision before testing

Reading or setting a property of a UserControl whose project is open calls that UserControl's own `Property Get/Let`, and adding one runs `InitProperties`. That is user code. So all designer operations are **deferred** (hard rule 14): `DESIGN` queues the operation under a request id, and `RESULT` returns the outcome. References run no user code and stay synchronous.

## Results

| Tried | Outcome |
|-------|---------|
| `VBComponent.Designer` with the form's designer window never opened | works; listing, setting, adding and removing all worked without opening it |
| `ControlType` of a Label | 2 (standard), not 1 (light) |
| `Properties("Index").Value` outside a control array | `-1` |
| `Properties("Font").Value` | **not a StdFont**: a VBIDE `Properties` collection (`TypeName` "Properties") with Name, Size, Bold, Italic, Underline, Strikethrough, Weight, Charset. The same holds for DragIcon, MouseIcon and DataFormat. Setting `Value("Bold").Value = True` worked and was saved as `BeginProperty Font ... Weight = 700` |
| String property with `–` and `é` | stored, then saved as 1252 bytes 0x96/0xE9 |
| `Val("abc")` for a numeric property | silently 0, so numbers are now validated (`-12`, `3.5`, `&H…&`) |
| **late-bound** `coll.Add("VB.Frame")` | **error 13 Type mismatch**, nothing added: the omitted typed optional `RelativeVBControl` went over as a missing argument |
| early-bound `ContainedVBControls.Add(progId, Nothing-typed var, False)` | works; also inside a Frame (`fraOptions.ContainedVBControls`) |
| `Add("MSComctlLib.ListViewCtrl.2")` (registered ProgID) | **returns Nothing**, no error |
| `Add("MSComctlLib.ListView")` (`.frm` class name) | works. `Add` takes the `.frm` class name, `AddToolboxProgID` the registered ProgID |
| `Add` of an OCX control whose component isn't in the project | returns Nothing as well |
| Setting `Index` to 0, then finding `Name(0)` | works (control array) |
| `ctl.Collection.Remove ctl` | works; the control's event procedures stay in the code |
| the form itself as a `VBControl` | not one; its properties are `VBComponent.Properties`. With the component's name as control, the form's Caption was set (second round) |
| picture properties (Picture, Icon, MouseIcon, Palette) | also `Properties` collections, with Handle, Type, Width, Height |
| Save after adding a ListView | VB wrote `Object = "{831FDD16-…}#2.2#0"; "mscomctl.OCX"` into the `.frm` header and `Object=…` into the `.vbp`, consistently |
| `AddFromGuid` ADODB 2.8 (GUID from the registry) | added |
| `AddFromFile` on an **unregistered** VB-built DLL | added with empty `FullPath` and `Description`, and not registered. After save: `Reference=*\G{9FAE05B6-…}#1.0#0#..\ComCompat\ComFixture.dll#ComFixture`, a relative path, and **the same LIBID that `LoadTypeLibEx` reads from the file** (this verifies `ws_affected`'s matching against a line VB wrote) |
| adding a duplicate reference | VB error 32813 "Name conflicts with existing module, project, or object library" |
| removing a `BuiltIn` reference | refused by us before VB is asked |
| `AddToolboxProgID "MSComctlLib.ListViewCtrl.2"` | the common controls appear in the toolbox; the component is lost if the project is closed unsaved |

## Decisions

- The designer stays deferred, with `RESULT` polling hidden in the server (`AddinClient.DesignAsync`).
- Object properties are listed and set as `Parent.Member` through the sub-collection. Whole-object (picture) and indexed properties aren't supported yet.
- `form_add_control` documents the `.frm` class name, and an `Add` that returns Nothing becomes a clear error.
- `refs_list` reads components from the `.vbp` on disk, because VBIDE can't list them.
- Not tested: project references inside a group (`*\A`), UserControls from an open project (the deferral exists for them), and the break-compatibility dialog when references change.
