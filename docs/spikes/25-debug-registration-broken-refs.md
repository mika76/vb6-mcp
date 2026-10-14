# Spike 25: debugger registrations, broken references, pictures and lists (Phase 5, last round)

**Date:** 2026-10-15. **Machine:** the owner's VM (VB6 SP6 elevated). Scratch projects only. The owner agreed to the kill test.

## 1. ActiveX projects running in the IDE (PLAN §9.2.5)

Scratch `DbgReg`: the ComFixture DLL with binary compatibility, its reference copy holding CLSID `{AB4DF5EF-…}`, LIBID `{3AFF3179-…}`.

| Step | Registry for the class / typelib |
|------|----------------------------------|
| before (not registered) | nothing |
| first Start | VB asked first: SourceSafe prompt, then the project's Debugging options ("Wait for components to be created", OK) |
| running | class → `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6DEBUG.DLL`; typelib → `D:\…\DbgReg\ComFixture.vbp` |
| End | both gone again: the earlier state, i.e. not registered |
| registered to `NoCompat\ComFixture.dll` (same CLSID), then Start | VB6DEBUG.DLL / `.vbp` again |
| End | **restored** to `NoCompat\ComFixture.dll` |
| Start, then `ide_kill force` | **left behind**: class → VB6DEBUG.DLL, typelib → `.vbp` |
| after the next VB6.EXE starts (the add-in build's `/make`, then the IDE) | restored to `NoCompat\ComFixture.dll` |

So VB keeps the previous registration and puts it back on End, and on its next start after a crash. Until then, client programs get the debugger stub with no IDE behind it.

**Decisions:**
- `com_orphans` lists debugger registrations (classes on VB6DEBUG.DLL, typelibs on a `.vbp`). It says whether the project is being debugged now or the entry is left over, and gives the repair: `com_register` of a build carrying that typelib, found in the workspace.
- `com_diagnose` reports it for 429.
- `ide_kill` warns when it kills an IDE in run or break mode.

Also found: `run_start` returned while VB's dialogs were open, so the server didn't record the run as its own and `run_end` asked for `force`. Fixed: a start that stopped at a dialog counts once the run begins after `ide_dialog_respond`.

## 2. Broken references

| Tried | Outcome |
|-------|---------|
| Form1x references ComFixture `{9FAE05B6}` 1.0, **not registered**, hint `..\ComCompat\ComFixture.dll` exists | **not broken**: VB loaded the typelib from the hint path (FullPath empty). The hint is a fallback when the registry has nothing |
| the same project one folder deeper (the hint leads nowhere) | broken: `IsBroken`, **Name unreadable** (empty), but Guid and version readable |
| `References.Remove` on it (`REFDEL #5`) | **error "Object library not registered"** (-2147319779) |
| `AddFromFile` of a file with the same LIBID while the broken one is there | added as a second reference; the broken one stays, and Remove still fails |
| after `com_register` of a file carrying the LIBID | Remove works |
| fresh broken copy, `com_register` of a carrier, **no remove** | **the reference healed by itself**: VB re-resolved it as soon as the typelib was registered |

**Decisions:**
- `refs_fix_broken` finds files carrying the exact LIBID and major version: the hint path under the roots, and workspace build outputs and compatibility copies. Its first fix is `com_register` of one of them.
- `refs_remove #position` is the fallback.
- With nothing to register, it says VBIDE can't remove the entry.

## 3. Pictures and lists in the designer

| Tried | Outcome |
|-------|---------|
| `Set Property.Object = LoadPicture(ico)` on the form's Icon | works; the handle changed |
| the same with a jpg on Picture | works; Type 1 (bitmap). Saved as a 6.9 MB `.frx`: VB stores it uncompressed |
| a picture outside the workspace roots | refused by the server; the add-in checks too |
| `IndexedValue(j) = item` for j = 0..2 on a new ListBox's List | works (3 items, `é` kept) |
| reading `IndexedValue(j)` past the end | **returns "" (List) / 0 (ItemData), no error**: the reader can't find the end by error. The server cuts List after its last non-empty item and ItemData to the same count |
| save | `Icon`, `Picture`, `List` and `ItemData` written to the `.frx` by VB itself (we never edit `.frx`, hard rule 2) |

## 4. Orphan scan noise

Registry paths don't always name a file directly:
- Bare system names (`combase.dll`) are found through the system folders.
- Typelib paths may end in `\<n>`, the resource index (`vbscript.dll\2`).
- Typelibs with no 32-bit file are 64-bit only, not broken.

After handling these, the owner's machine shows 137 classes and 15 typelibs with really missing files, as a 32-bit program sees them. Examples are Codejock OCXs registered from a deleted `Desktop\vs6installer4` folder, and 64-bit-only system files registered under System32.
