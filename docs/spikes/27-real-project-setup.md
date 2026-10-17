# Spike 27: setting up a real project on a fresh machine

**Date:** 2026-10-15. **Machine:** the owner's VM.

**Project:** a production project from the owner's workspace, copied to `D:\bench` as a test bed. It has 60 components (50 shared from sibling folders), 10 references and 14 OCX components.

This was the first real (not fixture) project used end to end. Getting it to build on the VM surfaced machine setup problems, VB behaviour, and two gaps in the tools.

## Steps and findings

1. **Shared sources:** 55 files from 7 sibling folders (`..\Common`, `..\Other`, …) had to be copied, keeping the same relative layout. They were found from the `.vbp` member lines. `.frm` files bring their `.frx`. Folder-name case in the `.vbp` (`Common`/`COMMON`) doesn't matter.
2. **`com_diagnose` on the project:** all references and components were registered, with their files present, so nothing COM-related needed copying.
3. **First build:** "frmMain.frm could not be loaded". The reason was only in `frmMain.log` next to the form: `Cannot load control WinHook1; license not found`.
   - `WinHook1` is Desaware SpyWorks 6 (`DWSHK36.OCX`), a **licensed** control. The registered OCX alone is the runtime part. Design time and compiling need the **developer licence**, and 25 of the owner's 43 projects use SpyWorks.
   - The owner's dev machine had the full SpyWorks install; the VM only had runtime files.
   - Guessing the licence key from GUIDs in the OCX was **wrong**: the one non-class GUID wasn't the licence key on the dev machine either.
   - Installing SpyWorks 6.3 from the owner's media added three `HKCR\Licenses` keys (`2A318503-8F9E-11D2-B89F-00AA00A8122A`, `CB300519-620F-11D2-8B8B-00AA00580C40`, `DesawareConstituentOCX`; diffed against a snapshot taken before). After that every form loaded.
4. **Second build:** compile error "Can't find project or library" at `Space(16)`, a built-in function. That error appears when a reference can't be loaded.
   - `refs_fix_broken` in the IDE named it: `Vendor_License_COM`, asked for as `1.a` (= 1.10, **hex**; this settles the open question from Phase 4) and registered as `1.14` (= 1.20).
   - **The registered file contained no type library:** `TypeLib\{…}\1.14\0\win32` pointed at `Vendor.License.COM.dll`, a 12 KB .NET COM wrapper, registered without a `.tlb`.
   - `com_diagnose` had reported this reference as "ok", because it only checked that the file exists. Fixed: it now also loads the type library. `refs_fix_broken` reports `hasTypeLib` and suggests `RegAsm <dll> /tlb` for this case.
5. **Copying the `.tlb` to the `.vbp`'s hint path** (`D:\Licensing\…`, where the dev machine has it) did **not** help:
   - VB stopped flagging the reference as broken.
   - But it stayed bound to the registered DLL, and the next compile failed on `Vendor_License_COM.COMValidator` ("User-defined type not defined").
   - **The hint path is a fallback only when the type library isn't registered at all** (spike 25). A registered-but-empty entry wins.
6. **The fix:** `RegAsm Vendor.License.COM.dll /tlb` (elevated, run by the owner) wrote the `.tlb` next to the DLL and pointed the registration at it.
   - The project, already open in VB6, still failed with the same error. **VB binds references when it loads a project**, so a registration change needs the project reloaded.
   - The IDE showed the project as changed (its reference bookkeeping), so `project_open replace` refused, by design (hard rule 4). The owner restarted VB6 without saving.
7. **`build_make` (`/make`)** then built `App.exe` with no errors in 15.7 s.

## Decisions and follow-ups

- **Done:** `com_diagnose` checks that a registered type library file loads; `refs_fix_broken` reports it and suggests `RegAsm /tlb`.
- **Backlog (PLAN.md):**
  - `project_close` with an explicit `discard: true`: real use needed it to reload a project after a registration change. It needs a spike: does `VBProjects.Remove` on a changed project close silently or prompt?
  - `build_make` should read the `<form>.log` files VB names in "Errors during load" and put their lines into the result: the actual reason ("license not found") was only in that file.
- **Machine setup notes for other users (README troubleshooting):** licensed controls need their developer licence, not just the registered OCX. .NET COM references need a registered `.tlb`.
