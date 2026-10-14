# Spike 24: COM registration, typelib reading and GUIDs without compatibility (Phase 5 round 2)

**Date:** 2026-10-15. **Machine:** the owner's VM (VB6 SP6 elevated, server medium). Scratch copies of `tests/fixtures/ComDll` and `scratch/ComCompat`; the owner's workspace (12 ActiveX projects) for the read tools.

**Questions:**
1. Can the medium server learn the outcome of an elevated child process?
2. Does registration through the add-in work and can it be verified?
3. Is the add-in's own boundary enough when the server's checks are bypassed?
4. What really happens to GUIDs when a DLL without binary compatibility is rebuilt?

## Results

**Elevated children from medium.** `Get-Process -Id <watchdog pid>` from a medium shell opened the elevated watchdog (`HasExited`, `Handle` worked). But `regsvr32 /s` can finish before the server even sees its pid, so the add-in keeps the process handle and answers `REGSTATUS <pid>` with `GetExitCodeProcess`, which doesn't block (hard rule 5).

**Concurrent typelib reads (bug found).** Two tools reading the same typelib at once failed with "COM object that has been separated from its underlying RCW". OLE caches loaded typelibs, so both reads shared one COM object, and `Marshal.ReleaseComObject` on one thread released it under the other. Fix: one lock for all typelib reads (`TypeLibInfo.ComLock`), and no `ReleaseComObject`; the GC releases them. After the fix, six tools sent together all succeeded.

**ActiveX controls aren't "creatable".** A UserControl's coclass in an OCX (XmText.ocx) has `TYPEFLAG_FCONTROL` but not `TYPEFLAG_FCANCREATE`, yet it is registered like a creatable class. Registration checks include control classes (`TlbType.Control`).

**Register and unregister** (scratch `Reg1\ComFixture.dll`):

| Step | Outcome |
|------|---------|
| `com_where` before | typelib and class `clsGreeter` not registered |
| `com_register` | exit 0; the registry shows typelib `{9FAE05B6-…}` 1.0 and the class (ProgID `ComFixture.clsGreeter`) pointing at this file; `verified: true` |
| `com_unregister` on `ComCompat\ComFixture.dll` (same CLSIDs, a different copy) | **refused**: its classes are registered to another file, and `force` wasn't given |
| `com_unregister` with `confirm: false` | refused |
| `com_register` on `C:\Windows\SysWOW64\scrrun.dll` | refused by the server: outside the roots |
| raw `REGSVR` straight to the add-in (server bypassed): system DLL / `..` path / `.cls` file | **400 from the add-in itself**: not under roots / absolute without `..` / only .dll and .ocx |
| `com_unregister` on Reg1 | exit 0; typelib and class gone; `verified: true` |

**Rebuild without binary compatibility** (scratch `NoCompat`, `CompatibleMode=0`, same source, `build_make` twice):

| Build | LIBID | CLSID of clsGreeter |
|-------|-------|---------------------|
| 1 | `{0B516EAB-EE67-43F4-AF74-36AE2A3D7516}` | `{01291818-A5AE-429A-8394-3E772D79FBEA}` |
| 2 | `{3AFF3179-EE79-4B16-A573-D089E146B326}` | `{AB4DF5EF-73A2-408D-ACC6-938D4D0F40FA}` |

- After build 2, build 1's CLSID and LIBID were **no longer registered**: `/make` unregistered the old binary before registering the new one.
- So a program compiled against build 1 can't find its class any more. That confirms `compat_status`'s warning for projects without binary compatibility.
- The test DLL was unregistered afterwards.

**Read tools on the owner's workspace:**
- `compat_status`: 12 projects, 0 errors, 20 warnings. Most of the owner's own controls (Xm* here) use their build output as their own compatibility file, and are registered from an installed copy under `C:\Program Files (x86)\Common Files\Acme\Controls\`, not from the build output.
- `com_where` on `XmText.ocx`: built by `XmText.vbp`; its typelib and its control class are registered to that installed copy.

## Decisions

- Registration runs only through the elevated add-in (`REGSVR`/`REGSTATUS`). Both sides check the path (the server: roots; the add-in: roots, absolute, no `..`, .dll/.ocx). Each call is audited, and `verified` comes from the registry, not the exit code.
- `com_unregister` refuses when the file's classes point at another copy, unless `force` is given (hard rule 9).
- Typelib reads are serialized and never force-released.
