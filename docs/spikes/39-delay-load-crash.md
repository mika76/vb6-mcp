# Spike 39: an IDE crash with 0xc06d007e (improvement notes 55)

**Question.** In a real-use session (2026-10-18), VB6 crashed while a production program ran a balance check with `nologin = 0`. The Application log had exception `0xc06d007e` in KERNELBASE and nothing else. Which DLL, and why does it take the whole IDE down?

**What the code does.** With `nologin <> 1`, `COMMON\frmLogin.frm` creates the DigitalPersona One Touch objects (`New DPFPCapture`, `DPFPFeatureExtraction`, `DPFPVerification`). The project references `DPFPDevX.dll`, `DPFPEngX.dll` and `DPFPShrX.dll` (`C:\Program Files (x86)\Common Files\Acme\DigitalPersona\Bin\Bin\COM-ActiveX`).

**What I tried.**
1. Read the import and delay-import tables of those DLLs with a small PE reader (a scratch Python script, no tools installed):
   - `DPFPDevX.dll` delay-loads `DPFPApi.dll`;
   - `DPFPEngX.dll` delay-loads `dpHFtrEx.dll` and `dpHMatch.dll`;
   - `DPFPShrX.dll` has no delay imports.
2. All three delay-loaded files exist in `SysWOW64`. But `DPFPApi.dll` there is from 2023 (the U.are.U SDK under `C:\Program Files\DigitalPersona`), while the One Touch files are from 2010. Its normal imports include `DPSCApi.dll`, which is in neither `SysWOW64` nor the folder next to it. A search of `Program Files`, `Program Files (x86)` and `System32` found no copy of it anywhere.
3. `LoadLibraryW` in a 32-bit process (`SysWOW64\WindowsPowerShell`): `dpHFtrEx.dll` and `dpHMatch.dll` load; `DPFPApi.dll` and `DPSCApi.dll` fail with **error 126** (module not found).

**Conclusion.**
- `0xc06d007e` is the code the MSVC delay-load helper raises when the module can't be loaded (`0xc06d007f`: a procedure can't be found).
- The first DPFP call that needs `DPFPApi.dll` raises it inside VB6.EXE. VB's error handling (`On Error`) doesn't catch SEH exceptions from native code, so the IDE process ends, as a compiled EXE would.
- The cause here is the machine, not the code: a newer `DPFPApi.dll` (U.are.U) in `SysWOW64` that needs a `DPSCApi.dll` that isn't installed.
- **Not reproduced in the IDE on purpose** (it would end VB6 again). The chain (delay import, then missing second-level DLL, then error 126) is proven step by step; that the crash came from exactly this call is inferred from the code path and the exception code.

**Decision for vb6-mcp.**
- The Application log's `KERNELBASE` names no DLL. A tool that walks a referenced DLL's imports **and delay imports** and tries `LoadLibraryEx` on each would have named `DPSCApi.dll` before the run.
- **Built 2026-10-19:** `PeDependencies` in `Vb6Mcp.Core` (BCL only) and the dependency checks of `com_diagnose`. Live on a production project's 33 references and controls: one error, `DPFPDevX.dll -> DPFPApi.DLL -> DPSCApi.DLL: not found`, flagged as delay-loaded; nothing else reported.
- A first version walked nothing under the Windows folder and missed it: `DPFPApi.dll` is third-party but sits in SysWOW64. Now only DLLs whose version info names Microsoft are skipped.
- **Built 2026-10-19 (improvement notes 55):** `run_start` runs the same check over every control and reference of the open projects (it can't know which ones the run will use) and reports `missingDlls` without refusing to start.
- Item 25 (2026-10-16, a production project, also `0xc06d007e`) happened while the IDE was idle, so it had another trigger. That entry had the .NET runtime loaded in the process. Not investigated.
