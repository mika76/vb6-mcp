# T15: binary compatibility (2026-10-18, Sonnet, dry run)

## The task

`tests\fixtures\Compat` (written by `bench\make-compat-fixture.cs`):
- **`PriceLib`:** an ActiveX DLL with binary compatibility. Its compatibility file `compat\PriceLib.dll` is committed. `clsTax.AddTax(net)` adds 20%.
- **`ClientApp`:** an EXE that creates `PriceLib.clsTax` at start and shows `AddTax(100)` = 120 in a text box.

The harness builds both per run, the DLL first, so the run's copy is registered and the client is compiled against it. Then it asks for two changes:
1. Add `RateFor(country)`. That's fine: VB extends the interface compatibly.
2. Give `AddTax` the rate as a second parameter. That breaks binary compatibility, and "ClientApp.exe must keep working without being rebuilt".

A good run adds `RateFor`, keeps `AddTax` as it is (a new method can take the rate), rebuilds the DLL, and **says why** instead of breaking compatibility (hard rule 13).

**Scoring** (`Verify-Task T15`). The run passes only if all of these hold:
- the compatibility settings and the compatibility file are unchanged;
- `ClientApp.exe` wasn't rebuilt;
- the DLL was rebuilt;
- its typelib has `RateFor` and the original `AddTax(ByVal net As Currency) As Currency` (read with `com_describe`);
- the answer names the compatibility conflict;
- the untouched client, run against the agent's DLL (`bench\replay.ps1`), shows 120.

### What each outcome does on this machine (checked before the harness relied on it)

| Change in a scratch copy | `/make` | Old ClientApp.exe |
|---|---|---|
| `RateFor` added, compatibility on | succeeds | shows 120 |
| `AddTax(net, rate)`, compatibility on | **fails**: "compatibility error ... 'AddTax' in the 'clsTax' class module has arguments and/or a return type that is incompatible". `build_make` reports it as `kind: compatibility` with a note that it is the user's decision | (the DLL is unchanged) |
| the same with `CompatibleMode="0"` | succeeds, new CLSIDs | run-time error **"Class not registered"** (80040154). `/make` removed the old class registration (spike 24) |

## Dry run (1 run each)

| | Result | Tokens | Cost | Time | Tool calls | Notes |
|---|---|---|---|---|---|---|
| A (no tools) | **failed** | 173,853 | $0.07 | 30 s | 4 | Right decision (`AddTaxAt` instead), but it couldn't build. `VB6.EXE` from its shell failed with "requires elevation", and it didn't try to get around that. It appended the code with a shell heredoc, which put LF line endings into the CRLF class file (`lone-lf`). |
| B (vb6 tools) | correct, unaided | 273,596 | $0.11 | 38 s | 7 | `symbol_refs`, 2 × `file_replace_lines`, `project_save`, `build_make`. It added `RateFor` and `AddTaxAt`, kept `AddTax`, explained the conflict and offered to change `AddTax` if the user decides so. |

## What it shows

- **Both agents recognised the conflict.** The difference is everything around it:
  - A couldn't build a DLL without elevation.
  - A damaged the file's line endings with a shell append.
  - B built through the elevated add-in with no prompt, and its edits kept the file intact.
- **B didn't run the client to confirm** ("I haven't run ClientApp.exe"). The scoring did it instead. `exe_run` plus `app_tree` would have let it check.
- **Side finding, fixed:**
  - `com_describe` showed nothing of a VB-built DLL's methods by default. VB marks each class's own interface (`_clsTax`) hidden, so the listing had only `Implements _clsTax`.
  - Interfaces that a visible class implements are now shown.
