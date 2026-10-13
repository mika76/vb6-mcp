# Spike 9 (part): project-map scan and 1252 round trip

Date: 2026-10-13. Code: `src/Vb6Mcp.Core/Files/ProjectMap.cs`, `VbpFile.cs`, `VbSourceFile.cs`; tests `ProjectMapTests.ScanRealRoots` (category `Scan`), `VbSourceFileTests.Every1252ByteRoundTrips`.

## Questions

1. Is a full scan of the real workspace fast enough to keep the project map in memory (PLAN §9), or does it need a cache file?
2. Do all bytes survive a Windows-1252 decode/encode round trip (hard rule 1)?

## Results

### Scan of the owner's workspace (`D:\src`)

| | |
|---|---|
| Projects | **43** (Exe 33, Control/OCX 9, OleDll 1) |
| Member files referenced | 1,498 |
| Files used by more than one project | **139** |
| `.vbp` parse errors | 0 |
| Scan time, cold file cache | **4.9 s** |
| Scan time, warm | **0.34 s** |

Only counts and timing were printed, no file contents.

### 1252 round trip

All 256 byte values decode and re-encode byte-exact through .NET's code page 1252 (with `CodePagesEncodingProvider`), including the five bytes 1252 leaves undefined (0x81, 0x8D, 0x8F, 0x90, 0x9D).

## Decisions

- **In-memory project map, no cache file.** The server builds it lazily on first use, caches it, and rescans on request (`refresh`) or when it's older than a few minutes. The cold scan (~5 s) happens at most once per server start.
- Shared files are common (139 of 1,498), which confirms that `ws_file_projects` (which projects use this file?) is worth having.
- `.NET` 1252 is lossless, so `Vb1252.Encoding` is safe for byte-exact round trips. `VbCode.PrepareForInsert` uses a *strict* 1252 encoder only to refuse characters 1252 can't represent.

## Still open (rest of spike 9)

- `/make /out` log formats for every case (success and compile error seen; missing reference not yet).
- Elevated `/make` of a DLL/OCX registering properly from the add-in (Phase 4).
