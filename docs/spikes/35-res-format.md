# Spike 35: reading VB6 resource files (.res)

**Date:** 2026-10-17. **Why:** the owner asked which other binary files an agent might need ("this is a general use MCP"). VB6 code loads resources by id (`LoadResString`, `LoadResPicture`, `LoadResData`), so changing UI text or icons means knowing which ids exist and what they hold. The layout was read from `.res` files on this machine, not taken from documentation: `AceScrollBarControl.res` and `VersionRipper.res` in the owner's workspace, and the VB98 sample `AECLIENT.RES`.

## Layout (32-bit .res)

- An **empty first entry**: data size 0, header size 32, type and name 0. Both workspace files start with it; it is how a 32-bit `.res` is recognised (16-bit ones aren't read).
- Then per resource: `u32 dataSize` · `u32 headerSize` · **type** and **name**, each `FFFF` + `u16` ordinal or a null-terminated UTF-16 string · padding to 4 · `u32 dataVersion` · `u16 memoryFlags` · `u16 language` · `u32 version` · `u32 characteristics` · the data · padding to 4.

| Seen | Entries |
|---|---|
| `AceScrollBarControl.res` | ICON 1 (744 bytes, language 0x1C09 en-ZA), GROUP_ICON `MAINICON` (a string name, header 48), BITMAP 1 (a 16×15 DIB) |
| `VersionRipper.res` | VERSION 1: file/product version 1.9.0.204 and its string values |
| `AECLIENT.RES` | STRING blocks: block n holds ids (n−1)·16 … (n−1)·16+15, each `u16 length` + UTF-16 text; 52 strings |

## Converting

- **BITMAP** is a DIB without the 14-byte file header: prepend `BM`, file size and the pixel offset (14 + header + palette). Checked: a valid 16×15, 4-bit `.bmp`.
- **GROUP_ICON / GROUP_CURSOR** lists its images by resource id (14-byte entries); an `.ico`/`.cur` file lists them by offset (16-byte entries). Cursor images start with their hotspot, which the `.cur` entry carries. Checked: a valid 32×32 `.ico`.
- **VERSION**: nested blocks (`u16 length`, `u16 valueLength`, `u16 type`, UTF-16 key, value, children); `VS_FIXEDFILEINFO` (signature `FEEF04BD`) gives the version numbers, `StringFileInfo` › table › strings the values.

## Decision

A read-only tool, `res_read`: lists resources, the string table and version info, and extracts resources (bitmaps as `.bmp`, icon/cursor groups as `.ico`/`.cur`, others raw) to a folder under the roots or shows a picture. Changing a `.res` stays with VB's Resource Editor; the hook refuses Edit/Write on `.res`.
