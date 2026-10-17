# Spike 34: what a .frx holds, and how to read it

**Date:** 2026-10-17. **Why:** the owner asked how an agent could take an image out of a form. VBIDE's `SavePicture` of a property's picture writes BMP whatever the original was; the `.frx` keeps the original bytes and is readable for closed projects too. The format was read from a production project's `frmMain.frx` and `frmFilter.frx` (benchmark copies), not taken from documentation.

## References in the .frm

Properties stored in the companion file point at a byte offset (hex) in it:

```
Icon            =   "frmMain.frx":0000
Picture         =   "frmMain.frx":08CA
DesignConnection=   $"frmMain.frx":659B      ($ = a long string)
ItemData        =   "frmFilter.frx":0000
List            =   "frmFilter.frx":000E
```

## Blocks

| Kind | Layout at the offset | Seen |
|---|---|---|
| Picture (Picture, Icon, MouseIcon, DragIcon, DownPicture, ...) | `u32 blockLength` · `"lt\0\0"` · `u32 imageLength` · the original image bytes | `Icon` at 0: block `0x8C6`, image `0x8BE`, an ICO (`00 00 01 00`, one 32×32). `Picture` at `0x08CA`: a **GIF89a**, kept as GIF. 4 + `0x8C6` = `0x8CA`, the next offset |
| Long string (`$"…"`) | `u32 length` · the text (ANSI) | `DesignConnection` at `0x659B`: a 139-byte connection string |
| List / ItemData | `u16 count` · `u16` (not needed) · count × (`u16 length` · text) | `ItemData` at 0: `-1`, `0`, `1`; `List` at `0x0E`: `No Selection`, `Not Active`, `Active` (ends exactly where the next block starts) |

A survey of every `.frx`/`.ctx` reference under `D:\src` (owner: "lots of things can be kept in frx, not just images") found more kinds:

| Kind | Layout | Seen in |
|---|---|---|
| Picture inside a saved StdPicture | `u32 block` · StdPicture CLSID `{0BE35204-8F91-11CE-9DE3-00AA004BB851}` (16 bytes) · `"lt\0\0"` · `u32 imageLength` · image | `MouseIcon`, `CellPicture` (grid), `Image`, **ImageList images** |
| ImageList images | in the `.frm`: `BeginProperty Images` › `BeginProperty ListImage1..n` with `Picture = "x.frx":…`, `Key`, `Tag`; each picture as above | a production `frmMain`: 4 GIFs, 32×32 |
| Short text | `u8 length` · text | `Text` (`05 Text4`) |
| Multi-line text | `u32 length` · text with CR LF | a slip layout in a `Text` |
| Font | `u32 block` · StdFont CLSID `{0BE35203-…}` · 10 bytes (version, charset, flags, weight, size) · `u8 length` · face name | `font` property of an ActiveX control (Tahoma) |
| An ActiveX control's own state | anything else: `u32 length` · bytes (`OleObjectBlob` starts `LB`, `MessagesAndKeys` with a CLSID, `PropertyPages` a list of names) | SpyWorks WinHook, a scanner control, UserControls |
| Empty picture | the `lt` block with image length 0 | SSTab `TabPicture(n)` |
| Long strings with `$` | `u32 length` · text | `Caption`, `Sql`, `ToolTipText`, `DesignConnection` |

Image types are told apart by their first bytes: `GIF8`, `FF D8` (JPEG), `BM` (BMP), `00 00 01 00` (ICO), `00 00 02 00` (CUR), `D7 CD C6 9A` (placeable WMF), `89 PNG`, EMF (` EMF` at offset 40).

## Decision

- A read-only tool, `form_resources`: lists a form's `.frx`-backed properties with their control and kind, and extracts pictures in their **original format** (files in a folder under the roots, never overwriting; or shown inline as PNG), lists as items and long strings as text.
- Long strings can be connection strings with passwords (this one is): `Password=`/`Pwd=` values are masked in results.
- Still no `.frx` **editing** (hard rule 2): changes go through `form_set_property` (`@picture:`, `List`), and VB rewrites the file on save.
- An open project's unsaved designer changes aren't in the `.frx` until VB saves; the tool reads disk and says when the IDE has the form dirty.
