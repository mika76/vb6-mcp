# Spike 22: the sync engine, live (Phase 4)

Date: 2026-10-14. Scratch copies of the Basic fixture (`tests/fixtures/scratch/Sync1`, `Sync2`), elevated IDE, the watchdog's `SyncPoller` (2 s), add-in `WRITES` / `RELOAD`.

## Results

| Test | Result |
|------|--------|
| Outside change to an unmodified module | ✓ reloaded 4 s later; the IDE showed the new text, still unmodified; event `reloaded` |
| Same, while VB6 showed the "Add this project to SourceSafe?" prompt | ✓ reloaded anyway (the deferred action ran from the prompt's message loop) |
| Outside change while the module has an unsaved edit | ✓ `conflict`; the IDE edit untouched; `sync_diff` showed both sides |
| `sync_take_disk` | ✓ IDE equals disk; the conflict resolved on the next poll ("no unsaved changes any more") |
| ... and its recovery copy | ✗ round 1: **deleted by the recovery poller** (the component was no longer unsaved). ✓ round 2: kept in its own session (`KeptId`, with a note), still there after several polls |
| `sync_take_ide` | ✗ round 1: **"Save Project Group As" dialog**. ✓ round 2: saved, disk equals IDE, `alsoSaved` empty |
| The IDE's own save after a conflict | ✓ `resolved` ("the IDE saved its version over the disk file"), **no reload**: the add-in's `WRITES` record works |
| `Group1.vbg` written by opening or replacing | ✓ none |

## Findings

1. **`FileControlEvents` has two similarly named interfaces.** The coclass sources `_dispVBFileControlEvents`, whose `AfterWriteFile` has 4 parameters (`VBProject, FileType, FileName, Result`). The type library also has `_FileControlEvents`, with a 2-parameter version. A handler written to the wrong one fails to compile: "Procedure declaration does not match description of event". Dump the coclass's source interface, not an interface found by name.
2. **`AddFromFile(path, False)` turns the IDE into a project group**, even with one project, and Save then asks for a group file name. `AddFromFile(path, True)` with projects still open writes a `Group1.vbg` of the closed projects. So: remove the open projects first, then `AddFromFile(path, True)`. Use `False` only to add a project next to others.
3. **Copies kept on purpose must not live in the IDE's own recovery session:** the recovery poller prunes copies of components that are no longer unsaved.
4. Deferred actions (`RELOAD`) also run while a modal prompt is open, because the prompt pumps messages.

## Decision

The sync engine stays as designed (PLAN.md §8), with the two fixes. Not built: a 3-way merge, and vetoing an IDE save over a newer file (`RequestWriteFile`); conflicts are reported and the user chooses.
