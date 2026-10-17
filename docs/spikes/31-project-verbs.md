# Spike 31: closing, adding, saving and configuring projects through VBIDE (Phase 7 A1-A4, B8)

**Setup:**
- First live run of the new add-in verbs (`ProjectVerbs.bas`), 2026-10-17.
- VB6 was started with `ide_launch` through a watchdog the owner had started with `--stay`: elevated, no UAC prompt, 4 s, pid 19332.
- Scratch copies of the Basic fixture: `D:\bench\spike-we`, and `spike-make` for `/make`.

## Results

| What | Result |
|---|---|
| `PRJGET` / `PRJSET` | Description, startup `Sub Main` and back to `frmMain` (a component) all read back correctly. |
| `COMPADD new` | A class and a form were added and named; a duplicate name and an unknown kind were refused before anything happened. |
| `COMPADD file` | An existing `.bas` was added under its `VB_Name` (`modExtra`). The add-in's log line right after `AddFile` showed `Module1`, but the name read back moments later was `modExtra`. |
| **Saving after `COMPADD new`** | **File > Save Project stops at a "Save File As" dialog** for each component that has never been saved. |
| **`CLOSE` with discard** | **Wrong: `VBComponent.IsDirty = False` saves the component.** A changed `modMath` was "discarded" and closed, and the change (`' DISCARD ME`) was on disk afterwards. The first try also wrote the two new components and the `.vbp` at the moment of the close. |
| `CLOSE` without discard | Refused with the changed components named, as intended. |
| `SAVECOMP` (`VBComponent.SaveAs` with its own path) | Writes that one component quietly: no dialog, `IsDirty` cleared, the `.vbp` byte-identical. |
| `.vbp` keys with `reopen` | `CondComp="spike = 1"` added and `Command32` changed byte-exactly; the project closed, reopened and showed nothing unsaved. |
| `MAKE` with `/d` | The fixture with `#If spike = 1 Then ThisSubDoesNotExist` built without defines and failed with `/d spike=1` on exactly that line. |

## Second round (after the rebuild)

- **`VBProjects.Remove` saves too.** With `IsDirty` left alone, a changed `modMath` was closed with discard and the change was on disk again. No prompt appeared at all, so there was nothing to answer No to. VBIDE has no way to close a project without saving.
- **New components:** saved straight away by `project_add_component` (`clsFresh.cls` written at once). `project_save` then needed no dialog, and its `.vbp` report showed just the new `Class=` line.
- **`project_save components`:** `modMath` saved alone, the `.vbp` byte-identical.
- **Discard by snapshot:** before closing, the server keeps the bytes of the `.vbp` and of each changed component's file (with its `.frx`/`.ctx`/`.pgx`) and lists the source files in the project folder. After VB has closed (and saved), it writes the snapshots back and deletes source files VB created for never-saved components. Live: `modMath.bas` and the `.vbp` had their earlier hashes after a discard, and no new files were left.

## Decisions

- **Discard:** never touch `IsDirty` (setting it to False *saves*), and expect no prompt: `VBProjects.Remove` saves silently too.
  - `discard: true` snapshots and restores the files the close writes (second round above).
  - It is refused if another IDE has one of those files loaded (hard rule 4), or if the project was never saved.
- **New components:** `project_add_component` saves a new component straight away to `<project folder>\<name>.<ext>`, with `SAVECOMP <project> <component> <path>` (`SaveAs` to a new file). The file must not exist yet, so nothing is overwritten. Later saves then need no dialog.
- **Component-only save:** `project_save` takes `components: [...]`, which saves just those components through `SAVECOMP`. File > Save Project rewrites the `.vbp`; for a production project that bumped control versions (PLAN B8).
- **Unchanged:** `PRJGET`/`PRJSET`, `COMPADD file`, the `.vbp` round trip and `/make` defines work as built.

AGENTS.md gotcha added: `IsDirty = False` saves.
