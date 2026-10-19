# Spike 43: creating a new project through VBIDE (2026-10-19)

## Why

An agent that needs a new VB6 project had no tool for it (notes from real use; roadmap "Soon" after 1.0.0):
- `ide_launch` without a project shows the New Project chooser, and answering it gives an unsaved `Project1`.
- `project_save` answered `nothing-to-save` for that project while it was unchanged, and refused it once changed ("never been saved").
- So agents wrote the `.vbp` and the form files by script, which the hook flags.

## Setup

- VB6 non-elevated (spike 41), started with `ide_launch elevation=normal`; the add-in built from the `feature/project-new` worktree with two new deferred verbs:
  - `PRJNEW <kind> <name> <path> <add|replace>`: `VBProjects.Add`, set the name, then save as below.
  - `PRJSAVEAS <project> <path>`: save a project that has no `.vbp` yet.
  - Saving: each never-saved component `VBComponent.SaveAs <folder>\<name>.<ext>`, then `VBProject.SaveAs <path>`.
- The new server tools (`project_new`, `project_save path`) driven from a script over stdio, and single add-in verbs sent straight over the socket. Scratch folders under a workspace root, deleted afterwards.
- The machine has Visual SourceSafe installed (VB's source-control add-in is active).

## Results

### Round 1 (component check by `FileCount = 0`)

| What | Result |
|---|---|
| New Project chooser, Open | `Project1` with `Form1`; `COMPONENTS` reports Form1's file as `""` |
| `project_save` without `path` | refused: "never been saved; pass path" (no longer `nothing-to-save`) |
| `PRJSAVEAS Project1 ...\Chooser.vbp` | **Only the `.vbp` was written, with no `Form=` line**, and no `Form1.frm`. The add-in's loop skipped Form1. |
| Why | **A new project's component reports `FileCount = 1` with an empty `FileNames(1)`**, not `FileCount = 0`. The add-in log for `PRJNEW dll`: `Class1 (type 2, files 1)`. |
| `VBProject.SaveAs` with a never-saved component | Saves the `.vbp` without it, **quietly**: no Save File As dialog. The project then reports `IsDirty = False`. |
| Repair attempts | `SAVECOMP` (with a path) wrote `Class1.cls` and VB showed the project clean, but the `.vbp` still had no `Class=` line. File > Save Project (`SAVE`) wrote nothing: VB counts the project as saved. Setting the description (`PRJSET`) didn't mark it changed either. **Once `VBProject.SaveAs` has run without the component, nothing VB does later adds the line.** |
| SourceSafe | Every `VBProject.SaveAs` of a new project asked **"Add this project to SourceSafe?"** (Source Code Control dialog, Yes/No/Help), *after* the files were written. Answering No still left a read-only `MSSCCPRJ.SCC` in the folder. |
| `PRJNEW dll ... add` with a project open | Added next to it: two projects, a group. `VBProjects.FileName` returned `Group1.vbg` although no group file exists. |
| `replace` with a changed project open | refused before anything happened |
| Quitting that IDE | two Save File As dialogs (for the never-saved Form1 and the group); the IDE was killed. |

### Round 2 (never saved = `FileCount = 0` or an empty `FileNames(1)`)

| What | Result |
|---|---|
| Chooser `Project1`, `project_save path` | `Form1.frm` saved quietly, then `Chooser.vbp` with `Form=Form1.frm`; the SourceSafe question after. The project keeps its name `Project1`. |
| `project_new exe ... replace` (one clean project open) | The open project closed without a prompt; `NewExe` created with `Exclusive` (nothing left open), `Form1.frm` saved and listed, project clean. |
| `project_new dll`, `ocx`, `axexe` (name `AxServer`) with `add` | One project each: `Type=OleDll` + `Class=Class1; Class1.cls`, `Type=Control` + `UserControl=UserControl1.ctl`, `Type=OleExe` + `Class=Class1; Class1.cls`, `Startup="(None)"`. Four projects open, all clean. |
| Refusals | existing `.vbp`; the default component's file already in the folder (`Form1.frm`); outside the roots; an invalid name from the file name (`My App`); a project of that name already open; no `mode` with projects open; unknown kind; `project_save path` for a project that has a `.vbp`. All refused before anything was sent. |
| `project_new` in an IDE with no project (chooser cancelled) | First refused: the title is just "Microsoft Visual Basic", with no `[design]` tag, so the server saw mode `unknown`. Now allowed when no project is open and no dialog shows (the add-in checks design mode itself: End disabled). Then `Hello.vbp` + `Form1.frm`, clean. |
| Quitting with the four-project group | "Save changes to the following files?" (the group only). No: VB closed, every project file byte-identical; VB wrote its `.vbw` workspace files as usual and no `.vbg`. |

## Decisions

- **`project_new`** (`kind`, `path`, `name?`, `mode?`) and **`project_save path`** for a never-saved project, both deferred (hard rule 14: `VBProjects.Add` creates a form or UserControl, and saving goes through VB's file code).
- **A component is never saved** when `FileCount = 0` or `FileNames(1)` is empty (`NeverSaved` in `ProjectVerbs.bas`; also used by `SAVECOMP`). Each one is saved next to the `.vbp` under its name **before** `VBProject.SaveAs`. An existing file stops the save before anything is written; the server checks the default component's file (`Form1.frm`, `Class1.cls`, `UserControl1.ctl`) before it sends anything, so a clash doesn't leave a new project open and unsaved.
- **The server checks the outcome on disk:** every component must have a file and be listed in the `.vbp`; otherwise the result is `incomplete` with `notInVbp`.
- **`mode`** like `project_open`: required when the IDE has a project; `replace` refused while any is dirty (spike 31: closing saves).
- **SourceSafe and other dialogs** are reported, not answered (hard rule 9). The result says the files are already written and how to check.
- **Not built:** templates (`VBProjects.AddFromTemplate`), saving a group file, Save As of a saved project to another path. No real use asked for them.
