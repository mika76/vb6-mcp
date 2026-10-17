# Spike 36: load errors that "keep coming", and dialog clicks that don't take (PLAN D3, D4)

**Date:** 2026-10-17. **Source:** improvement notes 10 and 11. Opening a production project by hand, "Errors during load" came back after every OK, more than ten times, while the watchdog's recovery poll logged add-in error 500 every 10 s; and a first click on "Continue Loading Project?" once returned `dialogClosed: false`.

## Fixture

`spikes/36-load-errors/make.cs` writes a copy of the Basic fixture whose form uses a control VB6 can't load: `Object=` MSCOMCTL's own LIBID, read from the OCX's type library (hard rule 13: no hand-written GUID), with major version 9, which isn't registered, and a `ProgressBar` on the form. An optional count adds more forms with the same control. Files go through `Vb1252` (hard rule 1).

## What happened (VB6 launched through the watchdog with the project, as a user opening it)

1. **"'…\MSCOMCTL.OCX' could not be loaded--Continue Loading Project?"**, once for the missing OCX. Answered Yes.
2. **"Errors during load. Refer to '…\frmMain.log' for details"**, then the same for `frmExtra2.log` and `frmExtra1.log`: **one per form that uses the control**, each naming its own log, none repeated. Then the SourceSafe question (the folder isn't bound).
3. While the first dialog was up, **VBIDE refused access**: the add-in's PROJECTS answered error 500 "Method '~' of object '~' failed". The watchdog's recovery poll logged it every 10 s, exactly as in improvement note 10, and `ide_launch` failed with it instead of reporting the dialog (and didn't claim the IDE).
4. **The polling didn't make anything repeat.** Each dialog appeared once.
5. **Every click took the first time** (`dialogClosed: true`, 9 of 9 over three launches), also when the next dialog had the same caption.

## Conclusions

- D3: the "more than ten" load errors in a production project were one per affected form (that project had two missing controls on many forms), not something our polling caused. No back-off of verbs is needed; the access error during a load is VB's.
- D4: not reproduced. The original miss may have been a dialog modal to the user's own File > Open; the tool already reports `dialogClosed` and the dialogs now, so the agent sees it and can click again.

## Changes

- `ide_launch` claims the IDE first, and treats an add-in error while reading the projects as "still loading": `status: "started"` with the dialog and a note (one "Errors during load" per form, each naming its `.log`).
- The watchdog's recovery poll logs "waits: a dialog is open" once instead of error 500 every 10 s.
