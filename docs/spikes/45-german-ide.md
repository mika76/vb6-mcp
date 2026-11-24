# Spike 45 — German VB6 IDE captions and build logs

A German Visual Basic 6 installation reports these captions while opening, running and hitting a breakpoint in the Debug fixture:

- `DebugFixture - Microsoft Visual Basic [Entwerfen]`
- `DebugFixture - Microsoft Visual Basic [Ausführen] - [frmDebug (Code)]`
- `DebugFixture - Microsoft Visual Basic [Unterbrechen] - [frmDebug (Code)]`

At the breakpoint, the top-level window of class `wndclass_desked_gsk` was hidden, while the same process had a visible `IDEOwner` with the caption `DebugFixture - Microsoft Visual Basic [Unterbrechen]`. The owner is a fallback for the status reader. Window lookup for UI actions still uses the IDE main frame.

The German compiler wrote these /make logs (the source path is shortened here):

- Success: `'Basic.exe' wurde erfolgreich erstellt.`
- Compile error: `Kompilierungsfehler in Datei 'D:\s\frmMain.frm', Zeile 6 : Variable nicht definiert`
- Failure status: `Das Erstellen von 'Basic.exe' wurde mit einem Fehler abgebrochen.`

The compile error points at IDE line 7, so the German log uses the same zero-based numbering as the English log in spike 21.
