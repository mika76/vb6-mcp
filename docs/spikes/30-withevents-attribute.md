# Spike 30: does VB write a WithEvents variable's VB_VarHelpID back on save?

**Question:** `file_replace_lines` refused to change a `Dim WithEvents m_comm As MSCommLib.MSComm` declaration (improvement notes item 12), because its hidden `Attribute m_comm.VB_VarHelpID = -1` line had no line to follow in the new text. Can the tool drop that attribute safely?

**Setup:** a scratch copy of the Basic fixture (`D:\bench\spike-we`). `frmMain.frm` was given `Private WithEvents m_btn As VB.CommandButton` **without** the attribute line, then opened in the elevated IDE (pid 20424, 2026-10-17). One comment line was inserted through `code_insert` so the form counted as changed, then the project was saved with `project_save`.

**Result:** the saved file has

```
Private WithEvents m_btn As VB.CommandButton
Attribute m_btn.VB_VarHelpID = -1
```

VB writes the attribute itself whenever it saves a `WithEvents` declaration. Losing it in a disk edit costs nothing once the file has been through the IDE.

**Also found:** the range logic had a bug. Hidden attribute lines right after the **last** replaced code line weren't part of the range. They stayed in the file and ended up after whatever the new text's last line was (`#End If` in the real case), or orphaned after the previous line when the declaration was deleted.

**Decision:**
- The range now includes attribute lines right after its last code line, so they are re-attached to their line in the new text, or the edit is refused.
- `file_replace_lines` has `drop_var_help_ids: true`, which drops only `VB_VarHelpID` attributes that have no line to follow. Others (`VB_UserMemId`, `VB_Description`, ...) can't be recreated by VB and still refuse.
- The refusal message names the option.

Side finding: the save also added `IconForm="frmMain"` to the `.vbp`. The new `project_save` report (B8) listed it.
