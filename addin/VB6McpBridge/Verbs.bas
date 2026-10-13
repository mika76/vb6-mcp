Attribute VB_Name = "Verbs"
' Request verbs of wire protocol v1 (docs/protocol.md "Verbs (v1)").
' Dispatch returns "OK" US payload, or "ERR" US code US message.
' Rules (spike 3): indexed loops only over VBIDE collections (For Each raises error 70 in run/break
' mode); nothing here runs user code, so nothing needs deferring yet (hard rule 14).
Option Explicit

' Error numbers raised inside verbs and mapped to protocol codes
Private Const E_BADREQUEST As Long = vbObjectError + 400
Private Const E_NOTFOUND As Long = vbObjectError + 404
Private Const E_CONFLICT As Long = vbObjectError + 409

Public Function Dispatch(ByVal verb As String, args() As String) As String
    Dim payload As String
    On Error GoTo Fail
    Select Case UCase$(verb)
    Case "PING":       NeedArgs args, 0, 0: payload = ProcessId() & US() & AddinVersion() & US() & PROTOCOL_VERSION
    Case "PROJECTS":   NeedArgs args, 0, 0: payload = Projects()
    Case "COMPONENTS": NeedArgs args, 1, 1: payload = Components(args(0))
    Case "READ":       NeedArgs args, 2, 4: payload = ReadCode(args)
    Case "MEMBERS":    NeedArgs args, 2, 2: payload = MembersOf(args(0), args(1))
    Case "PROC":       NeedArgs args, 4, 4: payload = ProcText(args(0), args(1), args(2), args(3))
    Case "OPEN":       NeedArgs args, 2, 2: Dispatch = QueueOpen(args(0), args(1)): Exit Function
    Case "REPLACE":    NeedArgs args, 6, 6: payload = ReplaceLines(args)
    Case "SAVE":       NeedArgs args, 1, 1: Dispatch = QueueSave(args(0)): Exit Function
    Case "CLICK":      NeedArgs args, 1, 1: payload = ClickButton(args(0))
    Case Else:         Err.Raise E_BADREQUEST, , "unknown verb " & verb
    End Select
    Dispatch = "OK" & US() & payload
    Exit Function
Fail:
    Select Case Err.Number
    Case E_BADREQUEST: Dispatch = "ERR" & US() & "400" & US() & Err.Description
    Case E_NOTFOUND:   Dispatch = "ERR" & US() & "404" & US() & Err.Description
    Case E_CONFLICT:   Dispatch = "ERR" & US() & "409" & US() & Err.Description
    Case Else:         Dispatch = "ERR" & US() & "500" & US() & Err.Number & ": " & Err.Description
    End Select
End Function

Private Sub NeedArgs(args() As String, ByVal minCount As Long, ByVal maxCount As Long)
    Dim n As Long
    n = UBound(args) + 1
    If n < minCount Or n > maxCount Then Err.Raise E_BADREQUEST, , "expected " & minCount & IIf(maxCount <> minCount, "-" & maxCount, "") & " arguments, got " & n
End Sub

Private Function B(ByVal v As Boolean) As String
    B = IIf(v, "1", "0")
End Function

' ---------- lookup ----------

Private Function FindProject(ByVal prjName As String) As VBIDE.VBProject
    Dim p As Long
    For p = 1 To g_VBE.VBProjects.Count
        If StrComp(g_VBE.VBProjects(p).Name, prjName, vbTextCompare) = 0 Then Set FindProject = g_VBE.VBProjects(p): Exit Function
    Next
    Err.Raise E_NOTFOUND, , "project not found: " & prjName
End Function

Private Function FindComponent(ByVal prjName As String, ByVal cmpName As String) As VBIDE.VBComponent
    Dim prj As VBIDE.VBProject, c As Long
    Set prj = FindProject(prjName)
    For c = 1 To prj.VBComponents.Count
        If StrComp(prj.VBComponents(c).Name, cmpName, vbTextCompare) = 0 Then Set FindComponent = prj.VBComponents(c): Exit Function
    Next
    Err.Raise E_NOTFOUND, , "component not found: " & prjName & "." & cmpName
End Function

Private Function FileOf(ByVal cmp As VBIDE.VBComponent) As String
    On Error Resume Next
    If cmp.FileCount > 0 Then FileOf = cmp.FileNames(1)
End Function

Private Function ProjectFile(ByVal prj As VBIDE.VBProject) As String
    On Error Resume Next                                     ' raises for a never-saved project
    ProjectFile = prj.FileName
End Function

' ---------- verbs ----------

' index US name US type US fileName US isDirty, records separated by RS
Private Function Projects() As String
    Dim p As Long, prj As VBIDE.VBProject, out As String
    For p = 1 To g_VBE.VBProjects.Count
        Set prj = g_VBE.VBProjects(p)
        If p > 1 Then out = out & RS()
        out = out & p & US() & prj.Name & US() & prj.Type & US() & ProjectFile(prj) & US() & B(prj.IsDirty)
    Next
    Projects = out
End Function

' name US type US isDirty US countOfLines US fileName, records separated by RS
Private Function Components(ByVal prjName As String) As String
    Dim prj As VBIDE.VBProject, c As Long, cmp As VBIDE.VBComponent, out As String, n As Long
    Set prj = FindProject(prjName)
    For c = 1 To prj.VBComponents.Count
        Set cmp = prj.VBComponents(c)
        n = -1
        On Error Resume Next
        n = cmp.CodeModule.CountOfLines
        On Error GoTo 0
        If c > 1 Then out = out & RS()
        out = out & cmp.Name & US() & cmp.Type & US() & B(cmp.IsDirty) & US() & n & US() & FileOf(cmp)
    Next
    Components = out
End Function

' args: project, component [, startLine, count] -> countOfLines US isDirty US text
Private Function ReadCode(args() As String) As String
    Dim cmp As VBIDE.VBComponent, cm As VBIDE.CodeModule, total As Long, first As Long, n As Long, text As String
    Set cmp = FindComponent(args(0), args(1))
    Set cm = cmp.CodeModule
    total = cm.CountOfLines
    first = 1: n = total
    If UBound(args) = 3 Then
        If Not IsNumeric(args(2)) Or Not IsNumeric(args(3)) Then Err.Raise E_BADREQUEST, , "startLine and count must be numbers"
        first = CLng(args(2)): n = CLng(args(3))
        If first < 1 Or n < 0 Then Err.Raise E_BADREQUEST, , "startLine must be >= 1 and count >= 0"
        If first + n - 1 > total Then n = total - first + 1
    ElseIf UBound(args) <> 1 Then
        Err.Raise E_BADREQUEST, , "READ takes 2 or 4 arguments"
    End If
    If n > 0 And first <= total Then text = cm.Lines(first, n)
    ReadCode = total & US() & B(cmp.IsDirty) & US() & text
End Function

' name US memberType US scope US static US codeLocation US procKind US procStart US procBody US procCount
' Members lists a property only once, under its first kind (Phase 1 finding), so properties are
' expanded to one record per kind that exists (Get, Let, Set).
Private Function MembersOf(ByVal prjName As String, ByVal cmpName As String) As String
    Dim cm As VBIDE.CodeModule, i As Long, m As VBIDE.Member, out As String, head As String, kinds As Variant, k As Long
    Set cm = FindComponent(prjName, cmpName).CodeModule
    For i = 1 To cm.Members.Count
        Set m = cm.Members(i)
        head = m.Name & US() & m.Type & US() & m.Scope & US() & B(m.Static) & US() & m.CodeLocation & US()
        If m.Type = vbext_mt_Property Then
            kinds = Array(vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)
            For k = 0 To 2
                If HasProc(cm, m.Name, kinds(k)) Then AppendRecord out, head & KindInfo(cm, m.Name, kinds(k))
            Next
        Else
            AppendRecord out, head & ProcInfo(cm, m)
        End If
    Next
    MembersOf = out
End Function

Private Sub AppendRecord(out As String, ByVal rec As String)
    If Len(out) > 0 Then out = out & RS()
    out = out & rec
End Sub

Private Function HasProc(ByVal cm As VBIDE.CodeModule, ByVal procName As String, ByVal kind As Long) As Boolean
    Dim n As Long
    On Error GoTo Done
    n = cm.ProcStartLine(procName, kind)
    HasProc = True
Done:
End Function

Private Function KindInfo(ByVal cm As VBIDE.CodeModule, ByVal procName As String, ByVal kind As Long) As String
    KindInfo = kind & US() & cm.ProcStartLine(procName, kind) & US() & cm.ProcBodyLine(procName, kind) & US() & cm.ProcCountLines(procName, kind)
End Function

' procKind US procStart US procBody US procCount for methods/properties; four empty fields otherwise
Private Function ProcInfo(ByVal cm As VBIDE.CodeModule, ByVal m As VBIDE.Member) As String
    Dim kind As VBIDE.vbext_ProcKind, nm As String
    On Error GoTo NoProc
    If m.Type <> vbext_mt_Method And m.Type <> vbext_mt_Property Then GoTo NoProc
    nm = cm.ProcOfLine(m.CodeLocation, kind)
    If Len(nm) = 0 Then GoTo NoProc
    ProcInfo = kind & US() & cm.ProcStartLine(nm, kind) & US() & cm.ProcBodyLine(nm, kind) & US() & cm.ProcCountLines(nm, kind)
    Exit Function
NoProc:
    ProcInfo = US() & US() & US()
End Function

' args: path, mode (add|replace) -> QUEUED. Validates now, opens later (Defer.bas, hard rule 14):
' opening a project can show modal dialogs, which must not run inside the request handler.
Private Function QueueOpen(ByVal path As String, ByVal mode As String) As String
    Dim ext As String, p As Long
    mode = LCase$(mode)
    If mode <> "add" And mode <> "replace" Then Err.Raise E_BADREQUEST, , "mode must be add or replace"
    ext = LCase$(Right$(path, 4))
    If ext <> ".vbp" And ext <> ".vbg" Then Err.Raise E_BADREQUEST, , "path must be a .vbp or .vbg file"
    If Len(Dir$(path)) = 0 Then Err.Raise E_NOTFOUND, , "file not found: " & path
    If mode = "replace" Then
        For p = 1 To g_VBE.VBProjects.Count
            If g_VBE.VBProjects(p).IsDirty Then Err.Raise E_CONFLICT, , "project " & g_VBE.VBProjects(p).Name & " has unsaved changes; replace would prompt to save"
        Next
    End If
    Enqueue "OPEN" & US() & path & US() & mode
    QueueOpen = "QUEUED" & US() & "open " & path & " (" & mode & ")"
End Function

' args: project, component, startLine, count, expectedOld, newText -> countOfLines US isDirty
' Replaces count lines at startLine with newText (CR LF separated). count 0 = insert before
' startLine; empty newText = delete. The write only happens if the target lines still equal
' expectedOld: checked and written in one go on the IDE thread, so a user edit in between is a
' 409 conflict, never overwritten. Mode (design only) and line limits are checked by the server.
Private Function ReplaceLines(args() As String) As String
    Dim cmp As VBIDE.VBComponent, cm As VBIDE.CodeModule, first As Long, n As Long, total As Long, current As String
    If Not IsNumeric(args(2)) Or Not IsNumeric(args(3)) Then Err.Raise E_BADREQUEST, , "startLine and count must be numbers"
    first = CLng(args(2)): n = CLng(args(3))
    Set cmp = FindComponent(args(0), args(1))
    Set cm = cmp.CodeModule
    total = cm.CountOfLines
    If first < 1 Or n < 0 Or first > total + 1 Or first + n - 1 > total Then
        Err.Raise E_BADREQUEST, , "line range " & first & "+" & n & " is outside the module (" & total & " lines)"
    End If
    If n > 0 Then current = cm.Lines(first, n)
    If current <> args(4) Then Err.Raise E_CONFLICT, , "the target lines changed since they were read"
    If n > 0 Then cm.DeleteLines first, n
    If Len(args(5)) > 0 Then cm.InsertLines first, args(5)
    ReplaceLines = cm.CountOfLines & US() & B(cmp.IsDirty)
End Function

' args: project -> QUEUED. Runs VB6's own File > Save Project command (deferred, hard rule 14):
' saving can show dialogs (source-control checkout, read-only file), so it must not run inside
' the request handler. A never-saved project is refused: Save would open a Save As dialog.
Private Function QueueSave(ByVal prjName As String) As String
    Dim prj As VBIDE.VBProject
    Set prj = FindProject(prjName)
    If Len(ProjectFile(prj)) = 0 Then Err.Raise E_CONFLICT, , "project " & prjName & " has never been saved; Save would open a Save As dialog"
    Enqueue "SAVE" & US() & prj.Name
    QueueSave = "QUEUED" & US() & "save " & prj.Name
End Function

' args: hwnd (hex) of a dialog button -> "clicked". Only buttons of this IDE process
' (AGENTS.md hard rule 8). Posted, so the click runs in the dialog's own loop after we return.
Private Function ClickButton(ByVal hexHwnd As String) As String
    Dim h As Long
    On Error GoTo Bad
    h = CLng("&H" & hexHwnd)
    On Error GoTo 0
    If Not IsOwnButton(h) Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " is not a button of this IDE"
    PostClick h
    ClickButton = "clicked " & hexHwnd
    Exit Function
Bad:
    Err.Raise E_BADREQUEST, , "bad window handle " & hexHwnd
End Function

' args: project, component, name, kind -> procStart US procBody US procCount US text
Private Function ProcText(ByVal prjName As String, ByVal cmpName As String, ByVal procName As String, ByVal kindText As String) As String
    Dim cm As VBIDE.CodeModule, kind As Long, first As Long, body As Long, n As Long
    If Not IsNumeric(kindText) Then Err.Raise E_BADREQUEST, , "procKind must be 0-3"
    kind = CLng(kindText)
    If kind < 0 Or kind > 3 Then Err.Raise E_BADREQUEST, , "procKind must be 0-3"
    Set cm = FindComponent(prjName, cmpName).CodeModule
    On Error GoTo NotFound
    first = cm.ProcStartLine(procName, kind)
    body = cm.ProcBodyLine(procName, kind)
    n = cm.ProcCountLines(procName, kind)
    On Error GoTo 0
    ProcText = first & US() & body & US() & n & US() & cm.Lines(first, n)
    Exit Function
NotFound:
    Err.Raise E_NOTFOUND, , "procedure not found: " & cmpName & "." & procName & " (kind " & kind & ")"
End Function
