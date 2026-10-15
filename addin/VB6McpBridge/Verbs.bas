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
    Case "CMD":        NeedArgs args, 1, 1: Dispatch = QueueCommand(args(0)): Exit Function
    Case "SETSEL":     NeedArgs args, 3, 3: payload = SetSelectionAt(args(0), args(1), args(2))
    Case "WHERE":      NeedArgs args, 0, 0: payload = WhereNow()
    Case "IMM":        NeedArgs args, 2, 2: Dispatch = QueueImmediate(args(0), args(1)): Exit Function
    Case "IMMSELECT":  NeedArgs args, 1, 1: payload = SelectImmediate(args(0))
    Case "IMMCOPY":    NeedArgs args, 0, 0: payload = CopyImmediate()
    Case "IMMPEEK":    NeedArgs args, 0, 0: payload = Clipboard.GetText(vbCFText)
    Case "IMMRESTORE": NeedArgs args, 0, 0: RestoreUserClipboard: payload = "restored"
    Case "APPCLICK":   NeedArgs args, 1, 1: payload = AppClick(args(0))
    Case "APPTEXT":    NeedArgs args, 2, 2: Dispatch = QueueAppText(args(0), args(1)): Exit Function
    Case "SNAP":       NeedArgs args, 2, 2: Dispatch = QueueSnap(args(0), args(1)): Exit Function
    Case "MOUSE":      NeedArgs args, 4, 6: payload = AppMouse(args)
    Case "SCROLL":     NeedArgs args, 4, 4: payload = AppScroll(args(0), args(1), args(2), args(3))
    Case "SCROLLPOS":  NeedArgs args, 1, 1: payload = AppScrollPos(args(0))
    Case "CALLSTACK":  NeedArgs args, 0, 0: Dispatch = QueueCallStack(): Exit Function
    Case "LISTITEMS":  NeedArgs args, 1, 1: payload = ListItems(args(0))
    Case "LISTSTATE":  NeedArgs args, 1, 1: payload = AppListState(args(0))
    Case "MAKE":       NeedArgs args, 1, 1: payload = MakeProject(args(0))
    Case "MAKESTATUS": NeedArgs args, 0, 0: payload = MakeStatus()
    Case "WRITES":     NeedArgs args, 1, 1: payload = WritesSince(CLng(Val(args(0))))
    Case "RELOAD":     NeedArgs args, 3, 3: Dispatch = QueueReload(args(0), args(1), args(2)): Exit Function
    Case "MAKEIDE":    NeedArgs args, 1, 1: Dispatch = QueueMakeInIde(args(0)): Exit Function
    Case "SELECT":     NeedArgs args, 2, 2: Dispatch = QueueSelect(args(0), args(1)): Exit Function
    Case "KEYS":       NeedArgs args, 2, 2: payload = AppKeys(args(0), args(1), False)
    Case "KEYSREAL":   NeedArgs args, 2, 2: Dispatch = "QUEUED" & US() & AppKeys(args(0), args(1), True): Exit Function
    Case "MOUSEREAL":  NeedArgs args, 4, 6: Dispatch = "QUEUED" & US() & AppMouseReal(args): Exit Function
    Case "MOUSERESTORE": NeedArgs args, 0, 0: payload = "restored" & US() & RestoreRealMouse()
    Case "DESIGN":     Dispatch = QueueDesign(args): Exit Function
    Case "RESULT":     NeedArgs args, 1, 1: payload = DesignResult(args(0))
    Case "REFS":       NeedArgs args, 1, 1: payload = ListRefs(args(0))
    Case "REFADD":     NeedArgs args, 3, 5: payload = AddRef(args)
    Case "REFDEL":     NeedArgs args, 2, 2: payload = RemoveRef(args(0), args(1))
    Case "TOOLBOX":    NeedArgs args, 2, 3: payload = AddToolbox(args)
    Case "REGSVR":     NeedArgs args, 2, 2: payload = RegisterFile(args(0), LCase$(args(1)))
    Case "REGSTATUS":  NeedArgs args, 1, 1: payload = RegsvrStatus(CLng(Val(args(0))))
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

Public Function FindProject(ByVal prjName As String) As VBIDE.VBProject
    Dim p As Long
    For p = 1 To g_VBE.VBProjects.Count
        If StrComp(g_VBE.VBProjects(p).Name, prjName, vbTextCompare) = 0 Then Set FindProject = g_VBE.VBProjects(p): Exit Function
    Next
    Err.Raise E_NOTFOUND, , "project not found: " & prjName
End Function

Public Function FindComponent(ByVal prjName As String, ByVal cmpName As String) As VBIDE.VBComponent
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

' ---------- debugging (Phase 2; spikes 3, 5, 6/8) ----------

' args: command id, or "fullcompile" -> QUEUED. Only an allow-list of debug commands (hard rule 8:
' no generic "execute any command"). Ids from the command-bar dump (spike 5); "Start With Full
' Compile" shares the generic id 746, so it's found by menu path. Deferred (hard rule 14): these
' commands run user code, and a break under the request handler deadlocks the IDE (spike 3).
Private Function QueueCommand(ByVal cmd As String) As String
    If LCase$(cmd) = "fullcompile" Then
        Enqueue "FULLCOMPILE"
        QueueCommand = "QUEUED" & US() & "Start With Full Compile"
        Exit Function
    End If
    Select Case cmd
    Case "186", "189", "228", "515", "188", "194", "2559", "1811", "51", "579"
    Case Else: Err.Raise E_BADREQUEST, , "command " & cmd & " is not an allowed debug command"
    End Select
    Enqueue "CMD" & US() & cmd
    QueueCommand = "QUEUED" & US() & "command " & cmd
End Function

' args: project, component, line -> "selected". Shows the code pane and puts the cursor at the
' start of the line (needed before Toggle Breakpoint / Run To Cursor). No user code runs.
Private Function SetSelectionAt(ByVal prjName As String, ByVal cmpName As String, ByVal lineText As String) As String
    Dim cm As VBIDE.CodeModule, ln As Long
    If Not IsNumeric(lineText) Then Err.Raise E_BADREQUEST, , "line must be a number"
    ln = CLng(lineText)
    Set cm = FindComponent(prjName, cmpName).CodeModule
    If ln < 1 Or ln > cm.CountOfLines Then Err.Raise E_BADREQUEST, , "line " & ln & " is outside the module (" & cm.CountOfLines & " lines)"
    cm.CodePane.Show
    cm.CodePane.SetSelection ln, 1, ln, 1
    SetSelectionAt = "selected " & cmpName & ":" & ln
End Function

' -> project US component US startLine US startCol US endLine US endCol US procName (active code pane)
Private Function WhereNow() As String
    Dim pane As VBIDE.CodePane, cm As VBIDE.CodeModule, sl As Long, sc As Long, el As Long, ec As Long, kind As VBIDE.vbext_ProcKind, proc As String
    Set pane = g_VBE.ActiveCodePane
    If pane Is Nothing Then Err.Raise E_NOTFOUND, , "no active code pane"
    Set cm = pane.CodeModule
    pane.GetSelection sl, sc, el, ec
    On Error Resume Next
    proc = cm.ProcOfLine(sl, kind)
    On Error GoTo 0
    WhereNow = cm.Parent.Collection.Parent.Name & US() & cm.Parent.Name & US() & sl & US() & sc & US() & el & US() & ec & US() & proc
End Function

' args: hwnd of the Immediate window (hex), text -> QUEUED. Typing into the Immediate window runs
' code (e.g. "? x" evaluates, "Call Foo" runs Foo), so it's deferred (hard rule 14).
Private Function QueueImmediate(ByVal hexHwnd As String, ByVal text As String) As String
    If Not IsOwnWindowOfClass(HexToLong(hexHwnd), "VbaWindow") Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " is not a VbaWindow of this IDE"
    If InStr(text, vbCr) > 0 Or InStr(text, vbLf) > 0 Then Err.Raise E_BADREQUEST, , "Immediate input must be a single line"
    Enqueue "IMM" & US() & hexHwnd & US() & text
    QueueImmediate = "QUEUED" & US() & "immediate"
End Function

' args: hwnd of the Immediate window (hex) -> the window's whole text. The Immediate window is
' custom-drawn, so WM_GETTEXT returns only its caption (spike 5): focus it, Select All + Copy,
' read the clipboard, and put the user's clipboard text back. No user code runs.
' Reading the Immediate window is a sequence of requests driven by the server, with the IDE idle in
' between (Phase 2 live tests):
'   IMMSELECT  save + clear the user's clipboard, focus the window, Select All
'   IMMCOPY    Copy ("copied"), or "empty" when Copy is disabled (nothing to select)
'   IMMPEEK    the clipboard text, repeated until the copy has arrived
'   IMMRESTORE put the user's clipboard text back
' Why: VB6 updates a command's Enabled state only when idle (Copy right after Select All in one
' handler was still disabled), and Copy itself completes asynchronously (reading the clipboard right
' after Execute found it empty, and the late copy then overwrote the restored user text).
' Focus uses the Win32 SetFocus on the handle (we're on the IDE thread).

Private Function SelectImmediate(ByVal hexHwnd As String) As String
    Dim h As Long, c As Object
    h = HexToLong(hexHwnd)
    If Not IsOwnWindowOfClass(h, "VbaWindow") Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " is not a VbaWindow of this IDE"
    Set c = g_VBE.CommandBars.FindControl(, 756)           ' Edit > Select All
    SaveUserClipboard
    Clipboard.Clear
    FocusWindow h
    g_readHwnd = h
    If Not c.Enabled Then Err.Raise E_CONFLICT, , "Select All is disabled right now; retry"
    c.Execute
    SelectImmediate = "selected"
End Function

Private Function CopyImmediate() As String
    Dim c As Object
    Set c = g_VBE.CommandBars.FindControl(, 19)            ' Edit > Copy
    If Not c.Enabled Then CopyImmediate = "empty": Exit Function
    c.Execute
    CopyImmediate = "copied"
End Function

' ---------- the running program (Phase 3) ----------

' args: hwnd -> "clicked". Command buttons, check boxes and option buttons of this process only
' (hard rule 8). BM_CLICK is posted, so the control's Click event runs after this handler returns
' (hard rule 14). The server verifies the effect.
Private Function AppClick(ByVal hexHwnd As String) As String
    Dim h As Long, cls As String
    h = HexToLong(hexHwnd)
    cls = OwnWindowClass(h)
    Select Case cls
    Case "ThunderCommandButton", "ThunderRT6CommandButton", "ThunderCheckBox", "ThunderRT6CheckBox", _
         "ThunderOptionButton", "ThunderRT6OptionButton", "Button"
    Case "": Err.Raise E_BADREQUEST, , "window " & hexHwnd & " doesn't belong to this IDE process"
    Case Else: Err.Raise E_BADREQUEST, , "window " & hexHwnd & " is a " & cls & ", not a button, check box or option button"
    End Select
    PostClick h
    AppClick = "clicked " & hexHwnd & " (" & cls & ")"
End Function

' args: hwnd, text -> QUEUED. Sets a text box's text with WM_SETTEXT. Deferred: it fires the
' control's Change event (user code).
Private Function QueueAppText(ByVal hexHwnd As String, ByVal text As String) As String
    Dim cls As String
    cls = OwnWindowClass(HexToLong(hexHwnd))
    Select Case cls
    Case "ThunderTextBox", "ThunderRT6TextBox", "Edit"
    Case "": Err.Raise E_BADREQUEST, , "window " & hexHwnd & " doesn't belong to this IDE process"
    Case Else: Err.Raise E_BADREQUEST, , "window " & hexHwnd & " is a " & cls & ", not a text box"
    End Select
    Enqueue "APPTEXT" & US() & hexHwnd & US() & text
    QueueAppText = "QUEUED" & US() & "set text"
End Function

' args: hwnd, id -> QUEUED. Captures the window into %LOCALAPPDATA%\vb6mcp\shots\<id>.bmp.
' The id is a plain name, never a path (hard rule 8). Deferred: painting can run user code.
Private Function QueueSnap(ByVal hexHwnd As String, ByVal id As String) As String
    Dim i As Long, ch As String
    If OwnWindowClass(HexToLong(hexHwnd)) = "" Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " doesn't belong to this IDE process"
    If Len(id) = 0 Or Len(id) > 40 Then Err.Raise E_BADREQUEST, , "id must be 1-40 characters"
    For i = 1 To Len(id)
        ch = Mid$(id, i, 1)
        If Not (ch Like "[A-Za-z0-9_-]") Then Err.Raise E_BADREQUEST, , "id may only contain letters, digits, _ and -"
    Next
    Enqueue "SNAP" & US() & hexHwnd & US() & id
    QueueSnap = "QUEUED" & US() & ShotsDir() & "\" & id & ".bmp"
End Function

' args: hwnd, action (click|dblclick|rclick|drag|wheel), x, y [, x2, y2] -> "mouse ...".
' Coordinates are client coordinates of hwnd; for drag (x2, y2) is the end point, for wheel x2 is
' the number of notches (negative = down). Windows of this process only (hard rule 8).
Private Function AppMouse(args() As String) As String
    Dim h As Long, action As String, i As Long, n(3) As Long
    h = HexToLong(args(0))
    If OwnWindowClass(h) = "" Then Err.Raise E_BADREQUEST, , "window " & args(0) & " doesn't belong to this IDE process"
    action = LCase$(args(1))
    Select Case action
    Case "click", "dblclick", "rclick", "drag", "wheel"
    Case Else: Err.Raise E_BADREQUEST, , "action must be click, dblclick, rclick, drag or wheel"
    End Select
    For i = 2 To UBound(args)
        If Not IsNumeric(args(i)) Then Err.Raise E_BADREQUEST, , "coordinates must be numbers"
        n(i - 2) = CLng(args(i))
        If n(i - 2) < -32768 Or n(i - 2) > 32767 Then Err.Raise E_BADREQUEST, , "coordinate out of range"
    Next
    If (action = "drag" Or action = "wheel") And UBound(args) < 4 Then Err.Raise E_BADREQUEST, , action & " needs a fifth argument"
    PostMouse h, action, n(0), n(1), n(2), n(3)
    AppMouse = "mouse " & action & " " & args(0)
End Function

' args: hwnd, v|h, code 0-7 (SB_LINEUP .. SB_BOTTOM), count 1-100 -> "scrolled ...".
Private Function AppScroll(ByVal hexHwnd As String, ByVal bar As String, ByVal codeText As String, ByVal countText As String) As String
    Dim h As Long, cls As String, code As Long, count As Long
    h = HexToLong(hexHwnd)
    cls = OwnWindowClass(h)
    If cls = "" Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " doesn't belong to this IDE process"
    If bar <> "v" And bar <> "h" Then Err.Raise E_BADREQUEST, , "bar must be v or h"
    If Not IsNumeric(codeText) Or Not IsNumeric(countText) Then Err.Raise E_BADREQUEST, , "code and count must be numbers"
    code = CLng(codeText): count = CLng(countText)
    If code < 0 Or code > 7 Then Err.Raise E_BADREQUEST, , "code must be 0-7"
    If count < 1 Or count > 100 Then Err.Raise E_BADREQUEST, , "count must be 1-100"
    PostScroll h, (bar = "v"), code, count, (InStr(1, cls, "ScrollBar", vbTextCompare) > 0)
    AppScroll = "scrolled " & hexHwnd & " (" & cls & ")"
End Function

' args: hwnd -> min US max US pos of a scroll bar control of this process.
Private Function AppScrollPos(ByVal hexHwnd As String) As String
    Dim h As Long
    h = HexToLong(hexHwnd)
    If InStr(1, OwnWindowClass(h), "ScrollBar", vbTextCompare) = 0 Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " isn't a scroll bar control of this IDE process"
    AppScrollPos = ScrollBarPos(h)
End Function

' args as MOUSE (click|dblclick|rclick|drag) -> "sent". Real input with SendInput (spike 17):
' brings the form to the front and moves the cursor; the server always follows with MOUSERESTORE.
Private Function AppMouseReal(args() As String) As String
    Dim h As Long, action As String, i As Long, n(3) As Long
    h = HexToLong(args(0))
    If OwnWindowClass(h) = "" Then Err.Raise E_BADREQUEST, , "window " & args(0) & " doesn't belong to this IDE process"
    action = LCase$(args(1))
    Select Case action
    Case "click", "dblclick", "rclick", "drag"
    Case Else: Err.Raise E_BADREQUEST, , "action must be click, dblclick, rclick or drag"
    End Select
    For i = 2 To UBound(args)
        If Not IsNumeric(args(i)) Then Err.Raise E_BADREQUEST, , "coordinates must be numbers"
        n(i - 2) = CLng(args(i))
        If n(i - 2) < -32768 Or n(i - 2) > 32767 Then Err.Raise E_BADREQUEST, , "coordinate out of range"
    Next
    If action = "drag" And UBound(args) < 5 Then Err.Raise E_BADREQUEST, , "drag needs x2 and y2"
    ' Deferred (hard rule 14): bringing the form to the front activates it, which raises its
    ' Activate event (user code). The outcome is returned by MOUSERESTORE.
    Enqueue "MOUSEREAL" & US() & Hex$(h) & US() & action & US() & n(0) & US() & n(1) & US() & n(2) & US() & n(3)
    AppMouseReal = "queued"
End Function

' -> QUEUED. Opens View > Call Stack (a modal dialog, so deferred; the server reads it with LISTITEMS
' and closes it with CLICK).
Private Function QueueCallStack() As String
    If g_VBE.CommandBars.FindControl(, 189) Is Nothing Then Err.Raise E_NOTFOUND, , "debug commands not found"
    Enqueue "CALLSTACK"
    QueueCallStack = "QUEUED" & US() & "call stack"
End Function

' args: hwnd of a list box of this process -> items separated by RS. Read in-process: UIPI blocks
' LB_* messages from the non-elevated server (Phase 2, Call Stack dialog). Read-only.
Private Function ListItems(ByVal hexHwnd As String) As String
    Dim h As Long, cls As String
    h = HexToLong(hexHwnd)
    cls = OwnWindowClass(h)
    If cls = "" Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " doesn't belong to this IDE process"
    If InStr(1, cls, "ListBox", vbTextCompare) = 0 Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " is a " & cls & ", not a list box"
    ListItems = ListBoxItems(h)

End Function

' args: hwnd of a list box or combo box of this process -> kind US count US selected US items (RS).
Private Function AppListState(ByVal hexHwnd As String) As String
    Dim h As Long, cls As String
    h = HexToLong(hexHwnd)
    cls = OwnWindowClass(h)
    If cls = "" Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " doesn't belong to this IDE process"
    If InStr(1, cls, "ListBox", vbTextCompare) = 0 And InStr(1, cls, "ComboBox", vbTextCompare) = 0 Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " is a " & cls & ", not a list box or combo box"
    AppListState = ListState(h, InStr(1, cls, "ComboBox", vbTextCompare) > 0)
End Function

' args: hwnd, index -> QUEUED. Deferred: the selection raises the control's Click (user code).
Private Function QueueSelect(ByVal hexHwnd As String, ByVal indexText As String) As String
    Dim h As Long, cls As String, n As Long, isCombo As Boolean
    h = HexToLong(hexHwnd)
    cls = OwnWindowClass(h)
    If cls = "" Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " doesn't belong to this IDE process"
    isCombo = (InStr(1, cls, "ComboBox", vbTextCompare) > 0)
    If Not isCombo And InStr(1, cls, "ListBox", vbTextCompare) = 0 Then Err.Raise E_BADREQUEST, , "window " & hexHwnd & " is a " & cls & ", not a list box or combo box"
    If Not IsNumeric(indexText) Then Err.Raise E_BADREQUEST, , "index must be a number"
    n = CLng(Split(ListState(h, isCombo), US())(1))
    If CLng(indexText) < 0 Or CLng(indexText) >= n Then Err.Raise E_BADREQUEST, , "index " & indexText & " is out of range (" & n & " items)"
    Enqueue "SELECT" & US() & hexHwnd & US() & CLng(indexText)
    QueueSelect = "QUEUED" & US() & "select " & indexText
End Function

' args: hwnd (or 0 = the program's focused control), events -> "keys N" / "sent". Windows of this
' process only (hard rule 8). Events: S<0-7>, D<vk>, U<vk>, C<1-255> separated by ";" (PostKeys).
Private Function AppKeys(ByVal hexHwnd As String, ByVal events As String, ByVal realInput As Boolean) As String
    Dim h As Long, tok() As String, i As Long, kind As String, v As String
    h = HexToLong(hexHwnd)
    If h = 0 Then h = FocusedWindow()
    If h = 0 Then Err.Raise E_CONFLICT, , "no control of the program has the focus; pass the control (from app_tree)"
    If OwnWindowClass(h) = "" Then Err.Raise E_BADREQUEST, , "window " & Hex$(h) & " doesn't belong to this IDE process"
    tok = Split(events, ";")
    If UBound(tok) > 500 Then Err.Raise E_BADREQUEST, , "too many key events (at most 500)"
    For i = 0 To UBound(tok)
        kind = Left$(tok(i), 1): v = Mid$(tok(i), 2)
        If Len(v) = 0 Or Len(v) > 3 Or Not IsNumeric(v) Or InStr(v, ".") > 0 Or InStr(v, "-") > 0 Then Err.Raise E_BADREQUEST, , "bad key event '" & tok(i) & "'"
        Select Case kind
        Case "S": If CLng(v) > 7 Then Err.Raise E_BADREQUEST, , "bad modifier set '" & tok(i) & "'"
        Case "D", "U": If CLng(v) < 1 Or CLng(v) > 254 Then Err.Raise E_BADREQUEST, , "bad virtual key '" & tok(i) & "'"
        Case "C": If CLng(v) < 1 Or CLng(v) > 255 Then Err.Raise E_BADREQUEST, , "bad character '" & tok(i) & "'"
        Case Else: Err.Raise E_BADREQUEST, , "bad key event '" & tok(i) & "'"
        End Select
    Next
    If realInput Then
        ' Deferred (hard rule 14): focusing the control raises GotFocus/LostFocus (user code)
        Enqueue "KEYSREAL" & US() & Hex$(h) & US() & events
        AppKeys = "queued" & US() & Hex$(h)
    Else
        PostKeys h, events
        AppKeys = "keys " & (UBound(tok) + 1) & US() & Hex$(h)
    End If
End Function

' args: full path of a .vbp -> pid US log path US vbp. VB6.EXE /make refuses a project that is open in
' any IDE ("The file '...' is already open.", Phase 4 spike), so this is for closed projects; open ones
' are built in their IDE (MAKEIDE). The path must be under a workspace root listed in
' %LOCALAPPDATA%\vb6mcp\roots.txt (written by install-server.ps1) and contain no "..". The token is the
' real boundary (anything that can read it can already drive this IDE); the roots check is defense
' in depth, because an elevated /make registers components machine-wide (hard rule 8).
Private Function MakeProject(ByVal vbp As String) As String
    Dim logPath As String, pid As Long, p As Long, name As String
    If LCase$(Right$(vbp, 4)) <> ".vbp" Then Err.Raise E_BADREQUEST, , "not a .vbp: " & vbp
    If InStr(vbp, "..") > 0 Or Mid$(vbp, 2, 2) <> ":\" Then Err.Raise E_BADREQUEST, , "the path must be absolute, without '..'"
    If Not UnderRoots(vbp) Then Err.Raise E_BADREQUEST, , vbp & " isn't under the workspace roots (" & RootsFile() & ")"
    If Len(Dir$(vbp)) = 0 Then Err.Raise E_NOTFOUND, , "project file not found: " & vbp
    For p = 1 To g_VBE.VBProjects.Count
        If StrComp(g_VBE.VBProjects(p).FileName, vbp, vbTextCompare) = 0 Then Err.Raise E_CONFLICT, , "the project is open in this IDE; /make can't build it (use MAKEIDE)"
    Next
    name = Mid$(vbp, InStrRev(vbp, "\") + 1)
    name = Left$(name, Len(name) - 4)
    EnsureDir BaseDir() & "\builds"
    logPath = BaseDir() & "\builds\" & Format$(Now, "yyyymmdd-hhnnss") & "-" & name & ".log"
    pid = StartBuild(vbp, logPath)
    If pid = 0 Then Err.Raise E_CONFLICT, , "VB6.EXE /make couldn't be started"
    MakeProject = pid & US() & logPath & US() & vbp
End Function

' args: path, r|u -> pid. Starts the 32-bit regsvr32 elevated (com_register / com_unregister, hard rule 13).
' Same path rules as MAKE: absolute, no "..", under the workspace roots (checked here too, on the elevated
' side, hard rule 8), and only .dll/.ocx. The server audits the call and checks the registry afterwards.
Private Function RegisterFile(ByVal path As String, ByVal mode As String) As String
    Dim ext As String, pid As Long
    ext = LCase$(Right$(path, 4))
    If ext <> ".dll" And ext <> ".ocx" Then Err.Raise E_BADREQUEST, , "only .dll and .ocx files can be registered: " & path
    If mode <> "r" And mode <> "u" Then Err.Raise E_BADREQUEST, , "mode must be r or u"
    If InStr(path, "..") > 0 Or Mid$(path, 2, 2) <> ":\" Or InStr(path, """") > 0 Then Err.Raise E_BADREQUEST, , "the path must be absolute, without '..' or quotes"
    If Not UnderRoots(path) Then Err.Raise E_BADREQUEST, , path & " isn't under the workspace roots (" & RootsFile() & ")"
    If Len(Dir$(path)) = 0 Then Err.Raise E_NOTFOUND, , "file not found: " & path
    pid = StartRegsvr(path, (mode = "u"))
    If pid = 0 Then Err.Raise E_CONFLICT, , "regsvr32 couldn't be started"
    RegisterFile = CStr(pid)
End Function

' The workspace roots file. An installed add-in (Program Files, Phase 6) reads roots.txt next to its DLL: that folder is
' admin-writable only, so ordinary processes can't widen what this elevated add-in builds, registers or loads (with a
' user-writable roots file, any process could get files registered elevated). A development build has no file there
' and reads %LOCALAPPDATA%\vb6mcp\roots.txt (install-server.ps1); its DLL folder is user-writable anyway.
Public Function RootsFile() As String
    If Len(Dir$(App.Path & "\roots.txt")) > 0 Then RootsFile = App.Path & "\roots.txt" Else RootsFile = BaseDir() & "\roots.txt"
End Function

' True if path is inside one of the roots in roots.txt (one absolute folder per line).
Public Function UnderRoots(ByVal path As String) As Boolean
    Dim f As Integer, root As String
    On Error GoTo Done
    If Len(Dir$(RootsFile())) = 0 Then Exit Function
    f = FreeFile
    Open RootsFile() For Input As #f
    Do While Not EOF(f)
        Line Input #f, root
        root = Trim$(root)
        If Len(root) > 2 Then
            If Right$(root, 1) <> "\" Then root = root & "\"
            If StrComp(Left$(path, Len(root)), root, vbTextCompare) = 0 Then UnderRoots = True: Exit Do
        End If
    Loop
Done:
    If f <> 0 Then Close #f
End Function

' -> outcome of the last MAKEIDE: "" (none yet), "running", "ok", or "failed US message".
' MakeCompiledFile shows no dialog for a compile error: it raises E_FAIL ("Method '~' of object '~'
' failed"), so the server needs this to know (Phase 4 spike).
Private Function MakeStatus() As String
    MakeStatus = LastMakeResult()
End Function

' args: project name -> QUEUED US output file. VBProject.MakeCompiledFile builds from IDE memory
' (unsaved edits included) to the project's BuildFileName. Deferred: it compiles and can show
' modal dialogs (compile errors, compatibility), which must not block the request handler.
Private Function QueueMakeInIde(ByVal prjName As String) As String
    Dim prj As VBIDE.VBProject
    Set prj = FindProject(prjName)
    If g_VBE.CommandBars.FindControl(, 228).Enabled Then Err.Raise E_CONFLICT, , "the program is running; end it first"
    MarkMakeQueued
    Enqueue "MAKEIDE" & US() & prj.Name
    QueueMakeInIde = "QUEUED" & US() & prj.BuildFileName
End Function

' args: project, component, mode (clean | discard) -> QUEUED. Reloads a component from its file
' (VBComponent.Reload, spike 8): refreshes code and designer. "clean" is the sync engine's automatic
' reload of an unmodified component; "discard" throws away unsaved IDE edits and is only sent by
' sync_take_disk with explicit confirmation (hard rule 9), after the server kept a recovery copy.
' Deferred: reloading a form can create design-time instances of user controls (their code runs).
' Design mode only (Reload fails with error 5 in break mode, spike 8).
Private Function QueueReload(ByVal prjName As String, ByVal cmpName As String, ByVal mode As String) As String
    Dim cmp As VBIDE.VBComponent
    If mode <> "clean" And mode <> "discard" Then Err.Raise E_BADREQUEST, , "mode must be clean or discard"
    Set cmp = FindComponent(prjName, cmpName)
    If g_VBE.CommandBars.FindControl(, 228).Enabled Then Err.Raise E_CONFLICT, , "the program is running or in break mode; reload in design mode"
    If mode = "clean" And cmp.IsDirty Then Err.Raise E_CONFLICT, , cmpName & " has unsaved changes; not reloaded"
    If cmp.FileCount = 0 Then Err.Raise E_CONFLICT, , cmpName & " has no file"
    Enqueue "RELOAD" & US() & prjName & US() & cmpName & US() & mode
    QueueReload = "QUEUED" & US() & "reload " & cmpName
End Function

Private Function HexToLong(ByVal hexText As String) As Long
    On Error GoTo Bad
    HexToLong = CLng("&H" & hexText)
    Exit Function
Bad:
    Err.Raise E_BADREQUEST, , "bad window handle " & hexText
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
