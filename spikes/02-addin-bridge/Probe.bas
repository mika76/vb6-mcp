Attribute VB_Name = "Probe"
' Spikes 3 + 5: probe verbs. Each one reports what worked and the exact error when it did not.
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Public Const WM_DEFER As Long = &H400 + 101   ' WM_USER + 101: run a queued IDE command

' Deferred command queue (one slot is enough for the spike)
Private m_pendingId As Long
Private m_pendingPath As String

' Returns Empty when the verb is not a probe verb
Public Function Dispatch(ByVal verb As String, ByVal args As String) As Variant
    Dim a() As String
    a = Split(args, vbTab)
    Select Case verb
    Case "STATE":    Dispatch = ProbeState()
    Case "CMDBARS":  Dispatch = CmdBars()
    Case "EXEC":     Dispatch = ExecControl(CLng(a(0)))
    Case "WINDOWS":  Dispatch = WindowsText()
    Case "FOCUS":    Dispatch = FocusWindow(a(0))
    Case "SENDKEY":  Dispatch = SendKey(CLng("&H" & a(0)), CLng(a(1)), False)
    Case "POSTKEY":  Dispatch = SendKey(CLng("&H" & a(0)), CLng(a(1)), True)
    Case "CHARS":    Dispatch = SendChars(CLng("&H" & a(0)), a(1))
    Case "CLIPGET":  Dispatch = Clipboard.GetText()
    Case "CLIPSET":  Clipboard.Clear: Clipboard.SetText a(0): Dispatch = "ok"
    Case "INSERT":   Dispatch = InsertLine(a(0), a(1), CLng(a(2)), a(3))
    Case "DELETE":   Dispatch = DeleteLine(a(0), a(1), CLng(a(2)))
    Case "QEXEC":    Dispatch = QueueExec(CLng(a(0)), "")
    Case "QEXECPATH": Dispatch = QueueExec(0, a(0))
    Case "FINDPATH": Dispatch = DescribeControl(FindByPath(a(0)))
    Case "INFO":     Dispatch = ComponentInfo(a(0), a(1))
    Case "RELOAD":   Dispatch = ReloadComponent(a(0), a(1))
    Case "SETSEL":   Dispatch = SetSel(a(0), a(1), CLng(a(2)))
    Case Else:       Dispatch = Empty
    End Select
End Function

' ---------- STATE: what can VBIDE do in the current mode? ----------

Public Function ProbeState() As String
    Dim i As Long, out As String
    For i = 1 To 16
        out = out & ProbeItem(i) & vbCrLf
    Next
    ProbeState = out
End Function

Private Function ProbeItem(ByVal i As Long) As String
    Dim label As String, r As String
    On Error GoTo Fail
    Select Case i
    Case 1:  label = "VBE.Version": r = g_VBE.Version
    Case 2:  label = "VBProjects.Count": r = g_VBE.VBProjects.Count
    Case 3:  label = "ActiveVBProject.Name": r = g_VBE.ActiveVBProject.Name
    Case 4:  label = "ActiveVBProject.VBComponents.Count": r = g_VBE.ActiveVBProject.VBComponents.Count
    Case 5:  label = "SelectedVBComponent.Name": r = g_VBE.SelectedVBComponent.Name
    Case 6:  label = "SelectedVBComponent.IsDirty": r = g_VBE.SelectedVBComponent.IsDirty
    Case 7:  label = "ActiveCodePane.CodeModule.Parent.Name": r = g_VBE.ActiveCodePane.CodeModule.Parent.Name
    Case 8:  label = "ActiveCodePane.GetSelection": r = SelText()
    Case 9:  label = "ActiveCodePane.CodeModule.CountOfLines": r = g_VBE.ActiveCodePane.CodeModule.CountOfLines
    Case 10: label = "ActiveCodePane.CodeModule.Lines(1,3)": r = Replace(g_VBE.ActiveCodePane.CodeModule.Lines(1, 3), vbCrLf, "\n")
    Case 11: label = "Members": r = MembersText()
    Case 12: label = "ProcOfLine(cursor)": r = ProcAtCursor()
    Case 13: label = "Designer controls": r = DesignerText()
    Case 14: label = "Windows.Count": r = g_VBE.Windows.Count
    Case 15: label = "CommandBars.Count": r = g_VBE.CommandBars.Count
    Case 16: label = "VBProjects(1).VBComponents(1).CodeModule.CountOfLines": r = g_VBE.VBProjects(1).VBComponents(1).CodeModule.CountOfLines
    End Select
    ProbeItem = label & " = " & r
    Exit Function
Fail:
    ProbeItem = label & " !! error " & Err.Number & ": " & Err.Description
End Function

Private Function SelText() As String
    Dim sl As Long, sc As Long, el As Long, ec As Long
    g_VBE.ActiveCodePane.GetSelection sl, sc, el, ec
    SelText = sl & "," & sc & " - " & el & "," & ec
End Function

Private Function MembersText() As String
    Dim m As VBIDE.Member, s As String
    For Each m In g_VBE.ActiveCodePane.CodeModule.Members
        s = s & m.Name & "(type=" & m.Type & ",scope=" & m.Scope & ",line=" & m.CodeLocation & ",static=" & m.Static & ") "
    Next
    MembersText = s
End Function

Private Function ProcAtCursor() As String
    Dim sl As Long, sc As Long, el As Long, ec As Long, kind As VBIDE.vbext_ProcKind, nm As String
    Dim cm As VBIDE.CodeModule
    Set cm = g_VBE.ActiveCodePane.CodeModule
    g_VBE.ActiveCodePane.GetSelection sl, sc, el, ec
    nm = cm.ProcOfLine(sl, kind)
    If Len(nm) = 0 Then ProcAtCursor = "(declarations)": Exit Function
    ProcAtCursor = nm & " kind=" & kind & " start=" & cm.ProcStartLine(nm, kind) & " body=" & cm.ProcBodyLine(nm, kind) & " count=" & cm.ProcCountLines(nm, kind)
End Function

Private Function DesignerText() As String
    Dim c As VBIDE.VBComponent, ctl As VBIDE.VBControl, s As String
    Set c = g_VBE.SelectedVBComponent
    s = "component " & c.Name & " caption=" & c.Properties("Caption").Value & "; "
    For Each ctl In c.Designer.VBControls
        s = s & ctl.Properties("Name").Value & ":" & ctl.ClassName & " "
    Next
    DesignerText = s
End Function

' ---------- command bars ----------

Private Function CmdBars() As String
    Dim bar As Object, s As String
    On Error Resume Next
    For Each bar In g_VBE.CommandBars
        s = s & "BAR " & bar.Name & " visible=" & bar.Visible & vbCrLf
        DumpControls bar.Controls, 1, s
    Next
    CmdBars = s
End Function

Private Sub DumpControls(ByVal ctls As Object, ByVal depth As Long, s As String)
    Dim c As Object
    On Error Resume Next
    For Each c In ctls
        s = s & String$(depth * 2, " ") & c.Id & vbTab & c.Type & vbTab & "en=" & c.Enabled & vbTab & c.Caption & vbCrLf
        If c.Type = 10 And depth < 3 Then DumpControls c.Controls, depth + 1, s   ' 10 = msoControlPopup
    Next
End Sub

Private Function ExecControl(ByVal id As Long) As String
    Dim c As Object
    Set c = g_VBE.CommandBars.FindControl(, id)
    If c Is Nothing Then ExecControl = "no control with id " & id: Exit Function
    c.Execute
    ExecControl = "executed " & id & " (" & c.Caption & ")"
End Function

' ---------- windows and input ----------

Private Function WindowsText() As String
    Dim w As VBIDE.Window, s As String
    For Each w In g_VBE.Windows
        s = s & Hex$(w.HWnd) & vbTab & "type=" & w.Type & vbTab & "visible=" & w.Visible & vbTab & w.Caption & vbCrLf
    Next
    WindowsText = s
End Function

Private Function FocusWindow(ByVal caption As String) As String
    Dim w As VBIDE.Window
    For Each w In g_VBE.Windows
        If StrComp(w.Caption, caption, vbTextCompare) = 0 Then
            w.Visible = True
            w.SetFocus
            FocusWindow = "focused " & Hex$(w.HWnd)
            Exit Function
        End If
    Next
    FocusWindow = "no window " & caption
End Function

Private Function SendKey(ByVal hWnd As Long, ByVal vk As Long, ByVal post As Boolean) As String
    If post Then
        PostMessage hWnd, WM_KEYDOWN, vk, 0
        PostMessage hWnd, WM_KEYUP, vk, &HC0000000
    Else
        SendMessage hWnd, WM_KEYDOWN, vk, 0
        SendMessage hWnd, WM_KEYUP, vk, &HC0000000
    End If
    SendKey = "sent vk " & vk & IIf(post, " (posted)", " (sent)")
End Function

Private Function SendChars(ByVal hWnd As Long, ByVal text As String) As String
    Dim i As Long
    For i = 1 To Len(text)
        SendMessage hWnd, WM_CHAR, Asc(Mid$(text, i, 1)), 0
    Next
    SendChars = "sent " & Len(text) & " chars"
End Function

' ---------- writes (undo test) ----------

Private Function FindModule(ByVal prjName As String, ByVal cmpName As String) As VBIDE.CodeModule
    Dim p As Long, c As Long, prj As VBIDE.VBProject, cmp As VBIDE.VBComponent
    For p = 1 To g_VBE.VBProjects.Count
        Set prj = g_VBE.VBProjects(p)
        If StrComp(prj.Name, prjName, vbTextCompare) = 0 Then
            For c = 1 To prj.VBComponents.Count
                Set cmp = prj.VBComponents(c)
                If StrComp(cmp.Name, cmpName, vbTextCompare) = 0 Then Set FindModule = cmp.CodeModule: Exit Function
            Next
        End If
    Next
    Err.Raise vbObjectError + 2, , "component not found: " & prjName & "." & cmpName
End Function

Private Function InsertLine(ByVal prj As String, ByVal cmp As String, ByVal lineNo As Long, ByVal text As String) As String
    Dim cm As VBIDE.CodeModule
    Set cm = FindModule(prj, cmp)
    cm.InsertLines lineNo, text
    InsertLine = "inserted at " & lineNo & "; now " & cm.CountOfLines & " lines"
End Function

Private Function DeleteLine(ByVal prj As String, ByVal cmp As String, ByVal lineNo As Long) As String
    Dim cm As VBIDE.CodeModule
    Set cm = FindModule(prj, cmp)
    cm.DeleteLines lineNo, 1
    DeleteLine = "deleted line " & lineNo & "; now " & cm.CountOfLines & " lines"
End Function

' ---------- deferred execution (rule 14) ----------
' The request handler only queues; the command runs from a fresh dispatch of our own
' window message, with no add-in request frames on the stack below it.

Private Function QueueExec(ByVal id As Long, ByVal menuPath As String) As String
    m_pendingId = id
    m_pendingPath = menuPath
    PostMessage DeferHwnd(), WM_DEFER, 0, 0
    QueueExec = "queued " & IIf(id <> 0, CStr(id), menuPath)
End Function

Public Sub RunDeferred()
    Dim c As Object
    On Error GoTo Fail
    If m_pendingId <> 0 Then Set c = g_VBE.CommandBars.FindControl(, m_pendingId) Else Set c = FindByPath(m_pendingPath)
    If c Is Nothing Then LogLine "deferred: control not found " & m_pendingId & m_pendingPath: Exit Sub
    LogLine "deferred: executing " & c.Id & " " & c.Caption
    c.Execute
    LogLine "deferred: Execute returned"
    Exit Sub
Fail:
    LogLine "deferred error " & Err.Number & ": " & Err.Description
End Sub

' "Run|Start With Full Compile": menu bar popup caption | item caption, '&' ignored
Private Function FindByPath(ByVal menuPath As String) As Object
    Dim parts() As String, bar As Object, i As Long, j As Long, pop As Object, item As Object
    parts = Split(menuPath, "|")
    Set bar = g_VBE.CommandBars("Menu Bar")
    For i = 1 To bar.Controls.Count
        Set pop = bar.Controls(i)
        If StrComp(Replace(pop.Caption, "&", ""), parts(0), vbTextCompare) = 0 Then
            For j = 1 To pop.Controls.Count
                Set item = pop.Controls(j)
                If StrComp(Replace(item.Caption, "&", ""), parts(1), vbTextCompare) = 0 Then Set FindByPath = item: Exit Function
            Next
        End If
    Next
End Function

Private Function DescribeControl(ByVal c As Object) As String
    If c Is Nothing Then DescribeControl = "not found" Else DescribeControl = "id=" & c.Id & " caption=" & c.Caption & " enabled=" & c.Enabled
End Function

' ---------- spike 8: component files and reload ----------

Private Function FindComponent(ByVal prjName As String, ByVal cmpName As String) As VBIDE.VBComponent
    Dim p As Long, c As Long, prj As VBIDE.VBProject
    For p = 1 To g_VBE.VBProjects.Count
        Set prj = g_VBE.VBProjects(p)
        If StrComp(prj.Name, prjName, vbTextCompare) = 0 Then
            For c = 1 To prj.VBComponents.Count
                If StrComp(prj.VBComponents(c).Name, cmpName, vbTextCompare) = 0 Then Set FindComponent = prj.VBComponents(c): Exit Function
            Next
        End If
    Next
    Err.Raise vbObjectError + 2, , "component not found: " & prjName & "." & cmpName
End Function

Private Function ComponentInfo(ByVal prj As String, ByVal cmp As String) As String
    Dim c As VBIDE.VBComponent, i As Integer, s As String
    Set c = FindComponent(prj, cmp)
    s = "dirty=" & c.IsDirty & " lines=" & c.CodeModule.CountOfLines & " files=" & c.FileCount
    For i = 1 To c.FileCount
        s = s & " [" & c.FileNames(i) & "]"
    Next
    ComponentInfo = s
End Function

Private Function ReloadComponent(ByVal prj As String, ByVal cmp As String) As String
    Dim c As VBIDE.VBComponent, before As String
    Set c = FindComponent(prj, cmp)
    before = "dirty=" & c.IsDirty & " lines=" & c.CodeModule.CountOfLines
    c.Reload
    Set c = FindComponent(prj, cmp)    ' the object may be replaced by Reload
    ReloadComponent = "before: " & before & " | after: dirty=" & c.IsDirty & " lines=" & c.CodeModule.CountOfLines
End Function

Private Function SetSel(ByVal prj As String, ByVal cmp As String, ByVal lineNo As Long) As String
    Dim c As VBIDE.VBComponent
    Set c = FindComponent(prj, cmp)
    c.CodeModule.CodePane.Show
    c.CodeModule.CodePane.SetSelection lineNo, 1, lineNo, 1
    SetSel = "selection at line " & lineNo
End Function

' ---------- used by the build-event hooks in Connect ----------

Public Function AccessTest() As String
    On Error GoTo Fail
    AccessTest = "VBProjects(1).VBComponents(1) lines=" & g_VBE.VBProjects(1).VBComponents(1).CodeModule.CountOfLines
    Exit Function
Fail:
    AccessTest = "access error " & Err.Number & ": " & Err.Description
End Function
