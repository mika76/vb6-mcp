Attribute VB_Name = "Defer"
' Deferred execution (AGENTS.md hard rule 14, spikes 3 and 6/8).
' Anything that can run user code or show modal UI must not run inside the socket handler:
' a break or a modal dialog there leaves the request handler on the stack (spike 3 deadlock:
' every debug command disabled, the IDE had to be killed). The handler queues the action and
' replies QUEUED; the action runs from a fresh dispatch of our own window message.
Option Explicit

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_DEFER As Long = &H400 + 101      ' WM_USER + 101

Private m_queue As Collection                     ' of String: "ACTION" US arg1 US arg2 ...
' Outcome of the last MakeCompiledFile (MAKESTATUS)
Private m_makeResult As String

Public Sub DeferInit()
    Set m_queue = New Collection
End Sub

' Queues an action and posts the window message that will run it.
Public Sub Enqueue(ByVal action As String)
    If m_queue Is Nothing Then DeferInit
    m_queue.Add action
    PostMessage ListenerHwnd(), WM_DEFER, 0, 0
End Sub

' Called from the listener's window procedure on WM_DEFER: runs ONE queued action. Each Enqueue
' posts one WM_DEFER, so actions keep their place among other posted messages: KEYS relies on a
' KEYSTATE action running between the key messages it posted (Phase 3).
Public Sub RunQueued()
    Dim action As String, f() As String
    On Error GoTo Fail
    If Not m_queue Is Nothing Then
        If m_queue.Count = 0 Then Exit Sub
        action = m_queue(1)
        m_queue.Remove 1
        f = Split(action, US())
        LogLine "deferred: " & Replace(action, US(), " | ")
        Select Case f(0)
        Case "OPEN": OpenProject f(1), (f(2) = "replace")
        Case "SAVE": SaveProject f(1)
        Case "CMD": ExecuteCommand CLng(f(1))
        Case "FULLCOMPILE": ExecuteMenu "Run", "Start With Full Compile"
        Case "CALLSTACK": ExecuteMenu "View", "Call Stack"
        Case "IMM": TypeImmediate CLng("&H" & f(1)), f(2)
        Case "APPTEXT": SetWindowText CLng("&H" & f(1)), f(2)
        Case "SNAP": CaptureWindow CLng("&H" & f(1)), f(2)
        Case "KEYSTATE": SetModifierState CLng(f(1))
        Case "MAKEIDE": MakeInIde f(1)
        Case "RELOAD": ReloadComponent f(1), f(2), (f(3) = "discard")
        Case "KEYSCONT": PostKeys CLng("&H" & f(1)), f(3), CLng(f(2))
        Case "SELECT": SelectItem CLng("&H" & f(1)), CLng(f(2))
        Case "MOUSEREAL": DeferredRealMouse CLng("&H" & f(1)), f(2), CLng(f(3)), CLng(f(4)), CLng(f(5)), CLng(f(6))
        Case "KEYSREAL": DeferredRealKeys CLng("&H" & f(1)), f(2)
        Case "DESIGN": RunDesign f
        End Select
    End If
    Exit Sub
Fail:
    LogLine "deferred error " & Err.Number & ": " & Err.Description
    Resume Next
End Sub

' Makes the project active, then runs File > Save Project (command bar Id 3, "Sa&ve Project"),
' which saves the dirty components and the .vbp exactly like Ctrl+S.
Private Sub SaveProject(ByVal prjName As String)
    Dim p As Long, c As Object
    On Error GoTo Fail
    For p = 1 To g_VBE.VBProjects.Count
        If StrComp(g_VBE.VBProjects(p).Name, prjName, vbTextCompare) = 0 Then Set g_VBE.ActiveVBProject = g_VBE.VBProjects(p)
    Next
    Set c = g_VBE.CommandBars.FindControl(, 3)
    If c Is Nothing Then LogLine "save: command 3 not found": Exit Sub
    c.Execute
    LogLine "saved " & prjName
    Exit Sub
Fail:
    LogLine "save failed " & prjName & ": " & Err.Number & ": " & Err.Description
End Sub

' Debug commands from the allow-list in Verbs.QueueCommand. Execute returns right away; the run
' or break follows (spike 6/8 log: "Execute returned" before EnterRunMode).
Private Sub ExecuteCommand(ByVal id As Long)
    Dim c As Object
    On Error GoTo Fail
    Set c = g_VBE.CommandBars.FindControl(, id)
    If c Is Nothing Then LogLine "command " & id & " not found": Exit Sub
    If Not c.Enabled Then LogLine "command " & id & " (" & c.Caption & ") is disabled now": Exit Sub
    c.Execute
    LogLine "executed " & id & " (" & c.Caption & ")"
    Exit Sub
Fail:
    LogLine "command " & id & " failed: " & Err.Number & ": " & Err.Description
End Sub

' Runs a menu item found by menu path, for commands whose id is the generic 746 ("Run > Start With
' Full Compile", spike 6/8) or unknown ("View > Call Stack...", Phase 2). Captions are compared
' without "&" and a trailing "...". Deferred only: these can run user code or show a modal dialog.
Private Sub ExecuteMenu(ByVal menu As String, ByVal item As String)
    Dim bar As Object, i As Long, j As Long, pop As Object, cap As String
    On Error GoTo Fail
    Set bar = g_VBE.CommandBars("Menu Bar")
    For i = 1 To bar.Controls.Count
        Set pop = bar.Controls(i)
        If Replace(pop.Caption, "&", "") = menu Then
            For j = 1 To pop.Controls.Count
                cap = Replace(pop.Controls(j).Caption, "&", "")
                If Right$(cap, 3) = "..." Then cap = Left$(cap, Len(cap) - 3)
                If cap = item Then
                    If Not pop.Controls(j).Enabled Then LogLine menu & " > " & item & " is disabled now": Exit Sub
                    pop.Controls(j).Execute
                    LogLine "executed " & menu & " > " & item
                    Exit Sub
                End If
            Next
        End If
    Next
    LogLine menu & " > " & item & " not found"
    Exit Sub
Fail:
    LogLine menu & " > " & item & " failed: " & Err.Number & ": " & Err.Description
End Sub

' Focuses the Immediate window, types the line and presses Enter (spike 5: WM_CHAR + VK_RETURN).
' Win32 SetFocus, not VBIDE Window.SetFocus (that failed in break mode, Phase 2 live test).
Private Sub TypeImmediate(ByVal hWnd As Long, ByVal text As String)
    On Error GoTo Fail
    FocusWindow hWnd
    CaretToEnd hWnd            ' never type into the middle of existing text or over a selection
    SendText hWnd, text
    SendEnter hWnd
    LogLine "immediate: " & text
    Exit Sub
Fail:
    LogLine "immediate failed: " & Err.Number & ": " & Err.Description
End Sub

' Opens a project. "replace" closes the open ones first (the server refused if any has unsaved changes).
' Two VBIDE traps decide how (Phase 3/4 live tests):
' - AddFromFile(path, True) with projects still open wrote a Group1.vbg of the closed projects into
'   VB6's current folder. So the open projects are removed first.
' - AddFromFile(path, False) makes the IDE a project group, even with only one project, and then
'   Save asks for a group file name ("Save Project Group As"). So a project opened into an empty IDE
'   uses True (nothing left to close); only "add" next to other projects uses False.
' May show modal dialogs (missing references, source control, ...); the server watches for them.
Private Sub OpenProject(ByVal path As String, ByVal exclusive As Boolean)
    Dim i As Long
    On Error GoTo Fail
    If exclusive Then
        For i = g_VBE.VBProjects.Count To 1 Step -1       ' indexed, never For Each (spike 3)
            g_VBE.VBProjects.Remove g_VBE.VBProjects(i)
        Next
    End If
    g_VBE.VBProjects.AddFromFile path, (g_VBE.VBProjects.Count = 0)
    LogLine "opened " & path & IIf(exclusive, " (replace)", " (add)")
    Exit Sub
Fail:
    LogLine "open failed " & path & ": " & Err.Number & ": " & Err.Description
End Sub

' VBProject.MakeCompiledFile from a fresh dispatch (build_in_ide). Errors show as dialogs, which the
' server reads; failures of the call itself are logged.
Private Sub MakeInIde(ByVal prjName As String)
    Dim p As Long
    On Error GoTo Fail
    m_makeResult = "running"
    For p = 1 To g_VBE.VBProjects.Count
        If StrComp(g_VBE.VBProjects(p).Name, prjName, vbTextCompare) = 0 Then
            LogLine "MakeCompiledFile " & prjName & " -> " & g_VBE.VBProjects(p).BuildFileName
            g_VBE.VBProjects(p).MakeCompiledFile
            LogLine "MakeCompiledFile returned"
            m_makeResult = "ok"
            Exit Sub
        End If
    Next
    LogLine "MakeCompiledFile: project not found: " & prjName
    m_makeResult = "failed" & US() & "project not found"
    Exit Sub
Fail:
    LogLine "MakeCompiledFile failed: " & Err.Number & ": " & Err.Description
    m_makeResult = "failed" & US() & Err.Description
End Sub

Public Sub MarkMakeQueued()
    m_makeResult = "queued"
End Sub

Public Function LastMakeResult() As String
    LastMakeResult = m_makeResult
End Function

' VBComponent.Reload from a fresh dispatch. Checks again: between queueing and now the user may
' have started typing (then a "clean" reload is skipped) or started the program.
Private Sub ReloadComponent(ByVal prjName As String, ByVal cmpName As String, ByVal discard As Boolean)
    Dim p As Long, c As Long, cmp As VBIDE.VBComponent
    On Error GoTo Fail
    If g_VBE.CommandBars.FindControl(, 228).Enabled Then LogLine "reload " & cmpName & " skipped: not in design mode": Exit Sub
    For p = 1 To g_VBE.VBProjects.Count
        If StrComp(g_VBE.VBProjects(p).Name, prjName, vbTextCompare) = 0 Then
            For c = 1 To g_VBE.VBProjects(p).VBComponents.Count
                Set cmp = g_VBE.VBProjects(p).VBComponents(c)
                If StrComp(cmp.Name, cmpName, vbTextCompare) = 0 Then
                    If cmp.IsDirty And Not discard Then LogLine "reload " & cmpName & " skipped: it has unsaved changes now": Exit Sub
                    cmp.Reload
                    LogLine "reloaded " & prjName & "." & cmpName & IIf(discard, " (unsaved edits discarded on request)", "")
                    Exit Sub
                End If
            Next
        End If
    Next
    LogLine "reload: " & prjName & "." & cmpName & " not found"
    Exit Sub
Fail:
    LogLine "reload " & cmpName & " failed: " & Err.Number & ": " & Err.Description
End Sub
