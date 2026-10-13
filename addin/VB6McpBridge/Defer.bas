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

Public Sub DeferInit()
    Set m_queue = New Collection
End Sub

' Queues an action and posts the window message that will run it.
Public Sub Enqueue(ByVal action As String)
    If m_queue Is Nothing Then DeferInit
    m_queue.Add action
    PostMessage ListenerHwnd(), WM_DEFER, 0, 0
End Sub

' Called from the listener's window procedure on WM_DEFER: runs every queued action.
Public Sub RunQueued()
    Dim action As String, f() As String
    On Error GoTo Fail
    Do While Not m_queue Is Nothing
        If m_queue.Count = 0 Then Exit Do
        action = m_queue(1)
        m_queue.Remove 1
        f = Split(action, US())
        LogLine "deferred: " & Replace(action, US(), " | ")
        Select Case f(0)
        Case "OPEN": OpenProject f(1), (f(2) = "replace")
        Case "SAVE": SaveProject f(1)
        End Select
    Loop
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

' VBProjects.AddFromFile(path, exclusive). Exclusive closes the open projects first.
' May show modal dialogs (missing references, source control, ...); the server watches for them.
Private Sub OpenProject(ByVal path As String, ByVal exclusive As Boolean)
    On Error GoTo Fail
    g_VBE.VBProjects.AddFromFile path, exclusive
    LogLine "opened " & path & IIf(exclusive, " (replace)", " (add)")
    Exit Sub
Fail:
    LogLine "open failed " & path & ": " & Err.Number & ": " & Err.Description
End Sub
