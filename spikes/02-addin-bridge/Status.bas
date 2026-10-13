Attribute VB_Name = "Status"
' Visible add-in status: a temporary toolbar with one caption button showing
' version, DLL build time, port and request count. Makes it obvious which build is loaded.
Option Explicit

Public g_StatusBtn As Object

Private m_bar As Object
Private m_requests As Long
Private Const BAR_NAME As String = "VB6 MCP"

Public Sub StatusCreate()
    On Error Resume Next
    g_VBE.CommandBars(BAR_NAME).Delete          ' leftover from a crashed session
    On Error GoTo Fail
    Set m_bar = g_VBE.CommandBars.Add(BAR_NAME, 1, False, True)   ' msoBarTop, temporary
    Set g_StatusBtn = m_bar.Controls.Add(1, , , , True)           ' msoControlButton, temporary
    g_StatusBtn.Style = 2                                         ' msoButtonCaption
    m_bar.Visible = True
    StatusRefresh
    Exit Sub
Fail:
    LogLine "StatusCreate error " & Err.Number & ": " & Err.Description
End Sub

Public Sub StatusDestroy()
    On Error Resume Next
    If Not m_bar Is Nothing Then m_bar.Delete
    Set g_StatusBtn = Nothing
    Set m_bar = Nothing
End Sub

Public Sub StatusCountRequest()
    m_requests = m_requests + 1
    StatusRefresh
End Sub

Public Sub StatusRefresh()
    On Error Resume Next
    If g_StatusBtn Is Nothing Then Exit Sub
    g_StatusBtn.Caption = StatusCaption()
    g_StatusBtn.TooltipText = StatusDetails()
End Sub

Public Function StatusCaption() As String
    Dim dot As String
    dot = " " & Chr$(183) & " "
    StatusCaption = "VB6 MCP " & App.Major & "." & App.Minor & dot & "built " & BuildTime() & dot & ":" & ServerPort() & dot & m_requests & " req"
End Function

Public Function StatusDetails() As String
    StatusDetails = "pid " & ServerPid() & ", port " & ServerPort() & ", requests " & m_requests & ", dll " & DllPath() & ", discovery " & DiscoveryFile()
End Function

Private Function DllPath() As String
    DllPath = App.Path & "\" & App.EXEName & ".dll"
End Function

Private Function BuildTime() As String
    On Error GoTo Fail
    BuildTime = Format$(FileDateTime(DllPath()), "dd-mm hh:nn")
    Exit Function
Fail:
    BuildTime = "?"
End Function
