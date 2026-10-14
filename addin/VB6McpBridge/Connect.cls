VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' VB6McpBridge add-in entry point (ProgID VB6McpBridge.Connect).
' Starts the loopback listener and the status toolbar; see docs/architecture.md.
' Rules (spikes): every public entry has an error handler (an add-in error can take the IDE
' down); no VBIDE access from build events; indexed loops only.
Option Explicit

Implements IDTExtensibility

Private WithEvents m_statusClick As VBIDE.CommandBarEvents
Attribute m_statusClick.VB_VarHelpID = -1
' The IDE's own file writes (Sync.bas), for all projects (Nothing)
Private WithEvents m_files As VBIDE.FileControlEvents
Attribute m_files.VB_VarHelpID = -1

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    On Error GoTo Fail
    Set g_VBE = VBInst
    LogLine "OnConnection mode=" & ConnectMode & " version=" & AddinVersion()
    ' VB6.EXE /make (build_make) loads startup add-ins too, as vbext_cm_CommandLine (3). A headless
    ' build must not open a listener, write a discovery file or start the watchdog (Phase 4).
    If ConnectMode = 3 Then LogLine "command-line build: listener not started": Exit Sub
    StartListener
    StartWatchdog
    StatusCreate
    If Not g_StatusBtn Is Nothing Then Set m_statusClick = g_VBE.Events.CommandBarEvents(g_StatusBtn)
    Set m_files = g_VBE.Events.FileControlEvents(Nothing)
    Exit Sub
Fail:
    LogLine "OnConnection error " & Err.Number & ": " & Err.Description
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
    On Error GoTo Fail
    LogLine "OnDisconnection mode=" & RemoveMode
    Set m_statusClick = Nothing
    Set m_files = Nothing
    StatusDestroy
    StopListener
    Set g_VBE = Nothing
    Exit Sub
Fail:
    LogLine "OnDisconnection error " & Err.Number & ": " & Err.Description
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
End Sub

' Status button: refresh and log details. No modal UI (AGENTS.md hard rule 5).
Private Sub m_statusClick_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    StatusRefresh
    LogLine "status: " & StatusDetails()
    handled = True
End Sub

' Signature from VB6EXT.OLB: the FileControlEvents coclass sources _dispVBFileControlEvents, whose
' AfterWriteFile has 4 parameters (an unrelated _FileControlEvents interface has 2; the first dump
' read that one and the build failed, Phase 4). Only records; no VBIDE access in event handlers.
Private Sub m_files_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal Result As Integer)
    On Error Resume Next
    RecordWrite FileName
End Sub