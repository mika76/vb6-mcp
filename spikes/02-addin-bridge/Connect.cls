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
' Spike 2: add-in entry point. Starts/stops the loopback TCP listener in Net.bas.
Option Explicit

Implements IDTExtensibility

Private WithEvents m_build As VBIDE.VBBuildEvents
Attribute m_build.VB_VarHelpID = -1
Private WithEvents m_statusClick As VBIDE.CommandBarEvents
Attribute m_statusClick.VB_VarHelpID = -1

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    Dim e2 As VBIDE.Events2
    On Error GoTo Fail
    Set g_VBE = VBInst
    LogLine "OnConnection mode=" & ConnectMode
    StartServer
    Set e2 = g_VBE.Events
    Set m_build = e2.VBBuildEvents
    LogLine "build events hooked"
    StatusCreate
    If Not g_StatusBtn Is Nothing Then Set m_statusClick = g_VBE.Events.CommandBarEvents(g_StatusBtn)
    LogLine "status: " & StatusCaption()
    StartWatchdog
    Exit Sub
Fail:
    LogLine "OnConnection error " & Err.Number & ": " & Err.Description
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
    On Error GoTo Fail
    LogLine "OnDisconnection mode=" & RemoveMode
    Set m_build = Nothing
    Set m_statusClick = Nothing
    StatusDestroy
    StopServer
    Set g_VBE = Nothing
    Exit Sub
Fail:
    LogLine "OnDisconnection error " & Err.Number & ": " & Err.Description
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
End Sub

' Status button: refresh the caption and log the details (no modal UI)
Private Sub m_statusClick_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    StatusRefresh
    LogLine "status clicked: " & StatusDetails()
    handled = True
End Sub

' Build events: LOG ONLY. Reading code inside EnterRunMode ended the debugged
' program as soon as Form_Load returned (spike 3); the handlers must not touch VBIDE.
Private Sub m_build_BeginCompile(ByVal VBProject As VBIDE.VBProject)
    On Error Resume Next
    LogLine "BeginCompile " & VBProject.Name
End Sub

Private Sub m_build_EnterRunMode()
    On Error Resume Next
    LogLine "EnterRunMode"
End Sub

Private Sub m_build_EnterDesignMode()
    On Error Resume Next
    LogLine "EnterDesignMode"
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
End Sub
