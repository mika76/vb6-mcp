Attribute VB_Name = "Log"
' Shared state and logging. Log file: %LOCALAPPDATA%\vb6mcp\addin.log
Option Explicit

Public g_VBE As VBIDE.VBE

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Function AddinVersion() As String
    AddinVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function ProcessId() As Long
    ProcessId = GetCurrentProcessId()
End Function

Public Function BaseDir() As String
    BaseDir = Environ$("LOCALAPPDATA") & "\vb6mcp"
End Function

Public Sub EnsureDir(ByVal path As String)
    On Error Resume Next
    MkDir path
End Sub

Public Sub LogLine(ByVal msg As String)
    Dim f As Integer
    On Error Resume Next
    EnsureDir BaseDir()
    f = FreeFile
    Open BaseDir() & "\addin.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & ProcessId() & "] " & msg
    Close #f
End Sub
