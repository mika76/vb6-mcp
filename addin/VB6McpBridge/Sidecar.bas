Attribute VB_Name = "Sidecar"
' Starts the elevated watchdog sidecar (PLAN.md section 3.1, spike 12): vb6mcp.exe in "watchdog" mode,
' from the "watchdog" folder next to this DLL (build-addin.ps1 puts it there). As a child of the
' elevated VB6 it inherits elevation with no UAC prompt, and it keeps running if VB6 crashes.
' A second copy exits at once (named mutex), so starting it from every IDE is fine.
' CreateProcess without waiting, never VB Shell: Shell blocked the IDE thread ~9 s (spike 12).
' Security note: the watchdog runs elevated, so it must live where this DLL lives; a user-writable
' DLL folder is the same exposure either way (docs/architecture.md "Security").
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const CREATE_NO_WINDOW As Long = &H8000000
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public Sub StartWatchdog()
    Dim exe As String, si As STARTUPINFO, pi As PROCESS_INFORMATION
    On Error GoTo Fail
    exe = App.Path & "\watchdog\vb6mcp.exe"
    If Len(Dir$(exe)) = 0 Then LogLine "watchdog not installed (" & exe & "); hang handling unavailable": Exit Sub
    si.cb = LenB(si)
    If CreateProcess(vbNullString, """" & exe & """ watchdog", 0, 0, 0, CREATE_NO_WINDOW, 0, vbNullString, si, pi) = 0 Then
        LogLine "watchdog: CreateProcess failed (" & Err.LastDllError & ")": Exit Sub
    End If
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    LogLine "watchdog started, pid " & pi.dwProcessId
    Exit Sub
Fail:
    LogLine "StartWatchdog error " & Err.Number & ": " & Err.Description
End Sub

' Starts "VB6.EXE /make <vbp> /out <log>" as a child process (build_make, Phase 4) and returns its
' pid, without waiting (hard rule 5). The child inherits VB6's elevation, so building an ActiveX
' DLL/OCX can register it without a UAC prompt. /make reads the files on disk, not IDE memory.
Public Function StartBuild(ByVal vbp As String, ByVal logPath As String) As Long
    Dim exe As String, n As Long, si As STARTUPINFO, pi As PROCESS_INFORMATION
    exe = String$(520, vbNullChar)
    n = GetModuleFileName(0, exe, 520)                 ' this process: VB6.EXE
    exe = Left$(exe, n)
    si.cb = LenB(si)
    If CreateProcess(vbNullString, """" & exe & """ /make """ & vbp & """ /out """ & logPath & """", 0, 0, 0, CREATE_NO_WINDOW, 0, vbNullString, si, pi) = 0 Then
        StartBuild = 0
        LogLine "build: CreateProcess failed (" & Err.LastDllError & ")"
        Exit Function
    End If
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    LogLine "build started, pid " & pi.dwProcessId & ": " & vbp
    StartBuild = pi.dwProcessId
End Function
