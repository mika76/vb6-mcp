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
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const STILL_ACTIVE As Long = 259

Private m_regHandles As Collection          ' "p" & pid -> process handle, until REGSTATUS saw it exit

Public Sub StartWatchdog()
    Dim exe As String, si As STARTUPINFO, pi As PROCESS_INFORMATION
    On Error GoTo Fail
    ' Installed (Phase 6): vb6mcp.exe sits next to this DLL in Program Files. Development: the watchdog folder that
    ' build-addin.ps1 publishes into.
    exe = App.Path & "\vb6mcp.exe"
    If Len(Dir$(exe)) = 0 Then exe = App.Path & "\watchdog\vb6mcp.exe"
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
Public Function StartBuild(ByVal vbp As String, ByVal logPath As String, Optional ByVal defines As String = "") As Long
    Dim exe As String, n As Long, si As STARTUPINFO, pi As PROCESS_INFORMATION, extra As String
    exe = String$(520, vbNullChar)
    n = GetModuleFileName(0, exe, 520)                 ' this process: VB6.EXE
    exe = Left$(exe, n)
    si.cb = LenB(si)
    ' /d name=value[:name=value]: conditional compilation constants, validated by Verbs.ValidDefines
    If Len(defines) > 0 Then extra = " /d " & defines
    If CreateProcess(vbNullString, """" & exe & """ /make """ & vbp & """ /out """ & logPath & """" & extra, 0, 0, 0, CREATE_NO_WINDOW, 0, vbNullString, si, pi) = 0 Then
        StartBuild = 0
        LogLine "build: CreateProcess failed (" & Err.LastDllError & ")"
        Exit Function
    End If
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    LogLine "build started, pid " & pi.dwProcessId & ": " & vbp
    StartBuild = pi.dwProcessId
End Function

' Starts regsvr32 /s [/u] "<path>" as a child (inherits VB6's elevation, so no UAC prompt) and keeps its
' process handle, so REGSTATUS can read the exit code without waiting (hard rule 5). GetSystemDirectory in
' this 32-bit process is redirected to SysWOW64: the 32-bit regsvr32, which registers in the 32-bit view.
' The caller (Verbs.RegisterFile) has validated the path; returns 0 if the process couldn't start.
Public Function StartRegsvr(ByVal path As String, ByVal unregister As Boolean) As Long
    Dim sysDir As String, n As Long, si As STARTUPINFO, pi As PROCESS_INFORMATION, cmd As String
    sysDir = String$(260, vbNullChar)
    n = GetSystemDirectory(sysDir, 260)
    sysDir = Left$(sysDir, n)
    cmd = """" & sysDir & "\regsvr32.exe"" /s " & IIf(unregister, "/u ", "") & """" & path & """"
    si.cb = LenB(si)
    If CreateProcess(vbNullString, cmd, 0, 0, 0, CREATE_NO_WINDOW, 0, vbNullString, si, pi) = 0 Then
        LogLine "regsvr32: CreateProcess failed (" & Err.LastDllError & ")"
        Exit Function
    End If
    CloseHandle pi.hThread
    If m_regHandles Is Nothing Then Set m_regHandles = New Collection
    m_regHandles.Add pi.hProcess, "p" & pi.dwProcessId
    LogLine "regsvr32 started, pid " & pi.dwProcessId & ": " & cmd
    StartRegsvr = pi.dwProcessId
End Function

' "running", or "exit" US code (then the handle is closed). "unknown" for a pid we didn't start.
Public Function RegsvrStatus(ByVal pid As Long) As String
    Dim h As Long, code As Long
    If m_regHandles Is Nothing Then RegsvrStatus = "unknown": Exit Function
    On Error Resume Next
    h = m_regHandles("p" & pid)
    If Err.Number <> 0 Then RegsvrStatus = "unknown": Exit Function
    On Error GoTo 0
    If GetExitCodeProcess(h, code) = 0 Then RegsvrStatus = "unknown": Exit Function
    If code = STILL_ACTIVE Then RegsvrStatus = "running": Exit Function
    CloseHandle h
    m_regHandles.Remove "p" & pid
    LogLine "regsvr32 pid " & pid & " exited with " & code
    RegsvrStatus = "exit" & US() & code
End Function
