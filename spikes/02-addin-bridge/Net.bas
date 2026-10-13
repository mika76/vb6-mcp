Attribute VB_Name = "Net"
' Spike 2: loopback TCP listener inside the VB6 IDE.
' Raw Winsock + WSAAsyncSelect on a hidden message-only window, so socket events
' arrive as window messages on the IDE thread (no polling, no blocking calls).
'
' Wire format (spike only):
'   request : "<token> <VERB> [args]" + LF          (args: tab-separated)
'   response: "OK <bytes>" + CRLF + payload          or  "ERR <bytes>" + CRLF + message
Option Explicit

Public g_VBE As VBIDE.VBE

' ---- Winsock ----
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare Function api_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal stype As Long, ByVal protocol As Long) As Long
Private Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, addr As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByVal addr As Long, ByVal addrlen As Long) As Long
Private Declare Function api_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function api_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function api_getsockname Lib "ws2_32.dll" Alias "getsockname" (ByVal s As Long, name As SOCKADDR_IN, namelen As Long) As Long
Private Declare Function api_htons Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Integer) As Integer
Private Declare Function api_ntohs Lib "ws2_32.dll" Alias "ntohs" (ByVal netshort As Integer) As Integer
Private Declare Function api_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long

Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const INVALID_SOCKET As Long = -1
Private Const FD_READ As Long = 1
Private Const FD_ACCEPT As Long = 8
Private Const FD_CLOSE As Long = 32

' ---- Window / misc ----
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CoCreateGuid Lib "ole32" (pguid As Any) As Long

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

Private Const HWND_MESSAGE As Long = -3
Private Const GWL_WNDPROC As Long = -4
Private Const WM_SOCKET As Long = &H400 + 100   ' WM_USER + 100

' ---- State ----
Private m_hWnd As Long
Private m_prevProc As Long
Private m_listen As Long
Private m_port As Long
Private m_token As String
Private m_discoveryFile As String
Private m_buffers As Collection   ' key "s<socket>" -> pending request text
Private m_wsaStarted As Boolean

Public Function ServerPort() As Long
    ServerPort = m_port
End Function

Public Function ServerPid() As Long
    ServerPid = GetCurrentProcessId()
End Function

Public Function DeferHwnd() As Long
    DeferHwnd = m_hWnd
End Function

Public Function DiscoveryFile() As String
    DiscoveryFile = m_discoveryFile
End Function

' ================= lifecycle =================

Public Sub StartServer()
    Dim wsa As WSADATA, addr As SOCKADDR_IN, nameLen As Long

    If WSAStartup(&H202, wsa) <> 0 Then LogLine "WSAStartup failed": Exit Sub
    m_wsaStarted = True
    Set m_buffers = New Collection

    m_hWnd = CreateWindowEx(0, "STATIC", "VB6McpSpike", 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, 0)
    If m_hWnd = 0 Then LogLine "CreateWindowEx failed": Exit Sub
    m_prevProc = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf WndProc)

    m_listen = api_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If m_listen = INVALID_SOCKET Then LogLine "socket failed " & WSAGetLastError(): Exit Sub

    addr.sin_family = AF_INET
    addr.sin_port = 0                                   ' let Windows pick a free port
    addr.sin_addr = api_inet_addr("127.0.0.1")          ' loopback only, never 0.0.0.0
    If api_bind(m_listen, addr, LenB(addr)) <> 0 Then LogLine "bind failed " & WSAGetLastError(): Exit Sub
    If api_listen(m_listen, 5) <> 0 Then LogLine "listen failed " & WSAGetLastError(): Exit Sub

    nameLen = LenB(addr)
    api_getsockname m_listen, addr, nameLen
    m_port = CLng(api_ntohs(addr.sin_port)) And &HFFFF&

    If WSAAsyncSelect(m_listen, m_hWnd, WM_SOCKET, FD_ACCEPT) <> 0 Then LogLine "WSAAsyncSelect failed " & WSAGetLastError(): Exit Sub

    m_token = NewToken()
    WriteDiscovery
    LogLine "listening on 127.0.0.1:" & m_port & " pid=" & GetCurrentProcessId()
End Sub

' Spike 12: start the watchdog as a child process (inherits elevation, no UAC prompt).
' A second instance exits by itself (named mutex), so starting it on every load is fine.
' CreateProcess, not Shell: Shell blocked the IDE thread ~9 s (spike 12).
Public Sub StartWatchdog()
    Dim exe As String, si As STARTUPINFO, pi As PROCESS_INFORMATION, t As Single
    On Error GoTo Fail
    exe = App.Path & "\..\..\12-watchdog\bin\Debug\net10.0-windows\Watchdog.exe"
    If Len(Dir$(exe)) = 0 Then LogLine "watchdog not found: " & exe: Exit Sub
    si.cb = LenB(si)
    t = Timer
    If CreateProcess(vbNullString, """" & exe & """", 0, 0, 0, CREATE_NO_WINDOW, 0, vbNullString, si, pi) = 0 Then
        LogLine "CreateProcess failed": Exit Sub
    End If
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    LogLine "watchdog started pid=" & pi.dwProcessId & " in " & Format$((Timer - t) * 1000, "0") & " ms"
    Exit Sub
Fail:
    LogLine "StartWatchdog error " & Err.Number & ": " & Err.Description
End Sub

Public Sub StopServer()
    On Error Resume Next
    If m_listen <> 0 And m_listen <> INVALID_SOCKET Then api_closesocket m_listen
    m_listen = 0
    If m_hWnd <> 0 Then
        If m_prevProc <> 0 Then SetWindowLong m_hWnd, GWL_WNDPROC, m_prevProc
        DestroyWindow m_hWnd
    End If
    m_hWnd = 0: m_prevProc = 0
    If Len(m_discoveryFile) > 0 Then Kill m_discoveryFile
    m_discoveryFile = ""
    If m_wsaStarted Then WSACleanup
    m_wsaStarted = False
    Set m_buffers = Nothing
    LogLine "stopped"
End Sub

' ================= window procedure (IDE thread) =================

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo Fail
    If uMsg = WM_SOCKET Then
        OnSocketEvent wParam, lParam And &HFFFF&
        Exit Function
    End If
    If uMsg = WM_DEFER Then
        RunDeferred
        Exit Function
    End If
    WndProc = CallWindowProc(m_prevProc, hWnd, uMsg, wParam, lParam)
    Exit Function
Fail:
    LogLine "WndProc error " & Err.Number & ": " & Err.Description
End Function

Private Sub OnSocketEvent(ByVal s As Long, ByVal evt As Long)
    Dim c As Long
    Select Case evt
    Case FD_ACCEPT
        c = api_accept(m_listen, 0, 0)
        If c <> INVALID_SOCKET Then
            WSAAsyncSelect c, m_hWnd, WM_SOCKET, FD_READ Or FD_CLOSE
            m_buffers.Add "", "s" & c
        End If
    Case FD_READ
        ReadFrom s
    Case FD_CLOSE
        CloseClient s
    End Select
End Sub

Private Sub ReadFrom(ByVal s As Long)
    Dim buf(0 To 8191) As Byte, n As Long, pending As String, chunk As String, p As Long
    n = api_recv(s, buf(0), 8192, 0)
    If n <= 0 Then Exit Sub
    chunk = StrConv(LeftB$(buf, n), vbUnicode)

    pending = m_buffers("s" & s) & chunk
    Do
        p = InStr(pending, vbLf)
        If p = 0 Then Exit Do
        HandleRequest s, Replace(Left$(pending, p - 1), vbCr, "")
        pending = Mid$(pending, p + 1)
    Loop
    m_buffers.Remove "s" & s
    m_buffers.Add pending, "s" & s
End Sub

Private Sub CloseClient(ByVal s As Long)
    On Error Resume Next
    api_closesocket s
    m_buffers.Remove "s" & s
End Sub

' ================= requests =================

Private Sub HandleRequest(ByVal s As Long, ByVal line As String)
    Dim sp As Long, tok As String, rest As String, verb As String, args As String, r As Variant
    On Error GoTo Fail

    sp = InStr(line, " ")
    If sp = 0 Then SendReply s, False, "bad request": Exit Sub
    tok = Left$(line, sp - 1): rest = Mid$(line, sp + 1)
    If tok <> m_token Then SendReply s, False, "bad token": Exit Sub
    StatusCountRequest

    sp = InStr(rest, " ")
    If sp = 0 Then verb = rest Else verb = Left$(rest, sp - 1): args = Mid$(rest, sp + 1)

    Select Case UCase$(verb)
    Case "PING": SendReply s, True, "pong pid=" & GetCurrentProcessId()
    Case "LIST": SendReply s, True, ListComponents()
    Case "READ": SendReply s, True, ReadComponent(args)
    Case Else
        r = Probe.Dispatch(UCase$(verb), args)
        If IsEmpty(r) Then SendReply s, False, "unknown verb " & verb Else SendReply s, True, CStr(r)
    End Select
    Exit Sub
Fail:
    SendReply s, False, "error " & Err.Number & ": " & Err.Description
End Sub

' project <TAB> component <TAB> type <TAB> dirty <TAB> lines, one per line.
' Indexed loops only: For Each over VBProjects raises error 70 in run/break mode (spike 3).
Private Function ListComponents() As String
    Dim p As Long, c As Long, prj As VBIDE.VBProject, cmp As VBIDE.VBComponent, out As String, nLines As Long
    For p = 1 To g_VBE.VBProjects.Count
        Set prj = g_VBE.VBProjects(p)
        For c = 1 To prj.VBComponents.Count
            Set cmp = prj.VBComponents(c)
            nLines = -1
            On Error Resume Next
            nLines = cmp.CodeModule.CountOfLines
            On Error GoTo 0
            out = out & prj.Name & vbTab & cmp.Name & vbTab & cmp.Type & vbTab & cmp.IsDirty & vbTab & nLines & vbCrLf
        Next
    Next
    ListComponents = out
End Function

' args: project <TAB> component
Private Function ReadComponent(ByVal args As String) As String
    Dim parts() As String, p As Long, c As Long, prj As VBIDE.VBProject, cmp As VBIDE.VBComponent, cm As VBIDE.CodeModule
    parts = Split(args, vbTab)
    If UBound(parts) < 1 Then Err.Raise vbObjectError + 1, , "READ needs project<TAB>component"
    For p = 1 To g_VBE.VBProjects.Count
        Set prj = g_VBE.VBProjects(p)
        If StrComp(prj.Name, parts(0), vbTextCompare) = 0 Then
            For c = 1 To prj.VBComponents.Count
                Set cmp = prj.VBComponents(c)
                If StrComp(cmp.Name, parts(1), vbTextCompare) = 0 Then
                    Set cm = cmp.CodeModule
                    If cm.CountOfLines = 0 Then ReadComponent = "" Else ReadComponent = cm.Lines(1, cm.CountOfLines)
                    Exit Function
                End If
            Next
        End If
    Next
    Err.Raise vbObjectError + 2, , "component not found: " & parts(0) & "." & parts(1)
End Function

Private Sub SendReply(ByVal s As Long, ByVal ok As Boolean, ByVal payload As String)
    Dim body() As Byte, head() As Byte, all() As Byte, i As Long, sent As Long, n As Long
    ' 1252 is single-byte, so the byte count equals Len(payload)
    body = StrConv(payload, vbFromUnicode)
    head = StrConv(IIf(ok, "OK ", "ERR ") & Len(payload) & vbCrLf, vbFromUnicode)
    ' Single buffer so header and body go out together
    n = UBound(head) + 1 + Len(payload)
    ReDim all(0 To n - 1)
    For i = 0 To UBound(head): all(i) = head(i): Next
    For i = 0 To Len(payload) - 1: all(UBound(head) + 1 + i) = body(i): Next
    Do While sent < n
        i = api_send(s, all(sent), n - sent, 0)
        If i <= 0 Then LogLine "send failed " & WSAGetLastError() & " after " & sent & "/" & n: Exit Sub
        sent = sent + i
    Loop
End Sub

' ================= helpers =================

Private Function NewToken() As String
    Dim g(0 To 15) As Byte, i As Long, s As String
    CoCreateGuid g(0)
    For i = 0 To 15: s = s & Right$("0" & Hex$(g(i)), 2): Next
    NewToken = LCase$(s)
End Function

Private Function BaseDir() As String
    BaseDir = Environ$("LOCALAPPDATA") & "\vb6mcp"
End Function

Private Sub EnsureDir(ByVal p As String)
    On Error Resume Next
    MkDir p
End Sub

Private Sub WriteDiscovery()
    Dim f As Integer, folder As String
    EnsureDir BaseDir()
    folder = BaseDir() & "\instances"
    EnsureDir folder
    m_discoveryFile = folder & "\" & GetCurrentProcessId() & ".json"
    f = FreeFile
    Open m_discoveryFile For Output As #f
    Print #f, "{""pid"":" & GetCurrentProcessId() & ",""port"":" & m_port & ",""token"":""" & m_token & """,""version"":""spike2""}"
    Close #f
End Sub

Public Sub LogLine(ByVal msg As String)
    Dim f As Integer
    On Error Resume Next
    EnsureDir BaseDir()
    f = FreeFile
    Open BaseDir() & "\addin-spike.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & GetCurrentProcessId() & "] " & msg
    Close #f
End Sub
