Attribute VB_Name = "Listener"
' Loopback TCP listener speaking wire protocol v1 (docs/protocol.md).
' Raw Winsock + WSAAsyncSelect on a hidden message-only window: socket events arrive as window
' messages on the IDE thread, so VBIDE can be used safely and nothing ever blocks (spike 2).
' Frames: "<length>" LF <body>, body fields separated by US (0x1F), text in the ANSI code page.
' Because the code page is single-byte (1252), string length = byte length after StrConv.
Option Explicit

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
Private Declare Function api_ntohs Lib "ws2_32.dll" Alias "ntohs" (ByVal netshort As Integer) As Integer
Private Declare Function api_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long

Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const INVALID_SOCKET As Long = -1
Private Const SOCKET_ERROR As Long = -1
Private Const WSAEWOULDBLOCK As Long = 10035
Private Const FD_READ As Long = 1
Private Const FD_WRITE As Long = 2
Private Const FD_ACCEPT As Long = 8
Private Const FD_CLOSE As Long = 32

' ---- window ----
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CoCreateGuid Lib "ole32" (pguid As Any) As Long

Private Const HWND_MESSAGE As Long = -3
Private Const GWL_WNDPROC As Long = -4
Private Const WM_SOCKET As Long = &H400 + 100   ' WM_USER + 100

' ---- protocol ----
Public Const MAGIC As String = "VB6MCP1"
Public Const PROTOCOL_VERSION As Long = 1
Private Const MAX_BODY As Long = 16777216        ' 16 MiB

' ---- state ----
Private m_hWnd As Long
Private m_prevProc As Long
Private m_listen As Long
Private m_port As Long
Private m_token As String
Private m_discoveryFile As String
Private m_wsaStarted As Boolean
Private m_inBuf As Collection      ' "s<socket>" -> unparsed input (String)
Private m_outBuf As Collection     ' "s<socket>" -> unsent output (String, ANSI chars)
Private m_closeAfter As Collection ' "s<socket>" -> True when the socket must close once flushed
Private m_requests As Long

Public Function US() As String
    US = Chr$(31)
End Function

Public Function RS() As String
    RS = Chr$(30)
End Function

Public Function ListenerHwnd() As Long
    ListenerHwnd = m_hWnd
End Function

Public Function ListenerPort() As Long
    ListenerPort = m_port
End Function

Public Function RequestCount() As Long
    RequestCount = m_requests
End Function

Public Function DiscoveryFilePath() As String
    DiscoveryFilePath = m_discoveryFile
End Function

' ================= lifecycle =================

Public Sub StartListener()
    Dim wsa As WSADATA, addr As SOCKADDR_IN, nameLen As Long
    If WSAStartup(&H202, wsa) <> 0 Then LogLine "WSAStartup failed": Exit Sub
    m_wsaStarted = True
    Set m_inBuf = New Collection
    Set m_outBuf = New Collection
    Set m_closeAfter = New Collection

    m_hWnd = CreateWindowEx(0, "STATIC", "VB6McpBridge", 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, 0)
    If m_hWnd = 0 Then LogLine "CreateWindowEx failed": Exit Sub
    m_prevProc = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf WndProc)

    m_listen = api_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If m_listen = INVALID_SOCKET Then LogLine "socket failed " & WSAGetLastError(): Exit Sub
    addr.sin_family = AF_INET
    addr.sin_port = 0                                   ' Windows picks a free port
    addr.sin_addr = api_inet_addr("127.0.0.1")          ' loopback only, never 0.0.0.0 (hard rule 8)
    If api_bind(m_listen, addr, LenB(addr)) <> 0 Then LogLine "bind failed " & WSAGetLastError(): Exit Sub
    If api_listen(m_listen, 5) <> 0 Then LogLine "listen failed " & WSAGetLastError(): Exit Sub
    nameLen = LenB(addr)
    api_getsockname m_listen, addr, nameLen
    m_port = CLng(api_ntohs(addr.sin_port)) And &HFFFF&
    If WSAAsyncSelect(m_listen, m_hWnd, WM_SOCKET, FD_ACCEPT) <> 0 Then LogLine "WSAAsyncSelect failed " & WSAGetLastError(): Exit Sub

    m_token = NewToken()
    WriteDiscovery
    LogLine "listening on 127.0.0.1:" & m_port
End Sub

Public Sub StopListener()
    On Error Resume Next
    If m_listen <> 0 And m_listen <> INVALID_SOCKET Then api_closesocket m_listen
    m_listen = 0
    If m_hWnd <> 0 Then
        If m_prevProc <> 0 Then SetWindowLong m_hWnd, GWL_WNDPROC, m_prevProc
        DestroyWindow m_hWnd
    End If
    m_hWnd = 0: m_prevProc = 0
    ' A clean unload deletes the discovery file; a left-over file means the IDE crashed or was killed
    If Len(m_discoveryFile) > 0 Then Kill m_discoveryFile
    m_discoveryFile = ""
    If m_wsaStarted Then WSACleanup
    m_wsaStarted = False
    Set m_inBuf = Nothing: Set m_outBuf = Nothing: Set m_closeAfter = Nothing
    LogLine "listener stopped"
End Sub

' ================= window procedure (IDE thread) =================

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo Fail
    If uMsg = WM_SOCKET Then
        OnSocketEvent wParam, lParam And &HFFFF&
        Exit Function
    End If
    If uMsg = WM_DEFER Then
        RunQueued
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
            WSAAsyncSelect c, m_hWnd, WM_SOCKET, FD_READ Or FD_WRITE Or FD_CLOSE
            m_inBuf.Add "", Key(c)
            m_outBuf.Add "", Key(c)
        End If
    Case FD_READ
        ReadFrom s
    Case FD_WRITE
        Flush s
    Case FD_CLOSE
        CloseClient s
    End Select
End Sub

Private Function Key(ByVal s As Long) As String
    Key = "s" & s
End Function

Private Sub SetBuf(ByVal col As Collection, ByVal s As Long, ByVal value As String)
    On Error Resume Next
    col.Remove Key(s)
    col.Add value, Key(s)
End Sub

Private Sub CloseClient(ByVal s As Long)
    On Error Resume Next
    api_closesocket s
    m_inBuf.Remove Key(s)
    m_outBuf.Remove Key(s)
    m_closeAfter.Remove Key(s)
End Sub

' ================= input: frames =================

Private Sub ReadFrom(ByVal s As Long)
    Dim buf(0 To 16383) As Byte, n As Long, pending As String, p As Long, header As String, length As Long, body As String
    n = api_recv(s, buf(0), 16384, 0)
    If n <= 0 Then Exit Sub
    pending = m_inBuf(Key(s)) & StrConv(LeftB$(buf, n), vbUnicode)
    Do
        p = InStr(pending, vbLf)
        If p = 0 Then
            If Len(pending) > 9 Then ProtocolError s, "frame header too long": Exit Sub
            Exit Do
        End If
        header = Left$(pending, p - 1)
        If Not IsDigits(header) Then ProtocolError s, "invalid frame header": Exit Sub
        length = CLng(header)
        If length > MAX_BODY Then ProtocolError s, "frame too large": Exit Sub
        If Len(pending) - p < length Then Exit Do          ' body not complete yet
        body = Mid$(pending, p + 1, length)
        pending = Mid$(pending, p + 1 + length)
        HandleRequest s, body
        If Not SocketOpen(s) Then Exit Sub                  ' e.g. closed after a bad token
    Loop
    SetBuf m_inBuf, s, pending
End Sub

Private Function SocketOpen(ByVal s As Long) As Boolean
    Dim v As String
    On Error GoTo Done
    v = m_outBuf(Key(s))
    SocketOpen = True
Done:
End Function

Private Function IsDigits(ByVal t As String) As Boolean
    Dim i As Long
    If Len(t) = 0 Or Len(t) > 9 Then Exit Function
    For i = 1 To Len(t)
        If Mid$(t, i, 1) < "0" Or Mid$(t, i, 1) > "9" Then Exit Function
    Next
    IsDigits = True
End Function

Private Sub ProtocolError(ByVal s As Long, ByVal why As String)
    LogLine "protocol error on socket " & s & ": " & why
    CloseClient s
End Sub

' Request: MAGIC US token US id US verb [US arg]...
Private Sub HandleRequest(ByVal s As Long, ByVal body As String)
    Dim f() As String, args() As String, i As Long, id As String
    On Error GoTo Fail
    f = Split(body, US())
    If UBound(f) < 3 Then SendFrame s, MAGIC & US() & "?" & US() & "ERR" & US() & "400" & US() & "bad request": Exit Sub
    id = f(2)
    If f(0) <> MAGIC Then SendFrame s, MAGIC & US() & id & US() & "ERR" & US() & "400" & US() & "unsupported protocol " & f(0): Exit Sub
    If f(1) <> m_token Then
        SendFrame s, MAGIC & US() & id & US() & "ERR" & US() & "401" & US() & "bad token"
        SetBuf m_closeAfter, s, "1"
        Flush s
        Exit Sub
    End If
    m_requests = m_requests + 1
    StatusRefresh

    If UBound(f) >= 4 Then
        ReDim args(0 To UBound(f) - 4)
        For i = 4 To UBound(f): args(i - 4) = f(i): Next
    Else
        args = Split("")                                    ' empty array (UBound = -1)
    End If
    SendFrame s, MAGIC & US() & id & US() & Dispatch(f(3), args)
    Exit Sub
Fail:
    LogLine "HandleRequest error " & Err.Number & ": " & Err.Description
End Sub

' ================= output: frames with buffering =================

Private Sub SendFrame(ByVal s As Long, ByVal body As String)
    SetBuf m_outBuf, s, m_outBuf(Key(s)) & Len(body) & vbLf & body
    Flush s
End Sub

' Sends as much as the socket takes; the rest waits for FD_WRITE (large modules on a
' non-blocking socket can return WSAEWOULDBLOCK).
Private Sub Flush(ByVal s As Long)
    Dim pending As String, bytes() As Byte, n As Long, total As Long
    On Error GoTo Fail
    pending = m_outBuf(Key(s))
    Do While Len(pending) > 0
        bytes = StrConv(pending, vbFromUnicode)
        total = UBound(bytes) + 1
        n = api_send(s, bytes(0), total, 0)
        If n = SOCKET_ERROR Then
            If WSAGetLastError() <> WSAEWOULDBLOCK Then LogLine "send failed " & WSAGetLastError(): CloseClient s: Exit Sub
            Exit Do                                         ' wait for FD_WRITE
        End If
        pending = Mid$(pending, n + 1)
    Loop
    SetBuf m_outBuf, s, pending
    If Len(pending) = 0 And IsCloseAfter(s) Then CloseClient s
    Exit Sub
Fail:
    ' socket already closed / unknown key
End Sub

Private Function IsCloseAfter(ByVal s As Long) As Boolean
    On Error Resume Next
    IsCloseAfter = (m_closeAfter(Key(s)) = "1")
End Function

' ================= helpers =================

Private Function NewToken() As String
    Dim g(0 To 15) As Byte, i As Long, t As String
    CoCreateGuid g(0)
    For i = 0 To 15: t = t & Right$("0" & Hex$(g(i)), 2): Next
    NewToken = LCase$(t)
End Function

' %LOCALAPPDATA%\vb6mcp\instances\<pid>.json (docs/protocol.md "Discovery"). Written to a temp
' file first and renamed, so the server never reads a half-written file.
Private Sub WriteDiscovery()
    Dim f As Integer, folder As String, tmp As String
    EnsureDir BaseDir()
    folder = BaseDir() & "\instances"
    EnsureDir folder
    m_discoveryFile = folder & "\" & ProcessId() & ".json"
    tmp = m_discoveryFile & ".tmp"
    f = FreeFile
    Open tmp For Output As #f
    Print #f, "{""pid"":" & ProcessId() & ",""port"":" & m_port & ",""token"":""" & m_token & """,""protocol"":" & PROTOCOL_VERSION & _
              ",""addin"":""" & AddinVersion() & """,""started"":""" & Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss") & """}"
    Close #f
    On Error Resume Next
    Kill m_discoveryFile
    On Error GoTo 0
    Name tmp As m_discoveryFile
End Sub
