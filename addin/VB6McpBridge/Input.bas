Attribute VB_Name = "Input"
' Input into this IDE's own windows (channel C writes). Medium-integrity processes can't send
' input to the elevated IDE (spike 10), so the add-in does it in-process. Only windows of this
' process are accepted (AGENTS.md hard rule 8).
Option Explicit

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function api_SetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_SETTEXT As Long = &HC
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long

Private Const BM_CLICK As Long = &HF5

' The user's clipboard text while the Immediate window is being read (IMMSELECT .. IMMRESTORE)
Private m_savedClip As String
Private m_hadClip As Boolean
' The window being read, so IMMRESTORE can collapse its selection again
Public g_readHwnd As Long
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const VK_RETURN As Long = &HD
Private Const VK_CONTROL As Long = &H11
Private Const VK_END As Long = &H23

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Long) As Long
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_HSCROLL As Long = &H114
Private Const WM_VSCROLL As Long = &H115
Private Const MK_LBUTTON As Long = &H1
Private Const MK_RBUTTON As Long = &H2
Private Const SB_ENDSCROLL As Long = 8
' Real input (spike 17): VB raises Click on a windowless control (Label, Image) only for real mouse
' input, never for posted or sent messages, so MOUSEREAL uses SendInput.
Private Type MOUSEIN
    itype As Long
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    time As Long
    extra As Long
End Type
Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As MOUSEIN, ByVal cbSize As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Private Const SBM_GETPOS As Long = &HE1
' Saved by MOUSEREAL, put back by MOUSERESTORE
Private m_realPending As Boolean
Private m_savedCursor(1) As Long
Private m_savedFore As Long
Public Function IsOwnButton(ByVal hWnd As Long) As Boolean
    Dim pid As Long, cls As String, n As Long
    If IsWindow(hWnd) = 0 Then Exit Function
    GetWindowThreadProcessId hWnd, pid
    If pid <> ProcessId() Then Exit Function
    cls = String$(64, vbNullChar)
    n = GetClassName(hWnd, cls, 64)
    IsOwnButton = (LCase$(Left$(cls, n)) = "button")
End Function

' Class of a window of this process, or "" if it isn't one (hard rule 8: input only to our windows).
Public Function OwnWindowClass(ByVal hWnd As Long) As String
    Dim pid As Long, cls As String, n As Long
    If IsWindow(hWnd) = 0 Then Exit Function
    GetWindowThreadProcessId hWnd, pid
    If pid <> ProcessId() Then Exit Function
    cls = String$(64, vbNullChar)
    n = GetClassName(hWnd, cls, 64)
    OwnWindowClass = Left$(cls, n)
End Function

' WM_SETTEXT (deferred caller only: a text box raises its Change event).
Public Sub SetWindowText(ByVal hWnd As Long, ByVal text As String)
    SendMessageStr hWnd, WM_SETTEXT, 0, text
    LogLine "set text of " & Hex$(hWnd)
End Sub

Public Function IsOwnWindowOfClass(ByVal hWnd As Long, ByVal className As String) As Boolean
    Dim pid As Long, cls As String, n As Long
    If IsWindow(hWnd) = 0 Then Exit Function
    GetWindowThreadProcessId hWnd, pid
    If pid <> ProcessId() Then Exit Function
    cls = String$(64, vbNullChar)
    n = GetClassName(hWnd, cls, 64)
    IsOwnWindowOfClass = (StrComp(Left$(cls, n), className, vbTextCompare) = 0)
End Function

Public Sub SaveUserClipboard()
    On Error Resume Next
    m_hadClip = Clipboard.GetFormat(vbCFText)
    m_savedClip = ""
    If m_hadClip Then m_savedClip = Clipboard.GetText(vbCFText)
End Sub

' Puts the user's clipboard text back (non-text formats can't be preserved).
Public Sub RestoreUserClipboard()
    On Error Resume Next
    If g_readHwnd <> 0 Then CaretToEnd g_readHwnd: g_readHwnd = 0     ' undo Select All
    Clipboard.Clear
    If m_hadClip Then Clipboard.SetText m_savedClip, vbCFText
    m_hadClip = False
    m_savedClip = ""
End Sub

' Keyboard focus to one of this IDE's windows. We run on the IDE's own thread, so the Win32
' SetFocus works directly (no foreground rules apply within the thread).
Public Sub FocusWindow(ByVal hWnd As Long)
    api_SetFocus hWnd
End Sub

' Ctrl+End in the window: caret to the end of the text, which also collapses a selection. Reading
' the Immediate window leaves everything selected (Select All), and the next typed line then
' replaced the whole content, erasing the user's Debug.Print output (Phase 2 live test, seen in a
' screenshot). Posted/sent key messages carry no modifier state, so Ctrl is set in this thread's
' keyboard state (we run on the IDE thread) and restored afterwards.
Public Sub CaretToEnd(ByVal hWnd As Long)
    Dim saved(0 To 255) As Byte, withCtrl(0 To 255) As Byte, i As Long
    GetKeyboardState saved(0)
    For i = 0 To 255: withCtrl(i) = saved(i): Next
    withCtrl(VK_CONTROL) = &H80
    SetKeyboardState withCtrl(0)
    SendMessage hWnd, WM_KEYDOWN, VK_END, 0
    SendMessage hWnd, WM_KEYUP, VK_END, &HC0000000
    SetKeyboardState saved(0)
End Sub

' Characters as WM_CHAR (only called from deferred actions, never inside a request handler).
Public Sub SendText(ByVal hWnd As Long, ByVal text As String)
    Dim i As Long
    For i = 1 To Len(text)
        SendMessage hWnd, WM_CHAR, Asc(Mid$(text, i, 1)), 0
    Next
End Sub

Public Sub SendEnter(ByVal hWnd As Long)
    SendMessage hWnd, WM_KEYDOWN, VK_RETURN, 0
    SendMessage hWnd, WM_KEYUP, VK_RETURN, &HC0000000
End Sub

' Posted, not sent: SendMessage would run the dialog's reaction inside our request handler.
Public Sub PostClick(ByVal hWnd As Long)
    PostMessage hWnd, BM_CLICK, 0, 0
End Sub

' Mouse input to a window of this process, in that window's client coordinates (posted, so any
' event code runs later from the message loop, never under the request handler: hard rule 14).
' A point over a Label/Image is posted to the form: VB's runtime routes it to the windowless control.
Public Sub PostMouse(ByVal hWnd As Long, ByVal action As String, ByVal x As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long)
    Dim i As Long, steps As Long, mx As Long, my As Long
    Select Case action
    Case "click"
        PostMessage hWnd, WM_MOUSEMOVE, 0, MakeLParam(x, y)
        PostMessage hWnd, WM_LBUTTONDOWN, MK_LBUTTON, MakeLParam(x, y)
        PostMessage hWnd, WM_LBUTTONUP, 0, MakeLParam(x, y)
    Case "dblclick"
        ' Windows' own sequence: down, up, DBLCLK (instead of the second down), up
        PostMessage hWnd, WM_MOUSEMOVE, 0, MakeLParam(x, y)
        PostMessage hWnd, WM_LBUTTONDOWN, MK_LBUTTON, MakeLParam(x, y)
        PostMessage hWnd, WM_LBUTTONUP, 0, MakeLParam(x, y)
        PostMessage hWnd, WM_LBUTTONDBLCLK, MK_LBUTTON, MakeLParam(x, y)
        PostMessage hWnd, WM_LBUTTONUP, 0, MakeLParam(x, y)
    Case "rclick"
        PostMessage hWnd, WM_MOUSEMOVE, 0, MakeLParam(x, y)
        PostMessage hWnd, WM_RBUTTONDOWN, MK_RBUTTON, MakeLParam(x, y)
        PostMessage hWnd, WM_RBUTTONUP, 0, MakeLParam(x, y)
    Case "drag"
        ' Button down at (x, y), moves in steps with the button held, button up at (x2, y2)
        PostMessage hWnd, WM_MOUSEMOVE, 0, MakeLParam(x, y)
        PostMessage hWnd, WM_LBUTTONDOWN, MK_LBUTTON, MakeLParam(x, y)
        steps = 10
        For i = 1 To steps
            mx = x + ((x2 - x) * i) \ steps
            my = y + ((y2 - y) * i) \ steps
            PostMessage hWnd, WM_MOUSEMOVE, MK_LBUTTON, MakeLParam(mx, my)
        Next
        PostMessage hWnd, WM_LBUTTONUP, 0, MakeLParam(x2, y2)
    Case "wheel"
        ' x2 = notches (positive = away from the user = up). WM_MOUSEWHEEL takes screen coordinates.
        Dim pt(1) As Long
        pt(0) = x: pt(1) = y
        ClientToScreen hWnd, pt(0)
        For i = 1 To Abs(x2)
            PostMessage hWnd, WM_MOUSEWHEEL, Sgn(x2) * 120 * &H10000, MakeLParam(pt(0), pt(1))
        Next
    End Select
    LogLine "mouse " & action & " to " & Hex$(hWnd) & " at " & x & "," & y
End Sub

' Scroll bar commands (SB_LINEUP=0 .. SB_BOTTOM=7). A scroll bar control notifies its parent, so for
' one the message goes to the parent with the control's handle in lParam; any other window gets it
' itself (its own scroll bars). Ends with SB_ENDSCROLL as a real scroll does.
Public Sub PostScroll(ByVal hWnd As Long, ByVal vertical As Boolean, ByVal code As Long, ByVal count As Long, ByVal isScrollBar As Boolean)
    Dim msg As Long, target As Long, lp As Long, i As Long
    If vertical Then msg = WM_VSCROLL Else msg = WM_HSCROLL
    If isScrollBar Then target = GetParent(hWnd): lp = hWnd Else target = hWnd: lp = 0
    For i = 1 To count
        PostMessage target, msg, code, lp
    Next
    PostMessage target, msg, SB_ENDSCROLL, lp
    LogLine "scroll " & code & " x" & count & " to " & Hex$(hWnd)
End Sub

' MAKELPARAM for signed 16-bit coordinates without overflowing a Long
Private Function MakeLParam(ByVal lo As Long, ByVal hi As Long) As Long
    lo = lo And &HFFFF&
    hi = hi And &HFFFF&
    If hi >= &H8000& Then MakeLParam = (hi - &H10000) * &H10000 + lo Else MakeLParam = hi * &H10000 + lo
End Function

' Real mouse input (SendInput) at client coordinates of hWnd. Brings hWnd's form to the front first
' and refuses if the point isn't on a window of this process (hard rule 8). Saves the cursor and
' foreground window for RestoreRealMouse, which the server calls once VB has processed the input
' (input is handled after this request returns; the add-in never waits, hard rule 5).
' Returns "" on success or the reason it refused.
Public Function RealMouse(ByVal hWnd As Long, ByVal action As String, ByVal x As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long) As String
    Dim top As Long, fore As Long, foreThread As Long, myThread As Long, attached As Boolean, dummy As Long
    Dim pt(1) As Long, pt2(1) As Long, hit As Long, ins(15) As MOUSEIN, n As Long, i As Long, sent As Long
    top = GetAncestor(hWnd, 2) ' GA_ROOT
    If Not m_realPending Then
        GetCursorPos m_savedCursor(0)
        m_savedFore = GetForegroundWindow()
        m_realPending = True
    End If
    fore = GetForegroundWindow()
    If fore <> top Then
        ' a background process may not take the foreground; attaching to the foreground thread's
        ' input lifts that (no synthetic keys needed)
        myThread = GetCurrentThreadId()
        foreThread = GetWindowThreadProcessId(fore, dummy)
        If foreThread <> 0 And foreThread <> myThread Then attached = (AttachThreadInput(myThread, foreThread, 1) <> 0)
        SetForegroundWindow top
        BringWindowToTop top
        If attached Then AttachThreadInput myThread, foreThread, 0
    End If
    If GetForegroundWindow() <> top Then RealMouse = "couldn't bring the form to the front": Exit Function
    pt(0) = x: pt(1) = y: ClientToScreen hWnd, pt(0)
    pt2(0) = x2: pt2(1) = y2: ClientToScreen hWnd, pt2(0)
    hit = WindowFromPoint(pt(0), pt(1))
    If OwnWindowClass(hit) = "" Then RealMouse = "the point is covered by another program's window": Exit Function
    Select Case action
    Case "click":    n = AddMove(ins, 0, pt(0), pt(1)): n = AddButton(ins, n, &H2): n = AddButton(ins, n, &H4)
    Case "dblclick": n = AddMove(ins, 0, pt(0), pt(1)): n = AddButton(ins, n, &H2): n = AddButton(ins, n, &H4): n = AddButton(ins, n, &H2): n = AddButton(ins, n, &H4)
    Case "rclick":   n = AddMove(ins, 0, pt(0), pt(1)): n = AddButton(ins, n, &H8): n = AddButton(ins, n, &H10)
    Case "drag"
        n = AddMove(ins, 0, pt(0), pt(1)): n = AddButton(ins, n, &H2)
        For i = 1 To 10
            n = AddMove(ins, n, pt(0) + ((pt2(0) - pt(0)) * i) \ 10, pt(1) + ((pt2(1) - pt(1)) * i) \ 10)
        Next
        n = AddButton(ins, n, &H4)
    End Select
    sent = SendInput(n, ins(0), Len(ins(0)))
    LogLine "real mouse " & action & " at screen " & pt(0) & "," & pt(1) & ": " & sent & "/" & n & " sent"
    If sent <> n Then RealMouse = "SendInput sent " & sent & " of " & n & " events"
End Function

' Puts the cursor and the previously active window back after RealMouse.
Public Sub RestoreRealMouse()
    If Not m_realPending Then Exit Sub
    SetCursorPos m_savedCursor(0), m_savedCursor(1)
    If m_savedFore <> 0 Then If IsWindow(m_savedFore) <> 0 Then SetForegroundWindow m_savedFore
    m_realPending = False
End Sub

' Scroll bar control range and position (read in-process: UIPI blocks SBM_* from the server).
Public Function ScrollBarPos(ByVal hWnd As Long) As String
    Dim mn As Long, mx As Long
    GetScrollRange hWnd, 2, mn, mx ' SB_CTL
    ScrollBarPos = mn & US() & mx & US() & SendMessage(hWnd, SBM_GETPOS, 0, 0)
End Function

' Absolute move to a screen point, normalized to the virtual desktop (all monitors)
Private Function AddMove(ins() As MOUSEIN, ByVal n As Long, ByVal sx As Long, ByVal sy As Long) As Long
    Dim vx As Long, vy As Long, vw As Long, vh As Long
    vx = GetSystemMetrics(76): vy = GetSystemMetrics(77): vw = GetSystemMetrics(78): vh = GetSystemMetrics(79)
    ins(n).itype = 0
    ins(n).dx = ((sx - vx) * 65535) \ (vw - 1)
    ins(n).dy = ((sy - vy) * 65535) \ (vh - 1)
    ins(n).dwFlags = &H1 Or &H8000& Or &H4000& ' MOVE | ABSOLUTE | VIRTUALDESK
    AddMove = n + 1
End Function

Private Function AddButton(ins() As MOUSEIN, ByVal n As Long, ByVal flags As Long) As Long
    ins(n).itype = 0
    ins(n).dwFlags = flags
    AddButton = n + 1
End Function

' Items of a list box of this process, separated by RS (LB_GETCOUNT / LB_GETTEXTLEN / LB_GETTEXT).
Public Function ListBoxItems(ByVal hWnd As Long) As String
    Dim n As Long, i As Long, ln As Long, s As String, out As String
    n = SendMessage(hWnd, &H18B, 0, 0)
    For i = 0 To n - 1
        ln = SendMessage(hWnd, &H18A, i, 0)
        If ln < 0 Then ln = 0
        s = String$(ln + 1, vbNullChar)
        ln = SendMessageStr(hWnd, &H189, i, s)
        If ln < 0 Then ln = 0
        If i > 0 Then out = out & RS()
        out = out & Left$(s, ln)
    Next
    ListBoxItems = out
End Function
