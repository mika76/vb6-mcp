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
' Outcome of the last deferred MOUSEREAL/KEYSREAL ("sent" or why not), returned by MOUSERESTORE
Private m_lastReal As String
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function api_GetFocus Lib "user32" Alias "GetFocus" () As Long
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
    Dim pt(1) As Long, pt2(1) As Long, hit As Long, ins(15) As MOUSEIN, n As Long, i As Long, sent As Long
    If Not BringFormToFront(hWnd) Then RealMouse = "couldn't bring the form to the front": Exit Function
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

' Puts the cursor and the previously active window back after real input, and returns the outcome
' of the deferred MOUSEREAL/KEYSREAL ("sent", the reason it refused, or "not run" if it hasn't run yet).
Public Function RestoreRealMouse() As String
    RestoreRealMouse = IIf(m_lastReal = "", "not run", m_lastReal)
    m_lastReal = ""
    If Not m_realPending Then Exit Function
    SetCursorPos m_savedCursor(0), m_savedCursor(1)
    If m_savedFore <> 0 Then If IsWindow(m_savedFore) <> 0 Then SetForegroundWindow m_savedFore
    m_realPending = False
End Function

' Deferred MOUSEREAL / KEYSREAL (hard rule 14): run from WM_DEFER, outcome kept for MOUSERESTORE.
Public Sub DeferredRealMouse(ByVal hWnd As Long, ByVal action As String, ByVal x As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long)
    Dim why As String
    why = RealMouse(hWnd, action, x, y, x2, y2)
    m_lastReal = IIf(why = "", "sent", why)
End Sub

Public Sub DeferredRealKeys(ByVal hWnd As Long, ByVal events As String)
    Dim why As String
    why = RealKeys(hWnd, events)
    m_lastReal = IIf(why = "", "sent", why)
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

' Brings hWnd's top-level window to the front for real input (saving cursor and foreground window
' for RestoreRealMouse). A background process may not take the foreground; attaching to the
' foreground thread's input lifts that (no synthetic keys needed). True if it worked.
Private Function BringFormToFront(ByVal hWnd As Long) As Boolean
    Dim top As Long, fore As Long, foreThread As Long, myThread As Long, attached As Boolean, dummy As Long
    top = GetAncestor(hWnd, 2) ' GA_ROOT
    If Not m_realPending Then
        GetCursorPos m_savedCursor(0)
        m_savedFore = GetForegroundWindow()
        m_realPending = True
    End If
    fore = GetForegroundWindow()
    If fore <> top Then
        myThread = GetCurrentThreadId()
        foreThread = GetWindowThreadProcessId(fore, dummy)
        If foreThread <> 0 And foreThread <> myThread Then attached = (AttachThreadInput(myThread, foreThread, 1) <> 0)
        SetForegroundWindow top
        BringWindowToTop top
        If attached Then AttachThreadInput myThread, foreThread, 0
    End If
    BringFormToFront = (GetForegroundWindow() = top)
End Function

' The program's focused control: the program runs on the IDE thread, so GetFocus here sees it.
Public Function FocusedWindow() As Long
    FocusedWindow = api_GetFocus()
End Function

' Sets Shift/Ctrl/Alt in this thread's key state (mods: 1 Shift, 2 Ctrl, 4 Alt). Runs deferred
' between posted key messages, so the program's GetKeyState and TranslateMessage see the modifiers
' while it processes them (posted messages don't change the key state themselves, spike 17).
Public Sub SetModifierState(ByVal mods As Long)
    Dim ks(255) As Byte
    GetKeyboardState ks(0)
    ks(&H10) = IIf(mods And 1, &H80, 0): ks(&HA0) = ks(&H10)
    ks(&H11) = IIf(mods And 2, &H80, 0): ks(&HA2) = ks(&H11)
    ks(&H12) = IIf(mods And 4, &H80, 0): ks(&HA4) = ks(&H12)
    SetKeyboardState ks(0)
End Sub

' Posts key events to hWnd. events: tokens separated by ";": S<mods> modifier state (queued as a
' deferred KEYSTATE so it takes effect in order), D<vk> key down, U<vk> key up, C<code> character
' (WM_CHAR; for text). With Alt held, keys go as WM_SYSKEYDOWN/UP (menus). Returns the count.
' After each key down the rest is queued as a deferred KEYSCONT: VB's loop runs TranslateMessage on
' the key down, and the WM_CHAR it makes (Backspace, Enter, Ctrl+letter) goes to the END of the
' queue, behind characters already posted (Phase 3 live test: "ab{BACKSPACE}c" gave "ab"). The
' continuation runs after the key down was processed, so it posts behind that generated WM_CHAR.
Public Function PostKeys(ByVal hWnd As Long, ByVal events As String, Optional ByVal mods As Long = 0) As Long
    Dim tok() As String, i As Long, kind As String, v As Long, lp As Long, n As Long, rest As String, j As Long
    tok = Split(events, ";")
    For i = 0 To UBound(tok)
        kind = Left$(tok(i), 1): v = CLng(Mid$(tok(i), 2))
        Select Case kind
        Case "S": mods = v: Enqueue "KEYSTATE" & US() & v
        Case "D", "U"
            lp = 1 Or (MapVirtualKey(v, 0) * &H10000)
            If (v >= &H21 And v <= &H28) Or v = &H2D Or v = &H2E Then lp = lp Or &H1000000   ' extended keys
            If mods And 4 Then lp = lp Or &H20000000                                          ' Alt held
            If kind = "U" Then lp = lp Or &HC0000000
            If mods And 4 Then
                PostMessage hWnd, IIf(kind = "D", &H104, &H105), v, lp    ' WM_SYSKEYDOWN / UP
            Else
                PostMessage hWnd, IIf(kind = "D", WM_KEYDOWN, WM_KEYUP), v, lp
            End If
            If kind = "D" And i < UBound(tok) Then
                For j = i + 1 To UBound(tok)
                    If j > i + 1 Then rest = rest & ";"
                    rest = rest & tok(j)
                Next
                Enqueue "KEYSCONT" & US() & Hex$(hWnd) & US() & mods & US() & rest
                PostKeys = n + 1
                Exit Function
            End If
        Case "C": PostMessage hWnd, WM_CHAR, v, 1
        End Select
        n = n + 1
    Next
    PostKeys = n
End Function

' Real keyboard input (SendInput) to hWnd: brings its form to the front, focuses hWnd, sends the
' events (same format as PostKeys; characters as Unicode). The server follows with MOUSERESTORE.
Public Function RealKeys(ByVal hWnd As Long, ByVal events As String) As String
    Dim tok() As String, i As Long, kind As String, v As Long, mods As Long, newMods As Long
    Dim ins() As MOUSEIN, n As Long, sent As Long, b As Long
    If Not BringFormToFront(hWnd) Then RealKeys = "couldn't bring the form to the front": Exit Function
    api_SetFocus hWnd
    tok = Split(events, ";")
    ReDim ins(0 To 4 * (UBound(tok) + 1) + 6)
    For i = 0 To UBound(tok)
        kind = Left$(tok(i), 1): v = CLng(Mid$(tok(i), 2))
        Select Case kind
        Case "S"
            newMods = v
            For b = 0 To 2      ' press newly held modifiers, release dropped ones
                If (newMods And (2 ^ b)) <> 0 And (mods And (2 ^ b)) = 0 Then n = AddKey(ins, n, Choose(b + 1, &H10, &H11, &H12), 0, False)
                If (newMods And (2 ^ b)) = 0 And (mods And (2 ^ b)) <> 0 Then n = AddKey(ins, n, Choose(b + 1, &H10, &H11, &H12), 0, True)
            Next
            mods = newMods
        Case "D": n = AddKey(ins, n, v, 0, False)
        Case "U": n = AddKey(ins, n, v, 0, True)
        Case "C": n = AddKey(ins, n, 0, AscW(Chr$(v)), False): n = AddKey(ins, n, 0, AscW(Chr$(v)), True)
        End Select
    Next
    For b = 0 To 2              ' never leave a modifier held down
        If (mods And (2 ^ b)) <> 0 Then n = AddKey(ins, n, Choose(b + 1, &H10, &H11, &H12), 0, True)
    Next
    If n = 0 Then Exit Function
    sent = SendInput(n, ins(0), Len(ins(0)))
    LogLine "real keys: " & sent & "/" & n & " sent"
    If sent <> n Then RealKeys = "SendInput sent " & sent & " of " & n & " events"
End Function

' One keyboard INPUT in the MOUSEIN layout (x86: type, then wVk|wScan<<16, dwFlags, time, extra).
' uni <> 0: a Unicode character (KEYEVENTF_UNICODE) instead of a virtual key.
Private Function AddKey(ins() As MOUSEIN, ByVal n As Long, ByVal vk As Long, ByVal uni As Long, ByVal up As Boolean) As Long
    Dim flags As Long
    ins(n).itype = 1                                   ' INPUT_KEYBOARD
    If uni <> 0 Then
        ins(n).dx = (uni And &HFFFF&) * &H10000         ' wVk = 0, wScan = character
        flags = &H4                                    ' KEYEVENTF_UNICODE
    Else
        ins(n).dx = vk Or (MapVirtualKey(vk, 0) * &H10000)
        If (vk >= &H21 And vk <= &H28) Or vk = &H2D Or vk = &H2E Then flags = &H1   ' KEYEVENTF_EXTENDEDKEY
    End If
    If up Then flags = flags Or &H2                    ' KEYEVENTF_KEYUP
    ins(n).dy = flags
    ins(n).mouseData = 0: ins(n).dwFlags = 0: ins(n).time = 0: ins(n).extra = 0
    AddKey = n + 1
End Function

' Check box or option button of this process: BM_GETCHECK, 0 unchecked, 1 checked, 2 grayed.
' Read in-process: UIPI blocks BM_* from the server, like LB_* (app_tree, app_click's change check).
Public Function ButtonCheck(ByVal hWnd As Long) As Long
    ButtonCheck = SendMessage(hWnd, &HF0, 0, 0)
End Function

' List box or combo box of this process: kind US count US selected index (-1 none) US items (RS).
' Read in-process: UIPI blocks LB_*/CB_* from the server (spike 18).
Public Function ListState(ByVal hWnd As Long, ByVal isCombo As Boolean) As String
    Dim n As Long, i As Long, ln As Long, s As String, out As String, sel As Long
    If isCombo Then
        n = SendMessage(hWnd, &H146, 0, 0): sel = SendMessage(hWnd, &H147, 0, 0)        ' CB_GETCOUNT, CB_GETCURSEL
        For i = 0 To n - 1
            ln = SendMessage(hWnd, &H149, i, 0): If ln < 0 Then ln = 0                   ' CB_GETLBTEXTLEN
            s = String$(ln + 1, vbNullChar): ln = SendMessageStr(hWnd, &H148, i, s): If ln < 0 Then ln = 0   ' CB_GETLBTEXT
            If i > 0 Then out = out & RS()
            out = out & Left$(s, ln)
        Next
        ListState = "combo" & US() & n & US() & sel & US() & out
    Else
        sel = SendMessage(hWnd, &H188, 0, 0)                                              ' LB_GETCURSEL
        out = ListBoxItems(hWnd)
        ListState = "list" & US() & SendMessage(hWnd, &H18B, 0, 0) & US() & sel & US() & out
    End If
End Function

' Selects an item and notifies the parent with LBN_SELCHANGE / CBN_SELCHANGE (what a user's click
' produces), which is where VB raises the control's Click event. Deferred: Click runs user code.
Public Sub SelectItem(ByVal hWnd As Long, ByVal index As Long)
    Dim isCombo As Boolean
    isCombo = (InStr(1, OwnWindowClass(hWnd), "ComboBox", vbTextCompare) > 0)
    SendMessage hWnd, IIf(isCombo, &H14E, &H186), index, 0                                ' CB_SETCURSEL / LB_SETCURSEL
    SendMessage GetParent(hWnd), &H111, (GetDlgCtrlID(hWnd) And &HFFFF&) Or &H10000, hWnd  ' WM_COMMAND, code 1 = *_SELCHANGE
    LogLine "selected item " & index & " of " & Hex$(hWnd)
End Sub
