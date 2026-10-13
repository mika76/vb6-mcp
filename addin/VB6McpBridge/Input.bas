Attribute VB_Name = "Input"
' Input into this IDE's own windows (channel C writes). Medium-integrity processes can't send
' input to the elevated IDE (spike 10), so the add-in does it in-process. Only windows of this
' process are accepted (AGENTS.md hard rule 8).
Option Explicit

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const BM_CLICK As Long = &HF5

Public Function IsOwnButton(ByVal hWnd As Long) As Boolean
    Dim pid As Long, cls As String, n As Long
    If IsWindow(hWnd) = 0 Then Exit Function
    GetWindowThreadProcessId hWnd, pid
    If pid <> ProcessId() Then Exit Function
    cls = String$(64, vbNullChar)
    n = GetClassName(hWnd, cls, 64)
    IsOwnButton = (LCase$(Left$(cls, n)) = "button")
End Function

' Posted, not sent: SendMessage would run the dialog's reaction inside our request handler.
Public Sub PostClick(ByVal hWnd As Long)
    PostMessage hWnd, BM_CLICK, 0, 0
End Sub
