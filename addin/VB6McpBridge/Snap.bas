Attribute VB_Name = "Snap"
' Window screenshots from inside the IDE process. A non-elevated process can't PrintWindow an
' elevated VB6 (spike 15), but the add-in runs inside VB6, so it can. PrintWindow renders covered
' windows too, including lightweight controls like Labels (spike 15). Output: a 32-bit BMP in
' %LOCALAPPDATA%\vb6mcp\shots\<id>.bmp, written to a temp file and renamed so the server never
' reads a half-written image. Always called deferred: painting can run user code (Form_Paint).
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long

Private Const PW_RENDERFULLCONTENT As Long = 2

Public Function ShotsDir() As String
    ShotsDir = BaseDir() & "\shots"
End Function

' Captures hWnd into ShotsDir\<id>.bmp. Logs and gives up on failure (the server times out waiting).
Public Sub CaptureWindow(ByVal hWnd As Long, ByVal id As String)
    Dim r As RECT, w As Long, h As Long, screenDC As Long, memDC As Long, bmp As Long, oldBmp As Long
    Dim bih As BITMAPINFOHEADER, bits() As Byte, f As Integer, path As String, tmp As String, ok As Long
    Dim magic(0 To 1) As Byte, fileSize As Long, reserved As Long, pixelOffset As Long
    On Error GoTo Fail
    GetWindowRect hWnd, r
    w = r.Right - r.Left: h = r.Bottom - r.Top
    If w <= 0 Or h <= 0 Then LogLine "snap: zero-size window": Exit Sub

    screenDC = GetDC(0)
    memDC = CreateCompatibleDC(screenDC)
    bmp = CreateCompatibleBitmap(screenDC, w, h)
    oldBmp = SelectObject(memDC, bmp)
    ok = PrintWindow(hWnd, memDC, PW_RENDERFULLCONTENT)
    SelectObject memDC, oldBmp

    bih.biSize = 40: bih.biWidth = w: bih.biHeight = h       ' positive height = bottom-up rows
    bih.biPlanes = 1: bih.biBitCount = 32: bih.biSizeImage = w * h * 4
    ReDim bits(0 To w * h * 4 - 1)
    GetDIBits memDC, bmp, 0, h, bits(0), bih, 0               ' DIB_RGB_COLORS
    DeleteObject bmp: DeleteDC memDC: ReleaseDC 0, screenDC

    EnsureDir BaseDir()
    EnsureDir ShotsDir()
    path = ShotsDir() & "\" & id & ".bmp"
    tmp = path & ".tmp"
    f = FreeFile
    Open tmp For Binary Access Write As #f
    ' BITMAPFILEHEADER, written field by field (Put needs variables, not expressions)
    magic(0) = 66: magic(1) = 77                              ' "BM"
    fileSize = 54 + w * h * 4
    reserved = 0
    pixelOffset = 54
    Put #f, , magic
    Put #f, , fileSize
    Put #f, , reserved
    Put #f, , pixelOffset
    Put #f, , bih
    Put #f, , bits
    Close #f
    On Error Resume Next
    Kill path
    On Error GoTo Fail
    Name tmp As path
    LogLine "snap " & id & ": " & w & "x" & h & IIf(ok = 0, " (PrintWindow returned 0)", "")
    Exit Sub
Fail:
    LogLine "snap failed: " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #f
End Sub
