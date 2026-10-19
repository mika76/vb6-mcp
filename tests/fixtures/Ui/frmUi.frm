VERSION 5.00
Begin VB.Form frmUi 
   Caption         =   "UI fixture"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKeys 
      Caption         =   "&Keys..."
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtDrop 
      Height          =   375
      Left            =   1680
      OLEDropMode     =   2  'Automatic
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDrag 
      Caption         =   "Drag me"
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Option"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox lstItems 
      Height          =   1815
      Left            =   240
      OLEDragMode     =   1  'Automatic
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.VScrollBar vsbValue 
      Height          =   1815
      Left            =   3120
      Max             =   100
      TabIndex        =   3
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox picDraw 
      Height          =   1815
      Left            =   3720
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   4
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lblEcho 
      Caption         =   "Echo:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   6615
   End
   Begin VB.Label lblClick 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click me"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   855
      Left            =   3120
      TabIndex        =   7
      Top             =   3240
      Width           =   3735
   End
End
Attribute VB_Name = "frmUi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UI fixture for the app_* tools (docs/tools.md "The running program"): every control reports what
' happened to it in a label, so screenshots and app_tree can verify input.
Option Explicit

' Logging the way compiled VB6 programs do it (Debug.Print is compiled out): debug_output captures it
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Private m_clicks As Long
Private m_dragX As Single
Private m_dragY As Single
Private m_dragging As Boolean

Private Sub Form_Load()
    Dim i As Long
    For i = 1 To 50
        lstItems.AddItem "Item " & i
    Next
    cboColor.AddItem "Red": cboColor.AddItem "Green": cboColor.AddItem "Blue"
End Sub

Private Sub txtName_Change()
    lblEcho.Caption = "Echo: " & txtName.Text
    OutputDebugString "UI fixture: text = " & txtName.Text & vbCrLf
End Sub

Private Sub chkOption_Click()
    lblStatus.Caption = "Option: " & chkOption.Value
End Sub

Private Sub lstItems_DblClick()
    lblStatus.Caption = "Double-clicked: " & lstItems.Text
End Sub

Private Sub lstItems_Click()
    lblStatus.Caption = "Selected: " & lstItems.Text
End Sub

Private Sub vsbValue_Change()
    lblStatus.Caption = "Scroll value: " & vsbValue.Value
End Sub

Private Sub lblClick_Click()
    m_clicks = m_clicks + 1
    lblClick.Caption = "Label clicked " & m_clicks
    ' also in the form caption: readable with GetWindowText, a Label has no window (spike 17)
    Me.Caption = "UI fixture - label clicks: " & m_clicks
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_dragging = True
    m_dragX = X: m_dragY = Y
    picDraw.Cls
    picDraw.PSet (X, Y)
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_dragging Then picDraw.Line -(X, Y)
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_dragging = False
    lblStatus.Caption = "Drag from (" & m_dragX & "," & m_dragY & ") to (" & X & "," & Y & ")"
End Sub

Private Sub cboColor_Click()
    lblStatus.Caption = "Color: " & cboColor.Text
End Sub

' Key code and Shift state in the form caption (readable with app_windows): shows whether modifiers arrive
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.Caption = "UI fixture - KeyDown " & KeyCode & " shift " & Shift
End Sub

' Classic VB drag (cmdDrag has DragMode automatic) dropped on the picture box
Private Sub picDraw_DragDrop(Source As Control, X As Single, Y As Single)
    lblStatus.Caption = "Dropped " & Source.Name & " at (" & X & "," & Y & ")"
End Sub

' OLE drop (txtDrop has OLEDropMode automatic): the text arrives by itself; report it
Private Sub txtDrop_Change()
    lblStatus.Caption = "OLE drop: " & txtDrop.Text
End Sub

Private Sub cmdKeys_Click()
    ' a form with only Labels takes keys itself (improvement notes 54)
    frmKeys.Show vbModal, Me
End Sub
