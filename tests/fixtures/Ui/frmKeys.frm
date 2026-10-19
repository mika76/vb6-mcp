VERSION 5.00
Begin VB.Form frmKeys 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblLast 
      Caption         =   "(no key yet)"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4200
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Press any key to continue (Esc closes)"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
End
Attribute VB_Name = "frmKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Labels only, like a production project's frmMessage ("Press Any Key To Continue"): no control can take the focus, so the form
' itself gets the keys (improvement notes 54). Each key is shown in lblLast and logged; Esc closes the form.

Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Private Sub Log(ByVal sText As String)
    lblLast.Caption = sText
    Debug.Print "frmKeys: " & sText
    OutputDebugString "frmKeys: " & sText & vbCrLf
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Log "KeyDown " & KeyCode & " shift " & Shift
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Log "KeyPress " & KeyAscii & " '" & Chr$(KeyAscii) & "'"
End Sub
