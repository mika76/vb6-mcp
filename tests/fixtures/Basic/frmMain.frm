VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Basic fixture"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblResult 
      Caption         =   "Result"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_counter As clsCounter

' Shows the running total in the label.
Private Sub cmdAdd_Click()
    m_counter.Increment
    lblResult.Caption = "Total: " & AddNumbers(m_counter.Value, 0)
End Sub

Private Sub Form_Load()
    Set m_counter = New clsCounter
End Sub