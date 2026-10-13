VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug fixture"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoop 
      Caption         =   "Loop"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdError 
      Caption         =   "Error"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblOut 
      Caption         =   "Result"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompute_Click()
    lblOut.Caption = "Outer(4) = " & Outer(4)
End Sub

' Division by zero: runtime error 11, unhandled
Private Sub cmdError_Click()
    lblOut.Caption = "Divide = " & Divide(1, 0)
End Sub

' Busy loop without DoEvents: keeps the IDE thread busy, so only the watchdog can break in (Phase 2b)
Private Sub cmdLoop_Click()
    Dim i As Long
    Do
        i = i + 1
        If i > 2000000000 Then i = 0
    Loop
End Sub
