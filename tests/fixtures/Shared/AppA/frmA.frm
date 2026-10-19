VERSION 5.00
Begin VB.Form frmA 
   Caption         =   "AppA"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblPrice 
      Caption         =   "Price"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim pricing As New SharedLib.clsPricing
    ' FormatAmount always shows two decimals here
    lblTotal.Caption = "Total: " & FormatAmount(TotalOf(3, 4.5))
    lblPrice.Caption = pricing.PriceText(100, 0.2)
    Me.Tag = "FormatAmount("
End Sub
