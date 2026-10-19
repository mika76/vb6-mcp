VERSION 5.00
Begin VB.Form frmClient 
   Caption         =   "Client"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim tax As New PriceLib.clsTax
    txtResult.Text = CStr(tax.AddTax(100))
End Sub
