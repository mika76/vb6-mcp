VERSION 5.00
Begin VB.Form frmCart 
   Caption         =   "Order"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotal 
      Locked          =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   6735
   End
   Begin VB.ListBox lstLines 
      Height          =   1425
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   6735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   450
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   450
      Width           =   1215
   End
   Begin VB.ComboBox cboCategory 
      Style           =   2  'Dropdown List
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblCategory 
      Caption         =   "Category"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblQty 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblPrice 
      Caption         =   "Price"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Gross amounts of the lines (net plus tax)
Private m_lines As Collection
Private m_rates As New RateLib.clsRates

Private Sub Form_Load()
    Set m_lines = New Collection
    cboCategory.AddItem "Food"
    cboCategory.AddItem "Standard"
    cboCategory.AddItem "Books"
    cboCategory.ListIndex = 1
    ShowTotal
End Sub

' The category's code for RateLib
Private Function CategoryCode() As String
    CategoryCode = UCase$(Left$(cboCategory.Text, 4))
End Function

Private Sub cmdAdd_Click()
    Dim net As Currency
    If Val(txtQty.Text) <= 0 Then
        MsgBox "Enter a quantity", vbExclamation
        Exit Sub
    End If
    net = CCur(Val(txtPrice.Text)) * CLng(Val(txtQty.Text))
    m_lines.Add net * (1 + m_rates.RateFor(CategoryCode()))
    lstLines.AddItem txtQty.Text & " x " & txtPrice.Text & " (" & cboCategory.Text & ")"
    ShowTotal
End Sub

Private Sub cmdClear_Click()
    Set m_lines = New Collection
    lstLines.Clear
    ShowTotal
End Sub

Private Sub ShowTotal()
    Dim i As Long, total As Currency
    For i = 1 To m_lines.Count
        total = total + m_lines(i)
    Next
    txtTotal.Text = "Total: " & Format$(total, "0.00")
End Sub
