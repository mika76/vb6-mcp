VERSION 5.00
Begin VB.Form frmShop 
   Caption         =   "Shop"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotal 
      Locked          =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   5415
   End
   Begin VB.ListBox lstLines 
      Height          =   1425
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   5415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   450
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   450
      Width           =   1215
   End
   Begin VB.TextBox txtQty 
      Text            =   "1"
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox cboProduct 
      Style           =   2  'Dropdown List
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblQty 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblProduct 
      Caption         =   "Product"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_total As Currency

Private Sub Form_Load()
    Dim i As Long
    LoadProducts App.Path & "\products.dat"
    For i = 1 To UBound(Products)
        If Not Listed(RTrim$(Products(i).Name)) Then cboProduct.AddItem RTrim$(Products(i).Name)
    Next
    cboProduct.ListIndex = 0
    ShowTotal
End Sub

' True if the combo already has this product
Private Function Listed(ByVal productName As String) As Boolean
    Dim i As Long
    For i = 0 To cboProduct.ListCount - 1
        If cboProduct.List(i) = productName Then Listed = True
    Next
End Function

Private Sub cmdAdd_Click()
    Dim amount As Currency
    If Val(txtQty.Text) <= 0 Then
        MsgBox "Enter a quantity", vbExclamation
        Exit Sub
    End If
    amount = FindPrice(cboProduct.Text) * CLng(Val(txtQty.Text))
    m_total = m_total + amount
    lstLines.AddItem txtQty.Text & " x " & cboProduct.Text & " = " & Format$(amount, "0.00")
    ShowTotal
End Sub

Private Sub cmdClear_Click()
    m_total = 0
    lstLines.Clear
    ShowTotal
End Sub

Private Sub ShowTotal()
    txtTotal.Text = "Total: " & Format$(m_total, "0.00")
End Sub
