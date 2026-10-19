VERSION 5.00
Begin VB.Form frmCart 
   Caption         =   "Order"
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
      Height          =   315
      Left            =   1680
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
   Begin VB.Label lblQty 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
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

Private m_lines As Collection
Private m_line As New clsLine

Private Sub Form_Load()
    Set m_lines = New Collection
    ShowTotal
End Sub

Private Sub cmdAdd_Click()
    If Val(txtQty.Text) <= 0 Then
        MsgBox "Enter a quantity", vbExclamation
        Exit Sub
    End If
    m_line.Price = CCur(Val(txtPrice.Text))
    m_line.Quantity = CLng(Val(txtQty.Text))
    m_lines.Add m_line
    lstLines.AddItem m_line.Describe
    ShowTotal
End Sub

Private Sub cmdClear_Click()
    Set m_lines = New Collection
    lstLines.Clear
    ShowTotal
End Sub

Private Sub ShowTotal()
    txtTotal.Text = "Total: " & Format$(CartTotal(m_lines), "0.00")
End Sub
