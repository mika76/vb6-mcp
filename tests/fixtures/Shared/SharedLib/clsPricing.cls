VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Net amount plus tax, as text.
Public Function PriceText(ByVal net As Currency, ByVal taxRate As Double) As String
    PriceText = FormatAmount(net) & " + tax " & _
        FormatAmount(net * taxRate)
End Function
