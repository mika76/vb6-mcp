Attribute VB_Name = "modOrder"
Option Explicit

Public Prices() As Currency
Public Qtys() As Long
Public LineCount As Long

Public Sub AddLine(ByVal price As Currency, ByVal qty As Long)
    LineCount = LineCount + 1
    ReDim Preserve Prices(1 To LineCount)
    ReDim Preserve Qtys(1 To LineCount)
    Prices(LineCount) = price
    Qtys(LineCount) = qty
End Sub

Public Sub ClearLines()
    LineCount = 0
    Erase Prices
    Erase Qtys
End Sub

Public Function OrderTotal() As Currency
    Dim i As Long, total As Currency
    For i = 1 To LineCount
        total = total + LineAmount(Prices(i), Qtys(i))
    Next
    OrderTotal = total
End Function

' 10 or more of an item: 10% off that line.
Private Function LineAmount(price As Currency, qty As Long) As Currency
    If qty >= 10 Then price = price * 0.9
    LineAmount = price * qty
End Function
