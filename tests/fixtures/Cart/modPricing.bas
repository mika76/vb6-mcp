Attribute VB_Name = "modPricing"
Option Explicit

' The third and later lines of an order get 10% off.
Public Function CartTotal(lines As Collection) As Currency
    Dim i As Long, total As Currency
    For i = 1 To lines.Count
        total = total + LineTotal(lines(i), i)
    Next
    CartTotal = total
End Function

Private Function LineTotal(ByVal item As clsLine, ByVal position As Long) As Currency
    LineTotal = item.Amount
    If position > 2 Then LineTotal = LineTotal * 0.9
End Function
