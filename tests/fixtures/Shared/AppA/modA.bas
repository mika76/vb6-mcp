Attribute VB_Name = "modA"
Option Explicit

Public Function TotalOf(ByVal count As Long, ByVal price As Currency) As Currency
    TotalOf = count * price
End Function

Public Function Summary(ByVal a As Currency, ByVal b As Currency) As String
    Summary = "A=" & FormatAmount(a) & _
              ", B=" & FormatAmount( _
                  b)
End Function
