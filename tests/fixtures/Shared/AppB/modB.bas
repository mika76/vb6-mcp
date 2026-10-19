Attribute VB_Name = "modB"
Option Explicit

Sub Main()
    Dim v(1) As Currency
    v(0) = 1.5: v(1) = 2.25
    Debug.Print FormatAmountList(v)
    Debug.Print "Ten: " & FormatAmount(10)
End Sub
