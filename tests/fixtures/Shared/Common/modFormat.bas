Attribute VB_Name = "modFormat"
' Shared by SharedLib, AppA and AppB
Option Explicit

' Formats an amount with two decimals and thousands separators.
Public Function FormatAmount(ByVal value As Currency) As String
    FormatAmount = Format$(value, "#,##0.00")
End Function

' A comma-separated list of amounts.
Public Function FormatAmountList(values() As Currency) As String
    Dim i As Long, s As String
    For i = LBound(values) To UBound(values)
        If Len(s) > 0 Then s = s & ", "
        s = s & FormatAmount(values(i))
    Next
    FormatAmountList = s
End Function
