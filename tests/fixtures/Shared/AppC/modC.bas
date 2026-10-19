Attribute VB_Name = "modC"
' AppC doesn't use Common\modFormat.bas: this FormatAmount is its own
Option Explicit

Private Function FormatAmount(ByVal value As Currency) As String
    FormatAmount = "$" & Format$(value, "0")
End Function

Sub Main()
    Debug.Print FormatAmount(42)
End Sub
