Attribute VB_Name = "modCalc"
' Debug fixture (Phase 2): a call chain with locals of several types, and a runtime error.
Option Explicit

Public g_runs As Long

Public Function Outer(ByVal n As Long) As Long
    Dim total As Long, label As String
    label = "outer"
    g_runs = g_runs + 1
    total = Inner(n, label) + 1
    Outer = total
End Function

Private Function Inner(ByVal n As Long, ByRef tag As String) As Long
    Dim i As Long, acc As Long, items(2) As Variant, flag As Boolean, when As Date
    items(0) = "a": items(1) = 2: items(2) = Null
    when = #1/2/2003#
    For i = 1 To n
        acc = acc + i
    Next
    flag = acc > 5
    tag = tag & "/inner"
    Inner = acc
End Function

Public Function Divide(ByVal a As Long, ByVal b As Long) As Long
    Divide = a \ b
End Function