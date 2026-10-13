Attribute VB_Name = "modMath"
Option Explicit

Public Const PI_APPROX As Double = 3.14159
Public g_callCount As Long

' Adds two numbers and counts the calls.
Public Function AddNumbers(ByVal a As Long, ByVal b As Long) As Long
    g_callCount = g_callCount + 1
    AddNumbers = a + b
End Function

Private Sub Unused()
End Sub