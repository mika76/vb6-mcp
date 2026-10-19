VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' The category codes RateFor knows, comma-separated.
Public Function Codes() As String
    Codes = Plain("GPPE") & "," & Plain("TUE") & "," & Plain("CPPLT")
End Function

' The tax rate for a category code; unknown codes get the standard rate.
Public Function RateFor(ByVal code As String) As Double
    Select Case code
    Case Plain("GPPE"): RateFor = 0.05
    Case Plain("CPPLT"): RateFor = 0
    Case Else: RateFor = 0.2
    End Select
End Function

' Codes are kept shifted by one letter
Private Function Plain(ByVal s As String) As String
    Dim i As Long
    For i = 1 To Len(s)
        Plain = Plain & Chr$(Asc(Mid$(s, i, 1)) - 1)
    Next
End Function
