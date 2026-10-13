VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Changed(ByVal newValue As Long)

Private m_value As Long

Public Property Get Value() As Long
    Value = m_value
End Property

Public Property Let Value(ByVal v As Long)
    m_value = v
    RaiseEvent Changed(m_value)
End Property

Public Sub Increment()
    Value = m_value + 1
End Sub