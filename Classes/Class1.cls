VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'just a dummy-class
Private m_Value As Double

Friend Sub New_(ByVal Value As Double)
    m_Value = Value
End Sub

Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Get Ptr() As LongPtr
    Ptr = ObjPtr(Me)
End Property

Public Property Get Key() As String
    Key = Ptr
End Property

Public Function ToStr() As String
    ToStr = Format(m_Value, "###,###,###,##0.000")
End Function

Public Function Compare(other As Class1) As Long
    Compare = Me.Value - other.Value
End Function
