VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Test2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Text As String

Private Sub Class_Initialize()
    m_Text = "Class Test2"
End Sub

Public Function Test2() As String
    Test2 = m_Text & "." & "Test2"
End Function

Public Function Test4() As String
    Test4 = m_Text & "." & "Test4"
End Function
