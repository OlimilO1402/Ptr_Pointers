VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Test1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Text As String

Private Sub Class_Initialize()
    m_Text = "Class Test1"
End Sub

Public Function Test1() As String
    Test1 = m_Text & "." & "Test1"
End Function

Public Function Test3() As String
    Test3 = m_Text & "." & "Test3"
End Function
