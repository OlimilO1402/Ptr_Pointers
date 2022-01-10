Attribute VB_Name = "MArray"
Option Explicit

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef Src As Any, ByVal BytLen As Long)

Public Declare Sub RtlZeroMemory Lib "kernel32" (ByRef Dst As Any, ByVal BytLen As Long)

'die Funktion ArrPtr geht bei allen Arrays außer bei String-Arrays
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef Arr() As Any) As Long

'deswegen hier eine Hilfsfunktion für StringArrays
Public Function StrArrPtr(ByRef strArr As Variant) As Long
    RtlMoveMemory StrArrPtr, ByVal VarPtr(strArr) + 8, 4
End Function

'jetzt kann das Property SAPtr für Alle Arrays verwendet werden,
'um den Zeiger auf den Safe-Array-Descriptor eines Arrays einem
'anderen Array zuzuweisen.
Public Property Get SAPtr(ByVal pArr As Long) As Long
    RtlMoveMemory SAPtr, ByVal pArr, 4
End Property

Public Property Let SAPtr(ByVal pArr As Long, ByVal RHS As Long)
    RtlMoveMemory ByVal pArr, RHS, 4
End Property

Public Sub ZeroSAPtr(ByVal pArr As Long)
    RtlZeroMemory ByVal pArr, 4
End Sub

'####################'   Class1 Helper   '####################'
Public Function New_Class1(ByVal dVal As Double) As Class1
    Set New_Class1 = New Class1
    New_Class1.DblVal = dVal
End Function
