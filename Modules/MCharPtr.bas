Attribute VB_Name = "MCharPtr"
Option Explicit
'Public Type TCharPointer
'    pudt    As TUDTPtr
'    Chars() As Integer
'End Type
'
'Public Sub New_CharPointer(ByRef this As TCharPointer, ByRef StrVal As String)
'    With this
'        New_UDTPtr .pudt, FADF_AUTO Or FADF_FIXEDSIZE, 2, Len(StrVal), 1
'        With .pudt
'            .pvData = StrPtr(StrVal)
'        End With
'        RtlMoveMemory ByVal ArrPtr(.Chars), ByVal VarPtr(.pudt), 4
'    End With
'End Sub
'
'Public Sub DeleteCharPointer(ByRef this As TCharPointer)
'    With this
'        RtlZeroMemory ByVal ArrPtr(.Chars), 4
'    End With
'End Sub
