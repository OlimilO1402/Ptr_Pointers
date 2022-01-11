Attribute VB_Name = "MByteSwapper"
'---------------------------------------------------------------------------------------
' Module    : MByteSwapper
' DateTime  : 11.10.2008 19:40
' Author    : Oliver Meyer (olimilo AT gmx DOT net)
' Purpose   : Tauscht die Bytereihenfolge von 16, 32 und 64 Bit Typen
'---------------------------------------------------------------------------------------
Option Explicit
'Modul MByteSwapper
' Ein SafeArray-Descriptor dient in VB als ein universaler Zeiger
'Public Type TUDTPtr
'    pSA        As Long
'    Reserved   As Long ' z.B. für vbVarType oder IRecordInfo
'    cDims      As Integer
'    fFeatures  As Integer
'    cbElements As Long
'    cLocks     As Long
'    pvData     As Long
'    cElements  As Long
'    lLBound    As Long
'End Type
'
'Public Enum SAFeature
'    FADF_AUTO = &H1
'    FADF_STATIC = &H2
'    FADF_EMBEDDED = &H4
'
'    FADF_FIXEDSIZE = &H10
'    FADF_RECORD = &H20
'    FADF_HAVEIID = &H40
'    FADF_HAVEVARTYPE = &H80
'
'    FADF_BSTR = &H100
'    FADF_UNKNOWN = &H200
'    FADF_DISPATCH = &H400
'    FADF_VARIANT = &H800
'    FADF_RESERVED = &HF008
'End Enum

Public Type TByteSwapper
    pB      As TUDTPtr
    tmpByte As Byte
    b()     As Byte
End Type

'Private Declare Sub PutMem4 Lib "msvbvm60" ( _
'    ByRef pDst As Any, _
'    ByVal Src As Long)
'
'Public Declare Function ArrPtr Lib "msvbvm60" _
'                        Alias "VarPtr" ( _
'                        ByRef pArr() As Any) As Long
'
'Public Sub New_UDTPtr(ByRef this As TUDTPtr, _
'                      ByVal Feature As SAFeature, _
'                      ByVal bytesPerElement As Long, _
'                      Optional ByVal CountElements As Long = 1, _
'                      Optional ByVal lLBound As Long = 0)
'
'    With this
'        .pSA = VarPtr(.cDims) 'nur als Sub wegen VarPtr(cDims)
'        .cDims = 1
'        .cbElements = bytesPerElement
'        .fFeatures = CInt(Feature)
'        .cElements = CountElements
'        .lLBound = lLBound
'    End With
'
'End Sub
'
Public Sub New_ByteSwapper(this As TByteSwapper, Optional ByVal CountBytes As Long = 2)
    With this
        Call New_UDTPtr(.pB, FADF_EMBEDDED Or FADF_STATIC, 1, CountBytes)
        Call PutMem4(ByVal ArrPtr(.b), .pB.pSA)
    End With
End Sub

Public Sub DeleteByteSwapper(this As TByteSwapper)
    With this
        Call PutMem4(ByVal ArrPtr(.b), 0)
    End With
End Sub

Public Sub Rotate2(this As TByteSwapper)
    With this
        .tmpByte = .b(0)
        .b(0) = .b(1)
        .b(1) = .tmpByte
    End With
End Sub
Public Sub Rotate4(this As TByteSwapper)
    With this
        .tmpByte = .b(0)
        .b(0) = .b(3)
        .b(3) = .tmpByte
        
        .tmpByte = .b(1)
        .b(1) = .b(2)
        .b(2) = .tmpByte
    End With
End Sub
Public Sub Rotate8(this As TByteSwapper)
    With this
        .tmpByte = .b(0)
        .b(0) = .b(7)
        .b(7) = .tmpByte
        
        .tmpByte = .b(1)
        .b(1) = .b(6)
        .b(6) = .tmpByte
        
        .tmpByte = .b(2)
        .b(2) = .b(5)
        .b(5) = .tmpByte
        
        .tmpByte = .b(3)
        .b(3) = .b(4)
        .b(4) = .tmpByte
    End With
End Sub

Public Sub Rotate(this As TByteSwapper)
    With this
        Select Case .pB.cElements
        Case 2
            .tmpByte = .b(0)
            .b(0) = .b(1)
            .b(1) = .tmpByte
        Case 4
            .tmpByte = .b(0)
            .b(0) = .b(3)
            .b(3) = .tmpByte
            
            .tmpByte = .b(1)
            .b(1) = .b(2)
            .b(2) = .tmpByte
        Case 8
            .tmpByte = .b(0)
            .b(0) = .b(7)
            .b(7) = .tmpByte
            
            .tmpByte = .b(1)
            .b(1) = .b(6)
            .b(6) = .tmpByte
            
            .tmpByte = .b(2)
            .b(2) = .b(5)
            .b(5) = .tmpByte
            
            .tmpByte = .b(3)
            .b(3) = .b(4)
            .b(4) = .tmpByte
        End Select
    End With
End Sub


