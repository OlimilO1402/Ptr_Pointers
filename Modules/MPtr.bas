Attribute VB_Name = "MPtr"
Option Explicit

'In this module everything for working with pointers

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

Public Enum SAFeature
    FADF_AUTO = &H1
    FADF_STATIC = &H2
    FADF_EMBEDDED = &H4

    FADF_FIXEDSIZE = &H10
    FADF_RECORD = &H20
    FADF_HAVEIID = &H40
    FADF_HAVEVARTYPE = &H80
    
    FADF_BSTR = &H100
    FADF_UNKNOWN = &H200
    FADF_DISPATCH = &H400
    FADF_VARIANT = &H800&
    FADF_RESERVED = &HF008
End Enum

'Public Const FADF_RECORD As Long = &H20&

Public Type TUDTPtr
    pSA        As Long
    Reserved   As Long    ' vbVarType / IRecordInfo
    cDims      As Integer
    fFeatures  As Integer ' SAFeature but int16
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements  As Long
    lLBound    As Long
End Type

Public Declare Sub GetMemArr Lib "msvbvm60" Alias "GetMem4" (ByRef Arr() As Any, ByRef Value As Long) 'same as ArrPtr

Public Declare Sub GetMem1 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any)
Public Declare Sub GetMem2 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any)
Public Declare Sub GetMem4 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any)
Public Declare Sub GetMem8 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any)

Public Declare Sub PutMem1 Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Byte)
Public Declare Sub PutMem2 Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Integer)
Public Declare Sub PutMemBol Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Boolean)
Public Declare Sub PutMem4 Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Long)
Public Declare Sub PutMemSng Lib "msvbvm60" Alias "PutMem4" (ByRef Dst As Any, ByVal Src As Single)
Public Declare Sub PutMem8 Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Currency)
Public Declare Sub PutMemDbl Lib "msvbvm60" Alias "PutMem8" (ByRef Dst As Any, ByVal Src As Double)
Public Declare Sub PutMemDat Lib "msvbvm60" Alias "PutMem8" (ByRef Dst As Any, ByVal Src As Date)

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)
Public Declare Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal BytLen As Long)

Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef pArr() As Any) As LongPtr

'here also everything concerning array pointers like VArrPtr and StrArrPtr

'helper function for StringArrays
Public Function StrArrPtr(ByRef strArr As Variant) As LongPtr
    RtlMoveMemory StrArrPtr, ByVal VarPtr(strArr) + 8, 4
End Function

Public Function VArrPtr(ByRef VArr As Variant) As LongPtr
    RtlMoveMemory VArrPtr, ByVal VarPtr(VArr) + 8, 4
End Function

'jetzt kann das Property SAPtr für Alle Arrays verwendet werden,
'um den Zeiger auf den Safe-Array-Descriptor eines Arrays einem
'anderen Array zuzuweisen.
Public Property Get SAPtr(ByVal pArr As LongPtr) As LongPtr
    RtlMoveMemory SAPtr, ByVal pArr, 4
End Property

Public Property Let SAPtr(ByVal pArr As LongPtr, ByVal RHS As LongPtr)
    RtlMoveMemory ByVal pArr, RHS, 4
End Property

Public Sub ZeroSAPtr(ByVal pArr As LongPtr)
    Call RtlZeroMemory(ByVal pArr, 4)
End Sub

Public Sub New_UDTPtr(ByRef this As TUDTPtr, _
                      ByVal Feature As SAFeature, _
                      ByVal bytesPerElement As Long, _
                      Optional ByVal CountElements As Long = 1, _
                      Optional ByVal lLBound As Long = 0)
    
    With this
        .pSA = VarPtr(.cDims) 'nur als Sub wegen VarPtr(cDims)
        .cDims = 1
        .cbElements = bytesPerElement
        .fFeatures = CInt(Feature)
        .cElements = CountElements
        .lLBound = lLBound
    End With
    
End Sub

' checks content of UDTPtr
Public Function UDTPtrToString(this As TUDTPtr) As String
    
    Dim s As String
    
    With this
        s = s & "pSA        : " & CStr(.pSA) & vbCrLf
        s = s & "Reserved   : " & CStr(.Reserved) & vbCrLf
        s = s & "cDims      : " & CStr(.cDims) & vbCrLf
        s = s & "fFeatures  : " & FeaturesToString(CLng(.fFeatures)) & vbCrLf
        s = s & "cbElements : " & CStr(.cbElements) & vbCrLf
        s = s & "cLocks     : " & CStr(.cLocks) & vbCrLf
        s = s & "pvData     : " & CStr(.pvData) & vbCrLf
        s = s & "cElements  : " & CStr(.cElements) & vbCrLf
        s = s & "lLBound    : " & CStr(.lLBound) & vbCrLf
    End With
    
    UDTPtrToString = s
    
End Function


Private Function FeaturesToString(ByVal Feature As SAFeature) As String
    
    Dim s As String
    Dim sOr As String: sOr = " Or "
    
    If Feature And FADF_AUTO Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_AUTO"
    If Feature And FADF_STATIC Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_STATIC"
    
    If Feature And FADF_EMBEDDED Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_EMBEDDED"
    
    If Feature And FADF_FIXEDSIZE Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_FIXEDSIZE"
    
    If Feature And FADF_RECORD Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_RECORD"
    If Feature And FADF_HAVEIID Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_HAVEIID"
    
    If Feature And FADF_HAVEVARTYPE Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_HAVEVARTYPE"
    
    If Feature And FADF_BSTR Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_BSTR"
    If Feature And FADF_UNKNOWN Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_UNKNOWN"
    
    If Feature And FADF_DISPATCH Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_DISPATCH"
    
    If Feature And FADF_VARIANT Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_VARIANT"
    
    If Feature And FADF_RESERVED Then s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_RESERVED"
    
    FeaturesToString = s
    
End Function

