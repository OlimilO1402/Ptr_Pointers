Attribute VB_Name = "MPtr"
Option Explicit

'In this module you can find everything you need for proper working with pointers, LongPtr, udt-pointer, array, function and collection

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

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
    Public Enum LongLong
        [_]
    End Enum
#End If

Public LongPtr_Empty As LongPtr

#If Win64 Then
    Public Const SizeOf_LongPtr As Long = 8
#Else
    Public Const SizeOf_LongPtr As Long = 4
#End If

Public Type TUDTPtr
    pSA        As LongPtr
    Reserved   As LongPtr ' vbVarType / IRecordInfo
    cDims      As Integer
    fFeatures  As Integer ' SAFeature but int16
    cbElements As Long
    cLocks     As Long
    pvData     As LongPtr
    cElements  As Long
    lLBound    As Long
End Type

#If Win64 Then
    Public Declare PtrSafe Sub GetMemArr Lib "msvbvm60" Alias "GetMem8" (ByRef Arr() As Any, ByRef Value As LongPtr) 'same as ArrPtr
    Public Declare PtrSafe Sub PutMemArr Lib "msvbvm60" Alias "PutMem8" (ByRef Dst As Any, ByVal Src As LongPtr)
#Else
    Public Declare Sub GetMemArr Lib "msvbvm60" Alias "GetMem4" (ByRef Arr() As Any, ByRef Value As LongPtr) 'same as ArrPtr
    Public Declare Sub PutMemArr Lib "msvbvm60" Alias "PutMem4" (ByRef Dst As Any, ByVal src As LongPtr)
#End If

#If VBA7 Then
    Public Declare PtrSafe Sub GetMem1 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any)
    Public Declare PtrSafe Sub GetMem2 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any)
    Public Declare PtrSafe Sub GetMem4 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any)
    Public Declare PtrSafe Sub GetMem8 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any)
    
    Public Declare PtrSafe Sub PutMem1 Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Byte)
    Public Declare PtrSafe Sub PutMem2 Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Integer)
    Public Declare PtrSafe Sub PutMemBol Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Boolean)
    Public Declare PtrSafe Sub PutMem4 Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Long)
    Public Declare PtrSafe Sub PutMemSng Lib "msvbvm60" Alias "PutMem4" (ByRef Dst As Any, ByVal Src As Single)
    Public Declare PtrSafe Sub PutMem8 Lib "msvbvm60" (ByRef Dst As Any, ByVal Src As Currency)
    Public Declare PtrSafe Sub PutMemDbl Lib "msvbvm60" Alias "PutMem8" (ByRef Dst As Any, ByVal Src As Double)
    Public Declare PtrSafe Sub PutMemDat Lib "msvbvm60" Alias "PutMem8" (ByRef Dst As Any, ByVal Src As Date)
    
    Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As LongLong)
    Public Declare PtrSafe Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal BytLen As LongLong)
    
    Public Declare PtrSafe Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef pArr() As Any) As LongPtr
#Else
    Public Declare Sub GetMem1 Lib "msvbvm60" (ByRef src As Any, ByRef Dst As Any)
    Public Declare Sub GetMem2 Lib "msvbvm60" (ByRef src As Any, ByRef Dst As Any)
    Public Declare Sub GetMem4 Lib "msvbvm60" (ByRef src As Any, ByRef Dst As Any)
    Public Declare Sub GetMem8 Lib "msvbvm60" (ByRef src As Any, ByRef Dst As Any)
    
    Public Declare Sub PutMem1 Lib "msvbvm60" (ByRef Dst As Any, ByVal src As Byte)
    Public Declare Sub PutMem2 Lib "msvbvm60" (ByRef Dst As Any, ByVal src As Integer)
    Public Declare Sub PutMemBol Lib "msvbvm60" (ByRef Dst As Any, ByVal src As Boolean)
    Public Declare Sub PutMem4 Lib "msvbvm60" (ByRef Dst As Any, ByVal src As Long)
    Public Declare Sub PutMemSng Lib "msvbvm60" Alias "PutMem4" (ByRef Dst As Any, ByVal src As Single)
    Public Declare Sub PutMem8 Lib "msvbvm60" (ByRef Dst As Any, ByVal src As Currency)
    Public Declare Sub PutMemDbl Lib "msvbvm60" Alias "PutMem8" (ByRef Dst As Any, ByVal src As Double)
    Public Declare Sub PutMemDat Lib "msvbvm60" Alias "PutMem8" (ByRef Dst As Any, ByVal src As Date)
    
    Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As LongLong)
    Public Declare Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal BytLen As LongLong)
    
    Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef pArr() As Any) As LongPtr
#End If

'here everything concerning array pointers like VArrPtr and StrArrPtr

'1. first use SAPtr, or StrArrPtr or VArrPtr to get the pointer to the safe-array-descriptor
'helper function for StringArrays
Public Function StrArrPtr(ByRef strArr As Variant) As LongPtr
'Attention, here 32bit-64bit-trap, so use only RtlMoveMemory to be variable in size of ptr
    RtlMoveMemory StrArrPtr, ByVal VarPtr(strArr) + 8, MPtr.SizeOf_LongPtr
End Function
Public Function VArrPtr(ByRef VArr As Variant) As LongPtr
    RtlMoveMemory VArrPtr, ByVal VarPtr(VArr) + 8, MPtr.SizeOf_LongPtr
End Function

'2. now you are able to use the Property SAPtr for all arrays, for assigning
'   the pointer to a safe-array-descriptor to another array.
Public Property Get SAPtr(ByVal pArr As LongPtr) As LongPtr
    RtlMoveMemory SAPtr, ByVal pArr, MPtr.SizeOf_LongPtr
End Property
Public Property Let SAPtr(ByVal pArr As LongPtr, ByVal RHS As LongPtr)
    RtlMoveMemory ByVal pArr, RHS, MPtr.SizeOf_LongPtr
End Property

'3. don't forget to delete the pointer before VB tries to do it.
Public Sub ZeroSAPtr(ByVal pArr As LongPtr)
    RtlZeroMemory ByVal pArr, MPtr.SizeOf_LongPtr
End Sub

Public Function FncPtr(ByVal pfn As Long) As Long
    FncPtr = pfn
End Function

Public Function Col_Contains(col As Collection, key As String) As Boolean
    'for this Function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original version of this function here: https://vb-tec.de/collctns.htm
    On Error Resume Next
    If IsEmpty(col(key)) Then: 'DoNothing
    Col_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function
    
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
    
    If Feature And FADF_AUTO Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_AUTO"
    If Feature And FADF_STATIC Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_STATIC"
    If Feature And FADF_EMBEDDED Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_EMBEDDED"
    If Feature And FADF_FIXEDSIZE Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_FIXEDSIZE"
    If Feature And FADF_RECORD Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_RECORD"
    If Feature And FADF_HAVEIID Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_HAVEIID"
    If Feature And FADF_HAVEVARTYPE Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_HAVEVARTYPE"
    If Feature And FADF_BSTR Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_BSTR"
    If Feature And FADF_UNKNOWN Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_UNKNOWN"
    If Feature And FADF_DISPATCH Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_DISPATCH"
    If Feature And FADF_VARIANT Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_VARIANT"
    If Feature And FADF_RESERVED Then _
                                        s = s & IIf(Len(s), sOr, vbNullString) & "FADF_RESERVED"
    FeaturesToString = s
    
End Function





