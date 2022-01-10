Attribute VB_Name = "ModUDTPtr"
Option Explicit
'Ein SafeArray-Descriptor dient in VB als ein universaler Zeiger
Public Type TUDTPtr
    pSA        As Long
    Reserved   As Long 'z.B. für vbVarType oder IRecordInfo
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements  As Long
    lLBound    As Long
End Type

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
    FADF_VARIANT = &H800
    FADF_RESERVED = &HF008
End Enum

Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
    ByRef pDst As Any, ByRef pSrc As Any, ByVal bLength As Long)
Public Declare Sub RtlZeroMemory Lib "kernel32" ( _
    ByRef pDst As Any, ByVal bLength As Long)
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" ( _
    ByRef pArr() As Any) As Long

Public Sub New_UDTPtr(ByRef this As TUDTPtr, _
                      ByVal Feature As SAFeature, _
                      ByVal bytesPerElement As Long, _
                      Optional ByVal CountElements As Long = 1, _
                      Optional ByVal lLBound As Long = 0)
                           
    With this
        .pSA = VarPtr(.cDims)
        .cDims = 1
        .cbElements = bytesPerElement
        .fFeatures = CInt(Feature)
        .cElements = CountElements
        .lLBound = lLBound
    End With
    Debug.Print UDTPtrToString(this)
End Sub

'Um zu überprüfen ob der UDTPtr auch das enthält was er soll
'kann man folgende Funktion verwenden
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
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_AUTO"
    If Feature And FADF_STATIC Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_STATIC"
    If Feature And FADF_EMBEDDED Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_EMBEDDED"
  
    If Feature And FADF_FIXEDSIZE Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_FIXEDSIZE"
    If Feature And FADF_RECORD Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_RECORD"
    If Feature And FADF_HAVEIID Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_HAVEIID"
    If Feature And FADF_HAVEVARTYPE Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_HAVEVARTYPE"
  
    If Feature And FADF_BSTR Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_BSTR"
    If Feature And FADF_UNKNOWN Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_UNKNOWN"
    If Feature And FADF_DISPATCH Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_DISPATCH"
    If Feature And FADF_VARIANT Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_VARIANT"
    If Feature And FADF_RESERVED Then _
        s = s & IIf(Len(s), sOr, vbNullString): s = s & "FADF_RESERVED"
    
    FeaturesToString = s
End Function
