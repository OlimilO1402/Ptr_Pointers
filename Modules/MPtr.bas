Attribute VB_Name = "MPtr"
Option Explicit

'In this module you can find everything you need concerning pointers in general and for proper working with it, LongPtr, safearray, udt-pointer, byte-swapper, array-ptr, function and collection

'this Enum will be used in the SafeArray-descriptor
Public Enum SAFeature
    FADF_AUTO = &H1         ' An array that is allocated on the stack.
    FADF_STATIC = &H2       ' An array that is statically allocated.
    FADF_EMBEDDED = &H4     ' An array that is embedded in a structure.
    
    FADF_FIXEDSIZE = &H10   ' An array that may not be resized or reallocated.
    FADF_RECORD = &H20      ' An array that contains records. When set, there will be a pointer to the IRecordInfo interface at negative offset 4 in the array descriptor.
    FADF_HAVEIID = &H40     ' An array that has an IID identifying interface. When set, there will be a GUID at negative offset 16 in the safe array descriptor. Flag is set only when FADF_DISPATCH or FADF_UNKNOWN is also set.
    FADF_HAVEVARTYPE = &H80 ' An array that has a variant type. The variant type can be retrieved with SafeArrayGetVartype.
    
    FADF_BSTR = &H100       ' An array of BSTRs.
    FADF_UNKNOWN = &H200    ' An array of IUnknown*.
    FADF_DISPATCH = &H400   ' An array of IDispatch*.
    FADF_VARIANT = &H800&   ' An array of VARIANTs.
    FADF_RESERVED = &HF008  ' Bits reserved for future use.
End Enum

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If
'#If Win64 = 0 Then
'    Public Enum LongLong
'        [_]
'    End Enum
'#End If

Public LongPtr_Empty As LongPtr

#If win64 Then
    Public Const SizeOf_LongPtr As Long = 8
    Public Const SizeOf_Variant As Long = 20
#Else
    Public Const SizeOf_LongPtr As Long = 4
    Public Const SizeOf_Variant As Long = 16
#End If

' https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-safearray
' https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-safearraybound
' a SafeArray-descriptor serves perfectly as a universal pointer
Public Type TUDTPtr
    pSA        As LongPtr ' pointer to cDims
    Reserved   As LongPtr ' vbVarType / IRecordInfo
    cDims      As Integer ' The number of Dimensions
    fFeatures  As Integer ' Flags SAFeature but int16
    cbElements As Long    ' The size of an array element.
    cLocks     As Long    ' The number of times the array has been locked without a corresponding unlock.
    pvData     As LongPtr ' The data
    cElements  As Long    ' The number of elements in the dimension.
    lLBound    As Long    ' The lower bound of the dimension.
End Type

' the TSafeArrayPtr lightweight-object structure
Public Type TSafeArrayPtr
    pSAPtr As TUDTPtr
    pSA()  As TUDTPtr
End Type

Public Type VBGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public GUID_NULL As VBGUID

#If win64 Then
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
    #If win64 Then
        Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As LongLong)
        Public Declare PtrSafe Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal BytLen As LongLong)
        Public Declare PtrSafe Sub RtlFillMemory Lib "kernel32" (ByRef pDst As Any, ByVal BytLen As LongLong)
        'https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlcomparememory
        Public Declare PtrSafe Function RtlCompareMemory Lib "kernel32" (ByRef pSrc0 As Any, ByRef pSrc1 As Any, ByVal BytLen As LongLong) as Long
    #Else
        Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)
        Public Declare PtrSafe Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal BytLen As Long)
        Public Declare PtrSafe Sub RtlFillMemory Lib "kernel32" (ByRef pDst As Any, ByVal BytLen As Long)
        'https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlcomparememory
        Public Declare PtrSafe Function RtlCompareMemory Lib "kernel32" (ByRef pSrc0 As Any, ByRef pSrc1 As Any, ByVal BytLen As Long) as Long
    #End If
    Public Declare PtrSafe Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef pArr() As Any) As LongPtr
#Else
    'GetMem and PutMem are copying memory just like RtlMoveMemory but only for a certain amount of bytes
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
    
    Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)
    Public Declare Sub RtlZeroMemory Lib "kernel32" (ByRef pDst As Any, ByVal BytLen As Long)
    Public Declare Sub RtlFillMemory Lib "kernel32" (ByRef pDst As Any, ByVal BytLen As Long)
    
    'https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlcomparememory
    Public Declare Function RtlCompareMemory Lib "ntdll" (ByRef pSrc1 As Any, ByRef pSrc2 As Any, ByVal BytLen As Long) As Long ' NTSYSAPI
    'RtlCompareMemory returns the number of bytes in the two blocks that match. If all bytes match up to the specified Length value, the Length value is returned.
    
    Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef pArr() As Any) As LongPtr
#End If

' v ############################## v '    MByteSwapper Declarations    ' v ############################## v '
'
'now the question
'should we maybe integrate MByteSwapper completely into MPtr?
' * on the one hand I am the type of programmer who puts/organizes different things in different modules, then everything is in a pretty order, separated from each other.
'   now the question, is SwapByteOrder something different than pointers, on the one hand, yes it is, it just uses heavily intensively pointers and the UDTPtr-method.
' * on the other hand, again one more module . . . Unicode will be used by MString, PathFileName uses it and this are classes used and contained in nearly every project,
'   so there would be no need for even upgrading a single project containing classes PathFilename or MString.
'OK we give it a try we are copying it into MPtr and separating it with special lines.

' the TByteSwapper lightweight-object structure
Public Type TByteSwapper
    pB      As TUDTPtr
    tmpByte As Byte
    b()     As Byte
End Type

#If VBA7 Then
    
    'The following functions can be found in the dll SwapByteOrder, written in assembler, and are much faster.
    Public Declare PtrSafe Sub SBO_Rotate2 Lib "SwapByteOrder" Alias "SwapByteOrder16" (ByRef Ptr As Any)
    Public Declare PtrSafe Sub SBO_Rotate4 Lib "SwapByteOrder" Alias "SwapByteOrder32" (ByRef Ptr As Any)
    Public Declare PtrSafe Sub SBO_Rotate8 Lib "SwapByteOrder" Alias "SwapByteOrder64" (ByRef Ptr As Any)
    Public Declare PtrSafe Function SBO_RotateArray Lib "SwapByteOrder" Alias "SwapByteOrderArray" (ByRef Value() As Any) As Long
    Public Declare PtrSafe Sub SBO_RotateUDTArray Lib "SwapByteOrder" Alias "SwapByteOrderUDTArray" (ByRef Arr() As Any, ByRef udtDescription() As Integer)
    
#Else
    
    'The following functions can be found in the dll SwapByteOrder, written in assembler, and are much faster.
    Public Declare Sub SBO_Rotate2 Lib "SwapByteOrder" Alias "SwapByteOrder16" (ByRef Ptr As Any)
    Public Declare Sub SBO_Rotate4 Lib "SwapByteOrder" Alias "SwapByteOrder32" (ByRef Ptr As Any)
    Public Declare Sub SBO_Rotate8 Lib "SwapByteOrder" Alias "SwapByteOrder64" (ByRef Ptr As Any)
    Public Declare Function SBO_RotateArray Lib "SwapByteOrder" Alias "SwapByteOrderArray" (ByRef Value() As Any) As Long
    Public Declare Sub SBO_RotateUDTArray Lib "SwapByteOrder" Alias "SwapByteOrderUDTArray" (ByRef Arr() As Any, ByRef udtDescription() As Integer)
    
#End If

' ^ ############################## ^ '    MByteSwapper Declarations    ' ^ ############################## ^ '



' v ############################## v '    Array-Ptr Functions   ' v ############################## v '

'1. first use ArrPtr, or StrArrPtr or VArrPtr to get the pointer to the safe-array-descriptor
'   from the array-variable, when it has a dimension, otherwise the pointer is 0
'   helper function for StringArrays
Public Function StrArrPtr(ByRef strArr As Variant) As LongPtr
'Attention, here 32bit-64bit-trap, so use only RtlMoveMemory to be variable in size of ptr
    RtlMoveMemory StrArrPtr, ByVal VarPtr(strArr) + 8, MPtr.SizeOf_LongPtr
End Function
Public Function VArrPtr(ByRef vArr As Variant) As LongPtr
    RtlMoveMemory VArrPtr, ByVal VarPtr(vArr) + 8, MPtr.SizeOf_LongPtr
End Function

Public Property Get VarSAPtr(ByRef vArr As Variant) As LongPtr
    '        VarSAPtr =
    PutMem4 VarSAPtr, VarPtr(vArr) + 8
    'should be the same as VArrPtr, shouldn't it?
End Property

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
' ^ ############################## ^ '    Array-Ptr Functions    ' ^ ############################## ^ '

' v ############################## v '    Array Functions   ' v ############################## v '

Public Function Array_Count(Arr, Optional nDim As Long = 1) As Long
Try: On Error GoTo Catch
    Array_Count = UBound(Arr, nDim) - LBound(Arr, nDim)
    Exit Function
Catch:
    Array_Count = 0
End Function

' ^ ############################## ^ '    Array Functions    ' ^ ############################## ^ '

'retrieve the pointer to a function by using FncPtr(Addressof myfunction)
Public Function FncPtr(ByVal PFN As LongPtr) As LongPtr
    FncPtr = PFN
End Function

' v ############################## v '    Collection Functions    ' v ############################## v '
Public Function Col_Contains(col As Collection, Key As String) As Boolean
    'for this Function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original version of this function here: https://vb-tec.de/collctns.htm
    On Error Resume Next
'  '"Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(col(Key)) Then: 'DoNothing
    Col_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Col_TryAddObject(col As Collection, Obj As Object, Key As String) As Boolean
Try: On Error GoTo Catch
    col.Add Obj, Key
    Col_TryAddObject = True
Catch: On Error GoTo 0
End Function

Public Sub Col_SwapItems(col As Collection, ByVal i1 As Long, i2 As Long)
    Dim c As Long: c = col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or col.Count <= i1 Then Exit Sub
    If i2 <= 0 Or col.Count < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim Obj1, Obj2
    If IsObject(col.Item(i1)) Then Set Obj1 = col.Item(i1) Else Obj1 = col.Item(i1)
    If IsObject(col.Item(i2)) Then Set Obj2 = col.Item(i2) Else Obj2 = col.Item(i2)
    col.Remove i1: col.Add Obj2, , i1:     col.Remove i2
    If i2 < c Then col.Add Obj1, , i2 Else col.Add Obj1
End Sub

Public Sub Col_MoveUp(col As Collection, ByVal i As Long)
    Col_SwapItems col, i, i - 1
End Sub

Public Sub Col_MoveDown(col As Collection, ByVal i As Long)
    Col_SwapItems col, i, i + 1
End Sub

' ^ ############################## ^ '    Collection Functions    ' ^ ############################## ^ '




' v ############################## v '    UDTPtr Functions   ' v ############################## v '

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

Public Sub UDTPtr_Assign(pDst As TUDTPtr, pSrc As TUDTPtr)
'hier wird nicht einfach nur zugewiesen in der Art pDst = pSrc
'sondern hier wird nur pvdata und cElements zugewiesen, wobei
'cElements in Abhängigkeit von cbElement entsprechend angepasst wird
    pDst.pvData = pSrc.pvData
    If pDst.cbElements > 0 Then
        pDst.cElements = pSrc.cElements * pSrc.cbElements \ pDst.cbElements + 1
    End If
End Sub

' Checks content of UDTPtr
Public Function UDTPtr_ToStr(this As TUDTPtr) As String
    
    Dim s As String
    Dim saf As SAFeature
    With this
        s = s & "pSA        : " & CStr(.pSA) & vbCrLf
        s = s & "cDims      : " & CStr(.cDims) & vbCrLf
        
        saf = .fFeatures
        If saf And FADF_HAVEVARTYPE Then
            s = s & "VarType    : " & VBVarType_ToStr(.Reserved) & vbCrLf
        ElseIf saf And FADF_DISPATCH Then
            s = s & "VarType    : " & VBVarType_ToStr(VbVarType.vbObject) & vbCrLf
            s = s & "pVTable    : " & .Reserved & vbCrLf
        Else
            s = s & "Reserved   : " & .Reserved & vbCrLf
        End If
        s = s & "fFeatures  : " & FeaturesToString(CLng(.fFeatures)) & vbCrLf
        s = s & "cbElements : " & CStr(.cbElements) & vbCrLf
        s = s & "cLocks     : " & CStr(.cLocks) & vbCrLf
        s = s & "pvData     : " & CStr(.pvData) & vbCrLf
        s = s & "cElements  : " & CStr(.cElements) & vbCrLf
        s = s & "lLBound    : " & CStr(.lLBound) & vbCrLf
    End With
    
    UDTPtr_ToStr = s
    
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

' ^ ############################## ^ '    UDTPtr Functions    ' ^ ############################## ^ '


' v ############################## v '  Object-WeakPtr Funcs  ' v ############################## v '

Public Function PtrToObject(ByVal p As LongPtr) As Object
    Dim Obj As IUnknown:  RtlMoveMemory Obj, p, MPtr.SizeOf_LongPtr
    Set PtrToObject = Obj: ZeroObject Obj
End Function

Public Sub ZeroObject(Obj As Object)
    'RtlZeroMemory ByVal VarPtr(obj), MPtr.SizeOf_LongPtr
    RtlZeroMemory Obj, MPtr.SizeOf_LongPtr
End Sub

Public Sub AssignSwap(Obj1 As IUnknown, Obj2 As IUnknown)
    Dim pObj1 As LongPtr: pObj1 = ObjPtr(Obj1)
    Dim pObj2 As LongPtr: pObj2 = ObjPtr(Obj2)
    RtlMoveMemory Obj1, pObj2, MPtr.SizeOf_LongPtr
    RtlMoveMemory Obj2, pObj1, MPtr.SizeOf_LongPtr
End Sub

Public Function VTablePtr(Obj As Object) As LongPtr
    RtlMoveMemory VTablePtr, ByVal ObjPtr(Obj), SizeOf_LongPtr
End Function

Public Property Get ObjectAddressOf(Obj As Object, ByVal Index As Long) As LongPtr
    Dim pVTable As LongPtr: pVTable = VTablePtr(Obj) 'first DeRef
    RtlMoveMemory ObjectAddressOf, ByVal pVTable + (7 + Index) * SizeOf_LongPtr, SizeOf_LongPtr
End Property

Public Property Let ObjectAddressOf(Obj As Object, ByVal Index As Long, ByVal Value As LongPtr)
    Dim pVTable As LongPtr: pVTable = VTablePtr(Obj) 'first DeRef
    RtlMoveMemory ByVal pVTable + (7 + Index) * SizeOf_LongPtr, Value, SizeOf_LongPtr
End Property

' ^ ############################## ^ '  Object-WeakPtr Funcs  ' ^ ############################## ^ '


' v ############################## v '    SafeArrayPtr Functions   ' v ############################## v '

Public Sub New_SafeArrayPtr(this As TSafeArrayPtr)
    ' creates a new SafeArrayPtr-lightweight-object
    ' works only as a Sub (with ByRef this) because of VarPtr(cDims)
    With this
        New_UDTPtr .pSAPtr, SAFeature.FADF_EMBEDDED Or SAFeature.FADF_STATIC Or SAFeature.FADF_RECORD, LenB(.pSAPtr)
        SAPtr(ArrPtr(.pSA)) = .pSAPtr.pSA
    End With
End Sub

Public Sub SafeArrayPtr_Delete(this As TSafeArrayPtr)
    ' deletes a TSafeArrayPtr-lightweight-object
    ZeroSAPtr ByVal ArrPtr(this.pSA)
End Sub

Public Property Let SafeArrayPtr_SAPtr(this As TSafeArrayPtr, ByVal Value As LongPtr)
    ' writes the pointer to a SafeArrayDescriptor-structure into a
    ' TSafeArrayPtr-lightweight-object
    Dim p As LongPtr
    'GetMem4 ByVal Value, p
    RtlMoveMemory p, ByVal Value, SizeOf_LongPtr
    
    this.pSAPtr.pvData = p - 2 * SizeOf_LongPtr
    ' -8 ist Mist
    ' -8 weil zuerst pSA und Reserved und dann kommt erst
    ' der Anfang der SafeArrayDesc-Struktur mit cDims
End Property

Public Function SafeArrayPtr_ToStr(this As TSafeArrayPtr) As String
    SafeArrayPtr_ToStr = MPtr.UDTPtr_ToStr(this.pSA(0))
End Function

Private Function VBVarType_ToStr(vt As VbVarType) As String
    Dim s As String
    Select Case vt
    Case VbVarType.vbByte:       s = "Byte"
    Case VbVarType.vbInteger:    s = "Integer"
    Case VbVarType.vbLong:       s = "Long"
    Case VbVarType.vbSingle:     s = "Single"
    Case VbVarType.vbDouble:     s = "Double"
    Case VbVarType.vbDate:       s = "Date"
    Case VbVarType.vbString:     s = "String"
    Case VbVarType.vbCurrency:   s = "Currency"
    Case VbVarType.vbDataObject: s = "DataObject"
    Case VbVarType.vbDecimal:    s = "Decimal"
    Case VbVarType.vbObject:     s = "Object"
    End Select
    VBVarType_ToStr = s
End Function

' ^ ############################## ^ '    SafeArrayPtr Functions    ' ^ ############################## ^ '


' v ############################## v '         Math Functions       ' v ############################## v '
Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function
' ^ ############################## ^ '         Math Functions       ' ^ ############################## ^ '

' v ############################## v '    MByteSwapper Functions   ' v ############################## v '

Public Sub New_ByteSwapper(this As TByteSwapper, Optional ByVal CountBytes As Long = 2)
    ' creates a new TByteSwapper lightweight-object
    With this
        New_UDTPtr .pB, FADF_EMBEDDED Or FADF_STATIC, 1, CountBytes
        'Call PutMem4(ByVal ArrPtr(.b), .pB.pSA)
        SAPtr(ArrPtr(.b)) = .pB.pSA
    End With
End Sub

'Public Sub New_ByteSwapperS(this As TByteSwapper, Value As String)
'
'End Sub

Public Sub ByteSwapper_Delete(this As TByteSwapper)
    ' deletes the pointer in the array of the TByteSwapper-structurr
    ZeroSAPtr ByVal ArrPtr(this.b)
End Sub

Public Sub Rotate2(this As TByteSwapper)
    ' swaps 2 bytes
    With this
        .tmpByte = .b(0)
        .b(0) = .b(1)
        .b(1) = .tmpByte
    End With
End Sub
Public Sub Rotate4(this As TByteSwapper)
    ' swaps 4 bytes
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
    ' swaps 8 bytes
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
    ' rotates the bytes about the size of the datatype found in cElements
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

Public Sub RotateArray(this As TByteSwapper, vArr)
    ' rotates the bytes of the elements of an arbitrary array in the Variant vArr.
    ' the Array may be of type Integer, Long, Currency, Single or Double.
    ' if you want to use an Array of UD-Type-elementes instead, use the function RotateUDTArray (see below).
    If Not IsArray(vArr) Then Exit Sub
    Dim pSA As TSafeArrayPtr: Call New_SafeArrayPtr(pSA)
    'SAPtr(pSA) = VarSAPtr(VArr)
    SafeArrayPtr_SAPtr(pSA) = VarSAPtr(vArr) ' = VArrPtr(vArr)
    Dim i  As Long
    Dim p  As Long
    Dim pc As Long
    Dim ub As Long: ub = UBound(vArr)
    Dim lb As Long: lb = LBound(vArr)
    Dim cnt As Long: cnt = ub - lb + 1
    
    'Debug.Print "TByteSwapper.pB(=TUDTPtr)" & vbCrLf & UDTPtr_ToStr(this.pB)
    
    'Debug.Print "SafeArrayPtr.pSAPtr:     " & vbCrLf & UDTPtr_ToStr(pSA.pSAPtr)
    
    If cnt > 0 Then
        With this
            If .pB.pvData = 0 Then
                .pB.pvData = pSA.pSA(0).pvData 'VarPtr(vArr(0))
            End If
            .pB.cElements = pSA.pSA(0).cbElements 'LenB(vArr(0))
            Select Case .pB.cbElements
            Case 2
                For i = lb To ub
                    .pB.pvData = .pB.pvData + 2
                    .tmpByte = .b(0)
                    .b(0) = .b(1)
                    .b(1) = .tmpByte
                Next
            Case 4
                For i = lb To ub
                    .pB.pvData = .pB.pvData + 4
                    .tmpByte = .b(0)
                    .b(0) = .b(3)
                    .b(3) = .tmpByte
                    
                    .tmpByte = .b(1)
                    .b(1) = .b(2)
                    .b(2) = .tmpByte
                Next
            Case 8
                For i = lb To ub
                    .pB.pvData = .pB.pvData + 8
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
                Next
            End Select
        End With
    End If
    SafeArrayPtr_Delete pSA
End Sub

'Eine entsprechende Funktion in der SwapByteOrderDll könnte in etwa so aussehen
'wie diese Funktion
'allerdings statt pData und Count könnte das Array direkt angegeben werden,
'As Any machts möglich.
'eine entsprechende Deklaration könnte so aussehen:
'Public Declare Sub RotateUDTArr Lib "SwapByteOrder.dll" _
'             Alias "SwapByteOrderUDTArray" ( _
'             ByRef ArrayOfUDType() As Any, ByRef udtDescription() As Integer)

Public Sub RotateUDTArray(this As TByteSwapper, _
                          ByVal pData As Long, _
                          ByVal Count As Long, _
                          ByRef udtDescription() As Integer)
    ' Rotiert die Elemente eines Array vom Typ eines beliebigen UD-Types
    ' this:  der ByteSwapper
    ' pData: der Zeiger auf das erste Element im Array (verwende VarPtr())
    ' Count: die Anzahl der Elemente im Array
    ' udtDescription(): liefert eine Beschreibung des UD-Types.
    '                   Der Wert der Integer-Elemente im Array repräsentiert
    '                   die Größe der einzelnen Variablen-Elemente des UD-Types
    '                   verwende dazu die Funktion LenB.
    '                   Variablen des UD-Types die nicht gedreht werden sollen,
    '                   müssen negativ angegeben werden.
    '                   Achtung: es müssen auch Padbytes berücksichtigt werden
    '
    Dim i As Long, j As Long
    Dim CountUDTElements As Long: CountUDTElements = UBound(udtDescription) + 1
    Dim udtLength As Long, ValLength As Long
    For i = 0 To CountUDTElements - 1
        udtLength = udtLength + Abs(udtDescription(i))
    Next
    With this
        .pB.pvData = pData
        For i = 0 To Count - 1
            For j = 0 To CountUDTElements - 1
                ValLength = udtDescription(j)
                .pB.cElements = Abs(ValLength)
                Select Case ValLength
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
                .pB.pvData = .pB.pvData + .pB.cElements 'Abs(ValLength)
            Next
        Next
    End With
End Sub

Public Function SwapBytesInt16(i As Integer) As Integer
    SBO_Rotate2 i: SwapBytesInt16 = i
End Function
Public Function SwapBytesInt32(i As Long) As Long
    SBO_Rotate4 i: SwapBytesInt32 = i
End Function

Public Sub String_Rotate2(s As String)
    Dim bs As TByteSwapper: New_ByteSwapper bs
    bs.pB.pvData = StrPtr(s)
    Dim i As Long
    'With bs
        For i = 0 To Len(s)
            bs.tmpByte = bs.b(0)
            bs.b(0) = bs.b(1)
            bs.b(1) = bs.tmpByte
            bs.pB.pvData = bs.pB.pvData + 2
        Next
    'End With
    ByteSwapper_Delete bs
End Sub

' ^ ############################## ^ '    MByteSwapper Functions    ' ^ ############################## ^ '

