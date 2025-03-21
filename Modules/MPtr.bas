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

'#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
'#End If
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

'Public Type VBGuid
'    Data1 As Long
'    Data2 As Integer
'    Data3 As Integer
'    Data5(0 To 7) As Byte
'End Type
'
'Public GUID_NULL As VBGuid

Public Type TCharPointer
    pudt    As TUDTPtr
    Chars() As Integer
End Type

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
    B()     As Byte
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
Private m_Col As Collection

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

'Get the pointer of something with VarPtr
'Public Property Get DeRef ?
'the "Property Get" of "Property Let DeRef" would just be VarPtr()
Public Property Let DeRef(ByVal Ptr As LongPtr, ByVal Value As LongPtr)
    RtlMoveMemory ByVal Ptr, Value, SizeOf_LongPtr
End Property
'use it like:
'Dim V  As Long:     V = 123: MsgBox V
'Dim pV As LongPtr: pV = VarPtr(V)
'DeRef(pV) = 456: MsgBox V

' v ############################## v '    Collection Functions    ' v ############################## v '
Public Function Col_Contains(Col As Collection, Key As String) As Boolean
    'for this Function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original Version of this function here: https://vb-tec.de/collctns.htm
    On Error Resume Next
'  '"Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(Col(Key)) Then: 'DoNothing
    Col_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Col_Add(Col As Collection, obj As Object) As Object
    Set Col_Add = obj:  Col.Add obj
End Function

'Public Function Col_Add(Col As Collection, Value)
'   'Nope just use eiter "Col.Add Value" or Col.Add Value, CStr(Value)
'    Col_AddV = Value:   Col.Add Value
'End Function

Public Function Col_AddKey(Col As Collection, obj As Object) As Object
    Set Col_AddKey = obj:  Col.Add obj, obj.Key ' the object needs to have a Public Function/PropertyGet Key As String
End Function

Public Function Col_AddOrGet(Col As Collection, obj As Object) As Object
    Dim Key As String: Key = obj.Key ' the object needs to have a Public Function Key As String
    If Col_Contains(Col, Key) Then
        Set Col_AddOrGet = Col.Item(Key)
    Else
        Set Col_AddOrGet = obj
        Col.Add obj, Key
    End If
End Function

Public Function Col_TryAddObject(Col As Collection, obj As Object, ByVal Key As String) As Boolean
Try: On Error GoTo Catch
    Col.Add obj, Key
    Col_TryAddObject = True
Catch: On Error GoTo 0
End Function

Public Sub Col_Remove(Col As Collection, obj As Object)
    Dim o As Object
    For Each o In Col
        If o.IsSame(obj) Then 'Obj needs Public Function IsSame(other) As Boolean
            If Col_Contains(Col, obj.Key) Then Col.Remove obj.Key 'Obj needs Public Property Key As String
        End If
    Next
End Sub

Public Sub Col_SwapItems(Col As Collection, ByVal i1 As Long, i2 As Long)
    Dim c As Long: c = Col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or c <= i1 Then Exit Sub
    If i2 <= 0 Or c < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim Obj1, Obj2
    If IsObject(Col.Item(i1)) Then Set Obj1 = Col.Item(i1) Else Obj1 = Col.Item(i1)
    If IsObject(Col.Item(i2)) Then Set Obj2 = Col.Item(i2) Else Obj2 = Col.Item(i2)
    Col.Remove i1: Col.Add Obj2, , i1:     Col.Remove i2
    If i2 < c Then Col.Add Obj1, , i2 Else Col.Add Obj1
End Sub

Public Sub Col_MoveUp(Col As Collection, ByVal i As Long)
    Col_SwapItems Col, i, i - 1
End Sub

Public Sub Col_MoveDown(Col As Collection, ByVal i As Long)
    Col_SwapItems Col, i, i + 1
End Sub

Public Sub Col_ToListBox(Col As Collection, aLB As ListBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    Col_ToListCtrl Col, aLB, addEmptyLineFirst
End Sub

Public Sub Col_ToComboBox(Col As Collection, aCB As ComboBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    Col_ToListCtrl Col, aCB, addEmptyLineFirst
End Sub

Public Sub Col_ToListCtrl(Col As Collection, ComboBoxOrListBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    If Col Is Nothing Then Exit Sub
    Dim i As Long, c As Long: c = Col.Count: If c = 0 Then Exit Sub
    Dim vt As VbVarType: vt = VarType(Col.Item(1))
    Dim v, obj As Object
    With ComboBoxOrListBox
        If .ListCount Then .Clear
        If addEmptyLineFirst Then .AddItem vbNullString
        Select Case vt
        Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal, vbString
            For i = 1 To c
                .AddItem Col.Item(i)
            Next
        Case vbObject
            For i = 1 To c
                Set obj = Col.Item(i)
                .AddItem obj.ToStr ' the object needs to have a Public Function ToStr As String
                If doPtrToItemData Then .ItemData(i - 1) = obj.Ptr ' and a Public Function Ptr As LongPtr
            Next
        End Select
    End With
End Sub

Public Property Get Col_ObjectFromListCtrl(Col As Collection, ComboBoxOrListBox, i_out As Long) As Object
    i_out = ComboBoxOrListBox.ListIndex
    If i_out < 0 Then Exit Property
    Dim Key As String: Key = ComboBoxOrListBox.ItemData(i_out)
    If Col_Contains(Col, Key) Then Set Col_ObjectFromListCtrl = Col.Item(Key)
End Property

Public Sub Col_Sort(Col As Collection)
    Set m_Col = Col
    Dim c As Long: c = m_Col.Count
    If c = 0 Then: Set m_Col = Nothing: Exit Sub
    Dim vt As VbVarType: vt = VarType(m_Col.Item(1))
    Select Case vt
    Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal
        Col_QuickSortVar 1, c
    Case vbString
        Col_QuickSortStr 1, c
    Case vbObject
        Col_QuickSortObj 1, c
    End Select
    Set m_Col = Nothing
End Sub

' The recursive data-independent QuickSort for primitive data-variables
Private Sub Col_QuickSortVar(ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i2 > i1 Then
        T = Col_DivideVar(i1, i2)
        Col_QuickSortVar i1, T - 1
        Col_QuickSortVar T + 1, i2
    End If
End Sub

Private Function Col_DivideVar(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (Col_CompareVar(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (Col_CompareVar(p, j) < 0))
        If i < j Then Col_SwapVar i, j
    Loop While (i < j)
    Col_SwapVar i, p
    Col_DivideVar = i
End Function

Private Function Col_CompareVar(ByVal i1 As Long, ByVal i2 As Long) As Variant
    Col_CompareVar = m_Col.Item(i1) - m_Col.Item(i2)
End Function

Private Sub Col_SwapVar(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Var1: Var1 = m_Col.Item(i1)
    Dim Var2: Var2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Var2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Var1, , i2 Else m_Col.Add Var1
End Sub

' The recursive data-independent QuickSort for strings
Private Sub Col_QuickSortStr(ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i1 < i2 Then
        T = Col_DivideStr(i1, i2)
        Col_QuickSortStr i1, T - 1
        Col_QuickSortStr T + 1, i2
    End If
End Sub

Private Function Col_DivideStr(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (Col_CompareStr(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (Col_CompareStr(p, j) < 0))
        If i < j Then Col_SwapStr i, j
    Loop While (i < j)
    Col_SwapStr i, p
    Col_DivideStr = i
End Function

Private Function Col_CompareStr(ByVal i1 As Long, ByVal i2 As Long)
    Col_CompareStr = StrComp(m_Col.Item(i1), m_Col.Item(i2))
    'Dim Str1 As String: Str1 = m_col.Item(i1)
    'Dim Str2 As String: Str2 = m_col.Item(i2)
    'CompareStr = StrComp(Str1, Str2)
End Function

Private Sub Col_SwapStr(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Str1 As String: Str1 = m_Col.Item(i1)
    Dim Str2 As String: Str2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Str2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Str1, , i2 Else m_Col.Add Str1
End Sub

' The recursive data-independent QuickSort for objects
Private Sub Col_QuickSortObj(ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i2 > i1 Then
        T = Col_DivideObj(i1, i2)
        Col_QuickSortObj i1, T - 1
        Col_QuickSortObj T + 1, i2
    End If
End Sub

Private Function Col_DivideObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (Col_CompareObj(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (Col_CompareObj(p, j) < 0))
        If i < j Then Col_SwapObj i, j
    Loop While (i < j)
    Col_SwapObj i, p
    Col_DivideObj = i
End Function

Private Function Col_CompareObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    Col_CompareObj = Obj1.compare(Obj2)
End Function

Private Sub Col_SwapObj(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Obj2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Obj1, , i2 Else m_Col.Add Obj1
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
'cElements in Abh�ngigkeit von cbElement entsprechend angepasst wird
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
    Dim obj As IUnknown:  RtlMoveMemory obj, p, MPtr.SizeOf_LongPtr
    Set PtrToObject = obj: ZeroObject obj
End Function

Public Sub ZeroObject(ByVal obj As Object)
    'RtlZeroMemory ByVal VarPtr(obj), MPtr.SizeOf_LongPtr
    RtlZeroMemory obj, MPtr.SizeOf_LongPtr
End Sub

Public Sub AssignSwap(Obj1 As IUnknown, Obj2 As IUnknown)
    Dim pObj1 As LongPtr: pObj1 = ObjPtr(Obj1)
    Dim pObj2 As LongPtr: pObj2 = ObjPtr(Obj2)
    RtlMoveMemory Obj1, pObj2, MPtr.SizeOf_LongPtr
    RtlMoveMemory Obj2, pObj1, MPtr.SizeOf_LongPtr
End Sub

Public Function VTablePtr(ByVal obj As Object) As LongPtr
    RtlMoveMemory VTablePtr, ByVal ObjPtr(obj), SizeOf_LongPtr
End Function

Public Property Get ObjectAddressOf(ByVal obj As Object, ByVal Index As Long) As LongPtr
    Dim pVTable As LongPtr: pVTable = VTablePtr(obj) 'first DeRef
    RtlMoveMemory ObjectAddressOf, ByVal pVTable + (7 + Index) * SizeOf_LongPtr, SizeOf_LongPtr
End Property

Public Property Let ObjectAddressOf(ByVal obj As Object, ByVal Index As Long, ByVal Value As LongPtr)
    Dim pVTable As LongPtr: pVTable = VTablePtr(obj) 'first DeRef
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
'Public Function Min(V1, V2)
'    If V1 < V2 Then Min = V1 Else Min = V2
'End Function
'
'Public Function Max(V1, V2)
'    If V1 > V2 Then Max = V1 Else Max = V2
'End Function
' ^ ############################## ^ '         Math Functions       ' ^ ############################## ^ '

' v ############################## v '    MByteSwapper Functions   ' v ############################## v '

Public Sub New_ByteSwapper(this As TByteSwapper, Optional ByVal CountBytes As Long = 2)
    ' creates a new TByteSwapper lightweight-object
    With this
        New_UDTPtr .pB, FADF_EMBEDDED Or FADF_STATIC, 1, CountBytes
        'Call PutMem4(ByVal ArrPtr(.b), .pB.pSA)
        SAPtr(ArrPtr(.B)) = .pB.pSA
    End With
End Sub

'Public Sub New_ByteSwapperS(this As TByteSwapper, Value As String)
'
'End Sub

Public Sub ByteSwapper_Delete(this As TByteSwapper)
    ' deletes the pointer in the array of the TByteSwapper-structurr
    ZeroSAPtr ByVal ArrPtr(this.B)
End Sub

Public Sub Rotate2(this As TByteSwapper)
    ' swaps 2 bytes
    With this
        .tmpByte = .B(0)
        .B(0) = .B(1)
        .B(1) = .tmpByte
    End With
End Sub
Public Sub Rotate4(this As TByteSwapper)
    ' swaps 4 bytes
    With this
        .tmpByte = .B(0)
        .B(0) = .B(3)
        .B(3) = .tmpByte
        
        .tmpByte = .B(1)
        .B(1) = .B(2)
        .B(2) = .tmpByte
    End With
End Sub
Public Sub Rotate8(this As TByteSwapper)
    ' swaps 8 bytes
    With this
        .tmpByte = .B(0)
        .B(0) = .B(7)
        .B(7) = .tmpByte
        
        .tmpByte = .B(1)
        .B(1) = .B(6)
        .B(6) = .tmpByte
        
        .tmpByte = .B(2)
        .B(2) = .B(5)
        .B(5) = .tmpByte
        
        .tmpByte = .B(3)
        .B(3) = .B(4)
        .B(4) = .tmpByte
    End With
End Sub
Public Sub Rotate(this As TByteSwapper)
    ' rotates the bytes about the size of the datatype found in cElements
    With this
        Select Case .pB.cElements
        Case 2
            .tmpByte = .B(0)
            .B(0) = .B(1)
            .B(1) = .tmpByte
        Case 4
            .tmpByte = .B(0)
            .B(0) = .B(3)
            .B(3) = .tmpByte
            
            .tmpByte = .B(1)
            .B(1) = .B(2)
            .B(2) = .tmpByte
        Case 8
            .tmpByte = .B(0)
            .B(0) = .B(7)
            .B(7) = .tmpByte
            
            .tmpByte = .B(1)
            .B(1) = .B(6)
            .B(6) = .tmpByte
            
            .tmpByte = .B(2)
            .B(2) = .B(5)
            .B(5) = .tmpByte
            
            .tmpByte = .B(3)
            .B(3) = .B(4)
            .B(4) = .tmpByte
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
                    .tmpByte = .B(0)
                    .B(0) = .B(1)
                    .B(1) = .tmpByte
                Next
            Case 4
                For i = lb To ub
                    .pB.pvData = .pB.pvData + 4
                    .tmpByte = .B(0)
                    .B(0) = .B(3)
                    .B(3) = .tmpByte
                    
                    .tmpByte = .B(1)
                    .B(1) = .B(2)
                    .B(2) = .tmpByte
                Next
            Case 8
                For i = lb To ub
                    .pB.pvData = .pB.pvData + 8
                    .tmpByte = .B(0)
                    .B(0) = .B(7)
                    .B(7) = .tmpByte
                    
                    .tmpByte = .B(1)
                    .B(1) = .B(6)
                    .B(6) = .tmpByte
                    
                    .tmpByte = .B(2)
                    .B(2) = .B(5)
                    .B(5) = .tmpByte
                    
                    .tmpByte = .B(3)
                    .B(3) = .B(4)
                    .B(4) = .tmpByte
                Next
            End Select
        End With
    End If
    SafeArrayPtr_Delete pSA
End Sub

'Eine entsprechende Funktion in der SwapByteOrderDll k�nnte in etwa so aussehen
'wie diese Funktion
'allerdings statt pData und Count k�nnte das Array direkt angegeben werden,
'As Any machts m�glich.
'eine entsprechende Deklaration k�nnte so aussehen:
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
    '                   Der Wert der Integer-Elemente im Array repr�sentiert
    '                   die Gr��e der einzelnen Variablen-Elemente des UD-Types
    '                   verwende dazu die Funktion LenB.
    '                   Variablen des UD-Types die nicht gedreht werden sollen,
    '                   m�ssen negativ angegeben werden.
    '                   Achtung: es m�ssen auch Padbytes ber�cksichtigt werden
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
                    .tmpByte = .B(0)
                    .B(0) = .B(1)
                    .B(1) = .tmpByte
                Case 4
                    .tmpByte = .B(0)
                    .B(0) = .B(3)
                    .B(3) = .tmpByte
                    
                    .tmpByte = .B(1)
                    .B(1) = .B(2)
                    .B(2) = .tmpByte
                Case 8
                    .tmpByte = .B(0)
                    .B(0) = .B(7)
                    .B(7) = .tmpByte
                    
                    .tmpByte = .B(1)
                    .B(1) = .B(6)
                    .B(6) = .tmpByte
                    
                    .tmpByte = .B(2)
                    .B(2) = .B(5)
                    .B(5) = .tmpByte
                    
                    .tmpByte = .B(3)
                    .B(3) = .B(4)
                    .B(4) = .tmpByte
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
            bs.tmpByte = bs.B(0)
            bs.B(0) = bs.B(1)
            bs.B(1) = bs.tmpByte
            bs.pB.pvData = bs.pB.pvData + 2
        Next
    'End With
    ByteSwapper_Delete bs
End Sub

' ^ ############################## ^ '    MByteSwapper Functions    ' ^ ############################## ^ '

' v ############################## v '    CharPointer Functions   ' v ############################## v '

Public Sub New_CharPointer(ByRef this As TCharPointer, ByRef StrVal As String, Optional Base1 As Boolean = False)
    With this
        New_UDTPtr .pudt, FADF_AUTO Or FADF_FIXEDSIZE, 2, Len(StrVal), IIf(Base1, 1, 0)
        With .pudt
            .pvData = StrPtr(StrVal)
        End With
        RtlMoveMemory ByVal ArrPtr(.Chars), ByVal VarPtr(.pudt), 4
    End With
End Sub

Public Sub DeleteCharPointer(ByRef this As TCharPointer)
    With this
        RtlZeroMemory ByVal ArrPtr(.Chars), 4
    End With
End Sub

' ^ ############################## ^ '    CharPointer Functions   ' ^ ############################## ^ '

' v ############################## v '      Random Functions      ' v ############################## v '
Public Function RndInt8() As Integer
    Randomize Timer
    RndInt8 = CInt(Rnd * 255 - 128)
End Function

Public Function RndUInt8() As Byte
    Randomize Timer
    RndUInt8 = CByte(Rnd * 255)
End Function

Public Function RndInt16() As Integer
    Randomize Timer
    RndInt16 = CInt(Rnd * 65536 - 32768)
End Function

Public Function RndUInt16() As Long
    Randomize Timer
    RndUInt16 = CLng(Rnd * 65536)
End Function

Public Function RndInt32() As Long
    Randomize Timer
    RndInt32 = CLng(Rnd * 4294967296# - 2147483648#)
End Function

Public Function RndUInt32() 'As Decimal
    Randomize Timer
    RndUInt32 = CDec(Rnd * 4294967296#)
End Function

'Public Function RndInt64() 'As Decimal
'    Randomize Timer
'    RndInt64 = CLng(Rnd * CDec("18446744073709551614") - CDec("9223372036854775808"))
'End Function
