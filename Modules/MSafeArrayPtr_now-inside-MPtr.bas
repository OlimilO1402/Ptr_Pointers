Attribute VB_Name = "MSafeArrayPtr"
Option Explicit

' the TSafeArrayPtr lightweight-object structure
Public Type TSafeArrayPtr
    pSAPtr As TUDTPtr
    pSA()  As TUDTPtr
End Type

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
    'Call GetMem4(ByVal Value, p)
    
    this.pSAPtr.pvData = p - 2 * SizeOf_LongPtr
    ' -8 ist Mist
    ' -8 weil zuerst pSA und Reserved und dann kommt erst
    ' der Anfang der SafeArrayDesc-Struktur mit cDims
End Property

Public Function SafeArrayPtr_ToStr(this As TSafeArrayPtr) As String
    SafeArrayPtr_ToStr = MPtr.UDTPtr_ToStr(this.pSA(0))
End Function
