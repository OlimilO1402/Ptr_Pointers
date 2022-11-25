Attribute VB_Name = "MByteSwapper"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : MByteSwapper
' DateTime  : 11.10.2008 19:40
' Author    : Oliver Meyer (olimilo AT gmx DOT net)
' Purpose   : swaps the sequence of bytes of 16, 32 and 64 Bit datatypes
'---------------------------------------------------------------------------------------

' the TByteSwapper lightweight-object structure
Public Type TByteSwapper
    pB      As TUDTPtr
    tmpByte As Byte
    b()     As Byte
End Type

'The following functions can be found in the dll SwapByteOrder, written in assembler, and are much faster.
Public Declare Sub SBO_Rotate2 Lib "SwapByteOrder" Alias "SwapByteOrder16" (ByRef Ptr As Any)
Public Declare Sub SBO_Rotate4 Lib "SwapByteOrder" Alias "SwapByteOrder32" (ByRef Ptr As Any)
Public Declare Sub SBO_Rotate8 Lib "SwapByteOrder" Alias "SwapByteOrder64" (ByRef Ptr As Any)
Public Declare Function SBO_RotateArray Lib "SwapByteOrder" Alias "SwapByteOrderArray" (ByRef Value() As Any) As Long
Public Declare Sub SBO_RotateUDTArray Lib "SwapByteOrder" Alias "SwapByteOrderUDTArray" (ByRef Arr() As Any, ByRef udtDescription() As Integer)
'
'now the big question
'should we maybe integrate MSwapByteOrder completely into MPtr?
' * einerseits bin ich eher der Typ der verschiedene Dinge in verschiedene Module ordnet, dann hat alles schön seine Ordnung und ist getrennt voneinander.
'   jetzt ist die Frage, ist SwapByteOrder was anderes als Zeiger, je einerseits doch schon, es benutzt die UDTPtr-Methode nur intensiv
' * andererseits noch ein Modul mehr . . . Unicode wird von MString gebraucht, PathFileName braucht es und das sind Klassen die in beinahe jedem Projekt
'   enthalten sind, so müßte man kein einziges Projekte updaten das PathFilename oder MString benützt
'
'OK now an idea what about copying it into MPtr and separating it with special lines

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
    
    Debug.Print "TByteSwapper.pB(=TUDTPtr)" & vbCrLf & UDTPtr_ToStr(this.pB)
    
    Debug.Print "SafeArrayPtr.pSAPtr:     " & vbCrLf & UDTPtr_ToStr(pSA.pSAPtr)
    
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
            bs.pB.pvData = bs.pB.pvData + 2
            bs.tmpByte = bs.b(0)
            bs.b(0) = bs.b(1)
            bs.b(1) = bs.tmpByte
        Next
    'End With
    ByteSwapper_Delete bs
End Sub
