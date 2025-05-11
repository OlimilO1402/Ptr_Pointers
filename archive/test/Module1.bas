Attribute VB_Name = "Module1"
Option Explicit
Private m_col As Collection

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

Public Function Col_TryAddObject(col As Collection, obj As Object, Key As String) As Boolean
Try: On Error GoTo Catch
    col.Add obj, Key
    Col_TryAddObject = True
Catch: On Error GoTo 0
End Function

Public Sub Col_MoveUp(col As Collection, ByVal i As Long)
    Col_SwapItems col, i, i - 1
End Sub

Public Sub Col_MoveDown(col As Collection, ByVal i As Long)
    Col_SwapItems col, i, i + 1
End Sub

Public Sub Col_SwapItems(col As Collection, ByVal i1 As Long, i2 As Long)
    Dim c As Long: c = col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or c <= i1 Then Exit Sub
    If i2 <= 0 Or c < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim Obj1, Obj2
    If IsObject(col.Item(i1)) Then Set Obj1 = col.Item(i1) Else Obj1 = col.Item(i1)
    If IsObject(col.Item(i2)) Then Set Obj2 = col.Item(i2) Else Obj2 = col.Item(i2)
    col.Remove i1: col.Add Obj2, , i1:     col.Remove i2
    If i2 < c Then col.Add Obj1, , i2 Else col.Add Obj1
End Sub

Public Sub Col_Sort(col As Collection)
    Set m_col = col
    Dim c As Long: c = m_col.Count
    If c = 0 Then Exit Sub
    Dim vt As VbVarType: vt = VarType(m_col.Item(1))
    Select Case vt
    Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal
        Call QuickSortVar(1, c)
    Case vbString
        Call QuickSortStr(1, c)
    Case vbObject
        Call QuickSortObj(1, c)
    End Select
    Set m_col = Nothing
End Sub

' Die Rekursive datenunabhängige Methode QuickSort
Private Sub QuickSortVar(ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i2 > i1 Then
        T = DivideVar(i1, i2)
        Call QuickSortVar(i1, T - 1)
        Call QuickSortVar(T + 1, i2)
    End If
End Sub
Private Function DivideVar(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareVar(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareVar(p, j) < 0))
        If i < j Then Call SwapVar(i, j)
    Loop While (i < j)
    Call SwapVar(i, p)
    DivideVar = i
End Function
Private Function CompareVar(ByVal i1 As Long, ByVal i2 As Long) As Variant
    CompareVar = m_col.Item(i1) - m_col.Item(i2)
End Function
Private Sub SwapVar(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Var1: Var1 = m_col.Item(i1)
    Dim Var2: Var2 = m_col.Item(i2)
    m_col.Remove i1: m_col.Add Obj2, , i1:   m_col.Remove i2
    If i2 < c Then m_col.Add Obj1, , i2 Else m_col.Add Obj1
End Sub

' Die Rekursive datenunabhängige Methode QuickSort für Strings
Private Sub QuickSortStr(ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i1 < i2 Then
        T = divideStr(i1, i2)
        Call QuickSortStr(i1, T - 1)
        Call QuickSortStr(T + 1, i2)
    End If
End Sub
Private Function divideStr(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareStr(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareStr(p, j) < 0))
        If i < j Then Call SwapStr(i, j)
    Loop While (i < j)
    Call SwapStr(i, p)
    divideStr = i
End Function
Private Function CompareStr(ByVal i1 As Long, ByVal i2 As Long)
    'CompareStr = StrComp(m_col.Item(i1), m_col.Item(i2))
    Dim Str1 As String: Str1 = m_col.Item(i1)
    Dim Str2 As String: Str2 = m_col.Item(i2)
    CompareStr = StrComp(Str1, Str2)
End Function
Private Sub SwapStr(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Str1 As String: Str1 = m_col.Item(i1)
    Dim Str2 As String: Str2 = m_col.Item(i2)
    m_col.Remove i1: m_col.Add Str2, , i1:   m_col.Remove i2
    If i2 < c Then m_col.Add Str1, , i2 Else m_col.Add Str1
End Sub

' Die Rekursive datenunabhängige Methode QuickSort
Private Sub QuickSortObj(ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i2 > i1 Then
        T = divideObj(i1, i2)
        Call QuickSortObj(i1, T - 1)
        Call QuickSortObj(T + 1, i2)
    End If
End Sub
Private Function divideObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareObj(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareObj(p, j) < 0))
        If i < j Then Call SwapObj(i, j)
    Loop While (i < j)
    Call SwapObj(i, p)
    divideObj = i
End Function
Private Function CompareObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim Obj1 As Object: Set Obj1 = m_col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_col.Item(i2)
    'CompareObj = m_Arr(i1).Compare(m_Arr(i2))
    CompareObj = Obj1.Compare(Obj2)
End Function
Private Sub SwapObj(ByVal i1 As Long, ByVal i2 As Long)
    'Dim tmp As Object: Set tmp = m_col.Item(i1)
    'Set m_Arr(i1) = m_Arr(i2): Set m_Arr(i2) = tmp
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_col.Count
    'If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    'If i1 <= 1 Or c <= i1 Then Exit Sub
    'If i2 <= 1 Or c < i2 Then Exit Sub
    Dim Obj1 As Object: Set Obj1 = m_col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_col.Item(i2)
    m_col.Remove i1: m_col.Add Obj2, , i1:   m_col.Remove i2
    If i2 < c Then m_col.Add Obj1, , i2 Else m_col.Add Obj1
    
End Sub
'
'Public Sub Swap(ByVal i1 As Long, ByVal i2 As Long)
'    'vertauscht die beiden Elemente z.B. für MoveUp, MoveDown
'    If m_DataType = vbObject Then
'        SwapObj i1, i2
'    ElseIf m_DataType = vbString Then
'        SwapStr i1, i2
'    Else
'        SwapVar i1, i2
'    End If
'End Sub


