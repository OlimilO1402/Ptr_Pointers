VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "Sort"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sort"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Swap i1, i2"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Swap i1, i2"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_List1 As New Collection 'enthält Long-Variablen
Private m_List2 As New Collection 'enthält Objekte

Private Sub Form_Load()
    
    Dim Key As String
    Dim i As Long
    
    'For i = 1000 To 1 Step -100
    For i = 1 To 10
        'Key = CStr(i)
        Key = CLng(Rnd * 1000) '
        m_List1.Add Key, Key
    Next
    UpdateView1
    
    Dim obj As Class1, v As Long
    For i = 1 To 10
        v = CLng(Rnd * 1000) '
        Set obj = New Class1: obj.Value = v
        Key = i
        m_List2.Add obj, Key
    Next
    UpdateView2
    
End Sub

Private Sub Command1_Click()
    Col_SwapItems m_List1, 5, 6
    UpdateView1
End Sub

Private Sub Command2_Click()
    Col_Sort m_List1
    UpdateView1
End Sub

Private Sub Command3_Click()
    Col_SwapItems m_List2, 5, 6
    UpdateView2
End Sub

Private Sub Command4_Click()
    Col_Sort m_List2
    UpdateView2
End Sub

Sub UpdateView1()
    List1.Clear
    Dim i As Long
    For i = 1 To m_List1.Count
        List1.AddItem m_List1.Item(i)
    Next
End Sub

Sub UpdateView2()
    List2.Clear
    Dim i As Long, o 'As Object
    For i = 1 To m_List2.Count
        Set o = m_List2.Item(i)
        List2.AddItem o.Value
    Next
End Sub

'Public Sub Sort(Optional ComparerObj As Object, Optional lambdaFunc As String)
''Sub List(Of T).Sort()
''Sub List(Of T).Sort(comparison As Comparison(Of T)
''Sub List(Of T).Sort(comparer As IComparer(Of T))
''Sub List(Of T).Sort(index As Integer, count As Integer, comparer As IComparer(Of T))
'    'Sortiert die Elemente in der gesamten List(Of T) mithilfe des Standardcomparers.
'    'Sortiert die Liste mit einem jeweils nach Datentyp angepassten QuickSort
'    If Not ComparerObj Is Nothing Then
'        Set m_LambObj = ComparerObj
'        m_LambFunc = lambdaFunc
'    End If
'    Select Case m_DataType
'    Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal
'        Call QuickSortVar(0, m_Count - 1)
'    Case vbString
'        Call QuickSortStr(0, m_Count - 1)
'    Case vbObject
'        Call QuickSortObj(0, m_Count - 1)
'    End Select
'    If m_IsHashed Then ReInitIndices
'    Set m_LambObj = Nothing
'    m_LambFunc = vbNullString
'End Sub
'Public Sub SortRev(Optional ComparerObj As Object, Optional lambdaFunc As String)
'    If Not ComparerObj Is Nothing Then
'        Set m_LambObj = ComparerObj
'        m_LambFunc = lambdaFunc
'    End If
'    Select Case m_DataType
'    Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal
'        Call QuickSortVar(0, m_Count - 1)
'    Case vbString
'        Call QuickSortStr(0, m_Count - 1)
'    Case vbObject
'        Call QuickSortObj(0, m_Count - 1)
'    End Select
'    'If m_IsHashed Then ReInitIndices
'    Reverse
'    Set m_LambObj = Nothing
'    m_LambFunc = vbNullString
'End Sub
'Public Sub Sort(Optional ByVal i1 As Long = 0, Optional ByVal i2 As Long = m_Count - 1)
'
'    Call QuickSort(i1, i2)
'
'End Sub

