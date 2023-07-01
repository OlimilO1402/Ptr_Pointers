VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Test Array-Pointer"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "Test Object Array"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test UD-Type Array"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test String Array"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Long Array"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TAnyType
    XVar As Long
    YVar As Single
    ZVar As Double
    SVar As String
End Type

Private Sub Command1_Click()
    Call TestLongArray
End Sub

Private Sub Command2_Click()
    Call TestStringArray
End Sub

Private Sub Command3_Click()
    Call TestUDTypeArray
End Sub

Private Sub Command4_Click()
    Call TestObjectArray
End Sub

Private Sub TestLongArray()
    Dim lngArr1() As Long
    Dim lngArr2() As Long
    ReDim lngArr1(0 To 1)
    lngArr1(0) = 123456
    lngArr1(1) = 456789
    
    'der Zeiger wird von lngArr1 ausgelesen und in lngArr2 hineinkopiert,
    'das manipulierte Array ist lngArr2
    SAPtr(ArrPtr(lngArr2)) = SAPtr(ArrPtr(lngArr1))
    MsgBox CStr(lngArr1(0)) & "    " & CStr(lngArr1(1))
    MsgBox CStr(lngArr2(0)) & "    " & CStr(lngArr2(1))
    
    'Achtung Wichtig:
    'den Zeiger des Manipulierten Arrays wieder Nullen, sonst gibt es
    'einen Absturz der IDE und des Programmes, beim Versuch von VB
    'beide Arrays zu löschen, bzw den Speicher wieder frei zu geben.
    'entweder wieder über das Property, oder mit einer Nuller-Funktion
    'SAPtr(ArrPtr(lngArr2)) = 0
    Call ZeroSAPtr(ArrPtr(lngArr2))
End Sub

Private Sub TestStringArray()
    Dim strArr1() As String
    Dim strArr2() As String
    ReDim strArr1(0 To 1)
    strArr1(0) = "der erste String"
    strArr1(1) = "der zweite String"
    
    'der Zeiger wird von strArr1 ausgelesen und in strArr2 hineinkopiert,
    'das manipulierte Array ist strArr2
    SAPtr(StrArrPtr(strArr2)) = SAPtr(StrArrPtr(strArr1))
    MsgBox strArr1(0) & "    " & strArr1(1)
    MsgBox strArr2(0) & "    " & strArr2(1)
    
    'Achtung Wichtig:
    'den Zeiger des Manipulierten Arrays wieder Nullen, sonst gibt es
    'einen Absturz der IDE und des Programmes, beim Versuch von VB
    'beide Arrays zu löschen, bzw den Speicher wieder frei zu geben.
    'entweder wieder über das Property, oder mit einer Nuller-Funktion
    'SAPtr(StrArrPtr(strArr2)) = 0
    Call ZeroSAPtr(StrArrPtr(strArr2))
End Sub

Private Sub TestUDTypeArray()
    Dim udtArr1() As TAnyType
    Dim udtArr2() As TAnyType
    ReDim udtArr1(0 To 1)
    udtArr1(0).XVar = 123456
    udtArr1(0).YVar = 123456.789
    udtArr1(0).ZVar = 123456789.123456
    udtArr1(0).SVar = "der erste String"
    udtArr1(1).XVar = 654321
    udtArr1(1).YVar = 987654.321
    udtArr1(1).ZVar = 987654321.987654
    udtArr1(1).SVar = "der zweite String"
    
    'der Zeiger wird von udtArr1 ausgelesen und in udtArr2 hineinkopiert,
    'das manipulierte Array ist udtArr2
    SAPtr(ArrPtr(udtArr2)) = SAPtr(ArrPtr(udtArr1))
    MsgBox TAnyTypeToStr(udtArr1(0)) & vbCrLf & TAnyTypeToStr(udtArr1(1))
    MsgBox TAnyTypeToStr(udtArr2(0)) & vbCrLf & TAnyTypeToStr(udtArr2(1))
    
    'Achtung Wichtig:
    'den Zeiger des Manipulierten Arrays wieder Nullen, sonst gibt es
    'einen Absturz der IDE und des Programmes, beim Versuch von VB
    'beide Arrays zu löschen, bzw den Speicher wieder frei zu geben.
    'entweder wieder über das Property, oder mit einer Nuller-Funktion
    'SAPtr(StrArrPtr(strArr2)) = 0
    Call ZeroSAPtr(ArrPtr(udtArr2))
End Sub

Private Function TAnyTypeToStr(A As TAnyType) As String
    Dim m As String
    m = m & CStr(A.XVar) & vbCrLf
    m = m & CStr(A.YVar) & vbCrLf
    m = m & CStr(A.ZVar) & vbCrLf
    m = m & A.SVar & vbCrLf
    TAnyTypeToStr = m
End Function

Private Sub TestObjectArray()
    ReDim objArr1(0 To 1) As Class1
    Set objArr1(0) = New_Class1(123456789.123456)
    Set objArr1(1) = New_Class1(987654321.987654)
    
    'der Zeiger wird von objArr1 ausgelesen und in objArr2 hineinkopiert,
    'das manipulierte Array ist objArr2
    
    Dim objArr2() As Class1
    SAPtr(ArrPtr(objArr2)) = SAPtr(ArrPtr(objArr1))
    MsgBox objArr1(0).ToStr & vbCrLf & _
           objArr1(1).ToStr
    MsgBox objArr2(0).ToStr & vbCrLf & _
           objArr2(1).ToStr
           
    'Achtung Wichtig:
    'den Zeiger des Manipulierten Arrays wieder Nullen, sonst gibt es
    'einen Absturz der IDE und des Programmes, beim Versuch von VB
    'beide Arrays zu löschen, bzw den Speicher wieder frei zu geben.
    'entweder wieder über das Property, oder mit einer Nuller-Funktion
    'SAPtr(StrArrPtr(strArr2)) = 0
    Call ZeroSAPtr(ArrPtr(objArr2))

End Sub

Public Function New_Class1(ByVal aValue As Double) As Class1
    Set New_Class1 = New Class1: New_Class1.Value = aValue
End Function
