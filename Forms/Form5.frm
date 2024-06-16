VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Test VB.Collection"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   11820
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnSort 
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox CmbVarType 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton BtnCreate 
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11580
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Col As Collection

Private Sub Form_Load()
    Randomize Timer
    InitCmbVarType
End Sub

Sub InitCmbVarType()
    Dim i As Integer
    With CmbVarType
        .Clear
        .AddItem "Byte     (uint8)  ": .ItemData(i) = VbVarType.vbByte:     i = i + 1
        .AddItem "Integer  (sint16) ": .ItemData(i) = VbVarType.vbInteger:  i = i + 1
        .AddItem "Long     (sint32) ": .ItemData(i) = VbVarType.vbLong:     i = i + 1
        .AddItem "Single   (flt32)  ": .ItemData(i) = VbVarType.vbSingle:   i = i + 1
        .AddItem "Double   (flt64)  ": .ItemData(i) = VbVarType.vbDouble:   i = i + 1
        .AddItem "Currency (sint64) ": .ItemData(i) = VbVarType.vbCurrency: i = i + 1
        .AddItem "Decimal  (sint128)": .ItemData(i) = VbVarType.vbDecimal:  i = i + 1
        .AddItem "Date              ": .ItemData(i) = VbVarType.vbDate:     i = i + 1
        .AddItem "String            ": .ItemData(i) = VbVarType.vbString:   i = i + 1
        .AddItem "Object            ": .ItemData(i) = VbVarType.vbObject:   i = i + 1
        .ListIndex = 0
    End With
End Sub

Private Sub BtnCreate_Click()
    Set m_Col = New Collection
    Dim i As Long: i = CmbVarType.ListIndex
    If i < 0 Then Exit Sub
    Dim y As Integer: y = Year(Now)
    Dim vt As VbVarType: vt = CmbVarType.ItemData(i)
    Dim n As Long: n = CLng(20 + Rnd() * 100)
    Select Case vt
    Case VbVarType.vbByte:        For i = 0 To n: m_Col.Add CByte(Rnd * 255):          Next
    Case VbVarType.vbInteger:     For i = 0 To n: m_Col.Add CInt(Rnd * 32767):         Next
    Case VbVarType.vbLong:        For i = 0 To n: m_Col.Add CLng(Rnd * 1000000):       Next
    Case VbVarType.vbSingle:      For i = 0 To n: m_Col.Add CSng(Rnd * 1000000!):      Next
    Case VbVarType.vbDouble:      For i = 0 To n: m_Col.Add CDbl(Rnd * 2147484000#):   Next
    Case VbVarType.vbCurrency:    For i = 0 To n: m_Col.Add CCur(Rnd * 2147484000@):   Next
    Case VbVarType.vbDecimal:     For i = 0 To n: m_Col.Add CDec(Rnd * 2147484000#):   Next
    Case VbVarType.vbDate:        For i = 0 To n: m_Col.Add DateSerial(CInt(Rnd * y), CInt(1 + Rnd * 11), CInt(1 + Rnd * 31)): Next
    Case VbVarType.vbString:      For i = 0 To n: m_Col.Add GetRndName(5 + Rnd * 22):  Next
    Case VbVarType.vbObject:      For i = 0 To n: m_Col.Add MNew.Class1(Rnd * 10000000#): Next
    End Select
    UpdateView
End Sub

Private Sub BtnSort_Click()
    If m_Col Is Nothing Then Exit Sub
    MPtr.Col_Sort m_Col
    UpdateView
End Sub

Private Function GetRndName(ByVal length As Byte) As String
    Dim s As String: s = ChrW(65 + Rnd * 25)
    Dim i As Long
    For i = 2 To length
        s = s & ChrW(97 + Rnd * 25)
    Next
    GetRndName = s
End Function
    
Private Sub UpdateView()
    MPtr.Col_ToListBox m_Col, Me.List1
End Sub

