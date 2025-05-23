VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "VBPointers"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestDeRef 
      Caption         =   "Test DeRef"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestVBCollection 
      Caption         =   "Test VB.Collection"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton BtnTestObjPtr 
      Caption         =   "Test ObjPtr"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton BtnTSafeArrayPtr 
      Caption         =   "Test SafeArrayPtr"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton BtnTestSAPtr 
      Caption         =   "Test SAPtr"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton BtnTestArrayPointer 
      Caption         =   "Test Array-Pointer"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton BtnTestCharArray 
      Caption         =   "Test Char-Pointer"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = "VBPointers v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub BtnTestCharArray_Click()
    Form1.Show vbModal, Me
End Sub

Private Sub BtnTSafeArrayPtr_Click()
    Form2.Show vbModal, Me
End Sub

Private Sub BtnTestArrayPointer_Click()
    Form3.Show vbModal, Me
End Sub

Private Sub BtnTestObjPtr_Click()
    Form4.Show
End Sub

Private Sub BtnTestSAPtr_Click()
    
    ReDim sa(0 To 10) As String
    sa(0) = "one"
    sa(1) = "two"
    
    Dim saX() As String
    SAPtr(StrArrPtr(saX)) = SAPtr(StrArrPtr(sa))
    
    MsgBox saX(0)
    
    ZeroSAPtr StrArrPtr(saX)
End Sub

Private Sub BtnTestVBCollection_Click()
    Form5.Show vbModal, Me
End Sub

Private Sub BtnTestDeRef_Click()
    
    Dim Value As Long: Value = 123456
    MsgBox Value
    
    Dim pValue As LongPtr: pValue = VarPtr(Value)
    
    DeRef(pValue) = 456789
    MsgBox Value
    
    TestDeRef1 Value
    MsgBox Value
    
    TestDeRef2 pValue
    MsgBox Value
    
End Sub

Sub TestDeRef1(ByRef Value As Long)
    Value = 321654
End Sub
Sub TestDeRef2(ByVal pValue As LongPtr)
    DeRef(pValue) = 654987
End Sub

