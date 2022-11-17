VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "VBPointers"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTSafeArrayPtr 
      Caption         =   "Test SafeArrayPtr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton BtnTestSAPtr 
      Caption         =   "Test SAPtr"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton BtnTestArrayPointer 
      Caption         =   "Test Array-Pointer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton BtnTestCharArray 
      Caption         =   "Test Char-Pointer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnTestSAPtr_Click()
    ReDim sa(0 To 10) As String
    sa(0) = "one"
    sa(1) = "two"
    
    Dim saX() As String
    SAPtr(StrArrPtr(saX)) = SAPtr(StrArrPtr(sa))
    
    MsgBox saX(0)
    
    ZeroSAPtr StrArrPtr(saX)
End Sub

Private Sub BtnTSafeArrayPtr_Click()
    Form2.Show
End Sub

Private Sub Form_Load()
    Me.Caption = "VBPointers v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub BtnTestCharArray_Click()
    Form1.Show vbModal, Me
End Sub

'Private Sub BtnTestByteSwapper_Click()
'    Form2.Show vbModal, Me
'End Sub

Private Sub BtnTestArrayPointer_Click()
    Form3.Show vbModal, Me
End Sub
