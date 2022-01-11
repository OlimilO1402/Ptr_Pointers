VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "VBPointers"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "FMain"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestArrayPointer 
      Caption         =   "Test Array-Pointer"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton BtnTestByteSwapper 
      Caption         =   "Test Byte-Swapper"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton BtnTestCharArray 
      Caption         =   "Test Char-Pointer"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnTestCharArray_Click()
    Form1.Show vbModal, Me
End Sub

Private Sub BtnTestByteSwapper_Click()
    Form2.Show vbModal, Me
End Sub

Private Sub BtnTestArrayPointer_Click()
    Form3.Show vbModal, Me
End Sub

