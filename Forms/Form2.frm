VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   ScaleHeight     =   4215
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnArrayOfObject 
      Caption         =   "Array Of Object"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton BtnArrayOfString 
      Caption         =   "Array Of String"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton BtnArrayOfDouble 
      Caption         =   "Array Of Double"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnArrayOfDouble_Click()
    
    ReDim DblArr(0 To 10) As Double ' eleven elements
    
    Dim sa As TSafeArrayPtr: MPtr.New_SafeArrayPtr sa
    
    MPtr.SafeArrayPtr_SAPtr(sa) = ArrPtr(DblArr)
    
    Text1.Text = MPtr.SafeArrayPtr_ToStr(sa)
    
    MPtr.SafeArrayPtr_Delete sa
    
End Sub

Private Sub BtnArrayOfString_Click()
    
    ReDim strArr(0 To 255) As String '256 elements
    
    Dim sa As TSafeArrayPtr: MPtr.New_SafeArrayPtr sa
    
    MPtr.SafeArrayPtr_SAPtr(sa) = StrArrPtr(strArr)
    
    Text1.Text = MPtr.SafeArrayPtr_ToStr(sa)
    
    MPtr.SafeArrayPtr_Delete sa
    
End Sub

Private Sub BtnArrayOfObject_Click()
    
    ReDim ObjArr(0 To 5) As Object '256 elements
    
    Dim sa As TSafeArrayPtr: MPtr.New_SafeArrayPtr sa
    
    MPtr.SafeArrayPtr_SAPtr(sa) = ArrPtr(ObjArr)
    
    Text1.Text = MPtr.SafeArrayPtr_ToStr(sa)
    
    MPtr.SafeArrayPtr_Delete sa
    
End Sub

