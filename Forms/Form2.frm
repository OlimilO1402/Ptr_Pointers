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
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
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

Private Sub Command1_Click()
    
    ReDim dblarr(0 To 10) As Double
    
    Dim sa As TSafeArrayPtr: MPtr.New_SafeArrayPtr sa
    
    MPtr.SafeArrayPtr_SAPtr(sa) = ArrPtr(dblarr)
    
    Text1.Text = MPtr.SafeArrayPtr_ToStr(sa)
    
    MPtr.SafeArrayPtr_Delete sa
    
End Sub
