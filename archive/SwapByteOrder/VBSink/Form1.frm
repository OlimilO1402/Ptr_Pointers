VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   5280
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Declare Sub SwapByteOrder16 Lib "SwapByteOrder.dll" (ByRef Value As Integer)
Private Declare Sub SwapByteOrder32 Lib "SwapByteOrder.dll" (ByRef Value As Single)
Private Declare Sub SwapByteOrder64 Lib "SwapByteOrder.dll" (ByRef Value As Double)
Private Declare Function SwapByteOrderArray Lib "SwapByteOrder.dll" (ByRef Value() As Any) As Long

Private Sub Command1_Click()
Dim a() As Integer, x As Long, y As Long, z As Long

    ReDim a(5000, 5000)
    'Text1.Text = "SwapByteOrder16 ->" & vbCrLf
    z = GetTickCount
    For x = 0 To 5000
        For y = 0 To 5000
            a(x, y) = Rnd * 16000
            SwapByteOrder16 a(x, y)
        Next
    Next
    z = GetTickCount - z
    Text1.Text = Text1.Text & "SwapByteOrder16: " & z & "ms" & vbCrLf
    Text1.Text = Text1.Text & "SwapByteOrderArray: "
    x = GetTickCount
    Call SwapByteOrderArray(a)
    y = GetTickCount
    Text1.Text = Text1.Text & y - x & "ms" & vbCrLf
End Sub

Private Sub Command2_Click()
Dim a() As Single, x As Long, y As Long, z As Long

    ReDim a(5000, 5000)
    'Text1.Text = "SwapByteOrder32 ->" & vbCrLf
    z = GetTickCount
    For x = 0 To 5000
        For y = 0 To 5000
            a(x, y) = Rnd * 160000
            SwapByteOrder32 a(x, y)
        Next
    Next
    z = GetTickCount - z
    Text1.Text = Text1.Text & "SwapByteOrder32: " & z & "ms" & vbCrLf
    Text1.Text = Text1.Text & "SwapByteOrderArray: "
    x = GetTickCount
    Call SwapByteOrderArray(a)
    y = GetTickCount
    Text1.Text = Text1.Text & y - x & "ms" & vbCrLf
End Sub

Private Sub Command3_Click()
Dim a() As Double, x As Long, y As Long, z As Long

    ReDim a(5000, 5000)
    'Text1.Text = "SwapByteOrder64 ->" & vbCrLf
    z = GetTickCount
    For x = 0 To 5000
        For y = 0 To 5000
            a(x, y) = Rnd * 160000000
            SwapByteOrder64 a(x, y)
        Next
    Next
    z = GetTickCount - z
    Text1.Text = Text1.Text & "SwapByteOrder64: " & z & "ms" & vbCrLf
    Text1.Text = Text1.Text & "SwapByteOrderArray: "
    x = GetTickCount
    Call SwapByteOrderArray(a)
    y = GetTickCount
    Text1.Text = Text1.Text & y - x & "ms" & vbCrLf
End Sub

Private Sub Form_Load()
    Command1.Caption = "16 Bit"
    Command2.Caption = "32 Bit"
    Command3.Caption = "64 Bit"
End Sub
