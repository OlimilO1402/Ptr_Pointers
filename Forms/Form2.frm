VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form2"
   ScaleHeight     =   3615
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestInt32 
      Caption         =   "Test"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   480
      Width           =   8655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private m_ArrInt32() As Long

Private s As String
Private m_n As Long

Private Sub Form_Load()
    BtnTestInt32.Caption = "Test Int 32"
    m_n = 2000 - 1
    Dim mr As VbMsgBoxResult
    mr = MsgBox("1. create Arrays of size 16MB")
    ReDim m_ArrInt32(0 To m_n, 0 To m_n) ' 16 MByte
    MsgBox "2. die Arrays mit Randomwerten füllen"
    Dim x As Long, y As Long
    For x = 0 To m_n
        For y = 0 To m_n
            m_ArrInt32(x, y) = CLng(Rnd * 2147483647)
        Next
    Next
    
End Sub

Private Sub BtnTestInt32_Click()
    Dim A() As Long: A = m_ArrInt32
    Dim x As Long, y As Long
    Dim t As Single
    
    Dim bs As TByteSwapper
    Call New_ByteSwapper(bs, LenB(A(0, 0)))
    
    't = Timer 'GetTickCount
    Dim sw As New StopWatch
    sw.Start
    For x = 0 To m_n
        For y = 0 To m_n
            bs.pB.pvData = VarPtr(A(x, y))
            MByteSwapper.Rotate4 bs
        Next
    Next
    sw.SStop
    't = Timer - t
    s = s & "MByteSwapper.Rotate4: Array of size 16MB " & sw.ElapsedToString & "s" & vbCrLf
    Call MByteSwapper.DeleteByteSwapper(bs)
    
    Text1.Text = s
    Text1.SelStart = Len(s)
End Sub

