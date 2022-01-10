VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestInt32 
      Caption         =   "Test"
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Declare Sub GetMem4 Lib "msvbvm60" (ByRef pSrc As Any, ByRef pDst As Any)

Private m_ArrInt32() As Long

Private s As String
Private m_n As Long

'
Private Sub Form_Load()
    BtnTestInt32.Caption = "Test Int 32"
    
    m_n = 2000 - 1
    Dim mr As VbMsgBoxResult
    mr = MsgBox("1. die Arrays anlegen")
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
    Dim a() As Long: a = m_ArrInt32
    Dim x As Long, y As Long
    Dim t As Long

    Dim bs As TByteSwapper
    Call New_ByteSwapper(bs, LenB(a(0, 0)))
    
    t = GetTickCount
    For x = 0 To m_n
        For y = 0 To m_n
            bs.pB.pvData = VarPtr(a(x, y))
            MByteSwapper.Rotate4 bs
        Next
    Next
    t = GetTickCount - t
    s = s & "MByteSwapper.Rotate4: " & CStr(t) & "ms" & vbCrLf
    Call MByteSwapper.DeleteByteSwapper(bs)
    
    Text1.Text = s
    Text1.SelStart = Len(s)
End Sub

