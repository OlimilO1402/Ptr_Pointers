VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Test CharPointer"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnWalkBArrUnic 
      Caption         =   "Byte Array Walk Unicode"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton BtnWalkBArrAnsi 
      Caption         =   "Byte Array Walk ANSI"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CommandButton BtnCharPtrWalk 
      Caption         =   "CharPtr Walk"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton BtnMidBWalk 
      Caption         =   "MidB Walk"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Frame FrmBuildString 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      Begin VB.OptionButton Option3 
         Caption         =   "80 Mb"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "8 Mb"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "800 Kb"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Run it compiled to native-exe!"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Beim Parsen eines Strings kommt es des Öfteren vor, daß ein String
'zeichenweise durchwandert werden soll. Mid und MidB ist hier die meist
'verwendete Möglichkeit, Teile eines Strings zu extrahieren und in einen
'anderen String zu kopieren.
'Soll dann eine Entscheidung anhand des rausgelesen Zeichens getroffen
'werden, bspw in einem Select case, so bietet der Select Case von VB die
'bequeme Möglichkeit selbst Strings als Argument übergeben zu dürfen.
'Diese bequeme Vorgehensweise erkauft man sich allerdings mit einer guten
'Portion Performance.
'Wer beim Parsen größerer Dateien auf Performance achten möchte, der greift
'ein wenig in die Trickkiste. Für den gewöhnlichen VB-Anwender stellt ein
'String einen besonderen Datentyp dar, der mit den in VB enthaltenen String-
'funktionen bearbeitet werden kann.
'Eine andere Sichtweise ist es, einen String als einen zusammenhängenden
'Speicherbereich bzw als ein Array von Zeichen (Character) zu betrachten.
'Es gibt jedoch in VBC keinen eigenen Datentyp für ein Zeichen, weshalb der
'Datentyp String meist für ein einzelnes Zeichen verwendet wird.
'Ein einzelnes Zeichen kann aber auch als ein 2-Byte langer Integer btrachtet
'werden. Ein Select Case mit einem numerischen Wert als Argument bringt
'ebenfalls einen Performancegewinn.
'Ein Array als universal verwendbaren Zeiger ist der Trick der Wahl.
'Dabei wird einer Arrayvariable der Zeiger auf einen selbstdefinierten
'SafeArraydescriptor injiziert.
'In anderen Programmiersprachen (Delphi, C++) ist diese Vorgehensweise seit
'jeher Bestandteil der Sprache, weswegen Parser in anderen Programmiersprachen
'meist viel zügiger als in VB zu Werke gehen.
'In diesem kleinen BspProjekt wird eine Performancemessung der gezeigten
'Methode im Vergleich zu Mid / MidB durchgeführt.
'Dabei ergibt sich gegenüber MidB in der kompilierten Exedatei ein etwa
'20-facher, in der VB-IDE immerhin noch 3-4facher Performancegewinn.

'Wohlgemerkt wird hier lediglich der Vorgang des Extrahierens eines einzelnen
'Zeichens betrachtet, wohingegen sich im gesamten Parse-Vorgang einer Datei,
'bzw überführen eines Strings in eine Objektdatenstruktur sich der Performance-
'gewinn allein durch die hier gezeigte Methode, selbstverständlich nicht so
'gravierend bemerkbar machen wird.

'Der String der durchwandert werden soll
Private mStrVal As String
Private mSW As StopWatch


Private Sub Form_Load()
    Set mSW = New StopWatch
End Sub

Private Sub Option1_Click()
    Call HourGlassBuildString
End Sub
Private Sub Option2_Click()
    Call HourGlassBuildString
End Sub
Private Sub Option3_Click()
    Call HourGlassBuildString
End Sub

'Dieser Programmteil dient lediglich zum Aufbau verschieden langer Strings
'die dann durchwandert werden.
Private Sub HourGlassBuildString()
    Dim mp As MousePointerConstants
    mp = Screen.MousePointer
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    Call BuildString
    Screen.MousePointer = mp
    Call MessStringLength
End Sub
Private Sub BuildString()

    mStrVal = vbNullString
    Dim s As String
    s = "quick brown fox jumps over the lazy dog " '40 = 80Byte
    Call AppendStringN(mStrVal, s, 1000)
    
    s = mStrVal
    Call AppendStringN(mStrVal, s, 10 - 1)
    If Option1.Value Then Exit Sub
    
    s = mStrVal
    Call AppendStringN(mStrVal, s, 10 - 1)
    If Option2.Value Then Exit Sub
    
    s = mStrVal
    Call AppendStringN(mStrVal, s, 10 - 1)
    If Option3.Value Then Exit Sub

End Sub
Private Sub AppendStringN(AppendTo As String, StrVal As String, ByVal n As Long)
    Dim i As Long, lA As Long, lS As Long
    lA = LenB(AppendTo)
    lS = LenB(StrVal)
    AppendTo = AppendTo & Space$(n * lS \ 2)
    For i = lA + 1 To lA + (n * lS) Step lS
        MidB$(AppendTo, i, lS) = StrVal
    Next
End Sub
Private Sub MessStringLength()
    Dim b As Double, c As String
    b = LenB(mStrVal)
    c = IIf(b > 1000# * 1000#, "MB", "KB")
    b = IIf(b > 1000# * 1000#, b / 1000# / 1000#, b / 1000#)
    MsgBox CStr((b)) & " " & c
End Sub

Private Sub Start(amp As MousePointerConstants)
    amp = Screen.MousePointer
    Screen.MousePointer = MousePointerConstants.vbArrowHourglass
    mSW.Reset
    mSW.Start
End Sub
Private Sub MessStop(mp As MousePointerConstants)
    mSW.SStop
    MsgBox CStr(mSW.ElapsedMilliseconds) & " ms"
    Screen.MousePointer = mp
End Sub

'hier gehts los mit den Vergleichsmessungen
Private Sub BtnMidBWalk_Click()
    
    If LenB(mStrVal) = 0 Then Call BuildString
    
    Dim mp As MousePointerConstants
    Call Start(mp)
    
    Dim i As Long
    Dim c As String
    For i = 1 To LenB(mStrVal) Step 2
        c = MidB$(mStrVal, i, 2)
    Next
    
    Call MessStop(mp)
    
End Sub

Private Sub BtnCharPtrWalk_Click()
    
    If LenB(mStrVal) = 0 Then Call BuildString
    
    Dim mp As MousePointerConstants
    Call Start(mp)
    
    Dim i As Long
    Dim c As Integer
    Dim cp As TCharPointer: Call New_CharPointer(cp, mStrVal)
    For i = 1 To Len(mStrVal)
        c = cp.Chars(i)
    Next
    Call DeleteCharPointer(cp)
    
    Call MessStop(mp)
    
End Sub

Private Sub BtnWalkBArrAnsi_Click()
    
    If LenB(mStrVal) = 0 Then Call BuildString
    
    Dim mp As MousePointerConstants
    Call Start(mp)
    
    Dim i As Long
    Dim c As Byte
    Dim bArray() As Byte
    bArray = StrConv(mStrVal, vbFromUnicode)
    For i = 0 To Len(mStrVal) - 1
        c = bArray(i)
    Next
    
    Call MessStop(mp)

End Sub

Private Sub BtnWalkBArrUnic_Click()
    
    If LenB(mStrVal) = 0 Then Call BuildString
    
    Dim mp As MousePointerConstants
    Call Start(mp)
    
    Dim i As Long
    Dim c As Byte
    Dim bArray() As Byte
    bArray = StrConv(mStrVal, vbUnicode)
    For i = 0 To UBound(bArray) Step 2
        c = bArray(i)
    Next
    
    Call MessStop(mp)

End Sub


