VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class StopWatch
'Namespace System.Diagnostics
'Der Datentyp Currency in VB ist ein 64Bit-Datentyp mit einem
'Festpunktanteil. D.h. beim Bitwert &H2710 (bzw decimal 10000)
'ist der Wert 1 definiert, bzw. der kleinste positive Wert des
'Datentyps Currency ist 0.0001
'Da die Einheit Ticks ebenfalls den Parameter 10000 beinhaltet,
'ist der VB-Datentyp Currency genau der richtige Datentyp für
'Zeitmessungen mit der QueryPerformance-API
Private mFrequency        As Currency
Private mIsHighResolution As Boolean
Private mIsRunning        As Boolean
Private mStartTimeStamp   As Currency
Private mElapsed          As Currency
Private mTickFrequency    As Double
'In der gleihnamigen VB.Net-Klasse wird hier der Datentyp Long,
'bzw Int64 verwendet.
'Der Korrekturfaktor für den Datentyp Currency
Private Const CurCorrect          As Long = 10000
Private Const TicksPerMillisecond As Long = 10000 / CurCorrect    '&H2710
Private Const TicksPerSecond      As Long = 10000000 / CurCorrect '&H989680
Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    ByRef lpFrequency As Currency) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
    ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLength As Long)
    
Private Sub Class_Initialize()
    If Not (QueryPerformanceFrequency(mFrequency) = 1) Then
        mIsHighResolution = False
        mFrequency = TicksPerSecond
        mTickFrequency = 1
    Else
        mIsHighResolution = True
        mTickFrequency = (TicksPerSecond / CDbl(mFrequency))
    End If
End Sub

'Public Function GetTimestamp() As Currency 'Long 'LongInt
Public Function GetTimestamp() As Currency
    GetTimestamp = pGetTimestamp * CurCorrect
End Function
Private Function pGetTimestamp() As Currency
    If mIsHighResolution Then
        Call QueryPerformanceCounter(pGetTimestamp)
    Else
        pGetTimestamp = DateTime.Now
    End If
End Function

Public Sub Start()
    If Not mIsRunning Then
        mStartTimeStamp = pGetTimestamp
        mIsRunning = True
    End If
End Sub

Public Function StartNew() As StopWatch
    Set StartNew = New StopWatch
    Call StartNew.Start
End Function

Public Sub SStop()
    If IsRunning Then
        Dim c As Currency: c = (pGetTimestamp - mStartTimeStamp)
        mElapsed = (mElapsed + c)
        mIsRunning = False
    End If
End Sub

Public Sub Reset()
    mElapsed = 0
    mIsRunning = False
    mStartTimeStamp = 0
End Sub


'Shared-Member Properties
Public Property Get Frequency() As Currency
    'Diese Multiplikation mit CurCorrect ist nur dazu da, die Klasse
    'nach außen hin konsistent zu halten mit der gleichnamigen Klasse
    'aus dem .netFX
    Frequency = mFrequency * CurCorrect
End Property
Public Property Get IsHighResolution() As Boolean
    IsHighResolution = mIsHighResolution
End Property

' #################### '   Public Properties   ' #################### '
Public Property Get ElapsedToString() As String 'TimeSpan
    ElapsedToString = TimeSpanToString(GetElapsedDateTimeTicks)
End Property
Private Function TimeSpanToString(ByVal ticks As Currency) As String
    Dim b As String
    Dim h As Long, m As Long, s As Long, n As Long
    Dim days As Long:  days = CInt((ticks / 86400000))
    Dim time As Currency: time = (ticks Mod 86400000)
    If (ticks < 0) Then
        b = b & "-"
        days = -days
        time = -time
    End If
    If (days <> 0) Then
        b = b & CStr(days) & "."
    End If
    
    'Stunden
    h = CInt(((time \ 3600000) Mod 24))
    b = b & IntToString(h, 2) & ":"
    
    'Minuten
    m = CInt(((time \ 60000) Mod 60))
    b = b & IntToString(m, 2) & ":"
    
    'Sekunden
    s = CInt(((time \ 1000) Mod 60))
    b = b & IntToString(s, 2)
    
    n = (ticks - (CCur(h) * CCur(3600000)) _
               - (CCur(m) * CCur(60000)) _
               - (CCur(s) * CCur(1000))) * 10000
               
    If (n <> 0) Then
        b = b & "."
        b = b & IntToString(n, 7)
    End If
    TimeSpanToString = b
End Function
Private Function IntToString(ByVal n As Long, ByVal digits As Long) As String
    Dim l As Long
    IntToString = CStr(n)
    l = Len(IntToString)
    If l < digits Then IntToString = String$(digits - l, "0") & IntToString
End Function

Public Property Get ElapsedMilliseconds() As Currency 'Long 'LongInt
    ElapsedMilliseconds = GetElapsedDateTimeTicks / TicksPerMillisecond
End Property

Public Property Get ElapsedTicks() As Currency 'As Long 'LongInt
    ElapsedTicks = GetRawElapsedTicks * CurCorrect
End Property

Public Property Get IsRunning() As Boolean
    IsRunning = mIsRunning
End Property

' #################### '   Private Functions   ' #################### '
Private Function GetElapsedDateTimeTicks() As Currency
    Dim rawElapsedTicks As Currency: rawElapsedTicks = GetRawElapsedTicks
    If mIsHighResolution Then
        Dim d As Double: d = CDbl(rawElapsedTicks)
        d = (d * mTickFrequency)
        GetElapsedDateTimeTicks = CCur(d)
    Else
        GetElapsedDateTimeTicks = rawElapsedTicks
    End If
End Function

Private Function GetRawElapsedTicks() As Currency
    Dim Elapsed As Currency: Elapsed = mElapsed
    If mIsRunning Then
        Dim c As Currency: c = (GetTimestamp - mStartTimeStamp)
        Elapsed = (Elapsed + c)
    End If
    GetRawElapsedTicks = Elapsed
End Function


