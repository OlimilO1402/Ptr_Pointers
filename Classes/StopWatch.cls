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
'The datatype Currency in VB is a 64Bit-integer-type with a fixed point.
'The value 1 is defined at the integer value &H2710 (resp decimal 10000)
'The smallest possible positive value of the datatype Currency is 0.0001
'The unit ticks also contains the parameter 10000, and so Currency is
'the best datatype for using it with the QueryPerformance-API in VB
Private m_Frequency        As Currency
Private m_IsHighResolution As Boolean
Private m_IsRunning        As Boolean
Private m_StartTimeStamp   As Currency
Private m_Elapsed          As Currency
Private m_TickFrequency    As Double
'In der gleihnamigen VB.Net-Klasse wird hier der Datentyp Long,
'bzw Int64 verwendet.
'Der Korrekturfaktor für den Datentyp Currency
Private Const CurCorrect          As Long = 10000
Private Const TicksPerMillisecond As Long = 10000 / CurCorrect    '&H2710
Private Const TicksPerSecond      As Long = 10000000 / CurCorrect '&H989680
#If VBA7 = 0 Then
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLength As Long)
#Else
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bytLength As Long)
#End If

Private Sub Class_Initialize()
    If Not (QueryPerformanceFrequency(m_Frequency) = 1) Then
        m_IsHighResolution = False
        m_Frequency = TicksPerSecond
        m_TickFrequency = 1
    Else
        m_IsHighResolution = True
        m_TickFrequency = (TicksPerSecond / CDbl(m_Frequency))
    End If
End Sub

'Public Function GetTimestamp() As Currency 'Long 'LongInt
Public Function GetTimestamp() As Currency
    GetTimestamp = pGetTimestamp * CurCorrect
End Function
Private Function pGetTimestamp() As Currency
    If m_IsHighResolution Then
        Call QueryPerformanceCounter(pGetTimestamp)
    Else
        pGetTimestamp = DateTime.Now
    End If
End Function

Public Sub Start()
    If Not m_IsRunning Then
        m_StartTimeStamp = pGetTimestamp
        m_IsRunning = True
    End If
End Sub

Public Function StartNew() As StopWatch
    Set StartNew = New StopWatch
    Call StartNew.Start
End Function

Public Sub SStop()
    If IsRunning Then
        Dim c As Currency: c = (pGetTimestamp - m_StartTimeStamp)
        m_Elapsed = (m_Elapsed + c)
        m_IsRunning = False
    End If
End Sub

Public Sub Reset()
    m_Elapsed = 0
    m_IsRunning = False
    m_StartTimeStamp = 0
End Sub

'Shared-Member Properties
Public Property Get Frequency() As Currency
    'Diese Multiplikation mit CurCorrect ist nur dazu da, die Klasse
    'nach außen hin konsistent zu halten mit der gleichnamigen Klasse
    'aus dem .netFX
    Frequency = m_Frequency * CurCorrect
End Property
Public Property Get IsHighResolution() As Boolean
    IsHighResolution = m_IsHighResolution
End Property

' #################### '   Public Properties   ' #################### '
Public Property Get ElapsedToString() As String 'TimeSpan
    ElapsedToString = TimeSpan_ToString(GetElapsedDateTimeTicks)
End Property
Private Function TimeSpan_ToString(ByVal ticks As Currency) As String
    Dim b As String
    Dim h As Long, m As Long, s As Long, n As Long
    Dim days As Long:     days = CInt((ticks / 86400000))
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
    TimeSpan_ToString = b
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
    IsRunning = m_IsRunning
End Property

' #################### '   Private Functions   ' #################### '
Private Function GetElapsedDateTimeTicks() As Currency
    Dim rawElapsedTicks As Currency: rawElapsedTicks = GetRawElapsedTicks
    If m_IsHighResolution Then
        Dim d As Double: d = CDbl(rawElapsedTicks)
        d = (d * m_TickFrequency)
        GetElapsedDateTimeTicks = CCur(d)
    Else
        GetElapsedDateTimeTicks = rawElapsedTicks
    End If
End Function

Private Function GetRawElapsedTicks() As Currency
    Dim Elapsed As Currency: Elapsed = m_Elapsed
    If m_IsRunning Then
        Dim c As Currency: c = (GetTimestamp - m_StartTimeStamp)
        Elapsed = (Elapsed + c)
    End If
    GetRawElapsedTicks = Elapsed
End Function

