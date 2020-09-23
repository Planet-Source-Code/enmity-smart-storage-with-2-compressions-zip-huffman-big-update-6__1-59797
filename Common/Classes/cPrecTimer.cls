VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrecTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class was posted here:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=31266&lngWId=1
Option Explicit

Private m_curFreq As Currency
Private m_curStart As Currency
Private m_curStop As Currency
Private m_dblLatency As Double
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long



Private Sub Class_Initialize()

    Dim i As Long
    Dim lngTotal As Double
    
    If QueryPerformanceFrequency(m_curFreq) = 0 Then
    Else
        For i = 1 To 1000
            ResetTimer
            StopTimer
            lngTotal = lngTotal + Elapsed
        Next i
        m_dblLatency = lngTotal / i
    End If
    
End Sub


Public Sub ResetTimer()
    QueryPerformanceCounter m_curStart
End Sub


Public Sub StopTimer()
    QueryPerformanceCounter m_curStop
End Sub


Public Function Elapsed() As Double
    Elapsed = (10000 * (m_curStop - m_curStart)) - m_dblLatency
End Function


Public Function ToString() As String
    ToString = Format$((10000 * (m_curStop - m_curStart)) - m_dblLatency, "###,###.00") & Space$(1) & Chr$(181) & "s"
End Function

