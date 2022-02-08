VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Instruments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "Kernel32" (X As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "Kernel32" (X As Currency) As Boolean

Private m_startTime As Currency
Private m_freq As Currency
Private m_overhead As Currency

Public Sub start()
    QueryPerformanceCounter m_startTime
End Sub

Public Function ElapsedSeconds() As Double
    Dim currentTime As Currency
    QueryPerformanceCounter currentTime
    ElapsedSeconds = (currentTime - m_startTime - m_overhead) / m_freq
End Function

Public Function ElapsedMilliseconds() As Double
    ElapsedMilliseconds = ElapsedSeconds * 1000
End Function

Private Sub Class_Initialize()
    QueryPerformanceFrequency m_freq
    Dim ctr1 As Currency
    Dim ctr2 As Currency
    QueryPerformanceCounter ctr1
    QueryPerformanceCounter ctr2
    m_overhead = ctr2 - ctr1
End Sub

