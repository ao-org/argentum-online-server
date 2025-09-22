VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFastString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFastString
' Author    : [/About]
' Date      : 23/07/2019
'---------------------------------------------------------------------------------------

Option Explicit

Private m_lRealLen              As Long
Private m_lSimulatedLen         As Long
Private m_sRealString           As String

Public Property Get ToString() As String

10  ToString = Left$(m_sRealString, m_lSimulatedLen)

End Property

Public Sub Append(sText As String)

Dim lLen                        As Long
Dim sTemp                       As String
Dim lNewLen                     As Long

    lLen = Len(sText)
    lNewLen = lLen + m_lSimulatedLen

    If lNewLen > m_lRealLen Then
        m_lRealLen = lNewLen + 50000
        sTemp = m_sRealString
        m_sRealString = Space$(m_lRealLen)
        Mid$(m_sRealString, 1, Len(sTemp)) = sTemp
    End If

    Mid$(m_sRealString, m_lSimulatedLen + 1, lLen) = sText

    m_lSimulatedLen = lNewLen

End Sub

Public Sub Clear()
    m_lSimulatedLen = 0
End Sub

