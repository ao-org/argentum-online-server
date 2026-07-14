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
Option Explicit

Private Const GROWTH_BLOCK_SIZE As Long = 50000

Private m_sRealString   As String
Private m_lRealLen      As Long
Private m_lSimulatedLen As Long
Private m_bCopyUsedOnly As Boolean

Public Property Get Length() As Long
    Length = m_lSimulatedLen
End Property

Public Property Get Capacity() As Long
    Capacity = m_lRealLen
End Property

Public Property Get CopyUsedOnly() As Boolean
    CopyUsedOnly = m_bCopyUsedOnly
End Property

Public Property Let CopyUsedOnly(ByVal value As Boolean)
    m_bCopyUsedOnly = value
End Property

Public Property Get ToString() As String
    If m_lSimulatedLen > 0 Then
        ToString = Left$(m_sRealString, m_lSimulatedLen)
    End If
End Property

Public Sub Clear()
    m_lSimulatedLen = 0
End Sub

Public Sub AppendNL(ByRef sThis As String)
    Append sThis
    Append vbCrLf
End Sub

Public Sub Append(ByRef sThis As String)
    Dim lNewLen As Long
    Dim sTemp   As String

    lNewLen = m_lSimulatedLen + Len(sThis)

    If lNewLen > m_lRealLen Then
        m_lRealLen = lNewLen + GROWTH_BLOCK_SIZE
        If m_bCopyUsedOnly Then
            sTemp = Left$(m_sRealString, m_lSimulatedLen)
        Else
            sTemp = m_sRealString
        End If
        m_sRealString = Space$(m_lRealLen)

        If Len(sTemp) > 0 Then
            Mid$(m_sRealString, 1, Len(sTemp)) = sTemp
        End If
    End If

    If Len(sThis) > 0 Then
        Mid$(m_sRealString, m_lSimulatedLen + 1, Len(sThis)) = sThis
    End If

    m_lSimulatedLen = lNewLen
End Sub
