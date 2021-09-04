VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================
' Name:     vbAccelerator cStringBuilder
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     1 January 2002
'
' Copyright © 2002 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' VB can be slow to append strings together because of the continual
' reallocation of string size.  This class pre-allocates a string in
' blocks and hence removes the performance restriction.
'
' Quicker insert and remove is also possible since string space does
' not have to be reallocated.
'
' Example:
' Adding "http://vbaccelerator.com/" 10,000 times to a string:
' Standard VB:   34s
' This Class:    0.35s
'
' ======================================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
      
Private m_sString    As String
Private m_iChunkSize As Long
Private m_iPos       As Long
Private m_iLen       As Long

Public Property Get Length() As Long
    Length = m_iPos \ 2
End Property

Public Property Get ByteLength() As Long
    ByteLength = m_iPos
End Property

Public Property Get Capacity() As Long
    Capacity = m_iLen \ 2
End Property

Public Property Get ChunkSize() As Long
        ' Return the unicode character chunk size:

100     ChunkSize = m_iChunkSize \ 2

End Property

Public Property Let ChunkSize(ByVal iChunkSize As Long)
        ' Set the chunksize.  We multiply by 2 because internally
        ' we are considering bytes:

100     m_iChunkSize = iChunkSize * 2

End Property

Public Property Get ToString() As String

        ' The internal string:
100     If m_iPos > 0 Then
102         ToString = Left$(m_sString, m_iPos \ 2)

        End If

End Property

Public Property Let TheString(ByRef sThis As String)

        Dim lLen As Long
   
        ' Setting the string:
100     lLen = LenB(sThis)

102     If lLen = 0 Then
            'Clear
104         m_sString = ""
106         m_iPos = 0
108         m_iLen = 0
        Else

110         If m_iLen < lLen Then
                ' Need to expand string to accommodate:
                Do
112                 m_sString = m_sString & Space$(m_iChunkSize \ 2)
114                 m_iLen = m_iLen + m_iChunkSize
116             Loop While m_iLen < lLen

            End If

118         CopyMemory ByVal StrPtr(m_sString), ByVal StrPtr(sThis), lLen
120         m_iPos = lLen

        End If

End Property

Public Sub Clear()

100     m_sString = ""
102     m_iPos = 0
104     m_iLen = 0

End Sub

Public Sub AppendNL(ByRef sThis As String)
        
        
100     Append sThis
102     Append vbCrLf

End Sub

Public Sub Append(ByRef sThis As String)

        Dim lLen        As Long

        Dim lLenPlusPos As Long
 
        ' Append an item to the string:
100     lLen = LenB(sThis)
102     lLenPlusPos = lLen + m_iPos

104     If lLenPlusPos > m_iLen Then

            Dim lTemp As Long
      
106         lTemp = m_iLen

108         Do While lTemp < lLenPlusPos
110             lTemp = lTemp + m_iChunkSize
            Loop
      
112         m_sString = m_sString & Space$((lTemp - m_iLen) \ 2)
114         m_iLen = lTemp

        End If
   
116     CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
118     m_iPos = m_iPos + lLen

End Sub

Public Sub AppendByVal(ByVal sThis As String)

100     Append sThis

End Sub

Public Sub Insert(ByVal iIndex As Long, ByRef sThis As String)

        Dim lLen  As Long

        Dim lPos  As Long

        Dim lSize As Long
   
        ' is iIndex within bounds?
100     If (iIndex * 2 > m_iPos) Then
102         Err.raise 9
        Else
   
104         lLen = LenB(sThis)

106         If (m_iPos + lLen) > m_iLen Then
108             m_sString = m_sString & Space$(m_iChunkSize \ 2)
110             m_iLen = m_iLen + m_iChunkSize

            End If
      
            ' Move existing characters from current position
112         lPos = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
114         lSize = m_iPos - iIndex * 2
      
            ' moving from iIndex to iIndex + lLen
116         CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize
      
            ' Insert new characters:
118         CopyMemory ByVal lPos, ByVal StrPtr(sThis), lLen
      
120         m_iPos = m_iPos + lLen

        End If

End Sub

Public Sub InsertByVal(ByVal iIndex As Long, ByVal sThis As String)
        
100     Insert iIndex, sThis


End Sub

Public Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)

        Dim lSrc  As Long

        Dim lDst  As Long

        Dim lSize As Long

        ' is iIndex within bounds?
100     If (iIndex * 2 > m_iPos) Then
102         Err.raise 9
        Else

            ' is there sufficient length?
104         If ((iIndex + lLen) * 2 > m_iPos) Then
106             Err.raise 9
            Else
                ' Need to copy characters from iIndex*2 to m_iPos back by lLen chars:
108             lSrc = UnsignedAdd(StrPtr(m_sString), (iIndex + lLen) * 2)
110             lDst = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
112             lSize = (m_iPos - (iIndex + lLen) * 2)
114             CopyMemory ByVal lDst, ByVal lSrc, lSize
116             m_iPos = m_iPos - lLen * 2

            End If

        End If

End Sub

Public Function Find(ByVal sToFind As String, _
                     Optional ByVal lStartIndex As Long = 1, _
                     Optional ByVal compare As VbCompareMethod = vbTextCompare) As Long

        Dim lInstr As Long

100     If (lStartIndex > 0) Then
102         lInstr = InStr(lStartIndex, m_sString, sToFind, compare)
        Else
104         lInstr = InStr(m_sString, sToFind, compare)

        End If

106     If (lInstr <= m_iPos \ 2) Then
108         Find = lInstr

        End If
End Function

Public Function SubStr(ByVal Index As Long, ByVal Length As Long) As String
    SubStr = Space$(Length)
    
    ' is iIndex within bounds?
    If (Index * 2 > m_iPos) Then
        Err.raise 9
    Else
    
        ' is there sufficient length?
        If ((Index + Length) * 2 > m_iPos) Then
            Err.raise 9
        Else
            Dim lStr As Long, Dest As Long
            lStr = UnsignedAdd(StrPtr(m_sString), Index * 2)
            Dest = StrPtr(SubStr)

            CopyMemory ByVal Dest, ByVal lStr, Length * 2
        End If
    
    End If
    
End Function

Public Sub HeapMinimize()

        Dim iLen As Long
   
        ' Reduce the string size so only the minimal chunks
        ' are allocated:
100     If (m_iLen - m_iPos) > m_iChunkSize Then
102         iLen = m_iLen

104         Do While (iLen - m_iPos) > m_iChunkSize
106             iLen = iLen - m_iChunkSize
            Loop
108         m_sString = Left$(m_sString, iLen \ 2)
110         m_iLen = iLen

        End If

End Sub

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
        ' This function is useful when doing pointer arithmetic,
        ' but note it only works for positive values of Incr
        

100     If Start And &H80000000 Then 'Start < 0
102         UnsignedAdd = Start + Incr
104     ElseIf (Start Or &H80000000) < -Incr Then
106         UnsignedAdd = Start + Incr
        Else
108         UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)

        End If
   
End Function

Private Sub Class_Initialize()
        ' The default allocation: 8192 characters.
        
        m_iChunkSize = 16384
   
End Sub

