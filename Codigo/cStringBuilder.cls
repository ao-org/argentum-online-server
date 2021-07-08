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
        
        On Error GoTo Length_Err
    
        
100     Length = m_iPos \ 2

        
        Exit Property

Length_Err:
102     Call TraceError(Err.Number, Err.Description, "cStringBuilder.Length", Erl)

        
End Property

Public Property Get ByteLength() As Long
    ByteLength = m_iPos
End Property

Public Property Get Capacity() As Long
        
        On Error GoTo Capacity_Err
    
        
100     Capacity = m_iLen \ 2

        
        Exit Property

Capacity_Err:
102     Call TraceError(Err.Number, Err.Description, "cStringBuilder.Capacity", Erl)

        
End Property

Public Property Get ChunkSize() As Long
        ' Return the unicode character chunk size:
        
        On Error GoTo ChunkSize_Err
    
        
100     ChunkSize = m_iChunkSize \ 2

        
        Exit Property

ChunkSize_Err:
102     Call TraceError(Err.Number, Err.Description, "cStringBuilder.ChunkSize", Erl)

        
End Property

Public Property Let ChunkSize(ByVal iChunkSize As Long)
        ' Set the chunksize.  We multiply by 2 because internally
        ' we are considering bytes:
        
        On Error GoTo ChunkSize_Err
    
        
100     m_iChunkSize = iChunkSize * 2

        
        Exit Property

ChunkSize_Err:
102     Call TraceError(Err.Number, Err.Description, "cStringBuilder.ChunkSize", Erl)

        
End Property

Public Property Get ToString() As String
        
        On Error GoTo toString_Err
    
        

        ' The internal string:
100     If m_iPos > 0 Then
102         ToString = Left$(m_sString, m_iPos \ 2)

        End If

        
        Exit Property

toString_Err:
104     Call TraceError(Err.Number, Err.Description, "cStringBuilder.toString", Erl)

        
End Property

Public Property Let TheString(ByRef sThis As String)
        
        On Error GoTo TheString_Err
    
        

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
   
        
        Exit Property

TheString_Err:
122     Call TraceError(Err.Number, Err.Description, "cStringBuilder.TheString", Erl)

        
End Property

Public Sub Clear()
        
        On Error GoTo Clear_Err
    
        
100     m_sString = ""
102     m_iPos = 0
104     m_iLen = 0

        
        Exit Sub

Clear_Err:
106     Call TraceError(Err.Number, Err.Description, "cStringBuilder.Clear", Erl)

        
End Sub

Public Sub AppendNL(ByRef sThis As String)
        
        On Error GoTo AppendNL_Err
    
        
100     Append sThis
102     Append vbCrLf

        
        Exit Sub

AppendNL_Err:
104     Call TraceError(Err.Number, Err.Description, "cStringBuilder.AppendNL", Erl)

        
End Sub

Public Sub Append(ByRef sThis As String)
        
        On Error GoTo Append_Err
    
        

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

        
        Exit Sub

Append_Err:
120     Call TraceError(Err.Number, Err.Description, "cStringBuilder.Append", Erl)

        
End Sub

Public Sub AppendByVal(ByVal sThis As String)
        
        On Error GoTo AppendByVal_Err
    
        
100     Append sThis

        
        Exit Sub

AppendByVal_Err:
102     Call TraceError(Err.Number, Err.Description, "cStringBuilder.AppendByVal", Erl)

        
End Sub

Public Sub Insert(ByVal iIndex As Long, ByRef sThis As String)
        
        On Error GoTo Insert_Err
    
        

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

        
        Exit Sub

Insert_Err:
122     Call TraceError(Err.Number, Err.Description, "cStringBuilder.Insert", Erl)

        
End Sub

Public Sub InsertByVal(ByVal iIndex As Long, ByVal sThis As String)
        
        On Error GoTo InsertByVal_Err
    
        
100     Insert iIndex, sThis

        
        Exit Sub

InsertByVal_Err:
102     Call TraceError(Err.Number, Err.Description, "cStringBuilder.InsertByVal", Erl)

        
End Sub

Public Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)
        
        On Error GoTo Remove_Err
    
        

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

        
        Exit Sub

Remove_Err:
118     Call TraceError(Err.Number, Err.Description, "cStringBuilder.Remove", Erl)

        
End Sub

Public Function Find(ByVal sToFind As String, _
                     Optional ByVal lStartIndex As Long = 1, _
                     Optional ByVal compare As VbCompareMethod = vbTextCompare) As Long
        
        On Error GoTo Find_Err
    
        
   
        Dim lInstr As Long

100     If (lStartIndex > 0) Then
102         lInstr = InStr(lStartIndex, m_sString, sToFind, compare)
        Else
104         lInstr = InStr(m_sString, sToFind, compare)

        End If

106     If (lInstr < m_iPos \ 2) Then
108         Find = lInstr

        End If

        
        Exit Function

Find_Err:
110     Call TraceError(Err.Number, Err.Description, "cStringBuilder.Find", Erl)

        
End Function

Public Sub HeapMinimize()
        
        On Error GoTo HeapMinimize_Err
    
        

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
   
        
        Exit Sub

HeapMinimize_Err:
112     Call TraceError(Err.Number, Err.Description, "cStringBuilder.HeapMinimize", Erl)

        
End Sub

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
        ' This function is useful when doing pointer arithmetic,
        ' but note it only works for positive values of Incr
        
        On Error GoTo UnsignedAdd_Err
    
        

100     If Start And &H80000000 Then 'Start < 0
102         UnsignedAdd = Start + Incr
104     ElseIf (Start Or &H80000000) < -Incr Then
106         UnsignedAdd = Start + Incr
        Else
108         UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)

        End If
   
        
        Exit Function

UnsignedAdd_Err:
110     Call TraceError(Err.Number, Err.Description, "cStringBuilder.UnsignedAdd", Erl)

        
End Function

Private Sub Class_Initialize()
        ' The default allocation: 8192 characters.
        
        On Error GoTo Class_Initialize_Err
    
        
100     m_iChunkSize = 16384

        
        Exit Sub

Class_Initialize_Err:
102     Call TraceError(Err.Number, Err.Description, "cStringBuilder.Class_Initialize", Erl)

        
End Sub

