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
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
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
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private m_sString    As String
Private m_iChunkSize As Long
Private m_iPos       As Long
Private m_iLen       As Long

Public Property Get length() As Long
    On Error GoTo Length_Err
    length = m_iPos \ 2
    Exit Property
Length_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.Length", Erl)
End Property

Public Property Get ByteLength() As Long
    ByteLength = m_iPos
End Property

Public Property Get Capacity() As Long
    On Error GoTo Capacity_Err
    Capacity = m_iLen \ 2
    Exit Property
Capacity_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.Capacity", Erl)
End Property

Public Property Get ChunkSize() As Long
    ' Return the unicode character chunk size:
    On Error GoTo ChunkSize_Err
    ChunkSize = m_iChunkSize \ 2
    Exit Property
ChunkSize_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.ChunkSize", Erl)
End Property

Public Property Let ChunkSize(ByVal iChunkSize As Long)
    ' Set the chunksize.  We multiply by 2 because internally
    ' we are considering bytes:
    On Error GoTo ChunkSize_Err
    m_iChunkSize = iChunkSize * 2
    Exit Property
ChunkSize_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.ChunkSize", Erl)
End Property

Public Property Get ToString() As String
    On Error GoTo toString_Err
    ' The internal string:
    If m_iPos > 0 Then
        ToString = Left$(m_sString, m_iPos \ 2)
    End If
    Exit Property
toString_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.toString", Erl)
End Property

Public Property Let TheString(ByRef sThis As String)
    On Error GoTo TheString_Err
    Dim lLen As Long
    ' Setting the string:
    lLen = LenB(sThis)
    If lLen = 0 Then
        'Clear
        m_sString = ""
        m_iPos = 0
        m_iLen = 0
    Else
        If m_iLen < lLen Then
            ' Need to expand string to accommodate:
            Do
                m_sString = m_sString & Space$(m_iChunkSize \ 2)
                m_iLen = m_iLen + m_iChunkSize
            Loop While m_iLen < lLen
        End If
        CopyMemory ByVal StrPtr(m_sString), ByVal StrPtr(sThis), lLen
        m_iPos = lLen
    End If
    Exit Property
TheString_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.TheString", Erl)
End Property

Public Sub Clear()
    On Error GoTo Clear_Err
    m_sString = ""
    m_iPos = 0
    m_iLen = 0
    Exit Sub
Clear_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.Clear", Erl)
End Sub

Public Sub AppendNL(ByRef sThis As String)
    On Error GoTo AppendNL_Err
    Append sThis
    Append vbCrLf
    Exit Sub
AppendNL_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.AppendNL", Erl)
End Sub

Public Sub Append(ByRef sThis As String)
    On Error GoTo Append_Err
    Dim lLen        As Long
    Dim lLenPlusPos As Long
    ' Append an item to the string:
    lLen = LenB(sThis)
    lLenPlusPos = lLen + m_iPos
    If lLenPlusPos > m_iLen Then
        Dim lTemp As Long
        lTemp = m_iLen
        Do While lTemp < lLenPlusPos
            lTemp = lTemp + m_iChunkSize
        Loop
        m_sString = m_sString & Space$((lTemp - m_iLen) \ 2)
        m_iLen = lTemp
    End If
    CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
    m_iPos = m_iPos + lLen
    Exit Sub
Append_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.Append", Erl)
End Sub

Public Sub AppendByVal(ByVal sThis As String)
    On Error GoTo AppendByVal_Err
    Append sThis
    Exit Sub
AppendByVal_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.AppendByVal", Erl)
End Sub

Public Sub Insert(ByVal iIndex As Long, ByRef sThis As String)
    On Error GoTo Insert_Err
    Dim lLen  As Long
    Dim lPos  As Long
    Dim lSize As Long
    ' is iIndex within bounds?
    If (iIndex * 2 > m_iPos) Then
        Err.Raise 9
    Else
        lLen = LenB(sThis)
        If (m_iPos + lLen) > m_iLen Then
            m_sString = m_sString & Space$(m_iChunkSize \ 2)
            m_iLen = m_iLen + m_iChunkSize
        End If
        ' Move existing characters from current position
        lPos = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
        lSize = m_iPos - iIndex * 2
        ' moving from iIndex to iIndex + lLen
        CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize
        ' Insert new characters:
        CopyMemory ByVal lPos, ByVal StrPtr(sThis), lLen
        m_iPos = m_iPos + lLen
    End If
    Exit Sub
Insert_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.Insert", Erl)
End Sub

Public Sub InsertByVal(ByVal iIndex As Long, ByVal sThis As String)
    On Error GoTo InsertByVal_Err
    Insert iIndex, sThis
    Exit Sub
InsertByVal_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.InsertByVal", Erl)
End Sub

Public Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)
    On Error GoTo Remove_Err
    Dim lSrc  As Long
    Dim lDst  As Long
    Dim lSize As Long
    ' is iIndex within bounds?
    If (iIndex * 2 > m_iPos) Then
        Err.Raise 9
    Else
        ' is there sufficient length?
        If ((iIndex + lLen) * 2 > m_iPos) Then
            Err.Raise 9
        Else
            ' Need to copy characters from iIndex*2 to m_iPos back by lLen chars:
            lSrc = UnsignedAdd(StrPtr(m_sString), (iIndex + lLen) * 2)
            lDst = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
            lSize = (m_iPos - (iIndex + lLen) * 2)
            CopyMemory ByVal lDst, ByVal lSrc, lSize
            m_iPos = m_iPos - lLen * 2
        End If
    End If
    Exit Sub
Remove_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.Remove", Erl)
End Sub

Public Function Find(ByVal sToFind As String, Optional ByVal lStartIndex As Long = 1, Optional ByVal compare As VbCompareMethod = vbTextCompare) As Long
    On Error GoTo Find_Err
    Dim lInstr As Long
    If (lStartIndex > 0) Then
        lInstr = InStr(lStartIndex, m_sString, sToFind, compare)
    Else
        lInstr = InStr(m_sString, sToFind, compare)
    End If
    If (lInstr < m_iPos \ 2) Then
        Find = lInstr
    End If
    Exit Function
Find_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.Find", Erl)
End Function

Public Sub HeapMinimize()
    On Error GoTo HeapMinimize_Err
    Dim iLen As Long
    ' Reduce the string size so only the minimal chunks
    ' are allocated:
    If (m_iLen - m_iPos) > m_iChunkSize Then
        iLen = m_iLen
        Do While (iLen - m_iPos) > m_iChunkSize
            iLen = iLen - m_iChunkSize
        Loop
        m_sString = Left$(m_sString, iLen \ 2)
        m_iLen = iLen
    End If
    Exit Sub
HeapMinimize_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.HeapMinimize", Erl)
End Sub

Private Function UnsignedAdd(start As Long, Incr As Long) As Long
    ' This function is useful when doing pointer arithmetic,
    ' but note it only works for positive values of Incr
    On Error GoTo UnsignedAdd_Err
    If start And &H80000000 Then 'Start < 0
        UnsignedAdd = start + Incr
    ElseIf (start Or &H80000000) < -Incr Then
        UnsignedAdd = start + Incr
    Else
        UnsignedAdd = (start + &H80000000) + (Incr + &H80000000)
    End If
    Exit Function
UnsignedAdd_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.UnsignedAdd", Erl)
End Function

Private Sub Class_Initialize()
    ' The default allocation: 8192 characters.
    On Error GoTo Class_Initialize_Err
    m_iChunkSize = 16384
    Exit Sub
Class_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "cStringBuilder.Class_Initialize", Erl)
End Sub
