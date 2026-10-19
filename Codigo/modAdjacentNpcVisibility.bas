Attribute VB_Name = "modAdjacentNpcVisibility"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit

' Phase 1: bounded direct-neighbor interest, independent of pursuit policy.
Public Const ADJACENT_NPC_RADIUS_X As Integer = 24
Public Const ADJACENT_NPC_RADIUS_Y As Integer = 18
Public Const ADJACENT_NPC_LIMIT As Integer = 128
Private Const UPDATE_INTERVAL_MS As Long = 100
Private mLastUpdate As Long
Private mGeneration As Long
Private mKnown As Scripting.Dictionary
Private mSessions As Scripting.Dictionary
Private mEpochs As Scripting.Dictionary
Private mObserved As Scripting.Dictionary

Public Function IsAdjacentNpcObserved(ByVal NpcIndex As Integer) As Boolean
    If mObserved Is Nothing Then Exit Function
    IsAdjacentNpcObserved = mObserved.Exists(CStr(NpcIndex))
End Function

Public Function AdjacentNpcInsideInterest(ByVal viewerX As Integer, ByVal viewerY As Integer, _
                                          ByVal npcX As Integer, ByVal npcY As Integer, _
                                          ByVal dx As Integer, ByVal dy As Integer) As Boolean
    AdjacentNpcInsideInterest = Abs(CLng(npcX) + dx - viewerX) <= ADJACENT_NPC_RADIUS_X And _
                                Abs(CLng(npcY) + dy - viewerY) <= ADJACENT_NPC_RADIUS_Y
End Function

Public Sub UpdateAdjacentNpcVisibility(ByVal currentTime As Long)
    On Error GoTo Failed
    If mLastUpdate <> 0 Then
        If TicksElapsed(mLastUpdate, currentTime) < UPDATE_INTERVAL_MS Then Exit Sub
    End If
    mLastUpdate = currentTime
    If mKnown Is Nothing Then
        Set mKnown = New Scripting.Dictionary
        Set mSessions = New Scripting.Dictionary
        Set mEpochs = New Scripting.Dictionary
    End If
    Set mObserved = New Scripting.Dictionary
    Dim userIndex As Integer
    Dim key As String
    For userIndex = 1 To LastUser
        key = CStr(userIndex)
        If UserList(userIndex).flags.UserLogged And UserList(userIndex).ConnectionDetails.ConnIDValida Then
            Call RefreshViewer(userIndex)
        ElseIf mKnown.Exists(key) Then
            mKnown.Remove key
            mSessions.Remove key
            mEpochs.Remove key
        End If
    Next userIndex
    ' LastUser can shrink after disconnect; release abandoned slots too.
    Dim oldKey As Variant
    For Each oldKey In mKnown.Keys
        If CLng(oldKey) > LastUser Then
            mKnown.Remove oldKey
            mSessions.Remove oldKey
            mEpochs.Remove oldKey
        End If
    Next oldKey
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.Update", Erl)
End Sub

Private Sub RefreshViewer(ByVal userIndex As Integer)
    On Error GoTo Failed
    Dim key As String
    Dim session As String
    Dim epoch As Long
    Dim previous As Scripting.Dictionary
    Dim current As New Scripting.Dictionary
    key = CStr(userIndex)
    session = CStr(UserList(userIndex).VersionId) & ":" & CStr(UserList(userIndex).pos.Map)
    If mSessions.Exists(key) Then
        If mSessions(key) <> session Then
            mKnown.Remove key
            mSessions.Remove key
            mEpochs.Remove key
        End If
    End If
    If Not mKnown.Exists(key) Then
        Set previous = New Scripting.Dictionary
        If mGeneration = &H7FFFFFFF Then mGeneration = 0
        mGeneration = mGeneration + 1
        mKnown.Add key, previous
        mSessions.Add key, session
        mEpochs.Add key, mGeneration
        Call SendReset(userIndex, mGeneration)
    End If
    Set previous = mKnown(key)
    epoch = CLng(mEpochs(key))
    Call CollectVisible(userIndex, current)
    Dim npcKey As Variant
    Dim npcIndex As Integer
    For Each npcKey In previous.Keys
        If Not current.Exists(npcKey) Then
            Call SendRemove(userIndex, epoch, CInt(npcKey), CInt(Split(CStr(previous(npcKey)), "|")(0)))
        ElseIf Split(CStr(previous(npcKey)), "|")(0) <> Split(CStr(current(npcKey)), "|")(0) Then
            Call SendRemove(userIndex, epoch, CInt(npcKey), CInt(Split(CStr(previous(npcKey)), "|")(0)))
        End If
    Next npcKey
    For Each npcKey In current.Keys
        npcIndex = CInt(npcKey)
        If Not previous.Exists(npcKey) Then
            Call SendUpsert(userIndex, epoch, npcIndex)
        ElseIf previous(npcKey) <> current(npcKey) Then
            Call SendUpsert(userIndex, epoch, npcIndex)
        End If
        If Not mObserved.Exists(npcKey) Then mObserved.Add npcKey, True
    Next npcKey
    Set mKnown(key) = current
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.RefreshViewer", Erl)
End Sub

Private Sub CollectVisible(ByVal userIndex As Integer, ByRef result As Scripting.Dictionary)
    On Error GoTo Failed
    Dim mapId As Integer
    mapId = UserList(userIndex).pos.Map
    If mapId <= 0 Or mapId > NumMaps Then Exit Sub
    ' Resource aliases/instances are deliberately unsupported in Phase 1.
    If MapInfo(mapId).MapResource <> mapId Then Exit Sub
    Dim neighbors As Collection
    Set neighbors = AdjacentMapNeighbors(mapId)
    If neighbors Is Nothing Then Exit Sub
    Dim item As Variant
    Dim neighbor As Integer
    Dim dx As Integer, dy As Integer
    Dim x As Long, y As Long
    Dim x0 As Long, x1 As Long, y0 As Long, y1 As Long
    Dim npcIndex As Integer
    Dim key As String
    For Each item In neighbors
        neighbor = CInt(item)
        If neighbor > 0 And neighbor <= NumMaps Then
            If MapInfo(neighbor).MapResource = neighbor Then
                If TryGetAdjacentEdge(mapId, neighbor, dx, dy) Then
                    x0 = max(XMinMapSize, CLng(UserList(userIndex).pos.x) - dx - ADJACENT_NPC_RADIUS_X)
                    x1 = min(XMaxMapSize, CLng(UserList(userIndex).pos.x) - dx + ADJACENT_NPC_RADIUS_X)
                    y0 = max(YMinMapSize, CLng(UserList(userIndex).pos.y) - dy - ADJACENT_NPC_RADIUS_Y)
                    y1 = min(YMaxMapSize, CLng(UserList(userIndex).pos.y) - dy + ADJACENT_NPC_RADIUS_Y)
                    For y = y0 To y1
                        For x = x0 To x1
                            npcIndex = MapData(x, y, neighbor).NpcIndex
                            If npcIndex > 0 Then
                                With NpcList(npcIndex)
                                    If .flags.NPCActive And .Char.charindex > 0 And .pos.Map = neighbor And .pos.x = x And .pos.y = y And _
                                        AdjacentNpcInsideInterest(UserList(userIndex).pos.x, UserList(userIndex).pos.y, .pos.x, .pos.y, dx, dy) Then
                                        ' Match normal NPC visibility for dead viewers on unsafe maps.
                                        If UserList(userIndex).flags.Muerto = 0 Or MapInfo(mapId).Seguro <> 0 Or .npcType = e_NPCType.Revividor Then
                                            key = CStr(npcIndex)
                                            If Not result.Exists(key) Then result.Add key, VisualFingerprint(npcIndex)
                                            If result.Count >= ADJACENT_NPC_LIMIT Then Exit Sub
                                        End If
                                    End If
                                End With
                            End If
                        Next x
                    Next y
                End If
            End If
        End If
    Next item
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.CollectVisible", Erl)
End Sub

Private Function VisualFingerprint(ByVal npcIndex As Integer) As String
    With NpcList(npcIndex)
        VisualFingerprint = CStr(.VersionId) & "|" & .pos.Map & "|" & .pos.x & "|" & .pos.y & "|" & _
            .Char.body & "|" & .Char.head & "|" & .Char.Heading & "|" & .Char.WeaponAnim & "|" & _
            .Char.ShieldAnim & "|" & .Char.CascoAnim & "|" & .Char.CartAnim & "|" & .Char.speeding & "|" & .flags.NPCIdle
    End With
End Function

Private Sub SendReset(ByVal userIndex As Integer, ByVal epoch As Long)
    On Error GoTo Failed
    Dim writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eAdjacentNpcReset)
    Call writer.WriteInt32(epoch)
    Call writer.WriteInt16(UserList(userIndex).pos.Map)
    Call modSendData.SendData(ToIndex, userIndex)
    Exit Sub
Failed:
    If Not writer Is Nothing Then Call writer.Clear()
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.SendReset", Erl)
End Sub

Private Sub SendRemove(ByVal userIndex As Integer, ByVal epoch As Long, ByVal npcIndex As Integer, ByVal version As Integer)
    On Error GoTo Failed
    Dim writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eAdjacentNpcRemove)
    Call writer.WriteInt32(epoch)
    Call writer.WriteInt16(npcIndex)
    Call writer.WriteInt16(version)
    Call modSendData.SendData(ToIndex, userIndex)
    Exit Sub
Failed:
    If Not writer Is Nothing Then Call writer.Clear()
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.SendRemove", Erl)
End Sub

Private Sub SendUpsert(ByVal userIndex As Integer, ByVal epoch As Long, ByVal npcIndex As Integer)
    On Error GoTo Failed
    Dim writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eAdjacentNpcUpsert)
    Call writer.WriteInt32(epoch)
    Call writer.WriteInt16(npcIndex)
    With NpcList(npcIndex)
        Call writer.WriteInt16(.VersionId)
        Call writer.WriteInt16(.pos.Map)
        Call writer.WriteInt8(.pos.x)
        Call writer.WriteInt8(.pos.y)
        Call writer.WriteInt16(.Char.body)
        Call writer.WriteInt16(.Char.head)
        Call writer.WriteInt8(.Char.Heading)
        Call writer.WriteInt16(.Char.WeaponAnim)
        Call writer.WriteInt16(.Char.ShieldAnim)
        Call writer.WriteInt16(.Char.CascoAnim)
        Call writer.WriteInt16(.Char.CartAnim)
        Call writer.WriteReal32(.Char.speeding)
        Call writer.WriteBool(.flags.NPCIdle)
    End With
    Call modSendData.SendData(ToIndex, userIndex)
    Exit Sub
Failed:
    If Not writer Is Nothing Then Call writer.Clear()
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.SendUpsert", Erl)
End Sub

Public Sub InvalidateAdjacentNpcViewer(ByVal userIndex As Integer)
    On Error GoTo Failed
    If mKnown Is Nothing Then Exit Sub
    Dim key As String
    key = CStr(userIndex)
    If Not mKnown.Exists(key) Then Exit Sub
    mKnown.Remove key
    mSessions.Remove key
    mEpochs.Remove key
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.InvalidateViewer", Erl)
End Sub
