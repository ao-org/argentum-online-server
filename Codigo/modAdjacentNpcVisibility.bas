Attribute VB_Name = "modAdjacentNpcVisibility"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit

' Phase 1: bounded two-hop interest, independent of pursuit policy.
Public Const ADJACENT_NPC_RADIUS_X As Integer = 24
Public Const ADJACENT_NPC_RADIUS_Y As Integer = 18
Public Const ADJACENT_NPC_LIMIT As Integer = 128
Private Const UPDATE_INTERVAL_MS As Long = 100
Private Const FAILURE_LOG_INTERVAL_MS As Long = 10000
Private mLastUpdate As Long
Private mLastFailureLog As Long
Private mFailureLogged As Boolean
Private mGeneration As Long
Private mKnown As Scripting.Dictionary
Private mSessions As Scripting.Dictionary
Private mEpochs As Scripting.Dictionary
Private mObserved As Scripting.Dictionary
#If UNIT_TEST = 1 Then
Private mScans As Scripting.Dictionary
#End If

Public Function IsAdjacentNpcObserved(ByVal NpcIndex As Integer) As Boolean
    If mObserved Is Nothing Then Exit Function
    IsAdjacentNpcObserved = mObserved.Exists(CStr(NpcIndex))
End Function

Public Function AdjacentNpcInsideInterest(ByVal viewerX As Integer, ByVal viewerY As Integer, _
                                          ByVal npcX As Integer, ByVal npcY As Integer, _
                                          ByVal dx As Long, ByVal dy As Long) As Boolean
    AdjacentNpcInsideInterest = Abs(CLng(npcX) + dx - viewerX) <= ADJACENT_NPC_RADIUS_X And _
                                Abs(CLng(npcY) + dy - viewerY) <= ADJACENT_NPC_RADIUS_Y
End Function

Public Sub UpdateAdjacentNpcVisibility(ByVal currentTime As Long)
    On Error GoTo Failed
    If mLastUpdate <> 0 Then
        If TicksElapsed(mLastUpdate, currentTime) < UPDATE_INTERVAL_MS Then Exit Sub
    End If
    mLastUpdate = currentTime
    #If UNIT_TEST = 1 Then
        Set mScans = New Scripting.Dictionary
    #End If
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
    session = CStr(UserList(userIndex).VersionId) & ":" & CStr(UserList(userIndex).pos.Map) & ":" & CStr(UserList(userIndex).ConnectionDetails.ConnID)
    If mSessions.Exists(key) Then
        If mSessions(key) <> session Then
            Call InvalidateAdjacentNpcViewer(userIndex)
        End If
    End If
    If Not CollectVisible(userIndex, current) Then Exit Sub
    Dim npcKey As Variant
    Dim npcIndex As Integer
    ' Interest keeps AI active even while a transport enqueue is being retried.
    For Each npcKey In current.Keys
        If Not mObserved.Exists(npcKey) Then mObserved.Add npcKey, True
    Next npcKey
    If Not mKnown.Exists(key) Then
        Set previous = New Scripting.Dictionary
        If mGeneration = &H7FFFFFFF Then mGeneration = 0
        mGeneration = mGeneration + 1
        ' Do not publish initialization until the reset was accepted by transport.
        If Not SendReset(userIndex, mGeneration) Then Exit Sub
        mKnown.Add key, previous
        mSessions.Add key, session
        mEpochs.Add key, mGeneration
    End If
    Set previous = mKnown(key)
    epoch = CLng(mEpochs(key))
    ' Commit each accepted difference, so partial success is neither lost nor replayed.
    ' Keys is a snapshot: removing entries during this loop is safe.
    For Each npcKey In previous.Keys
        If Not current.Exists(npcKey) Then
            If Not SendRemove(userIndex, epoch, CInt(npcKey), CInt(Split(CStr(previous(npcKey)), "|")(0))) Then Exit Sub
            previous.Remove npcKey
        ElseIf Split(CStr(previous(npcKey)), "|")(0) <> Split(CStr(current(npcKey)), "|")(0) Then
            If Not SendRemove(userIndex, epoch, CInt(npcKey), CInt(Split(CStr(previous(npcKey)), "|")(0))) Then Exit Sub
            previous.Remove npcKey
        End If
    Next npcKey
    For Each npcKey In current.Keys
        npcIndex = CInt(npcKey)
        If Not previous.Exists(npcKey) Then
            If Not SendUpsert(userIndex, epoch, npcIndex) Then Exit Sub
            previous.Add npcKey, current(npcKey)
        ElseIf previous(npcKey) <> current(npcKey) Then
            If Not SendUpsert(userIndex, epoch, npcIndex) Then Exit Sub
            previous(npcKey) = current(npcKey)
        End If
    Next npcKey
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.RefreshViewer", Erl)
    Call InvalidateAdjacentNpcViewer(userIndex)
End Sub

Private Function CollectVisible(ByVal userIndex As Integer, ByRef result As Scripting.Dictionary) As Boolean
    On Error GoTo Failed
    Dim mapId As Integer
    mapId = UserList(userIndex).pos.Map
    CollectVisible = True
    If mapId <= 0 Or mapId > NumMaps Then Exit Function
    ' Resource aliases/instances are deliberately unsupported in Phase 1.
    If MapInfo(mapId).MapResource <> mapId Then Exit Function
    Dim placements() As t_AdjacentNpcPlacement, count As Long, i As Long
    If Not TryGetAdjacentVisibilityPlacements(mapId, placements, count) Then
        CollectVisible = False
        Exit Function
    End If
    Dim neighbor As Integer
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    Dim x0 As Long, x1 As Long, y0 As Long, y1 As Long
    Dim npcIndex As Integer
    Dim key As String
    For i = 1 To count
        neighbor = placements(i).MapId
        dx = placements(i).dx
        dy = placements(i).dy
        If neighbor > 0 And neighbor <= NumMaps Then
            If MapInfo(neighbor).MapResource = neighbor Then
                x0 = max(XMinMapSize, CLng(UserList(userIndex).pos.x) - dx - ADJACENT_NPC_RADIUS_X)
                x1 = min(XMaxMapSize, CLng(UserList(userIndex).pos.x) - dx + ADJACENT_NPC_RADIUS_X)
                y0 = max(YMinMapSize, CLng(UserList(userIndex).pos.y) - dy - ADJACENT_NPC_RADIUS_Y)
                y1 = min(YMaxMapSize, CLng(UserList(userIndex).pos.y) - dy + ADJACENT_NPC_RADIUS_Y)
                If x0 <= x1 And y0 <= y1 Then
                    #If UNIT_TEST = 1 Then
                        key = CStr(neighbor)
                        If Not mScans.Exists(key) Then Call mScans.Add(key, 0)
                        mScans(key) = CLng(mScans(key)) + 1
                    #End If
                    For y = y0 To y1
                        For x = x0 To x1
                            npcIndex = MapData(x, y, neighbor).NpcIndex
                            If npcIndex > 0 Then
                                With NpcList(npcIndex)
                                    If .flags.NPCActive And .Char.charindex > 0 And .pos.Map = neighbor And .pos.x = x And .pos.y = y And _
                                        AdjacentNpcInsideInterest(UserList(userIndex).pos.x, UserList(userIndex).pos.y, .pos.x, .pos.y, dx, dy) Then
                                        If CanViewerSeeAdjacentNpc(userIndex, npcIndex) Then
                                            key = CStr(npcIndex)
                                            If Not result.Exists(key) Then result.Add key, VisualFingerprint(npcIndex)
                                            If result.Count >= ADJACENT_NPC_LIMIT Then Exit Function
                                        End If
                                    End If
                                End With
                            End If
                        Next x
                    Next y
                End If
            End If
        End If
    Next i
    Exit Function
Failed:
    CollectVisible = False
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.CollectVisible", Erl)
End Function

Private Function CanViewerSeeAdjacentNpc(ByVal userIndex As Integer, ByVal npcIndex As Integer) As Boolean
    With NpcList(npcIndex)
        ' The visual-only packet cannot represent hidden/translucent NPCs. There is
        ' no staff/owner exception in the NPC path (player visibility differs).
        If .flags.invisible <> 0 Or .Char.body <= 0 Then Exit Function
        ' Match MakeNPCChar's dead-viewer rule using the viewer's map safety.
        CanViewerSeeAdjacentNpc = UserList(userIndex).flags.Muerto = 0 Or _
            MapInfo(UserList(userIndex).pos.Map).Seguro <> 0 Or .npcType = e_NPCType.Revividor
    End With
End Function

Private Function VisualFingerprint(ByVal npcIndex As Integer) As String
    With NpcList(npcIndex)
        VisualFingerprint = CStr(.VersionId) & "|" & .pos.Map & "|" & .pos.x & "|" & .pos.y & "|" & _
            .Char.body & "|" & .Char.head & "|" & .Char.Heading & "|" & .Char.WeaponAnim & "|" & _
            .Char.ShieldAnim & "|" & .Char.CascoAnim & "|" & .Char.CartAnim & "|" & .Char.speeding & "|" & .flags.NPCIdle
    End With
End Function

Private Function SendReset(ByVal userIndex As Integer, ByVal epoch As Long) As Boolean
    On Error GoTo Failed
    Dim writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eAdjacentNpcReset)
    Call writer.WriteInt32(epoch)
    Call writer.WriteInt16(UserList(userIndex).pos.Map)
    Call modSendData.SendData(ToIndex, userIndex, , False, SendReset, True)
    If Not SendReset Then Call ReportSendFailure(userIndex, "reset")
    Exit Function
Failed:
    Call ReportSendFailure(userIndex, "reset serialization", Err.Number, Err.Description)
    If Not writer Is Nothing Then Call writer.Clear()
End Function

Private Function SendRemove(ByVal userIndex As Integer, ByVal epoch As Long, ByVal npcIndex As Integer, ByVal version As Integer) As Boolean
    On Error GoTo Failed
    Dim writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eAdjacentNpcRemove)
    Call writer.WriteInt32(epoch)
    Call writer.WriteInt16(npcIndex)
    Call writer.WriteInt16(version)
    Call modSendData.SendData(ToIndex, userIndex, , False, SendRemove, True)
    If Not SendRemove Then Call ReportSendFailure(userIndex, "remove")
    Exit Function
Failed:
    Call ReportSendFailure(userIndex, "remove serialization", Err.Number, Err.Description)
    If Not writer Is Nothing Then Call writer.Clear()
End Function

Private Function SendUpsert(ByVal userIndex As Integer, ByVal epoch As Long, ByVal npcIndex As Integer) As Boolean
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
    Call modSendData.SendData(ToIndex, userIndex, , False, SendUpsert, True)
    If Not SendUpsert Then Call ReportSendFailure(userIndex, "upsert")
    Exit Function
Failed:
    Call ReportSendFailure(userIndex, "upsert serialization", Err.Number, Err.Description)
    If Not writer Is Nothing Then Call writer.Clear()
End Function

Private Sub ReportSendFailure(ByVal userIndex As Integer, ByVal operation As String, Optional ByVal errorNumber As Long = 0, Optional ByVal description As String = "enqueue not accepted")
    ' A broken connection must not generate a log line on every 100 ms retry.
    Dim now As Long
    now = GetTickCountRaw()
    If mFailureLogged Then
        If TicksElapsed(mLastFailureLog, now) < FAILURE_LOG_INTERVAL_MS Then Exit Sub
    End If
    mFailureLogged = True
    mLastFailureLog = now
    Call TraceError(errorNumber, "user=" & CStr(userIndex) & " " & operation & ": " & description, "modAdjacentNpcVisibility.Send", 0)
End Sub

Public Sub InvalidateAdjacentNpcViewer(ByVal userIndex As Integer)
    On Error GoTo Failed
    If mKnown Is Nothing Then Exit Sub
    Dim key As String
    key = CStr(userIndex)
    If mKnown.Exists(key) Then mKnown.Remove key
    If mSessions.Exists(key) Then mSessions.Remove key
    If mEpochs.Exists(key) Then mEpochs.Remove key
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modAdjacentNpcVisibility.InvalidateViewer", Erl)
End Sub

#If UNIT_TEST = 1 Then
Public Function AdjacentNpcTestScanCount(ByVal MapId As Integer) As Long
    If mScans Is Nothing Then Exit Function
    If mScans.Exists(CStr(MapId)) Then AdjacentNpcTestScanCount = CLng(mScans(CStr(MapId)))
End Function

Public Sub AdjacentNpcTestResetState()
    mLastUpdate = 0
    mGeneration = 0
    Set mKnown = Nothing
    Set mSessions = Nothing
    Set mEpochs = Nothing
    Set mObserved = Nothing
    Set mScans = Nothing
End Sub

Public Function AdjacentNpcTestKnownCount(ByVal userIndex As Integer) As Long
    If mKnown Is Nothing Then Exit Function
    If Not mKnown.Exists(CStr(userIndex)) Then Exit Function
    AdjacentNpcTestKnownCount = mKnown(CStr(userIndex)).Count
End Function

Public Function AdjacentNpcTestEpoch(ByVal userIndex As Integer) As Long
    If mEpochs Is Nothing Then Exit Function
    If mEpochs.Exists(CStr(userIndex)) Then AdjacentNpcTestEpoch = CLng(mEpochs(CStr(userIndex)))
End Function

Public Function AdjacentNpcTestKnownVersion(ByVal userIndex As Integer, ByVal npcIndex As Integer) As Integer
    AdjacentNpcTestKnownVersion = -1
    If mKnown Is Nothing Then Exit Function
    If Not mKnown.Exists(CStr(userIndex)) Then Exit Function
    Dim known As Scripting.Dictionary
    Set known = mKnown(CStr(userIndex))
    If known.Exists(CStr(npcIndex)) Then AdjacentNpcTestKnownVersion = CInt(Split(CStr(known(CStr(npcIndex))), "|")(0))
End Function
#End If
