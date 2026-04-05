Attribute VB_Name = "ModContinuousMap"
Option Explicit

Private Const BORDER_STRIP_DEFAULT As Byte = 8
Private Const BORDER_SYNC_INTERVAL_MS As Long = 150 ' ~6.6 Hz
Private Const MIRROR_VIEW_MARGIN As Byte = 2

Private Type t_MapNeighbors
    North As Integer
    South As Integer
    East As Integer
    West As Integer
End Type

Private Type t_BorderFlags
    North As Boolean
    South As Boolean
    East As Boolean
    West As Boolean
End Type

Public Sub SyncBorderEntitiesForUser(ByVal UserIndex As Integer, Optional ByVal Force As Boolean = False)
    On Error GoTo SyncBorderEntitiesForUser_Err
    If Not IsFeatureEnabled("continuous_map_border") Then Exit Sub
    If UserIndex <= 0 Or UserIndex > LastUser Then Exit Sub
    If Not UserList(UserIndex).flags.UserLogged Then Exit Sub

    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If Not Force Then
        If TicksElapsed(UserList(UserIndex).Counters.LastBorderMirrorSync, nowRaw) < BORDER_SYNC_INTERVAL_MS Then Exit Sub
    End If
    UserList(UserIndex).Counters.LastBorderMirrorSync = nowRaw

    Dim strip As Byte
    strip = GetBorderStripWidth()

    Dim active As t_BorderFlags
    Call ResolveActiveBorders(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, strip, active)

    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageBorderEntitiesReset())
    Call TraceBorderMirrorReset(UserIndex)

    If Not (active.North Or active.South Or active.East Or active.West) Then Exit Sub

    Dim neighbors As t_MapNeighbors
    Call GetMapNeighbors(UserList(UserIndex).pos.Map, neighbors)

    Dim totalCreates As Long
    Dim totalCandidates As Long

    If active.West Then Call SendBorderEntitiesFromMap(UserIndex, neighbors.West, e_Heading.WEST, strip, totalCandidates, totalCreates)
    If active.East Then Call SendBorderEntitiesFromMap(UserIndex, neighbors.East, e_Heading.EAST, strip, totalCandidates, totalCreates)
    If active.North Then Call SendBorderEntitiesFromMap(UserIndex, neighbors.North, e_Heading.NORTH, strip, totalCandidates, totalCreates)
    If active.South Then Call SendBorderEntitiesFromMap(UserIndex, neighbors.South, e_Heading.SOUTH, strip, totalCandidates, totalCreates)

    If totalCreates = 0 Then
        Call TraceBorderMirrorNoCreates(UserIndex, neighbors, active, totalCandidates)
    End If

    Exit Sub
SyncBorderEntitiesForUser_Err:
    Call TraceError(Err.Number, Err.Description, "ModContinuousMap.SyncBorderEntitiesForUser", Erl)
End Sub

Public Sub NotifyMapBorderViewers(ByVal Map As Integer)
    On Error GoTo NotifyMapBorderViewers_Err
    If Not IsFeatureEnabled("continuous_map_border") Then Exit Sub
    If Not MapaValido(Map) Then Exit Sub

    Dim i As Integer
    For i = 1 To ConnGroups(Map).CountEntrys
        Call SyncBorderEntitiesForUser(ConnGroups(Map).UserEntrys(i))
    Next i

    Dim neighbors As t_MapNeighbors
    Call GetMapNeighbors(Map, neighbors)
    Call SyncMapUsers(neighbors.North)
    Call SyncMapUsers(neighbors.South)
    Call SyncMapUsers(neighbors.East)
    Call SyncMapUsers(neighbors.West)

    Exit Sub
NotifyMapBorderViewers_Err:
    Call TraceError(Err.Number, Err.Description, "ModContinuousMap.NotifyMapBorderViewers", Erl)
End Sub

Private Sub SyncMapUsers(ByVal Map As Integer)
    On Error GoTo SyncMapUsers_Err
    If Not MapaValido(Map) Then Exit Sub

    Dim i As Integer
    For i = 1 To ConnGroups(Map).CountEntrys
        Call SyncBorderEntitiesForUser(ConnGroups(Map).UserEntrys(i))
    Next i

    Exit Sub
SyncMapUsers_Err:
    Call TraceError(Err.Number, Err.Description, "ModContinuousMap.SyncMapUsers", Erl)
End Sub

Private Sub SendBorderEntitiesFromMap(ByVal ReceiverUserIndex As Integer, ByVal NeighborMap As Integer, ByVal direction As e_Heading, ByVal strip As Byte, _
                                      ByRef totalCandidates As Long, ByRef totalCreates As Long)
    On Error GoTo SendBorderEntitiesFromMap_Err

    Dim candidates As Long
    Dim creates As Long

    If MapaValido(NeighborMap) Then
        Call SendNeighborUsers(ReceiverUserIndex, NeighborMap, direction, strip, candidates, creates)
        Call SendNeighborNpcs(ReceiverUserIndex, NeighborMap, direction, strip, candidates, creates)
    End If

    totalCandidates = totalCandidates + candidates
    totalCreates = totalCreates + creates

    Call TraceBorderMirrorEdge(ReceiverUserIndex, direction, candidates)

    Exit Sub
SendBorderEntitiesFromMap_Err:
    Call TraceError(Err.Number, Err.Description, "ModContinuousMap.SendBorderEntitiesFromMap", Erl)
End Sub

Private Sub SendNeighborUsers(ByVal ReceiverUserIndex As Integer, ByVal NeighborMap As Integer, ByVal direction As e_Heading, ByVal strip As Byte, _
                              ByRef candidates As Long, ByRef creates As Long)
    On Error GoTo SendNeighborUsers_Err

    Dim i As Integer
    Dim candidateUser As Integer
    Dim projectedX As Integer
    Dim projectedY As Integer

    For i = 1 To ConnGroups(NeighborMap).CountEntrys
        candidateUser = ConnGroups(NeighborMap).UserEntrys(i)
        If candidateUser > 0 And candidateUser <> ReceiverUserIndex Then
            If ShouldMirrorByDirection(UserList(candidateUser).pos.x, UserList(candidateUser).pos.y, direction, strip) Then
                candidates = candidates + 1
                projectedX = UserList(candidateUser).pos.x
                projectedY = UserList(candidateUser).pos.y
                Call ProjectByDirection(projectedX, projectedY, direction)

                If IsInsideMirrorWindow(ReceiverUserIndex, projectedX, projectedY, strip) Then
                    creates = creates + 1
                    Call SendData(SendTarget.ToIndex, ReceiverUserIndex, PrepareMessageBorderCharacterCreate(UserList(candidateUser).pos.Map, UserList(candidateUser).Char.charindex, _
                            projectedX, projectedY, UserList(candidateUser).Char.Heading, False, UserList(candidateUser).name))
                    Call TraceBorderMirrorCreate(ReceiverUserIndex, UserList(candidateUser).pos.Map, UserList(candidateUser).Char.charindex, projectedX, projectedY, False, UserList(candidateUser).name)
                End If
            End If
        End If
    Next i

    Exit Sub
SendNeighborUsers_Err:
    Call TraceError(Err.Number, Err.Description, "ModContinuousMap.SendNeighborUsers", Erl)
End Sub

Private Sub SendNeighborNpcs(ByVal ReceiverUserIndex As Integer, ByVal NeighborMap As Integer, ByVal direction As e_Heading, ByVal strip As Byte, _
                             ByRef candidates As Long, ByRef creates As Long)
    On Error GoTo SendNeighborNpcs_Err

    Dim x As Integer
    Dim y As Integer
    Dim npcIndex As Integer
    Dim projectedX As Integer
    Dim projectedY As Integer

    For x = MinXBorder To MaxXBorder
        For y = MinYBorder To MaxYBorder
            npcIndex = MapData(NeighborMap, x, y).NpcIndex
            If npcIndex > 0 Then
                If ShouldMirrorByDirection(NpcList(npcIndex).pos.x, NpcList(npcIndex).pos.y, direction, strip) Then
                    candidates = candidates + 1
                    projectedX = NpcList(npcIndex).pos.x
                    projectedY = NpcList(npcIndex).pos.y
                    Call ProjectByDirection(projectedX, projectedY, direction)

                    If IsInsideMirrorWindow(ReceiverUserIndex, projectedX, projectedY, strip) Then
                        creates = creates + 1
                        Call SendData(SendTarget.ToIndex, ReceiverUserIndex, PrepareMessageBorderCharacterCreate(NpcList(npcIndex).pos.Map, NpcList(npcIndex).Char.charindex, _
                                projectedX, projectedY, NpcList(npcIndex).Char.Heading, True, NpcList(npcIndex).name))
                        Call TraceBorderMirrorCreate(ReceiverUserIndex, NpcList(npcIndex).pos.Map, NpcList(npcIndex).Char.charindex, projectedX, projectedY, True, NpcList(npcIndex).name)
                    End If
                End If
            End If
        Next y
    Next x

    Exit Sub
SendNeighborNpcs_Err:
    Call TraceError(Err.Number, Err.Description, "ModContinuousMap.SendNeighborNpcs", Erl)
End Sub

Private Function IsInsideMirrorWindow(ByVal ReceiverUserIndex As Integer, ByVal projectedX As Integer, ByVal projectedY As Integer, ByVal strip As Byte) As Boolean
    Dim rangeX As Integer
    Dim rangeY As Integer

    rangeX = RANGO_VISION_X + strip + MIRROR_VIEW_MARGIN
    rangeY = RANGO_VISION_Y + strip + MIRROR_VIEW_MARGIN

    IsInsideMirrorWindow = Abs(projectedX - UserList(ReceiverUserIndex).pos.x) <= rangeX And Abs(projectedY - UserList(ReceiverUserIndex).pos.y) <= rangeY
End Function

Private Sub ResolveActiveBorders(ByVal x As Integer, ByVal y As Integer, ByVal strip As Byte, ByRef active As t_BorderFlags)
    active.West = x <= MinXBorder + strip
    active.East = x >= MaxXBorder - strip
    active.North = y <= MinYBorder + strip
    active.South = y >= MaxYBorder - strip
End Sub

Private Function GetBorderStripWidth() As Byte
    On Error GoTo GetBorderStripWidth_Err

    Dim rawValue As Long
    rawValue = val(SvrConfig.GetValue("ContinuousMapBorderStrip"))

    If rawValue < 1 Or rawValue > 20 Then
        GetBorderStripWidth = BORDER_STRIP_DEFAULT
    Else
        GetBorderStripWidth = CByte(rawValue)
    End If

    Exit Function
GetBorderStripWidth_Err:
    GetBorderStripWidth = BORDER_STRIP_DEFAULT
End Function

Private Sub GetMapNeighbors(ByVal Map As Integer, ByRef neighbors As t_MapNeighbors)
    On Error GoTo GetMapNeighbors_Err
    neighbors.North = GetDominantBorderExitMap(Map, e_Heading.NORTH)
    neighbors.South = GetDominantBorderExitMap(Map, e_Heading.SOUTH)
    neighbors.East = GetDominantBorderExitMap(Map, e_Heading.EAST)
    neighbors.West = GetDominantBorderExitMap(Map, e_Heading.WEST)
    Exit Sub
GetMapNeighbors_Err:
    Call TraceError(Err.Number, Err.Description, "ModContinuousMap.GetMapNeighbors", Erl)
End Sub

Private Function GetDominantBorderExitMap(ByVal Map As Integer, ByVal direction As e_Heading) As Integer
    On Error GoTo GetDominantBorderExitMap_Err

    Dim counts() As Integer
    ReDim counts(1 To NumMaps) As Integer

    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    Dim targetMap As Integer

    For i = MinXBorder To MaxXBorder
        Select Case direction
            Case e_Heading.NORTH
                x = i
                y = MinYBorder
            Case e_Heading.SOUTH
                x = i
                y = MaxYBorder
            Case e_Heading.WEST
                x = MinXBorder
                y = i
            Case e_Heading.EAST
                x = MaxXBorder
                y = i
            Case Else
                Exit For
        End Select

        targetMap = MapData(Map, x, y).TileExit.Map
        If targetMap > 0 And targetMap <= NumMaps Then
            counts(targetMap) = counts(targetMap) + 1
        End If
    Next i

    Dim bestCount As Integer
    bestCount = 0
    GetDominantBorderExitMap = 0

    For i = 1 To NumMaps
        If counts(i) > bestCount Then
            bestCount = counts(i)
            GetDominantBorderExitMap = i
        End If
    Next i

    Exit Function
GetDominantBorderExitMap_Err:
    Call TraceError(Err.Number, Err.Description, "ModContinuousMap.GetDominantBorderExitMap", Erl)
End Function

Private Function ShouldMirrorByDirection(ByVal x As Integer, ByVal y As Integer, ByVal direction As e_Heading, ByVal strip As Byte) As Boolean
    Select Case direction
        Case e_Heading.NORTH
            ShouldMirrorByDirection = y >= (MaxYBorder - strip)
        Case e_Heading.SOUTH
            ShouldMirrorByDirection = y <= (MinYBorder + strip)
        Case e_Heading.EAST
            ShouldMirrorByDirection = x <= (MinXBorder + strip)
        Case e_Heading.WEST
            ShouldMirrorByDirection = x >= (MaxXBorder - strip)
        Case Else
            ShouldMirrorByDirection = False
    End Select
End Function

Private Sub ProjectByDirection(ByRef x As Integer, ByRef y As Integer, ByVal direction As e_Heading)
    Select Case direction
        Case e_Heading.NORTH
            y = y - 100
        Case e_Heading.SOUTH
            y = y + 100
        Case e_Heading.EAST
            x = x + 100
        Case e_Heading.WEST
            x = x - 100
    End Select
End Sub

Private Sub TraceBorderMirrorEdge(ByVal ReceiverUserIndex As Integer, ByVal direction As e_Heading, ByVal candidates As Long)
    If Not IsFeatureEnabled("continuous_map_border_logs") Then Exit Sub
    Call LogError("BorderMirror: receiver=" & UserList(ReceiverUserIndex).name & " map=" & UserList(ReceiverUserIndex).pos.Map & " edge=" & GetEdgeName(direction) & _
            " candidates=" & candidates)
End Sub

Private Sub TraceBorderMirrorReset(ByVal UserIndex As Integer)
    If Not IsFeatureEnabled("continuous_map_border_logs") Then Exit Sub
    Call LogError("BorderMirror: send reset receiver=" & UserList(UserIndex).name)
End Sub

Private Sub TraceBorderMirrorCreate(ByVal UserIndex As Integer, ByVal OriginMap As Integer, ByVal CharIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal IsNpc As Boolean, _
                                    ByVal EntityName As String)
    If Not IsFeatureEnabled("continuous_map_border_logs") Then Exit Sub
    Call LogError("BorderMirror: send create receiver=" & UserList(UserIndex).name & " originMap=" & OriginMap & " idx=" & CharIndex & " x=" & x & " y=" & y & " isNpc=" _
            & IsNpc & " name=" & EntityName)
End Sub

Private Sub TraceBorderMirrorNoCreates(ByVal UserIndex As Integer, ByRef neighbors As t_MapNeighbors, ByRef active As t_BorderFlags, ByVal totalCandidates As Long)
    If Not IsFeatureEnabled("continuous_map_border_logs") Then Exit Sub
    Call LogError("BorderMirror: reset without creates receiver=" & UserList(UserIndex).name & " map=" & UserList(UserIndex).pos.Map)
    Call LogError("BorderMirror: validar selecciÃ³n de vecino N=" & neighbors.North & " S=" & neighbors.South & " E=" & neighbors.East & " W=" & neighbors.West)
    Call LogError("BorderMirror: validar filtro de borde activeN=" & active.North & " activeS=" & active.South & " activeE=" & active.East & " activeW=" & active.West)
    Call LogError("BorderMirror: validar loop sobre entidades del mapa vecino candidates=" & totalCandidates)
    Call LogError("BorderMirror: validar que proyecciÃ³n no descarte todas")
End Sub

Private Function GetEdgeName(ByVal direction As e_Heading) As String
    Select Case direction
        Case e_Heading.WEST
            GetEdgeName = "W"
        Case e_Heading.EAST
            GetEdgeName = "E"
        Case e_Heading.NORTH
            GetEdgeName = "N"
        Case e_Heading.SOUTH
            GetEdgeName = "S"
        Case Else
            GetEdgeName = "?"
    End Select
End Function
