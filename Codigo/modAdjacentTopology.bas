Attribute VB_Name = "modAdjacentTopology"
Option Explicit

Private Const TOPOLOGY_VERSION As Long = 1
Private Const TOPOLOGY_ID_SPACE As String = "resource_map_id"

Private mAvailable As Boolean
Private mMaps As Scripting.Dictionary
Private mEdges As Scripting.Dictionary
Private mEdgeDx As Scripting.Dictionary
Private mEdgeDy As Scripting.Dictionary
Private mNeighbors As Scripting.Dictionary
Private mFailureLogged As Boolean

Public Type t_AdjacentNpcPlacement
    MapId As Integer
    dx As Long
    dy As Long
    Depth As Integer
End Type

' Visibility only: never add composed edges to movement/pursuit routing.
Public Function TryGetAdjacentVisibilityPlacements(ByVal FromMap As Integer, ByRef placements() As t_AdjacentNpcPlacement, ByRef count As Long) As Boolean
    On Error GoTo Failed
    count = 0
    Erase placements
    TryGetAdjacentVisibilityPlacements = True
    If Not mAvailable Then Exit Function
    If Not mMaps.Exists(CStr(FromMap)) Then Exit Function
    ReDim placements(1 To mMaps.Count)
    Dim indices As New Scripting.Dictionary, conflicts As New Scripting.Dictionary
    Dim firstHop As Collection, secondHop As Collection
    Dim first As Variant, second As Variant
    Dim dx As Integer, dy As Integer, nextDx As Integer, nextDy As Integer
    Set firstHop = mNeighbors(CStr(FromMap))
    For Each first In firstHop
        If TryGetAdjacentEdge(FromMap, CInt(first), dx, dy) Then
            Call AddVisibilityPlacement(placements, count, indices, conflicts, CInt(first), CLng(dx), CLng(dy), 1)
        End If
    Next first
    ' Expand every direct neighbor, regardless of NPCs or interest intersection.
    ' Exactly two edge levels bound cycles and discovery work.
    For Each first In firstHop
        If TryGetAdjacentEdge(FromMap, CInt(first), dx, dy) Then
            Set secondHop = mNeighbors(CStr(first))
            For Each second In secondHop
                If CInt(second) <> FromMap Then
                    If TryGetAdjacentEdge(CInt(first), CInt(second), nextDx, nextDy) Then
                        Call AddVisibilityPlacement(placements, count, indices, conflicts, CInt(second), CLng(dx) + CLng(nextDx), CLng(dy) + CLng(nextDy), 2)
                    End If
                End If
            Next second
        End If
    Next first
    ' Packets identify source maps, not paths. Suppress ambiguous transforms rather
    ' than picking one that may disagree with the client's prepared terrain view.
    Dim i As Long, kept As Long, j As Long, placement As t_AdjacentNpcPlacement
    For i = 1 To count
        If Not conflicts.Exists(CStr(placements(i).MapId)) Then
            kept = kept + 1
            placements(kept) = placements(i)
        End If
    Next i
    count = kept
    ' HOO sorts by depth, map ID, y, x; one unambiguous placement per map means
    ' the last two tie-breakers cannot differ here. NPCs then follow tile order.
    For i = 2 To count
        placement = placements(i)
        j = i - 1
        Do While j > 0
            If placements(j).Depth < placement.Depth Then Exit Do
            If placements(j).Depth = placement.Depth And placements(j).MapId < placement.MapId Then Exit Do
            placements(j + 1) = placements(j)
            j = j - 1
        Loop
        placements(j + 1) = placement
    Next i
    Exit Function
Failed:
    count = 0
    TryGetAdjacentVisibilityPlacements = False
    Call TraceError(Err.Number, Err.Description, "modAdjacentTopology.TryGetAdjacentVisibilityPlacements", Erl)
End Function

Private Sub AddVisibilityPlacement(ByRef placements() As t_AdjacentNpcPlacement, ByRef count As Long, ByVal indices As Scripting.Dictionary, ByVal conflicts As Scripting.Dictionary, ByVal MapId As Integer, ByVal dx As Long, ByVal dy As Long, ByVal Depth As Integer)
    Dim key As String, index As Long
    key = CStr(MapId)
    If indices.Exists(key) Then
        index = CLng(indices(key))
        If placements(index).dx <> dx Or placements(index).dy <> dy Then
            If Not conflicts.Exists(key) Then Call conflicts.Add(key, True)
        End If
        Exit Sub
    End If
    count = count + 1
    Call indices.Add(key, count)
    With placements(count)
        .MapId = MapId
        .dx = dx
        .dy = dy
        .Depth = Depth
    End With
End Sub

Public Sub LoadAdjacentTopology()
    Dim topologyPath As String
    topologyPath = App.Path & "\Recursos\config\legacy_overworld_topology.json"

    If Not FileExist(topologyPath, vbNormal) Then
        Call DisableAdjacentTopology("manifest not found")
        Exit Sub
    End If

    If LoadAdjacentTopologyJson(FileText(topologyPath)) Then
        Call LogInfoServidor("Adjacent topology loaded maps=" & mMaps.Count & " edges=" & mEdges.Count)
    End If
End Sub

Public Function LoadAdjacentTopologyJson(ByVal json As String, Optional ByVal LogFailure As Boolean = True) As Boolean
    On Error GoTo LoadFailed

    Dim root As Scripting.Dictionary
    Set root = mod_JSON.parse(json)
    If root Is Nothing Then Err.Raise vbObjectError + 700, "modAdjacentTopology", "invalid JSON"
    If mod_JSON.GetParserErrors() <> vbNullString Then Err.Raise vbObjectError + 701, "modAdjacentTopology", mod_JSON.GetParserErrors()
    If Not root.Exists("version") Or CLng(root("version")) <> TOPOLOGY_VERSION Then Err.Raise vbObjectError + 702, "modAdjacentTopology", "unsupported version"
    If Not root.Exists("id_space") Or CStr(root("id_space")) <> TOPOLOGY_ID_SPACE Then Err.Raise vbObjectError + 703, "modAdjacentTopology", "invalid id_space"
    If Not root.Exists("maps") Or Not root.Exists("edges") Then Err.Raise vbObjectError + 704, "modAdjacentTopology", "missing maps or edges"

    Dim maps As New Scripting.Dictionary
    Dim edges As New Scripting.Dictionary
    Dim edgeDx As New Scripting.Dictionary
    Dim edgeDy As New Scripting.Dictionary
    Dim neighbors As New Scripting.Dictionary
    Dim mapItems As Collection
    Dim edgeItems As Collection
    Dim edge As Collection
    Dim item As Variant
    Dim mapId As Long
    Dim fromMap As Long
    Dim toMap As Long
    Dim dx As Long
    Dim dy As Long
    Dim key As String

    Set mapItems = root("maps")
    For Each item In mapItems
        mapId = CLng(item)
        If mapId <= 0 Or mapId > 32767 Then Err.Raise vbObjectError + 705, "modAdjacentTopology", "invalid map id"
        key = CStr(mapId)
        If maps.Exists(key) Then Err.Raise vbObjectError + 706, "modAdjacentTopology", "duplicate map id"
        maps.Add key, True
        Dim newNeighbors As Collection
        Set newNeighbors = New Collection
        neighbors.Add key, newNeighbors
    Next item

    Set edgeItems = root("edges")
    For Each item In edgeItems
        Set edge = item
        If edge.Count <> 4 Then Err.Raise vbObjectError + 707, "modAdjacentTopology", "invalid edge"
        fromMap = CLng(edge(1))
        toMap = CLng(edge(2))
        dx = CLng(edge(3))
        dy = CLng(edge(4))
        If fromMap = toMap Or (dx = 0 And dy = 0) Then Err.Raise vbObjectError + 708, "modAdjacentTopology", "invalid edge transform"
        If Not maps.Exists(CStr(fromMap)) Or Not maps.Exists(CStr(toMap)) Then Err.Raise vbObjectError + 709, "modAdjacentTopology", "edge references unknown map"
        key = EdgeKey(fromMap, toMap)
        If edges.Exists(key) Then Err.Raise vbObjectError + 710, "modAdjacentTopology", "duplicate directed edge"
        edges.Add key, True
        edgeDx.Add key, CInt(dx)
        edgeDy.Add key, CInt(dy)
        Dim mapNeighbors As Collection
        Set mapNeighbors = neighbors(CStr(fromMap))
        mapNeighbors.Add CInt(toMap)
    Next item

    Dim reverseKey As String
    Dim edgeKeyItem As Variant
    For Each edgeKeyItem In edges.Keys
        ParseEdgeKey CStr(edgeKeyItem), fromMap, toMap
        reverseKey = EdgeKey(toMap, fromMap)
        If Not edges.Exists(reverseKey) Then Err.Raise vbObjectError + 711, "modAdjacentTopology", "missing reciprocal edge"
        If CInt(edgeDx(edgeKeyItem)) <> -CInt(edgeDx(reverseKey)) Or CInt(edgeDy(edgeKeyItem)) <> -CInt(edgeDy(reverseKey)) Then
            Err.Raise vbObjectError + 712, "modAdjacentTopology", "inconsistent reciprocal edge"
        End If
    Next edgeKeyItem

    Set mMaps = maps
    Set mEdges = edges
    Set mEdgeDx = edgeDx
    Set mEdgeDy = edgeDy
    Set mNeighbors = neighbors
    mAvailable = True
    mFailureLogged = False
    LoadAdjacentTopologyJson = True
    Exit Function

LoadFailed:
    mAvailable = False
    If LogFailure Then Call DisableAdjacentTopology(Err.Description)
End Function

Public Sub ResetAdjacentTopology()
    mAvailable = False
    Set mMaps = Nothing
    Set mEdges = Nothing
    Set mEdgeDx = Nothing
    Set mEdgeDy = Nothing
    Set mNeighbors = Nothing
End Sub

Public Function AdjacentTopologyAvailable() As Boolean
    AdjacentTopologyAvailable = mAvailable
End Function

Public Function AdjacentTopologyContainsMap(ByVal MapId As Integer) As Boolean
    If Not mAvailable Then Exit Function
    AdjacentTopologyContainsMap = mMaps.Exists(CStr(MapId))
End Function

Public Function TryGetAdjacentEdge(ByVal FromMap As Integer, ByVal ToMap As Integer, ByRef dx As Integer, ByRef dy As Integer) As Boolean
    If Not mAvailable Then Exit Function
    Dim key As String
    key = EdgeKey(FromMap, ToMap)
    If Not mEdges.Exists(key) Then Exit Function
    dx = CInt(mEdgeDx(key))
    dy = CInt(mEdgeDy(key))
    TryGetAdjacentEdge = True
End Function

Public Function TryGetNextSpatialHop(ByVal FromMap As Integer, _
                                     ByVal TargetMap As Integer, _
                                     ByRef NextMap As Integer, _
                                     Optional ByVal MaxHops As Integer = 0, _
                                     Optional ByVal AvoidSafeMaps As Boolean = False) As Boolean
    On Error GoTo RoutingFailed
    If Not mAvailable Then Exit Function
    If FromMap = TargetMap Then
        NextMap = TargetMap
        TryGetNextSpatialHop = True
        Exit Function
    End If
    If Not mMaps.Exists(CStr(FromMap)) Or Not mMaps.Exists(CStr(TargetMap)) Then Exit Function

    Dim visited As New Scripting.Dictionary
    Dim firstHop As New Scripting.Dictionary
    Dim depth As New Scripting.Dictionary
    Dim queue() As Integer
    ReDim queue(0 To mMaps.Count - 1)
    Dim queueHead As Long
    Dim queueTail As Long
    Dim currentMap As Integer
    Dim candidate As Integer
    Dim currentDepth As Integer
    Dim neighbor As Variant
    Dim mapNeighbors As Collection

    queue(0) = FromMap
    visited.Add CStr(FromMap), True
    depth.Add CStr(FromMap), 0

    Do While queueHead <= queueTail
        currentMap = queue(queueHead)
        queueHead = queueHead + 1
        currentDepth = CInt(depth(CStr(currentMap)))
        If MaxHops = 0 Or currentDepth < MaxHops Then
            Set mapNeighbors = mNeighbors(CStr(currentMap))
            For Each neighbor In mapNeighbors
                candidate = CInt(neighbor)
                If Not visited.Exists(CStr(candidate)) Then
                    If Not AvoidSafeMaps Or Not IsTopologyMapSafe(candidate) Then
                        visited.Add CStr(candidate), True
                        depth.Add CStr(candidate), currentDepth + 1
                        If currentMap = FromMap Then
                            firstHop.Add CStr(candidate), candidate
                        Else
                            firstHop.Add CStr(candidate), CInt(firstHop(CStr(currentMap)))
                        End If
                        If candidate = TargetMap Then
                            NextMap = CInt(firstHop(CStr(candidate)))
                            TryGetNextSpatialHop = True
                            Exit Function
                        End If
                        queueTail = queueTail + 1
                        queue(queueTail) = candidate
                    End If
                End If
            Next neighbor
        End If
    Loop
    Exit Function

RoutingFailed:
    Call TraceError(Err.Number, Err.Description, "modAdjacentTopology.TryGetNextSpatialHop", Erl)
End Function

Public Function SpatialTransitionMatchesEdge(ByVal FromMap As Integer, _
                                             ByVal ExitX As Integer, _
                                             ByVal ExitY As Integer, _
                                             ByVal ToMap As Integer, _
                                             ByVal DestinationX As Integer, _
                                             ByVal DestinationY As Integer) As Boolean
    Dim dx As Integer
    Dim dy As Integer
    If Not TryGetAdjacentEdge(FromMap, ToMap, dx, dy) Then Exit Function
    SpatialTransitionMatchesEdge = (DestinationX + dx = ExitX And DestinationY + dy = ExitY)
End Function

Private Function IsTopologyMapSafe(ByVal MapId As Integer) As Boolean
    If MapId > 0 And MapId <= NumMaps Then IsTopologyMapSafe = (MapInfo(MapId).Seguro <> 0)
End Function

Private Function EdgeKey(ByVal FromMap As Long, ByVal ToMap As Long) As String
    EdgeKey = CStr(FromMap) & ":" & CStr(ToMap)
End Function

Private Sub ParseEdgeKey(ByVal key As String, ByRef FromMap As Long, ByRef ToMap As Long)
    Dim separator As Long
    separator = InStr(1, key, ":", vbBinaryCompare)
    FromMap = CLng(Left$(key, separator - 1))
    ToMap = CLng(mid$(key, separator + 1))
End Sub

Private Sub DisableAdjacentTopology(ByVal reason As String)
    mAvailable = False
    If Not mFailureLogged Then
        mFailureLogged = True
        Call LogError("Adjacent topology unavailable: " & reason)
    End If
End Sub

Public Function AdjacentMapNeighbors(ByVal MapId As Integer) As Collection
    On Error GoTo Failed
    If Not mAvailable Then Exit Function
    If Not mNeighbors.Exists(CStr(MapId)) Then Exit Function
    Set AdjacentMapNeighbors = mNeighbors(CStr(MapId))
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "modAdjacentTopology.AdjacentMapNeighbors", Erl)
End Function
