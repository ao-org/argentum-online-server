Attribute VB_Name = "modNpcCrossMapPursuit"
Option Explicit

Public Const NPC_CROSS_MAP_MAX_CHASE_HOPS As Integer = 3
Public Const NPC_CROSS_MAP_PURSUIT_FEATURE As String = "npc-cross-map-pursuit"
Private Const NPC_CROSS_MAP_MAX_PATH_FAILURES As Byte = 3

Public Function IsVerifiedNpcSpatialTransition(ByVal FromMap As Integer, _
                                               ByVal ExitX As Integer, _
                                               ByVal ExitY As Integer, _
                                               ByRef Destination As t_WorldPos, _
                                               ByVal IsTeleport As Boolean) As Boolean
    If IsTeleport Then Exit Function
    If Not IsFeatureEnabled(NPC_CROSS_MAP_PURSUIT_FEATURE) Then Exit Function
    If Not AdjacentTopologyAvailable() Then Exit Function
    If Destination.Map <= 0 Or Destination.Map > NumMaps Then Exit Function
    If Not NpcCrossMapDestinationAllowed(Destination.Map) Then Exit Function
    IsVerifiedNpcSpatialTransition = SpatialTransitionMatchesEdge(FromMap, ExitX, ExitY, Destination.Map, Destination.x, Destination.y)
End Function

Public Sub RetainNpcPursuitForSpatialTransition(ByVal UserIndex As Integer, _
                                                ByVal FromMap As Integer, _
                                                ByVal ToMap As Integer)
    On Error GoTo RetainFailed
    If Not NpcCrossMapPursuitEnabled() Then Exit Sub
    Dim NpcIndex As Integer
    Dim matchingTargets As Integer
    Dim sourceMapTargets As Integer
    Dim retainedTargets As Integer
    For NpcIndex = 1 To LastNPC
        If IsValidUserRef(NpcList(NpcIndex).TargetUser) Then
            If NpcList(NpcIndex).TargetUser.ArrayIndex = UserIndex Then
                matchingTargets = matchingTargets + 1
                If NpcList(NpcIndex).pos.Map = FromMap Then sourceMapTargets = sourceMapTargets + 1
            End If
        End If
        'Only NPCs already pursuing from the map being left may retain aggro.
        'Destination-map NPCs can acquire the user while WarpUserChar rebuilds visibility.
        If NpcList(NpcIndex).pos.Map = FromMap Then
            If IsValidUserRef(NpcList(NpcIndex).TargetUser) Then
                If NpcList(NpcIndex).TargetUser.ArrayIndex = UserIndex Then
                    If NpcCanPursueAcrossMaps(NpcIndex) Then
                        retainedTargets = retainedTargets + 1
                        With NpcList(NpcIndex).CrossMapRoute
                            .Mode = eNpcCrossMapRouteChase
                            .TargetMap = ToMap
                            .NextMap = 0
                            .ExitX = 0
                            .ExitY = 0
                            .DestinationX = 0
                            .DestinationY = 0
                            .PathFailures = 0
                        End With
                        NpcList(NpcIndex).pathFindingInfo.PathLength = 0
                        Call LogNpcCrossMap(NpcIndex, "chase retained fromMap=" & FromMap & " targetMap=" & ToMap)
                        'Prime the first hop while the transition context is still known. The AI loop
                        'will continue the route after the source map becomes empty.
                        If Not PrepareNpcNextHop(NpcIndex, ToMap, True) Then
                            Call ClearUserRef(NpcList(NpcIndex).TargetUser)
                            Call BeginNpcReturnHome(NpcIndex, "initial route unavailable")
                        End If
                    End If
                End If
            End If
        End If
    Next NpcIndex
    Call LogInfoServidor("NPC cross-map retention scan user=" & UserIndex & " fromMap=" & FromMap & " toMap=" & ToMap & " matchingTargets=" & matchingTargets & " sourceMapTargets=" & sourceMapTargets & " retained=" & retainedTargets)
    Exit Sub
RetainFailed:
    Call TraceError(Err.Number, Err.Description, "modNpcCrossMapPursuit.RetainNpcPursuitForSpatialTransition", Erl)
End Sub

Public Sub CancelNpcCrossMapPursuitForUser(ByVal UserIndex As Integer)
    On Error GoTo CancelFailed
    Dim NpcIndex As Integer
    For NpcIndex = 1 To LastNPC
        If NpcList(NpcIndex).CrossMapRoute.Mode = eNpcCrossMapRouteChase Then
            If IsValidUserRef(NpcList(NpcIndex).TargetUser) Then
                If NpcList(NpcIndex).TargetUser.ArrayIndex = UserIndex Then
                    Call ClearUserRef(NpcList(NpcIndex).TargetUser)
                    Call BeginNpcReturnHome(NpcIndex, "target transition is not spatial")
                End If
            End If
        End If
    Next NpcIndex
    Exit Sub
CancelFailed:
    Call TraceError(Err.Number, Err.Description, "modNpcCrossMapPursuit.CancelNpcCrossMapPursuitForUser", Erl)
End Sub

Public Function ShouldPreserveNpcAttackerReference(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    If NpcIndex <= 0 Or NpcIndex > LastNPC Then Exit Function
    If Not NpcCrossMapPursuitEnabled() Then Exit Function
    If NpcList(NpcIndex).CrossMapRoute.Mode <> eNpcCrossMapRouteChase Then Exit Function
    If Not IsValidUserRef(NpcList(NpcIndex).TargetUser) Then Exit Function
    ShouldPreserveNpcAttackerReference = (NpcList(NpcIndex).TargetUser.ArrayIndex = UserIndex)
End Function

Public Function HandleNpcCrossMapRoute(ByVal NpcIndex As Integer) As Boolean
    On Error GoTo RouteFailed
    Dim routeDestination As t_WorldPos
    If NpcIndex <= 0 Or NpcIndex > LastNPC Then Exit Function
    If Not NpcCrossMapPursuitEnabled() Then
        If NpcList(NpcIndex).CrossMapRoute.Mode <> eNpcCrossMapRouteNone Then Call ResetNpcCrossMapRoute(NpcIndex)
        Exit Function
    End If

    With NpcList(NpcIndex)
        If .CrossMapRoute.Mode = eNpcCrossMapRouteNone Then Exit Function

        If .CrossMapRoute.Mode = eNpcCrossMapRouteChase Then
            If Not IsValidUserRef(.TargetUser) Then
                Call BeginNpcReturnHome(NpcIndex, "target reference invalid")
            ElseIf Not NpcTargetAllowsRetainedChase(.TargetUser.ArrayIndex) Then
                Call ClearUserRef(.TargetUser)
                Call BeginNpcReturnHome(NpcIndex, "target no longer attackable")
            ElseIf .pos.Map = UserList(.TargetUser.ArrayIndex).pos.Map Then
                .CrossMapRoute.TargetMap = .pos.Map
                Exit Function
            Else
                .CrossMapRoute.TargetMap = UserList(.TargetUser.ArrayIndex).pos.Map
                If Not PrepareNpcNextHop(NpcIndex, .CrossMapRoute.TargetMap, True) Then
                    Call ClearUserRef(.TargetUser)
                    Call BeginNpcReturnHome(NpcIndex, "no allowed chase route")
                Else
                    routeDestination = NpcCrossMapExitPosition(NpcIndex)
                    Call AI_CaminarConRumbo(NpcIndex, routeDestination)
                    HandleNpcCrossMapRoute = True
                    Exit Function
                End If
            End If
        End If

        If .CrossMapRoute.Mode = eNpcCrossMapRouteReturnHome Then
            HandleNpcCrossMapRoute = True
            If .pos.Map = .Orig.Map Then
                If .pos.x = .Orig.x And .pos.y = .Orig.y Then
                    Call ResetNpcCrossMapRoute(NpcIndex)
                    HandleNpcCrossMapRoute = False
                Else
                    Call AI_CaminarConRumbo(NpcIndex, .Orig)
                End If
            ElseIf PrepareNpcNextHop(NpcIndex, .Orig.Map, False) Then
                routeDestination = NpcCrossMapExitPosition(NpcIndex)
                Call AI_CaminarConRumbo(NpcIndex, routeDestination)
            Else
                Call LogNpcCrossMap(NpcIndex, "return route unavailable; warping home")
                Call WarpNpcChar(NpcIndex, .Orig.Map, .Orig.x, .Orig.y, False)
                Call ResetNpcCrossMapRoute(NpcIndex)
            End If
        End If
    End With
    Exit Function

RouteFailed:
    Call TraceError(Err.Number, Err.Description, "modNpcCrossMapPursuit.HandleNpcCrossMapRoute", Erl)
End Function

Public Function NpcPathMayUseExit(ByVal NpcIndex As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If NpcIndex <= 0 Or NpcIndex > LastNPC Then Exit Function
    If Not NpcCrossMapPursuitEnabled() Then Exit Function
    With NpcList(NpcIndex).CrossMapRoute
        If .Mode = eNpcCrossMapRouteNone Then Exit Function
        If .ExitX <> x Or .ExitY <> y Then Exit Function
        If NpcList(NpcIndex).pathFindingInfo.destination.x <> x Or NpcList(NpcIndex).pathFindingInfo.destination.y <> y Then Exit Function
        NpcPathMayUseExit = SpatialTransitionMatchesEdge(NpcList(NpcIndex).pos.Map, x, y, .NextMap, .DestinationX, .DestinationY)
    End With
End Function

Public Function IsNpcAuthorizedSpatialExit(ByVal NpcIndex As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If Not NpcPathMayUseExit(NpcIndex, x, y) Then Exit Function
    IsNpcAuthorizedSpatialExit = (MapData(NpcList(NpcIndex).pos.Map, x, y).TileExit.Map = NpcList(NpcIndex).CrossMapRoute.NextMap)
End Function

Public Function CrossNpcSpatialExit(ByVal NpcIndex As Integer, ByVal ExitX As Integer, ByVal ExitY As Integer) As Boolean
    On Error GoTo CrossFailed
    If Not IsNpcAuthorizedSpatialExit(NpcIndex, ExitX, ExitY) Then Exit Function

    With NpcList(NpcIndex)
        If Abs(.pos.x - ExitX) + Abs(.pos.y - ExitY) <> 1 Then Exit Function
        Dim destination As t_WorldPos
        destination = MapData(.pos.Map, ExitX, ExitY).TileExit
        If destination.Map <> .CrossMapRoute.NextMap Then Exit Function
        If destination.x <> .CrossMapRoute.DestinationX Or destination.y <> .CrossMapRoute.DestinationY Then Exit Function
        If Not NpcDestinationAvailable(NpcIndex, destination) Then
            Call RecordNpcCrossMapPathFailure(NpcIndex)
            Exit Function
        End If
        If .CrossMapRoute.Mode = eNpcCrossMapRouteChase Then
            If Not IsValidUserRef(.TargetUser) Then Exit Function
            If Not NpcTargetAllowsRetainedChase(.TargetUser.ArrayIndex) Then Exit Function
        End If

        Dim oldMap As Integer
        oldMap = .pos.Map
        Call EraseNPCChar(NpcIndex)
        .pos = destination
        Call MakeNPCChar(True, 0, NpcIndex, destination.Map, destination.x, destination.y)
        .pathFindingInfo.PathLength = 0
        .pathFindingInfo.NextPathRecomputeAt = 0
        .CrossMapRoute.NextMap = 0
        .CrossMapRoute.ExitX = 0
        .CrossMapRoute.ExitY = 0
        .CrossMapRoute.DestinationX = 0
        .CrossMapRoute.DestinationY = 0
        .CrossMapRoute.PathFailures = 0
        If .CrossMapRoute.HopsCrossed < 255 Then .CrossMapRoute.HopsCrossed = .CrossMapRoute.HopsCrossed + 1
        Call LogNpcCrossMap(NpcIndex, "spatial exit crossed " & oldMap & " -> " & destination.Map & " target preserved=" & CStr(IsValidUserRef(.TargetUser)))
        CrossNpcSpatialExit = True
    End With
    Exit Function

CrossFailed:
    Call TraceError(Err.Number, Err.Description, "modNpcCrossMapPursuit.CrossNpcSpatialExit", Erl)
End Function

Public Sub RecordNpcCrossMapPathFailure(ByVal NpcIndex As Integer)
    If NpcIndex <= 0 Or NpcIndex > LastNPC Then Exit Sub
    With NpcList(NpcIndex).CrossMapRoute
        If .Mode = eNpcCrossMapRouteNone Then Exit Sub
        If .PathFailures < 255 Then .PathFailures = .PathFailures + 1
        If .PathFailures >= NPC_CROSS_MAP_MAX_PATH_FAILURES Then
            If .Mode = eNpcCrossMapRouteChase Then
                Call ClearUserRef(NpcList(NpcIndex).TargetUser)
                Call BeginNpcReturnHome(NpcIndex, "pathfinding failure limit")
            Else
                Call LogNpcCrossMap(NpcIndex, "return path failed; warping home")
                Call WarpNpcChar(NpcIndex, NpcList(NpcIndex).Orig.Map, NpcList(NpcIndex).Orig.x, NpcList(NpcIndex).Orig.y, False)
                Call ResetNpcCrossMapRoute(NpcIndex)
            End If
        End If
    End With
End Sub

Public Sub ResetNpcCrossMapRoute(ByVal NpcIndex As Integer)
    If NpcIndex <= 0 Or NpcIndex > UBound(NpcList) Then Exit Sub
    With NpcList(NpcIndex).CrossMapRoute
        .Mode = eNpcCrossMapRouteNone
        .TargetMap = 0
        .NextMap = 0
        .ExitX = 0
        .ExitY = 0
        .DestinationX = 0
        .DestinationY = 0
        .HopsCrossed = 0
        .PathFailures = 0
    End With
End Sub

Private Function PrepareNpcNextHop(ByVal NpcIndex As Integer, ByVal TargetMap As Integer, ByVal IsChase As Boolean) As Boolean
    Dim nextMap As Integer
    Dim avoidSafeMaps As Boolean
    avoidSafeMaps = Not NpcCrossMapSafeMapsAllowed()
    If IsChase Then
        Dim originHop As Integer
        If Not TryGetNextSpatialHop(NpcList(NpcIndex).Orig.Map, TargetMap, originHop, NPC_CROSS_MAP_MAX_CHASE_HOPS, avoidSafeMaps) Then
            Call LogNpcCrossMap(NpcIndex, "route rejected reason=origin-leash originMap=" & NpcList(NpcIndex).Orig.Map & " targetMap=" & TargetMap)
            Exit Function
        End If
        If Not TryGetNextSpatialHop(NpcList(NpcIndex).pos.Map, TargetMap, nextMap, NPC_CROSS_MAP_MAX_CHASE_HOPS, avoidSafeMaps) Then
            Call LogNpcCrossMap(NpcIndex, "route rejected reason=no-route fromMap=" & NpcList(NpcIndex).pos.Map & " targetMap=" & TargetMap)
            Exit Function
        End If
    Else
        If Not TryGetNextSpatialHop(NpcList(NpcIndex).pos.Map, TargetMap, nextMap, , avoidSafeMaps) Then
            Call LogNpcCrossMap(NpcIndex, "route rejected reason=no-return-route fromMap=" & NpcList(NpcIndex).pos.Map & " targetMap=" & TargetMap)
            Exit Function
        End If
    End If

    With NpcList(NpcIndex).CrossMapRoute
        If .NextMap <> nextMap Or .ExitX = 0 Or .ExitY = 0 Then
            If Not ResolveNpcSpatialExit(NpcIndex, nextMap) Then
                Call LogNpcCrossMap(NpcIndex, "route rejected reason=no-valid-spatial-exit fromMap=" & NpcList(NpcIndex).pos.Map & " nextMap=" & nextMap)
                Exit Function
            End If
            .NextMap = nextMap
            .PathFailures = 0
            NpcList(NpcIndex).pathFindingInfo.PathLength = 0
            NpcList(NpcIndex).pathFindingInfo.NextPathRecomputeAt = 0
            Call LogNpcCrossMap(NpcIndex, "spatial exit selected fromMap=" & NpcList(NpcIndex).pos.Map & " exit=(" & .ExitX & "," & .ExitY & ") toMap=" & nextMap & " destination=(" & .DestinationX & "," & .DestinationY & ")")
        End If
    End With
    PrepareNpcNextHop = True
End Function

Private Function ResolveNpcSpatialExit(ByVal NpcIndex As Integer, ByVal NextMap As Integer) As Boolean
    Dim x As Integer
    Dim y As Integer
    Dim bestDistance As Long
    Dim candidateDistance As Long
    Dim destination As t_WorldPos
    bestDistance = &H7FFFFFFF

    For x = MinXBorder To MaxXBorder
        For y = MinYBorder To MaxYBorder
            destination = MapData(NpcList(NpcIndex).pos.Map, x, y).TileExit
            If destination.Map = NextMap Then
                If SpatialTransitionMatchesEdge(NpcList(NpcIndex).pos.Map, x, y, destination.Map, destination.x, destination.y) Then
                    If NpcDestinationSurfaceAllowed(NpcIndex, destination) Then
                        candidateDistance = Abs(CLng(NpcList(NpcIndex).pos.x) - x) + Abs(CLng(NpcList(NpcIndex).pos.y) - y)
                        If candidateDistance < bestDistance Then
                            bestDistance = candidateDistance
                            With NpcList(NpcIndex).CrossMapRoute
                                .ExitX = CByte(x)
                                .ExitY = CByte(y)
                                .DestinationX = CByte(destination.x)
                                .DestinationY = CByte(destination.y)
                            End With
                        End If
                    End If
                End If
            End If
        Next y
    Next x
    ResolveNpcSpatialExit = (bestDistance <> &H7FFFFFFF)
End Function

Private Function NpcDestinationSurfaceAllowed(ByVal NpcIndex As Integer, ByRef destination As t_WorldPos) As Boolean
    If Not InMapBounds(destination.Map, destination.x, destination.y) Then Exit Function
    With MapData(destination.Map, destination.x, destination.y)
        If (.Blocked And e_Block.ALL_SIDES) = e_Block.ALL_SIDES Then Exit Function
        If .trigger = e_Trigger.POSINVALIDA Then Exit Function
        If (.Blocked And FLAG_AGUA) <> 0 Then
            If NpcList(NpcIndex).flags.AguaValida = 0 Then Exit Function
        ElseIf NpcList(NpcIndex).flags.TierraInvalida <> 0 Then
            Exit Function
        End If
        If NpcList(NpcIndex).flags.LavaValida = 1 And Not HayLava(destination.Map, destination.x, destination.y) Then Exit Function
    End With
    NpcDestinationSurfaceAllowed = True
End Function

Private Function NpcDestinationAvailable(ByVal NpcIndex As Integer, ByRef destination As t_WorldPos) As Boolean
    If Not NpcDestinationSurfaceAllowed(NpcIndex, destination) Then Exit Function
    With MapData(destination.Map, destination.x, destination.y)
        If .NpcIndex <> 0 Or .UserIndex <> 0 Then Exit Function
    End With
    NpcDestinationAvailable = True
End Function

Private Function NpcCrossMapExitPosition(ByVal NpcIndex As Integer) As t_WorldPos
    NpcCrossMapExitPosition.Map = NpcList(NpcIndex).pos.Map
    NpcCrossMapExitPosition.x = NpcList(NpcIndex).CrossMapRoute.ExitX
    NpcCrossMapExitPosition.y = NpcList(NpcIndex).CrossMapRoute.ExitY
End Function

Private Function NpcCanPursueAcrossMaps(ByVal NpcIndex As Integer) As Boolean
    With NpcList(NpcIndex)
        If .Hostile <> 1 Then Exit Function
        If .Movement <> e_TipoAI.MueveAlAzar Then Exit Function
        If IsValidUserRef(.MaestroUser) Or IsValidNpcRef(.MaestroNPC) Then Exit Function
        If Not IsValidUserRef(.TargetUser) Then Exit Function
    End With
    NpcCanPursueAcrossMaps = True
End Function

Private Function NpcTargetAllowsRetainedChase(ByVal UserIndex As Integer) As Boolean
    If UserIndex <= 0 Or UserIndex > MaxUsers Then Exit Function
    With UserList(UserIndex)
        If Not .flags.UserLogged Then Exit Function
        If .flags.Muerto <> 0 Or .flags.EnConsulta <> 0 Then Exit Function
        If EsGM(UserIndex) And Not .flags.AdminPerseguible Then Exit Function
        If .pos.Map <= 0 Or .pos.Map > NumMaps Then Exit Function
        If Not NpcCrossMapDestinationAllowed(.pos.Map) Then Exit Function
    End With
    NpcTargetAllowsRetainedChase = True
End Function

Public Function NpcCrossMapMapAllowedByPolicy(ByVal MapId As Integer, ByVal AllowSafeMaps As Boolean) As Boolean
    If MapId <= 0 Or MapId > NumMaps Then Exit Function
    NpcCrossMapMapAllowedByPolicy = AllowSafeMaps Or MapInfo(MapId).Seguro = 0
End Function

Private Function NpcCrossMapDestinationAllowed(ByVal MapId As Integer) As Boolean
    NpcCrossMapDestinationAllowed = NpcCrossMapMapAllowedByPolicy(MapId, NpcCrossMapSafeMapsAllowed())
End Function

Private Function NpcCrossMapSafeMapsAllowed() As Boolean
    If SvrConfig Is Nothing Then Exit Function
    NpcCrossMapSafeMapsAllowed = CBool(SvrConfig.GetValue("NPC_CROSS_MAP_ALLOW_SAFE_MAPS"))
End Function

Private Sub BeginNpcReturnHome(ByVal NpcIndex As Integer, ByVal reason As String)
    With NpcList(NpcIndex)
        .pathFindingInfo.PathLength = 0
        .pathFindingInfo.NextPathRecomputeAt = 0
        If .pos.Map = .Orig.Map Then
            Call LogNpcCrossMap(NpcIndex, "chase cancelled reason=" & reason & " alreadyOnOriginMap=" & .Orig.Map)
            Call ResetNpcCrossMapRoute(NpcIndex)
        Else
            .CrossMapRoute.Mode = eNpcCrossMapRouteReturnHome
            .CrossMapRoute.TargetMap = .Orig.Map
            .CrossMapRoute.NextMap = 0
            .CrossMapRoute.ExitX = 0
            .CrossMapRoute.ExitY = 0
            .CrossMapRoute.DestinationX = 0
            .CrossMapRoute.DestinationY = 0
            .CrossMapRoute.PathFailures = 0
            Call LogNpcCrossMap(NpcIndex, "chase cancelled reason=" & reason & " returningToMap=" & .Orig.Map)
        End If
    End With
End Sub

Private Sub LogNpcCrossMap(ByVal NpcIndex As Integer, ByVal message As String)
    Call LogInfoServidor("NPC cross-map npc=" & NpcIndex & "/" & NpcList(NpcIndex).VersionId & " " & message)
End Sub

Private Function NpcCrossMapPursuitEnabled() As Boolean
    NpcCrossMapPursuitEnabled = IsFeatureEnabled(NPC_CROSS_MAP_PURSUIT_FEATURE) And AdjacentTopologyAvailable()
End Function
