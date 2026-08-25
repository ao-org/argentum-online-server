Attribute VB_Name = "Unit_NpcCrossMapPursuit"
Option Explicit
#If UNIT_TEST = 1 Then

Private Const TEST_TOPOLOGY As String = "{""version"":1,""id_space"":""resource_map_id"",""maps"":[34,35,36,78],""edges"":[[34,35,0,-80],[35,34,0,80],[35,36,74,0],[36,35,-74,0],[34,78,-74,0],[78,34,74,0]]}"

Public Function test_suite_npc_cross_map_pursuit() As Boolean
    Call UnitTesting.RunTest("npc topology loads verified edges", test_topology_load_and_edges())
    Call UnitTesting.RunTest("npc topology routes shortest next hop", test_topology_next_hop())
    Call UnitTesting.RunTest("npc topology routes multiple hops", test_topology_multi_hop())
    Call UnitTesting.RunTest("npc topology reports disconnected maps", test_topology_disconnected())
    Call UnitTesting.RunTest("npc topology validates spatial exit geometry", test_spatial_exit_geometry())
    Call UnitTesting.RunTest("npc spatial pursuit allows safe maps when configured", test_safe_destination_allowed())
    Call UnitTesting.RunTest("npc spatial pursuit blocks safe maps when configured", test_safe_destination_blocked())
    Call UnitTesting.RunTest("npc topology rejects malformed reciprocity", test_topology_rejects_invalid())
    Call UnitTesting.RunTest("npc topology checked-in manifest loads", test_checked_in_manifest())
    Call UnitTesting.RunTest("npc topology feature toggle disables routing", test_feature_toggle())
    Call UnitTesting.RunTest("npc spatial transition retains existing target only", test_target_retention())
    Call UnitTesting.RunTest("npc combat rejects targets on another map", test_cross_map_combat_guard())
    Call UnitTesting.RunTest("npc cross-map route keeps AI active on empty map", test_route_keeps_ai_active())
    Call UnitTesting.RunTest("npc respawn keeps original spawn map", test_respawn_uses_original_map())
    Call LoadAdjacentTopology
    test_suite_npc_cross_map_pursuit = True
End Function

Private Function test_safe_destination_allowed() As Boolean
    Dim originalToggle As Boolean
    Dim destination As t_WorldPos

    originalToggle = IsFeatureEnabled(NPC_CROSS_MAP_PURSUIT_FEATURE)
    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, True)
    If MapInfo(1).Seguro = 0 Then GoTo Cleanup

    destination.Map = 1
    destination.x = 31
    destination.y = 11
    test_safe_destination_allowed = IsVerifiedNpcSpatialTransition(5, 31, 91, destination, False)

Cleanup:
    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, originalToggle)
End Function

Private Function test_respawn_uses_original_map() As Boolean
    Dim npc As t_Npc
    npc.Orig.Map = 34
    npc.pos.Map = 35
    test_respawn_uses_original_map = (NpcRespawnMap(npc) = 34)
End Function

Private Function test_safe_destination_blocked() As Boolean
    If MapInfo(1).Seguro = 0 Then Exit Function
    test_safe_destination_blocked = Not NpcCrossMapMapAllowedByPolicy(1, False)
End Function

Private Function test_route_keeps_ai_active() As Boolean
    Dim npcIndex As Integer
    Dim originalUsers As Integer
    Dim originalForceUpdate As Boolean

    npcIndex = UBound(NpcList)
    originalUsers = MapInfo(34).NumUsers
    originalForceUpdate = MapInfo(34).ForceUpdate
    NpcList(npcIndex).pos.Map = 34
    MapInfo(34).NumUsers = 0
    MapInfo(34).ForceUpdate = False

    Call ResetNpcCrossMapRoute(npcIndex)
    If NpcRequiresAiUpdate(npcIndex) Then GoTo Cleanup
    NpcList(npcIndex).CrossMapRoute.Mode = eNpcCrossMapRouteChase
    test_route_keeps_ai_active = NpcRequiresAiUpdate(npcIndex)

Cleanup:
    Call ResetNpcCrossMapRoute(npcIndex)
    NpcList(npcIndex).pos.Map = 0
    MapInfo(34).NumUsers = originalUsers
    MapInfo(34).ForceUpdate = originalForceUpdate
End Function

Private Function test_target_retention() As Boolean
    Dim originalToggle As Boolean
    Dim originalLastNpc As Integer
    Dim npcIndex As Integer
    Dim userIndex As Integer

    originalToggle = IsFeatureEnabled(NPC_CROSS_MAP_PURSUIT_FEATURE)
    originalLastNpc = LastNPC
    npcIndex = UBound(NpcList)
    userIndex = UBound(UserList)
    If npcIndex <= originalLastNpc Then Exit Function

    UserList(userIndex).VersionId = 32000
    UserList(userIndex).flags.UserLogged = True
    UserList(userIndex).pos.Map = 35
    NpcList(npcIndex).pos.Map = 35
    NpcList(npcIndex).Orig.Map = 34
    NpcList(npcIndex).Hostile = 1
    NpcList(npcIndex).Movement = e_TipoAI.MueveAlAzar
    Call ClearUserRef(NpcList(npcIndex).MaestroUser)
    Call ClearNpcRef(NpcList(npcIndex).MaestroNPC)
    Call SetUserRef(NpcList(npcIndex).TargetUser, userIndex)
    LastNPC = npcIndex

    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, True)
    Call RetainNpcPursuitForSpatialTransition(userIndex, 34, 35)
    If NpcList(npcIndex).CrossMapRoute.Mode <> eNpcCrossMapRouteNone Then GoTo Cleanup

    NpcList(npcIndex).pos.Map = 34
    Call RetainNpcPursuitForSpatialTransition(userIndex, 34, 35)
    If NpcList(npcIndex).CrossMapRoute.Mode <> eNpcCrossMapRouteChase Then GoTo Cleanup
    If Not IsValidUserRef(NpcList(npcIndex).TargetUser) Then GoTo Cleanup
    If NpcList(npcIndex).CrossMapRoute.NextMap <> 35 Then GoTo Cleanup
    If NpcList(npcIndex).CrossMapRoute.ExitX = 0 Or NpcList(npcIndex).CrossMapRoute.ExitY = 0 Then GoTo Cleanup

    Call ClearUserRef(NpcList(npcIndex).TargetUser)
    Call ResetNpcCrossMapRoute(npcIndex)
    Call RetainNpcPursuitForSpatialTransition(userIndex, 34, 35)
    test_target_retention = (NpcList(npcIndex).CrossMapRoute.Mode = eNpcCrossMapRouteNone)

Cleanup:
    Call ClearUserRef(NpcList(npcIndex).TargetUser)
    Call ResetNpcCrossMapRoute(npcIndex)
    NpcList(npcIndex).pos.Map = 0
    NpcList(npcIndex).Orig.Map = 0
    NpcList(npcIndex).Hostile = 0
    NpcList(npcIndex).Movement = 0
    UserList(userIndex).flags.UserLogged = False
    UserList(userIndex).pos.Map = 0
    LastNPC = originalLastNpc
    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, originalToggle)
End Function

Private Function test_cross_map_combat_guard() As Boolean
    Dim npcIndex As Integer
    Dim userIndex As Integer
    Dim originalHp As Integer

    npcIndex = UBound(NpcList)
    userIndex = UBound(UserList)
    NpcList(npcIndex).pos.Map = 34
    UserList(userIndex).pos.Map = 35
    UserList(userIndex).Stats.MinHp = 100
    originalHp = UserList(userIndex).Stats.MinHp

    Call NpcDoDamageToUser(npcIndex, userIndex, 25, e_DamageSourceType.e_phisical, 0)
    Call NpcLanzaSpellSobreUser(npcIndex, userIndex, 1)
    Call NpcAtacaUser(npcIndex, userIndex, e_Heading.NORTH)
    test_cross_map_combat_guard = (UserList(userIndex).Stats.MinHp = originalHp)

    NpcList(npcIndex).pos.Map = 0
    UserList(userIndex).pos.Map = 0
    UserList(userIndex).Stats.MinHp = 0
End Function

Private Function test_checked_in_manifest() As Boolean
    Dim topologyPath As String
    Dim dx As Integer
    Dim dy As Integer
    topologyPath = App.Path & "\Recursos\config\legacy_overworld_topology.json"
    If Not FileExist(topologyPath, vbNormal) Then Exit Function
    If Not LoadAdjacentTopologyJson(FileText(topologyPath), False) Then Exit Function
    If Not TryGetAdjacentEdge(34, 35, dx, dy) Then Exit Function
    test_checked_in_manifest = (dx = 0 And dy = -80 And AdjacentTopologyContainsMap(557))
End Function

Private Function test_feature_toggle() As Boolean
    Dim originalValue As Boolean
    originalValue = IsFeatureEnabled(NPC_CROSS_MAP_PURSUIT_FEATURE)
    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, True)
    If Not AdjacentTopologyAvailable() Then GoTo RestoreToggle
    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, False)
    test_feature_toggle = Not AdjacentTopologyAvailable()
RestoreToggle:
    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, originalValue)
End Function

Private Function test_topology_load_and_edges() As Boolean
    Dim dx As Integer
    Dim dy As Integer
    If Not LoadAdjacentTopologyJson(TEST_TOPOLOGY, False) Then Exit Function
    If Not AdjacentTopologyContainsMap(34) Then Exit Function
    If Not TryGetAdjacentEdge(34, 35, dx, dy) Then Exit Function
    test_topology_load_and_edges = (dx = 0 And dy = -80)
End Function

Private Function test_topology_next_hop() As Boolean
    Dim nextMap As Integer
    If Not LoadAdjacentTopologyJson(TEST_TOPOLOGY, False) Then Exit Function
    If Not TryGetNextSpatialHop(34, 78, nextMap) Then Exit Function
    test_topology_next_hop = (nextMap = 78)
End Function

Private Function test_topology_multi_hop() As Boolean
    Dim nextMap As Integer
    If Not LoadAdjacentTopologyJson(TEST_TOPOLOGY, False) Then Exit Function
    If Not TryGetNextSpatialHop(34, 36, nextMap) Then Exit Function
    test_topology_multi_hop = (nextMap = 35)
End Function

Private Function test_topology_disconnected() As Boolean
    Dim nextMap As Integer
    If Not LoadAdjacentTopologyJson(TEST_TOPOLOGY, False) Then Exit Function
    test_topology_disconnected = Not TryGetNextSpatialHop(36, 78, nextMap, 1)
End Function

Private Function test_spatial_exit_geometry() As Boolean
    If Not LoadAdjacentTopologyJson(TEST_TOPOLOGY, False) Then Exit Function
    If Not SpatialTransitionMatchesEdge(34, 50, 10, 35, 50, 90) Then Exit Function
    If SpatialTransitionMatchesEdge(34, 50, 10, 35, 51, 90) Then Exit Function
    test_spatial_exit_geometry = SpatialTransitionMatchesEdge(35, 50, 91, 34, 50, 11)
End Function

Private Function test_topology_rejects_invalid() As Boolean
    Dim invalidJson As String
    invalidJson = "{""version"":1,""id_space"":""resource_map_id"",""maps"":[1,2],""edges"":[[1,2,74,0]]}"
    test_topology_rejects_invalid = Not LoadAdjacentTopologyJson(invalidJson, False)
    test_topology_rejects_invalid = test_topology_rejects_invalid And Not AdjacentTopologyAvailable()
End Function

#End If
