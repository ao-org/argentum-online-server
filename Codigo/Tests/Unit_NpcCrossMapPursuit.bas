Attribute VB_Name = "Unit_NpcCrossMapPursuit"
Option Explicit
#If UNIT_TEST = 1 Then

Private Const TEST_TOPOLOGY As String = "{""version"":1,""id_space"":""resource_map_id"",""maps"":[34,35,36,78],""edges"":[[34,35,0,-80],[35,34,0,80],[35,36,74,0],[36,35,-74,0],[34,78,-74,0],[78,34,74,0]]}"
Private Const ADJACENT_TEST_TOPOLOGY As String = "{""version"":1,""id_space"":""resource_map_id"",""maps"":[34,35],""edges"":[[34,35,0,-80],[35,34,0,80]]}"
Private adjacentCaptureActive As Boolean
Private adjacentPackets As Collection
Private adjacentReaderBytes() As Byte
Private adjacentAccepted As Collection
Private adjacentFailAttempt As Long
Private adjacentCaptureError As String
Private adjacentExpectedRoute As SendTarget
Private adjacentExpectedRecipient As Integer
Private adjacentTick As Long
Private adjacentNpc As Integer
Private diagnosticsCapture As Boolean
Private diagnosticsCount As Long
Private diagnosticsOverflow As Boolean
Private diagnosticsText(1 To 32) As String
Private diagnosticsSeverity(1 To 32) As LogEventTypeConstants

' Observe only while these tests run; LogThis still writes the real Windows event.
Public Sub ObserveDiagnosticLog(ByVal text As String, ByVal severity As LogEventTypeConstants)
    If Not diagnosticsCapture Then Exit Sub
    If diagnosticsCount = UBound(diagnosticsText) Then
        diagnosticsOverflow = True
        Exit Sub
    End If
    diagnosticsCount = diagnosticsCount + 1
    diagnosticsText(diagnosticsCount) = Left$(text, 512)
    diagnosticsSeverity(diagnosticsCount) = severity
End Sub

Private Sub ClearDiagnosticLogs()
    diagnosticsCount = 0
    diagnosticsOverflow = False
End Sub

Private Sub ExpectDiagnosticLog(ByVal fragment As String, ByVal expectedCount As Long, Optional ByVal severity As LogEventTypeConstants = vbLogEventTypeInformation)
    Dim i As Long, found As Long
    For i = 1 To diagnosticsCount
        If InStr(1, diagnosticsText(i), fragment, vbBinaryCompare) > 0 Then
            found = found + 1
            Call AdjacentExpect(diagnosticsSeverity(i) = severity, "diagnostic severity: " & fragment)
        End If
    Next i
    Call AdjacentExpect(Not diagnosticsOverflow And found = expectedCount, "diagnostic count: " & fragment)
End Sub

Private Sub LoadDiagnosticsSetting(Optional ByVal value As String = vbNullString)
    Dim source As New clsIniManager
    Call source.ChangeValue("AI", "NPCCrossMapAllowSafeMaps", "0")
    If value <> vbNullString Then Call source.ChangeValue("AI", "NPCCrossMapDiagnostics", value)
    Set SvrConfig = New ServerConfig
    Call AdjacentExpect(SvrConfig.LoadSettings(vbNullString, source) = 0, "load diagnostic fixture settings")
    Call source.ChangeValue("AI", "NPCCrossMapDiagnostics", "1")
    ' Changing the source after loading must not change the cached setting.
    Call AdjacentExpect(NpcCrossMapDiagnosticsEnabled() = (value = "1"), "cached diagnostic setting")
End Sub

Private Sub CheckDiagnosticsCase(ByVal group As Long)
    Dim i As Long, expectedCount As Long
    LastNPC = adjacentNpc
    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, True)
    Call AdjacentExpect(LoadAdjacentTopologyJson(ADJACENT_TEST_TOPOLOGY, False), "diagnostic fixture after feature reload")
    With NpcList(adjacentNpc)
        .Char.charindex = 0
        .Orig = .pos
        .Hostile = 1
        .Movement = e_TipoAI.MueveAlAzar
    End With
    diagnosticsCapture = True
    Call ClearDiagnosticLogs()
    Select Case group
        Case 42
            Set SvrConfig = Nothing
            Call AdjacentExpect(Not NpcCrossMapDiagnosticsEnabled(), "missing configuration is off")
            Call LoadDiagnosticsSetting()
            Call LoadDiagnosticsSetting("0")
            Call LoadDiagnosticsSetting("1")
            Call LoadDiagnosticsSetting("not-a-number")
            Call LoadDiagnosticsSetting("1")
            Call LoadDiagnosticsSetting("0")
        Case 43
            For i = 0 To 2
                expectedCount = 0
                If i = 1 Then expectedCount = 1
                Call LoadDiagnosticsSetting(CStr(expectedCount))
                UserList(1).flags.Muerto = 0
                Call ClearDiagnosticLogs()
                Call RetainNpcPursuitForSpatialTransition(1, 35, 34)
                Call ExpectDiagnosticLog("matchingTargets=0 sourceMapTargets=0 retained=0", expectedCount)
                Call AdjacentExpect(diagnosticsCount = expectedCount, "zero-retention logging gate")

                With MapData(50, 90, 35).TileExit
                    .Map = 34: .x = 50: .y = 10
                End With
                Call SetUserRef(NpcList(adjacentNpc).TargetUser, 1)
                Call ClearDiagnosticLogs()
                Call RetainNpcPursuitForSpatialTransition(1, 35, 34)
                Call AdjacentExpect(NpcList(adjacentNpc).CrossMapRoute.Mode = eNpcCrossMapRouteChase And NpcList(adjacentNpc).CrossMapRoute.NextMap = 34, "retained chase unchanged by diagnostics")
                Call ExpectDiagnosticLog("chase retained fromMap=35 targetMap=34", expectedCount)
                Call ExpectDiagnosticLog("spatial exit selected fromMap=35 exit=(50,90) toMap=34 destination=(50,10)", expectedCount)
                Call ExpectDiagnosticLog("matchingTargets=1 sourceMapTargets=1 retained=1", expectedCount)
                Call AdjacentExpect(diagnosticsCount = 3 * expectedCount, "retained logging gate")

                UserList(1).flags.Muerto = 1
                Call ClearDiagnosticLogs()
                Call AdjacentExpect(Not HandleNpcCrossMapRoute(adjacentNpc), "cancel at origin does not move")
                Call AdjacentExpect(NpcList(adjacentNpc).CrossMapRoute.Mode = eNpcCrossMapRouteNone And Not IsValidUserRef(NpcList(adjacentNpc).TargetUser), "cancellation state identical off/on/off")
                Call ExpectDiagnosticLog("chase cancelled reason=target no longer attackable alreadyOnOriginMap=35", expectedCount)
                Call AdjacentExpect(diagnosticsCount = expectedCount, "cancellation logging gate")

                UserList(1).flags.Muerto = 0
                MapData(50, 90, 35).TileExit.Map = 0
                Call SetUserRef(NpcList(adjacentNpc).TargetUser, 1)
                NpcList(adjacentNpc).CrossMapRoute.Mode = eNpcCrossMapRouteChase
                Call ClearDiagnosticLogs()
                Call AdjacentExpect(Not HandleNpcCrossMapRoute(adjacentNpc), "missing exit rejects route")
                Call AdjacentExpect(NpcList(adjacentNpc).CrossMapRoute.Mode = eNpcCrossMapRouteNone And Not IsValidUserRef(NpcList(adjacentNpc).TargetUser), "rejection state independent of diagnostics")
                Call ExpectDiagnosticLog("route rejected reason=no-valid-spatial-exit", expectedCount)
                Call ExpectDiagnosticLog("chase cancelled reason=no allowed chase route", expectedCount)
                Call AdjacentExpect(diagnosticsCount = 2 * expectedCount, "rejection logging gate")
            Next i
        Case 44
            Call LoadDiagnosticsSetting("0")
            ' The real route error handler must remain visible; no valid NPC is touched.
            LastNPC = UBound(NpcList) + 1
            Call HandleNpcCrossMapRoute(LastNPC)
            LastNPC = adjacentNpc
            Call ExpectDiagnosticLog("Component: modNpcCrossMapPursuit.HandleNpcCrossMapRoute", 1, vbLogEventTypeError)
            Call ExpectDiagnosticLog("Error number: 9", 1, vbLogEventTypeError)
        Case 45
            Call LoadDiagnosticsSetting("0")
            Call ClearAdjacentCapture(0, ToNPCArea, adjacentNpc)
            For i = 1 To 2
                Call ClearDiagnosticLogs()
                With NpcList(adjacentNpc)
                    .Orig.Map = 33 + i: .Orig.x = 50: .Orig.y = 50
                    .CrossMapRoute.Mode = eNpcCrossMapRouteReturnHome
                    If i = 1 Then
                        Call AdjacentExpect(HandleNpcCrossMapRoute(adjacentNpc), "missing return exit uses fallback")
                        Call ExpectDiagnosticLog("return route unavailable; warping home", 1)
                    Else
                        .CrossMapRoute.PathFailures = 2
                        Call RecordNpcCrossMapPathFailure(adjacentNpc)
                        Call ExpectDiagnosticLog("return path failed; warping home", 1)
                    End If
                    Call AdjacentExpect(.pos.Map = .Orig.Map And .pos.x = 50 And .pos.y = 50 And .CrossMapRoute.Mode = eNpcCrossMapRouteNone, "fallback really warps home and resets route")
                End With
                Call AdjacentExpect(diagnosticsCount = 1, "only unconditional Info fallback logged")
            Next i
    End Select
End Sub

Public Function CaptureAdjacentPacket(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByRef Queued As Boolean) As Boolean
    If Not adjacentCaptureActive Then Exit Function
    On Error GoTo Failed
    CaptureAdjacentPacket = True
    Queued = False
    Dim writer As Network.Writer, bytes() As Byte, packet As Variant
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    packet = bytes
    Call adjacentPackets.Add(packet)
    Queued = adjacentPackets.Count <> adjacentFailAttempt
    Call adjacentAccepted.Add(Queued)
    If sndRoute <> adjacentExpectedRoute Or sndIndex <> adjacentExpectedRecipient Then adjacentCaptureError = "unexpected recipient or route"
    Exit Function
Failed:
    adjacentCaptureError = CStr(Err.Number) & ": " & Err.Description
    Queued = False
    If Not writer Is Nothing Then Call writer.Clear()
End Function

Public Function test_suite_npc_cross_map_pursuit() As Boolean
    Dim failuresBefore As Long, names As Variant, group As Long
    failuresBefore = UnitTesting.GetFailedTests()
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
    Call UnitTesting.RunTest("npc unrestricted respawn works in both periods", test_respawn_schedule_any_time())
    Call UnitTesting.RunTest("npc day respawn only works during daytime", test_respawn_schedule_day_only())
    Call UnitTesting.RunTest("npc night respawn only works during nighttime", test_respawn_schedule_night_only())
    Call UnitTesting.RunTest("npc respawn day boundaries are inclusive", test_respawn_schedule_boundaries())
    Call UnitTesting.RunTest("npc invalid respawn schedule remains compatible", test_respawn_schedule_invalid())
    Call UnitTesting.RunTest("npc scheduled initial spawn waits for a free origin", test_scheduled_initial_spawn_requires_free_origin())
    Call UnitTesting.RunTest("adjacent visibility includes north seam", AdjacentNpcInsideInterest(50, 11, 50, 90, 0, -80))
    Call UnitTesting.RunTest("adjacent visibility includes west seam", AdjacentNpcInsideInterest(14, 50, 87, 50, -74, 0))
    Call UnitTesting.RunTest("adjacent visibility rejects distant NPC", Not AdjacentNpcInsideInterest(50, 50, 50, 50, 74, 0))
    Call UnitTesting.RunTest("adjacent visibility includes radius boundary", AdjacentNpcInsideInterest(50, 50, 74, 68, 0, 0))
    Call UnitTesting.RunTest("adjacent visibility excludes outside radius", Not AdjacentNpcInsideInterest(50, 50, 75, 68, 0, 0))
    names = Split("wire and unchanged silence|initial hidden|visible to hidden|hidden to visible|unsafe dead viewer reviver only|safe dead viewer|hidden owned NPC|hidden NPC staff viewer|reset failure|initial upsert failure|changed upsert failure|remove failure|reuse failed remove|reuse ordering|partial upsert success|partial remove success|viewer version reset|viewer map reset|explicit same map reset|LastUser shrink|disconnect cleanup|invisible reviver|visible pets and summons|zero body|reuse failed replacement|version during failed upsert|map during failed remove|connection during failed upsert", "|")
    For group = 0 To UBound(names)
        Call RunAdjacentCase(group + 1, CStr(names(group)))
    Next group
    names = Split("normal and adjacent visual parity|definition refresh|runtime name refresh|runtime description refresh|showName hide and reveal|empty display text|UTF-8 varuint boundary|maximum display text|oversized text recovery|failed display refresh retry", "|")
    For group = 0 To UBound(names)
        Call RunAdjacentCase(group + 29, CStr(names(group)))
    Next group
    Call RunAdjacentCase(39, "stationary South idle changes preserve equipment")
    Call RunAdjacentCase(40, "entrance guard AI same-heading refresh preserves equipment")
    Call RunAdjacentCase(41, "character change preserves body semantics and current equipment")
    Call RunAdjacentCase(42, "diagnostics cached configuration defaults and reload")
    Call RunAdjacentCase(43, "diagnostics retention rejection cancellation off-on-off")
    Call RunAdjacentCase(44, "diagnostics off preserves route error reporting")
    Call RunAdjacentCase(45, "diagnostics off preserves both warp-home fallback logs")
    names = Split("exact connection and length|wrong connection|partial enqueue|oversized enqueue|late callback|nested probe preserves outer|same connection closes|other connection closes|duplicate callback", "|")
    For group = 0 To UBound(names)
        Call UnitTesting.RunTest("adjacent queue " & CStr(names(group)), modNetwork.TestQueueAcceptanceProbe(CInt(group + 1)))
    Next group
    names = Split("checked-in diagonal map transitions|southeast boundaries|southwest boundaries|northeast boundaries|northwest boundaries|cycles and two edge cap|ambiguous source excluded|positive wide composition|negative wide composition|instance and visibility retain|diagonal retry and map reset", "|")
    For group = 0 To UBound(names)
        Call RunAdjacentTwoHopCase(group + 1, CStr(names(group)))
    Next group
    Call TestNpcRespawnEffects()
    Call LoadAdjacentTopology
    test_suite_npc_cross_map_pursuit = UnitTesting.GetFailedTests() = failuresBefore
End Function

Private Sub AdjacentExpect(ByVal condition As Boolean, ByVal description As String)
    If Not condition Then Err.Raise vbObjectError + 183, "Adjacent NPC regression", description
End Sub

Private Sub ClearAdjacentCapture(Optional ByVal failAttempt As Long = 0, Optional ByVal expectedRoute As SendTarget = ToIndex, Optional ByVal expectedRecipient As Integer = 1)
    Set adjacentPackets = New Collection
    Set adjacentAccepted = New Collection
    adjacentFailAttempt = failAttempt
    adjacentCaptureError = vbNullString
    adjacentExpectedRoute = expectedRoute
    adjacentExpectedRecipient = expectedRecipient
End Sub

Private Function AdjacentPacketReader(ByVal index As Long, ByVal packetId As Integer, ByVal length As Long) As Network.reader
    Dim reader As New Network.reader
    ' Aurora readers borrow their input; retain bytes after this helper returns.
    adjacentReaderBytes = adjacentPackets(index)
    Call AdjacentExpect(UBound(adjacentReaderBytes) - LBound(adjacentReaderBytes) + 1 = length, "exact wire length")
    Call reader.SetData(adjacentReaderBytes)
    Call AdjacentExpect(reader.ReadInt16() = packetId, "packet order/ID")
    Set AdjacentPacketReader = reader
End Function

Private Sub ExpectAdjacentActor(ByVal index As Long, ByVal packetId As Integer, ByVal npcIndex As Integer, ByVal version As Integer, ByVal accepted As Boolean)
    Dim reader As Network.reader, length As Long
    length = 10
    If packetId = ServerPacketID.eAdjacentNpcUpsert Then length = 35
    Set reader = AdjacentPacketReader(index, packetId, length)
    Call AdjacentExpect(reader.ReadInt32() = AdjacentNpcTestEpoch(1), "current epoch")
    Call AdjacentExpect(reader.ReadInt16() = npcIndex, "exact NPC index")
    Call AdjacentExpect(reader.ReadInt16() = version, "exact NPC lifetime")
    Call AdjacentExpect(CBool(adjacentAccepted(index)) = accepted, "enqueue outcome")
End Sub

Private Sub TickAdjacent()
    adjacentTick = adjacentTick + 100
    Call UpdateAdjacentNpcVisibility(adjacentTick)
    Call AdjacentExpect(adjacentCaptureError = vbNullString, "capture: " & adjacentCaptureError)
    ' A marker after every production update catches leaked shared-writer bytes.
    Dim writer As Network.Writer, bytes() As Byte
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(&H6A5B)
    Call writer.GetData(bytes)
    Call writer.Clear()
    Call AdjacentExpect(UBound(bytes) = 1, "shared writer was not cleared")
    Call AdjacentExpect(bytes(0) = &H5B And bytes(1) = &H6A, "writer marker")
End Sub

Private Sub AddAdjacentSecondNpc()
    NpcList(adjacentNpc + 1) = NpcList(adjacentNpc)
    With NpcList(adjacentNpc + 1)
        .pos.x = 51
        .Char.charindex = 124
        .npcType = e_NPCType.Comun
    End With
    MapData(51, 90, 35).NpcIndex = adjacentNpc + 1
End Sub

Private Sub CheckAdjacentCase(ByVal group As Long)
    Dim reader As Network.reader, epoch As Long, bytes() As Byte, i As Long
    Select Case group
        Case 1
            With NpcList(adjacentNpc).Char
                .head = 267: .WeaponAnim = 524: .ShieldAnim = 781
                .CascoAnim = 1038: .CartAnim = 1295
            End With
            NpcList(adjacentNpc).flags.NPCIdle = True
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2, "reset then one upsert")
            Set reader = AdjacentPacketReader(1, ServerPacketID.eAdjacentNpcReset, 8)
            Call AdjacentExpect(reader.ReadInt32() = AdjacentNpcTestEpoch(1), "reset epoch")
            Call AdjacentExpect(reader.ReadInt16() = 34, "viewer map")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            Set reader = AdjacentPacketReader(2, ServerPacketID.eAdjacentNpcUpsert, 35)
            Call reader.ReadInt32(): Call reader.ReadInt16(): Call reader.ReadInt16()
            Call AdjacentExpect(reader.ReadInt16() = 35, "source map")
            Call AdjacentExpect(reader.ReadInt8() = 50 And reader.ReadInt8() = 90, "source position")
            Call AdjacentExpect(reader.ReadInt16() = 10 And reader.ReadInt16() = 267, "body/head widths")
            Call AdjacentExpect(reader.ReadInt8() = e_Heading.NORTH, "heading")
            Call AdjacentExpect(reader.ReadInt16() = 524 And reader.ReadInt16() = 781, "weapon/shield widths")
            Call AdjacentExpect(reader.ReadInt16() = 1038 And reader.ReadInt16() = 1295, "helmet/cart widths")
            Call AdjacentExpect(reader.ReadReal32() = 1 And reader.ReadBool(), "speed and idle")
            Call AdjacentExpect(reader.ReadInt16() = 1543 And reader.ReadString8() = vbNullString, "definition and empty display text")
            Call AdjacentExpect(reader.GetAvailable() = 0, "complete payload consumed")
            #If PYMMO = 1 Then
                Call AdjacentExpect(ServerPacketID.eAdjacentNpcReset = 214 And ServerPacketID.eAdjacentNpcUpsert = 215 And ServerPacketID.eAdjacentNpcRemove = 216, "HOO literal packet IDs")
            #Else
                Call AdjacentExpect(ServerPacketID.eAdjacentNpcReset = 215 And ServerPacketID.eAdjacentNpcUpsert = 216 And ServerPacketID.eAdjacentNpcRemove = 217, "PYMMO=0 ordinals")
            #End If
            bytes = adjacentPackets(2)
            Call AdjacentExpect(bytes(10) = 35 And bytes(11) = 0 And bytes(16) = 11 And bytes(17) = 1, "little endian fields")
            Call AdjacentExpect(bytes(27) = 0 And bytes(28) = 0 And bytes(29) = 128 And bytes(30) = 63 And bytes(31) = 1, "Single and Boolean representation")
            Call AdjacentExpect(bytes(32) = 7 And bytes(33) = 6 And bytes(34) = 0, "definition little endian and empty String8 prefix")
            Call AdjacentExpect(AdjacentNpcTestKnownCount(1) = 1 And IsAdjacentNpcObserved(adjacentNpc), "accepted cache and AI observation")
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 0, "unchanged state is silent")
        Case 2, 7, 8, 22, 24
            NpcList(adjacentNpc).flags.invisible = 1
            If group = 7 Then Call SetUserRef(NpcList(adjacentNpc).MaestroUser, 1)
            If group = 8 Then UserList(1).flags.Privilegios = e_PlayerType.Admin
            If group = 22 Then
                NpcList(adjacentNpc).npcType = e_NPCType.Revividor
                UserList(1).flags.Muerto = 1
            End If
            If group = 24 Then
                NpcList(adjacentNpc).flags.invisible = 0
                NpcList(adjacentNpc).Char.body = 0
            End If
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1, "excluded actor sends reset only")
            Call AdjacentExpect(AdjacentNpcTestKnownCount(1) = 0, "excluded actor not cached")
            Call AdjacentExpect(Not IsAdjacentNpcObserved(adjacentNpc), "excluded actor not observed")
        Case 3, 4
            If group = 4 Then NpcList(adjacentNpc).flags.invisible = 1
            Call TickAdjacent
            Call ClearAdjacentCapture
            If group = 3 Then
                NpcList(adjacentNpc).flags.invisible = 1
            Else
                NpcList(adjacentNpc).flags.invisible = 0
            End If
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1, "one visibility transition")
            If group = 3 Then
                Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc, 3, True)
                Call AdjacentExpect(AdjacentNpcTestKnownCount(1) = 0 And Not IsAdjacentNpcObserved(adjacentNpc), "hide retires observation")
            Else
                Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
                Call AdjacentExpect(AdjacentNpcTestKnownCount(1) = 1 And IsAdjacentNpcObserved(adjacentNpc), "visible actor restored")
            End If
        Case 5
            Call AddAdjacentSecondNpc
            NpcList(adjacentNpc + 1).npcType = e_NPCType.Revividor
            UserList(1).flags.Muerto = 1
            MapInfo(35).Seguro = 1 ' Viewer map, not source map, controls ghost visibility.
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2, "dead unsafe sees only reviver")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc + 1, 3, True)
            Call AdjacentExpect(AdjacentNpcTestKnownVersion(1, adjacentNpc) = -1, "ordinary NPC excluded")
        Case 6
            UserList(1).flags.Muerto = 1
            MapInfo(34).Seguro = 1
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2, "safe dead viewer sees ordinary NPC")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
        Case 9
            adjacentFailAttempt = 1
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1 And Not CBool(adjacentAccepted(1)), "failed reset stops snapshot")
            Call AdjacentExpect(AdjacentNpcTestEpoch(1) = 0 And AdjacentNpcTestKnownCount(1) = 0, "failed reset does not initialize cache")
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2 And AdjacentNpcTestEpoch(1) > 0, "retry reset then snapshot")
            Set reader = AdjacentPacketReader(1, ServerPacketID.eAdjacentNpcReset, 8)
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
        Case 10
            adjacentFailAttempt = 2
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2, "reset accepted, upsert attempted")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, False)
            epoch = AdjacentNpcTestEpoch(1)
            Call AdjacentExpect(epoch > 0 And AdjacentNpcTestKnownCount(1) = 0, "no cache commit for rejected upsert")
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestEpoch(1) = epoch, "retry upsert without reset")
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
        Case 11
            Call TickAdjacent
            NpcList(adjacentNpc).Char.body = 11
            Call ClearAdjacentCapture(1)
            Call TickAdjacent
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, False)
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1, "unchanged failed update retried")
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            bytes = adjacentPackets(1)
            Call AdjacentExpect(bytes(14) = 11, "retried new appearance")
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 0, "successful update cached")
        Case 12
            Call TickAdjacent
            NpcList(adjacentNpc).flags.invisible = 1
            Call ClearAdjacentCapture(1)
            Call TickAdjacent
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc, 3, False)
            Call AdjacentExpect(AdjacentNpcTestKnownVersion(1, adjacentNpc) = 3, "failed remove retains delivered lifetime")
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1, "failed removal retried")
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc, 3, True)
            Call AdjacentExpect(AdjacentNpcTestKnownCount(1) = 0, "successful removal retired")
        Case 13, 14, 25
            Call TickAdjacent
            NpcList(adjacentNpc).VersionId = 4
            If group = 13 Then
                Call ClearAdjacentCapture(1)
                Call TickAdjacent
                Call AdjacentExpect(adjacentPackets.Count = 1, "failed removal prevents replacement upsert")
                Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc, 3, False)
                Call AdjacentExpect(AdjacentNpcTestKnownVersion(1, adjacentNpc) = 3, "old lifetime retained")
            End If
            If group = 25 Then
                Call ClearAdjacentCapture(2)
            Else
                Call ClearAdjacentCapture
            End If
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2, "reuse removes before upsert")
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc, 3, True)
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 4, group <> 25)
            If group = 25 Then
                Call AdjacentExpect(AdjacentNpcTestKnownVersion(1, adjacentNpc) = -1, "accepted removal survives failed replacement")
                Call ClearAdjacentCapture
                Call TickAdjacent
                Call AdjacentExpect(adjacentPackets.Count = 1, "only replacement retried")
                Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 4, True)
            End If
            Call AdjacentExpect(AdjacentNpcTestKnownVersion(1, adjacentNpc) = 4, "replacement committed")
        Case 15
            Call AddAdjacentSecondNpc
            adjacentFailAttempt = 3
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 3, "partial initial snapshot")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            Call ExpectAdjacentActor(3, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc + 1, 3, False)
            Call AdjacentExpect(AdjacentNpcTestKnownCount(1) = 1, "accepted prefix retained")
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1, "accepted prefix not resent")
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc + 1, 3, True)
            Call AdjacentExpect(AdjacentNpcTestKnownCount(1) = 2, "both delivered actors cached")
        Case 16
            Call AddAdjacentSecondNpc
            Call TickAdjacent
            NpcList(adjacentNpc).flags.invisible = 1
            NpcList(adjacentNpc + 1).flags.invisible = 1
            Call ClearAdjacentCapture(2)
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2, "two removals attempted")
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc, 3, True)
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcRemove, adjacentNpc + 1, 3, False)
            Call AdjacentExpect(AdjacentNpcTestKnownVersion(1, adjacentNpc) = -1 And AdjacentNpcTestKnownVersion(1, adjacentNpc + 1) = 3, "partial removals committed independently")
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1, "only failed removal retried")
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc + 1, 3, True)
        Case 17, 18, 19
            Call TickAdjacent
            epoch = AdjacentNpcTestEpoch(1)
            If group = 17 Then UserList(1).VersionId = 8
            If group = 18 Then UserList(1).pos.Map = 35
            If group = 19 Then Call InvalidateAdjacentNpcViewer(1)
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(AdjacentNpcTestEpoch(1) > epoch, "new subscription epoch")
            Set reader = AdjacentPacketReader(1, ServerPacketID.eAdjacentNpcReset, 8)
            Call reader.ReadInt32()
            Call AdjacentExpect(reader.ReadInt16() = UserList(1).pos.Map, "reset map")
            If group = 18 Then
                Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestKnownCount(1) = 0, "source becomes local, no adjacent duplicate")
            Else
                Call AdjacentExpect(adjacentPackets.Count = 2, "new subscription snapshot")
                Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            End If
        Case 20, 21
            Call TickAdjacent
            epoch = AdjacentNpcTestEpoch(1)
            If group = 20 Then
                LastUser = 0
            Else
                UserList(1).ConnectionDetails.ConnIDValida = False
            End If
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 0 And AdjacentNpcTestKnownCount(1) = 0 And AdjacentNpcTestEpoch(1) = 0, "retired viewer releases cache without packets")
            Call AdjacentExpect(Not IsAdjacentNpcObserved(adjacentNpc), "retired viewer no longer drives AI")
            LastUser = 1
            UserList(1).ConnectionDetails.ConnIDValida = True
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2 And AdjacentNpcTestEpoch(1) > epoch, "reconnect creates fresh snapshot")
        Case 23
            Call AddAdjacentSecondNpc
            Call SetUserRef(NpcList(adjacentNpc).MaestroUser, 1)
            Call SetNpcRef(NpcList(adjacentNpc + 1).MaestroNPC, adjacentNpc)
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 3 And AdjacentNpcTestKnownCount(1) = 2, "visible pets and summons retained")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            Call ExpectAdjacentActor(3, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc + 1, 3, True)
        Case 26, 28
            adjacentFailAttempt = 2
            Call TickAdjacent
            epoch = AdjacentNpcTestEpoch(1)
            Call AdjacentExpect(epoch > 0 And AdjacentNpcTestKnownCount(1) = 0, "failed snapshot awaiting retry")
            If group = 26 Then
                UserList(1).VersionId = 8
            Else
                UserList(1).ConnectionDetails.ConnID = 43
            End If
            NpcList(adjacentNpc).Char.body = 11
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2 And AdjacentNpcTestEpoch(1) > epoch, "new identity resets instead of retrying old epoch")
            Set reader = AdjacentPacketReader(1, ServerPacketID.eAdjacentNpcReset, 8)
            Call AdjacentExpect(reader.ReadInt32() = AdjacentNpcTestEpoch(1), "replacement reset epoch")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            bytes = adjacentPackets(2)
            Call AdjacentExpect(bytes(14) = 11, "new snapshot uses current appearance")
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 0, "obsolete retry does not reappear")
        Case 27
            Call TickAdjacent
            epoch = AdjacentNpcTestEpoch(1)
            NpcList(adjacentNpc).flags.invisible = 1
            Call ClearAdjacentCapture(1)
            Call TickAdjacent
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc, 3, False)
            UserList(1).pos.Map = 35
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestEpoch(1) > epoch, "map replacement discards failed old removal")
            Set reader = AdjacentPacketReader(1, ServerPacketID.eAdjacentNpcReset, 8)
            Call reader.ReadInt32()
            Call AdjacentExpect(reader.ReadInt16() = 35 And AdjacentNpcTestKnownCount(1) = 0, "new map starts empty")
            Call ClearAdjacentCapture
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 0, "no stale removal replay on new map")
        Case 29 To 38
            Call CheckAdjacentAppearanceCase(group)
        Case 39 To 41
            Call CheckNpcAppearanceChangeCase(group)
        Case 46 To 67
            Call CheckRespawnCase(group)
        Case 42 To 45
            Call CheckDiagnosticsCase(group)
    End Select
End Sub

Public Sub TestNpcRespawnEffects()
    Call UnitTesting.RunTest("respawn queue retains death provenance and excludes initial spawn", TestNpcRespawnQueueProvenance())
    Dim group As Long
    For group = 46 To 67
        Call RunAdjacentCase(group, "respawn " & CStr(group))
    Next group
End Sub

Private Sub CheckRespawnCase(ByVal group As Long)
    Dim reader As Network.reader, extra As Integer, x As Integer, y As Integer
    Select Case group
        Case 58, 59: UserList(1).pos.y = 80
        Case 61
            Call AdjacentExpect(LoadAdjacentTopologyJson("{""version"":1,""id_space"":""resource_map_id"",""maps"":[34,35,36],""edges"":[[34,36,200,200],[36,34,-200,-200],[36,35,-200,-280],[35,36,200,280]]}", False), "two-hop non-cardinal topology")
        Case 62
            Call AdjacentExpect(LoadAdjacentTopologyJson("{""version"":1,""id_space"":""resource_map_id"",""maps"":[34,35,36],""edges"":[[34,35,0,-80],[35,34,0,80],[34,36,200,200],[36,34,-200,-200],[36,35,-199,-280],[35,36,199,280]]}", False), "ambiguous placement topology")
        Case 63: MapInfo(35).MapResource = 135
        Case 64: MapInfo(34).MapResource = 134
        Case 65: Call ClearAdjacentCapture(3)
        Case 66
            Call AddAdjacentSecondNpc()
            Call ClearAdjacentCapture(3)
        Case 67
            For extra = 1 To ADJACENT_NPC_LIMIT
                x = 26 + ((extra - 1) Mod 16): y = 73 + ((extra - 1) \ 16)
                NpcList(adjacentNpc - extra) = NpcList(adjacentNpc)
                NpcList(adjacentNpc - extra).pos.x = x: NpcList(adjacentNpc - extra).pos.y = y
                NpcList(adjacentNpc - extra).Char.charindex = 400 + extra
                MapData(x, y, 35).NpcIndex = adjacentNpc - extra
            Next extra
        Case 47: NpcList(adjacentNpc).flags.invisible = 1
        Case 48: UserList(1).flags.Muerto = 1
        Case 49: UserList(1).pos.y = 80
        Case 50: NpcList(adjacentNpc).MaestroUser.ArrayIndex = 1
        Case 51: MapData(50, 90, 35).NpcIndex = 0
        Case 52: Call ClearAdjacentCapture(2)
        Case 53
            UserList(1).pos.Map = 35: UserList(1).pos.y = 90
        Case 54
            UserList(1).flags.Muerto = 1: MapInfo(34).Seguro = 1
        Case 55
            UserList(1).flags.Muerto = 1: NpcList(adjacentNpc).npcType = e_NPCType.Revividor
        Case 56
            Call TickAdjacent()
            Call ClearAdjacentCapture()
            Call TickAdjacent()
            Call AdjacentExpect(adjacentPackets.Count = 0, "ordinary visibility never emits respawn")
            Exit Sub
        Case 57
            Call ClearAdjacentCapture(1)
    End Select
    ' Exercise the production map connection index, including same-map relocation.
    ConnGroups(UserList(1).pos.Map).CountEntrys = 1
    ConnGroups(UserList(1).pos.Map).UserEntrys(1) = 1
    Call NotifyNpcRespawn(adjacentNpc)
    If group = 53 Then
        Call AdjacentExpect(AdjacentNpcTestCollectionCount() = 0, "same-map appearance needs no adjacent collection")
    ElseIf group = 46 Or group = 52 Or group = 54 Or group = 55 Or group = 57 Or group = 60 Or group = 61 Or group >= 65 Then
        Call AdjacentExpect(AdjacentNpcTestCollectionCount() = 1, "one collection per eligible viewer")
        Call AdjacentExpect(AdjacentNpcTestScanCount(35) = 1, "source tiles scanned once")
    Else
        Call AdjacentExpect(AdjacentNpcTestCollectionCount() = 0 And AdjacentNpcTestScanCount(35) = 0, "no collection for irrelevant viewer")
    End If
    If group = 58 Then Call AdjacentExpect(AdjacentNpcTestRespawnViewerVisits() = 1, "64 unrelated connections are not visited")
    Select Case group
        Case 46, 54, 55, 60, 61
            Call AdjacentExpect(adjacentPackets.Count = 3, "reset upsert respawn ordering")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            Set reader = AdjacentPacketReader(3, ServerPacketID.eNpcRespawnEffect, 20)
            Call AdjacentExpect(reader.ReadInt32() > 0, "one-shot sequence")
            Call AdjacentExpect(reader.ReadInt16() = 34 And reader.ReadInt16() = 35, "viewer and source maps")
            Call AdjacentExpect(reader.ReadInt16() = adjacentNpc And reader.ReadInt16() = 3, "NPC lifetime")
            Call AdjacentExpect(reader.ReadInt16() = 123 And reader.ReadInt16() = 1543, "character and definition")
            Call AdjacentExpect(reader.ReadInt8() = 50 And reader.ReadInt8() = 90, "source tile")
        Case 65
            Call AdjacentExpect(adjacentPackets.Count = 3 And Not CBool(adjacentAccepted(3)), "effect enqueue failed after accepted appearance")
        Case 66
            Call AdjacentExpect(adjacentPackets.Count = 4 And Not CBool(adjacentAccepted(3)), "partial appearance prefix accepted")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            Set reader = AdjacentPacketReader(4, ServerPacketID.eNpcRespawnEffect, 20)
        Case 53
            Call AdjacentExpect(adjacentPackets.Count = 2, "active create then effect")
            Set reader = AdjacentPacketReader(2, ServerPacketID.eNpcRespawnEffect, 20)
        Case 52
            Call AdjacentExpect(adjacentPackets.Count = 2, "failed upsert suppresses effect")
        Case 57
            Call AdjacentExpect(adjacentPackets.Count = 1, "failed reset suppresses effect")
        Case Else
            Call AdjacentExpect(adjacentPackets.Count = 0, "ineligible viewer or placement suppressed")
    End Select
    ' Unrelated sessions have served their notification-cost assertion. End them
    ' before the global tick, whose normal resets are outside this one-viewer capture.
    If group = 58 Then
        For extra = 2 To 65
            UserList(extra).flags.UserLogged = False
        Next extra
    End If
    Call ClearAdjacentCapture()
    Call TickAdjacent()
    ' An eventual successful retry is appearance only, never the missed effect.
    Dim i As Long, bytes() As Byte
    For i = 1 To adjacentPackets.Count
        bytes = adjacentPackets(i)
        Call AdjacentExpect(CLng(bytes(0)) + 256& * bytes(1) <> ServerPacketID.eNpcRespawnEffect, "no delayed effect")
    Next i
End Sub

Private Sub ExpectAdjacentAppearance(ByVal packetIndex As Long, ByVal expectedText As String, ByVal expectedLength As Long)
    Dim reader As Network.reader
    Set reader = AdjacentPacketReader(packetIndex, ServerPacketID.eAdjacentNpcUpsert, expectedLength)
    With NpcList(adjacentNpc)
        Call AdjacentExpect(reader.ReadInt32() = AdjacentNpcTestEpoch(1), "appearance scene")
        Call AdjacentExpect(reader.ReadInt16() = adjacentNpc And reader.ReadInt16() = .VersionId, "appearance lifetime")
        Call AdjacentExpect(reader.ReadInt16() = .pos.Map, "appearance source map")
        Call AdjacentExpect(reader.ReadInt8() = .pos.x And reader.ReadInt8() = .pos.y, "appearance source tile")
        Call AdjacentExpect(reader.ReadInt16() = .Char.body And reader.ReadInt16() = .Char.head, "appearance body/head")
        Call AdjacentExpect(reader.ReadInt8() = .Char.Heading, "appearance heading")
        Call AdjacentExpect(reader.ReadInt16() = .Char.WeaponAnim And reader.ReadInt16() = .Char.ShieldAnim, "appearance weapon/shield")
        Call AdjacentExpect(reader.ReadInt16() = .Char.CascoAnim And reader.ReadInt16() = .Char.CartAnim, "appearance helmet/cart")
        Call AdjacentExpect(reader.ReadReal32() = .Char.speeding And reader.ReadBool() = .flags.NPCIdle, "appearance speed/idle")
        Call AdjacentExpect(reader.ReadInt16() = .Numero, "definition is not NPC slot, character index or lifetime")
        Call AdjacentExpect(reader.ReadString8() = expectedText, "authorized runtime display text")
    End With
    Call AdjacentExpect(reader.GetAvailable() = 0, "no unparsed adjacent appearance bytes")
End Sub

Private Sub ExpectNormalAppearance(ByVal expectedText As String)
    Dim bytes() As Byte, reader As New Network.reader, i As Long
    Call ClearAdjacentCapture()
    With NpcList(adjacentNpc)
        ' Exercise the same production creation/refresh used by ModAreas and taming.
        Call MakeNPCChar(False, 1, adjacentNpc, .pos.Map, .pos.x, .pos.y)
        Call AdjacentExpect(adjacentCaptureError = vbNullString And adjacentPackets.Count = 1, "one normal create, no unintended effect packet")
        bytes = adjacentPackets(1)
        Call reader.SetData(bytes)
        Call AdjacentExpect(reader.ReadInt16() = ServerPacketID.eCharacterCreate, "normal create packet")
        Call AdjacentExpect(reader.ReadInt16() = .Char.charindex, "normal character index")
        Call AdjacentExpect(reader.ReadInt16() = .Char.body And reader.ReadInt16() = .Char.head, "normal body/head")
        Call AdjacentExpect(reader.ReadInt8() = .Char.Heading, "normal heading")
        Call AdjacentExpect(reader.ReadInt8() = .pos.x And reader.ReadInt8() = .pos.y, "normal tile")
        Call AdjacentExpect(reader.ReadInt16() = .Char.WeaponAnim And reader.ReadInt16() = .Char.ShieldAnim, "normal weapon/shield")
        Call AdjacentExpect(reader.ReadInt16() = .Char.CascoAnim And reader.ReadInt16() = .Char.CartAnim, "normal helmet/cart argument mapping")
        Call AdjacentExpect(reader.ReadInt16() = 0, "NPC backpack remains absent")
        Call AdjacentExpect(reader.ReadInt16() = 0 And reader.ReadInt16() = 0, "equipment must not become FX or FXLoops")
        Call AdjacentExpect(reader.ReadString8() = expectedText, "normal authorized runtime text")
        Call reader.ReadInt8(): Call reader.ReadInt8()
        Call AdjacentExpect(reader.ReadInt8() = 0, "no character particle")
        For i = 1 To 7
            Call reader.ReadString8()
        Next i
        Call AdjacentExpect(reader.ReadReal32() = .Char.speeding And reader.ReadInt8() = 1, "normal speed/NPC kind")
        Call reader.ReadInt8(): Call reader.ReadInt16(): Call reader.ReadInt16(): Call reader.ReadInt8()
        For i = 1 To 4
            Call reader.ReadInt32()
        Next i
        Call reader.ReadInt8()
        Call AdjacentExpect(reader.ReadInt8() = Abs(CInt(.flags.NPCIdle)), "normal idle flag")
        Call reader.ReadInt8(): Call reader.ReadInt8(): Call reader.ReadInt8()
        Call AdjacentExpect(reader.ReadInt16() = .Numero And reader.GetAvailable() = 0, "normal definition and complete payload")
    End With
    Call ClearAdjacentCapture()
End Sub

Private Sub ExpectNpcAppearanceChange(ByVal expectedBody As Integer)
    Dim reader As Network.reader
    Call AdjacentExpect(adjacentCaptureError = vbNullString And adjacentPackets.Count = 1, "one alive-area character change, no extra FX")
    Set reader = AdjacentPacketReader(1, ServerPacketID.eCharacterChange, 24)
    With NpcList(adjacentNpc)
        Call AdjacentExpect(reader.ReadInt16() = .Char.charindex, "change keeps character index")
        Call AdjacentExpect(reader.ReadInt8() = Abs(CInt(.flags.NPCIdle)), "change idle flag and not navigating")
        Call AdjacentExpect(reader.ReadInt16() = expectedBody And reader.ReadInt16() = .Char.head, "change body argument and authoritative head")
        Call AdjacentExpect(reader.ReadInt8() = .Char.Heading, "change authoritative heading")
        Call AdjacentExpect(reader.ReadInt16() = .Char.WeaponAnim And reader.ReadInt16() = .Char.ShieldAnim, "change retains weapon/shield")
        Call AdjacentExpect(reader.ReadInt16() = .Char.CascoAnim And reader.ReadInt16() = .Char.CartAnim, "change retains helmet/cart")
        Call AdjacentExpect(reader.ReadInt16() = 0, "change has no NPC backpack")
        Call AdjacentExpect(reader.ReadInt16() = 0 And reader.ReadInt16() = 0, "change equipment does not become FX/loops")
    End With
    Call AdjacentExpect(reader.GetAvailable() = 0, "complete character change payload")
End Sub

Private Sub CheckNpcAppearanceChangeCase(ByVal group As Long)
    Dim iteration As Long, originalBody As Integer
    With NpcList(adjacentNpc)
        .Char.body = 257: .Char.head = 514: .Char.WeaponAnim = 771: .Char.ShieldAnim = 1028
        .Char.CascoAnim = 1285: .Char.CartAnim = 1542: .Char.Heading = e_Heading.SOUTH
        .flags.NPCIdle = False
        If group = 40 Then
            ' NPC269 in Recursos/Dat/npcs.dat: fixed-origin guard, already South.
            .Numero = 269: .npcType = e_NPCType.GuardiaNpc: .Movement = e_TipoAI.GuardiaPersigueNpc
            .Char.body = 578: .Char.head = 42: .Char.WeaponAnim = 68: .Char.ShieldAnim = 16
            .Char.CascoAnim = 0: .Char.CartAnim = 0: .Orig = .pos
        End If
        originalBody = .Char.body
        Call ExpectNormalAppearance(vbNullString)
        Call TickAdjacent()
        Call ExpectAdjacentAppearance(2, vbNullString, 35)
        For iteration = 1 To 2
            Call ClearAdjacentCapture(0, ToNPCAliveArea, adjacentNpc)
            Select Case group
                Case 39
                    Call AnimacionIdle(adjacentNpc, iteration = 1)
                Case 40
                    Call NpcAI(adjacentNpc)
                Case 41
                    If iteration = 2 Then
                        .Char.WeaponAnim = 0: .Char.ShieldAnim = 0: .Char.CascoAnim = 0: .Char.CartAnim = 0
                    End If
                    ' The caller's body override is sent, not stored into the NPC.
                    Call ChangeNPCChar(adjacentNpc, 0, 2056, e_Heading.SOUTH)
            End Select
            Call ExpectNpcAppearanceChange(IIf(group = 41, 0, originalBody))
            Call AdjacentExpect(.Char.body = originalBody, "change does not mutate stored body")
            Call AdjacentExpect(.Char.Heading = e_Heading.SOUTH And .pos.Map = 35 And .pos.x = 50 And .pos.y = 90, "same South heading and stationary tile")
            If group = 39 Then Call AdjacentExpect(.flags.NPCIdle = (iteration = 1), "idle changes both directions")
            If group = 40 Then Call AdjacentExpect(Not IsValidNpcRef(.TargetNPC), "guard remains at origin without a target")
            Call ClearAdjacentCapture()
            Call TickAdjacent()
            If group = 40 Then
                Call AdjacentExpect(adjacentPackets.Count = 0, "same-heading guard refresh leaves adjacent state unchanged")
            Else
                Call AdjacentExpect(adjacentPackets.Count = 1, "changed idle/head/equipment reaches adjacent view")
                Call ExpectAdjacentAppearance(1, vbNullString, 35)
            End If
        Next iteration
        ' A fresh adjacent scene must still agree after the active-map refreshes.
        Call InvalidateAdjacentNpcViewer(1)
        Call ClearAdjacentCapture()
        Call TickAdjacent()
        Call AdjacentExpect(adjacentPackets.Count = 2, "new scene reset then current appearance")
        Call ExpectAdjacentAppearance(2, vbNullString, 35)
    End With
End Sub

Private Sub CheckAdjacentAppearanceCase(ByVal group As Long)
    Dim expectedText As String, bytes() As Byte
    With NpcList(adjacentNpc)
        Select Case group
            Case 29
                .Char.body = 257: .Char.head = 514: .Char.WeaponAnim = 771: .Char.ShieldAnim = 1028
                .Char.CascoAnim = 1285: .Char.CartAnim = 1542: .Char.Heading = e_Heading.SOUTH
                .Char.speeding = 1.25: .flags.NPCIdle = True
                .showName = 1: .name = "Centinela Núñez": .SubName = "<Guardián de Ullathorpe>"
                expectedText = "Centinela Núñez<Guardián de Ullathorpe>"
                Call ExpectNormalAppearance(expectedText)
                Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 2, "normal create then adjacent snapshot")
                Call ExpectAdjacentAppearance(2, expectedText, 77)
                bytes = adjacentPackets(2)
                Call UnitTesting.Progress("adjacent appearance production upsert: " & AO20CryptoSysWrapper.ByteArrayToHex(bytes))
                Call AdjacentExpect(bytes(34) = 42 And bytes(46) = &HC3 And bytes(47) = &HBA, "UTF-8 accented name, not ANSI")
            Case 30 To 32
                .showName = 1: .name = "Guard": .SubName = "<South>"
                Call TickAdjacent(): Call ClearAdjacentCapture()
                Select Case group
                    Case 30: .Numero = 2049
                    Case 31: .name = "Watch"
                    Case 32: .SubName = "<North>"
                End Select
                expectedText = .name & .SubName
                Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 1, "stationary definition/name/description change sends upsert")
                Call ExpectAdjacentAppearance(1, expectedText, 35 + Len(expectedText))
                Call ClearAdjacentCapture(): Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 0, "accepted display change is silent next sample")
            Case 33
                .showName = 1: .name = "Guard": .SubName = "<South>"
                Call TickAdjacent(): Call ClearAdjacentCapture()
                .showName = 0
                Call ExpectNormalAppearance(vbNullString)
                Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 1, "hide label without removing NPC")
                Call ExpectAdjacentAppearance(1, vbNullString, 35)
                Call ClearAdjacentCapture()
                .name = "Watch": .SubName = "<North>"
                Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 0, "hidden text changes do not leak")
                .showName = 2
                Call ExpectNormalAppearance("Watch<North>")
                Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 1, "showName greater than zero reveals label without movement")
                Call ExpectAdjacentAppearance(1, "Watch<North>", 47)
            Case 34
                .showName = 1: .name = vbNullString: .SubName = vbNullString
                Call ExpectNormalAppearance(vbNullString)
                Call TickAdjacent()
                Call ExpectAdjacentAppearance(2, vbNullString, 35)
            Case 35
                .showName = 1: .name = String$(126, "x") & "ñ"
                Call ExpectNormalAppearance(.name)
                Call TickAdjacent()
                Call ExpectAdjacentAppearance(2, .name, 164)
                bytes = adjacentPackets(2)
                Call AdjacentExpect(bytes(34) = &H80 And bytes(35) = 1, "128 UTF-8 bytes use two prefix bytes")
                Call AdjacentExpect(bytes(162) = &HC3 And bytes(163) = &HB1, "accented ending roundtrips")
            Case 36
                .showName = 1: .name = String$(65496, "x") & "ñ"
                Call TickAdjacent()
                Call ExpectAdjacentAppearance(2, .name, 65535)
                bytes = adjacentPackets(2)
                Call AdjacentExpect(bytes(34) = &HDA And bytes(35) = &HFF And bytes(36) = 3, "65498 UTF-8 bytes use three prefix bytes")
                Call ClearAdjacentCapture(): Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 0, "maximum-length label is stable")
            Case 37
                .showName = 1: .name = String$(65499, "x")
                Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestKnownCount(1) = 0, "oversized upsert not sent or cached after reset")
                .name = "Guard"
                Call ClearAdjacentCapture(): Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestKnownCount(1) = 1, "corrected label retries without another reset")
                Call ExpectAdjacentAppearance(1, "Guard", 40)
            Case 38
                .showName = 1: .name = "Guard"
                Call TickAdjacent(): Call ClearAdjacentCapture(1)
                .SubName = "<New>"
                Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 1 And Not CBool(adjacentAccepted(1)), "display refresh enqueue fails")
                Call ExpectAdjacentAppearance(1, "Guard<New>", 45)
                Call ClearAdjacentCapture(): Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 1 And CBool(adjacentAccepted(1)), "unchanged failed display refresh retried")
                Call ExpectAdjacentAppearance(1, "Guard<New>", 45)
                Call ClearAdjacentCapture(): Call TickAdjacent()
                Call AdjacentExpect(adjacentPackets.Count = 0, "accepted display retry converges")
        End Select
    End With
End Sub

Private Sub RunAdjacentCase(ByVal group As Long, ByVal name As String)
    On Error GoTo Failed
    Dim oldUser As t_User, oldNpc As t_Npc, oldSecond As t_Npc, emptyUser As t_User, emptyNpc As t_Npc
    Dim oldMap34 As t_MapInfo, oldMap35 As t_MapInfo, emptyTile As t_MapBlock
    Dim tiles34(1 To 100, 1 To 100) As t_MapBlock, tiles35(1 To 100, 1 To 100) As t_MapBlock
    Dim oldLastUser As Integer, x As Integer, y As Integer, saved As Boolean, passed As Boolean, failure As String
    Dim oldConfig As ServerConfig, oldToggle As Boolean, oldLastNpc As Integer
    Dim oldNumChars As Integer, oldLastChar As Integer, oldEntries34 As Integer, oldEntries35 As Integer
    Dim oldGroup34 As t_ConnGroup, oldGroup35 As t_ConnGroup, oldGroup36 As t_ConnGroup, oldGroup1 As t_ConnGroup
    Dim oldResource36 As Integer, oldUserUpper As Long, extra As Integer
    Dim extraUsers(2 To 65) As t_User, extraNpcs(1 To ADJACENT_NPC_LIMIT) As t_Npc
    Call AdjacentExpect(NumMaps >= 35 And UBound(UserList) >= 1 And UBound(NpcList) >= 2, "initialized test world")
    adjacentNpc = UBound(NpcList) - 1
    If group >= 42 And group <= 45 Then adjacentNpc = 1
    Set oldConfig = SvrConfig
    oldToggle = IsFeatureEnabled(NPC_CROSS_MAP_PURSUIT_FEATURE): oldLastNpc = LastNPC
    oldNumChars = NumChars: oldLastChar = LastChar
    oldEntries34 = ConnGroups(34).CountEntrys: oldEntries35 = ConnGroups(35).CountEntrys
    oldUser = UserList(1): oldNpc = NpcList(adjacentNpc): oldSecond = NpcList(adjacentNpc + 1)
    oldLastUser = LastUser
    oldMap34 = MapInfo(34): oldMap35 = MapInfo(35)
    For y = 1 To 100
        For x = 1 To 100
            tiles34(x, y) = MapData(x, y, 34)
            tiles35(x, y) = MapData(x, y, 35)
        Next x
    Next y
    If group >= 46 Then
        oldGroup34 = ConnGroups(34): oldGroup35 = ConnGroups(35): oldGroup36 = ConnGroups(36)
        oldResource36 = MapInfo(36).MapResource
        If group = 58 Then
            oldGroup1 = ConnGroups(1): oldUserUpper = UBound(UserList)
            For extra = 2 To 65
                If extra <= oldUserUpper Then extraUsers(extra) = UserList(extra)
            Next extra
        End If
        If group = 67 Then
            For extra = 1 To ADJACENT_NPC_LIMIT
                extraNpcs(extra) = NpcList(adjacentNpc - extra)
            Next extra
        End If
    End If
    saved = True
    If group >= 46 Then
        ConnGroups(34).CountEntrys = 0: ConnGroups(35).CountEntrys = 0: ConnGroups(36).CountEntrys = 0
        ReDim ConnGroups(34).UserEntrys(1 To 1): ReDim ConnGroups(35).UserEntrys(1 To 1)
        MapInfo(36).MapResource = 36
        If group = 58 Then
            If oldUserUpper < 65 Then ReDim Preserve UserList(1 To 65)
            ReDim ConnGroups(1).UserEntrys(1 To 64)
            ConnGroups(1).CountEntrys = 64
        End If
    End If
    If group >= 42 And group <= 45 Then
        ConnGroups(34).CountEntrys = 0: ConnGroups(35).CountEntrys = 0
    End If
    Call AdjacentNpcTestResetState
    Call AdjacentExpect(LoadAdjacentTopologyJson(ADJACENT_TEST_TOPOLOGY, False), "fixture topology")
    For y = 1 To 100
        For x = 1 To 100
            MapData(x, y, 34) = emptyTile
            MapData(x, y, 35) = emptyTile
        Next x
    Next y
    MapInfo(34).MapResource = 34: MapInfo(35).MapResource = 35
    MapInfo(34).Seguro = 0: MapInfo(35).Seguro = 0
    UserList(1) = emptyUser
    LastUser = 1
    With UserList(1)
        .Id = 101: .VersionId = 7
        .flags.UserLogged = True
        .flags.Privilegios = e_PlayerType.User
        .ConnectionDetails.ConnIDValida = True
        .ConnectionDetails.ConnID = 42
        .pos.Map = 34: .pos.x = 50: .pos.y = 11
    End With
    If group = 58 Then
        LastUser = 65
        For extra = 2 To 65
            UserList(extra) = UserList(1): UserList(extra).pos.Map = 1
            ConnGroups(1).UserEntrys(extra - 1) = extra
        Next extra
    End If
    NpcList(adjacentNpc) = emptyNpc
    NpcList(adjacentNpc + 1) = emptyNpc
    With NpcList(adjacentNpc)
        .flags.NPCActive = True
        .VersionId = 3
        .Numero = 1543
        .Char.charindex = 123: .Char.body = 10: .Char.Heading = e_Heading.NORTH: .Char.speeding = 1
        .pos.Map = 35: .pos.x = 50: .pos.y = 90
        .npcType = e_NPCType.Comun
    End With
    MapData(50, 90, 35).NpcIndex = adjacentNpc
    Call InitializeAuxiliaryBuffer
    Call ClearAdjacentCapture
    adjacentTick = 1000
    adjacentCaptureActive = True
    Call CheckAdjacentCase(group)
    passed = True
    GoTo Cleanup
Failed:
    failure = CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error GoTo 0
    diagnosticsCapture = False
    adjacentCaptureActive = False
    Call AdjacentNpcTestResetState
    If saved Then
        If group >= 46 Then
            ConnGroups(34) = oldGroup34: ConnGroups(35) = oldGroup35: ConnGroups(36) = oldGroup36
            MapInfo(36).MapResource = oldResource36
            If group = 58 Then
                ConnGroups(1) = oldGroup1
                For extra = 2 To 65
                    If extra <= oldUserUpper Then UserList(extra) = extraUsers(extra)
                Next extra
                If oldUserUpper < 65 Then ReDim Preserve UserList(1 To oldUserUpper)
            End If
            If group = 67 Then
                For extra = 1 To ADJACENT_NPC_LIMIT
                    NpcList(adjacentNpc - extra) = extraNpcs(extra)
                Next extra
            End If
        End If
        If group >= 42 And group <= 45 Then
            If NpcList(adjacentNpc).Char.charindex > 0 Then CharList(NpcList(adjacentNpc).Char.charindex) = 0
            Set SvrConfig = oldConfig
            Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, oldToggle)
            LastNPC = oldLastNpc: NumChars = oldNumChars: LastChar = oldLastChar
            ConnGroups(34).CountEntrys = oldEntries34: ConnGroups(35).CountEntrys = oldEntries35
        End If
        UserList(1) = oldUser: NpcList(adjacentNpc) = oldNpc: NpcList(adjacentNpc + 1) = oldSecond
        LastUser = oldLastUser
        MapInfo(34) = oldMap34: MapInfo(35) = oldMap35
        For y = 1 To 100
            For x = 1 To 100
                MapData(x, y, 34) = tiles34(x, y)
                MapData(x, y, 35) = tiles35(x, y)
            Next x
        Next y
    End If
    If passed Then
        Call UnitTesting.RunTest("adjacent stream " & name, True)
    Else
        Call UnitTesting.Progress("adjacent stream " & name & " FAILED: " & failure)
        Call UnitTesting.RunTestError("adjacent stream " & name, failure)
    End If
End Sub

Private Function TwoHopJson(ByVal edges As String) As String
    TwoHopJson = "{""version"":1,""id_space"":""resource_map_id"",""maps"":[34,35,78,80],""edges"":[" & edges & "]}"
End Function

Private Function TwoHopDiamond(ByVal dx As Long, ByVal dy As Long, Optional ByVal conflict As Boolean = False) As String
    Dim otherDx As Long
    otherDx = dx
    If conflict Then otherDx = dx - 1
    TwoHopDiamond = TwoHopJson("[34,35," & dx & ",0],[35,34," & -dx & ",0]," & _
        "[34,78,0," & dy & "],[78,34,0," & -dy & "]," & _
        "[35,80,0," & dy & "],[80,35,0," & -dy & "]," & _
        "[78,80," & otherDx & ",0],[80,78," & -otherDx & ",0]")
End Function

Private Sub PlaceTwoHopNpc(ByVal mapId As Integer, ByVal x As Byte, ByVal y As Byte)
    With NpcList(adjacentNpc)
        If .pos.Map > 0 Then MapData(.pos.x, .pos.y, .pos.Map).NpcIndex = 0
        .pos.Map = mapId: .pos.x = x: .pos.y = y
    End With
    MapData(x, y, mapId).NpcIndex = adjacentNpc
End Sub

Private Sub ExpectTwoHopPlacement(ByVal sourceMap As Integer, ByVal targetMap As Integer, ByVal dx As Long, ByVal dy As Long, ByVal depth As Integer)
    Dim placements() As t_AdjacentNpcPlacement, count As Long, i As Long, found As Long
    Call AdjacentExpect(TryGetAdjacentVisibilityPlacements(sourceMap, placements, count), "placements available")
    For i = 1 To count
        Call AdjacentExpect(placements(i).MapId <> sourceMap, "active map excluded")
        If i > 1 Then
            Call AdjacentExpect(placements(i).Depth > placements(i - 1).Depth Or _
                (placements(i).Depth = placements(i - 1).Depth And placements(i).MapId > placements(i - 1).MapId), "stable depth/map order")
        End If
        If placements(i).MapId = targetMap Then
            found = found + 1
            Call AdjacentExpect(placements(i).Dx = dx And placements(i).Dy = dy And placements(i).Depth = depth, "exact composed transform")
        End If
    Next i
    Call AdjacentExpect(found = 1, "source included once")
End Sub

Private Sub ExpectTwoHopAbsent(ByVal sourceMap As Integer, ByVal targetMap As Integer)
    Dim placements() As t_AdjacentNpcPlacement, count As Long, i As Long
    Call AdjacentExpect(TryGetAdjacentVisibilityPlacements(sourceMap, placements, count), "placements available")
    For i = 1 To count
        Call AdjacentExpect(placements(i).MapId <> targetMap, "excluded source absent")
    Next i
End Sub

Private Sub CheckAdjacentTwoHopCase(ByVal group As Long)
    Dim dx As Long, dy As Long, sx As Long, sy As Long, epoch As Long
    Dim reader As Network.reader, topologyPath As String, bytes() As Byte
    Select Case group
        Case 1
            topologyPath = App.Path & "\Recursos\config\legacy_overworld_topology.json"
            Call AdjacentExpect(FileExist(topologyPath, vbNormal), "checked-in topology fixture exists")
            Call AdjacentExpect(LoadAdjacentTopologyJson(FileText(topologyPath), False), "checked-in topology loads")
            Call ExpectTwoHopPlacement(78, 34, 74, 0, 1)
            Call ExpectTwoHopPlacement(80, 34, 74, 80, 2)
            Call PlaceTwoHopNpc(34, 15, 10)
            UserList(1).pos.Map = 78: UserList(1).pos.x = 90: UserList(1).pos.y = 10
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2, "first map reset/upsert")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            epoch = AdjacentNpcTestEpoch(1)
            UserList(1).pos.Map = 80: UserList(1).pos.y = 90
            Call ClearAdjacentCapture: Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2 And AdjacentNpcTestEpoch(1) > epoch, "diagonal entry fresh reset/upsert")
            Set reader = AdjacentPacketReader(1, ServerPacketID.eAdjacentNpcReset, 8)
            Call reader.ReadInt32()
            Call AdjacentExpect(reader.ReadInt16() = 80, "new active map serialized")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            bytes = adjacentPackets(2)
            Call AdjacentExpect(bytes(10) = 34 And bytes(11) = 0 And bytes(12) = 15 And bytes(13) = 10, "wire keeps source map and tile")
            Call AdjacentExpect(AdjacentNpcTestScanCount(34) = 1 And AdjacentNpcTestKnownCount(1) = 1, "two routes one scan one actor")
            epoch = AdjacentNpcTestEpoch(1)
            UserList(1).pos.Map = 78: UserList(1).pos.y = 10
            Call ClearAdjacentCapture: Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2 And AdjacentNpcTestEpoch(1) > epoch, "return resets old diagonal view")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            Call AdjacentExpect(AdjacentNpcTestKnownCount(1) = 1 And NpcList(adjacentNpc).VersionId = 3, "same lifetime no ghosts")
            Call ClearAdjacentCapture: Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 0, "return snapshot idempotent")
        Case 2 To 5
            sx = 1: sy = 1
            If group = 3 Or group = 5 Then sx = -1
            If group = 4 Or group = 5 Then sy = -1
            dx = sx * 74: dy = sy * 80
            Call AdjacentExpect(LoadAdjacentTopologyJson(TwoHopDiamond(dx, dy), False), "corner topology")
            Call ExpectTwoHopPlacement(34, 80, dx, dy, 2)
            UserList(1).pos.x = 50 + sx * 40: UserList(1).pos.y = 50 + sy * 40
            Call PlaceTwoHopNpc(80, CByte(50 - sx * 10), CByte(50 - sy * 22))
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 2 And AdjacentNpcTestKnownCount(1) = 1, "inclusive corner boundary")
            Call ExpectAdjacentActor(2, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
            Call AdjacentExpect(AdjacentNpcTestScanCount(80) = 1, "diamond source scanned once")
            Call AdjacentExpect(AdjacentNpcTestScanCount(35) = 1 And AdjacentNpcTestScanCount(78) = 1, "empty intermediates scanned without blocking diagonal")
            Call PlaceTwoHopNpc(80, CByte(50 - sx * 10 + sx), CByte(50 - sy * 22))
            Call ClearAdjacentCapture: Call TickAdjacent
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc, 3, True)
            Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestKnownCount(1) = 0, "one past horizontal boundary excluded")
            Call PlaceTwoHopNpc(80, CByte(50 - sx * 10), CByte(50 - sy * 22 + sy))
            Call ClearAdjacentCapture: Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 0, "one past vertical boundary excluded")
            Call PlaceTwoHopNpc(80, CByte(50 - sx * 10), CByte(50 - sy * 22))
            Call ClearAdjacentCapture: Call TickAdjacent
            Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
        Case 6
            Call AdjacentExpect(LoadAdjacentTopologyJson(TwoHopJson("[34,35,10,0],[35,34,-10,0],[35,78,10,0],[78,35,-10,0],[78,80,10,0],[80,78,-10,0]"), False), "cycle chain")
            Call ExpectTwoHopPlacement(34, 35, 10, 0, 1)
            Call ExpectTwoHopPlacement(34, 78, 20, 0, 2)
            Call ExpectTwoHopAbsent(34, 34): Call ExpectTwoHopAbsent(34, 80)
            Call PlaceTwoHopNpc(80, 20, 50)
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestKnownCount(1) = 0, "third edge NPC excluded despite nearby projection")
            Call AdjacentExpect(AdjacentNpcTestScanCount(34) = 0 And AdjacentNpcTestScanCount(80) = 0, "active and depth-three maps not scanned")
            Call AdjacentExpect(AdjacentNpcTestScanCount(35) = 1 And AdjacentNpcTestScanCount(78) = 1, "reachable empty sources scanned once")
        Case 7
            Call AdjacentExpect(LoadAdjacentTopologyJson(TwoHopDiamond(74, 80, True), False), "contradictory reciprocal topology")
            Call ExpectTwoHopAbsent(34, 80)
            UserList(1).pos.x = 90: UserList(1).pos.y = 90
            Call PlaceTwoHopNpc(80, 15, 10)
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestKnownCount(1) = 0, "ambiguous source fails closed")
            Call AdjacentExpect(AdjacentNpcTestScanCount(80) = 0, "ambiguous source never scanned")
        Case 8, 9
            dx = 30000
            If group = 9 Then dx = -30000
            Call AdjacentExpect(LoadAdjacentTopologyJson(TwoHopJson("[34,35," & dx & ",0],[35,34," & -dx & ",0],[35,80," & dx & ",0],[80,35," & -dx & ",0]"), False), "wide offset topology")
            Call ExpectTwoHopPlacement(34, 80, dx * 2, 0, 2)
            Call PlaceTwoHopNpc(80, 50, 50)
            Call TickAdjacent
            Call AdjacentExpect(adjacentPackets.Count = 1, "far transforms do not overflow into view")
            Call AdjacentExpect(AdjacentNpcTestScanCount(35) = 0 And AdjacentNpcTestScanCount(80) = 0, "nonintersecting maps never scanned")
        Case 10, 11
            Call AdjacentExpect(LoadAdjacentTopologyJson(TwoHopDiamond(74, 80), False), "diagonal lifecycle topology")
            UserList(1).pos.x = 90: UserList(1).pos.y = 90
            Call PlaceTwoHopNpc(80, 15, 10)
            If group = 10 Then
                MapInfo(80).MapResource = 35
                Call TickAdjacent
                Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestScanCount(80) = 0, "source instance excluded")
                MapInfo(80).MapResource = 80
                NpcList(adjacentNpc).flags.invisible = 1
                Call ClearAdjacentCapture: Call TickAdjacent
                Call AdjacentExpect(adjacentPackets.Count = 0 And AdjacentNpcTestKnownCount(1) = 0, "invisible diagonal excluded")
                NpcList(adjacentNpc).flags.invisible = 0
                Call ClearAdjacentCapture: Call TickAdjacent
                Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
                NpcList(adjacentNpc).flags.invisible = 1
                Call ClearAdjacentCapture: Call TickAdjacent
                Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcRemove, adjacentNpc, 3, True)
                MapInfo(34).MapResource = 35
                NpcList(adjacentNpc).flags.invisible = 0
                Call ClearAdjacentCapture: Call TickAdjacent
                Call AdjacentExpect(adjacentPackets.Count = 0 And AdjacentNpcTestScanCount(80) = 0, "viewer instance excluded")
            Else
                Call ClearAdjacentCapture(2): Call TickAdjacent
                epoch = AdjacentNpcTestEpoch(1)
                Call AdjacentExpect(adjacentPackets.Count = 2 And AdjacentNpcTestKnownCount(1) = 0, "failed diagonal upsert not cached")
                Call ClearAdjacentCapture: Call TickAdjacent
                Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestEpoch(1) = epoch, "same epoch retries only missing upsert")
                Call ExpectAdjacentActor(1, ServerPacketID.eAdjacentNpcUpsert, adjacentNpc, 3, True)
                NpcList(adjacentNpc).flags.invisible = 1
                Call ClearAdjacentCapture(1): Call TickAdjacent
                Call AdjacentExpect(AdjacentNpcTestKnownCount(1) = 1, "failed diagonal remove retains known actor")
                UserList(1).pos.Map = 80
                Call ClearAdjacentCapture: Call TickAdjacent
                Call AdjacentExpect(adjacentPackets.Count = 1 And AdjacentNpcTestEpoch(1) > epoch And AdjacentNpcTestKnownCount(1) = 0, "map reset drops obsolete failed remove")
                Set reader = AdjacentPacketReader(1, ServerPacketID.eAdjacentNpcReset, 8)
                Call ClearAdjacentCapture: Call TickAdjacent
                Call AdjacentExpect(adjacentPackets.Count = 0, "obsolete retry never returns")
            End If
    End Select
End Sub

Private Sub RunAdjacentTwoHopCase(ByVal group As Long, ByVal name As String)
    On Error GoTo Failed
    Dim ids As Variant, i As Long, x As Integer, y As Integer, mapId As Integer
    Dim maps(0 To 3) As t_MapInfo, tiles() As t_MapBlock, emptyTile As t_MapBlock
    Dim oldUser As t_User, oldNpc As t_Npc, emptyUser As t_User, emptyNpc As t_Npc
    Dim oldLastUser As Integer, saved As Boolean, passed As Boolean, failure As String
    Call AdjacentExpect(NumMaps >= 80 And UBound(UserList) >= 1 And UBound(NpcList) >= 2, "initialized two-hop world")
    ids = Array(34, 35, 78, 80)
    ReDim tiles(1 To 100, 1 To 100, 0 To 3)
    adjacentNpc = UBound(NpcList) - 1
    oldUser = UserList(1): oldNpc = NpcList(adjacentNpc): oldLastUser = LastUser
    For i = 0 To 3
        mapId = CInt(ids(i)): maps(i) = MapInfo(mapId)
        For y = 1 To 100
            For x = 1 To 100
                tiles(x, y, i) = MapData(x, y, mapId)
            Next x
        Next y
    Next i
    saved = True
    Call AdjacentNpcTestResetState
    For i = 0 To 3
        mapId = CInt(ids(i))
        MapInfo(mapId).MapResource = mapId: MapInfo(mapId).Seguro = 0
        For y = 1 To 100
            For x = 1 To 100
                MapData(x, y, mapId) = emptyTile
            Next x
        Next y
    Next i
    UserList(1) = emptyUser: LastUser = 1
    With UserList(1)
        .Id = 101: .VersionId = 7: .flags.UserLogged = True: .flags.Privilegios = e_PlayerType.User
        .ConnectionDetails.ConnIDValida = True: .ConnectionDetails.ConnID = 42
        .pos.Map = 34: .pos.x = 50: .pos.y = 50
    End With
    NpcList(adjacentNpc) = emptyNpc
    With NpcList(adjacentNpc)
        .flags.NPCActive = True: .VersionId = 3: .npcType = e_NPCType.Comun
        .Numero = 1543
        .Char.charindex = 123: .Char.body = 10: .Char.Heading = e_Heading.NORTH: .Char.speeding = 1
    End With
    Call InitializeAuxiliaryBuffer: Call ClearAdjacentCapture
    adjacentTick = 1000: adjacentCaptureActive = True
    Call CheckAdjacentTwoHopCase(group)
    passed = True
    GoTo Cleanup
Failed:
    failure = CStr(Err.Number) & ": " & Err.Description
    Resume Cleanup
Cleanup:
    On Error GoTo 0
    adjacentCaptureActive = False
    Call AdjacentNpcTestResetState
    If saved Then
        UserList(1) = oldUser: NpcList(adjacentNpc) = oldNpc: LastUser = oldLastUser
        For i = 0 To 3
            mapId = CInt(ids(i)): MapInfo(mapId) = maps(i)
            For y = 1 To 100
                For x = 1 To 100
                    MapData(x, y, mapId) = tiles(x, y, i)
                Next x
            Next y
        Next i
    End If
    If passed Then
        Call UnitTesting.RunTest("adjacent two-hop " & name, True)
    Else
        Call UnitTesting.RunTestError("adjacent two-hop " & name, failure)
    End If
End Sub

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

Private Function test_respawn_schedule_any_time() As Boolean
    test_respawn_schedule_any_time = IsNpcRespawnScheduleAllowed(eNpcRespawnAnyTime, True) And _
                                     IsNpcRespawnScheduleAllowed(eNpcRespawnAnyTime, False)
End Function

Private Function test_respawn_schedule_day_only() As Boolean
    test_respawn_schedule_day_only = IsNpcRespawnScheduleAllowed(eNpcRespawnDayOnly, True) And _
                                     Not IsNpcRespawnScheduleAllowed(eNpcRespawnDayOnly, False)
End Function

Private Function test_respawn_schedule_night_only() As Boolean
    test_respawn_schedule_night_only = Not IsNpcRespawnScheduleAllowed(eNpcRespawnNightOnly, True) And _
                                       IsNpcRespawnScheduleAllowed(eNpcRespawnNightOnly, False)
End Function

Private Function test_respawn_schedule_boundaries() As Boolean
    test_respawn_schedule_boundaries = Not WorldTime_IsDaytimeAtMinute(359, 6, 20) And _
                                       WorldTime_IsDaytimeAtMinute(360, 6, 20) And _
                                       WorldTime_IsDaytimeAtMinute(1199, 6, 20) And _
                                       Not WorldTime_IsDaytimeAtMinute(1200, 6, 20)
End Function

Private Function test_respawn_schedule_invalid() As Boolean
    test_respawn_schedule_invalid = IsNpcRespawnScheduleAllowed(3, True) And _
                                    IsNpcRespawnScheduleAllowed(3, False)
End Function

Private Function test_scheduled_initial_spawn_requires_free_origin() As Boolean
    Dim spawnPosition     As t_WorldPos
    Dim originalNpcIndex  As Integer
    Dim originalUserIndex As Integer

    spawnPosition.Map = 34
    spawnPosition.x = 50
    spawnPosition.y = 50
    originalNpcIndex = MapData(spawnPosition.x, spawnPosition.y, spawnPosition.Map).NpcIndex
    originalUserIndex = MapData(spawnPosition.x, spawnPosition.y, spawnPosition.Map).UserIndex

    MapData(spawnPosition.x, spawnPosition.y, spawnPosition.Map).NpcIndex = 0
    MapData(spawnPosition.x, spawnPosition.y, spawnPosition.Map).UserIndex = 0
    If Not IsNpcSpawnPositionAvailable(spawnPosition) Then GoTo Cleanup

    MapData(spawnPosition.x, spawnPosition.y, spawnPosition.Map).NpcIndex = 1
    If IsNpcSpawnPositionAvailable(spawnPosition) Then GoTo Cleanup

    MapData(spawnPosition.x, spawnPosition.y, spawnPosition.Map).NpcIndex = 0
    MapData(spawnPosition.x, spawnPosition.y, spawnPosition.Map).UserIndex = 1
    If IsNpcSpawnPositionAvailable(spawnPosition) Then GoTo Cleanup

    test_scheduled_initial_spawn_requires_free_origin = True

Cleanup:
    MapData(spawnPosition.x, spawnPosition.y, spawnPosition.Map).NpcIndex = originalNpcIndex
    MapData(spawnPosition.x, spawnPosition.y, spawnPosition.Map).UserIndex = originalUserIndex
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
    If dx <> 0 Or dy <> -80 Or Not AdjacentTopologyContainsMap(557) Then Exit Function
    ' Keep the previously omitted client connections in the shared server data.
    Dim reviewedEdges As Variant, reviewedEdge As Variant, parts As Variant
    reviewedEdges = Array("40,41,0,80", "41,42,0,80", "42,43,0,80", "43,44,0,80", _
        "44,45,-74,0", "342,343,0,80", "520,529,0,80", "521,530,0,80", _
        "529,530,74,0", "529,535,-74,0", "529,536,0,80")
    For Each reviewedEdge In reviewedEdges
        parts = Split(CStr(reviewedEdge), ",")
        If Not TryGetAdjacentEdge(CInt(parts(0)), CInt(parts(1)), dx, dy) Then Exit Function
        If dx <> CInt(parts(2)) Or dy <> CInt(parts(3)) Then Exit Function
        If Not TryGetAdjacentEdge(CInt(parts(1)), CInt(parts(0)), dx, dy) Then Exit Function
        If dx <> -CInt(parts(2)) Or dy <> -CInt(parts(3)) Then Exit Function
    Next reviewedEdge
    ' Terrain packs and server visibility load separate manifest copies.
    ' Keep the repaired Veril seam present in the server's installed copy too.
    If Not AdjacentTopologyContainsMap(144) Or Not AdjacentTopologyContainsMap(145) Then Exit Function
    If Not TryGetAdjacentEdge(144, 145, dx, dy) Then Exit Function
    If dx <> 74 Or dy <> 0 Then Exit Function
    If Not TryGetAdjacentEdge(145, 144, dx, dy) Then Exit Function
    If dx <> -74 Or dy <> 0 Then Exit Function
    If Not TryGetAdjacentEdge(144, 143, dx, dy) Then Exit Function
    If dx <> 0 Or dy <> -80 Then Exit Function
    If Not TryGetAdjacentEdge(143, 144, dx, dy) Then Exit Function
    If dx <> 0 Or dy <> 80 Then Exit Function
    Dim placements() As t_AdjacentNpcPlacement, count As Long
    ' The reviewed 143-144 connection extends Veril's two-hop neighborhood.
    ' Assert the complete depth-ordered sets, including offsets and no duplicates.
    If Not TryGetAdjacentVisibilityPlacements(145, placements, count) Then Exit Function
    If count <> 2 Then Exit Function
    If placements(1).MapId <> 144 Or placements(1).dx <> -74 Or placements(1).dy <> 0 Then Exit Function
    If placements(1).Depth <> 1 Or placements(2).Depth <> 2 Then Exit Function
    If placements(2).MapId <> 143 Or placements(2).dx <> -74 Or placements(2).dy <> -80 Then Exit Function
    If Not TryGetAdjacentVisibilityPlacements(144, placements, count) Then Exit Function
    ' The reviewed 143-141 edge adds map 141 at depth two from map 144.
    If count <> 3 Then Exit Function
    If placements(1).MapId <> 143 Or placements(1).dx <> 0 Or placements(1).dy <> -80 Then Exit Function
    If placements(1).Depth <> 1 Or placements(2).Depth <> 1 Then Exit Function
    If placements(2).MapId <> 145 Or placements(2).dx <> 74 Or placements(2).dy <> 0 Then Exit Function
    If placements(3).MapId <> 141 Or placements(3).Depth <> 2 Then Exit Function
    test_checked_in_manifest = (placements(3).dx = -74 And placements(3).dy = -80)
End Function

Private Function test_feature_toggle() As Boolean
    Dim originalValue As Boolean
    originalValue = IsFeatureEnabled(NPC_CROSS_MAP_PURSUIT_FEATURE)
    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, True)
    If Not AdjacentTopologyAvailable() Then GoTo RestoreToggle
    Call SetFeatureToggle(NPC_CROSS_MAP_PURSUIT_FEATURE, False)
    Dim destination As t_WorldPos
    destination.Map = 35
    destination.x = 50
    destination.y = 90
    test_feature_toggle = AdjacentTopologyAvailable() And _
        Not IsVerifiedNpcSpatialTransition(34, 50, 10, destination, False)
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
