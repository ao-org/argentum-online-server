' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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

Option Explicit
Const configPath = "../dat/Scenarios/RoomRush.ini"
Implements IBaseScenario

Const LANE_COUNT = 5
Const ROOM_COUNT = 5   ' Salas con NPC-guardián: 4 en el mapa de habitaciones + 1 en el mapa de espera/final
Const TIER_COUNT = 4
Const VICTORY_ROOM = 6 ' Sala 6: solo el teleport de salida, sin NPC

Private Type t_RoomConfig
    Map As Integer
    DoorX As Byte
    DoorY As Byte
    NpcSpawnX As Byte
    NpcSpawnY As Byte
End Type

Private Type t_LaneConfig
    EntranceMap As Integer
    EntranceX As Byte
    EntranceY As Byte
    Rooms(1 To ROOM_COUNT) As t_RoomConfig
    VictoryMap As Integer
    VictoryTopLeftX As Byte
    VictoryTopLeftY As Byte
    VictoryBottomRightX As Byte
    VictoryBottomRightY As Byte
End Type

Private Type t_LevelTier
    MinLevel As Byte
    MaxLevel As Byte
End Type

Dim Lanes(1 To LANE_COUNT)     As t_LaneConfig
Dim Tiers(1 To TIER_COUNT)     As t_LevelTier
Dim SelectedTier               As Integer
Dim MapRooms                   As Integer  ' 164
Dim MapWaiting                 As Integer  ' 165
Dim WaitingX                   As Byte
Dim WaitingY                   As Byte
Dim LaneGuardianNpc(1 To LANE_COUNT, 1 To ROOM_COUNT) As t_NpcReference
Dim LaneRoomCleared(1 To LANE_COUNT, 1 To ROOM_COUNT) As Boolean
Dim RoomRushKeyClaveMin         As Integer  ' rango de "clave" reservado para llaves de RoomRush, para poder limpiarlas
Dim RoomRushKeyClaveMax         As Integer
Dim LobbyIndex                  As Integer
Dim Winner                      As Integer  ' carril ganador, -1 si ninguno todavía
Dim UpdateIndex                 As Integer
Dim LastVictoryCheckTick        As Long
Dim LaneColorName(1 To LANE_COUNT) As String
Dim SharedRoomNpcId(1 To ROOM_COUNT, 1 To TIER_COUNT) As Integer
Dim SharedRoomDoorLockedObjIndex(1 To ROOM_COUNT) As Integer
Const VictoryCheckIntervalMs = 500

Friend Sub Class_Initialize()
    On Error GoTo Class_Initialize_Err
    Debug.Assert FileExist(DatPath & configPath)
    Dim ScenarioIni As clsIniManager
    Dim t As Integer
    Dim r As Integer
    Dim L As Integer
    Set ScenarioIni = New clsIniManager
    Call ScenarioIni.Initialize(DatPath & configPath)
    MapRooms = val(ScenarioIni.GetValue("INIT", "MapRooms"))
    MapWaiting = val(ScenarioIni.GetValue("INIT", "MapWaiting"))
    WaitingX = val(ScenarioIni.GetValue("INIT", "WaitingX"))
    WaitingY = val(ScenarioIni.GetValue("INIT", "WaitingY"))
    RoomRushKeyClaveMin = val(ScenarioIni.GetValue("INIT", "KeyClaveMin"))
    RoomRushKeyClaveMax = val(ScenarioIni.GetValue("INIT", "KeyClaveMax"))

    For t = 1 To TIER_COUNT
        Tiers(t).MinLevel = val(ScenarioIni.GetValue("LevelTiers", "Tier" & t & "Min"))
        Tiers(t).MaxLevel = val(ScenarioIni.GetValue("LevelTiers", "Tier" & t & "Max"))
    Next t

    For r = 1 To ROOM_COUNT
        SharedRoomDoorLockedObjIndex(r) = val(ScenarioIni.GetValue("Rooms", "Room" & r & "DoorLockedObjIndex"))
        For t = 1 To TIER_COUNT
            SharedRoomNpcId(r, t) = val(ScenarioIni.GetValue("Rooms", "Room" & r & "NpcId_Tier" & t))
        Next t
    Next r

    For L = 1 To LANE_COUNT
        With Lanes(L)
            .EntranceMap = MapRooms
            .EntranceX = val(ScenarioIni.GetValue("Lane" & L, "EntranceX"))
            .EntranceY = val(ScenarioIni.GetValue("Lane" & L, "EntranceY"))
            .VictoryMap = MapWaiting
            .VictoryTopLeftX = val(ScenarioIni.GetValue("Lane" & L, "VictoryTopLeftX"))
            .VictoryTopLeftY = val(ScenarioIni.GetValue("Lane" & L, "VictoryTopLeftY"))
            .VictoryBottomRightX = val(ScenarioIni.GetValue("Lane" & L, "VictoryBottomRightX"))
            .VictoryBottomRightY = val(ScenarioIni.GetValue("Lane" & L, "VictoryBottomRightY"))
            For r = 1 To ROOM_COUNT
                .Rooms(r).Map = IIf(r <= 4, MapRooms, MapWaiting)
                .Rooms(r).DoorX = val(ScenarioIni.GetValue("Lane" & L, "Room" & r & "DoorX"))
                .Rooms(r).DoorY = val(ScenarioIni.GetValue("Lane" & L, "Room" & r & "DoorY"))
                .Rooms(r).NpcSpawnX = val(ScenarioIni.GetValue("Lane" & L, "Room" & r & "NpcSpawnX"))
                .Rooms(r).NpcSpawnY = val(ScenarioIni.GetValue("Lane" & L, "Room" & r & "NpcSpawnY"))
            Next r
        End With
    Next L

    UpdateIndex = AddUpdateScenario(Me)
    Call CustomScenarios.SetCustomScenario(MapRooms, Me)
    Call CustomScenarios.SetCustomScenario(MapWaiting, Me)
    Winner = -1
    LaneColorName(1) = "Verde"
    LaneColorName(2) = "Rojo"
    LaneColorName(3) = "Celeste"
    LaneColorName(4) = "Azul"
    LaneColorName(5) = "Amarillo"
    Exit Sub
Class_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioRoomRush.Class_Initialize", Erl)
End Sub

Private Function SelectTierForLevel(ByVal MinLevel As Byte) As Integer
    Dim t As Integer
    For t = 1 To TIER_COUNT
        If MinLevel >= Tiers(t).MinLevel And (Tiers(t).MaxLevel = 0 Or MinLevel <= Tiers(t).MaxLevel) Then
            SelectTierForLevel = t
            Exit Function
        End If
    Next t
    SelectTierForLevel = 1 ' fallback si no matchea ningún tier
End Function

Private Sub ResetDoors()
    Dim L As Integer, r As Integer
    For L = 1 To LANE_COUNT
        For r = 1 To ROOM_COUNT
            With Lanes(L).Rooms(r)
                Debug.Assert LegalPos(.Map, .DoorX, .DoorY)
                MapData(.Map, .DoorX, .DoorY).ObjInfo.ObjIndex = SharedRoomDoorLockedObjIndex(r)
            End With
        Next r
    Next L
End Sub

Private Sub SpawnAllGuardians()
    Dim L As Integer, r As Integer
    Dim SpawnPos As t_WorldPos
    Dim NpcId As Integer
    For L = 1 To LANE_COUNT
        For r = 1 To ROOM_COUNT
            With Lanes(L).Rooms(r)
                SpawnPos.Map = .Map
                SpawnPos.x = .NpcSpawnX
                SpawnPos.y = .NpcSpawnY
                NpcId = CrearNPC(SharedRoomNpcId(r, SelectedTier), .Map, SpawnPos)
                If NpcId > 0 Then
                    Call SetNpcRef(LaneGuardianNpc(L, r), NpcId)
                    NpcList(NpcId).flags.Respawn = 0
                Else
                    Call TraceError(0, "RoomRush: fallo al spawnear guardian carril " & L & " sala " & r, "ScenarioRoomRush.SpawnAllGuardians", 0)
                End If
                LaneRoomCleared(L, r) = False
            End With
        Next r
    Next L
End Sub

Private Sub RemoveRoomRushKeys(ByVal UserIndex As Integer)
    Dim i As Integer
    Dim thisObjIndex As Integer
    With UserList(UserIndex)
        For i = 1 To UBound(.invent.Object)
            thisObjIndex = .invent.Object(i).ObjIndex
            If thisObjIndex > 0 Then
                If thisObjIndex >= LBound(ObjData) And thisObjIndex <= UBound(ObjData) Then
                    If ObjData(thisObjIndex).OBJType = e_OBJType.otKeys Then
                        If ObjData(thisObjIndex).clave >= RoomRushKeyClaveMin And ObjData(thisObjIndex).clave <= RoomRushKeyClaveMax Then
                            Call QuitarUserInvItem(UserIndex, i, .invent.Object(i).amount)
                        End If
                    End If
                End If
            End If
        Next i
        Call UpdateUserInv(True, UserIndex, 0)
    End With
End Sub

Public Sub IBaseScenario_SetLobbyIndex(ByVal value As Integer)
    LobbyIndex = value
    SelectedTier = SelectTierForLevel(LobbyList(LobbyIndex).MinLevel)
    Call ResetDoors
    Call ConfigMapLobby
End Sub

Private Sub ConfigMapLobby()
    Call ModLobby.SetTeamCount(LobbyList(LobbyIndex), LANE_COUNT, LobbyList(LobbyIndex).TeamType)
    Call SetSummonCoordinates(LobbyList(LobbyIndex), MapWaiting, WaitingX, WaitingY)
    MapInfo(MapRooms).Seguro = 0
    MapInfo(MapRooms).SinInviOcul = True
    MapInfo(MapRooms).SafeFightMap = True
    MapInfo(MapWaiting).SinInviOcul = True
End Sub

Public Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef NewState As e_LobbyState)
    If oldState = NewState Then Exit Sub
    If NewState = InProgress Then
        Call ResetDoors
        Call SpawnAllGuardians
        Winner = -1
        Call SendData(SendTarget.toMap, MapRooms, PrepareMessageLocaleMsg(MsgRoomRushStarted, "", e_FontTypeNames.FONTTYPE_GUILD))
        Call ModLobby.SummonAll(LobbyList(LobbyIndex))
    End If
    If NewState = Closed Then
        Call CleanupMatchState
        Call CleanupScenario
    End If
End Sub

Public Function IBaseScenario_GetScenarioName() As String
    IBaseScenario_GetScenarioName = "Room Rush"
End Function

Public Sub IBaseScenario_BroadcastOpenScenario()
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgRoomRushIntro, "", e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Public Sub IBaseScenario_SendRules(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MsgRoomRushInstructions, "", e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Public Function IBaseScenario_GetExtraInterfaces() As Long
    IBaseScenario_GetExtraInterfaces = 0
End Function

Public Sub IBaseScenario_Reset()
End Sub

Public Function IBaseScenario_RequiresSpawn() As Boolean
    IBaseScenario_RequiresSpawn = False
End Function

Public Function IBaseScenario_ValidateUser(ByVal UserIndex As Integer) As Integer
    IBaseScenario_ValidateUser = 0
End Function

Public Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    Dim SpawnPos As t_WorldPos
    Dim team As Integer
    team = UserList(UserIndex).flags.CurrentTeam
    If team >= 1 And team <= LANE_COUNT Then
        SpawnPos.Map = Lanes(team).EntranceMap
        SpawnPos.x = Lanes(team).EntranceX
        SpawnPos.y = Lanes(team).EntranceY
    Else
        SpawnPos.Map = MapWaiting
        SpawnPos.x = WaitingX
        SpawnPos.y = WaitingY
    End If
    LobbyList(LobbyIndex).SummonCoordinates = SpawnPos
End Sub

Public Sub IBaseScenario_Update()
    If Winner <> -1 Then Exit Sub
    If LobbyList(LobbyIndex).State <> e_LobbyState.InProgress Then Exit Sub
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If TicksElapsed(LastVictoryCheckTick, nowRaw) < VictoryCheckIntervalMs Then Exit Sub
    LastVictoryCheckTick = nowRaw
    Dim i As Integer, L As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            Dim UserIndex As Integer
            UserIndex = LobbyList(LobbyIndex).Players(i).User.ArrayIndex
            L = LobbyList(LobbyIndex).Players(i).team
            If L >= 1 And L <= LANE_COUNT Then
                If UserList(UserIndex).pos.Map = Lanes(L).VictoryMap Then
                    If UserList(UserIndex).pos.x >= Lanes(L).VictoryTopLeftX And UserList(UserIndex).pos.x <= Lanes(L).VictoryBottomRightX _
                            And UserList(UserIndex).pos.y >= Lanes(L).VictoryTopLeftY And UserList(UserIndex).pos.y <= Lanes(L).VictoryBottomRightY Then
                        Winner = L
                        Call MatchCompleted
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next i
End Sub

Public Sub IBaseScenario_PlayerKillNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    Dim L As Integer, r As Integer
    For L = 1 To LANE_COUNT
        For r = 1 To ROOM_COUNT
            If IsValidNpcRef(LaneGuardianNpc(L, r)) Then
                If LaneGuardianNpc(L, r).ArrayIndex = NpcIndex Then
                    LaneRoomCleared(L, r) = True
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgRoomRushRoomCleared, GetLaneDisplayName(L) & "?" & r, e_FontTypeNames.FONTTYPE_GUILD))
                    Exit Sub
                End If
            End If
        Next r
    Next L
End Sub

Public Sub IBaseScenario_NPcKillPlayer(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Public Sub IBaseScenario_PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Public Sub IBaseScenario_NpcDie(ByVal NpcIndex)
End Sub

Public Sub IBaseScenario_UserDie(ByVal UserIndex)
    ' ASUNCIÓN: morir no penaliza ni resetea progreso; el jugador revive donde el sistema normal indique.
End Sub

Private Sub CleanupMatchState()
    Dim i As Integer, L As Integer, r As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            Call RemoveRoomRushKeys(LobbyList(LobbyIndex).Players(i).User.ArrayIndex)
        End If
    Next i
    For L = 1 To LANE_COUNT
        For r = 1 To ROOM_COUNT
            If IsValidNpcRef(LaneGuardianNpc(L, r)) Then
                Call QuitarNPC(LaneGuardianNpc(L, r).ArrayIndex, eClearHunt)
            End If
        Next r
    Next L
    Call ResetDoors
End Sub

Private Sub MatchCompleted()
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgRoomRushWinnerTeam, GetLaneDisplayName(Winner), e_FontTypeNames.FONTTYPE_GUILD))
    Call GiveRewards
    Dim i As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            If LobbyList(LobbyIndex).Players(i).team <> Winner Then
                Call WriteLocaleMsg(LobbyList(LobbyIndex).Players(i).User.ArrayIndex, MsgRoomRushOtherTeamWon, e_FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    Next i
    Call ModLobby.ReturnAllPlayers(LobbyList(LobbyIndex))
    Call ModLobby.UpdateLobbyState(LobbyList(LobbyIndex), e_LobbyState.Closed)
End Sub

Private Sub CleanupScenario()
    Call CustomScenarios.RemoveCustomScenario(MapRooms)
    Call CustomScenarios.RemoveCustomScenario(MapWaiting)
    Call CustomScenarios.RemoveUpdateScenario(UpdateIndex)
    Call ReleaseLobby(LobbyIndex)
End Sub

Private Sub GiveRewards()
    If LobbyList(LobbyIndex).AvailableInscriptionMoney <= 0 Then Exit Sub
    Dim WinnerCount As Integer
    Dim i As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            If LobbyList(LobbyIndex).Players(i).team = Winner Then WinnerCount = WinnerCount + 1
        End If
    Next i
    If WinnerCount = 0 Then Exit Sub
    Dim Reward As Long
    Reward = LobbyList(LobbyIndex).AvailableInscriptionMoney / WinnerCount
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            If LobbyList(LobbyIndex).Players(i).team = Winner Then
                Call GiveGoldToPlayer(LobbyList(LobbyIndex), i, Reward)
            End If
        End If
    Next i
End Sub

Public Sub IBaseScenario_OnUserDisconnected(ByVal UserIndex As Integer)
    Call RemoveRoomRushKeys(UserIndex)
End Sub

Public Sub IBaseScenario_OnUserReconnect(ByVal UserIndex As Integer)
End Sub

Public Sub IBaseScenario_GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
    ' Los guardianes son estáticos, no patrullan.
End Sub

Private Function GetLaneDisplayName(ByVal L As Integer) As String
    Dim i As Integer
    Dim Names As String
    Dim playerCount As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If LobbyList(LobbyIndex).Players(i).team = L Then
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                If LenB(Names) > 0 Then Names = Names & ", "
                Names = Names & UserList(LobbyList(LobbyIndex).Players(i).User.ArrayIndex).name
                playerCount = playerCount + 1
            End If
        End If
    Next i
    If playerCount = 1 Then
        GetLaneDisplayName = Names & " (" & LaneColorName(L) & ")"
    Else
        GetLaneDisplayName = "Equipo " & LaneColorName(L) & " (" & Names & ")"
    End If
End Function
