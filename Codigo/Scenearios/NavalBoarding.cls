VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavalBoarding"
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

Const configPath = "../dat/Scenarios/NavalBattle.ini"

Implements IBaseScenario
Implements IInventoryInterface

Const TopMapNumber = 0
Const CargoMapNumber = 1
Const CannonSpell = 256

Private Type t_Waypoint
    PosX As Byte
    PosY As Byte
    NextWpCount As Byte
    NextWp() As Byte
End Type

Private Type t_TilePos
    PosX As Byte
    PosY As Byte
End Type

Private Type t_NpcSpawnInfo
    NpcId As Integer
    Waypoint As Integer
End Type

Private Type t_CargoElement
    WorldPos As t_WorldPos 'keeps track where it is in the work
    CarrierIndex As Integer 'keeps track of the user that carry this item
    id As Integer
    ItemIndex As Integer
End Type

Private Type t_CargoArea
    TopLeft As t_TilePos
    BottomRight As t_TilePos
End Type

Private Type t_Cargo
    CargoItemId As Integer
    OriginalCargoArea As t_CargoArea
    DropArea As t_CargoArea
    ItemList() As t_CargoElement
End Type

Private Type t_CaptinState
    LastShotTime As Long
    NpcRef As t_NpcReference
    TimeSinceLastFire As Long
    NextFireTime As Long
    TargetLocations(3) As t_WorldPos
End Type

Dim MapNumbers(1) As Integer
Dim TileBlocks() As t_TilePos

Dim MaxTime As Long
Dim CountdownTimer As t_Timer
Dim BroadCastTimer As t_Timer
Dim ScoreBroadcastTimer As t_Timer
Dim UpdateIndex As Integer
Dim LastFrameTime As Long
Dim ElapsedTime As Long
Dim Team1Waypoints() As t_Waypoint
Dim Team2Waypoints() As t_Waypoint
Dim PlayerRespawn As New Dictionary
Dim ForbbidenItemId() As Integer
Dim CaptainState(1 To 2) As t_CaptinState
Dim MinFireTime As Long
Dim MaxFireTime As Long
Dim FireStartTime As Long
Dim Team1SpawnX As Byte
Dim Team2SpawnX As Byte
Dim NpcSpawnList() As t_NpcSpawnInfo
Dim RespawnWp As Byte
Dim PlayerMinRespawnTime As Long
Dim RespawnInterval As Long
Dim ElapsedResapwnTime As Long
Dim CaptianNpcId As Integer
Dim UserLoot As New Dictionary
Dim TeamCargo(1 To 2) As t_Cargo

Private Sub Class_Initialize()
    On Error GoTo NavalBoarding_Initialize_Err:
100     Debug.Assert FileExist(DatPath & configPath)
        Dim ScenarioIni As clsIniManager
        Dim WpCount As Byte
102     Set ScenarioIni = New clsIniManager
104     Call ScenarioIni.Initialize(DatPath & configPath)
106     MapNumbers(0) = val(ScenarioIni.GetValue("INIT", "MapNumber1"))
108     MapNumbers(1) = val(ScenarioIni.GetValue("INIT", "MapNumber2"))
110     MaxTime = val(ScenarioIni.GetValue("INIT", "MaxTime"))
112     Team1SpawnX = val(ScenarioIni.GetValue("INIT", "Team1SpawnX"))
114     Team2SpawnX = val(ScenarioIni.GetValue("INIT", "Team2SpawnX"))
116     RespawnWp = val(ScenarioIni.GetValue("RESPAWN", "Waypoint"))
118     PlayerMinRespawnTime = val(ScenarioIni.GetValue("RESPAWN", "PlayerMinRespawnTime"))
120     RespawnInterval = val(ScenarioIni.GetValue("RESPAWN", "RespawnInterval"))
122     MaxTime = MaxTime * 1000
124     RespawnInterval = RespawnInterval * 1000
126     PlayerMinRespawnTime = PlayerMinRespawnTime * 1000
128     WpCount = val(ScenarioIni.GetValue("INIT", "WaypointCount"))
130     CaptianNpcId = val(ScenarioIni.GetValue("CaptainNpc", "NpcId"))
132     MinFireTime = val(ScenarioIni.GetValue("INIT", "MinFireTime"))
134     MaxFireTime = val(ScenarioIni.GetValue("INIT", "MaxFireTime"))
136     FireStartTime = val(ScenarioIni.GetValue("INIT", "FireStartTime"))
138     Call LoadTeamWeapoints(1, Team1Waypoints, WpCount, ScenarioIni)
140     Call LoadTeamWeapoints(2, Team2Waypoints, WpCount, ScenarioIni)
142     Call LoadForbidenItems(ScenarioIni)
144     Call LoadTileBlock(TileBlocks, ScenarioIni)
146     UpdateIndex = AddUpdateScenario(Me)
148     Call CustomScenarios.SetCustomScenario(MapNumbers(0), Me)
150     Call CustomScenarios.SetCustomScenario(MapNumbers(1), Me)
152     Call ConfigMapLobby
154     Call LoadNpcSpawnInfo(NpcSpawnList, ScenarioIni)
156     Call LoadTeamCargo(TeamCargo(1), 1, ScenarioIni)
158     Call LoadTeamCargo(TeamCargo(2), 2, ScenarioIni)
        Exit Sub
NavalBoarding_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "DeathMatch.Initialize ", Erl)
End Sub

Private Sub LoadTileBlock(ByRef Tiles() As t_TilePos, ScenearioIni As clsIniManager)
    Dim TileCount As Integer
    TileCount = val(ScenearioIni.GetValue("LOBBY-BLOCK", "Count"))
    ReDim Tiles(1 To TileCount) As t_TilePos
    Dim i As Integer
    For i = 1 To TileCount
        Tiles(i).PosX = val(ScenearioIni.GetValue("LOBBY-BLOCK", "TileX" & i))
        Tiles(i).PosY = val(ScenearioIni.GetValue("LOBBY-BLOCK", "TileY" & i))
        Call BlockAndInform(MapNumbers(1), Tiles(i).PosX, Tiles(i).PosY, 1)
    Next i
End Sub

Private Sub LoadTeamCargo(ByRef TeamCargo As t_Cargo, ByVal team As Integer, ScenearioIni As clsIniManager)
    TeamCargo.CargoItemId = val(ScenearioIni.GetValue("Team" & team & "Cargo", "CargoItemId"))
    TeamCargo.OriginalCargoArea.TopLeft.PosX = val(ScenearioIni.GetValue("Team" & team & "Cargo", "TopLeftX"))
    TeamCargo.OriginalCargoArea.TopLeft.PosY = val(ScenearioIni.GetValue("Team" & team & "Cargo", "TopLeftY"))
    TeamCargo.OriginalCargoArea.BottomRight.PosX = val(ScenearioIni.GetValue("Team" & team & "Cargo", "BottomLeftX"))
    TeamCargo.OriginalCargoArea.BottomRight.PosY = val(ScenearioIni.GetValue("Team" & team & "Cargo", "BottomLeftY"))
    TeamCargo.DropArea.TopLeft.PosX = val(ScenearioIni.GetValue("Team" & team & "Storage", "TopLeftX"))
    TeamCargo.DropArea.TopLeft.PosY = val(ScenearioIni.GetValue("Team" & team & "Storage", "TopLeftY"))
    TeamCargo.DropArea.BottomRight.PosX = val(ScenearioIni.GetValue("Team" & team & "Storage", "BottomLeftX"))
    TeamCargo.DropArea.BottomRight.PosY = val(ScenearioIni.GetValue("Team" & team & "Storage", "BottomLeftY"))
    Dim CargoCount As Integer
    Dim Width, Height As Integer
    Width = TeamCargo.OriginalCargoArea.BottomRight.PosX - TeamCargo.OriginalCargoArea.TopLeft.PosX + 1
    Height = TeamCargo.OriginalCargoArea.BottomRight.PosY - TeamCargo.OriginalCargoArea.TopLeft.PosY + 1
    CargoCount = Width * Height
    ReDim TeamCargo.ItemList(1 To CargoCount) As t_CargoElement
    Dim i As Integer
    For i = 1 To CargoCount
        TeamCargo.ItemList(i).ItemIndex = TeamCargo.CargoItemId
        TeamCargo.ItemList(i).CarrierIndex = 0
        TeamCargo.ItemList(i).id = team * 1000 + i
        TeamCargo.ItemList(i).WorldPos.x = TeamCargo.OriginalCargoArea.TopLeft.PosX + ((i - 1) Mod Width)
        TeamCargo.ItemList(i).WorldPos.y = TeamCargo.OriginalCargoArea.TopLeft.PosY + ((i - 1) \ Width)
        TeamCargo.ItemList(i).WorldPos.Map = MapNumbers(1)
    Next i
End Sub

Private Sub LoadNpcSpawnInfo(ByRef NpcSpawnList() As t_NpcSpawnInfo, ScenarioIni As clsIniManager)
    Dim StartIndex, TankCount, SupportCount, RandgedCount, CaptainCount As Integer
    TankCount = val(ScenarioIni.GetValue("TankNpc", "Count"))
    SupportCount = val(ScenarioIni.GetValue("SupportNpc", "Count"))
    RandgedCount = val(ScenarioIni.GetValue("RangedNpc", "Count"))
    CaptainCount = val(ScenarioIni.GetValue("CaptainNpc", "Count"))
    ReDim NpcSpawnList(1 To TankCount + SupportCount + RandgedCount + CaptainCount) As t_NpcSpawnInfo
    StartIndex = LoadNpcData(NpcSpawnList, ScenarioIni, "TankNpc", 0, TankCount)
    StartIndex = LoadNpcData(NpcSpawnList, ScenarioIni, "SupportNpc", StartIndex, SupportCount)
    StartIndex = LoadNpcData(NpcSpawnList, ScenarioIni, "RangedNpc", StartIndex, RandgedCount)
    StartIndex = LoadNpcData(NpcSpawnList, ScenarioIni, "CaptainNpc", StartIndex, CaptainCount)
End Sub

Private Function LoadNpcData(ByRef NpcSpawnList() As t_NpcSpawnInfo, ScenearioIni As clsIniManager, ByVal SectionName As String, _
                             ByVal StartIndex As Integer, ByVal NpcCount As Integer) As Long
    Dim NpcId, i As Integer
    NpcId = val(ScenearioIni.GetValue(SectionName, "NpcId"))
    For i = 1 To NpcCount
        NpcSpawnList(StartIndex + i).NpcId = NpcId
        NpcSpawnList(StartIndex + i).Waypoint = val(ScenearioIni.GetValue(SectionName, "Wp" & i))
    Next i
    LoadNpcData = NpcCount + StartIndex
End Function

Private Sub LoadTeamWeapoints(ByVal team As Integer, ByRef WpArray() As t_Waypoint, ByVal WpCount As Byte, ScenearioIni As clsIniManager)
    Dim i, j As Integer
    ReDim WpArray(1 To WpCount) As t_Waypoint
    Dim NextWpCount As Byte
    For i = 1 To WpCount
        WpArray(i).PosX = val(ScenearioIni.GetValue("T" & team & "WP" & i, "X"))
        WpArray(i).PosY = val(ScenearioIni.GetValue("T" & team & "WP" & i, "Y"))
        WpArray(i).NextWpCount = val(ScenearioIni.GetValue("T" & team & "WP" & i, "NextWpSize"))
        If WpArray(i).NextWpCount > 0 Then
            ReDim WpArray(i).NextWp(1 To WpArray(i).NextWpCount) As Byte
            For j = 1 To WpArray(i).NextWpCount
                WpArray(i).NextWp(j) = val(ScenearioIni.GetValue("T" & team & "WP" & i, "NextWp" & j))
            Next j
        End If
    Next i
End Sub

Private Sub LoadForbidenItems(ScenearioIni As clsIniManager)
    Dim forbiddenItemCount As Integer
    Dim i As Integer
    forbiddenItemCount = val(ScenearioIni.GetValue("FORBIDDEN", "Count"))
    ReDim ForbbidenItemId(forbiddenItemCount) As Integer
    For i = 0 To forbiddenItemCount - 1
        ForbbidenItemId(i) = val(ScenearioIni.GetValue("FORBIDDEN", "ItemId" & i + 1))
    Next i
End Sub

Private Sub RemoveForbiddenItems(ByVal UserIndex As Integer)
    Dim i As Integer
    With UserList(UserIndex)
        For i = 1 To UBound(.invent.Object)
            If IsForbidden(.invent.Object(i).ObjIndex) Then
                Call QuitarUserInvItem(UserIndex, i, .invent.Object(i).amount)
            End If
        Next i
        If .flags.ModificoInventario Then
            Call UpdateUserInv(True, UserIndex, 0)
        End If
    End With
End Sub

Private Sub CloseScenario()
On Error GoTo CloseScenario_Err:
100    Call ModLobby.ReturnAllPlayers(GenericGlobalLobby)
102    Set GenericGlobalLobby.Scenario = Nothing
104    Call ModLobby.UpdateLobbyState(GenericGlobalLobby, e_LobbyState.UnInitilized)
118    Call ClearMap
    Exit Sub
CloseScenario_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.CloseScenario ", Erl)
End Sub

Private Sub ClearMap()
On Error GoTo ClearMap_Err:
106    CustomScenarios.RemoveCustomScenario (MapNumbers(0))
107    CustomScenarios.RemoveCustomScenario (MapNumbers(1))
108    CustomScenarios.RemoveUpdateScenario (UpdateIndex)
       MapInfo(MapNumbers(TopMapNumber)).ForceUpdate = False
    Exit Sub
ClearMap_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.ClearMap", Erl)
End Sub

Private Function IsForbidden(ByVal ItemId As Integer) As Boolean
    Dim i As Integer
    If ItemId = 0 Then
        IsForbidden = False
        Exit Function
    End If
    For i = 0 To UBound(ForbbidenItemId) - 1
        If ForbbidenItemId(i) = ItemId Then
            IsForbidden = True
            Exit Function
        End If
    Next i
    IsForbidden = False
End Function

Public Function IBaseScenario_GetScenarioName() As String
    IBaseScenario_GetScenarioName = "Abordaje"
End Function

Public Sub IBaseScenario_BroadcastOpenScenario(ByVal CreatorIndex As Integer)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgNavalConquestIntro, "", e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Public Sub IBaseScenario_SendRules(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MsgNavalConquestInstructions, "", e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Public Function IBaseScenario_GetExtraInterfaces() As Long
    IBaseScenario_GetExtraInterfaces = eInventoryInterface
End Function

Public Sub IBaseScenario_Reset()
End Sub

Public Function IBaseScenario_RequiresSpawn() As Boolean
    IBaseScenario_RequiresSpawn = False
End Function

Public Function IBaseScenario_ValidateUser(ByVal UserIndex As Integer) As Integer
    Dim i As Integer
    With UserList(UserIndex)
        For i = 1 To UBound(.invent.Object)
            If IsForbidden(.invent.Object(i).ObjIndex) Then
                IBaseScenario_ValidateUser = 404
                Exit Function
            End If
        Next i
        IBaseScenario_ValidateUser = 0
    End With
End Function

Public Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MapNumbers(1)
    If UserList(UserIndex).flags.CurrentTeam = 1 Then
        SpawnPos.x = Team1SpawnX
    ElseIf UserList(UserIndex).flags.CurrentTeam = 2 Then
        SpawnPos.x = Team2SpawnX
    Else
        SpawnPos.x = Team1SpawnX
    End If
    SpawnPos.y = 50
    GenericGlobalLobby.SummonCoordinates = SpawnPos
End Sub

Public Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef NewState As e_LobbyState)
    If oldState = NewState Then
        Exit Sub
    End If
    If NewState = InProgress Then
        Call ModLobby.SummonAll(GenericGlobalLobby)
        Call SetTimer(CountdownTimer, 1000)
    End If
    If NewState = Closed Then
        Call MatchCompleted
    End If
End Sub

Public Sub IBaseScenario_PlayerKillNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Public Sub IBaseScenario_NPcKillPlayer(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Public Sub IBaseScenario_PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    
End Sub

Public Sub IBaseScenario_Update()
    Dim frametime As Long
    frametime = GetTickCount() - LastFrameTime
    LastFrameTime = GetTickCount()
    If GenericGlobalLobby.State = e_LobbyState.InProgress Then
        If CountdownTimer.Occurrences < 10 Then
            If UpdateTime(CountdownTimer, frametime) Then
                If CountdownTimer.Occurrences >= 10 Then
                    Call StartGame
                Else
                    Call SendData(SendTarget.toMap, MapNumbers(1), PrepareMessageConsoleMsg("El abordaje iniciara en " & 10 - CountdownTimer.Occurrences, e_FontTypeNames.FONTTYPE_GUILD))
                End If
            End If
            Exit Sub
        End If
        ElapsedTime = ElapsedTime + frametime
        ElapsedResapwnTime = ElapsedResapwnTime + frametime
        Call UpdateCaptains(frametime)
        If ElapsedResapwnTime > RespawnInterval Then
            Call RespawnPlayers
            ElapsedResapwnTime = 0
        End If
        If ElapsedTime > MaxTime Then
            Call SendData(SendTarget.toMap, MapNumbers(0), PrepareMessageConsoleMsg("Se acabo el tiempo!", e_FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.toMap, MapNumbers(1), PrepareMessageConsoleMsg("Se acabo el tiempo!", e_FontTypeNames.FONTTYPE_GUILD))
            Call CheckWinner
            Call MatchCompleted
        Else
            If UpdateTime(BroadCastTimer, frametime) Then
                Call BroadcastPendingTime(frametime)
            End If
            If UpdateTime(ScoreBroadcastTimer, frametime) Then
                'Call Board.PublishScoreboard(SendTarget.toMap, MapNumbers(0), GenericGlobalLobby, False)
                'Call Board.PublishScoreboard(SendTarget.toMap, MapNumbers(1), GenericGlobalLobby, False)
            End If
        End If
    ElseIf GenericGlobalLobby.State = e_LobbyState.Completed Then
        If CountdownTimer.Occurrences < 10 Then
            If UpdateTime(CountdownTimer, frametime) Then
                If CountdownTimer.Occurrences >= 10 Then
                    Call CloseScenario
                Else
                    Call SendData(SendTarget.toMap, MapNumbers(0), PrepareMessageConsoleMsg("El mapa se cerrara en " & 10 - CountdownTimer.Occurrences, e_FontTypeNames.FONTTYPE_GUILD))
                    Call SendData(SendTarget.toMap, MapNumbers(1), PrepareMessageConsoleMsg("El mapa se cerrara en " & 10 - CountdownTimer.Occurrences, e_FontTypeNames.FONTTYPE_GUILD))
                End If
            End If
            Exit Sub
        End If
    ElseIf GenericGlobalLobby.State = e_LobbyState.Closed Then
        Call CloseScenario
    End If
End Sub

Public Sub IBaseScenario_OnUserDisconnected(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If UserLoot.Exists(UserIndex) Then
            Dim i, ItemTeam As Integer
            ItemTeam = UserLoot.Item(UserIndex) \ 1000
            For i = LBound(.invent.Object) To UBound(.invent.Object)
                If .invent.Object(i).ObjIndex = TeamCargo(ItemTeam).CargoItemId Then
                    Call DropObj(UserIndex, i, .invent.Object(i).amount, .pos.Map, .pos.x, .pos.y)
                    Exit For
                End If
            Next i
        End If
    End With
    Call RemoveForbiddenItems(UserIndex)
End Sub

Public Sub IBaseScenario_OnUserReconnect(ByVal UserIndex As Integer)
End Sub

Private Sub StartGame()
    Call SetTimer(BroadCastTimer, 60000)
    Call SetTimer(ScoreBroadcastTimer, 20000)
    LastFrameTime = GetTickCount()
    Call SendData(SendTarget.toMap, MapNumbers(1), PrepareMessageConsoleMsg("Marineros al abordaje!!!", e_FontTypeNames.FONTTYPE_GUILD))
    ElapsedTime = 0
    Call ConfigMapEvent
End Sub

Private Sub Respawn(ByVal player As Integer)
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MapNumbers(0)
    SpawnPos.x = 10
    SpawnPos.y = 10
    If UserList(player).flags.CurrentTeam = 1 Then
        SpawnPos.x = Team1Waypoints(RespawnWp).PosX
        SpawnPos.y = Team1Waypoints(RespawnWp).PosY
    ElseIf UserList(player).flags.CurrentTeam = 2 Then
        SpawnPos.x = Team2Waypoints(RespawnWp).PosX
        SpawnPos.y = Team2Waypoints(RespawnWp).PosY
    End If
    
    Call WarpToLegalPos(player, SpawnPos.Map, SpawnPos.x, SpawnPos.y, True, True)
    Call RevivirUsuario(player)
    UserList(player).Counters.timeFx = 2
114 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessageParticleFX(UserList(player).Char.charindex, e_ParticulasIndex.Curar, 100, False, , UserList(player).pos.x, UserList(player).pos.y))
116 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessagePlayWave("104", UserList(player).pos.x, UserList(player).pos.y))
End Sub

Private Sub MatchCompleted()
    MapInfo(MapNumbers(0)).Seguro = 1
    MapInfo(MapNumbers(1)).Seguro = 1
    Call SetTimer(CountdownTimer, 1000)
    Dim i As Integer
    For i = 0 To GenericGlobalLobby.RegisteredPlayers - 1
        If IsValidUserRef(GenericGlobalLobby.Players(i).user) Then
            Call RemoveForbiddenItems(GenericGlobalLobby.Players(i).user.ArrayIndex)
        End If
    Next i
    GenericGlobalLobby.State = Completed
    ElapsedTime = MaxTime
    'Call Board.PublishScoreboard(SendTarget.toMap, MapNumbers(0), GenericGlobalLobby, True)
    'Call Board.PublishScoreboard(SendTarget.toMap, MapNumbers(1), GenericGlobalLobby, True)
    Call CustomScenarios.ClearMap(MapNumbers(0))
    Call CustomScenarios.ClearMap(MapNumbers(1))
    MapInfo(MapNumbers(0)).SafeFightMap = False
    MapInfo(MapNumbers(1)).SafeFightMap = False
    Dim NpcIndex As Integer
    Do While ModAreas.ConnGroups(MapNumbers(0)).RegisteredNpc > 0
        NpcIndex = ModAreas.ConnGroups(MapNumbers(0)).NpcForAi(0)
        Call QuitarNPC(NpcIndex, eClearHunt)
    Loop
End Sub

Private Sub BroadcastPendingTime(ByVal frametime As Long)
    Dim pendingTime As Long
    'we add frame time here again so we round up
    pendingTime = (MaxTime - ElapsedTime + frametime) / 1000
    If pendingTime > MaxTime / 2000 Then
        Exit Sub
    ElseIf pendingTime >= 60 Then
        pendingTime = pendingTime / 60
        Call SendData(SendTarget.toMap, MapNumbers(0), PrepareMessageConsoleMsg("quedan " & pendingTime & " minutos.", e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.toMap, MapNumbers(1), PrepareMessageConsoleMsg("quedan " & pendingTime & " minutos.", e_FontTypeNames.FONTTYPE_GUILD))
    ElseIf pendingTime <= 10 Then
        Call SendData(SendTarget.toMap, MapNumbers(0), PrepareMessageConsoleMsg(pendingTime, e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.toMap, MapNumbers(1), PrepareMessageConsoleMsg(pendingTime, e_FontTypeNames.FONTTYPE_GUILD))
    End If
    If pendingTime <= 1 Then
        Call SetTimer(BroadCastTimer, 1000)
    End If
End Sub

Private Sub RespawnPlayers()
    Dim key As Variant
    Dim Keys As Variant
    Dim currentTime As Long
    Dim RespawnInfo As clsRespawnInfo
    Dim i As Integer
    currentTime = GetTickCount()
    Keys = PlayerRespawn.Keys
    For Each key In Keys
        Set RespawnInfo = PlayerRespawn.Item(key)
        If currentTime - RespawnInfo.RespawnTime > PlayerMinRespawnTime Then
            Call Respawn(key)
            For i = 0 To RespawnInfo.EquipedElementCount - 1
                Call EquiparInvItem(key, RespawnInfo.GetActiveSlot(i))
            Next i
            If RespawnInfo.HasPowerUp Then
                With UserList(key)
                    .Stats.UserAtributos(1) = RespawnInfo.Strength
                    .Stats.UserAtributos(2) = RespawnInfo.Agility
                    .flags.DuracionEfecto = RespawnInfo.PowerUpTime
                    Call WriteFYA(key)
                End With
            End If
            Call PlayerRespawn.Remove(key)
        End If
    Next key
End Sub

Private Sub ConfigMapEvent()
    Dim i As Integer
    For i = 0 To UBound(MapNumbers)
        MapInfo(MapNumbers(i)).Seguro = 0
        MapInfo(MapNumbers(i)).SinInviOcul = True
        MapInfo(MapNumbers(i)).SafeFightMap = True
        MapInfo(MapNumbers(i)).FriendlyFire = False
    Next i
    MapInfo(MapNumbers(TopMapNumber)).ForceUpdate = True
    For i = 1 To UBound(NpcSpawnList)
        Call SpawnNpc(NpcSpawnList(i), Team1Waypoints, 1)
        Call SpawnNpc(NpcSpawnList(i), Team2Waypoints, 2)
    Next i
    For i = 1 To UBound(TileBlocks)
        Call BlockAndInform(MapNumbers(1), TileBlocks(i).PosX, TileBlocks(i).PosY, 0)
    Next i
    For i = LBound(TeamCargo(1).ItemList) To UBound(TeamCargo(1).ItemList)
        Call SpawnItem(TeamCargo(1).ItemList(i).ItemIndex, 1, TeamCargo(1).ItemList(i).WorldPos.Map, _
                       TeamCargo(1).ItemList(i).WorldPos.x, TeamCargo(1).ItemList(i).WorldPos.y)
        Call SpawnItem(TeamCargo(2).ItemList(i).ItemIndex, 1, TeamCargo(2).ItemList(i).WorldPos.Map, _
                       TeamCargo(2).ItemList(i).WorldPos.x, TeamCargo(2).ItemList(i).WorldPos.y)
    Next i
End Sub

Private Sub SpawnNpc(ByRef NpcInfo As t_NpcSpawnInfo, ByRef Waypoints() As t_Waypoint, ByVal team As Byte)
    Dim SpawnPos As t_WorldPos
    Dim NpcId As Integer
    SpawnPos.Map = MapNumbers(0)
    SpawnPos.x = Waypoints(NpcInfo.Waypoint).PosX
    SpawnPos.y = Waypoints(NpcInfo.Waypoint).PosY
    NpcId = CrearNPC(NpcInfo.NpcId, MapNumbers(0), SpawnPos)
    If NpcId > 0 Then
        If CaptianNpcId = NpcInfo.NpcId Then
            Call SetNpcRef(CaptainState(team).NpcRef, NpcId)
            Call IncreaseSingle(NpcList(NpcId).Modifiers.MovementSpeed, 0.2)
            Call UpdateNpcSpeed(NpcId)
        End If
        NpcList(NpcId).flags.Respawn = 1
        NpcList(NpcId).flags.team = team
    Else
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("No se pudo spawnear el npc " & NpcInfo.NpcId & ", revisar.", e_FontTypeNames.FONTTYPE_TALK))
        Call TraceError(Err.Number, "failed to spawn Npc " & NpcInfo.NpcId, "NavalBoarding.SpawnNpc ", Erl)
    End If
End Sub
Private Sub ConfigMapLobby()
    Dim i As Integer
    Call ModLobby.SetTeamCount(GenericGlobalLobby, 2, GenericGlobalLobby.TeamType)
    For i = 0 To UBound(MapNumbers)
        MapInfo(MapNumbers(i)).Seguro = 1
        MapInfo(MapNumbers(i)).SinInviOcul = True
        MapInfo(MapNumbers(i)).DropItems = False
    Next i
End Sub

Private Function GetClosestWp(ByRef Npc As t_Npc, ByRef Waypoints() As t_Waypoint) As Integer
    Dim ClosestWp As Integer
    Dim MinDistance As Single
    Dim i As Integer
    Dim CurrWpDistance As Single
    ClosestWp = -1
    For i = LBound(Waypoints) To UBound(Waypoints)
        CurrWpDistance = Distance(Npc.pos.x, Npc.pos.y, Waypoints(i).PosX, Waypoints(i).PosY)
        If ClosestWp < 0 Or CurrWpDistance < MinDistance Then
            ClosestWp = i
            MinDistance = CurrWpDistance
        End If
    Next i
    GetClosestWp = ClosestWp
    If ClosestWp < 0 Then Exit Function
    If Waypoints(ClosestWp).NextWpCount > 0 Then 'get the closes wp from the next ones
        ClosestWp = -1
        For i = 1 To Waypoints(GetClosestWp).NextWpCount
            Dim NextPossibleWp As Byte
            NextPossibleWp = Waypoints(GetClosestWp).NextWp(i)
            CurrWpDistance = Distance(Npc.pos.x, Npc.pos.y, _
                                      Waypoints(NextPossibleWp).PosX, _
                                      Waypoints(NextPossibleWp).PosY)
            If ClosestWp < 0 Or CurrWpDistance < MinDistance Then
                ClosestWp = NextPossibleWp
                MinDistance = CurrWpDistance
            End If
        Next i
    End If
    If ClosestWp < 0 Then Exit Function
    GetClosestWp = ClosestWp
End Function

Public Sub IBaseScenario_GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
    With NpcList(NpcIndex)
        Dim TargetWp As Byte
        If .flags.team = 1 Then
            TargetWp = GetClosestWp(NpcList(NpcIndex), Team1Waypoints)
            If TargetWp <= 0 Then
                PosX = NpcList(NpcIndex).pos.x
                PosY = NpcList(NpcIndex).pos.y
            End If
            PosX = Team1Waypoints(TargetWp).PosX
            PosY = Team1Waypoints(TargetWp).PosY
        ElseIf .flags.team = 2 Then
            TargetWp = GetClosestWp(NpcList(NpcIndex), Team2Waypoints)
            If TargetWp <= 0 Then
                PosX = NpcList(NpcIndex).pos.x
                PosY = NpcList(NpcIndex).pos.y
            End If
            PosX = Team2Waypoints(TargetWp).PosX
            PosY = Team2Waypoints(TargetWp).PosY
        End If
    End With
End Sub

Public Function GetOposingFaction(ByVal team As Integer) As Integer
    GetOposingFaction = IIf(team = 1, 2, 1)
End Function
Public Sub IBaseScenario_NpcDie(ByVal NpcIndex)
    If NpcList(NpcIndex).Numero = CaptianNpcId Then
        Call SendData(SendTarget.toMap, MapNumbers(0), PrepareMessageLocaleMsg(MsgCaptainIsDeath, NpcList(NpcIndex).flags.team, e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.toMap, MapNumbers(0), PrepareMessageLocaleMsg(MsgTeamNumberWin, GetOposingFaction(NpcList(NpcIndex).flags.team), e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.toMap, MapNumbers(1), PrepareMessageLocaleMsg(MsgCaptainIsDeath, NpcList(NpcIndex).flags.team, e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.toMap, MapNumbers(1), PrepareMessageLocaleMsg(MsgTeamNumberWin, GetOposingFaction(NpcList(NpcIndex).flags.team), e_FontTypeNames.FONTTYPE_GUILD))
        Call MatchCompleted
    End If
End Sub

Public Sub IBaseScenario_UserDie(ByVal UserIndex)
    Dim RespawnInfo As clsRespawnInfo
    Set RespawnInfo = New clsRespawnInfo
    RespawnInfo.RespawnTime = GetTickCount()
    With UserList(UserIndex)
        Dim i As Integer
        For i = LBound(.invent.Object) To UBound(.invent.Object)
            If .invent.Object(i).Equipped Then
                Call RespawnInfo.AddEquipedElement(i)
            End If
        Next i
        If .flags.TomoPocion Then
            RespawnInfo.HasPowerUp = True
            RespawnInfo.PowerUpTime = .flags.DuracionEfecto
            RespawnInfo.Strength = .Stats.UserAtributos(1)
            RespawnInfo.Agility = .Stats.UserAtributos(2)
        End If
        If UserLoot.Exists(UserIndex) Then
            Dim ItemTeam As Integer
            ItemTeam = UserLoot.Item(UserIndex) \ 1000
            
            For i = LBound(.invent.Object) To UBound(.invent.Object)
                If .invent.Object(i).ObjIndex = TeamCargo(ItemTeam).CargoItemId Then
                    Call DropObj(UserIndex, i, .invent.Object(i).amount, .pos.Map, .pos.x, .pos.y)
                    Exit For
                End If
            Next i
        End If
    End With
    Call PlayerRespawn.Add(UserIndex, RespawnInfo)
    
End Sub

Public Function IInventoryInterface_UserCanPickUpItem(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        Dim ObjIndex As Integer
        ObjIndex = MapData(.pos.Map, .pos.x, .pos.y).ObjInfo.ObjIndex
        Dim ItemTeam As Integer
        ItemTeam = GetCargoItemTeam(ObjIndex)
        If ItemTeam > 0 Then
            IInventoryInterface_UserCanPickUpItem = TryPickupItemAt(UserIndex)
            Exit Function
        End If
    End With
    IInventoryInterface_UserCanPickUpItem = True
End Function

Public Sub IInventoryInterface_UserDidPickupItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    Dim i As Integer
    Dim ItemTeam As Integer
    With UserList(UserIndex)
        ItemTeam = GetCargoItemTeam(ItemIndex)
        If ItemTeam = 0 Then
            Exit Sub
        End If
        For i = LBound(TeamCargo(ItemTeam).ItemList) To UBound(TeamCargo(ItemTeam).ItemList)
            If TeamCargo(ItemTeam).ItemList(i).WorldPos.x = .pos.x And _
               TeamCargo(ItemTeam).ItemList(i).WorldPos.y = .pos.y And _
               TeamCargo(ItemTeam).ItemList(i).WorldPos.Map = .pos.Map Then
               TeamCargo(ItemTeam).ItemList(i).CarrierIndex = UserIndex
               Call UserLoot.Add(UserIndex, TeamCargo(ItemTeam).ItemList(i).id)
            Exit Sub
            End If
        Next i
    End With
End Sub

Public Function IInventoryInterface_UserCanDropItemAt(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer) As Boolean
    With UserList(UserIndex)
        Dim ItemTeam As Integer
        ItemTeam = GetCargoItemTeam(.invent.Object(Slot).ObjIndex)
        If ItemTeam > 0 Then
            If MapData(Map, TileX, TileY).ObjInfo.ObjIndex > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MsgCantDropCargoAtPos, "", e_FontTypeNames.FONTTYPE_GUILD))
                Exit Function
            End If
        End If
        IInventoryInterface_UserCanDropItemAt = True
        Exit Function
    End With
End Function

'inform that user want to drop an item
Public Sub IInventoryInterface_UserDropItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    Dim ItemTeam As Integer
    With UserList(UserIndex)
        ItemTeam = GetCargoItemTeam(.invent.Object(Slot).ObjIndex)
        If ItemTeam > 0 Then
            If Not UserLoot.Exists(UserIndex) Then
                'Debug.Assert False
                'TODO: LOG ERROR
                Exit Sub
            End If
            Dim id As Integer
            id = UserLoot.Item(UserIndex)
            id = id - ItemTeam * 1000
            UserLoot.Remove (UserIndex)
            If TeamCargo(ItemTeam).ItemList(id).CarrierIndex = UserIndex Then
                TeamCargo(ItemTeam).ItemList(id).WorldPos.Map = Map
                TeamCargo(ItemTeam).ItemList(id).WorldPos.x = TileX
                TeamCargo(ItemTeam).ItemList(id).WorldPos.y = TileY
                TeamCargo(ItemTeam).ItemList(id).CarrierIndex = 0
                If IsInsideArea(TeamCargo(.flags.CurrentTeam).DropArea, TileX, TileY) Or _
                   IsInsideArea(TeamCargo(.flags.CurrentTeam).OriginalCargoArea, TileX, TileY) Then
                    Call ValidateCargoAreasForTeam(.flags.CurrentTeam)
                End If
            Else
                Debug.Assert False
                'TODO: LOG ERROR
                Exit Sub
            End If
            
        End If
    End With
End Sub

Private Sub SpawnItem(ByVal ItemId As Integer, ByVal StackSize As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    Dim MiObj As t_Obj
    MiObj.amount = StackSize
    MiObj.ObjIndex = ItemId
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = Map
    SpawnPos.x = TileX
    SpawnPos.y = TileY
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
End Sub

Private Function GetCargoItemTeam(ByVal ItemIndex As Integer) As Integer
    If ItemIndex = TeamCargo(1).CargoItemId Then
        GetCargoItemTeam = 1
    ElseIf ItemIndex = TeamCargo(2).CargoItemId Then
        GetCargoItemTeam = 2
    End If
End Function

Private Function GetOposingTeam(ByVal TeamIndex As Integer) As Integer
    If TeamIndex = 1 Then GetOposingTeam = 2
    If TeamIndex = 2 Then GetOposingTeam = 1
End Function

Private Function TryPickupItemAt(ByVal UserIndex As Integer) As Boolean
    With UserList(UserIndex)
        If UserLoot.Exists(UserIndex) Then
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MsgCantCarryMoreThanOne, "", e_FontTypeNames.FONTTYPE_GUILD))
            Exit Function
        End If
        Dim ObjIndex As Integer
        ObjIndex = MapData(.pos.Map, .pos.x, .pos.y).ObjInfo.ObjIndex
        Dim ItemTeam As Integer
        ItemTeam = GetCargoItemTeam(ObjIndex)
        If ItemTeam > 0 Then
            If ItemTeam = .flags.CurrentTeam Then
                If IsInsideArea(TeamCargo(ItemTeam).OriginalCargoArea, .pos.x, .pos.y) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MsgCantPickFromYourStorage, "", e_FontTypeNames.FONTTYPE_GUILD))
                    TryPickupItemAt = False
                    Exit Function
                End If
            Else
                If IsInsideArea(TeamCargo(.flags.CurrentTeam).DropArea, .pos.x, .pos.y) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MsgCantPickFromYourStorage, "", e_FontTypeNames.FONTTYPE_GUILD))
                    TryPickupItemAt = False
                    Exit Function
                End If
            End If
        End If
    End With
    TryPickupItemAt = True
End Function

Private Function IsInsideArea(ByRef Area As t_CargoArea, ByVal PosX As Integer, ByVal PosY As Integer) As Boolean
100     If PosX < Area.TopLeft.PosX Then Exit Function
102     If PosX > Area.BottomRight.PosX Then Exit Function
104     If PosY < Area.TopLeft.PosY Then Exit Function
106     If PosY > Area.BottomRight.PosY Then Exit Function
108     IsInsideArea = True
End Function

Private Sub ValidateCargoAreasForTeam(ByVal team As Integer)
    If GetTeamCargoCount(team) < UBound(TeamCargo(team).ItemList) * 2 Then
        Exit Sub
    End If
    Call SendData(SendTarget.toMap, MapNumbers(0), PrepareMessageLocaleMsg(MsgTeamGotAllCargo, team, e_FontTypeNames.FONTTYPE_GUILD))
    Call SendData(SendTarget.toMap, MapNumbers(1), PrepareMessageLocaleMsg(MsgTeamGotAllCargo, team, e_FontTypeNames.FONTTYPE_GUILD))
    Call MatchCompleted
End Sub

Private Function GetTeamCargoCount(ByVal team As Integer) As Integer
    Dim i, j As Integer
    Dim CargoCount As Integer
    For i = TeamCargo(team).OriginalCargoArea.TopLeft.PosX To _
        TeamCargo(team).OriginalCargoArea.BottomRight.PosX
        For j = TeamCargo(team).OriginalCargoArea.TopLeft.PosY To _
            TeamCargo(team).OriginalCargoArea.BottomRight.PosY
            If MapData(MapNumbers(1), i, j).ObjInfo.ObjIndex = TeamCargo(1).CargoItemId Or _
               MapData(MapNumbers(1), i, j).ObjInfo.ObjIndex = TeamCargo(2).CargoItemId Then
                CargoCount = CargoCount + 1
            End If
        Next j
    Next i
    For i = TeamCargo(team).DropArea.TopLeft.PosX To _
        TeamCargo(team).DropArea.BottomRight.PosX
        For j = TeamCargo(team).DropArea.TopLeft.PosY To _
            TeamCargo(team).DropArea.BottomRight.PosY
            If MapData(MapNumbers(1), i, j).ObjInfo.ObjIndex = TeamCargo(1).CargoItemId Or _
               MapData(MapNumbers(1), i, j).ObjInfo.ObjIndex = TeamCargo(2).CargoItemId Then
                CargoCount = CargoCount + 1
            End If
        Next j
    Next i
    GetTeamCargoCount = CargoCount
End Function

Private Sub CheckWinner()
    Dim Team1Cargo, Team2Cargo As Integer
    Team1Cargo = GetTeamCargoCount(1)
    Team2Cargo = GetTeamCargoCount(2)
    Call SendData(SendTarget.toMap, MapNumbers(TopMapNumber), PrepareMessageLocaleMsg(MsgBothTeamCargo, Team1Cargo & "¬" & Team2Cargo, e_FontTypeNames.FONTTYPE_GUILD))
    Call SendData(SendTarget.toMap, MapNumbers(CargoMapNumber), PrepareMessageLocaleMsg(MsgBothTeamCargo, Team1Cargo & "¬" & Team2Cargo, e_FontTypeNames.FONTTYPE_GUILD))
    If Team1Cargo > Team2Cargo Then
        Call SendData(SendTarget.toMap, MapNumbers(TopMapNumber), PrepareMessageLocaleMsg(MsgNavalConquestWinnerTeam, 1, e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.toMap, MapNumbers(CargoMapNumber), PrepareMessageLocaleMsg(MsgNavalConquestWinnerTeam, 1, e_FontTypeNames.FONTTYPE_GUILD))
    ElseIf Team2Cargo > Team1Cargo Then
        Call SendData(SendTarget.toMap, MapNumbers(TopMapNumber), PrepareMessageLocaleMsg(MsgNavalConquestWinnerTeam, 2, e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.toMap, MapNumbers(CargoMapNumber), PrepareMessageLocaleMsg(MsgNavalConquestWinnerTeam, 2, e_FontTypeNames.FONTTYPE_GUILD))
    Else
        Call SendData(SendTarget.toMap, MapNumbers(TopMapNumber), PrepareMessageLocaleMsg(MsgNavalConquestEvenMatch, "", e_FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.toMap, MapNumbers(CargoMapNumber), PrepareMessageLocaleMsg(MsgNavalConquestEvenMatch, "", e_FontTypeNames.FONTTYPE_GUILD))
    End If
End Sub

Private Sub UpdateCaptains(ByVal DeltaTime As Long)
    If ElapsedTime <= FireStartTime Then Exit Sub
    Call HandleCaptain(DeltaTime, 1)
    Call HandleCaptain(DeltaTime, 2)
End Sub

Private Sub HandleCaptain(ByVal DeltaTime As Long, ByVal CaptainIndex As Integer)
    With CaptainState(CaptainIndex)
        If Not IsValidNpcRef(CaptainState(CaptainIndex).NpcRef) Then Exit Sub
        .TimeSinceLastFire = .TimeSinceLastFire + DeltaTime
        If .TimeSinceLastFire > .NextFireTime Then
            .TimeSinceLastFire = 0
            .NextFireTime = RandomNumber(MinFireTime, MaxFireTime)
            If Not CaptainFire(CaptainIndex) Then
                .NextFireTime = .NextFireTime / 2
            Else
                Call SendData(SendTarget.toMap, NpcList(.NpcRef.ArrayIndex).pos.Map, PrepareMessageLocaleMsg(MsgCaptainFire, CaptainIndex, e_FontTypeNames.FONTTYPE_WARNING))
                Call SendData(SendTarget.toMap, NpcList(.NpcRef.ArrayIndex).pos.Map, PrepareLocaleChatOverHead(MsgFire, "", NpcList(.NpcRef.ArrayIndex).Char.charindex, vbRed, , NpcList(.NpcRef.ArrayIndex).pos.x, NpcList(.NpcRef.ArrayIndex).pos.y))
            End If
        End If
    End With
End Sub

Private Function CaptainFire(ByVal CaptainIndex As Integer) As Boolean
    With CaptainState(CaptainIndex)
        Dim TargetTeam As Integer
        TargetTeam = GetOposingTeam(CaptainIndex)
        Dim PossibleTargets(10) As Integer
        Dim PossibleTargetCount As Integer
        Dim UserIndex As Integer
        Dim i As Integer
        For i = 1 To ModAreas.ConnGroups(NpcList(.NpcRef.ArrayIndex).pos.Map).CountEntrys
            UserIndex = ModAreas.ConnGroups(NpcList(.NpcRef.ArrayIndex).pos.Map).UserEntrys(i)
            If UserList(UserIndex).flags.CurrentTeam = TargetTeam Then
                If UserList(UserIndex).Stats.MinHp > 0 Then
                    PossibleTargets(PossibleTargetCount) = UserIndex
                    PossibleTargetCount = PossibleTargetCount + 1
                End If
            End If
        Next i
        If PossibleTargetCount = 0 Then
            Exit Function
        End If
        Dim CannonCount As Integer
        Dim TargetIndex As Integer
        CannonCount = Min(3, PossibleTargetCount)
        Dim Particula As Integer
        Dim Tiempo    As Long
        Dim CannonProyectile As Integer
        Dim TargetPos As t_WorldPos
        CannonProyectile = 4
        Particula = Hechizos(CannonSpell).Particle
        Tiempo = Hechizos(CannonSpell).TimeParticula
        For i = 0 To CannonCount - 1
            TargetIndex = RandomNumber(0, PossibleTargetCount - 1)
            UserIndex = PossibleTargets(TargetIndex)
            TargetPos = UserList(UserIndex).pos
            Call SendData(SendTarget.toMap, NpcList(.NpcRef.ArrayIndex).pos.Map, PrepareMessageParticleFX(NpcList(.NpcRef.ArrayIndex).Char.charindex, Particula, Tiempo, False, , NpcList(.NpcRef.ArrayIndex).pos.x, NpcList(.NpcRef.ArrayIndex).pos.y))
            Call SendData(SendTarget.toMap, NpcList(.NpcRef.ArrayIndex).pos.Map, PrepareCreateProjectile(NpcList(.NpcRef.ArrayIndex).pos.x, NpcList(.NpcRef.ArrayIndex).pos.y, TargetPos.x, TargetPos.y, CannonProyectile))
            If Hechizos(CannonSpell).wav <> 0 Then Call SendData(SendTarget.ToNPCAliveArea, .NpcRef.ArrayIndex, PrepareMessagePlayWave(Hechizos(CannonSpell).wav, NpcList(.NpcRef.ArrayIndex).pos.x, NpcList(.NpcRef.ArrayIndex).pos.y))
            Call CreateDelayedBlast(.NpcRef.ArrayIndex, eNpc, TargetPos.Map, TargetPos.x, TargetPos.y, Hechizos(CannonSpell).EotId, -1)
            PossibleTargets(TargetIndex) = PossibleTargets(PossibleTargetCount - 1)
            PossibleTargetCount = PossibleTargetCount - 1
        Next i
    End With
    CaptainFire = True
End Function
