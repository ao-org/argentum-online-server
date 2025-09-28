VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScenarioDeathMatch"
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
Const configPath = "../dat/Scenarios/DeathMatch.ini"
Implements IBaseScenario

Private Type t_SpawnItems
    ItemId As Integer
    ItemCount As Integer
    StackCount As Single
    InitialStacks As Single
    timer As t_Timer
End Type

Private Type t_RemoveItem
    ItemId As Integer
    ItemCount As Integer
End Type

Private Type t_UserItemInfo
    RemovedItems(10) As t_RemoveItem
    RemovedItemCount As Integer
    UserId As Long
End Type

Private Type t_UserList
    UserList(100) As t_UserItemInfo
    RegisteredUserCount As Integer
End Type

Dim SpawnedUserList            As t_UserList
Dim MapNumber                  As Integer
Dim TargetScore                As Integer
Dim MaxTime                    As Long
Dim PlayerRespawnTime          As Long
Dim CountdownTimer             As t_Timer
Dim BroadCastTimer             As t_Timer
Dim ScoreBroadcastTimer        As t_Timer
Dim ForbbidenItemId()          As Integer
Dim UpdateIndex                As Integer
Dim Board                      As New ScoreBoard
Dim LastFrameTime              As Long
Dim ElapsedTime                As Long
Dim PlayerRespawn              As New Dictionary
Dim ItemsToSpawn()             As t_SpawnItems
Dim MapDropItems               As Boolean
Dim PrevMapState               As Integer
Dim SafeFigthState             As Boolean
Dim FriendlyFireState          As Boolean
Dim FightAreaTopLeft           As t_Position
Dim FightAreaBottomRight       As t_Position
Dim PlayerAreaIncreaseModifier As Single
Dim RingLimiterHItemID         As Long
Dim RingLimiterVItemID         As Long
Dim RingLimiterTRItemID        As Long
Dim RingLimiterTLItemID        As Long
Dim RingLimiterBRItemID        As Long
Dim RingLimiterBLItemID        As Long
Dim MinArenaSize               As Integer
Dim LobbyIndex                 As Integer

Private Sub Class_Initialize()
    On Error GoTo DeathMatch_Initialize_Err:
    Debug.Assert FileExist(DatPath & configPath)
    Dim ScenarioIni As clsIniManager
    Dim OriginalMap As Integer
    Set ScenarioIni = New clsIniManager
    Call ScenarioIni.Initialize(DatPath & configPath)
    OriginalMap = val(ScenarioIni.GetValue("INIT", "MapNumber"))
    MapNumber = InstanceManager.GetNextAvailableInstance
    Call InstanceManager.CloneMap(OriginalMap, MapNumber)
    MaxTime = val(ScenarioIni.GetValue("INIT", "MaxTime"))
    PlayerRespawnTime = val(ScenarioIni.GetValue("INIT", "PlayerRespawnTime"))
    MaxTime = MaxTime * 1000
    PlayerRespawnTime = PlayerRespawnTime * 1000
    PlayerAreaIncreaseModifier = val(ScenarioIni.GetValue("INIT", "PlayerAreaIncreaseModifier")) / 10
    RingLimiterHItemID = val(ScenarioIni.GetValue("INIT", "RingLimiterHItemId"))
    RingLimiterVItemID = val(ScenarioIni.GetValue("INIT", "RingLimiterVItemID"))
    RingLimiterTRItemID = val(ScenarioIni.GetValue("INIT", "RingLimiterTRItemID"))
    RingLimiterTLItemID = val(ScenarioIni.GetValue("INIT", "RingLimiterTLItemID"))
    RingLimiterBRItemID = val(ScenarioIni.GetValue("INIT", "RingLimiterBRItemID"))
    RingLimiterBLItemID = val(ScenarioIni.GetValue("INIT", "RingLimiterBLItemID"))
    MinArenaSize = val(ScenarioIni.GetValue("INIT", "MinArenaSize"))
    TargetScore = val(ScenarioIni.GetValue("INIT", "TargetScore"))
    Call LoadForbidenItems(ScenarioIni)
    UpdateIndex = AddUpdateScenario(Me)
    Call CustomScenarios.SetCustomScenario(MapNumber, Me)
    Call LoadItemsToSpawn(ScenarioIni)
    Exit Sub
DeathMatch_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "DeathMatch.Initialize ", Erl)
End Sub

Private Sub LoadItemsToSpawn(ScenearioIni As clsIniManager)
    Dim ItemCount    As Integer
    Dim i            As Integer
    Dim spawInterval As Long
    ItemCount = val(ScenearioIni.GetValue("INIT", "SpawnItems"))
    ReDim ItemsToSpawn(ItemCount) As t_SpawnItems
    For i = 0 To ItemCount - 1
        ItemsToSpawn(i).ItemId = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "ItemId"))
        ItemsToSpawn(i).ItemCount = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "ItemCount"))
        ItemsToSpawn(i).StackCount = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "StackCount")) / 10
        ItemsToSpawn(i).InitialStacks = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "InitialStacks")) / 10
        spawInterval = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "SpawnInterval"))
        Call SetTimer(ItemsToSpawn(i).timer, spawInterval * 1000)
    Next i
End Sub

Private Sub LoadForbidenItems(ScenearioIni As clsIniManager)
    Dim forbiddenItemCount As Integer
    Dim i                  As Integer
    forbiddenItemCount = val(ScenearioIni.GetValue("FORBIDDEN", "Count"))
    ReDim ForbbidenItemId(forbiddenItemCount) As Integer
    For i = 0 To forbiddenItemCount - 1
        ForbbidenItemId(i) = val(ScenearioIni.GetValue("FORBIDDEN", "ItemId" & i + 1))
    Next i
End Sub

Private Sub UpdateFightAreaForPlayerCount()
    Dim Size As Integer
    Size = MinArenaSize + LobbyList(LobbyIndex).RegisteredPlayers * PlayerAreaIncreaseModifier
    FightAreaTopLeft.x = 50 - Size
    FightAreaTopLeft.y = 50 - Size
    FightAreaBottomRight.x = 50 + Size
    FightAreaBottomRight.y = 50 + Size
End Sub

Private Sub AddItemToUser(ByVal UserIndex As Integer, ByVal ItemId As Integer, ByVal ItemCount As Integer)
    With UserList(UserIndex)
        Dim Slot As Integer
        For Slot = 1 To UBound(.invent.Object)
            If .invent.Object(Slot).ObjIndex = 0 Then
                .invent.Object(Slot).ObjIndex = ItemId
                .invent.Object(Slot).amount = ItemCount
                Exit Sub
            End If
        Next Slot
    End With
End Sub

Private Sub RemoveForbiddenItems(ByVal UserIndex As Integer)
    Dim i As Integer
    Dim j As Integer
    With UserList(UserIndex)
        For i = 1 To UBound(.invent.Object)
            If IsForbidden(.invent.Object(i).ObjIndex) Then
                Call QuitarUserInvItem(UserIndex, i, .invent.Object(i).amount)
            End If
        Next i
        If .flags.ModificoInventario Then
            Call UpdateUserInv(True, UserIndex, 0)
        End If
        For i = 0 To SpawnedUserList.RegisteredUserCount - 1
            If SpawnedUserList.UserList(i).UserId = .Id Then
                Exit For
            End If
        Next i
        If i < SpawnedUserList.RegisteredUserCount Then
            For j = 0 To SpawnedUserList.UserList(i).RemovedItemCount - 1
                If SpawnedUserList.UserList(i).RemovedItems(j).ItemId > 0 Then
                    Call AddItemToUser(UserIndex, SpawnedUserList.UserList(i).RemovedItems(j).ItemId, SpawnedUserList.UserList(i).RemovedItems(j).ItemCount)
                End If
            Next j
            SpawnedUserList.UserList(i) = SpawnedUserList.UserList(SpawnedUserList.RegisteredUserCount - 1)
            SpawnedUserList.RegisteredUserCount = SpawnedUserList.RegisteredUserCount - 1
            Call UpdateUserInv(True, UserIndex, 0)
        End If
    End With
End Sub

Private Sub LimitFightArea()
    Dim x     As Integer
    Dim y     As Integer
    Dim MiObj As t_Obj
    MiObj.amount = 1
    MiObj.ObjIndex = RingLimiterHItemID
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MapNumber
    Call UpdateFightAreaForPlayerCount
    For x = FightAreaTopLeft.x To FightAreaBottomRight.x
        SpawnPos.x = x
        SpawnPos.y = FightAreaTopLeft.y - 1
        Call TirarItemAlPiso(SpawnPos, MiObj, 0)
        SpawnPos.y = FightAreaBottomRight.y + 1
        Call TirarItemAlPiso(SpawnPos, MiObj, 0)
        Call BlockAndInform(MapNumber, x, FightAreaTopLeft.y - 1, 1)
        Call BlockAndInform(MapNumber, x, FightAreaBottomRight.y + 1, 1)
    Next x
    MiObj.ObjIndex = RingLimiterVItemID
    For y = FightAreaTopLeft.y To FightAreaBottomRight.y
        SpawnPos.x = FightAreaTopLeft.x - 1
        SpawnPos.y = y
        Call TirarItemAlPiso(SpawnPos, MiObj, 0)
        SpawnPos.x = FightAreaBottomRight.x + 1
        Call TirarItemAlPiso(SpawnPos, MiObj, 0)
        Call BlockAndInform(MapNumber, FightAreaTopLeft.x - 1, y, 1)
        Call BlockAndInform(MapNumber, FightAreaBottomRight.x + 1, y, 1)
    Next y
    MiObj.ObjIndex = RingLimiterTLItemID
    SpawnPos.x = FightAreaTopLeft.x - 1
    SpawnPos.y = FightAreaTopLeft.y - 1
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
    MiObj.ObjIndex = RingLimiterTRItemID
    SpawnPos.x = FightAreaBottomRight.x + 1
    SpawnPos.y = FightAreaTopLeft.y - 1
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
    MiObj.ObjIndex = RingLimiterBLItemID
    SpawnPos.x = FightAreaTopLeft.x - 1
    SpawnPos.y = FightAreaBottomRight.y + 1
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
    MiObj.ObjIndex = RingLimiterBRItemID
    SpawnPos.x = FightAreaBottomRight.x + 1
    SpawnPos.y = FightAreaBottomRight.y + 1
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
End Sub

Private Sub CloseScenario()
    On Error GoTo CloseScenario_Err:
    Call ModLobby.ReturnAllPlayers(LobbyList(LobbyIndex))
    Set LobbyList(LobbyIndex).Scenario = Nothing
    Call ModLobby.UpdateLobbyState(LobbyList(LobbyIndex), e_LobbyState.UnInitilized)
    Call ClearMap
    If MapNumber > 0 Then
        Call ReleaseInstance(MapNumber)
        MapNumber = 0
        Call ReleaseLobby(LobbyIndex)
    End If
    Exit Sub
CloseScenario_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.CloseScenario ", Erl)
End Sub

Private Sub ClearMap()
    On Error GoTo ClearMap_Err:
    CustomScenarios.RemoveCustomScenario (MapNumber)
    CustomScenarios.RemoveUpdateScenario (UpdateIndex)
    MapInfo(MapNumber).DropItems = MapDropItems
    MapInfo(MapNumber).Seguro = PrevMapState
    MapInfo(MapNumber).SafeFightMap = SafeFigthState
    MapInfo(MapNumber).FriendlyFire = FriendlyFireState
    Dim i As Integer
    If FightAreaTopLeft.x < FightAreaBottomRight.x Then
        For i = FightAreaTopLeft.x - 1 To FightAreaBottomRight.x + 1
            Call BlockAndInform(MapNumber, i, FightAreaTopLeft.y - 1, 0)
            Call BlockAndInform(MapNumber, i, FightAreaBottomRight.y + 1, 0)
        Next i
        For i = FightAreaTopLeft.y To FightAreaBottomRight.y
            Call BlockAndInform(MapNumber, FightAreaTopLeft.x - 1, i, 0)
            Call BlockAndInform(MapNumber, FightAreaBottomRight.x + 1, i, 0)
        Next i
    End If
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

Public Sub IBaseScenario_SetLobbyIndex(ByVal value As Integer)
    LobbyIndex = value
    Call ConfigMapLobby
End Sub

Public Function IBaseScenario_GetScenarioName() As String
    IBaseScenario_GetScenarioName = "DeathMatch"
End Function

Public Sub IBaseScenario_BroadcastOpenScenario()
    'queda pendiente hacer un mensaje de broadcast de deathmatch
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgDeathMatchInstructions, "", e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Public Sub IBaseScenario_SendRules(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MsgDeathMathInstructions, "", e_FontTypeNames.FONTTYPE_GUILD))
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
End Function

Public Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MapNumber
    SpawnPos.x = RandomNumber(50 - MinArenaSize, 50 - MinArenaSize)
    SpawnPos.y = RandomNumber(50 - MinArenaSize, 50 - MinArenaSize)
    LobbyList(LobbyIndex).SummonCoordinates = SpawnPos
    Call Board.AddPlayer(UserIndex)
    Dim UserItemInfo As t_UserItemInfo
    Dim i            As Integer
    With UserList(UserIndex)
        For i = 0 To SpawnedUserList.RegisteredUserCount - 1
            If SpawnedUserList.UserList(i).UserId = .Id Then
                Exit For
            End If
        Next i
        If i >= SpawnedUserList.RegisteredUserCount Then
            For i = 1 To UBound(.invent.Object)
                If IsForbidden(.invent.Object(i).ObjIndex) Then
                    UserItemInfo.RemovedItems(UserItemInfo.RemovedItemCount).ItemId = .invent.Object(i).ObjIndex
                    UserItemInfo.RemovedItems(UserItemInfo.RemovedItemCount).ItemCount = .invent.Object(i).amount
                    UserItemInfo.RemovedItemCount = UserItemInfo.RemovedItemCount + 1
                    .invent.Object(i).ObjIndex = 0
                    .invent.Object(i).amount = 0
                    .invent.Object(i).Equipped = 0
                End If
            Next i
            If UserItemInfo.RemovedItemCount > 0 Then
                SpawnedUserList.UserList(SpawnedUserList.RegisteredUserCount) = UserItemInfo
                SpawnedUserList.UserList(SpawnedUserList.RegisteredUserCount).UserId = .Id
                SpawnedUserList.RegisteredUserCount = SpawnedUserList.RegisteredUserCount + 1
                Call UpdateUserInv(True, UserIndex, 0)
            End If
        End If
    End With
End Sub

Public Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef NewState As e_LobbyState)
    If oldState = NewState Then
        Exit Sub
    End If
    If NewState = InProgress Then
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
    Call Board.UpdatePlayerScore(deadIndex, -1)
    If killerIndex <> deadIndex Then
        If Board.UpdatePlayerScore(killerIndex, 3) >= TargetScore Then
            Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(1889, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1889=¡Tenemos un ganador!
            Call MatchCompleted
        End If
        If GlobalFrameTime - UserList(deadIndex).flags.LastAttackedByUserTime < AssistDamageValidTime Then
            If IsValidUserRef(UserList(deadIndex).flags.LastAttacker) And UserList(deadIndex).flags.LastAttacker.ArrayIndex <> killerIndex Then
                If Board.UpdatePlayerScore(UserList(deadIndex).flags.LastAttacker.ArrayIndex, 1) >= TargetScore Then
                    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(1889, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1889=¡Tenemos un ganador!
                    Call MatchCompleted
                End If
            End If
        End If
    End If
End Sub

Public Sub IBaseScenario_Update()
    Dim frametime As Long
    frametime = TicksElapsed(LastFrameTime)
    LastFrameTime = GetTickCount()
    If LobbyList(LobbyIndex).State = e_LobbyState.InProgress Then
        If CountdownTimer.Occurrences < 10 Then
            If UpdateTime(CountdownTimer, frametime) Then
                If CountdownTimer.Occurrences >= 10 Then
                    Call StartGame
                Else
                    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(1890, (10 - CountdownTimer.Occurrences), e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1890=La matanza iniciará en ¬1 minutos.
                End If
            End If
            Exit Sub
        End If
        Call RespawnPlayers
        Call RespawnItems(frametime)
        ElapsedTime = ElapsedTime + frametime
        If ElapsedTime > MaxTime Then
            Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(1881, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1881=¡Se acabó el tiempo!
            Call MatchCompleted
        Else
            If UpdateTime(BroadCastTimer, frametime) Then
                Call BroadcastPendingTime(frametime)
            End If
            If UpdateTime(ScoreBroadcastTimer, frametime) Then
                Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, LobbyList(LobbyIndex), False)
            End If
        End If
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.Completed Then
        If CountdownTimer.Occurrences < 10 Then
            If UpdateTime(CountdownTimer, frametime) Then
                If CountdownTimer.Occurrences >= 10 Then
                    Call CloseScenario
                Else
                    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(1882, 10 - CountdownTimer.Occurrences, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1882=El mapa se cerrará en ¬1
                End If
            End If
            Exit Sub
        End If
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.AcceptingPlayers Then
        Call UpdateWaitingForPlayers(frametime, LobbyList(LobbyIndex))
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.Closed Then
        Call CloseScenario
    End If
End Sub

Public Sub IBaseScenario_OnUserDisconnected(ByVal UserIndex As Integer)
    Call Board.DisconnectUser(UserIndex, UserList(UserIndex).Id)
    Call RemoveForbiddenItems(UserIndex)
End Sub

Public Sub IBaseScenario_OnUserReconnect(ByVal UserIndex As Integer)
    Call Board.RestoreDisconnectUser(UserIndex, UserList(UserIndex).Id)
End Sub

Private Sub StartGame()
    Call SetTimer(BroadCastTimer, 60000)
    Call SetTimer(ScoreBroadcastTimer, 20000)
    LastFrameTime = GetTickCount()
    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(1893, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1893=¡Que inicie la masacre!
    ElapsedTime = 0
    Call LimitFightArea
    Dim i As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        Dim SpawnPos As t_WorldPos
        SpawnPos.Map = MapNumber
        SpawnPos.x = RandomNumber(FightAreaTopLeft.x, FightAreaBottomRight.x)
        SpawnPos.y = RandomNumber(FightAreaTopLeft.y, FightAreaBottomRight.y)
        If LobbyList(LobbyIndex).Players(i).IsSummoned Then
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                Call WarpToLegalPos(LobbyList(LobbyIndex).Players(i).User.ArrayIndex, MapNumber, SpawnPos.x, SpawnPos.y, True, True)
            End If
        End If
    Next i
    Dim j             As Integer
    Dim InitialStacks As Integer
    For i = 0 To UBound(ItemsToSpawn) - 1
        InitialStacks = ItemsToSpawn(i).InitialStacks * LobbyList(LobbyIndex).RegisteredPlayers
        For j = 0 To InitialStacks
            Call SpawnItem(ItemsToSpawn(i).ItemId, ItemsToSpawn(i).ItemCount)
        Next j
    Next i
    Call ConfigMapEvent
End Sub

Private Sub Respawn(ByVal player As Integer)
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MapNumber
    SpawnPos.x = RandomNumber(FightAreaTopLeft.x, FightAreaBottomRight.x)
    SpawnPos.y = RandomNumber(FightAreaTopLeft.y, FightAreaBottomRight.y)
    Call WarpToLegalPos(player, MapNumber, SpawnPos.x, SpawnPos.y, True, True)
    Call RevivirUsuario(player)
    UserList(player).Counters.timeFx = 3
    Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessageParticleFX(UserList(player).Char.charindex, e_ParticleEffects.Corazones, 100, False, , UserList(player).pos.x, _
            UserList(player).pos.y))
    Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessagePlayWave("104", UserList(player).pos.x, UserList(player).pos.y))
End Sub

Private Sub MatchCompleted()
    MapInfo(MapNumber).Seguro = 1
    Call SetTimer(CountdownTimer, 1000)
    Dim i As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            Call RemoveForbiddenItems(LobbyList(LobbyIndex).Players(i).User.ArrayIndex)
        End If
    Next i
    LobbyList(LobbyIndex).State = Completed
    ElapsedTime = MaxTime
    If Not LobbyList(LobbyIndex).Canceled Then
        Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, LobbyList(LobbyIndex), True)
        Call GiveRewards
    End If
    Call CustomScenarios.ClearMap(MapNumber)
    MapInfo(MapNumber).Seguro = 1
    MapInfo(MapNumber).SafeFightMap = False
End Sub

Private Sub BroadcastPendingTime(ByVal frametime As Long)
    Dim pendingTime As Long
    'we add frame time here again so we round up
    pendingTime = (MaxTime - ElapsedTime + frametime) / 1000
    If pendingTime > MaxTime / 2000 Then
        Exit Sub
    ElseIf pendingTime >= 60 Then
        pendingTime = pendingTime / 60
        Call SendData(SendTarget.toMap, MapNumber, PrepareMessageLocaleMsg(1894, CStr(pendingTime), e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1894=Quedan ¬1 minutos.
    ElseIf pendingTime <= 10 Then
        Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg(pendingTime, e_FontTypeNames.FONTTYPE_GUILD))
    End If
    If pendingTime <= 1 Then
        Call SetTimer(BroadCastTimer, 1000)
    End If
End Sub

Private Sub RespawnPlayers()
    Dim key         As Variant
    Dim Keys        As Variant
    Dim currentTime As Long
    Dim RespawnInfo As clsRespawnInfo
    Dim i           As Integer
    currentTime = GetTickCount()
    Keys = PlayerRespawn.Keys
    For Each key In Keys
        Set RespawnInfo = PlayerRespawn.Item(key)
        If currentTime - RespawnInfo.RespawnTime > PlayerRespawnTime Then
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

Private Sub RespawnItems(ByVal frametime As Long)
    Dim i, j As Integer
    Dim SpawnCount    As Integer
    Dim StackToCreate As Integer
    StackToCreate = (ItemsToSpawn(i).StackCount * LobbyList(LobbyIndex).RegisteredPlayers) - 1
    For i = 0 To UBound(ItemsToSpawn) - 1
        If UpdateTime(ItemsToSpawn(i).timer, frametime) Then
            For j = 0 To StackToCreate
                SpawnCount = ItemsToSpawn(i).ItemCount
                Call SpawnItem(ItemsToSpawn(i).ItemId, SpawnCount)
            Next j
        End If
    Next i
End Sub

Private Sub SpawnItem(ByVal ItemId As Integer, ByVal StackSize As Integer)
    Dim MiObj As t_Obj
    MiObj.amount = StackSize
    MiObj.ObjIndex = ItemId
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MapNumber
    SpawnPos.x = RandomNumber(FightAreaTopLeft.x, FightAreaBottomRight.x)
    SpawnPos.y = RandomNumber(FightAreaTopLeft.y, FightAreaBottomRight.y)
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
End Sub

Private Sub ConfigMapEvent()
    MapInfo(MapNumber).Seguro = 0
    MapInfo(MapNumber).SinInviOcul = True
    MapInfo(MapNumber).SafeFightMap = True
    MapInfo(MapNumber).FriendlyFire = False
End Sub

Private Sub ConfigMapLobby()
    PrevMapState = MapInfo(MapNumber).Seguro
    MapDropItems = MapInfo(MapNumber).DropItems
    SafeFigthState = MapInfo(MapNumber).SafeFightMap
    FriendlyFireState = MapInfo(MapNumber).FriendlyFire
    MapInfo(MapNumber).Seguro = 1
    MapInfo(MapNumber).SinInviOcul = True
    MapInfo(MapNumber).DropItems = False
End Sub

Public Sub IBaseScenario_GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
End Sub

Public Sub IBaseScenario_NpcDie(ByVal NpcIndex)
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
    End With
    Call PlayerRespawn.Add(UserIndex, RespawnInfo)
End Sub

Sub GiveRewards()
    On Error GoTo GiveRewards_Err:
    Dim Teams()         As String
    Dim Rewards(1 To 3) As Long
    Dim Reward          As Long
    Dim Ranking()       As e_Rank
    Dim RewardStr       As String
    Dim i               As Integer
    Dim j               As Integer
    If (LobbyList(LobbyIndex).TeamSize = 0) Then
        LobbyList(LobbyIndex).TeamSize = 1 'cero division check
    End If
    If (LobbyList(LobbyIndex).RegisteredPlayers / LobbyList(LobbyIndex).TeamSize) = 2 Then
        Rewards(1) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.55
        Rewards(2) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.35
    Else
        Rewards(1) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.4
        Rewards(2) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.3
        Rewards(3) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.2
    End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgMatchComplete, Me.IBaseScenario_GetScenarioName(), e_FontTypeNames.FONTTYPE_GUILD))
    If LobbyList(LobbyIndex).TeamSize > 1 Then
        ReDim Teams(0 To LobbyList(LobbyIndex).RegisteredPlayers / LobbyList(LobbyIndex).TeamSize) As String
        Ranking = Board.GetTeamRanking(LobbyList(LobbyIndex))
        For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                Teams(LobbyList(LobbyIndex).Players(i).team) = Teams(LobbyList(LobbyIndex).Players(i).team) & UserList(LobbyList(LobbyIndex).Players(i).User.ArrayIndex).name & " "
            End If
        Next i
        For i = 1 To 3
            If LobbyList(LobbyIndex).AvailableInscriptionMoney > 0 Then
                Reward = Rewards(i) / LobbyList(LobbyIndex).TeamSize
                For j = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
                    If IsValidUserRef(LobbyList(LobbyIndex).Players(j).User) Then
                        If LobbyList(LobbyIndex).Players(j).team = Ranking(i).PlayerIndex Then
                            Call GiveGoldToPlayer(LobbyList(LobbyIndex), j, Reward)
                        End If
                    End If
                Next j
                RewardStr = ", ganaron " & PonerPuntos(Reward) & " Modenas de oro cada uno."
            End If
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1744, i & "¬" & Teams(Ranking(i).PlayerIndex) & "¬" & Ranking(i).Score & "¬" & RewardStr, _
                    e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1744=¬1) ¬2: ¬3 puntos¬4
        Next i
    Else
        Ranking = Board.GetRanking(3)
        For i = 0 To UBound(Ranking) - 1
            If LobbyList(LobbyIndex).AvailableInscriptionMoney > 0 Then
                For j = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
                    If IsValidUserRef(LobbyList(LobbyIndex).Players(j).User) Then
                        If LobbyList(LobbyIndex).Players(j).User.ArrayIndex = Ranking(i).PlayerIndex Then
                            Call GiveGoldToPlayer(LobbyList(LobbyIndex), j, Rewards(i + 1))
                            Exit For
                        End If
                    End If
                Next j
                RewardStr = ", gano " & PonerPuntos(Rewards(i + 1)) & " Modenas de oro."
            End If
            If Ranking(i).PlayerIndex > 0 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1824, (i + 1) & "¬" & UserList(Ranking(i).PlayerIndex).name & "¬" & Ranking(i).Score & "¬" & _
                        RewardStr, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1824=¬1) ¬2: ¬3 puntos¬4
            End If
        Next i
    End If
    Exit Sub
GiveRewards_Err:
    Call TraceError(Err.Number, Err.Description, "DeathMatch.GiveRewards ", Erl)
End Sub
