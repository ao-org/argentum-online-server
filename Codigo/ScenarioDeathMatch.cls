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
    itemId As Integer
    itemCount As Integer
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

Dim SpawnedUserList As t_UserList
Dim MapNumber As Integer
Dim TargetScore As Integer
Dim MaxTime As Long
Dim PlayerRespawnTime As Long
Dim CountdownTimer As t_Timer
Dim BroadCastTimer As t_Timer
Dim ScoreBroadcastTimer As t_Timer
Dim ForbbidenItemId() As Integer
Dim UpdateIndex As Integer
Dim Board As New ScoreBoard
Dim LastFrameTime As Long
Dim ElapsedTime As Long
Dim PlayerRespawn As New Dictionary
Dim ItemsToSpawn() As t_SpawnItems
Dim MapDropItems As Boolean
Dim PrevMapState As Integer
Dim SafeFigthState As Boolean
Dim FriendlyFireState As Boolean
Dim FightAreaTopLeft As t_Position
Dim FightAreaBottomRight As t_Position
Dim PlayerAreaIncreaseModifier As Single
Dim RingLimiterHItemID As Long
Dim RingLimiterVItemID As Long
Dim RingLimiterTRItemID As Long
Dim RingLimiterTLItemID As Long
Dim RingLimiterBRItemID As Long
Dim RingLimiterBLItemID As Long
Dim MinArenaSize As Integer
Dim LobbyIndex As Integer

Private Sub Class_Initialize()
On Error GoTo DeathMatch_Initialize_Err:
    Debug.Assert FileExist(DatPath & configPath)
    Dim ScenarioIni As clsIniManager
    Dim OriginalMap As Integer
104 Set ScenarioIni = New clsIniManager
    Call ScenarioIni.Initialize(DatPath & configPath)
112 OriginalMap = val(ScenarioIni.GetValue("INIT", "MapNumber"))
    MapNumber = InstanceManager.GetNextAvailableInstance
    Call InstanceManager.CloneMap(OriginalMap, MapNumber)
114 MaxTime = val(ScenarioIni.GetValue("INIT", "MaxTime"))
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
116 TargetScore = val(ScenarioIni.GetValue("INIT", "TargetScore"))
118 Call LoadForbidenItems(ScenarioIni)
120 UpdateIndex = AddUpdateScenario(Me)
122 Call CustomScenarios.SetCustomScenario(MapNumber, Me)
124 Call LoadItemsToSpawn(ScenarioIni)
    Exit Sub
DeathMatch_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "DeathMatch.Initialize ", Erl)
End Sub

Private Sub LoadItemsToSpawn(ScenearioIni As clsIniManager)
    Dim itemCount As Integer
    Dim i As Integer
    Dim spawInterval As Long
    itemCount = val(ScenearioIni.GetValue("INIT", "SpawnItems"))
    ReDim ItemsToSpawn(itemCount) As t_SpawnItems
    For i = 0 To itemCount - 1
        ItemsToSpawn(i).itemId = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "ItemId"))
        ItemsToSpawn(i).itemCount = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "ItemCount"))
        ItemsToSpawn(i).StackCount = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "StackCount")) / 10
        ItemsToSpawn(i).InitialStacks = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "InitialStacks")) / 10
        spawInterval = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "SpawnInterval"))
        Call SetTimer(ItemsToSpawn(i).timer, spawInterval * 1000)
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

Private Sub UpdateFightAreaForPlayerCount()
    Dim size As Integer
    Size = MinArenaSize + LobbyList(LobbyIndex).RegisteredPlayers * PlayerAreaIncreaseModifier
    FightAreaTopLeft.X = 50 - size
    FightAreaTopLeft.y = 50 - size
    FightAreaBottomRight.X = 50 + size
    FightAreaBottomRight.y = 50 + size
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
            If IsForbidden(.invent.Object(i).objIndex) Then
                Call QuitarUserInvItem(UserIndex, i, .invent.Object(i).amount)
            End If
        Next i
        If .flags.ModificoInventario Then
            Call UpdateUserInv(True, UserIndex, 0)
        End If
        
        For i = 0 To SpawnedUserList.RegisteredUserCount - 1
            If SpawnedUserList.UserList(i).UserId = .id Then
                Exit For
            End If
        Next i
        If i < SpawnedUserList.RegisteredUserCount Then
            For j = 0 To SpawnedUserList.UserList(i).RemovedItemCount - 1
                If SpawnedUserList.UserList(i).RemovedItems(j).ItemId > 0 Then
                    Call AddItemToUser(UserIndex, _
                                       SpawnedUserList.UserList(i).RemovedItems(j).ItemId, _
                                       SpawnedUserList.UserList(i).RemovedItems(j).ItemCount)
                End If
            Next j
            SpawnedUserList.UserList(i) = SpawnedUserList.UserList(SpawnedUserList.RegisteredUserCount - 1)
            SpawnedUserList.RegisteredUserCount = SpawnedUserList.RegisteredUserCount - 1
            Call UpdateUserInv(True, UserIndex, 0)
        End If
    End With
End Sub

Private Sub LimitFightArea()
    Dim X As Integer
    Dim y As Integer
    Dim MiObj As t_Obj
    MiObj.amount = 1
    MiObj.objIndex = RingLimiterHItemID
    Dim SpawnPos As t_WorldPos
    SpawnPos.map = MapNumber
    Call UpdateFightAreaForPlayerCount
    For X = FightAreaTopLeft.X To FightAreaBottomRight.X
       SpawnPos.X = X
       SpawnPos.y = FightAreaTopLeft.y - 1
       Call TirarItemAlPiso(SpawnPos, MiObj, 0)
       SpawnPos.y = FightAreaBottomRight.y + 1
       Call TirarItemAlPiso(SpawnPos, MiObj, 0)
       Call BlockAndInform(MapNumber, X, FightAreaTopLeft.y - 1, 1)
       Call BlockAndInform(MapNumber, X, FightAreaBottomRight.y + 1, 1)
    Next X
    MiObj.objIndex = RingLimiterVItemID
    For y = FightAreaTopLeft.y To FightAreaBottomRight.y
       SpawnPos.X = FightAreaTopLeft.X - 1
       SpawnPos.y = y
       Call TirarItemAlPiso(SpawnPos, MiObj, 0)
       SpawnPos.X = FightAreaBottomRight.X + 1
       Call TirarItemAlPiso(SpawnPos, MiObj, 0)
       Call BlockAndInform(MapNumber, FightAreaTopLeft.X - 1, y, 1)
       Call BlockAndInform(MapNumber, FightAreaBottomRight.X + 1, y, 1)
    Next y
    MiObj.objIndex = RingLimiterTLItemID
    SpawnPos.X = FightAreaTopLeft.X - 1
    SpawnPos.y = FightAreaTopLeft.y - 1
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
    
    MiObj.objIndex = RingLimiterTRItemID
    SpawnPos.X = FightAreaBottomRight.X + 1
    SpawnPos.y = FightAreaTopLeft.y - 1
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
    
    MiObj.objIndex = RingLimiterBLItemID
    SpawnPos.X = FightAreaTopLeft.X - 1
    SpawnPos.y = FightAreaBottomRight.y + 1
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
       
    MiObj.objIndex = RingLimiterBRItemID
    SpawnPos.X = FightAreaBottomRight.X + 1
    SpawnPos.y = FightAreaBottomRight.y + 1
    Call TirarItemAlPiso(SpawnPos, MiObj, 0)
End Sub

Private Sub CloseScenario()
On Error GoTo CloseScenario_Err:
100    Call ModLobby.ReturnAllPlayers(LobbyList(LobbyIndex))
102    Set LobbyList(LobbyIndex).Scenario = Nothing
104    Call ModLobby.UpdateLobbyState(LobbyList(LobbyIndex), e_LobbyState.UnInitilized)
106    Call ClearMap
108    If MapNumber > 0 Then
110         Call ReleaseInstance(MapNumber)
112         MapNumber = 0
114         Call ReleaseLobby(LobbyIndex)
       End If
    Exit Sub
CloseScenario_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.CloseScenario ", Erl)
End Sub

Private Sub ClearMap()
    On Error GoTo ClearMap_Err:
100     CustomScenarios.RemoveCustomScenario (MapNumber)
102     CustomScenarios.RemoveUpdateScenario (UpdateIndex)
104     MapInfo(MapNumber).DropItems = MapDropItems
106     MapInfo(MapNumber).Seguro = PrevMapState
108     MapInfo(MapNumber).SafeFightMap = SafeFigthState
110     MapInfo(MapNumber).FriendlyFire = FriendlyFireState
        Dim i As Integer
112     If FightAreaTopLeft.x < FightAreaBottomRight.x Then
114         For i = FightAreaTopLeft.x - 1 To FightAreaBottomRight.x + 1
116             Call BlockAndInform(MapNumber, i, FightAreaTopLeft.y - 1, 0)
118             Call BlockAndInform(MapNumber, i, FightAreaBottomRight.y + 1, 0)
120         Next i
122         For i = FightAreaTopLeft.y To FightAreaBottomRight.y
124             Call BlockAndInform(MapNumber, FightAreaTopLeft.x - 1, i, 0)
126             Call BlockAndInform(MapNumber, FightAreaBottomRight.x + 1, i, 0)
128         Next i
        End If
    Exit Sub
ClearMap_Err:
130 Call TraceError(Err.Number, Err.Description, "ScenarioHunt.ClearMap", Erl)
End Sub

Private Function IsForbidden(ByVal itemId As Integer) As Boolean
    Dim i As Integer
    If itemId = 0 Then
        IsForbidden = False
        Exit Function
    End If
    For i = 0 To UBound(ForbbidenItemId) - 1
        If ForbbidenItemId(i) = itemId Then
            IsForbidden = True
            Exit Function
        End If
    Next i
    IsForbidden = False
End Function

Public Sub IBaseScenario_SetLobbyIndex(ByVal Value As Integer)
    LobbyIndex = Value
    Call ConfigMapLobby
End Sub

Public Function IBaseScenario_GetScenarioName() As String
    IBaseScenario_GetScenarioName = "DeathMatch"
End Function

Public Sub IBaseScenario_BroadcastOpenScenario(ByVal CreatorIndex As Integer)
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
    Dim spawnPos As t_WorldPos
    spawnPos.map = MapNumber
    spawnPos.X = RandomNumber(50 - MinArenaSize, 50 - MinArenaSize)
    spawnPos.y = RandomNumber(50 - MinArenaSize, 50 - MinArenaSize)
    LobbyList(LobbyIndex).SummonCoordinates = SpawnPos
    Call Board.AddPlayer(UserIndex)
    Dim UserItemInfo As t_UserItemInfo
    Dim i As Integer
    With UserList(UserIndex)
        For i = 0 To SpawnedUserList.RegisteredUserCount - 1
            If SpawnedUserList.UserList(i).UserId = .id Then
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
                SpawnedUserList.UserList(SpawnedUserList.RegisteredUserCount).UserId = .id
                SpawnedUserList.RegisteredUserCount = SpawnedUserList.RegisteredUserCount + 1
                Call UpdateUserInv(True, UserIndex, 0)
            End If
        End If
    End With
    
End Sub

Public Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef newState As e_LobbyState)
    If oldState = newState Then
        Exit Sub
    End If
    If newState = InProgress Then
        Call SetTimer(CountdownTimer, 1000)
    End If
    If newState = Closed Then
        Call MatchCompleted
    End If
End Sub

Public Sub IBaseScenario_PlayerKillNpc(ByVal npcIndex As Integer, ByVal UserIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
End Sub

Public Sub IBaseScenario_NPcKillPlayer(ByVal npcIndex As Integer, ByVal UserIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
End Sub

Public Sub IBaseScenario_PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    Call Board.UpdatePlayerScore(deadIndex, -1)
    If killerIndex <> deadIndex Then
        If Board.UpdatePlayerScore(killerIndex, 3) >= TargetScore Then
            Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg("Tenemos un ganador!", e_FontTypeNames.FONTTYPE_GUILD))
            Call MatchCompleted
        End If
        If GlobalFrameTime - UserList(deadIndex).flags.LastAttackedByUserTime < AssistDamageValidTime Then
            If IsValidUserRef(UserList(deadIndex).flags.LastAttacker) And UserList(deadIndex).flags.LastAttacker.ArrayIndex <> killerIndex Then
                If Board.UpdatePlayerScore(UserList(deadIndex).flags.LastAttacker.ArrayIndex, 1) >= TargetScore Then
                    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg("Tenemos un ganador!", e_FontTypeNames.FONTTYPE_GUILD))
                    Call MatchCompleted
                End If
            End If
        End If
    End If
End Sub

Public Sub IBaseScenario_Update()
    Dim frametime As Long
    frametime = GetTickCount() - LastFrameTime
    LastFrameTime = GetTickCount()
    If LobbyList(LobbyIndex).State = e_LobbyState.InProgress Then
        If CountdownTimer.Occurrences < 10 Then
            If UpdateTime(CountdownTimer, frametime) Then
                If CountdownTimer.Occurrences >= 10 Then
                    Call StartGame
                Else
                    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg("La matanza iniciara en " & 10 - CountdownTimer.Occurrences, e_FontTypeNames.FONTTYPE_GUILD))
                End If
            End If
            Exit Sub
        End If
        Call RespawnPlayers
        Call RespawnItems(frametime)
        ElapsedTime = ElapsedTime + frametime
        If ElapsedTime > MaxTime Then
            Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg("Se acabo el tiempo!", e_FontTypeNames.FONTTYPE_GUILD))
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
                    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg("El mapa se cerrara en " & 10 - CountdownTimer.Occurrences, e_FontTypeNames.FONTTYPE_GUILD))
                End If
            End If
            Exit Sub
        End If
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.AcceptingPlayers Then
        Call UpdateWaitingForPlayers(FrameTime, LobbyList(LobbyIndex))
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.Closed Then
        Call CloseScenario
    End If
End Sub

Public Sub IBaseScenario_OnUserDisconnected(ByVal userIndex As Integer)
    Call Board.DisconnectUser(userIndex, UserList(userIndex).ID)
    Call RemoveForbiddenItems(UserIndex)
End Sub

Public Sub IBaseScenario_OnUserReconnect(ByVal userIndex As Integer)
    Call Board.RestoreDisconnectUser(userIndex, UserList(userIndex).ID)
End Sub

Private Sub StartGame()
    Call SetTimer(BroadCastTimer, 60000)
    Call SetTimer(ScoreBroadcastTimer, 20000)
    LastFrameTime = GetTickCount()
    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg("Que inicie la masacre!!!", e_FontTypeNames.FONTTYPE_GUILD))
    ElapsedTime = 0
    Call LimitFightArea
    Dim i As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        Dim spawnPos As t_WorldPos
        spawnPos.map = MapNumber
        SpawnPos.X = RandomNumber(FightAreaTopLeft.X, FightAreaBottomRight.X)
        SpawnPos.y = RandomNumber(FightAreaTopLeft.y, FightAreaBottomRight.y)
        If LobbyList(LobbyIndex).Players(i).IsSummoned Then
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).user) Then
                Call WarpToLegalPos(LobbyList(LobbyIndex).Players(i).user.ArrayIndex, MapNumber, SpawnPos.x, SpawnPos.y, True, True)
            End If
        End If
    Next i
    Dim j As Integer
    Dim InitialStacks As Integer
    For i = 0 To UBound(ItemsToSpawn) - 1
        InitialStacks = ItemsToSpawn(i).InitialStacks * LobbyList(LobbyIndex).RegisteredPlayers
        For j = 0 To InitialStacks
            Call SpawnItem(ItemsToSpawn(i).itemId, ItemsToSpawn(i).itemCount)
        Next j
    Next i
    Call ConfigMapEvent
End Sub

Private Sub Respawn(ByVal player As Integer)
    Dim spawnPos As t_WorldPos
    spawnPos.map = MapNumber
    SpawnPos.X = RandomNumber(FightAreaTopLeft.X, FightAreaBottomRight.X)
    SpawnPos.y = RandomNumber(FightAreaTopLeft.y, FightAreaBottomRight.y)
    Call WarpToLegalPos(player, MapNumber, spawnPos.x, spawnPos.y, True, True)
    Call RevivirUsuario(player)
    UserList(player).Counters.timeFx = 3
114 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessageParticleFX(UserList(player).Char.charindex, e_ParticulasIndex.Curar, 100, False, , UserList(player).Pos.x, UserList(player).Pos.y))
116 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessagePlayWave("104", UserList(player).Pos.x, UserList(player).Pos.y))
End Sub

Private Sub MatchCompleted()
    MapInfo(MapNumber).Seguro = 1
    Call SetTimer(CountdownTimer, 1000)
    Dim i As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).user) Then
            Call RemoveForbiddenItems(LobbyList(LobbyIndex).Players(i).user.ArrayIndex)
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
        Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg("quedan " & pendingTime & " minutos.", e_FontTypeNames.FONTTYPE_GUILD))
    ElseIf pendingTime <= 10 Then
        Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg(pendingTime, e_FontTypeNames.FONTTYPE_GUILD))
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
    Dim SpawnCount As Integer
    Dim StackToCreate As Integer
    StackToCreate = (ItemsToSpawn(i).StackCount * LobbyList(LobbyIndex).RegisteredPlayers) - 1
    For i = 0 To UBound(ItemsToSpawn) - 1
        If UpdateTime(ItemsToSpawn(i).timer, frametime) Then
            For j = 0 To StackToCreate
                SpawnCount = ItemsToSpawn(i).itemCount
                Call SpawnItem(ItemsToSpawn(i).itemId, SpawnCount)
            Next j
        End If
    Next i
End Sub

Private Sub SpawnItem(ByVal itemId As Integer, ByVal stackSize As Integer)
    Dim MiObj As t_Obj
    MiObj.amount = stackSize
    MiObj.objIndex = itemId
    Dim spawnPos As t_WorldPos
    spawnPos.map = MapNumber
    SpawnPos.X = RandomNumber(FightAreaTopLeft.X, FightAreaBottomRight.X)
    SpawnPos.y = RandomNumber(FightAreaTopLeft.y, FightAreaBottomRight.y)
    Call TirarItemAlPiso(spawnPos, MiObj, 0)
End Sub

Private Sub ConfigMapEvent()
    MapInfo(MapNumber).Seguro = 0
    MapInfo(MapNumber).SinInviOcul = True
    MapInfo(mapNumber).SafeFightMap = True
    MapInfo(MapNumber).FriendlyFire = False
End Sub

Private Sub ConfigMapLobby()
    PrevMapState = MapInfo(MapNumber).Seguro
    MapDropItems = MapInfo(MapNumber).DropItems
    SafeFigthState = MapInfo(mapNumber).SafeFightMap
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
        Dim Teams() As String
        Dim Rewards(1 To 3) As Long
        Dim Reward As Long
        Dim Ranking() As e_Rank
        Dim RewardStr As String
        Dim i As Integer
        Dim j As Integer
100     If (LobbyList(LobbyIndex).RegisteredPlayers / LobbyList(LobbyIndex).TeamSize) = 2 Then
102         Rewards(1) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.55
104         Rewards(2) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.35
        Else
106         Rewards(1) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.4
108         Rewards(2) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.3
110         Rewards(3) = LobbyList(LobbyIndex).AvailableInscriptionMoney * 0.2
        End If
112     Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgMatchComplete, Me.IBaseScenario_GetScenarioName(), e_FontTypeNames.FONTTYPE_GUILD))
114     If LobbyList(LobbyIndex).TeamSize > 1 Then
116         ReDim Teams(0 To LobbyList(LobbyIndex).RegisteredPlayers / LobbyList(LobbyIndex).TeamSize) As String
118         Ranking = Board.GetTeamRanking(LobbyList(LobbyIndex))
120         For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
122             If IsValidUserRef(LobbyList(LobbyIndex).Players(i).user) Then
124                 Teams(LobbyList(LobbyIndex).Players(i).team) = Teams(LobbyList(LobbyIndex).Players(i).team) & UserList(LobbyList(LobbyIndex).Players(i).user.ArrayIndex).name & " "
                End If
126         Next i
128         For i = 1 To 3
130             If LobbyList(LobbyIndex).AvailableInscriptionMoney > 0 Then
132                 Reward = Rewards(i) / LobbyList(LobbyIndex).TeamSize
134                 For j = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
136                     If IsValidUserRef(LobbyList(LobbyIndex).Players(j).user) Then
138                         If LobbyList(LobbyIndex).Players(j).team = Ranking(i).PlayerIndex Then
140                             Call GiveGoldToPlayer(LobbyList(LobbyIndex), j, Reward)
                            End If
                        End If
142                 Next j
144                 RewardStr = ", ganaron " & PonerPuntos(Reward) & " Modenas de oro cada uno."
                End If
146             Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1744, i & "¬" & Teams(Ranking(i).PlayerIndex) & "¬" & Ranking(i).Score & "¬" & RewardStr, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1744=¬1) ¬2: ¬3 puntos¬4
148         Next i
        Else
150         Ranking = Board.GetRanking(3)
152         For i = 0 To UBound(Ranking) - 1
154             If LobbyList(LobbyIndex).AvailableInscriptionMoney > 0 Then
156                 For j = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
158                     If IsValidUserRef(LobbyList(LobbyIndex).Players(j).user) Then
160                         If LobbyList(LobbyIndex).Players(j).user.ArrayIndex = Ranking(i).PlayerIndex Then
162                             Call GiveGoldToPlayer(LobbyList(LobbyIndex), j, Rewards(i + 1))
                                Exit For
                            End If
                        End If
164                 Next j
166                 RewardStr = ", gano " & PonerPuntos(Rewards(i + 1)) & " Modenas de oro."
                End If
168             If Ranking(i).PlayerIndex > 0 Then Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg((i + 1) & ") " & UserList(Ranking(i).PlayerIndex).name & ": " & Ranking(i).Score & " puntos" & RewardStr, e_FontTypeNames.FONTTYPE_GUILD))
170         Next i
        End If
        Exit Sub
GiveRewards_Err:
    Call TraceError(Err.Number, Err.Description, "DeathMatch.GiveRewards ", Erl)
End Sub
