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
Option Explicit

Const configPath = "../init/Scenarios/DeathMatch.ini"

Implements IBaseScenario

Private Type t_SpawnItems
    itemId As Integer
    itemCount As Integer
    StackCount As Single
    InitialStacks As Single
    timer As t_Timer
End Type

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

Private Sub Class_Initialize()
On Error GoTo DeathMatch_Initialize_Err:
    Debug.Assert FileExist(DatPath & configPath)
    Dim ScenarioIni As clsIniManager
104 Set ScenarioIni = New clsIniManager
    Call ScenarioIni.Initialize(DatPath & configPath)
112 MapNumber = val(ScenarioIni.GetValue("INIT", "MapNumber"))
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
    Call ConfigMapLobby
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
    size = MinArenaSize + GenericGlobalLobby.RegisteredPlayers * PlayerAreaIncreaseModifier
    FightAreaTopLeft.X = 50 - size
    FightAreaTopLeft.y = 50 - size
    FightAreaBottomRight.X = 50 + size
    FightAreaBottomRight.y = 50 + size
End Sub

Private Sub RemoveForbiddenItems(ByVal UserIndex As Integer)
    Dim i As Integer
    With UserList(UserIndex)
        For i = 1 To UBound(.invent.Object)
            If IsForbidden(.invent.Object(i).objIndex) Then
                Call QuitarUserInvItem(UserIndex, i, .invent.Object(i).amount)
            End If
        Next i
        If .flags.ModificoInventario Then
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
100    Call ModLobby.ReturnAllPlayers(GenericGlobalLobby)
102    Set GenericGlobalLobby.scenario = Nothing
104    Call ModLobby.UpdateLobbyState(GenericGlobalLobby, e_LobbyState.UnInitilized)
118    Call ClearMap
    Exit Sub
CloseScenario_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.CloseScenario ", Erl)
End Sub

Private Sub ClearMap()
On Error GoTo ClearMap_Err:
106    CustomScenarios.RemoveCustomScenario (MapNumber)
108    CustomScenarios.RemoveUpdateScenario (UpdateIndex)
110    MapInfo(MapNumber).DropItems = MapDropItems
112    MapInfo(MapNumber).Seguro = PrevMapState
114    MapInfo(MapNumber).SafeFightMap = SafeFigthState
115    MapInfo(MapNumber).FriendlyFire = FriendlyFireState
116    Dim i As Integer
118    For i = FightAreaTopLeft.X - 1 To FightAreaBottomRight.X + 1
120        Call BlockAndInform(MapNumber, i, FightAreaTopLeft.y - 1, 0)
122        Call BlockAndInform(MapNumber, i, FightAreaBottomRight.y + 1, 0)
124    Next i
126    For i = FightAreaTopLeft.y To FightAreaBottomRight.y
128        Call BlockAndInform(MapNumber, FightAreaTopLeft.X - 1, i, 0)
130        Call BlockAndInform(MapNumber, FightAreaBottomRight.X + 1, i, 0)
132    Next i
    Exit Sub
ClearMap_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.ClearMap", Erl)
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

Public Sub IBaseScenario_Reset()
End Sub

Public Function IBaseScenario_RequiresSpawn() As Boolean
    IBaseScenario_RequiresSpawn = False
End Function

Public Function IBaseScenario_ValidateUser(ByVal UserIndex As Integer) As Integer
    Dim i As Integer
    With UserList(UserIndex)
        For i = 1 To UBound(.Invent.Object)
            If IsForbidden(.Invent.Object(i).objIndex) Then
                IBaseScenario_ValidateUser = 404
                Exit Function
            End If
        Next i
        IBaseScenario_ValidateUser = 0
    End With
End Function

Public Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    Dim spawnPos As t_WorldPos
    spawnPos.map = MapNumber
    spawnPos.X = RandomNumber(50 - MinArenaSize, 50 - MinArenaSize)
    spawnPos.y = RandomNumber(50 - MinArenaSize, 50 - MinArenaSize)
    GenericGlobalLobby.SummonCoordinates = spawnPos
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
    If Board.UpdatePlayerScore(killerIndex, 3) >= TargetScore Then
        Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg("Tenemos un ganador!", e_FontTypeNames.FONTTYPE_GUILD))
        Call MatchCompleted
    End If
    Dim RespawnInfo As clsRespawnInfo
    Set RespawnInfo = New clsRespawnInfo
    RespawnInfo.RespawnTime = GetTickCount()
    With UserList(deadIndex)
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
    Call PlayerRespawn.Add(deadIndex, RespawnInfo)
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
                Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, GenericGlobalLobby, False)
            End If
        End If
    ElseIf GenericGlobalLobby.State = e_LobbyState.Completed Then
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
    ElseIf GenericGlobalLobby.State = e_LobbyState.Closed Then
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
    For i = 0 To GenericGlobalLobby.RegisteredPlayers - 1
        Dim spawnPos As t_WorldPos
        spawnPos.map = MapNumber
        SpawnPos.X = RandomNumber(FightAreaTopLeft.X, FightAreaBottomRight.X)
        SpawnPos.y = RandomNumber(FightAreaTopLeft.y, FightAreaBottomRight.y)
        If GenericGlobalLobby.Players(i).IsSummoned Then
            If IsValidUserRef(GenericGlobalLobby.Players(i).user) Then
                Call WarpToLegalPos(GenericGlobalLobby.Players(i).user.ArrayIndex, MapNumber, SpawnPos.X, SpawnPos.y, True, True)
            End If
        End If
    Next i
    Dim j As Integer
    Dim InitialStacks As Integer
    For i = 0 To UBound(ItemsToSpawn) - 1
        InitialStacks = ItemsToSpawn(i).InitialStacks * GenericGlobalLobby.RegisteredPlayers
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
    UserList(player).Counters.timeFx = 2
114 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessageParticleFX(UserList(player).Char.charindex, e_ParticulasIndex.Curar, 100, False, , UserList(player).Pos.x, UserList(player).Pos.y))
116 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessagePlayWave("104", UserList(player).Pos.x, UserList(player).Pos.y))
End Sub

Private Sub MatchCompleted()
    MapInfo(MapNumber).Seguro = 1
    Call SetTimer(CountdownTimer, 1000)
    Dim i As Integer
    For i = 0 To GenericGlobalLobby.RegisteredPlayers - 1
        If IsValidUserRef(GenericGlobalLobby.Players(i).user) Then
            Call RemoveForbiddenItems(GenericGlobalLobby.Players(i).user.ArrayIndex)
        End If
    Next i
    GenericGlobalLobby.State = Completed
    ElapsedTime = MaxTime
    Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, GenericGlobalLobby, True)
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
    StackToCreate = (ItemsToSpawn(i).StackCount * GenericGlobalLobby.RegisteredPlayers) - 1
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


