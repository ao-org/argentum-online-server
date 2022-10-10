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
    StackCount As Integer
    InitialStacks As Integer
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
        ItemsToSpawn(i).StackCount = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "StackCount"))
        ItemsToSpawn(i).InitialStacks = val(ScenearioIni.GetValue("SpawnItem" & i + 1, "InitialStacks"))
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

Private Sub CloseScenario()
On Error GoTo CloseScenario_Err:
100    Call ModLobby.ReturnAllPlayers(GenericGlobalLobby)
102    Set GenericGlobalLobby.scenario = Nothing
104    Call ModLobby.UpdateLobbyState(GenericGlobalLobby, e_LobbyState.UnInitilized)
106    CustomScenarios.RemoveCustomScenario (MapNumber)
108    CustomScenarios.RemoveUpdateScenario (UpdateIndex)
110    MapInfo(MapNumber).DropItems = MapDropItems
112    MapInfo(MapNumber).Seguro = PrevMapState
114    MapInfo(mapNumber).SafeFightMap = SafeFigthState
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
    spawnPos.X = RandomNumber(12, 85)
    spawnPos.y = RandomNumber(12, 85)
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
    Call PlayerRespawn.Add(deadIndex, GetTickCount())
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
                Call Board.PublishScoreboard(SendTarget.toMap, MapNumber)
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
    
    Dim i As Integer
    For i = 0 To GenericGlobalLobby.RegisteredPlayers - 1
        Dim spawnPos As t_WorldPos
        spawnPos.map = MapNumber
        spawnPos.x = RandomNumber(12, 85)
        spawnPos.y = RandomNumber(12, 85)
        If GenericGlobalLobby.Players(i).IsSummoned Then
            If IsValidUserRef(GenericGlobalLobby.Players(i).user) Then
                Call WarpToLegalPos(GenericGlobalLobby.Players(i).user.ArrayIndex, MapNumber, SpawnPos.X, SpawnPos.y, True, True)
            End If
        End If
    Next i
    Dim j As Integer
    For i = 0 To UBound(ItemsToSpawn) - 1
        For j = 0 To ItemsToSpawn(i).InitialStacks - 1
            Call SpawnItem(ItemsToSpawn(i).itemId, ItemsToSpawn(i).itemCount)
        Next j
    Next i
    Call ConfigMapEvent
End Sub

Private Sub Respawn(ByVal player As Integer)
    Dim spawnPos As t_WorldPos
    spawnPos.map = MapNumber
    spawnPos.x = RandomNumber(12, 85)
    spawnPos.y = RandomNumber(12, 85)
    Call WarpToLegalPos(player, MapNumber, spawnPos.x, spawnPos.y, True, True)
    Call RevivirUsuario(player)
    UserList(player).Counters.timeFx = 2
114 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessageParticleFX(UserList(player).Char.charindex, e_ParticulasIndex.Curar, 100, False, , UserList(player).Pos.x, UserList(player).Pos.y))
116 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessagePlayWave("104", UserList(player).Pos.x, UserList(player).Pos.y))
End Sub

Private Sub MatchCompleted()
    MapInfo(MapNumber).Seguro = 1
    Call SetTimer(CountdownTimer, 1000)
    GenericGlobalLobby.State = Completed
    ElapsedTime = MaxTime
    Call Board.PublishScoreboard(SendTarget.toMap, MapNumber)
    Call CustomScenarios.ClearMap(MapNumber)
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
    currentTime = GetTickCount()
    Keys = PlayerRespawn.Keys
    For Each key In Keys
        If currentTime - PlayerRespawn.Item(key) > PlayerRespawnTime Then
            Call Respawn(key)
            Call PlayerRespawn.Remove(key)
        End If
    Next key
End Sub

Private Sub RespawnItems(ByVal frametime As Long)
    Dim i, j As Integer
    For i = 0 To UBound(ItemsToSpawn) - 1
        If UpdateTime(ItemsToSpawn(i).timer, frametime) Then
            For j = 0 To ItemsToSpawn(i).StackCount - 1
                Call SpawnItem(ItemsToSpawn(i).itemId, ItemsToSpawn(i).itemCount)
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
    spawnPos.x = RandomNumber(12, 85)
    spawnPos.y = RandomNumber(12, 85)
    Call TirarItemAlPiso(spawnPos, MiObj, 0)
End Sub

Private Sub ConfigMapEvent()
    MapInfo(MapNumber).Seguro = 0
    MapInfo(MapNumber).SinInviOcul = True
    MapInfo(mapNumber).SafeFightMap = True
End Sub

Private Sub ConfigMapLobby()
    PrevMapState = MapInfo(MapNumber).Seguro
    MapDropItems = MapInfo(MapNumber).DropItems
    SafeFigthState = MapInfo(mapNumber).SafeFightMap
    MapInfo(MapNumber).Seguro = 1
    MapInfo(MapNumber).SinInviOcul = True
    MapInfo(MapNumber).DropItems = False
End Sub


