VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScenarioHunt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const RequiredWeaponId = 1848
Const MaxActiveNpc = 30
Const configPath = "../init/Scenarios/Snakes.ini"
Const RequireItemMessage = 403
Const MillisecondInSecond = 1000

Implements IBaseScenario

Private Type e_Rank
    userIndex As Integer
    score As Integer
End Type

Dim MatchNpc As New Dictionary
Dim SpawnCreatureCount As Integer
Dim ElapsedTime As Long
Dim LastFrameTime As Long
Dim LastScoreUpdate As Long
Dim ActiveNpc As Integer
Dim SpawnedCreaturesId As New Dictionary
Dim MapNumber As Integer
Dim UpdateIndex As Integer
Dim MaxTime As Long
Dim TargetScore As Integer
Dim BroadCastInterval As Long
Attribute BroadCastInterval.VB_VarHelpID = -1
Dim LastBroadCastTime As Long
Dim FrameTime As Long
Dim board As New ScoreBoard
Dim StartTimer As t_Timer
Dim PrevMapState As Integer
Dim MapDropItems As Boolean


Friend Sub RegisterNpc(ByVal Tag, ByVal sectionSize, ByRef registry As Dictionary, ByRef File As clsIniManager)
    Dim i As Integer
    Dim npcEntry As EventNpcDefinition
    For i = 1 To sectionSize
        Set npcEntry = New EventNpcDefinition
        npcEntry.NpcId = val(File.GetValue(Tag & i, "NPCID"))
        npcEntry.SpawnCount = val(File.GetValue(Tag & i, "count"))
        npcEntry.Score = val(File.GetValue(Tag & i, "score"))
        If Not registry.Exists(npcEntry.NpcId) Then
            Call registry.Add(npcEntry.NpcId, npcEntry)
        End If
    Next i
End Sub

Private Sub ConfigMapLobby()
    PrevMapState = MapInfo(MapNumber).Seguro
    MapDropItems = MapInfo(MapNumber).DropItems
    MapInfo(MapNumber).Seguro = 1
    MapInfo(MapNumber).DropItems = False
End Sub

Private Sub Class_Initialize()
On Error GoTo Hunt_Initialize_Err:
        Dim validNpcCount As Integer
        Dim ScenearioIni As clsIniManager
104     Set ScenearioIni = New clsIniManager
        Debug.Assert FileExist(DatPath & configPath)
        
106     Call ScenearioIni.Initialize(DatPath & configPath)
108     validNpcCount = val(ScenearioIni.GetValue("INIT", "TargetNpc"))
112     mapNumber = val(ScenearioIni.GetValue("INIT", "MapNumber"))
114     MaxTime = val(ScenearioIni.GetValue("INIT", "MaxTime"))
        MaxTime = MaxTime * 1000
116     TargetScore = val(ScenearioIni.GetValue("INIT", "TargetScore"))
118     Call RegisterNpc("TARGET", validNpcCount, MatchNpc, ScenearioIni)
120     UpdateIndex = AddUpdateScenario(Me)
122     Call CustomScenarios.SetCustomScenario(mapNumber, Me)
124     LastBroadCastTime = 0
        Call ConfigMapLobby
        Exit Sub
Hunt_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.Initialize ", Erl)
End Sub

Private Sub CloseScenario()
On Error GoTo CloseScenario_Err:
100    Call ModLobby.ReturnAllPlayers(GenericGlobalLobby)
102    Set GenericGlobalLobby.scenario = Nothing
104    Call ModLobby.UpdateLobbyState(GenericGlobalLobby, e_LobbyState.UnInitilized)
106    Call ClearMap
    Exit Sub
CloseScenario_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.CloseScenario ", Erl)
End Sub

Private Sub ClearMap()
On Error GoTo ClearMap_Err:
106    CustomScenarios.RemoveCustomScenario (mapNumber)
108    CustomScenarios.RemoveUpdateScenario (UpdateIndex)
110    MapInfo(MapNumber).Seguro = PrevMapState
112    MapInfo(MapNumber).DropItems = MapDropItems
    Exit Sub
ClearMap_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.ClearMap", Erl)
End Sub

Public Sub IBaseScenario_Reset()
106        Call ClearMap
End Sub

Private Function IBaseScenario_RequiresSpawn() As Boolean
    IBaseScenario_RequiresSpawn = False
End Function

Private Function IBaseScenario_ValidateUser(ByVal userIndex As Integer) As Integer
On Error GoTo ValidateUser_Err:
       Dim i As Integer
100    With UserList(userIndex)
102        For i = 1 To UBound(.Invent.Object)
104            If .Invent.Object(i).objIndex = RequiredWeaponId Then
106                IBaseScenario_ValidateUser = 0
108                Exit Function
110            End If
112        Next i
114        IBaseScenario_ValidateUser = RequireItemMessage
116    End With
       Exit Function
ValidateUser_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.ValidateUser userIndex: " & userIndex, Erl)
End Function

Private Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    Dim SpawnPos As t_WorldPos
    SpawnPos.map = mapNumber
    SpawnPos.X = RandomNumber(12, 85)
    SpawnPos.y = RandomNumber(12, 85)
    GenericGlobalLobby.SummonCoordinates = SpawnPos
    Call board.AddPlayer(UserIndex)
End Sub

Private Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef newState As e_LobbyState)
    If oldState = newState Then
        Exit Sub
    End If
    If newState = InProgress Then
        Call SetTimer(StartTimer, MillisecondInSecond)
    End If
    If newState = Closed Then
        Call MatchComplete
    End If
End Sub

Private Sub StartGame()
    Dim key As Variant
    Dim i As Integer
    For i = 0 To GenericGlobalLobby.RegisteredPlayers - 1
        Dim spawnPos As t_WorldPos
        spawnPos.map = MapNumber
        spawnPos.x = RandomNumber(12, 85)
        spawnPos.y = RandomNumber(12, 85)
        If GenericGlobalLobby.Players(i).IsSummoned Then
            If IsValidUserRef(GenericGlobalLobby.Players(i).user) Then
                Call WarpToLegalPos(GenericGlobalLobby.Players(i).user.ArrayIndex, MapNumber, spawnPos.x, spawnPos.y, True, True)
            End If
        End If
    Next i
    For Each key In MatchNpc.Keys
        For i = 1 To MatchNpc(key).SpawnCount
            SpawnNpc (MatchNpc(key).NpcId)
        Next i
    Next key
    LastFrameTime = GetTickCount()
    BroadCastInterval = 60000
    Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg("Que inicie la caceria!!!", e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Private Sub IBaseScenario_PlayerKillNpc(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    If GenericGlobalLobby.State <> e_LobbyState.InProgress Then
        Exit Sub
    End If
    Dim npcType As Integer
    NPCtype = NpcList(npcIndex).Numero
    If sourceType = e_phisical And sourceIndex = RequiredWeaponId Then
        If MatchNpc.Exists(NPCtype) Then
            If board.UpdatePlayerScore(UserIndex, MatchNpc.Item(npcType).Score) >= TargetScore Then
                GenericGlobalLobby.State = Completed
                Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg("Tenemos un ganador!", e_FontTypeNames.FONTTYPE_GUILD))
                Call MatchComplete
                Exit Sub
            End If
        End If
    Else
        Call WriteLocaleMsg(userIndex, 402, FONTTYPE_INFO, ObjData(RequiredWeaponId).name)
    End If
    Call SpawnedCreaturesId.Remove(npcIndex)
    Call SpawnNpc(npcType)
End Sub

Private Sub IBaseScenario_NPcKillPlayer(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    'implements interface
End Sub

Private Sub IBaseScenario_PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    'implements interface
End Sub

Private Sub IBaseScenario_Update()
    Dim frameTime As Long
    frametime = GetTickCount() - LastFrameTime
    LastFrameTime = GetTickCount()
    If GenericGlobalLobby.State = e_LobbyState.InProgress Then
        If StartTimer.Occurrences < 10 Then
            If UpdateTime(StartTimer, frametime) Then
                If StartTimer.Occurrences >= 10 Then
                    Call StartGame
                Else
                    Call SendData(SendTarget.toMap, MapNumber, PrepareMessageConsoleMsg("La caceria iniciara en " & 10 - StartTimer.Occurrences, e_FontTypeNames.FONTTYPE_GUILD))
                End If
            End If
            Exit Sub
        End If
        ElapsedTime = ElapsedTime + FrameTime
        If ElapsedTime - LastScoreUpdate > 30000 Then
            LastScoreUpdate = ElapsedTime
            Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, GenericGlobalLobby, False)
        End If
        If ElapsedTime > MaxTime Then
            Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg("Se acabo el tiempo!", e_FontTypeNames.FONTTYPE_GUILD))
            Call MatchComplete
        End If
        Call BroadCastTimeUpdate
    ElseIf GenericGlobalLobby.State = e_LobbyState.Completed Then
        ElapsedTime = ElapsedTime + frametime
        If ElapsedTime - LastScoreUpdate > 1000 Then
            LastScoreUpdate = ElapsedTime
            Call CloseCountDown
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

Private Sub MatchComplete()
On Error GoTo MatchComplete_Err:
100    Dim i As Integer
102    GenericGlobalLobby.State = Completed
104    Call CustomScenarios.ClearMap(MapNumber)
106    Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, GenericGlobalLobby, True)
108    ElapsedTime = MaxTime
110    Dim key As Variant
112    For Each key In SpawnedCreaturesId.Keys
114        Call QuitarNPC(key, eClearHunt)
116    Next key
       Exit Sub
MatchComplete_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.MatchComplete ", Erl)
    Resume Next
End Sub

Private Sub SpawnNpc(ByVal npcType As Integer)
    Dim spawnPos As t_WorldPos
    Dim npcId As Integer
    spawnPos.map = MapNumber
    SpawnPos.X = RandomNumber(12, 85)
    SpawnPos.y = RandomNumber(12, 85)
    npcId = CrearNPC(npcType, MapNumber, spawnPos)
    NpcList(NpcId).flags.Respawn = 1
    NpcList(NpcId).Contadores.IntervaloRespawn = 0
    Call SpawnedCreaturesId.Add(NpcId, 0)
End Sub

Private Sub CloseCountDown()
    Dim pendingTime As Integer
    pendingTime = 10 - (ElapsedTime - MaxTime) / 1000
    If pendingTime <= 0 Then
        Call CloseScenario
        Exit Sub
    End If
    Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg("El mapa se cerrara en " & pendingTime & " segundos.", e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Private Sub BroadCastTimeUpdate()
    If BroadCastInterval > (ElapsedTime - LastBroadCastTime) Then
        Exit Sub
    End If
    LastBroadCastTime = ElapsedTime
    Dim pendingTime As Long
    'we add frame time here again so we round up
    pendingTime = (MaxTime - ElapsedTime + FrameTime) / 1000
    If pendingTime > MaxTime / 2000 Then
        Exit Sub
    ElseIf pendingTime >= 60 Then
        pendingTime = pendingTime / 60
        Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg("quedan " & pendingTime & " minutos.", e_FontTypeNames.FONTTYPE_GUILD))
        BroadCastInterval = 60000
    ElseIf pendingTime <= 10 Then
        Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg(pendingTime, e_FontTypeNames.FONTTYPE_GUILD))
    End If
    If pendingTime <= 1 Then
        BroadCastInterval = 1000
    End If
End Sub

