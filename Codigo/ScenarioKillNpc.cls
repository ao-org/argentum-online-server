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
Dim PlayersScore As New Dictionary
Dim SpawnedCreaturesId As New Dictionary
Dim MapNumber As Integer
Dim TopFive(5) As e_Rank
Dim UpdateIndex As Integer
Dim MaxTime As Long
Dim TargetScore As Integer
Dim BroadCastInterval As Long
Attribute BroadCastInterval.VB_VarHelpID = -1
Dim LastBroadCastTime As Long
Dim FrameTime As Long


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
116     TargetScore = val(ScenearioIni.GetValue("INIT", "TargetScore"))
118     Call RegisterNpc("TARGET", validNpcCount, MatchNpc, ScenearioIni)
120     UpdateIndex = AddUpdateScenario(Me)
122     Call CustomScenarios.SetCustomScenario(mapNumber, Me)
124     LastBroadCastTime = 0
        Exit Sub
Hunt_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.Initialize ", Erl)
End Sub

Private Sub CloseScenario()
On Error GoTo CloseScenario_Err:
100    Call ModLobby.ReturnAllPlayers(GenericGlobalLobby)
102    Set GenericGlobalLobby.scenario = Nothing
104    Call ModLobby.UpdateLobbyState(GenericGlobalLobby, e_LobbyState.UnInitilized)
106    CustomScenarios.RemoveCustomScenario (mapNumber)
108    CustomScenarios.RemoveUpdateScenario (UpdateIndex)
110    Dim key As Variant
112    For Each key In SpawnedCreaturesId.Keys
114        Call QuitarNPC(key)
116    Next key
    Exit Sub
CloseScenario_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.CloseScenario ", Erl)
End Sub

Public Sub IBaseScenario_Reset()
    'implements interface
End Sub

Private Function IBaseScenario_RequiresSpawn() As Boolean
    IBaseScenario_RequiresSpawn = False
End Function

Private Function IBaseScenario_ValidateUser(ByVal userIndex As Integer) As Integer
    Dim i As Integer
    With UserList(userIndex)
        For i = 1 To UBound(.Invent.Object)
            If .Invent.Object(i).objIndex = RequiredWeaponId Then
                IBaseScenario_ValidateUser = 0
                Exit Function
            End If
        Next i
        IBaseScenario_ValidateUser = 403 'TODO create valid message here
    End With
End Function

Private Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    Dim SpawnPos As t_WorldPos
    SpawnPos.map = mapNumber
    SpawnPos.X = RandomNumber(12, 85)
    SpawnPos.y = RandomNumber(12, 85)
    GenericGlobalLobby.SummonCoordinates = SpawnPos
    If Not PlayersScore.Exists(UserIndex) Then
        Call PlayersScore.Add(UserIndex, 0)
    End If
End Sub

Private Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef newState As e_LobbyState)
    If oldState = newState Then
        Exit Sub
    End If
    If newState = InProgress Then
        Call StartGame
    End If
End Sub

Private Sub StartGame()
    Dim key As Variant
    Dim i As Integer
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
            PlayersScore.Item(UserIndex) = PlayersScore.Item(UserIndex) + MatchNpc.Item(NPCtype).Score
            Call UpdateRanking(PlayersScore.Item(userIndex), userIndex)
            If PlayersScore.Item(UserIndex) >= TargetScore Then
                GenericGlobalLobby.State = Completed
                Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg("Tenemos un ganador!", e_FontTypeNames.FONTTYPE_GUILD))
                ElapsedTime = MaxTime
                Call PublishScoreboard
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

Private Sub PublishScoreboard()
    Dim i As Integer
    For i = 0 To UBound(TopFive)
        If TopFive(i).score <= 0 Then
            Exit Sub
        End If
        If i = 0 Then
            Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg("Posiciones:", e_FontTypeNames.FONTTYPE_GUILD))
        End If
        Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg((i + 1) & ") " & UserList(TopFive(i).UserIndex).name & ": " & TopFive(i).Score, e_FontTypeNames.FONTTYPE_GUILD))
    Next i
End Sub

Private Sub IBaseScenario_Update()
    Dim frameTime As Long
    If GenericGlobalLobby.State = e_LobbyState.InProgress Then
        FrameTime = GetTickCount() - LastFrameTime
        ElapsedTime = ElapsedTime + FrameTime
        LastFrameTime = GetTickCount()
        If ElapsedTime - LastScoreUpdate > 30000 Then
            LastScoreUpdate = ElapsedTime
            Call PublishScoreboard
        End If
        If ElapsedTime > MaxTime Then
            GenericGlobalLobby.State = Completed
            Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg("Se acabo el tiempo!", e_FontTypeNames.FONTTYPE_GUILD))
            Call PublishScoreboard
        End If
        Call BroadCastTimeUpdate
    ElseIf GenericGlobalLobby.State = e_LobbyState.Completed Then
        ElapsedTime = ElapsedTime + GetTickCount() - LastFrameTime
        LastFrameTime = GetTickCount()
        If ElapsedTime - LastScoreUpdate > 1000 Then
            LastScoreUpdate = ElapsedTime
            Call CloseCountDown
        End If
    End If
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

Public Sub UpdateRanking(ByVal value As Integer, ByVal userIndex As Integer)
    Dim insertPos As Integer
    insertPos = InsertionPos(value)
    If (insertPos <= UBound(TopFive)) Then
        If TopFive(insertPos).UserIndex <> UserIndex Then
            Call ShiftElement(insertPos, CurrentScorePos(UserIndex))
        End If
        TopFive(insertPos).userIndex = userIndex
        TopFive(insertPos).score = value
    End If
End Sub

Private Function InsertionPos(ByVal score As Integer) As Integer
    If TopFive(5).score >= score Then
        InsertionPos = UBound(TopFive) + 1
        Exit Function
    End If
    
    Dim currentPosition As Integer
    For currentPosition = 0 To UBound(TopFive)
        If TopFive(currentPosition).Score < Score Then
            Exit For
        End If
    Next currentPosition
    InsertionPos = currentPosition
End Function

Private Function CurrentScorePos(ByVal UserIndex As Integer) As Integer
    Dim i As Integer
    For i = 0 To UBound(TopFive) - 1
        If TopFive(i).UserIndex = UserIndex Then
            Exit For
        End If
    Next i
    CurrentScorePos = i
End Function

Sub ShiftElement(ByVal position As Integer, ByVal starPos As Integer)
    Dim i As Integer
    For i = starPos To position + 1 Step -1
        TopFive(i) = TopFive(i - 1)
    Next i
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
