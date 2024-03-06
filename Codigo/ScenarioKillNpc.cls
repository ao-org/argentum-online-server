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

Const RequiredWeaponId = 1848
Const MaxActiveNpc = 30
Const configPath = "../dat/Scenarios/Snakes.ini"
Const RequireItemMessage = 403
Const MillisecondInSecond = 1000

Implements IBaseScenario

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
Dim PlayerRespawn As New Dictionary
Dim PlayerRespawnTime As Long
Dim LobbyIndex As Integer

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
        Dim ScenarioIni As clsIniManager
100     Set ScenarioIni = New clsIniManager
102     Debug.Assert FileExist(DatPath & configPath)
        Dim OriginalMap As Integer
104     Call ScenarioIni.Initialize(DatPath & configPath)
106     validNpcCount = val(ScenarioIni.GetValue("INIT", "TargetNpc"))
108     OriginalMap = val(ScenarioIni.GetValue("INIT", "MapNumber"))
110     MapNumber = InstanceManager.GetNextAvailableInstance
112     Call InstanceManager.CloneMap(OriginalMap, MapNumber)
114     MaxTime = val(ScenarioIni.GetValue("INIT", "MaxTime"))
116     MaxTime = MaxTime * 1000
118     PlayerRespawnTime = val(ScenarioIni.GetValue("INIT", "PlayerRespawnTime"))
120     PlayerRespawnTime = PlayerRespawnTime * 1000
122     TargetScore = val(ScenarioIni.GetValue("INIT", "TargetScore"))
124     Call RegisterNpc("TARGET", validNpcCount, MatchNpc, ScenarioIni)
126     UpdateIndex = AddUpdateScenario(Me)
128     Call CustomScenarios.SetCustomScenario(MapNumber, Me)
130     LastBroadCastTime = 0
        Exit Sub
Hunt_Initialize_Err:
136 Call TraceError(Err.Number, Err.Description, "ScenarioHunt.Initialize ", Erl)
End Sub

Private Sub CloseScenario()
    On Error GoTo CloseScenario_Err:
100    Call ModLobby.ReturnAllPlayers(LobbyList(LobbyIndex))
102    Set LobbyList(LobbyIndex).Scenario = Nothing
104    Call ModLobby.UpdateLobbyState(LobbyList(LobbyIndex), e_LobbyState.UnInitilized)
106    Call ClearMap
108    If MapNumber > 0 Then
110         Call ReleaseLobby(LobbyIndex)
112         Call ReleaseInstance(MapNumber)
114         MapNumber = 0
       End If
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

Public Sub IBaseScenario_SetLobbyIndex(ByVal Value As Integer)
    LobbyIndex = Value
    Call ConfigMapLobby
End Sub

Public Function IBaseScenario_GetScenarioName() As String
    IBaseScenario_GetScenarioName = "Dia del garrote"
End Function

Public Sub IBaseScenario_BroadcastOpenScenario(ByVal CreatorIndex As Integer)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgHuntScenearioIntro, "", e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Public Sub IBaseScenario_SendRules(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageLocaleMsg(MsgHuntScenarioInstructions, "", e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Public Function IBaseScenario_GetExtraInterfaces() As Long
    IBaseScenario_GetExtraInterfaces = 0
End Function

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
    LobbyList(LobbyIndex).SummonCoordinates = SpawnPos
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
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        Dim spawnPos As t_WorldPos
        spawnPos.map = MapNumber
        spawnPos.x = RandomNumber(12, 85)
        spawnPos.y = RandomNumber(12, 85)
        If LobbyList(LobbyIndex).Players(i).IsSummoned Then
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).user) Then
                Call WarpToLegalPos(LobbyList(LobbyIndex).Players(i).user.ArrayIndex, MapNumber, SpawnPos.x, SpawnPos.y, True, True)
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
    If LobbyList(LobbyIndex).State <> e_LobbyState.InProgress Then
        Exit Sub
    End If
    Dim npcType As Integer
    NPCtype = NpcList(npcIndex).Numero
    If sourceType = e_phisical And sourceIndex = RequiredWeaponId Then
        If MatchNpc.Exists(NPCtype) Then
            If board.UpdatePlayerScore(UserIndex, MatchNpc.Item(npcType).Score) >= TargetScore Then
                LobbyList(LobbyIndex).State = Completed
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
    If LobbyList(LobbyIndex).State = e_LobbyState.InProgress Then
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
            Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, LobbyList(LobbyIndex), False)
        End If
        If ElapsedTime > MaxTime Then
            Call SendData(SendTarget.toMap, mapNumber, PrepareMessageConsoleMsg("Se acabo el tiempo!", e_FontTypeNames.FONTTYPE_GUILD))
            Call MatchComplete
        End If
        Call RespawnPlayers
        Call BroadCastTimeUpdate
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.Completed Then
        ElapsedTime = ElapsedTime + frametime
        If ElapsedTime - LastScoreUpdate > 1000 Then
            LastScoreUpdate = ElapsedTime
            Call CloseCountDown
        End If
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.AcceptingPlayers Then
        Call UpdateWaitingForPlayers(FrameTime, LobbyList(LobbyIndex))
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.Closed Then
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
       If Not LobbyList(LobbyIndex).Canceled Then
101        Call Board.PublishScoreboard(SendTarget.toMap, MapNumber, LobbyList(LobbyIndex), True)
102        Call GiveRewards
       End If
103    LobbyList(LobbyIndex).State = Completed
104    Call CustomScenarios.ClearMap(MapNumber)

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
    If NpcId > 0 Then
        NpcList(NpcId).flags.Respawn = 1
        NpcList(NpcId).Contadores.IntervaloRespawn = 0
        Call SpawnedCreaturesId.Add(NpcId, 0)
    Else
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("No se pudo spawnear el npc " & npcType & ", revisar.", e_FontTypeNames.FONTTYPE_TALK))
        Call TraceError(Err.Number, "failed to spawn Npc " & npcType, "ScenarioHunt.SpawnNpc ", Erl)
    End If
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

Private Sub Respawn(ByVal player As Integer)
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MapNumber
    SpawnPos.x = RandomNumber(12, 85)
    SpawnPos.y = RandomNumber(12, 85)
    Call WarpToLegalPos(player, MapNumber, SpawnPos.x, SpawnPos.y, True, True)
    Call RevivirUsuario(player)
    UserList(player).Counters.timeFx = 3
114 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessageParticleFX(UserList(player).Char.charindex, e_ParticulasIndex.Curar, 100, False, , UserList(player).pos.x, UserList(player).pos.y))
116 Call SendData(SendTarget.ToPCAliveArea, player, PrepareMessagePlayWave("104", UserList(player).pos.x, UserList(player).pos.y))
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
146             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(i & ") " & Teams(Ranking(i).PlayerIndex) & ": " & Ranking(i).Score & " puntos" & RewardStr, e_FontTypeNames.FONTTYPE_GUILD))
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
168             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg((i + 1) & ") " & UserList(Ranking(i).PlayerIndex).name & ": " & Ranking(i).Score & " puntos" & RewardStr, e_FontTypeNames.FONTTYPE_GUILD))
170         Next i
        End If
        Exit Sub
GiveRewards_Err:
    Call TraceError(Err.Number, Err.Description, "ScenarioHunt.GiveRewards ", Erl)
End Sub
