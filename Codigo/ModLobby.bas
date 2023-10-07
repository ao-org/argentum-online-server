Attribute VB_Name = "ModLobby"
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
Const ForbiddenLevelMessage = 396
Const LobbyIsFullMessage = 397
Const ForbiddenClassMessage = 398
Const JoinSuccessMessage = 399
Const AlreadyRegisteredMessage = 405
Public Type PlayerInLobby
    SummonedFrom As t_WorldPos
    IsSummoned As Boolean
    user As t_UserReference
    UserId As Long
    Connected As Boolean
    ReturnOnReconnect As Boolean
    Team As Integer
End Type

Public Enum e_LobbyState
    UnInitilized
    Initialized
    AcceptingPlayers
    InProgress
    Completed
    Closed
End Enum

Public Enum e_TeamTypes
    ePremade
    eRandom
End Enum

Public Enum e_SortType
    eFixedTeamSize
    eFixedTeamCount
End Enum

Type t_Lobby
    MinLevel As Byte
    MaxLevel As Byte
    MaxPlayers As Integer
    MinPlayers As Integer
    Players() As PlayerInLobby
    SummonCoordinates As t_WorldPos
    RegisteredPlayers As Integer
    ClassFilter As Integer 'check for e_Class or <= 0 for no filter
    State As e_LobbyState
    SummonAfterInscription As Boolean
    Scenario As IBaseScenario
    TeamSize As Integer
    IsPublic As Boolean
    TeamType As e_TeamTypes
    SortType As e_SortType
    TeamSortDone As Boolean
    NextTeamId As Integer
    InscriptionPrice As Long
    AvailableInscriptionMoney As Long
End Type

Public Type t_response
    Success As Boolean
    Message As Integer
End Type

Public Enum e_EventType
    Generic = 0
    CaptureTheFlag = 1
    NpcHunt = 2
    DeathMatch = 3
    NavalBattle = 4
End Enum

Public Enum e_LobbyCommandId
    eSetSpawnPos
    eSetMaxLevel
    eSetMinLevel
    eSetClassLimit
    eRegisterPlayer
    eSummonSinglePlayer
    eSummonAll
    eReturnSinglePlayer
    eReturnAllSummoned
    eOpenLobby
    eStartEvent
    eEndEvent
    eCancelEvent
    eListPlayers
    eKickPlayer
    eForceReset
    eSetTeamSize
    eAddPlayer
    eSetInscriptionPrice
End Enum
Public GenericGlobalLobby As t_Lobby
Public CurrentActiveEventType As e_EventType

Public Sub InitializeLobby(ByRef instance As t_Lobby)
    instance.MinLevel = 1
    instance.MaxLevel = 47
    instance.MaxPlayers = 100
    instance.MinPlayers = 1
    instance.SummonAfterInscription = True
    instance.RegisteredPlayers = 0
    instance.State = Initialized
    instance.SummonCoordinates.map = -1
    instance.TeamSize = -1
    instance.TeamType = eRandom
    instance.TeamSortDone = False
    instance.NextTeamId = 1
    instance.AvailableInscriptionMoney = 0
    instance.InscriptionPrice = 0
End Sub

Public Sub SetSummonCoordinates(ByRef instance As t_Lobby, ByVal map As Integer, ByVal posX As Integer, ByVal posY As Integer)
    instance.SummonCoordinates.map = map
    instance.SummonCoordinates.X = posX
    instance.SummonCoordinates.y = posY
End Sub

Public Sub SetMaxPlayers(ByRef instance As t_Lobby, ByVal playerCount As Integer)
    instance.MaxPlayers = playerCount
    ReDim instance.Players(0 To playerCount - 1)
End Sub

Public Sub SetMinPlayers(ByRef instance As t_Lobby, ByVal playerCount As Integer)
    instance.MinPlayers = playerCount
End Sub

Public Sub SetMinLevel(ByRef instance As t_Lobby, ByVal level As Byte)
    instance.MinLevel = level
End Sub

Public Sub SetMaxLevel(ByRef instance As t_Lobby, ByVal level As Byte)
    instance.MaxLevel = level
End Sub

Public Sub SetClassFilter(ByRef instance As t_Lobby, ByVal Class As Integer)
    instance.ClassFilter = Class
End Sub

Public Sub UpdateLobbyState(ByRef instance As t_Lobby, ByVal newState As e_LobbyState)
    If Not instance.Scenario Is Nothing Then
        Call instance.Scenario.UpdateLobbyState(instance.State, newState)
    End If
    instance.State = newState
End Sub

Private Sub ClearUserSocket(ByRef instance As t_Lobby, ByVal index As Integer)
    Dim i As Integer
    For i = index To instance.RegisteredPlayers - 2
        instance.Players(i) = instance.Players(i + 1)
    Next i
    instance.Players(i).connected = False
    instance.Players(i).IsSummoned = False
    instance.Players(i).ReturnOnReconnect = False
    instance.Players(i).Team = -1
    Call ClearUserRef(instance.Players(i).user)
    instance.Players(i).userID = 0
    instance.RegisteredPlayers = instance.RegisteredPlayers - 1
End Sub

Public Function CanPlayerJoin(ByRef instance As t_Lobby, ByVal UserIndex As Integer) As t_response
    On Error GoTo CanPlayerJoin_Err
100    With UserList(userIndex)
102        If .Stats.ELV < instance.MinLevel Or .Stats.ELV > instance.MaxLevel Then
104            CanPlayerJoin.Success = False
106            CanPlayerJoin.Message = ForbiddenLevelMessage
               Exit Function
           End If
108        If instance.RegisteredPlayers >= instance.MaxPlayers Then
110            CanPlayerJoin.Success = False
112            CanPlayerJoin.Message = LobbyIsFullMessage
               Exit Function
           End If
114        If instance.ClassFilter > 0 And .clase <> instance.ClassFilter Then
116            CanPlayerJoin.Success = False
118            CanPlayerJoin.Message = ForbiddenClassMessage
               Exit Function
           End If
120     If .flags.Muerto = 1 Then
122          CanPlayerJoin.Success = False
124          CanPlayerJoin.Message = MsgCantJoinEventDeath
             Exit Function
        End If
126     If .flags.EnReto Then
128         CanPlayerJoin.Success = False
130         CanPlayerJoin.Message = MsgCantJoinWhileAnotherEvent
            Exit Function
        End If
132     If .flags.EnConsulta Then
134        CanPlayerJoin.Success = False
136        CanPlayerJoin.Message = MsgCantJoinWhileAnotherEvent
           Exit Function
        End If
138     If .pos.Map = 0 Or .pos.x = 0 Or .pos.y = 0 Then
140         CanPlayerJoin.Success = False
142         CanPlayerJoin.Message = MsgCantJoinWhileAnotherEvent
            Exit Function
        End If
144     If .flags.EnTorneo Then
146         CanPlayerJoin.Success = False
148         CanPlayerJoin.Message = MsgCantJoinWhileAnotherEvent
            Exit Function
        End If
150     If .Stats.GLD < instance.InscriptionPrice Then
152         CanPlayerJoin.Success = False
154         CanPlayerJoin.Message = MsgNotEnouthMoneyToParticipate
            Exit Function
        End If
156     If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
158         CanPlayerJoin.Success = False
160         CanPlayerJoin.Message = MsgYouAreInJail
            Exit Function
        End If
162        If Not instance.Scenario Is Nothing Then
164            CanPlayerJoin.Message = instance.Scenario.ValidateUser(UserIndex)
166            If CanPlayerJoin.Message > 0 Then
168                CanPlayerJoin.Success = False
                   Exit Function
               End If
           End If
           Dim i As Integer
170        For i = 0 To instance.RegisteredPlayers - 1
172            If instance.Players(i).UserId = .id Then
174                CanPlayerJoin.Success = False
176                CanPlayerJoin.Message = AlreadyRegisteredMessage
                   Exit Function
               End If
178        Next i
180     CanPlayerJoin.Success = True
182     CanPlayerJoin.Message = 0
        End With
        Exit Function
CanPlayerJoin_Err:
184    Call TraceError(Err.Number, Err.Description, "ModLobby.CanPlayerJoin", Erl)
End Function

Public Function AddPlayer(ByRef instance As t_Lobby, ByVal UserIndex As Integer, Optional Team As Integer = 0) As t_response
On Error GoTo AddPlayer_Err
   With UserList(UserIndex)
       AddPlayer = CanPlayerJoin(instance, UserIndex)
       If Not AddPlayer.Success Then
           Exit Function
       End If
       If instance.InscriptionPrice > 0 Then
        If Not RemoveGold(UserIndex, instance.InscriptionPrice) Then
            AddPlayer.Success = False
            AddPlayer.Message = MsgNotEnouthMoneyToParticipate
            Exit Function
        Else
            instance.AvailableInscriptionMoney = instance.AvailableInscriptionMoney + instance.InscriptionPrice
        End If
       End If
       Dim playerPos As Integer: playerPos = instance.RegisteredPlayers
       Call SetUserRef(instance.Players(playerPos).user, UserIndex)
       instance.Players(playerPos).userID = UserList(UserIndex).id
       instance.Players(playerPos).IsSummoned = False
       instance.Players(playerPos).Connected = True
       UserList(UserIndex).flags.CurrentTeam = team
       instance.Players(playerPos).ReturnOnReconnect = False
       instance.Players(playerPos).team = team
       instance.RegisteredPlayers = instance.RegisteredPlayers + 1
       AddPlayer.Message = JoinSuccessMessage
       AddPlayer.Success = True
       If Not instance.Scenario Is Nothing Then
           Call instance.Scenario.SendRules(UserIndex)
       End If
       If instance.SummonAfterInscription Then
           Call SummonPlayer(instance, playerPos)
       End If
   End With
   Exit Function
AddPlayer_Err:
   Call TraceError(Err.Number, Err.Description, "ModLobby.AddPlayer", Erl)
End Function

Public Function AddPlayerOrGroup(ByRef instance As t_Lobby, ByVal UserIndex As Integer) As t_response
On Error GoTo AddPlayerOrGroup_Err
100    With UserList(UserIndex)
102        If instance.TeamSize > 0 And instance.TeamType = ePremade Then
104             If Not .Grupo.EnGrupo Then
106                 AddPlayerOrGroup.Message = MsgTeamRequiredToJoin
108                 AddPlayerOrGroup.Success = False
                    Exit Function
                End If
110             If Not .Grupo.Lider.ArrayIndex = UserIndex Then
112                 AddPlayerOrGroup.Message = MsgOnlyLeaderCanJoin
114                 AddPlayerOrGroup.Success = False
116                 Exit Function
                End If
118             If .Grupo.CantidadMiembros <> instance.TeamSize Then
120                 AddPlayerOrGroup.Message = MsgNotEnoughPlayersInGroup
122                 AddPlayerOrGroup.Success = False
124                 Exit Function
                End If
                Dim i As Integer
136             For i = 1 To UBound(.Grupo.Miembros)
138                 If IsValidUserRef(.Grupo.Miembros(i)) Then
140                     AddPlayerOrGroup = CanPlayerJoin(instance, .Grupo.Miembros(i).ArrayIndex)
142                     If Not AddPlayerOrGroup.Success Then
                            Call WriteConsoleMsg(UserIndex, UserList(.Grupo.Miembros(i).ArrayIndex).name & ": no puede participar, motivo: ", e_FontTypeNames.FONTTYPE_New_Verde_Oscuro)
                            Call WriteLocaleMsg(UserIndex, AddPlayerOrGroup.Message, e_FontTypeNames.FONTTYPE_INFO)
150                         Exit Function
                        End If
                    End If
                Next i
160             For i = 1 To UBound(.Grupo.Miembros)
162                 If IsValidUserRef(.Grupo.Miembros(i)) Then
164                     AddPlayerOrGroup = AddPlayer(instance, .Grupo.Miembros(i).ArrayIndex, instance.NextTeamId)
                    End If
                Next i
                instance.NextTeamId = instance.NextTeamId + 1
170        Else
172             AddPlayerOrGroup = CanPlayerJoin(instance, UserIndex)
174             If Not AddPlayerOrGroup.Success Then
176                 Exit Function
                End If
180             AddPlayerOrGroup = AddPlayer(instance, UserIndex)
           End If
186    End With
188    Exit Function
AddPlayerOrGroup_Err:
       Call TraceError(Err.Number, Err.Description, "ModLobby.AddPlayerOrGroup", Erl)
End Function

Public Sub SummonPlayer(ByRef instance As t_Lobby, ByVal user As Integer)
On Error GoTo SummonPlayer_Err
100        Dim userIndex As Integer
102        With instance.Players(user)
103            If Not IsValidUserRef(.user) Then
104                Call LogUserRefError(.user, "SummonPlayer")
105                Exit Sub
106            End If
108            If Not .IsSummoned And .SummonedFrom.map = 0 Then
109                .SummonedFrom = UserList(.user.ArrayIndex).Pos
110            End If
112            If Not instance.scenario Is Nothing Then
114                Call instance.scenario.WillSummonPlayer(.user.ArrayIndex)
116            End If
118            Call WarpToLegalPos(.user.ArrayIndex, instance.SummonCoordinates.map, instance.SummonCoordinates.X, instance.SummonCoordinates.y, True, True)
120            .IsSummoned = True
122        End With
124        Exit Sub
SummonPlayer_Err:
126    Call TraceError(Err.Number, Err.Description, "ModLobby.SummonPlayer_Err", Erl)
End Sub

Public Sub SummonAll(ByRef instance As t_Lobby)
On Error GoTo ReturnAllPlayer_Err
100    Dim i As Integer
102    For i = 0 To instance.RegisteredPlayers - 1
104        Call SummonPlayer(instance, i)
106    Next i
108    Exit Sub
ReturnAllPlayer_Err:
110     Call TraceError(Err.Number, Err.Description, "ModLobby.SummonAll", Erl)
End Sub

Public Sub ReturnPlayer(ByRef instance As t_Lobby, ByVal user As Integer)
On Error GoTo ReturnPlayer_Err
100    With instance.Players(user)
103        If Not IsValidUserRef(.user) Then
104            Call LogUserRefError(.user, "ReturnPlayer")
105            Exit Sub
106        End If
108        If Not .IsSummoned Then
110            Exit Sub
112        End If
114        UserList(.user.ArrayIndex).flags.CurrentTeam = 0
116        Call WarpToLegalPos(.user.ArrayIndex, .SummonedFrom.map, .SummonedFrom.x, .SummonedFrom.y, True, True)
118        .IsSummoned = False
    End With
    Exit Sub
ReturnPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnPlayer return user:" & user, Erl)
End Sub

Public Sub ReturnAllPlayers(ByRef instance As t_Lobby)
On Error GoTo ReturnAllPlayer_Err
100    Dim i As Integer
102    For i = 0 To instance.RegisteredPlayers - 1
104        Call ReturnPlayer(instance, i)
106    Next i
108    Exit Sub
ReturnAllPlayer_Err:
110     Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnAllPlayer", Erl)
End Sub

Public Sub CancelLobby(ByRef instance As t_Lobby)
On Error GoTo CancelLobby_Err
       If instance.InscriptionPrice > 0 Then
            Dim i As Integer
            For i = 0 To instance.RegisteredPlayers - 1
                Call GiveGoldToPlayer(instance, i, instance.InscriptionPrice)
            Next i
       End If
100    Call ReturnAllPlayers(instance)
104    Call UpdateLobbyState(instance, Closed)
105    instance.RegisteredPlayers = 0
106    Exit Sub
CancelLobby_Err:
108     Call TraceError(Err.Number, Err.Description, "ModLobby.CancelLobby", Erl)
End Sub

Public Function GiveGoldToPlayer(ByRef instance As t_Lobby, ByVal UserSlotIndex As Integer, ByVal amount As Long) As Boolean
On Error GoTo GiveMoneyToPlayer_Err
    If amount > instance.AvailableInscriptionMoney Then
        Call LogError("Instance is trying to give gold to " & instance.Players(UserSlotIndex).UserId & " but there is not enought gold collected")
        Exit Function
    End If
    If IsValidUserRef(instance.Players(UserSlotIndex).user) Then
        Call AddGold(instance.Players(UserSlotIndex).user.ArrayIndex, amount)
        instance.AvailableInscriptionMoney = instance.AvailableInscriptionMoney - amount
    End If
GiveMoneyToPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.CancelLobby", Erl)
End Function

Public Sub ListPlayers(ByRef instance As t_Lobby, ByVal UserIndex As Integer)
On Error GoTo ListPlayers_Err
       Dim i As Integer
100    For i = 0 To instance.RegisteredPlayers - 1
102        If instance.Players(i).Connected And IsValidUserRef(instance.Players(i).user) Then
104            Call WriteConsoleMsg(userIndex, i & ") " & UserList(instance.Players(i).user.ArrayIndex).name, e_FontTypeNames.FONTTYPE_INFOBOLD)
106        Else
108            Call WriteConsoleMsg(userIndex, i & ") " & "Disconnected player.", e_FontTypeNames.FONTTYPE_New_Verde_Oscuro)
110        End If
112    Next i
114    Exit Sub
ListPlayers_Err:
116    Call TraceError(Err.Number, Err.Description, "ModLobby.ListPlayers", Erl)
End Sub

Public Function OpenLobby(ByRef instance As t_Lobby, ByVal IsPublic As Boolean, ByVal UserIndex As Integer) As t_response
On Error GoTo OpenLobby_Err
    Dim Ret As t_response
    Dim RequiresSpawn As Boolean
    If Not instance.Scenario Is Nothing Then
        RequiresSpawn = instance.Scenario.RequiresSpawn
    End If
    RequiresSpawn = RequiresSpawn Or instance.SummonCoordinates.Map > 0
    If RequiresSpawn Then
        Ret.Success = False
        Ret.Message = 400
        OpenLobby = Ret
        Exit Function
    End If
    instance.IsPublic = IsPublic
    Call UpdateLobbyState(instance, AcceptingPlayers)
    If IsPublic Then
        Dim EventName As String: EventName = "Evento"
        If Not instance.Scenario Is Nothing Then
             EventName = instance.Scenario.GetScenarioName()
        End If
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgCreateEventRoom, UserList(UserIndex).name & "¬" & EventName, e_FontTypeNames.FONTTYPE_GUILD))
        If Not instance.Scenario Is Nothing Then
             Call instance.Scenario.BroadcastOpenScenario(UserIndex)
        End If
        If instance.InscriptionPrice > 0 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgBoardcastInscriptionPrice, instance.InscriptionPrice, e_FontTypeNames.FONTTYPE_GUILD))
        End If
    End If
    Ret.Message = 401
    Ret.Success = True
    OpenLobby = Ret
   Exit Function
OpenLobby_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.OpenLobby", Erl)
End Function

Public Sub BroadcastOpenLobby(ByRef instance As t_Lobby)
    Dim EventName As String: EventName = "Evento"
        If Not instance.Scenario Is Nothing Then
             EventName = instance.Scenario.GetScenarioName()
        End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgOpenEventBroadcast, EventName, e_FontTypeNames.FONTTYPE_GUILD))
    If GenericGlobalLobby.InscriptionPrice > 0 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgBoardcastInscriptionPrice, GenericGlobalLobby.InscriptionPrice, e_FontTypeNames.FONTTYPE_GUILD))
    End If
End Sub

Public Sub ForceReset(ByRef instance As t_Lobby)
On Error GoTo ForceReset_Err

100    instance.MinLevel = 1
102    instance.MaxLevel = 47
104    instance.MaxPlayers = 0
106    instance.MinPlayers = 1
108    instance.SummonAfterInscription = True
110    instance.RegisteredPlayers = 0
112    instance.State = UnInitilized
114    instance.SummonCoordinates.map = -1
116    instance.ClassFilter = -1
118    If Not scenario Is Nothing Then
120        Call scenario.Reset
122    End If
124    Set scenario = Nothing
126    Erase instance.Players
       Exit Sub
ForceReset_Err:
128     Call TraceError(Err.Number, Err.Description, "ModLobby.ForceReset", Erl)
        Resume Next
End Sub

Public Sub RegisterDisconnectedUser(ByRef instance As t_Lobby, ByVal DisconnectedUserIndex As Integer)
On Error GoTo RegisterDisconnectedUser_Err
100    If instance.State < AcceptingPlayers Then
102        Exit Sub
104    End If
106    Dim i As Integer
108    For i = 0 To instance.RegisteredPlayers - 1
110        If instance.Players(i).User.ArrayIndex = DisconnectedUserIndex And IsValidUserRef(instance.Players(i).User) Then
112            instance.Players(i).connected = False
114            If Not instance.Scenario Is Nothing Then
116                instance.Scenario.OnUserDisconnected (DisconnectedUserIndex)
118            End If
120            If instance.Players(i).IsSummoned Then
122                instance.Players(i).ReturnOnReconnect = True
124                Call ReturnPlayer(instance, i)
126            End If

128            Exit Sub
130        End If
132    Next i
134    Exit Sub
RegisterDisconnectedUser_Err:
136     Call TraceError(Err.Number, Err.Description, "ModLobby.RegisterDisconnectedUser", Erl)
End Sub

Public Sub RegisterReconnectedUser(ByRef instance As t_Lobby, ByVal userIndex As Integer)
On Error GoTo RegisterReconnectedUser_Err
100    If instance.State < AcceptingPlayers Or instance.State >= Closed Then
102        Exit Sub
104    End If
106    Dim i As Integer
108    Dim userID As Long
110    userID = UserList(userIndex).ID
112    For i = 0 To instance.RegisteredPlayers - 1
114        If instance.Players(i).UserId = UserId Then
116            instance.Players(i).connected = True
118            Call SetUserRef(instance.Players(i).User, UserIndex)
119             UserList(instance.Players(i).user.ArrayIndex).flags.CurrentTeam = instance.Players(i).Team
120            If instance.Players(i).ReturnOnReconnect Then
122                Call SummonPlayer(instance, i)
124            End If
126            If Not instance.scenario Is Nothing Then
128                instance.scenario.OnUserReconnect (userIndex)
130            End If
132            Exit Sub
134        End If
136    Next i
138    Exit Sub
RegisterReconnectedUser_Err:
140     Call TraceError(Err.Number, Err.Description, "ModLobby.RegisterReconnectedUser", Erl)
End Sub

Public Function SetTeamSize(ByRef instance As t_Lobby, ByVal TeamSize As Integer, ByVal TeamType As e_TeamTypes) As t_response
On Error GoTo SetTeamSize_Err
100 Dim response As t_response
102 If instance.MaxPlayers Mod TeamSize <> 0 Then
104     response.Success = False
106     response.Message = MsgInvalidGroupCount
108     SetTeamSize = response
110     Exit Function
112 End If
114 If instance.State <> Initialized Then
116     reponse.Success = False
118     response.Message = MsgCantChangeGroupSizeNow
120     SetTeamSize = response
122     Exit Function
124 End If
126 response.Message = MsgTeamConfigSuccess
128 instance.TeamSize = TeamSize
130 instance.TeamType = TeamType
    instance.SortType = eFixedTeamSize
132 response.Success = True
134 SetTeamSize = response
    Exit Function
SetTeamSize_Err:
140     Call TraceError(Err.Number, Err.Description, "ModLobby.SetTeamSize", Erl)
End Function

Public Function SetTeamCount(ByRef instance As t_Lobby, ByVal TeamCount As Integer, ByVal TeamType As e_TeamTypes) As t_response
On Error GoTo SetTeamSize_Err
100 Dim response As t_response
102 If instance.MaxPlayers Mod TeamCount <> 0 Then
104     response.Success = False
106     response.Message = MsgInvalidGroupCount
108     SetTeamCount = response
110     Exit Function
112 End If
114 If instance.State <> Initialized Then
116     reponse.Success = False
118     response.Message = MsgCantChangeGroupSizeNow
120     SetTeamCount = response
122     Exit Function
124 End If
126 response.Message = MsgTeamConfigSuccess
128 instance.TeamSize = instance.MaxPlayers / TeamCount
130 instance.TeamType = TeamType
    instance.SortType = eFixedTeamCount
132 response.Success = True
134 SetTeamCount = response
    Exit Function
SetTeamSize_Err:
140     Call TraceError(Err.Number, Err.Description, "ModLobby.SetTeamSize", Erl)
End Function

Public Sub StartLobby(ByRef instance As t_Lobby, ByVal UserIndex As Integer)
    If instance.State = Initialized Then
        Call WriteConsoleMsg(UserIndex, "El evento ya fue iniciado.", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If instance.TeamSize > 0 And instance.TeamType = eRandom Then
        Call SortTeams(instance)
    End If
    Call ModLobby.UpdateLobbyState(GenericGlobalLobby, e_LobbyState.InProgress)
    Call WriteConsoleMsg(UserIndex, "Evento iniciado", e_FontTypeNames.FONTTYPE_INFO)
End Sub

Public Function HandleRemoteLobbyCommand(ByVal Command, ByVal Params As String, ByVal UserIndex As Integer) As Boolean
On Error GoTo HandleRemoteLobbyCommand_Err
100 Dim Arguments()    As String
    Dim RetValue As t_response
    Dim tUser As t_UserReference
102 Arguments = Split(Params, " ")
    HandleRemoteLobbyCommand = True
    With UserList(UserIndex)
    Select Case Command
            Case e_LobbyCommandId.eSetSpawnPos
110             Call SetSummonCoordinates(GenericGlobalLobby, .pos.map, .pos.x, .pos.y)
            Case e_LobbyCommandId.eEndEvent
120             Call CancelLobby(GenericGlobalLobby)
            Case e_LobbyCommandId.eReturnAllSummoned
128             Call ModLobby.ReturnAllPlayers(GenericGlobalLobby)
            Case e_LobbyCommandId.eReturnSinglePlayer
132             Call ModLobby.ReturnPlayer(GenericGlobalLobby, Arguments(0))
             Case e_LobbyCommandId.eSetClassLimit
136             Call ModLobby.SetClassFilter(GenericGlobalLobby, Arguments(0))
             Case e_LobbyCommandId.eSetMaxLevel
140             Call ModLobby.SetMaxLevel(GenericGlobalLobby, Arguments(0))
             Case e_LobbyCommandId.eSetMinLevel
144              Call ModLobby.SetMinLevel(GenericGlobalLobby, Arguments(0))
             Case e_LobbyCommandId.eOpenLobby
148             RetValue = ModLobby.OpenLobby(GenericGlobalLobby, Arguments(0), UserIndex)
            Case e_LobbyCommandId.eStartEvent
158             Call StartLobby(GenericGlobalLobby, UserIndex)
            Case e_LobbyCommandId.eSummonAll
164             Call ModLobby.SummonAll(GenericGlobalLobby)
            Case e_LobbyCommandId.eSummonSinglePlayer
168            Call ModLobby.SummonPlayer(GenericGlobalLobby, Arguments(0))
            Case e_LobbyCommandId.eListPlayers
172             Call ModLobby.ListPlayers(GenericGlobalLobby, UserIndex)
            Case e_LobbyCommandId.eForceReset
176             Call ModLobby.ForceReset(GenericGlobalLobby)
178             Call WriteConsoleMsg(UserIndex, "Reset done.", e_FontTypeNames.FONTTYPE_INFO)
            Case e_LobbyCommandId.eSetTeamSize
182             Call ModLobby.SetTeamSize(GenericGlobalLobby, Arguments(0), Arguments(1))
184             Call WriteConsoleMsg(UserIndex, "Team size set.", e_FontTypeNames.FONTTYPE_INFO)
            Case e_LobbyCommandId.eAddPlayer
186             tUser = NameIndex(Params)
188             If Not IsValidUserRef(tUser) Then
190                 Call WriteConsoleMsg(UserIndex, "User " & Params & " not found.", e_FontTypeNames.FONTTYPE_INFO)
192                 HandleRemoteLobbyCommand = False
194                 Exit Function
196             End If
198             RetValue = ModLobby.AddPlayerOrGroup(GenericGlobalLobby, tUser.ArrayIndex)
200             If Not RetValue.Success Then
202                 Call WriteConsoleMsg(UserIndex, "Failed to add player with message:", e_FontTypeNames.FONTTYPE_INFO)
204                 Call WriteLocaleMsg(UserIndex, RetValue.Message, e_FontTypeNames.FONTTYPE_INFO)
206             Else
208                 Call WriteConsoleMsg(UserIndex, "Player has been registered", e_FontTypeNames.FONTTYPE_INFO)
210             End If
212             Call WriteLocaleMsg(tUser.ArrayIndex, RetValue.Message, e_FontTypeNames.FONTTYPE_INFO)
            Case e_LobbyCommandId.eSetInscriptionPrice
                If SetIncriptionPrice(GenericGlobalLobby, Arguments(0)) Then
                    Call WriteConsoleMsg(UserIndex, "Inscription Price updated", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Failed to update insription price", e_FontTypeNames.FONTTYPE_INFO)
                End If
214         Case Else
216             HandleRemoteLobbyCommand = False
218             Exit Function
    End Select
220 End With
    Exit Function
HandleRemoteLobbyCommand_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.HandleRemoteLobbyCommand", Erl)
End Function

Function SetIncriptionPrice(ByRef instance As t_Lobby, ByVal price As Long) As Boolean
    If instance.State <> Initialized Then
        Exit Function
    End If
    instance.InscriptionPrice = price
    SetIncriptionPrice = True
End Function

Private Function GetHigherLvlWithoutTeam(ByRef instance As t_Lobby) As Integer
    Dim i As Integer
    Dim currentMaxLevel As Integer
    Dim currentIndex As Integer
    currentMaxLevel = 0
    currentIndex = -1
    For i = 0 To instance.RegisteredPlayers - 1
        If instance.Players(i).Team <= 0 Then
            If IsValidUserRef(instance.Players(i).user) Then
                If UserList(instance.Players(i).user.ArrayIndex).Stats.ELV > currentMaxLevel Then
                    currentMaxLevel = UserList(instance.Players(i).user.ArrayIndex).Stats.ELV
                    currentIndex = i
                End If
            End If
        End If
    Next i
    GetHigherLvlWithoutTeam = currentIndex
End Function

    
Public Sub SortTeams(ByRef instance As t_Lobby)
On Error GoTo SortTeams_Err
100 If instance.TeamSize < 1 Or (instance.MaxPlayers / instance.TeamSize) < 1 Then Exit Sub
102 Dim currentIndex As Integer
104 Dim TeamCount As Integer
106 Dim MaxPossiblePlayers As Integer
108 TeamCount = instance.MaxPlayers / instance.TeamSize
    If instance.SortType = eFixedTeamSize Then
110     MaxPossiblePlayers = (instance.RegisteredPlayers / instance.TeamSize)
112     MaxPossiblePlayers = MaxPossiblePlayers * instance.TeamSize
    Else
        MaxPossiblePlayers = instance.RegisteredPlayers / TeamCount
        MaxPossiblePlayers = MaxPossiblePlayers * TeamCount
    End If
114 Dim i As Integer
116 For i = instance.RegisteredPlayers - 1 To MaxPossiblePlayers Step -1
118     If IsValidUserRef(instance.Players(i).user) Then
120         Call WriteLocaleMsg(instance.Players(i).user.ArrayIndex, MsgNotEnoughPlayerForTeam, e_FontTypeNames.FONTTYPE_INFO)
122     End If
124     Call KickPlayer(instance, i)
126 Next i
128 TeamCount = instance.RegisteredPlayers / instance.TeamSize
130 currentIndex = GetHigherLvlWithoutTeam(instance)
132 Dim CurrentAssignTeam As Integer
138 Dim Direction As Integer
140 Direction = 1
142 CurrentAssignTeam = 1
    
144 While currentIndex >= 0
146     instance.Players(currentIndex).Team = CurrentAssignTeam
147     UserList(instance.Players(currentIndex).user.ArrayIndex).flags.CurrentTeam = CurrentAssignTeam
148     CurrentAssignTeam = CurrentAssignTeam + Direction
150     If CurrentAssignTeam > TeamCount Then 'we want to bound but repeat the team we add so we co 1 2 3 3 2 1 1 2 3....
152         Direction = -1
151         CurrentAssignTeam = TeamCount
154     ElseIf CurrentAssignTeam < 1 Then
156         Direction = 1
157         CurrentAssignTeam = 1
158     End If
160     currentIndex = GetHigherLvlWithoutTeam(instance)
162 Wend
174 instance.TeamSortDone = True
    Exit Sub
SortTeams_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.SortTeams", Erl)
End Sub

Public Function KickPlayer(ByRef instance As t_Lobby, ByVal index As Integer) As t_response
On Error GoTo KickPlayer_Err
    Call ReturnPlayer(instance, index)
    Call ClearUserSocket(instance, index)
    Exit Function
KickPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.KickPlayer", Erl)
End Function

Public Function AllPlayersReady(ByRef instance As t_Lobby) As t_response
On Error GoTo AllPlayersReady_Err
100 Dim Ret As t_response
102 Dim i As Integer
104 Ret.Success = True
106 For i = 0 To instance.RegisteredPlayers - 1
108     If Not IsValidUserRef(instance.Players(i).user) Then
110         Ret.Success = False
112         Ret.Message = MsgDisconnectedPlayers
114     End If
116 Next i
118 AllPlayersReady = Ret
    Exit Function
AllPlayersReady_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.AllPlayersReady", Erl)
End Function
