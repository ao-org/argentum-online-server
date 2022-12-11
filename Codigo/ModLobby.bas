Attribute VB_Name = "ModLobby"
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
    TeamSortDone As Boolean
    NextTeamId As Integer
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
108            Exit Function
110        End If
112        If instance.RegisteredPlayers >= instance.MaxPlayers Then
114            CanPlayerJoin.Success = False
116            CanPlayerJoin.Message = LobbyIsFullMessage
118            Exit Function
120        End If
122        If instance.ClassFilter > 0 And .clase <> instance.ClassFilter Then
124            CanPlayerJoin.Success = False
126            CanPlayerJoin.Message = ForbiddenClassMessage
128            Exit Function
130        End If
132        If Not instance.scenario Is Nothing Then
134            CanPlayerJoin.Message = instance.scenario.ValidateUser(UserIndex)
136            If CanPlayerJoin.Message > 0 Then
138                CanPlayerJoin.Success = False
140                Exit Function
142            End If
144        End If
146        Dim i As Integer
148        For i = 0 To instance.RegisteredPlayers - 1
150            If instance.Players(i).UserId = .ID Then
152                CanPlayerJoin.Success = False
154                CanPlayerJoin.Message = AlreadyRegisteredMessage
156                Exit Function
158            End If
160        Next i
        CanPlayerJoin.Success = True
        CanPlayerJoin.Message = 0
        End With
        Exit Function
CanPlayerJoin_Err:
190    Call TraceError(Err.Number, Err.Description, "ModLobby.CanPlayerJoin", Erl)
End Function

Public Function AddPlayer(ByRef instance As t_Lobby, ByVal UserIndex As Integer, Optional Team As Integer = 0) As t_response
On Error GoTo AddPlayer_Err
100    With UserList(UserIndex)
150        AddPlayer = CanPlayerJoin(instance, UserIndex)
152        If Not AddPlayer.Success Then
154            Exit Function
156        End If
162        Dim playerPos As Integer: playerPos = instance.RegisteredPlayers
164        Call SetUserRef(instance.Players(playerPos).user, UserIndex)
165        instance.Players(playerPos).UserId = UserList(UserIndex).ID
166        instance.Players(playerPos).IsSummoned = False
168        instance.Players(playerPos).Connected = True
170        UserList(UserIndex).flags.CurrentTeam = Team
172        instance.Players(playerPos).ReturnOnReconnect = False
173        instance.Players(playerPos).Team = Team
174        instance.RegisteredPlayers = instance.RegisteredPlayers + 1
176        AddPlayer.Message = JoinSuccessMessage
178        AddPlayer.Success = True
180        If instance.SummonAfterInscription Then
182            Call SummonPlayer(instance, playerPos)
184        End If
186    End With
188    Exit Function
AddPlayer_Err:
190    Call TraceError(Err.Number, Err.Description, "ModLobby.AddPlayer", Erl)
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
102    For i = 0 To instance.RegisteredPlayers
104        Call SummonPlayer(instance, i)
106    Next i
108    Exit Sub
ReturnAllPlayer_Err:
110     Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnAllPlayer", Erl)
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
100    Call ReturnAllPlayers(instance)
104    Call UpdateLobbyState(instance, Closed)
105    instance.RegisteredPlayers = 0
106    Exit Sub
CancelLobby_Err:
108     Call TraceError(Err.Number, Err.Description, "ModLobby.CancelLobby", Erl)
End Sub

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

Public Function OpenLobby(ByRef instance As t_Lobby, ByVal IsPublic As Boolean) As t_response
On Error GoTo OpenLobby_Err
       Dim Ret As t_response
       Dim RequiresSpawn As Boolean
100        If Not instance.Scenario Is Nothing Then
102            RequiresSpawn = instance.Scenario.RequiresSpawn
104        End If
106        RequiresSpawn = RequiresSpawn Or instance.SummonCoordinates.map > 0
108        If RequiresSpawn Then
110            Ret.Success = False
112            Ret.Message = 400
               OpenLobby = Ret
114            Exit Function
116        End If
117        instance.IsPublic = IsPublic
118        Call UpdateLobbyState(instance, AcceptingPlayers)
120        Ret.Message = 401
124        Ret.Success = True
           OpenLobby = Ret
126    Exit Function
OpenLobby_Err:
128     Call TraceError(Err.Number, Err.Description, "ModLobby.OpenLobby", Erl)
End Function

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
114            If instance.Players(i).IsSummoned Then
116                instance.Players(i).ReturnOnReconnect = True
118                Call ReturnPlayer(instance, i)
120            End If
122            If Not instance.scenario Is Nothing Then
124                instance.scenario.OnUserDisconnected (DisconnectedUserIndex)
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
132 response.Success = True
134 SetTeamSize = response
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
148             RetValue = ModLobby.OpenLobby(GenericGlobalLobby, Arguments(0))
150             If Arguments(0) And RetValue.Success Then
152                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " creó un nuevo evento, para participar ingresá /participar", e_FontTypeNames.FONTTYPE_GUILD))
154             End If
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
214         Case Else
216             HandleRemoteLobbyCommand = False
218             Exit Function
    End Select
220 End With
    Exit Function
HandleRemoteLobbyCommand_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.HandleRemoteLobbyCommand", Erl)
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
110 MaxPossiblePlayers = (instance.RegisteredPlayers / instance.TeamSize)
112 MaxPossiblePlayers = MaxPossiblePlayers * instance.TeamSize
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
