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
    Connected As Boolean
    ReturnOnReconnect As Boolean
End Type

Public Enum e_LobbyState
    UnInitilized
    Initialized
    AcceptingPlayers
    InProgress
    Completed
    Closed
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

Public Function AddPlayer(ByRef instance As t_Lobby, ByVal UserIndex As Integer) As t_response
On Error GoTo AddPlayer_Err
100    With UserList(userIndex)
102        If .Stats.ELV < instance.MinLevel Or .Stats.ELV > instance.MaxLevel Then
104            AddPlayer.Success = False
106            AddPlayer.Message = ForbiddenLevelMessage
108            Exit Function
110        End If
112        If instance.RegisteredPlayers >= instance.MaxPlayers Then
114            AddPlayer.Success = False
116            AddPlayer.Message = LobbyIsFullMessage
118            Exit Function
120        End If
122        If instance.ClassFilter > 0 And .clase <> instance.ClassFilter Then
124            AddPlayer.Success = False
126            AddPlayer.Message = ForbiddenClassMessage
128            Exit Function
130        End If
132        If Not instance.scenario Is Nothing Then
134            AddPlayer.Message = instance.scenario.ValidateUser(userIndex)
136            If AddPlayer.Message > 0 Then
138                AddPlayer.Success = False
140                Exit Function
142            End If
144        End If
146        Dim i As Integer
148        For i = 0 To instance.RegisteredPlayers - 1
150            If instance.Players(i).user.ExpectedId = .ID Then
152                AddPlayer.Success = False
154                AddPlayer.Message = AlreadyRegisteredMessage
156                Exit Function
158            End If
160        Next i
162        Dim playerPos As Integer: playerPos = instance.RegisteredPlayers
164        Call SetUserRef(instance.Players(playerPos).user, UserIndex)
166        instance.Players(playerPos).IsSummoned = False
168        instance.Players(playerPos).Connected = True
172        instance.Players(playerPos).ReturnOnReconnect = False
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
104            Call LogUserRefError(.user, "SummonPlayer")
105            Exit Sub
106        End If
108        If Not .IsSummoned Or Not .connected Then
110            Exit Sub
112        End If
114        Call WarpToLegalPos(.user.ArrayIndex, .SummonedFrom.map, .SummonedFrom.X, .SummonedFrom.y, True, True)
116        .IsSummoned = False
    End With
    Exit Sub
ReturnPlayer_Err:
118     Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnPlayer return user:" & user, Erl)
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
102    instance.RegisteredPlayers = 0
104    Call UpdateLobbyState(instance, Closed)
106    Exit Sub
CancelLobby_Err:
108     Call TraceError(Err.Number, Err.Description, "ModLobby.CancelLobby", Erl)
End Sub

Public Sub ListPlayers(ByRef instance As t_Lobby, ByVal UserIndex As Integer)
On Error GoTo ListPlayers_Err
    Dim i As Integer
    For i = 0 To instance.RegisteredPlayers
        If instance.Players(i).connected And IsValidUserRef(instance.Players(i).user) Then
            Call WriteConsoleMsg(UserIndex, i & ") " & UserList(instance.Players(i).user.ArrayIndex).name, e_FontTypeNames.FONTTYPE_INFOBOLD)
        Else
            Call WriteConsoleMsg(userIndex, i & ") " & "Disconnected player.", e_FontTypeNames.FONTTYPE_New_Verde_Oscuro)
        End If
    Next i
    Exit Sub
ListPlayers_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.CancelLobby", Erl)
End Sub

Public Function OpenLobby(ByRef instance As t_Lobby) As t_response
On Error GoTo OpenLobby_Err
       Dim RequiresSpawn As Boolean
100        If Not instance.Scenario Is Nothing Then
102            RequiresSpawn = instance.Scenario.RequiresSpawn
104        End If
106        RequiresSpawn = RequiresSpawn Or instance.SummonCoordinates.map > 0
108        If RequiresSpawn Then
110            StartLobby.Success = False
112            StartLobby.Message = 400
114            Exit Function
116        End If
118        Call UpdateLobbyState(instance, AcceptingPlayers)
120        StartLobby.Message = 401
124        StartLobby.Success = True
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

Public Sub RegisterDisconnectedUser(ByRef instance As t_Lobby, ByVal userIndex As Integer)
On Error GoTo RegisterDisconnectedUser_Err
100    If instance.State < AcceptingPlayers Then
102        Exit Sub
104    End If
106    Dim i As Integer
108    For i = 0 To instance.RegisteredPlayers - 1
110        If instance.Players(i).user.ArrayIndex = UserIndex And IsValidUserRef(instance.Players(i).user) Then
112            instance.Players(i).connected = False
114            If instance.Players(i).IsSummoned Then
116                instance.Players(i).ReturnOnReconnect = True
118                Call ReturnPlayer(instance, i)
120            End If
122            If Not instance.scenario Is Nothing Then
124                instance.scenario.OnUserDisconnected (userIndex)
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
108    Dim userID As Integer
110    userID = UserList(userIndex).ID
112    For i = 0 To instance.RegisteredPlayers - 1
114        If instance.Players(i).user.ExpectedId = userID Then
116            instance.Players(i).connected = True
118            instance.Players(i).user.ArrayIndex = UserIndex
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
