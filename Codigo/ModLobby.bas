Attribute VB_Name = "ModLobby"

Type PlayerInLobby
    SummonedFrom As t_WorldPos
    IsSummoned As Boolean
    UserId As Integer
    Connected As Boolean
    dbId As Integer
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
    With UserList(UserIndex)
        If .Stats.ELV < instance.MinLevel Or .Stats.ELV > instance.MaxLevel Then
            AddPlayer.Success = False
            AddPlayer.Message = 396
            Exit Function
        End If
        If instance.RegisteredPlayers >= instance.MaxPlayers Then
            AddPlayer.Success = False
            AddPlayer.Message = 397
            Exit Function
        End If
        If instance.ClassFilter > 0 And .clase <> instance.ClassFilter Then
            AddPlayer.Success = False
            AddPlayer.Message = 398
            Exit Function
        End If
        If Not instance.Scenario Is Nothing Then
            AddPlayer.Message = instance.Scenario.ValidateUser(userIndex)
            If AddPlayer.Message > 0 Then
                AddPlayer.Success = False
                Exit Function
            End If
        End If
        Dim playerPos As Integer: playerPos = instance.RegisteredPlayers
        instance.Players(playerPos).UserId = UserIndex
        instance.Players(playerPos).IsSummoned = False
        instance.Players(playerPos).Connected = True
        instance.Players(playerPos).dbId = .ID
        instance.Players(playerPos).ReturnOnReconnect = False
        instance.RegisteredPlayers = instance.RegisteredPlayers + 1
        AddPlayer.Message = 399
        AddPlayer.Success = True
        If instance.SummonAfterInscription Then
            Call SummonPlayer(instance, playerPos)
        End If
        
    End With
    Exit Function
AddPlayer_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.AddPlayer", Erl)

End Function

Public Sub SummonPlayer(ByRef instance As t_Lobby, ByVal user As Integer)
   On Error GoTo SummonPlayer_Err
        Dim UserIndex As Integer
        With instance.Players(user)
            UserIndex = .UserId
            If Not .IsSummoned And .SummonedFrom.map = 0 Then
                .SummonedFrom = UserList(UserIndex).Pos
            End If
            If Not instance.Scenario Is Nothing Then
                Call instance.scenario.WillSummonPlayer(UserIndex)
            End If
100         Call WarpToLegalPos(UserIndex, instance.SummonCoordinates.map, instance.SummonCoordinates.X, instance.SummonCoordinates.y, True, True)
            .IsSummoned = True
        End With
    Exit Sub
SummonPlayer_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.SummonPlayer_Err", Erl)
End Sub

Public Sub SummonAll(ByRef instance As t_Lobby)
On Error GoTo ReturnAllPlayer_Err
    Dim i As Integer
    For i = 0 To instance.RegisteredPlayers
        Call SummonPlayer(instance, i)
    Next i
    Exit Sub
ReturnAllPlayer_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnAllPlayer", Erl)
End Sub

Public Sub ReturnPlayer(ByRef instance As t_Lobby, ByVal user As Integer)
On Error GoTo ReturnPlayer_Err
    Dim UserIndex As Integer
    With instance.Players(user)
        UserIndex = .UserId
        If Not .IsSummoned Then
            Exit Sub
        End If
100         Call WarpToLegalPos(UserIndex, .SummonedFrom.map, .SummonedFrom.X, .SummonedFrom.y, True, True)
        .IsSummoned = False
    End With
    Exit Sub
ReturnPlayer_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnPlayer return user:" & User, Erl)
End Sub

Public Sub ReturnAllPlayers(ByRef instance As t_Lobby)
On Error GoTo ReturnAllPlayer_Err
    Dim i As Integer
    For i = 0 To instance.RegisteredPlayers - 1
        Call ReturnPlayer(instance, i)
    Next i
    Exit Sub
ReturnAllPlayer_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnAllPlayer", Erl)
End Sub

Public Sub CancelLobby(ByRef instance As t_Lobby)
On Error GoTo CancelLobby_Err
    Call ReturnAllPlayers(instance)
    instance.RegisteredPlayers = 0
    Call UpdateLobbyState(instance, Closed)
    Exit Sub
CancelLobby_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.CancelLobby", Erl)
End Sub

Public Sub ListPlayers(ByRef instance As t_Lobby, ByVal UserIndex As Integer)
On Error GoTo ListPlayers_Err
    Dim i As Integer
    For i = 0 To instance.RegisteredPlayers
        Call WriteConsoleMsg(UserIndex, i & ") " & UserList(instance.Players(i).UserId).name, e_FontTypeNames.FONTTYPE_INFOBOLD)
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

Public Sub RegisterDisconnectedUser(ByRef instance As t_Lobby, ByVal userIndex As Integer)
On Error GoTo RegisterDisconnectedUser_Err
    If instance.State = UnInitilized Then
        Exit Sub
    End If
    Dim i As Integer
    For i = 0 To UBound(instance.Players)
        If instance.Players(i).userID = userIndex Then
            instance.Players(i).Connected = False
            If instance.Players(i).IsSummoned Then
                instance.Players(i).ReturnOnReconnect = True
                Call ReturnPlayer(instance, i)
            End If
            If Not instance.Scenario Is Nothing Then
                instance.Scenario.OnUserDisconnected (userIndex)
            End If
            Exit Sub
        End If
    Next i
    Exit Sub
RegisterDisconnectedUser_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.RegisterDisconnectedUser", Erl)
End Sub

Public Sub RegisterReconnectedUser(ByRef instance As t_Lobby, ByVal userIndex As Integer)
On Error GoTo RegisterReconnectedUser_Err
    If instance.State = UnInitilized Then
        Exit Sub
    End If
    Dim i As Integer
    Dim userID As Integer
    userID = UserList(userIndex).ID
    For i = 0 To UBound(instance.Players)
        If instance.Players(i).dbId = userID Then
            instance.Players(i).Connected = True
            instance.Players(i).userID = userIndex
            If instance.Players(i).ReturnOnReconnect Then
                Call SummonPlayer(instance, i)
            End If
            If Not instance.Scenario Is Nothing Then
                instance.Scenario.OnUserReconnect (userIndex)
            End If
            Exit Sub
        End If
    Next i
    Exit Sub
RegisterReconnectedUser_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.RegisterReconnectedUser", Erl)
End Sub
