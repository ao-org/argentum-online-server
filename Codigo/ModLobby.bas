Attribute VB_Name = "ModLobby"

Type PlayerInLobby
    SummonedFrom As t_WorldPos
    IsSummoned As Boolean
    UserId As Integer
End Type

Public Enum e_LobbyState
    UnInitilized
    Initialized
    AcceptingPlayers
    InProgress
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
    eStartEvent
    eEndEvent
    eCancelEvent
    eListPlayers
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
            If Not .IsSummoned Then
                .SummonedFrom = UserList(UserIndex).Pos
            End If
            If Not instance.Scenario Is Nothing Then
                Call instance.Scenario.WillSummonPlayer(userIndex, instance)
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
102     Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnPlayer", Erl)
End Sub

Public Sub ReturnAllPlayers(ByRef instance As t_Lobby)
On Error GoTo ReturnAllPlayer_Err
    Dim i As Integer
    For i = 0 To instance.RegisteredPlayers
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

Public Function StartLobby(ByRef instance As t_Lobby) As t_response
On Error GoTo StartLobby_Err
        If instance.SummonCoordinates.map < 0 Then
            StartLobby.Success = False
            StartLobby.Message = 400
            Exit Function
        End If
        Call UpdateLobbyState(instance, AcceptingPlayers)
        StartLobby.Message = 401
        StartLobby.Success = True
    Exit Function
StartLobby_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLobby.StartLobby", Erl)
End Function
