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
Const WaitingForPlayersTime = 300000 '5 minutes

Public Type PlayerInLobby
    SummonedFrom As t_WorldPos
    IsSummoned As Boolean
    User As t_UserReference
    UserId As Long
    Connected As Boolean
    ReturnOnReconnect As Boolean
    team As Integer
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

Public Type t_NewScenearioSettings
    MinLevel As Byte
    MaxLevel As Byte
    MinPlayers As Byte
    MaxPlayers As Byte
    TeamSize As Byte
    InscriptionFee As Long
    ScenearioType As Byte
    TeamType As Byte
    RoundNumber As Byte
    Description As String
    Password As String
End Type

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
    Canceled As Boolean
    Description As String
    Password As String
    IsGlobal As Boolean
    MapOpenTime As Long
    BroadOpenEvent As t_Timer
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

Public GlobalLobbyIndex         As Integer
Public CurrentActiveEventType   As e_EventType
Public LastAutoEventAttempt     As Long
Public AlreadyDidAutoEventToday As Boolean
Const LobbyCount = 200
Public LobbyList(0 To LobbyCount) As t_Lobby
Private AvailableLobby            As t_IndexHeap
Private ActiveLobby               As t_IndexHeap

Public Sub InitializeLobbyList()
    ReDim AvailableLobby.IndexInfo(0 To LobbyCount) As Integer
    ReDim ActiveLobby.IndexInfo(0 To LobbyCount) As Integer
    For i = 0 To LobbyCount
        AvailableLobby.IndexInfo(i) = LobbyCount - i
    Next i
    AvailableLobby.currentIndex = LobbyCount
    ActiveLobby.currentIndex = -1
    GlobalLobbyIndex = -1
End Sub

Public Sub ReleaseLobby(ByVal LobbyIndex As Integer)
    Dim i                As Integer
    Dim FoundActiveLobby As Boolean
    For i = 0 To ActiveLobby.currentIndex
        If ActiveLobby.IndexInfo(i) = LobbyIndex Then
            ActiveLobby.IndexInfo(i) = ActiveLobby.IndexInfo(ActiveLobby.currentIndex)
            ActiveLobby.currentIndex = ActiveLobby.currentIndex - 1
            FoundActiveLobby = True
            Exit For
        End If
    Next i
    If Not FoundActiveLobby Then
        LogError ("Trying to release a lobby twice")
        Exit Sub
    End If
    AvailableLobby.currentIndex = AvailableLobby.currentIndex + 1
    AvailableLobby.IndexInfo(AvailableLobby.currentIndex) = LobbyIndex
    If GlobalLobbyIndex = LobbyIndex Then
        GlobalLobbyIndex = -1
    End If
End Sub

Public Function GetAvailableLobby() As Integer
    If AvailableLobby.currentIndex < 0 Then
        GetAvailableLobby = -1
        Exit Function
    End If
    GetAvailableLobby = AvailableLobby.IndexInfo(AvailableLobby.currentIndex)
    AvailableLobby.currentIndex = AvailableLobby.currentIndex - 1
    ActiveLobby.currentIndex = ActiveLobby.currentIndex + 1
    ActiveLobby.IndexInfo(ActiveLobby.currentIndex) = GetAvailableLobby
End Function

Public Sub InitializeLobby(ByRef instance As t_Lobby)
    instance.MinLevel = 1
    instance.MaxLevel = 47
    instance.MaxPlayers = 100
    instance.MinPlayers = 1
    instance.SummonAfterInscription = True
    instance.RegisteredPlayers = 0
    instance.State = Initialized
    instance.SummonCoordinates.Map = -1
    instance.TeamSize = -1
    instance.TeamType = eRandom
    instance.TeamSortDone = False
    instance.NextTeamId = 1
    instance.AvailableInscriptionMoney = 0
    instance.InscriptionPrice = 0
    instance.Canceled = False
    instance.Password = ""
    instance.Description = ""
    instance.MapOpenTime = 0
    instance.IsGlobal = False
End Sub

Public Sub SetupLobby(ByRef instance As t_Lobby, ByRef LobbySettings As t_NewScenearioSettings)
    instance.MinLevel = LobbySettings.MinLevel
    instance.MaxLevel = LobbySettings.MaxLevel
    instance.MinPlayers = LobbySettings.MinPlayers
    Call SetMaxPlayers(instance, LobbySettings.MaxPlayers)
    instance.TeamSize = LobbySettings.TeamSize
    instance.TeamType = LobbySettings.TeamType
    instance.Description = LobbySettings.Description
    instance.Password = LobbySettings.Password
    instance.InscriptionPrice = LobbySettings.InscriptionFee
End Sub

Public Sub SetSummonCoordinates(ByRef instance As t_Lobby, ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
    instance.SummonCoordinates.Map = Map
    instance.SummonCoordinates.x = PosX
    instance.SummonCoordinates.y = PosY
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

Public Sub UpdateLobbyState(ByRef instance As t_Lobby, ByVal NewState As e_LobbyState)
    If Not instance.Scenario Is Nothing Then
        Call instance.Scenario.UpdateLobbyState(instance.State, NewState)
    End If
    instance.State = NewState
End Sub

Private Sub ClearUserSocket(ByRef instance As t_Lobby, ByVal Index As Integer)
    Dim i As Integer
    For i = Index To instance.RegisteredPlayers - 2
        instance.Players(i) = instance.Players(i + 1)
    Next i
    instance.Players(i).Connected = False
    instance.Players(i).IsSummoned = False
    instance.Players(i).ReturnOnReconnect = False
    instance.Players(i).team = -1
    Call ClearUserRef(instance.Players(i).User)
    instance.Players(i).UserId = 0
    instance.RegisteredPlayers = instance.RegisteredPlayers - 1
End Sub

Public Function CanPlayerJoin(ByRef instance As t_Lobby, ByVal UserIndex As Integer) As t_response
    On Error GoTo CanPlayerJoin_Err
    With UserList(UserIndex)
        If .Stats.ELV < instance.MinLevel Or .Stats.ELV > instance.MaxLevel Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = ForbiddenLevelMessage
            Exit Function
        End If
        If instance.RegisteredPlayers >= instance.MaxPlayers Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = LobbyIsFullMessage
            Exit Function
        End If
        If instance.ClassFilter > 0 And .clase <> instance.ClassFilter Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = ForbiddenClassMessage
            Exit Function
        End If
        If .flags.Muerto = 1 Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = MsgCantJoinEventDeath
            Exit Function
        End If
        If .flags.EnReto Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = MsgCantJoinWhileAnotherEvent
            Exit Function
        End If
        If .flags.EnConsulta Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = MsgCantJoinWhileAnotherEvent
            Exit Function
        End If
        If .pos.Map = 0 Or .pos.x = 0 Or .pos.y = 0 Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = MsgCantJoinWhileAnotherEvent
            Exit Function
        End If
        If .flags.EnTorneo Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = MsgCantJoinWhileAnotherEvent
            Exit Function
        End If
        If .Stats.GLD < instance.InscriptionPrice Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = MsgNotEnouthMoneyToParticipate
            Exit Function
        End If
        If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
            CanPlayerJoin.Success = False
            CanPlayerJoin.Message = MsgYouAreInJail
            Exit Function
        End If
        If Not instance.Scenario Is Nothing Then
            CanPlayerJoin.Message = instance.Scenario.ValidateUser(UserIndex)
            If CanPlayerJoin.Message > 0 Then
                CanPlayerJoin.Success = False
                Exit Function
            End If
        End If
        Dim i As Integer
        For i = 0 To instance.RegisteredPlayers - 1
            If instance.Players(i).UserId = .Id Then
                CanPlayerJoin.Success = False
                CanPlayerJoin.Message = AlreadyRegisteredMessage
                Exit Function
            End If
        Next i
        CanPlayerJoin.Success = True
        CanPlayerJoin.Message = 0
    End With
    Exit Function
CanPlayerJoin_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.CanPlayerJoin", Erl)
End Function

Public Function AddPlayer(ByRef instance As t_Lobby, ByVal UserIndex As Integer, Optional team As Integer = 0) As t_response
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
        Call SetUserRef(instance.Players(playerPos).User, UserIndex)
        instance.Players(playerPos).UserId = UserList(UserIndex).Id
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

Public Function AddPlayerOrGroup(ByRef instance As t_Lobby, ByVal UserIndex As Integer, ByVal Password As String) As t_response
    On Error GoTo AddPlayerOrGroup_Err
    With UserList(UserIndex)
        If Password <> instance.Password Then
            AddPlayerOrGroup.Message = MsgInvalidPassword
            AddPlayerOrGroup.Success = False
            Exit Function
        End If
        If instance.TeamSize > 1 And instance.TeamType = ePremade Then
            If Not .Grupo.EnGrupo Then
                AddPlayerOrGroup.Message = MsgTeamRequiredToJoin
                AddPlayerOrGroup.Success = False
                Exit Function
            End If
            If Not .Grupo.Lider.ArrayIndex = UserIndex Then
                AddPlayerOrGroup.Message = MsgOnlyLeaderCanJoin
                AddPlayerOrGroup.Success = False
                Exit Function
            End If
            If .Grupo.CantidadMiembros <> instance.TeamSize Then
                AddPlayerOrGroup.Message = MsgNotEnoughPlayersInGroup
                AddPlayerOrGroup.Success = False
                Exit Function
            End If
            Dim i As Integer
            For i = 1 To UBound(.Grupo.Miembros)
                If IsValidUserRef(.Grupo.Miembros(i)) Then
                    AddPlayerOrGroup = CanPlayerJoin(instance, .Grupo.Miembros(i).ArrayIndex)
                    If Not AddPlayerOrGroup.Success Then
                        Call WriteLocaleMsg(UserIndex, 1604, UserList(.Grupo.Miembros(i).ArrayIndex).name, e_FontTypeNames.FONTTYPE_New_Verde_Oscuro) 'Msg1604= ¬1: no puede participar, motivo: 'ver ReyarB
                        Call WriteLocaleMsg(UserIndex, AddPlayerOrGroup.Message, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                End If
            Next i
            For i = 1 To UBound(.Grupo.Miembros)
                If IsValidUserRef(.Grupo.Miembros(i)) Then
                    AddPlayerOrGroup = AddPlayer(instance, .Grupo.Miembros(i).ArrayIndex, instance.NextTeamId)
                End If
            Next i
            instance.NextTeamId = instance.NextTeamId + 1
        Else
            AddPlayerOrGroup = CanPlayerJoin(instance, UserIndex)
            If Not AddPlayerOrGroup.Success Then
                Exit Function
            End If
            AddPlayerOrGroup = AddPlayer(instance, UserIndex)
        End If
    End With
    Exit Function
AddPlayerOrGroup_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.AddPlayerOrGroup", Erl)
End Function

Public Sub SummonPlayer(ByRef instance As t_Lobby, ByVal User As Integer)
    On Error GoTo SummonPlayer_Err
    Dim UserIndex As Integer
    With instance.Players(User)
        If Not IsValidUserRef(.User) Then
            Call LogUserRefError(.User, "SummonPlayer")
            Exit Sub
        End If
        If Not .IsSummoned And .SummonedFrom.Map = 0 Then
            .SummonedFrom = UserList(.User.ArrayIndex).pos
        End If
        If Not instance.Scenario Is Nothing Then
            Call instance.Scenario.WillSummonPlayer(.User.ArrayIndex)
        End If
        Call WarpToLegalPos(.User.ArrayIndex, instance.SummonCoordinates.Map, instance.SummonCoordinates.x, instance.SummonCoordinates.y, True, True)
        .IsSummoned = True
    End With
    Exit Sub
SummonPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.SummonPlayer_Err", Erl)
End Sub

Public Sub SummonAll(ByRef instance As t_Lobby)
    On Error GoTo ReturnAllPlayer_Err
    Dim i As Integer
    For i = 0 To instance.RegisteredPlayers - 1
        Call SummonPlayer(instance, i)
    Next i
    Exit Sub
ReturnAllPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.SummonAll", Erl)
End Sub

Public Sub ReturnPlayer(ByRef instance As t_Lobby, ByVal User As Integer)
    On Error GoTo ReturnPlayer_Err
    With instance.Players(User)
        If Not IsValidUserRef(.User) Then
            Call LogUserRefError(.User, "ReturnPlayer")
            Exit Sub
        End If
        If Not .IsSummoned Then
            Exit Sub
        End If
        UserList(.User.ArrayIndex).flags.CurrentTeam = 0
        Call WarpToLegalPos(.User.ArrayIndex, .SummonedFrom.Map, .SummonedFrom.x, .SummonedFrom.y, True, True)
        .IsSummoned = False
    End With
    Exit Sub
ReturnPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnPlayer return user:" & User, Erl)
End Sub

Public Sub ReturnAllPlayers(ByRef instance As t_Lobby)
    On Error GoTo ReturnAllPlayer_Err
    Dim i As Integer
    For i = 0 To instance.RegisteredPlayers - 1
        Call ReturnPlayer(instance, i)
    Next i
    Exit Sub
ReturnAllPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.ReturnAllPlayer", Erl)
End Sub

Public Sub CancelLobby(ByRef instance As t_Lobby)
    On Error GoTo CancelLobby_Err
    instance.Canceled = True
    If instance.InscriptionPrice > 0 Then
        Dim i As Integer
        For i = 0 To instance.RegisteredPlayers - 1
            Call GiveGoldToPlayer(instance, i, instance.InscriptionPrice)
        Next i
    End If
    Call ReturnAllPlayers(instance)
    Call UpdateLobbyState(instance, Closed)
    instance.RegisteredPlayers = 0
    Exit Sub
CancelLobby_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.CancelLobby", Erl)
End Sub

Public Function GiveGoldToPlayer(ByRef instance As t_Lobby, ByVal UserSlotIndex As Integer, ByVal amount As Long) As Boolean
    On Error GoTo GiveMoneyToPlayer_Err
    If amount > instance.AvailableInscriptionMoney Then
        Call LogError("Instance is trying to give gold to " & instance.Players(UserSlotIndex).UserId & " but there is not enought gold collected")
        Exit Function
    End If
    If IsValidUserRef(instance.Players(UserSlotIndex).User) Then
        Call AddGold(instance.Players(UserSlotIndex).User.ArrayIndex, amount)
        instance.AvailableInscriptionMoney = instance.AvailableInscriptionMoney - amount
        GiveGoldToPlayer = True
    End If
    Exit Function
GiveMoneyToPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.GiveGoldToPlayer", Erl)
End Function

Public Sub ListPlayers(ByRef instance As t_Lobby, ByVal UserIndex As Integer)
    On Error GoTo ListPlayers_Err
    Dim i As Integer
    For i = 0 To instance.RegisteredPlayers - 1
        If instance.Players(i).Connected And IsValidUserRef(instance.Players(i).User) Then
            Call WriteConsoleMsg(UserIndex, i & ") " & UserList(instance.Players(i).User.ArrayIndex).name, e_FontTypeNames.FONTTYPE_INFOBOLD)
        Else
            Call WriteConsoleMsg(UserIndex, i & ") " & "Disconnected player.", e_FontTypeNames.FONTTYPE_New_Verde_Oscuro)
        End If
    Next i
    Exit Sub
ListPlayers_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.ListPlayers", Erl)
End Sub

Public Function OpenLobby(ByRef instance As t_Lobby, ByVal IsPublic As Boolean) As t_response
    On Error GoTo OpenLobby_Err
    Dim Ret           As t_response
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
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgCreateEventRoom, EventName & "¬" & instance.MaxPlayers & "¬" & instance.MinLevel & "¬" & instance.MaxLevel _
                & "¬" & instance.InscriptionPrice, e_FontTypeNames.FONTTYPE_GLOBAL))
        If Not instance.Scenario Is Nothing Then
            Call instance.Scenario.BroadcastOpenScenario
        End If
        If instance.InscriptionPrice > 0 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgBoardcastInscriptionPrice, instance.InscriptionPrice, e_FontTypeNames.FONTTYPE_GUILD))
        End If
    End If
    Call SetTimer(instance.BroadOpenEvent, 30000)
    instance.MapOpenTime = GlobalFrameTime
    Ret.Message = 401
    Ret.Success = True
    OpenLobby = Ret
    Exit Function
OpenLobby_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.OpenLobby", Erl)
End Function

Public Function WaitForPlayersTimeUp(ByRef instance As t_Lobby) As Boolean
    'global events waiting time is handled by game masters
    If instance.IsGlobal Then Exit Function
    WaitForPlayersTimeUp = GlobalFrameTime - instance.MapOpenTime > WaitingForPlayersTime
End Function

Public Sub UpdateWaitingForPlayers(ByVal frametime As Long, ByRef instance As t_Lobby)
    Dim i As Integer
    If instance.IsPublic Then
        If UpdateTime(instance.BroadOpenEvent, frametime) Then
            If instance.IsGlobal Then
                Call BroadcastOpenLobby(instance)
            Else
                For i = 0 To instance.RegisteredPlayers - 1
                    If IsValidUserRef(instance.Players(i).User) Then
                        Dim Seconds As Long
                        Dim Minutes As Long
                        Seconds = (WaitingForPlayersTime - (GlobalFrameTime - instance.MapOpenTime)) / 1000
                        Minutes = Seconds / 60
                        Seconds = Seconds - (Minutes * 60)
                        Call SendData(SendTarget.ToIndex, instance.Players(i).User.ArrayIndex, PrepareMessageLocaleMsg(1727, GetTimeString(Minutes, Seconds), _
                                e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1727=Esperando jugadores, La partida iniciara en ¬1 o cuando se llene la sala
                        Call SendData(SendTarget.ToIndex, instance.Players(i).User.ArrayIndex, PrepareMessageLocaleMsg(1728, instance.RegisteredPlayers & "¬" & _
                                instance.MaxPlayers & "¬" & instance.MinPlayers, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1728=En este momento hay ¬1 / ¬2 y se requiere un minimo de ¬3 para que pueda iniciar
                    End If
                Next i
            End If
        End If
    End If
    If WaitForPlayersTimeUp(instance) Then
        If instance.RegisteredPlayers >= instance.MinPlayers Then
            Call StartLobby(instance, -1)
        Else
            For i = 0 To instance.RegisteredPlayers - 1
                If IsValidUserRef(instance.Players(i).User) Then
                    Call SendData(SendTarget.ToIndex, instance.Players(i).User.ArrayIndex, PrepareMessageLocaleMsg(1729, "", e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1729=Evento cancelado por falta de jugadores
                End If
            Next i
            Call CancelLobby(instance)
        End If
    Else
        If instance.RegisteredPlayers >= instance.MaxPlayers Then
            Call StartLobby(instance, -1)
        End If
    End If
End Sub

Public Sub BroadcastOpenLobby(ByRef instance As t_Lobby)
    If Not instance.IsGlobal Then Exit Sub
    Dim EventName As String: EventName = "Evento"
    If Not instance.Scenario Is Nothing Then
        EventName = instance.Scenario.GetScenarioName()
    End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgOpenEventBroadcast, EventName, e_FontTypeNames.FONTTYPE_GUILD))
    If instance.InscriptionPrice > 0 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MsgBoardcastInscriptionPrice, instance.InscriptionPrice, e_FontTypeNames.FONTTYPE_GUILD))
    End If
End Sub

Public Sub ForceReset(ByRef instance As t_Lobby)
    On Error GoTo ForceReset_Err
    instance.MinLevel = 1
    instance.MaxLevel = 47
    instance.MaxPlayers = 0
    instance.MinPlayers = 1
    instance.SummonAfterInscription = True
    instance.RegisteredPlayers = 0
    instance.State = UnInitilized
    instance.SummonCoordinates.Map = -1
    instance.ClassFilter = -1
    If Not Scenario Is Nothing Then
        Call Scenario.Reset
    End If
    Set Scenario = Nothing
    Erase instance.Players
    Exit Sub
ForceReset_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.ForceReset", Erl)
    Resume Next
End Sub

Public Sub RegisterDisconnectedUser(ByVal DisconnectedUserIndex As Integer)
    Dim i As Integer
    For i = 0 To ActiveLobby.currentIndex
        Call RegisterDisconnectedUserOnLobby(LobbyList(i), DisconnectedUserIndex)
    Next i
End Sub

Public Sub RegisterDisconnectedUserOnLobby(ByRef instance As t_Lobby, ByVal DisconnectedUserIndex As Integer)
    On Error GoTo RegisterDisconnectedUser_Err
    If instance.State < AcceptingPlayers Then
        Exit Sub
    End If
    Dim i As Integer
    For i = 0 To instance.RegisteredPlayers - 1
        If instance.Players(i).User.ArrayIndex = DisconnectedUserIndex And IsValidUserRef(instance.Players(i).User) Then
            instance.Players(i).Connected = False
            If Not instance.Scenario Is Nothing Then
                instance.Scenario.OnUserDisconnected (DisconnectedUserIndex)
            End If
            If instance.Players(i).IsSummoned Then
                instance.Players(i).ReturnOnReconnect = True
                Call ReturnPlayer(instance, i)
            End If
            Exit Sub
        End If
    Next i
    Exit Sub
RegisterDisconnectedUser_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.RegisterDisconnectedUser", Erl)
End Sub

Public Sub RegisterReconnectedUser(ByVal DisconnectedUserIndex As Integer)
    Dim i As Integer
    For i = 0 To ActiveLobby.currentIndex
        Call RegisterReconnectedUserOnLobby(LobbyList(i), DisconnectedUserIndex)
    Next i
End Sub

Public Sub RegisterReconnectedUserOnLobby(ByRef instance As t_Lobby, ByVal UserIndex As Integer)
    On Error GoTo RegisterReconnectedUser_Err
    If instance.State < AcceptingPlayers Or instance.State >= Closed Then
        Exit Sub
    End If
    Dim i      As Integer
    Dim UserId As Long
    UserId = UserList(UserIndex).Id
    For i = 0 To instance.RegisteredPlayers - 1
        If instance.Players(i).UserId = UserId Then
            instance.Players(i).Connected = True
            Call SetUserRef(instance.Players(i).User, UserIndex)
            UserList(instance.Players(i).User.ArrayIndex).flags.CurrentTeam = instance.Players(i).team
            If instance.Players(i).ReturnOnReconnect Then
                Call SummonPlayer(instance, i)
            End If
            If Not instance.Scenario Is Nothing Then
                instance.Scenario.OnUserReconnect (UserIndex)
            End If
            Exit Sub
        End If
    Next i
    Exit Sub
RegisterReconnectedUser_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.RegisterReconnectedUser", Erl)
End Sub

Public Function SetTeamSize(ByRef instance As t_Lobby, ByVal TeamSize As Integer, ByVal TeamType As e_TeamTypes) As t_response
    On Error GoTo SetTeamSize_Err
    Dim response As t_response
    If instance.MaxPlayers Mod TeamSize <> 0 Then
        response.Success = False
        response.Message = MsgInvalidGroupCount
        SetTeamSize = response
        Exit Function
    End If
    If instance.State <> Initialized Then
        reponse.Success = False
        response.Message = MsgCantChangeGroupSizeNow
        SetTeamSize = response
        Exit Function
    End If
    response.Message = MsgTeamConfigSuccess
    instance.TeamSize = TeamSize
    instance.TeamType = TeamType
    instance.SortType = eFixedTeamSize
    response.Success = True
    SetTeamSize = response
    Exit Function
SetTeamSize_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.SetTeamSize", Erl)
End Function

Public Function SetTeamCount(ByRef instance As t_Lobby, ByVal TeamCount As Integer, ByVal TeamType As e_TeamTypes) As t_response
    On Error GoTo SetTeamSize_Err
    Dim response As t_response
    If instance.MaxPlayers Mod TeamCount <> 0 Then
        response.Success = False
        response.Message = MsgInvalidGroupCount
        SetTeamCount = response
        Exit Function
    End If
    If instance.State <> Initialized Then
        reponse.Success = False
        response.Message = MsgCantChangeGroupSizeNow
        SetTeamCount = response
        Exit Function
    End If
    response.Message = MsgTeamConfigSuccess
    instance.TeamSize = instance.MaxPlayers / TeamCount
    instance.TeamType = TeamType
    instance.SortType = eFixedTeamCount
    response.Success = True
    SetTeamCount = response
    Exit Function
SetTeamSize_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.SetTeamSize", Erl)
End Function

Public Sub StartLobby(ByRef instance As t_Lobby, ByVal UserIndex As Integer)
    If instance.State = Initialized And UserIndex >= 0 Then
        Call WriteLocaleMsg(UserIndex, 1605, e_FontTypeNames.FONTTYPE_INFO) 'Msg1605= El evento ya fue iniciado.
        Exit Sub
    End If
    If (instance.TeamSize > 1 Or instance.TeamType = eFixedTeamCount) And instance.TeamType = eRandom Then
        Call SortTeams(instance)
    End If
    Call ModLobby.UpdateLobbyState(instance, e_LobbyState.InProgress)
    If UserIndex >= 0 Then Call WriteLocaleMsg(UserIndex, 1606, e_FontTypeNames.FONTTYPE_INFO) 'Msg1606= Evento iniciado
End Sub

Public Function HandleRemoteLobbyCommand(ByVal Command, ByVal Params As String, ByVal UserIndex As Integer, ByVal LobbyIndex As Integer) As Boolean
    On Error GoTo HandleRemoteLobbyCommand_Err
    Dim Arguments() As String
    Dim RetValue    As t_response
    Dim tUser       As t_UserReference
    Arguments = Split(Params, " ")
    HandleRemoteLobbyCommand = True
    With UserList(UserIndex)
        Select Case Command
            Case e_LobbyCommandId.eSetSpawnPos
                Call SetSummonCoordinates(LobbyList(LobbyIndex), .pos.Map, .pos.x, .pos.y)
            Case e_LobbyCommandId.eEndEvent
                Call CancelLobby(LobbyList(LobbyIndex))
            Case e_LobbyCommandId.eReturnAllSummoned
                Call ModLobby.ReturnAllPlayers(LobbyList(LobbyIndex))
            Case e_LobbyCommandId.eReturnSinglePlayer
                Call ModLobby.ReturnPlayer(LobbyList(LobbyIndex), Arguments(0))
            Case e_LobbyCommandId.eSetClassLimit
                Call ModLobby.SetClassFilter(LobbyList(LobbyIndex), Arguments(0))
            Case e_LobbyCommandId.eSetMaxLevel
                Call ModLobby.SetMaxLevel(LobbyList(LobbyIndex), Arguments(0))
            Case e_LobbyCommandId.eSetMinLevel
                Call ModLobby.SetMinLevel(LobbyList(LobbyIndex), Arguments(0))
            Case e_LobbyCommandId.eOpenLobby
                RetValue = ModLobby.OpenLobby(LobbyList(LobbyIndex), Arguments(0))
            Case e_LobbyCommandId.eStartEvent
                Call StartLobby(LobbyList(LobbyIndex), UserIndex)
            Case e_LobbyCommandId.eSummonAll
                Call ModLobby.SummonAll(LobbyList(LobbyIndex))
            Case e_LobbyCommandId.eSummonSinglePlayer
                Call ModLobby.SummonPlayer(LobbyList(LobbyIndex), Arguments(0))
            Case e_LobbyCommandId.eListPlayers
                Call ModLobby.ListPlayers(LobbyList(LobbyIndex), UserIndex)
            Case e_LobbyCommandId.eForceReset
                Call ModLobby.ForceReset(LobbyList(LobbyIndex))
                Call WriteConsoleMsg(UserIndex, "Reset done.", e_FontTypeNames.FONTTYPE_INFO)
            Case e_LobbyCommandId.eSetTeamSize
                Call ModLobby.SetTeamSize(LobbyList(LobbyIndex), Arguments(0), Arguments(1))
                Call WriteConsoleMsg(UserIndex, "Team size set.", e_FontTypeNames.FONTTYPE_INFO)
            Case e_LobbyCommandId.eAddPlayer
                tUser = NameIndex(Params)
                If Not IsValidUserRef(tUser) Then
                    Call WriteConsoleMsg(UserIndex, "User " & Params & " not found.", e_FontTypeNames.FONTTYPE_INFO)
                    HandleRemoteLobbyCommand = False
                    Exit Function
                End If
                RetValue = ModLobby.AddPlayerOrGroup(LobbyList(LobbyIndex), tUser.ArrayIndex, "")
                If Not RetValue.Success Then
                    Call WriteConsoleMsg(UserIndex, "Failed to add player with message:", e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, RetValue.Message, e_FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Player has been registered", e_FontTypeNames.FONTTYPE_INFO)
                End If
                Call WriteLocaleMsg(tUser.ArrayIndex, RetValue.Message, e_FontTypeNames.FONTTYPE_INFO)
            Case e_LobbyCommandId.eSetInscriptionPrice
                If SetIncriptionPrice(LobbyList(LobbyIndex), Arguments(0)) Then
                    Call WriteConsoleMsg(UserIndex, "Inscription Price updated", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Failed to update insription price", e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case Else
                HandleRemoteLobbyCommand = False
                Exit Function
        End Select
    End With
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
    Dim i               As Integer
    Dim currentMaxLevel As Integer
    Dim currentIndex    As Integer
    currentMaxLevel = 0
    currentIndex = -1
    For i = 0 To instance.RegisteredPlayers - 1
        If instance.Players(i).team <= 0 Then
            If IsValidUserRef(instance.Players(i).User) Then
                If UserList(instance.Players(i).User.ArrayIndex).Stats.ELV > currentMaxLevel Then
                    currentMaxLevel = UserList(instance.Players(i).User.ArrayIndex).Stats.ELV
                    currentIndex = i
                End If
            End If
        End If
    Next i
    GetHigherLvlWithoutTeam = currentIndex
End Function

Public Sub SortTeams(ByRef instance As t_Lobby)
    On Error GoTo SortTeams_Err
    If instance.TeamSize < 1 Or (instance.MaxPlayers / instance.TeamSize) < 1 Then Exit Sub
    Dim currentIndex       As Integer
    Dim TeamCount          As Integer
    Dim MaxPossiblePlayers As Integer
    TeamCount = instance.MaxPlayers / instance.TeamSize
    If instance.SortType = eFixedTeamSize Then
        MaxPossiblePlayers = (instance.RegisteredPlayers / instance.TeamSize)
        MaxPossiblePlayers = MaxPossiblePlayers * instance.TeamSize
    Else
        MaxPossiblePlayers = instance.RegisteredPlayers / TeamCount
        MaxPossiblePlayers = MaxPossiblePlayers * TeamCount
    End If
    Dim i As Integer
    For i = instance.RegisteredPlayers - 1 To MaxPossiblePlayers Step -1
        If IsValidUserRef(instance.Players(i).User) Then
            Call WriteLocaleMsg(instance.Players(i).User.ArrayIndex, MsgNotEnoughPlayerForTeam, e_FontTypeNames.FONTTYPE_INFO)
        End If
        Call KickPlayer(instance, i)
    Next i
    TeamCount = instance.RegisteredPlayers / instance.TeamSize
    currentIndex = GetHigherLvlWithoutTeam(instance)
    Dim CurrentAssignTeam As Integer
    Dim Direction         As Integer
    Direction = 1
    CurrentAssignTeam = 1
    While currentIndex >= 0
        instance.Players(currentIndex).team = CurrentAssignTeam
        UserList(instance.Players(currentIndex).User.ArrayIndex).flags.CurrentTeam = CurrentAssignTeam
        CurrentAssignTeam = CurrentAssignTeam + Direction
        If CurrentAssignTeam > TeamCount Then 'we want to bound but repeat the team we add so we co 1 2 3 3 2 1 1 2 3....
            Direction = -1
            CurrentAssignTeam = TeamCount
        ElseIf CurrentAssignTeam < 1 Then
            Direction = 1
            CurrentAssignTeam = 1
        End If
        currentIndex = GetHigherLvlWithoutTeam(instance)
    Wend
    instance.TeamSortDone = True
    Exit Sub
SortTeams_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.SortTeams", Erl)
End Sub

Public Function KickPlayer(ByRef instance As t_Lobby, ByVal Index As Integer) As t_response
    On Error GoTo KickPlayer_Err
    Call ReturnPlayer(instance, Index)
    Call ClearUserSocket(instance, Index)
    Exit Function
KickPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.KickPlayer", Erl)
End Function

Public Function AllPlayersReady(ByRef instance As t_Lobby) As t_response
    On Error GoTo AllPlayersReady_Err
    Dim Ret As t_response
    Dim i   As Integer
    Ret.Success = True
    For i = 0 To instance.RegisteredPlayers - 1
        If Not IsValidUserRef(instance.Players(i).User) Then
            Ret.Success = False
            Ret.Message = MsgDisconnectedPlayers
        End If
    Next i
    AllPlayersReady = Ret
    Exit Function
AllPlayersReady_Err:
    Call TraceError(Err.Number, Err.Description, "ModLobby.AllPlayersReady", Erl)
End Function

Public Function GetOpenLobbyList(ByRef IdList() As Integer) As Integer
    Dim i         As Integer
    Dim OpenCount As Integer
    If ActiveLobby.currentIndex < 0 Then
        GetOpenLobbyList = 0
        Exit Function
    End If
    ReDim IdList(ActiveLobby.currentIndex) As Integer
    For i = 0 To ActiveLobby.currentIndex
        If LobbyList(ActiveLobby.IndexInfo(i)).State = AcceptingPlayers And LobbyList(ActiveLobby.IndexInfo(i)).IsPublic Then
            IdList(OpenCount) = ActiveLobby.IndexInfo(i)
            OpenCount = OpenCount + 1
        End If
    Next i
    GetOpenLobbyList = OpenCount
End Function

Public Function ValidateLobbySettings(ByRef LobbySettings As t_NewScenearioSettings)
    If LobbySettings.MinLevel < 1 Or LobbySettings.MaxLevel > 47 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1607, "", e_FontTypeNames.FONTTYPE_GLOBAL))
        Exit Function
    End If
    If LobbySettings.MinLevel > LobbySettings.MaxLevel Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1609, "", e_FontTypeNames.FONTTYPE_GLOBAL))
        Exit Function
    End If
    ValidateLobbySettings = True
End Function

Public Sub CreatePublicEvent(ByRef LobbySettings As t_NewScenearioSettings)
    GlobalLobbyIndex = GetAvailableLobby()
    If Not ValidateLobbySettings(LobbySettings) Then
        Exit Sub
    End If
    Call InitializeLobby(LobbyList(GlobalLobbyIndex))
    Call ModLobby.SetupLobby(LobbyList(GlobalLobbyIndex), LobbySettings)
    Call CustomScenarios.PrepareNewEvent(LobbySettings.ScenearioType, GlobalLobbyIndex)
    Call OpenLobby(LobbyList(GlobalLobbyIndex), True)
End Sub

Public Sub initEventLobby(ByVal UserIndex As Integer, ByVal eventType As Integer, LobbySettings As t_NewScenearioSettings)
    'aca se podria validar por nivel de patreon
    If eventType = 0 Then
        'a esto que esta aca abajo solo se accede si el lobby fue creado mediante comando GM
        CurrentActiveEventType = LobbySettings.ScenearioType
        Select Case LobbySettings.ScenearioType
            Case e_EventType.CaptureTheFlag
                Call HandleIniciarCaptura(LobbySettings)
            Case Else
                Call CreatePublicEvent(LobbySettings)
        End Select
    Else
        With UserList(UserIndex)
            If IsValidNpcRef(.flags.TargetNPC) Then
                If NpcList(.flags.TargetNPC.ArrayIndex).npcType = e_NPCType.EventMaster And .flags.Muerto = 0 Then
                    Call CreatePublicEvent(LobbySettings)
                End If
            End If
        End With
    End If
End Sub
