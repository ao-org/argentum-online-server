Attribute VB_Name = "modNetwork"
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


Public DisconnectTimeout As Long

Public Type t_ConnectionMapping
    UserRef As t_UserReference
    ConnectionDetails As t_ConnectionInfo
    TimeLastReset As Long
    PacketCount As Long
End Type




#If DIRECT_PLAY = 0 Then
Public Mapping() As t_ConnectionMapping

Private Const TIME_RECV_FREQUENCY As Long = 0  ' In milliseconds
Private Const TIME_SEND_FREQUENCY As Long = 0 ' In milliseconds

Private Server  As Network.Server
Private Time(2) As Single

Private FramePacketCount As Long
Private NewFrameConnections As Long

Const MaxActiveConnections = 10000

Private PendingConnections As New Dictionary

Public Sub Listen(ByVal Limit As Long, ByVal Address As String, ByVal Service As String)
    Set Server = New Network.Server
    ReDim Mapping(1 To MaxActiveConnections) As t_ConnectionMapping
    
    Call Server.Attach(AddressOf OnServerConnect, AddressOf OnServerClose, AddressOf OnServerSend, AddressOf OnServerRecv)
    
    Call Server.Listen(Limit, Address, Service)
End Sub

Public Sub Disconnect()
    Call Server.Close
End Sub

Public Sub Tick(ByVal Delta As Single)
    Time(0) = Time(0) + Delta
    Time(1) = Time(1) + Delta
    Dim PerformanceTimer As Long
    FramePacketCount = 0
    NewFrameConnections = 0
    Call PerformanceTestStart(PerformanceTimer)
    If (Time(0) >= TIME_RECV_FREQUENCY) Then
        Time(0) = 0
        Call Server.Poll
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "modNetwork Poll, packets: " & FramePacketCount & " new connections: " & NewFrameConnections, 200)
    If (Time(1) >= TIME_SEND_FREQUENCY) Then
        Time(1) = 0
        
        Call Server.Flush
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "modNetwork flush", 200)
End Sub

Public Sub Poll()
    Call Server.Poll
    Call Server.Flush
End Sub

Public Sub Send(ByVal UserIndex As Long, ByRef Buffer As Network.Writer)
    Call Server.Send(UserList(UserIndex).ConnectionDetails.ConnID, False, Buffer)
End Sub

Public Sub SendToConnection(ByVal ConnectionId As Long, ByRef Buffer As Network.Writer)
    Call Server.Send(ConnectionId, False, Buffer)
End Sub

Public Sub Flush(ByVal UserIndex As Long)
    Call Server.Flush(UserList(UserIndex).ConnectionDetails.ConnID)
End Sub

Public Sub Kick(ByVal Connection As Long, Optional ByVal message As String = vbNullString)
On Error GoTo Kick_ErrHandler:
    If IsFeatureEnabled("debug_connections") Then
        If (Message <> vbNullString) Then
        Call AddLogToCircularBuffer("Kick connection: " & Connection & " reason: " & Message)
        Else
            Call AddLogToCircularBuffer("Kick connection: " & Connection)
        End If
    End If
    Dim UserRef As t_UserReference
    UserRef = Mapping(Connection).UserRef
    If (message <> vbNullString) Then
        If UserRef.ArrayIndex > 0 Then
            Call Protocol_Writes.WriteErrorMsg(UserRef.ArrayIndex, Message)
            If UserList(UserRef.ArrayIndex).flags.UserLogged Then
                Call Cerrar_Usuario(UserRef.ArrayIndex)
            End If
        End If
    End If
        
    Call Server.Flush(Connection)
    Call Server.Kick(Connection, True)
    Exit Sub
    
Kick_ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Kick", Erl)
End Sub

Public Function GetTimeOfNextFlush() As Single
    GetTimeOfNextFlush = max(0, TIME_SEND_FREQUENCY - Time(1))
End Function


Public Sub close_not_logged_sockets_if_timeout()
On Error GoTo close_not_logged_sockets_if_timeout_ErrHandler:
        Dim i As Integer
        Dim key As Variant
        Dim Ticks As Long, Delta As Long
100     Ticks = GetTickCount
102     For Each key In PendingConnections.Keys
104         With Mapping(key)
                Dim ConnectionId As Long
106             ConnectionId = key
108             Delta = Ticks - Mapping(ConnectionId).ConnectionDetails.OnConnectTimestamp
110             If Delta > PendingConnectionTimeout Then
112                 If IsValidUserRef(.UserRef) Then
114                     LogError ("trying to kick an assigned connection: " & ConnectionId & " assigned to: " & .UserRef.ArrayIndex)
                    Else
116                     Call KickConnection(ConnectionId)
                    End If
                End If
            End With
118     Next key
        Exit Sub
close_not_logged_sockets_if_timeout_ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Kick", Erl)
End Sub

Private Sub OnServerConnect(ByVal Connection As Long, ByVal Address As String)
    On Error GoTo OnServerConnect_Err:
100     NewFrameConnections = NewFrameConnections + 1
102     If IsFeatureEnabled("debug_connections") Then
104         Call AddLogToCircularBuffer("OnServerConnect connecting new user on id: " & Connection & " ip: " & Address)
        End If
106     If Mapping(Connection).UserRef.ArrayIndex > 0 Then
108         Call TraceError(Err.Number, Err.Description, "OnServerConnect Mapping(Connection).UserRef.ArrayIndex > 0, connection: " & Connection & " value: " & Mapping(Connection).UserRef.ArrayIndex & ", is valid: " & IsValidUserRef(Mapping(Connection).UserRef), Erl)
        End If
        
110     If Connection <= MaxActiveConnections Then
112         If IsValidUserRef(Mapping(Connection).UserRef) Then
114             LogError ("opening a new connection: " & Connection & " to a connection mapped to a user " & Mapping(Connection).UserRef.ArrayIndex)
            End If
116         If PendingConnections.Exists(Connection) Then
118             LogError ("opening a new connection id " & Connection & " with ip: " & Address & " but there already a pending connection with this id and ip: " & Mapping(Connection).ConnectionDetails.IP)
120             PendingConnections.Remove (Connection)
            End If
122         With Mapping(Connection)
124             .ConnectionDetails.ConnIDValida = True
126             .ConnectionDetails.IP = Address
128             .ConnectionDetails.ConnID = Connection
130             .ConnectionDetails.OnConnectTimestamp = GetTickCount()
132             .PacketCount = 0
134             .TimeLastReset = 0
            End With
136         Call PendingConnections.Add(Connection, Connection)
138         Call modSendData.SendToConnection(Connection, PrepareConnected())
        Else
142         Call Kick(Connection, "El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
        End If
    
        Exit Sub
    
OnServerConnect_Err:
144     Call Kick(Connection)
146     Call TraceError(Err.Number, Err.Description, "modNetwork.OnServerConnect", Erl)
End Sub

Private Sub OnServerClose(ByVal Connection As Long)
On Error GoTo OnServerClose_Err:
    
    Dim UserRef As t_UserReference
100    UserRef = Mapping(Connection).UserRef
102    If IsFeatureEnabled("debug_connections") Then
104        If UserRef.ArrayIndex > 0 Then
106            Call AddLogToCircularBuffer("OnServerClose disconnected user index: " & UserRef.ArrayIndex & " With connection id: " & Connection & " with name: " & UserList(UserRef.ArrayIndex).name & " and ip" & UserList(UserRef.ArrayIndex).ConnectionDetails.IP)
108        Else
110            Call AddLogToCircularBuffer("OnServerClose disconnected user index: " & UserRef.ArrayIndex & " With connection id: " & Connection)
112        End If
114    End If
    
118    If IsValidUserRef(UserRef) Then
120        If UserList(UserRef.ArrayIndex).flags.UserLogged Then
122            Call CloseSocketSL(UserRef.ArrayIndex)
124            Call Cerrar_Usuario(UserRef.ArrayIndex)
126        Else
128            Call CloseSocket(UserRef.ArrayIndex)
130        End If
    
132        UserList(UserRef.ArrayIndex).ConnectionDetails.ConnIDValida = False
134        UserList(UserRef.ArrayIndex).ConnectionDetails.ConnID = 0
       ElseIf PendingConnections.Exists(Connection) Then
136        PendingConnections.Remove (Connection)
       End If
138    Call ClearConnection(Connection)
        
140    Exit Sub
    
OnServerClose_Err:
    Call ForcedClose(UserRef.ArrayIndex, Connection)
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnServerClose", Erl)
End Sub

Private Sub OnServerSend(ByVal Connection As Long, ByVal Message As Network.Reader)
On Error GoTo OnServerSend_Err:
    
    Exit Sub
    
OnServerSend_Err:
    Call Kick(Connection)
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnServerSend", Erl)
End Sub

Private Sub OnServerRecv(ByVal Connection As Long, ByVal Message As Network.Reader)
On Error GoTo OnServerRecv_Err:
    
    Dim UserRef As t_UserReference
    UserRef = Mapping(Connection).UserRef
    FramePacketCount = FramePacketCount + 1
    
    If IsValidUserRef(UserRef) Then
        ' The user index is assigned to the connection in:
        '       HandleLoginExistingChar(ConnectionID)
        '       HandleLoginNewChar(ConnectionID)
        ' It does not make sense to pass the index if it has not being assigned
        Call Protocol.HandleIncomingData(connection, Message, UserRef.ArrayIndex)
    Else
        Call Protocol.HandleIncomingData(connection, Message)
    End If
    
    Exit Sub
    
OnServerRecv_Err:
    Call Kick(Connection)
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnServerRecv", Erl)
End Sub

Private Sub ForcedClose(ByVal UserIndex As Integer, Connection As Long)
On Error GoTo ForcedClose_Err:
100     UserList(UserIndex).ConnectionDetails.ConnIDValida = False
102     UserList(UserIndex).ConnectionDetails.ConnID = 0
104     Call Server.Flush(Connection)
106     Call Server.Kick(Connection, True)
108     Call ClearUserRef(Mapping(Connection).UserRef)
110     Call IncreaseVersionId(userIndex)
        Exit Sub
ForcedClose_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.ForcedClose", Erl)
End Sub

Public Sub KickConnection(Connection As Long)
On Error GoTo ForcedClose_Err:
104     Call Server.Flush(Connection)
106     Call Server.Kick(Connection, True)
108     Call ClearConnection(Connection)
110     If PendingConnections.Exists(Connection) Then
112         Call PendingConnections.Remove(Connection)
        End If
        Exit Sub
ForcedClose_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.KickConnection", Erl)
End Sub

Public Sub CheckDisconnectedUsers()
    On Error GoTo CheckDisconnectedUsers_Err:
100     If DisconnectTimeout <= 0 Then
            Exit Sub
        End If
        Dim currentTime As Long
        Dim iUserIndex As Integer
102     currentTime = GetTickCount()
104     For iUserIndex = 1 To MaxUsers
106         With UserList(iUserIndex)
                'Conexion activa? y es un usuario loggeado?
108             If .ConnectionDetails.ConnIDValida = 0 And .flags.UserLogged Then
110                 If .ConnectionDetails.ConnID > 0 Then
112                     If currentTime - Mapping(.ConnectionDetails.ConnID).TimeLastReset > DisconnectTimeout Then
                            'mato los comercios seguros
114                         If .ComUsu.DestUsu.ArrayIndex > 0 Then
116                             If IsValidUserRef(.ComUsu.DestUsu) And UserList(.ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
118                                 If UserList(.ComUsu.DestUsu.ArrayIndex).ComUsu.DestUsu.ArrayIndex = iUserIndex Then
120                                     Call WriteConsoleMsg(.ComUsu.DestUsu.ArrayIndex, "Comercio cancelado por el otro usuario.", e_FontTypeNames.FONTTYPE_TALK)
122                                     Call FinComerciarUsu(.ComUsu.DestUsu.ArrayIndex)
                                    End If
                                End If
124                             Call FinComerciarUsu(iUserIndex)
                            End If
126                         Call Cerrar_Usuario(iUserIndex, True)
                        End If
                    End If
                End If
           End With
128     Next iUserIndex
    
        Exit Sub
CheckDisconnectedUsers_Err:
130     Call TraceError(Err.Number, Err.Description, "modNetwork.CheckDisconnectedUsers", Erl)
End Sub

Public Function MapConnectionToUser(ByVal ConnectionId As Long) As Integer
     On Error GoTo CheckDisconnectedUsers_Err:
        Dim FreeUser As Long
100     If Not PendingConnections.Exists(ConnectionId) Then
102         Call LogError("Connection " & ConnectionId & " is not waiting for assign")
            Exit Function
        End If
        
104     FreeUser = NextOpenUser()
106     If IsFeatureEnabled("debug_id_assign") Then
108         Call LogError("Assign userId: " & FreeUser & " to connection: " & ConnectionID)
        End If
        
110     If FreeUser < 0 Then
112         If IsFeatureEnabled("debug_connections") Then
114             Call LogError("Failed to find slot for new user, connection: " & Connection & " LastUser: " & LastUser)
            End If
116         Call Kick(ConnectionId, "El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
            Exit Function
        End If
        
118     If UserList(FreeUser).InUse Then
120        Call LogError("Trying to use an user slot marked as in use! slot: " & FreeUser)
122        FreeUser = NextOpenUser()
        End If
        
124     Call PendingConnections.Remove(ConnectionId)
126     UserList(FreeUser).ConnectionDetails = Mapping(ConnectionId).ConnectionDetails
128     Call SetUserRef(Mapping(ConnectionId).UserRef, FreeUser)
130     MapConnectionToUser = FreeUser
132     If FreeUser > LastUser Then
134         LastUser = FreeUser
        End If
        Exit Function
CheckDisconnectedUsers_Err:
136     Call TraceError(Err.Number, Err.Description, "modNetwork.MapConnectionToUser", Erl)
End Function

Public Sub ClearConnection(ByVal Connection)
    With Mapping(Connection)
        .TimeLastReset = 0
        .PacketCount = 0
        Call ClearUserRef(.UserRef)
    End With
End Sub
#Else
'DIRECT_PLAY

Public Mapping As New Dictionary


Public Function MapConnectionToUser(ByVal ConnectionID As Long) As Integer
On Error GoTo CheckDisconnectedUsers_Err:
        MapConnectionToUser = 0
        Dim FreeUser As Long
        FreeUser = NextOpenUser()
        If IsFeatureEnabled("debug_id_assign") Then
            Call LogError("Assign userId: " & FreeUser & " to connection: " & ConnectionID)
        End If
        
        If FreeUser < 0 Then
            If IsFeatureEnabled("debug_connections") Then
                Call LogError("Failed to find slot for new user, connection: " & connection & " LastUser: " & LastUser)
            End If
            KickConnection (ConnectionID)
            Exit Function
        End If
        
        If UserList(FreeUser).InUse Then
            Call LogError("Trying to use an user slot marked as in use! slot: " & FreeUser)
            FreeUser = NextOpenUser()
        End If
        

        Mapping.Add ConnectionID, FreeUser
        
        
        Dim cdetail As t_ConnectionInfo
        With cdetail
             .ConnIDValida = True
             .IP = "127.0.0.1"
             .ConnID = ConnectionID
             .OnConnectTimestamp = GetTickCount()
        End With
        UserList(FreeUser).ConnectionDetails = cdetail
 
        MapConnectionToUser = FreeUser
        If FreeUser > LastUser Then
           LastUser = FreeUser
        End If
        Exit Function
        
CheckDisconnectedUsers_Err:
        Call TraceError(Err.Number, Err.Description, "modNetwork.MapConnectionToUser", Erl)
End Function


'DirectPlay
Public Sub CheckDisconnectedUsers()
End Sub


Public Sub Listen(ByVal Limit As Long, ByVal Address As String, ByVal Service As String)
On Error GoTo listen_err
    Err.Clear
    Dim AppDesc As DPN_APPLICATION_DESC
    dpa.SetSP DP8SP_TCPIP
    dpa.AddComponentLong DPN_KEY_PORT, CLng(Service)
    Debug.Assert Err.Number = 0
    'Now set up the app description
    With AppDesc
        .guidApplication = AppGuid
        .lMaxPlayers = 800
        .SessionName = "vbArgentumServer"
        .lFlags = DPNSESSION_CLIENT_SERVER 'We must pass the client server flags if we are a server
    End With
    
    'Now set up our address value
    dpa.SetSP dps.GetServiceProvider(1).Guid
    
    Dim pInfo As DPN_PLAYER_INFO
    pInfo.Name = "server"
    pInfo.lInfoFlags = DPNINFO_NAME
    dps.SetServerInfo pInfo
    
    'Now start the server
    dps.Host AppDesc, dpa
    
    gfStarted = True
    Exit Sub
listen_err:
    If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.CreatePlayer", Erl)
    End If
End Sub
Public Sub close_not_logged_sockets_if_timeout()

End Sub

Public Sub CreatePlayer(ByVal lPlayerID As Long, fRejectMsg As Boolean)
On Error GoTo create_player_err
    Debug.Print "DPLAY > CreatePlayer ID:" & lPlayerID
    Err.Clear
    Dim dpPeer As DPN_PLAYER_INFO
    'Try to get player info, note this will fail if this is LOCAL_PLAYER aka server and that's okay
    dpPeer = dps.GetClientInfo(lPlayerID, 0)
    Dim addr As DirectPlay8Address
    Set addr = dps.GetClientAddress(lPlayerID)
    Dim port As Long
    port = addr.GetComponentLong(DPN_KEY_PORT)
    Dim IP As String
    IP = addr.GetComponentString(DPN_KEY_HOSTNAME)
    Dim Address As String
    Address = IP & ":" & port
    If IsFeatureEnabled("debug_connections") Then
        Call AddLogToCircularBuffer("OnServerConnect connecting new user on id: " & lPlayerID & " ip: " & Address)
    End If
    'Mapping maps ConnectionID -> UserIndex
    If Mapping.Exists(lPlayerID) Then
        'Check if the connection already has a UserIndex assigned, then raise an error and kick it out
        With Mapping.Item(lPlayerID)
          Call TraceError(Err.Number, Err.Description, "OnServerConnect Mapping(lPlayerID) > 0, connection: " & lPlayerID & " value: " & Mapping.Item(lPlayerID), Erl)
        End With
        Call KickConnection(lPlayerID)
        Exit Sub
    End If
       
    'Upong connection send PrepareConnected msg
    Call modSendData.SendToConnection(lPlayerID, PrepareConnected())
    Exit Sub

create_player_err:
    'If there is an error, handle it and kick out
    If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.CreatePlayer", Erl)
    End If
    Call KickConnection(lPlayerID)

End Sub

Public Sub DestroyPlayer(ByVal lPlayerID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
On Error GoTo OnServerClose_Err:
    Debug.Print "DPLAY > DestroyPlayer ID:" & lPlayerID
    Err.Clear
    If Mapping.Exists(lPlayerID) Then
        Dim user_index As Integer
        user_index = Mapping.Item(lPlayerID)
        With UserList(user_index)
            ' With UPD there is no way to send a msg after DirectPlay8Event.DestroyPlayer has been called so
            ' we set ConnIDValida to false to prevent sending msg and getting errors
            .ConnectionDetails.ConnIDValida = False
            If .flags.UserLogged Then
                Call Cerrar_Usuario(user_index)
            End If
        End With
        Mapping.Remove lPlayerID
    End If
    
    Exit Sub
OnServerClose_Err:
    If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.DestroyPlayer", Erl)
    End If
End Sub
Public Sub Receive(dpnotify As DxVBLibA.DPNMSG_RECEIVE, fRejectMsg As Boolean)
On Error GoTo receive_error:
    Err.Clear
    With dpnotify
            'Two cases here
            ' A) The client has not yet been assigned a UserIndex, so we receive a message which can be only LoginExistingChar/LoginNewChar
            ' B) The client has already got a valid UserIndex
            ' HandleIncomingData uses IsMissing(UserIndex) to test which case we are dealing with.
            If Mapping.Exists(.idSender) Then
                ' Case B: we got a UserIndex, retrieve it and pass it to HandleIncomingData
                Dim user_index As Integer
                user_index = Mapping.Item(dpnotify.idSender)
                Call Protocol.HandleIncomingData(.idSender, dpnotify, user_index)
            Else
                ' Case A: Client has not yet been assgined a UserIndex
                Call Protocol.HandleIncomingData(.idSender, dpnotify)
            End If
    End With
    Exit Sub
receive_error:
     If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.Receive", Erl)
    End If
End Sub
Public Sub Send(ByVal user_index As Long, ByRef writer As clsNetWriter)
On Error GoTo send_error:
    Debug.Assert user_index >= LBound(UserList) And user_index <= UBound(UserList)
    Err.Clear
    With UserList(user_index)
        Call SendToConnection(.ConnectionDetails.ConnID, writer)
    End With
    Exit Sub
send_error:
     If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.Send", Erl)
    End If
End Sub
Public Sub SendToConnection(ByVal ConnectionID As Long, ByRef writer As clsNetWriter)
On Error GoTo sendtoconnection_error:
    Err.Clear
    Call writer.Send(ConnectionID)
    Exit Sub
sendtoconnection_error:
     If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.SendToConnection", Erl)
    End If
End Sub
Public Sub Flush(ByVal user_index As Long)
    'Nothing
End Sub

Public Sub KickConnection(ByVal connection As Long)
On Error GoTo KickConnection_err:
    Err.Clear
    Call dps.DestroyClient(connection, 0, 0, 0)
    Exit Sub
KickConnection_err:
    If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.KickConnection", Erl)
    End If
End Sub

Public Sub Kick(ByVal connection As Long, Optional ByVal Message As String = vbNullString)
On Error GoTo Kick_ErrHandler:
    If IsFeatureEnabled("debug_connections") Then
        If (Message <> vbNullString) Then
            Call AddLogToCircularBuffer("Kick connection: " & connection & " reason: " & Message)
        Else
            Call AddLogToCircularBuffer("Kick connection: " & connection)
        End If
    End If
    If (Message <> vbNullString) Then
        Dim user_index As Integer
        
        If Mapping.Exists(connection) Then
            user_index = Mapping.Item(connection)
            Call Protocol_Writes.WriteErrorMsg(user_index, Message)
            If UserList(user_index).flags.UserLogged Then
                Call Cerrar_Usuario(user_index)
            End If
        End If
    End If
    KickConnection connection
    Exit Sub
    
Kick_ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Kick", Erl)
End Sub

Public Sub Disconnect()
    'Nothing
End Sub
Public Sub Poll()
    'Nothing to do here when using DPLAY
End Sub
#End If


