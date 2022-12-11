Attribute VB_Name = "modNetwork"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

Private Const TIME_RECV_FREQUENCY As Long = 0  ' In milliseconds
Private Const TIME_SEND_FREQUENCY As Long = 0 ' In milliseconds

Private Server  As Network.Server
Private Time(2) As Single
Private Mapping() As t_UserReference
Public DisconnectTimeout As Long

Public Sub Listen(ByVal Limit As Long, ByVal Address As String, ByVal Service As String)
    Set Server = New Network.Server
    ReDim Mapping(1 To MaxUsers) As t_UserReference
    
    Call Server.Attach(AddressOf OnServerConnect, AddressOf OnServerClose, AddressOf OnServerSend, AddressOf OnServerRecv)
    
    Call Server.Listen(Limit, Address, Service)
End Sub

Public Sub Disconnect()
    Call Server.Close
End Sub

Public Sub Tick(ByVal Delta As Single)
    Time(0) = Time(0) + Delta
    Time(1) = Time(1) + Delta
    
    If (Time(0) >= TIME_RECV_FREQUENCY) Then
        Time(0) = 0
        
        Call Server.Poll
    End If
        
    If (Time(1) >= TIME_SEND_FREQUENCY) Then
        Time(1) = 0
        
        Call Server.Flush
    End If
End Sub

Public Sub Poll()
    Call Server.Poll
    Call Server.Flush
End Sub

Public Sub Send(ByVal UserIndex As Long, ByVal Buffer As Network.Writer)
    Call Server.Send(UserList(UserIndex).ConnID, False, Buffer)
End Sub

Public Sub Flush(ByVal UserIndex As Long)
    Call Server.Flush(UserList(UserIndex).ConnID)
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
    If (message <> vbNullString) Then
        Dim UserRef As t_UserReference
        UserRef = Mapping(Connection)
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
    Dim i As Integer
    For i = 1 To LastUser
         With UserList(i)
                If Not .flags.UserLogged And .ConnID > 0 Then
                    Dim Ticks As Long, Delta As Long
                    Ticks = GetTickCount
                    Delta = Ticks - .Counters.OnConnectTimestamp
                    If Delta > 3000 Then
                        If Mapping(.ConnID).ArrayIndex = i Then
                            Call Kick(.ConnID, "Connection timeout")
                        Else
                            .ConnID = 0
                            .ConnIDValida = False
                            Call TraceError(Err.Number, Err.Description, "trying to kick an invalid mapping", Erl)
                        End If
                    End If
                End If
            End With
    Next i
End Sub
Private Sub OnServerConnect(ByVal Connection As Long, ByVal Address As String)
On Error GoTo OnServerConnect_Err:
  
    If IsFeatureEnabled("debug_connections") Then
        Call AddLogToCircularBuffer("OnServerConnect connecting new user on id: " & Connection & " ip: " & Address)
    End If
    If IP_Blacklist.Exists(Address) <> 0 Then 'Busca si esta banneada la ip
        Call Kick(Connection, "Se te ha prohibido la entrada al servidor. Cod: #0003")
        Exit Sub
    End If
    If Mapping(Connection).ArrayIndex > 0 Then
        Call TraceError(Err.Number, Err.Description, "OnServerConnect Mapping(Connection) > 0, connection: " & Connection & " value: " & Mapping(Connection).ArrayIndex & ", is valid: " & IsValidUserRef(Mapping(Connection)), Erl)
    End If
    If Connection <= MaxUsers Then
        'By Ladder y Wolfenstein
        Dim FreeUser As Long
        FreeUser = NextOpenUser()
        If UserList(FreeUser).InUse Then
           Call LogError("Trying to use an user slot marked as in use! slot: " & FreeUser)
           FreeUser = NextOpenUser()
        End If
        UserList(FreeUser).ConnIDValida = True
        UserList(FreeUser).IP = Address
        UserList(FreeUser).ConnID = Connection
        UserList(FreeUser).Counters.OnConnectTimestamp = GetTickCount()
        
        If FreeUser >= LastUser Then LastUser = FreeUser
        Debug.Assert Not IsValidUserRef(Mapping(Connection))
        If Not SetUserRef(Mapping(Connection), FreeUser) Then
            Call TraceError(Err.Number, Err.Description, "OnServerConnect failed to map connection (" & Connection & ") to user: " & FreeUser, Erl)
        End If
        
        Call WriteConnected(Mapping(Connection).ArrayIndex)
    Else
        Call Kick(Connection, "El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
    End If
    
    Exit Sub
    
OnServerConnect_Err:
    Call Kick(Connection)
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnServerConnect", Erl)
End Sub

Private Sub OnServerClose(ByVal Connection As Long)
On Error GoTo OnServerClose_Err:
    
    Dim UserRef As t_UserReference
100    UserRef = Mapping(Connection)
102    If IsFeatureEnabled("debug_connections") Then
104        If UserRef.ArrayIndex > 0 Then
106            Call AddLogToCircularBuffer("OnServerClose disconnected user index: " & UserRef.ArrayIndex & " With connection id: " & Connection & " with name: " & UserList(UserRef.ArrayIndex).name & " and ip" & UserList(UserRef.ArrayIndex).IP)
108        Else
110            Call AddLogToCircularBuffer("OnServerClose disconnected user index: " & UserRef.ArrayIndex & " With connection id: " & Connection)
112        End If
114    End If
    
116    Debug.Assert IsValidUserRef(UserRef)
118    If IsValidUserRef(UserRef) Then
120        If UserList(UserRef.ArrayIndex).flags.UserLogged Then
122            Call CloseSocketSL(UserRef.ArrayIndex)
124            Call Cerrar_Usuario(UserRef.ArrayIndex)
126        Else
128            Call CloseSocket(UserRef.ArrayIndex)
130        End If
    
132        UserList(UserRef.ArrayIndex).ConnIDValida = False
134        UserList(UserRef.ArrayIndex).ConnID = 0
       End If
138    Call ClearUserRef(Mapping(Connection))

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
    UserRef = Mapping(Connection)

    Call Protocol.HandleIncomingData(UserRef.ArrayIndex, Message)
    
    Exit Sub
    
OnServerRecv_Err:
    Call Kick(Connection)
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnServerRecv", Erl)
End Sub

Private Sub ForcedClose(ByVal UserIndex As Integer, Connection As Long)
On Error GoTo ForcedClose_Err:
100     UserList(UserIndex).ConnIDValida = False
102     UserList(UserIndex).ConnID = 0
104     Call Server.Flush(Connection)
106     Call Server.Kick(Connection, True)
108     Call ClearUserRef(Mapping(Connection))
110     Call IncreaseVersionId(userIndex)
        Exit Sub
ForcedClose_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.ForcedClose", Erl)
End Sub

Public Sub CheckDisconnectedUsers()
On Error GoTo CheckDisconnectedUsers_Err:
    If DisconnectTimeout <= 0 Then
        Exit Sub
    End If
    Dim currentTime As Long
    Dim iUserIndex As Integer
    currentTime = GetTickCount()
    For iUserIndex = 1 To MaxUsers
        'Conexion activa? y es un usuario loggeado?
102     If UserList(iUserIndex).ConnIDValida = 0 And UserList(iUserIndex).flags.UserLogged And currentTime - UserList(iUserIndex).Counters.TimeLastReset > DisconnectTimeout Then
106         'mato los comercios seguros
110         If UserList(iUserIndex).ComUsu.DestUsu.ArrayIndex > 0 Then
112             If IsValidUserRef(UserList(iUserIndex).ComUsu.DestUsu) And UserList(UserList(iUserIndex).ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
114                 If UserList(UserList(iUserIndex).ComUsu.DestUsu.ArrayIndex).ComUsu.DestUsu.ArrayIndex = iUserIndex Then
116                     Call WriteConsoleMsg(UserList(iUserIndex).ComUsu.DestUsu.ArrayIndex, "Comercio cancelado por el otro usuario.", e_FontTypeNames.FONTTYPE_TALK)
118                     Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu.ArrayIndex)
                    End If
                End If
120             Call FinComerciarUsu(iUserIndex)
            End If
122         Call Cerrar_Usuario(iUserIndex, True)
        End If

124 Next iUserIndex
    Exit Sub
CheckDisconnectedUsers_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.CheckDisconnectedUsers", Erl)
End Sub

