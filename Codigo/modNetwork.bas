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

Const TimeOutTime As Long = 40000
Private Server  As Network.Server
Private Time(2) As Single
Private Mapping() As Long

Public Sub Listen(ByVal Limit As Long, ByVal Address As String, ByVal Service As String)
    Set Server = New Network.Server
    ReDim Mapping(1 To MaxUsers) As Long
    
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

Public Sub Kick(ByVal Connection As Long, Optional ByVal Message As String = vbNullString)
On Error GoTo Kick_ErrHandler:
    If (Message <> vbNullString) Then
        Dim UserIndex As Long
        UserIndex = Mapping(Connection)
        If UserIndex > 0 Then
            Call Protocol_Writes.WriteErrorMsg(UserIndex, Message)
            If UserList(UserIndex).flags.UserLogged Then
                Call Cerrar_Usuario(UserIndex)
            End If
        End If
    End If
        
    Call Server.Flush(Connection)
    Call Server.Kick(Connection, True)
    
    Exit Sub
    
Kick_ErrHandler:
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
                        Call Kick(.ConnID, "Connection timeout")
                    End If
                End If
            End With
    Next i
End Sub
Private Sub OnServerConnect(ByVal Connection As Long, ByVal Address As String)
On Error GoTo OnServerConnect_Err:
  
    If IP_Blacklist.Exists(Address) <> 0 Then 'Busca si esta banneada la ip
        Call Kick(Connection, "Se te ha prohibido la entrada al servidor. Cod: #0003")
        Exit Sub
    End If
    
    If Connection <= MaxUsers Then
        'By Ladder y Wolfenstein
        Dim FreeUser As Long
        FreeUser = NextOpenUser()
                
        UserList(FreeUser).ConnIDValida = True
        UserList(FreeUser).IP = Address
        UserList(FreeUser).ConnID = Connection
        UserList(FreeUser).Counters.OnConnectTimestamp = GetTickCount()
        
        If FreeUser >= LastUser Then LastUser = FreeUser
        
        Mapping(Connection) = FreeUser
        
        Call WriteConnected(FreeUser)
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
    
    Dim UserIndex As Long
    UserIndex = Mapping(Connection)

    If UserIndex <= 0 Then Exit Sub
    'Es el mismo user al que está revisando el centinela??
    'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
    If Centinela.RevisandoUserIndex = UserIndex Then
        Call modCentinela.CentinelaUserLogout
    End If
    
    If UserList(UserIndex).flags.UserLogged Then
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    Else
        Call CloseSocket(UserIndex)
    End If
    
    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).ConnID = 0
    Mapping(Connection) = 0
    
    
    Exit Sub
    
OnServerClose_Err:
    Call ForcedClose(UserIndex, Connection)
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
    
    Dim UserIndex As Long
    UserIndex = Mapping(Connection)

    Call Protocol.HandleIncomingData(UserIndex, Message)
    
    Exit Sub
    
OnServerRecv_Err:
    Call Kick(Connection)
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnServerRecv", Erl)
End Sub

Private Sub ForcedClose(ByVal UserIndex As Integer, Connection As Long)
On Error GoTo ForcedClose_Err:
    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).ConnID = 0
    Call Server.Flush(Connection)
    Call Server.Kick(Connection, True)
    Mapping(Connection) = 0
    
ForcedClose_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.ForcedClose", Erl)
End Sub

Public Sub CheckDisconnectedUsers()
On Error GoTo CheckDisconnectedUsers_Err:
    
    Dim currentTime As Long
    Dim iUserIndex As Integer
    currentTime = GetTickCount()
    For iUserIndex = 1 To MaxUsers
        'Conexion activa? y es un usuario loggeado?
102     If UserList(iUserIndex).ConnIDValida = 0 And UserList(iUserIndex).flags.UserLogged And currentTime - UserList(iUserIndex).Counters.TimeLastReset > TimeOutTime Then
106         'mato los comercios seguros
110         If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
112             If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
114                 If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
116                     Call WriteConsoleMsg(UserList(iUserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", e_FontTypeNames.FONTTYPE_TALK)
118                     Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                    End If
                End If
120             Call FinComerciarUsu(iUserIndex)
            End If
122         Call Cerrar_Usuario(iUserIndex, True)
        End If
    
124 Next iUserIndex
CheckDisconnectedUsers_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.ForcedClose", Erl)
End Sub
