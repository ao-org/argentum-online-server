Attribute VB_Name = "modNetwork"
Option Explicit

Private Const WARNING_INFINITE_LOOP        As Long = 100
Private Const WARNING_INFINITE_LOOP_REPEAT As Long = 1000

Private Const TIME_RECV_FREQUENCY As Long = 5  ' In milliseconds
Private Const TIME_SEND_FREQUENCY As Long = 10 ' In milliseconds

Private Server  As Network.Server
Private Time(2) As Single

Public Sub Listen(ByVal Limit As Long, ByVal Address As String, ByVal Service As String)
    Set Server = New Network.Server
    
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

Public Sub Send(ByVal Connection As Long, ByVal Buffer As Network.Writer)
    Call Server.Send(Connection, False, Buffer)
End Sub

Public Sub Flush(ByVal Connection As Long)
    Call Server.Flush(Connection)
End Sub

Public Sub Kick(ByVal Connection As Long, Optional ByVal Message As String = vbNullString)
    If (Message <> vbNullString) Then
        Call Protocol_Writes.WriteErrorMsg(Connection, Message)
    End If
        
    Call Server.Flush(Connection)
    Call Server.Kick(Connection)
End Sub

Public Function GetTimeOfNextFlush() As Single
    GetTimeOfNextFlush = max(0, TIME_SEND_FREQUENCY - Time(1))
End Function

Private Sub OnServerConnect(ByVal Connection As Long, ByVal Address As String)
On Error GoTo OnServerConnect_Err:

    If Connection <= MaxUsers Then
    
        UserList(Connection).ConnIDValida = True
        UserList(Connection).IP = Address
        UserList(Connection).UUID = GenGUID()
        UserList(Connection).ID = -1
        UserList(Connection).AccountID = -1
        UserList(Connection).WaitingPacket = -1

        If IP_Blacklist.Exists(UserList(Connection).IP) <> 0 Then 'Busca si esta banneada la ip
            Call Kick(Connection, "Se te ha prohibido la entrada al servidor. Cod: #0003")
            Exit Sub
        End If

        If Connection > LastUser Then LastUser = Connection
        
        Call WriteConnected(Connection)
    Else
        Call Kick(Connection, "El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
    End If
    
    Exit Sub
    
OnServerConnect_Err:
    Call Kick(Connection)
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnServerConnect", Erl)
End Sub

Private Sub OnServerClose(ByVal Connection As Long)
On Error GoTo OnServerClose_Err:

    'Es el mismo user al que está revisando el centinela??
    'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
    If Centinela.RevisandoUserIndex = Connection Then
        Call modCentinela.CentinelaUserLogout
    End If
    
    If UserList(Connection).flags.UserLogged Then
        Call CloseSocketSL(Connection)
        Call Cerrar_Usuario(Connection)
    Else
        Call CloseSocket(Connection)
    End If
    
    UserList(Connection).ConnIDValida = False
    Exit Sub
    
OnServerClose_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnServerClose", Erl)
End Sub

Private Sub OnServerSend(ByVal Connection As Long, ByVal Message As Network.Reader)
On Error GoTo OnServerSend_Err:

#If AntiExternos = 1 Then
    Dim BytesRef() As Byte
    Call Message.getData(BytesRef) ' Is only a view of the buffer as a SafeArrayPtr ;-)

    Call Security.XorData(BytesRef, UBound(BytesRef) - 1, UserList(Connection).XorIndexOut)
#End If

    Exit Sub
    
OnServerSend_Err:
    Call Kick(Connection)
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnServerSend", Erl)
End Sub

Private Sub OnServerRecv(ByVal Connection As Long, ByVal Message As Network.Reader)
On Error GoTo OnServerRecv_Err:

#If AntiExternos = 1 Then
    Dim BytesRef() As Byte
    Call Message.getData(BytesRef) ' Is only a view of the buffer as a SafeArrayPtr ;-)

    Call Security.XorData(BytesRef, UBound(BytesRef) - 1, UserList(Connection).XorIndexIn)
#End If

    Call Protocol.HandleIncomingData(Connection, Message)
    
    Exit Sub
    
OnServerRecv_Err:
    Call Kick(Connection)
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnServerRecv", Erl)
End Sub

