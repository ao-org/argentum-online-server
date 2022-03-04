Attribute VB_Name = "UnitClient"
Option Explicit

Private Client As Network.Client

Public connected As Boolean
'public Public Enum ClientPacketID

Public Sub Connect(ByVal Address As String, ByVal Service As String)
    connected = False
    
    If (Address = vbNullString Or Service = vbNullString) Then
        Exit Sub
    End If
    
    Call Unit_Protocol_Writes.Initialize
    
    Set Client = New Network.Client
    Call Client.Attach(AddressOf OnClientConnect, AddressOf OnClientClose, AddressOf OnClientSend, AddressOf OnClientRecv)
    Call Client.Connect(Address, Service)
End Sub

Public Sub Disconnect()
    connected = False
    If Not Client Is Nothing Then
        Call Client.Close(True)
    End If
End Sub

Public Sub Poll()
    If (Client Is Nothing) Then
        Exit Sub
    End If
    
    Call Client.Flush
    Call Client.Poll
End Sub

Public Sub Send(ByVal Buffer As Network.Writer)
    If (Connected) Then
        Call Client.Send(False, Buffer)
    End If
    
    Call Buffer.Clear
End Sub

Private Sub OnClientConnect()
    Debug.Print ("UnitClient.OnClientConnect")
    connected = True
    Call Unit_Protocol_Writes.WriteLoginNewChar(UnitTesting.public_key, "morgolock2002@yahoo.com.ar")
    
End Sub

Private Sub OnClientClose(ByVal Code As Long)
    Call Unit_Protocol_Writes.Clear
    Debug.Print "OnClientClose " & Code
    Call Client.Close(True)
    connected = False
End Sub

Private Sub OnClientSend(ByVal Message As Network.Reader)

End Sub

Private Sub OnClientRecv(ByVal Message As Network.Reader)

End Sub


