Attribute VB_Name = "UnitClient"
Option Explicit

Private Client As Network.Client

Public connected As Boolean


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
        Set Client = Nothing
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
    If (connected) Then
        Call Client.Send(False, Buffer)
    End If
    
    Call Buffer.Clear
End Sub

Private Sub TestInvalidPacketID()
    Call Unit_Protocol_Writes.WriteLong(ClientPacketID.PacketCount + 1)
    
End Sub

Private Sub TestWriteLoginExcistingChar()
 Dim good_md5, md5 As String
    good_md5 = "a944087c826163c4ed658b1ea00594be"
    Call WriteLoginExistingChar(UnitTesting.encrypted_token, UnitTesting.public_key, _
        "zeno", 2, 0, 4, good_md5)
End Sub

Private Sub TestWriteLoginNewChar()
    Dim good_md5, md5 As String
    good_md5 = "a944087c826163c4ed658b1ea00594be"
    md5 = good_md5
    Dim app_major, app_minor, app_revision, race, gender, Class, body, head, home As Byte
        
    app_major = 2
    app_minor = 0
    app_revision = 4
    race = 1
    gender = 1
    Class = 1
    body = 1
    head = 1
    home = 1
        
    Call Unit_Protocol_Writes.WriteLoginNewChar( _
        UnitTesting.public_key, UnitTesting.character_name, app_major, app_minor, app_revision, _
        md5, race, gender, Class, body, head, home)

End Sub



Private Sub OnClientConnect()
    Debug.Print ("UnitClient.OnClientConnect")
    connected = True
   
    'Call TestInvalidPacketID
    'Call TestWriteLoginNewCharFail
    Call TestWriteLoginExcistingChar
    'Call TestWriteLoginNewChar
    
    
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
    Dim Reader As Network.Reader
    Set Reader = Message
    Dim PacketId As Long:
    PacketId = Reader.ReadInt16
    Debug.Print "UnitTesting recv PacketId" & PacketId
    Select Case PacketId
        Case ServerPacketID.connected
            Debug.Print "ServerPacketID.connected"
            
        Case ServerPacketID.logged
            Debug.Print "ServerPacketID.logged"
            
            
        Case ServerPacketID.Disconnect
            Debug.Print "ServerPacketID.Disconnect"
            
            
        Case ServerPacketID.CharacterChange
            Call Unit_Protocol_Writes.HandleCharacterChange(Reader)
      
        Case ServerPacketID.ShowMessageBox
            Call Unit_Protocol_Writes.HandleShowMessageBox(Reader)
            
        Case Else
            While Reader.GetAvailable() > 0
                Reader.ReadBool
            Wend
    End Select
    
      
    Debug.Assert Message.GetAvailable() = 0
    '"You are in deep shit dude"
    
End Sub


