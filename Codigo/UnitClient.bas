Attribute VB_Name = "UnitClient"
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

Public Enum ClientTests
    TestInvalidBigPacketID = 100
    TestInvalidNegativePacketID
    TestWriteLoginExistingChar
    TestEnd
End Enum

Private NextTest As ClientTests

Private Client As Network.Client

Public connected As Boolean

Public Function GetNextTest() As ClientTests
GetNextTest = NextTest
End Function



Public Sub Init()
    NextTest = ClientTests.TestInvalidBigPacketID
    Set Client = New Network.Client
    Call Unit_Protocol_Writes.Initialize
    Call Client.Attach(AddressOf OnClientConnect, AddressOf OnClientClose, AddressOf OnClientSend, AddressOf OnClientRecv)
End Sub

Public Sub Connect(ByVal Address As String, ByVal Service As String)
    Debug.Assert Not Client Is Nothing
    Debug.Assert Address <> vbNullString And Service <> vbNullString
    Connected = False
    Call Client.Connect(Address, Service)
End Sub

Public Sub Disconnect()
    connected = False
    If Not Client Is Nothing Then
        Call Client.Close(True)
    End If
End Sub


Public Sub Poll()
    
    If (Client Is Nothing Or NextTest = TestEnd) Then
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

Private Sub fTestInvalidBigPacketID()
    Debug.Print "Running TestInvalidBigPacketID()"
    Call Unit_Protocol_Writes.WriteLong(ClientPacketID.PacketCount + 1)
End Sub

Private Sub fTestInvalidNegativePacketID()
    Debug.Print "Running TestInvalidNegativePacketID()"
    Call Unit_Protocol_Writes.WriteLong(-1)
End Sub
Private Sub fTestWriteLoginExcistingChar()
    Debug.Print "Running TestWriteLoginExcistingChar"
 Dim good_md5, md5 As String
    good_md5 = "a944087c826163c4ed658b1ea00594be"
    Call WriteLoginExistingChar(UnitTesting.encrypted_token, UnitTesting.public_key, _
        "zeno", 2, 0, 4, good_md5)
End Sub

Private Sub fTestWriteLoginNewChar()
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
   
    Select Case NextTest
        Case ClientTests.TestInvalidBigPacketID
            Call fTestInvalidBigPacketID
        Case ClientTests.TestInvalidNegativePacketID
            Call fTestInvalidNegativePacketID
        Case ClientTests.TestWriteLoginExistingChar
            Call fTestWriteLoginExcistingChar
        Case ClientTests.TestEnd
            Debug.Print "Executed all Client tests"
        Case Else
            Debug.Assert False
    End Select
    
End Sub

Private Sub OnClientClose(ByVal Code As Long)
On Error GoTo OnClientClose_Err:
    
    Call Unit_Protocol_Writes.Clear

    Debug.Print "UnitClient.OnClientClose"

    If NextTest <> TestEnd Then
        NextTest = NextTest + 1
        Call UnitClient.Connect("127.0.0.1", "7667")
    End If
    Exit Sub
    
OnClientClose_Err:
    
End Sub

Private Sub OnClientSend(ByVal Message As Network.Reader)

End Sub

Private Sub OnClientRecv(ByVal Message As Network.Reader)
    Dim Reader As Network.Reader
    Set Reader = Message
    Dim PacketId As Long
    PacketId = Reader.ReadInt16
    Debug.Print "UnitTesting recv PacketId " & PacketId
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
        Case ServerPacketID.ErrorMsg
            Call Unit_Protocol_Writes.HandleErrorMessageBox(Reader)
        Case Else
            While Reader.GetAvailable() > 0
                Reader.ReadBool
            Wend
    End Select
    
      
    Debug.Assert Message.GetAvailable() = 0
    '"You are in deep shit dude"
    
End Sub


