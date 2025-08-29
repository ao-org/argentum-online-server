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

#If UNIT_TEST = 1 Then

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
    On Error Goto GetNextTest_Err
GetNextTest = NextTest
    Exit Function
GetNextTest_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.GetNextTest", Erl)
End Function



Public Sub Init()
    On Error Goto Init_Err
    NextTest = ClientTests.TestInvalidBigPacketID
    Set Client = New Network.Client
    Call Unit_Protocol_Writes.Initialize
    Call Client.Attach(AddressOf OnClientConnect, AddressOf OnClientClose, AddressOf OnClientSend, AddressOf OnClientRecv)
    Exit Sub
Init_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.Init", Erl)
End Sub

Public Sub Connect(ByVal Address As String, ByVal Service As String)
    On Error Goto Connect_Err
    Debug.Assert Not Client Is Nothing
    Debug.Assert Address <> vbNullString And Service <> vbNullString
    Connected = False
    Call Client.Connect(Address, Service)
    Exit Sub
Connect_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.Connect", Erl)
End Sub

Public Sub Disconnect()
    On Error Goto Disconnect_Err
    connected = False
    If Not Client Is Nothing Then
        Call Client.Close(True)
    End If
    Exit Sub
Disconnect_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.Disconnect", Erl)
End Sub


Public Sub Poll()
    On Error Goto Poll_Err
    
    If (Client Is Nothing Or NextTest = TestEnd) Then
        Exit Sub
    End If
    
    Call Client.Flush
    Call Client.Poll
    Exit Sub
Poll_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.Poll", Erl)
End Sub

Public Sub Send(ByVal Buffer As Network.Writer)
    On Error Goto Send_Err
    If (connected) Then
        Call Client.Send(False, Buffer)
    End If
    
    Call Buffer.Clear
    Exit Sub
Send_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.Send", Erl)
End Sub

Private Sub fTestInvalidBigPacketID()
    On Error Goto fTestInvalidBigPacketID_Err
    Debug.Print "Running TestInvalidBigPacketID()"
    Call Unit_Protocol_Writes.WriteLong(ClientPacketID.PacketCount + 1)
    Exit Sub
fTestInvalidBigPacketID_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.fTestInvalidBigPacketID", Erl)
End Sub

Private Sub fTestInvalidNegativePacketID()
    On Error Goto fTestInvalidNegativePacketID_Err
    Debug.Print "Running TestInvalidNegativePacketID()"
    Call Unit_Protocol_Writes.WriteLong(-1)
    Exit Sub
fTestInvalidNegativePacketID_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.fTestInvalidNegativePacketID", Erl)
End Sub
Private Sub fTestWriteLoginExcistingChar()
    On Error Goto fTestWriteLoginExcistingChar_Err
    Debug.Print "Running TestWriteLoginExcistingChar"
 Dim good_md5, md5 As String
    good_md5 = "a944087c826163c4ed658b1ea00594be"
    Call WriteLoginExistingChar(UnitTesting.encrypted_token, UnitTesting.public_key, _
        "zeno", 2, 0, 4, good_md5)
    Exit Sub
fTestWriteLoginExcistingChar_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.fTestWriteLoginExcistingChar", Erl)
End Sub

Private Sub fTestWriteLoginNewChar()
    On Error Goto fTestWriteLoginNewChar_Err
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

    Exit Sub
fTestWriteLoginNewChar_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.fTestWriteLoginNewChar", Erl)
End Sub


Private Sub OnClientConnect()
    On Error Goto OnClientConnect_Err
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
    
    Exit Sub
OnClientConnect_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.OnClientConnect", Erl)
End Sub

Private Sub OnClientClose(ByVal Code As Long)
    On Error Goto OnClientClose_Err
On Error GoTo OnClientClose_Err:
    
    Call Unit_Protocol_Writes.Clear

    Debug.Print "UnitClient.OnClientClose"
    
    
    If NextTest <> TestEnd Then
        NextTest = NextTest + 1
        Call UnitClient.Connect("127.0.0.1", "7667")
    End If
    Exit Sub
    
OnClientClose_Err:
    
    Exit Sub
OnClientClose_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.OnClientClose", Erl)
End Sub

Private Sub OnClientSend(ByVal Message As Network.Reader)
    On Error Goto OnClientSend_Err

    Exit Sub
OnClientSend_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.OnClientSend", Erl)
End Sub

Private Sub OnClientRecv(ByVal Message As Network.Reader)
    On Error Goto OnClientRecv_Err
    Dim Reader As Network.Reader
    Set Reader = Message
    Dim PacketId As Long
    PacketId = Reader.ReadInt16
    Debug.Print "UnitTesting recv PacketId " & PacketId
    Select Case PacketId
        Case ServerPacketID.eConnected
            Debug.Print "ServerPacketID.econnected"
        Case ServerPacketID.elogged
            Debug.Print "ServerPacketID.elogged"
        Case ServerPacketID.eDisconnect
            Debug.Print "ServerPacketID.eDisconnect"
        Case ServerPacketID.eCharacterChange
            Debug.Print "CharacterChange"
            Call Unit_Protocol_Writes.HandleCharacterChange(Reader)
        Case ServerPacketID.eShowMessageBox
            Call Unit_Protocol_Writes.HandleShowMessageBox(Reader)
        Case ServerPacketID.eErrorMsg
            Call Unit_Protocol_Writes.HandleErrorMessageBox(Reader)
        Case Else
            While Reader.GetAvailable() > 0
                Reader.ReadBool
            Wend
    End Select
    
      
    Debug.Assert Message.GetAvailable() = 0
    '"You are in deep shit dude"
    
    Exit Sub
OnClientRecv_Err:
    Call TraceError(Err.Number, Err.Description, "UnitClient.OnClientRecv", Erl)
End Sub

#End If
