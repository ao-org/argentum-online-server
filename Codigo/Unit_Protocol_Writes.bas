Attribute VB_Name = "Unit_Protocol_Writes"
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
Private Writer As Network.Writer

Public Function writer_is_nothing() As Boolean
    writer_is_nothing = Writer Is Nothing
End Function
Public Sub Initialize()
    Set Writer = New Network.Writer
End Sub

Public Sub Clear()
    Call Writer.Clear
End Sub


Public Sub WriteLoginExistingChar(ByVal encrypted_session_token As String, ByVal public_key As String, ByVal username As String, _
    ByVal app_major As Byte, ByVal app_minor As Byte, ByVal app_revision As Byte, ByVal md5 As String)
    
    Call Writer.WriteInt16(ClientPacketID.eLoginExistingChar)
    Call Writer.WriteString8(encrypted_session_token)
    Dim encrypted_username_b64 As String
    encrypted_username_b64 = AO20CryptoSysWrapper.ENCRYPT(cnvHexStrFromString(public_key), username)
    Call Writer.WriteString8(encrypted_username_b64)
    Call Writer.WriteInt8(app_major)
    Call Writer.WriteInt8(app_minor)
    Call Writer.WriteInt8(app_revision)
    Call Writer.WriteString8(md5)
     Call UnitClient.Send(Writer)
End Sub

Public Sub WriteLoginNewChar(ByVal public_key As String, ByVal username As String, _
    ByVal app_major As Byte, ByVal app_minor As Byte, ByVal app_revision As Byte, ByVal md5 As String, _
    ByVal race As Byte, ByVal gender As Byte, ByVal class As Byte, ByVal body As Byte, _
    ByVal head As Byte, ByVal home As Byte)
     
     Dim encrypted_username_b64 As String
     encrypted_username_b64 = AO20CryptoSysWrapper.ENCRYPT(cnvHexStrFromString(public_key), username)
     Call Writer.WriteInt16(ClientPacketID.eLoginNewChar)
     Call Writer.WriteString8(UnitTesting.encrypted_token)
     Call Writer.WriteString8(encrypted_username_b64)
     Call Writer.WriteInt8(App.Major)
     Call Writer.WriteInt8(App.Minor)
     Call Writer.WriteInt8(App.Revision)
     Call Writer.WriteString8(md5)
     Call Writer.WriteInt8(race)
     Call Writer.WriteInt8(gender)
     Call Writer.WriteInt8(class)
     Call Writer.WriteInt16(head)
     Call Writer.WriteInt8(home)
     Call UnitClient.Send(Writer)
End Sub

Public Sub WriteLong(ByVal value_to_send As Long)
    Call Writer.WriteInt16(value_to_send)
    Call UnitClient.Send(Writer)
End Sub
Public Sub HandleErrorMessageBox(ByRef Reader As Network.Reader)
    Dim mensaje As String
    mensaje = Reader.ReadString8()
    Debug.Print "HandleErrorMessageBox " & mensaje
End Sub
Public Sub HandleShowMessageBox(ByRef Reader As Network.Reader)
    Dim mensaje As String
    mensaje = Reader.ReadString8()
    Debug.Print "HandleShowMessageBox " & mensaje
End Sub

Public Sub HandleCharacterChange(ByRef Reader As Network.Reader)
    Dim charindex As Integer
    Dim TempInt   As Integer
    Dim headIndex As Integer
    Call Reader.ReadInt16
    TempInt = Reader.ReadInt16()
    headIndex = Reader.ReadInt16()
    Call Reader.ReadInt8
    TempInt = Reader.ReadInt16()
    TempInt = Reader.ReadInt16()
    TempInt = Reader.ReadInt16()
    Reader.ReadInt16
    Reader.ReadInt16
    Reader.ReadInt8
End Sub



