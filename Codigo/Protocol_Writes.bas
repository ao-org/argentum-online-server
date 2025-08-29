Attribute VB_Name = "Protocol_Writes"
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

#If DIRECT_PLAY = 0 Then
Private Writer  As Network.Writer

Public Sub InitializeAuxiliaryBuffer()
    On Error Goto InitializeAuxiliaryBuffer_Err
    Set writer = New Network.writer
    Exit Sub
InitializeAuxiliaryBuffer_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.InitializeAuxiliaryBuffer", Erl)
End Sub
    
Public Function GetWriterBuffer() As Network.Writer
    On Error Goto GetWriterBuffer_Err
    Set GetWriterBuffer = writer
    Exit Function
GetWriterBuffer_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.GetWriterBuffer", Erl)
End Function

#Else

Public writer As New clsNetWriter

#End If


#If PYMMO = 0 Then
Public Sub WriteAccountCharacterList(ByVal UserIndex As Integer, ByRef Personajes() As t_PersonajeCuenta, ByVal Count As Long)
    On Error Goto WriteAccountCharacterList_Err
        
        On Error GoTo WriteAccountCharacterList_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eAccountCharacterList)
        
        Call Writer.WriteInt(Count)
        
        Dim i As Long
        For i = 1 To Count
            With Personajes(i)
                Call Writer.WriteString8(.nombre)
                Call Writer.WriteInt(.cuerpo)
                Call Writer.WriteInt(.Cabeza)
                Call Writer.WriteInt(.clase)
                Call Writer.WriteInt(.Mapa)
                Call Writer.WriteInt(.posX)
                Call Writer.WriteInt(.posY)
                Call Writer.WriteInt(.nivel)
                Call Writer.WriteInt(.Status)
                Call Writer.WriteInt(.Casco)
                Call Writer.WriteInt(.Escudo)
                Call Writer.WriteInt(.Arma)
            End With
        Next i

102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteAccountCharacterList_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAccountCharacterList", Erl)
        
    Exit Sub
WriteAccountCharacterList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteAccountCharacterList", Erl)
End Sub
#End If
' \Begin: [Writes]

Public Function PrepareConnected()
    On Error Goto PrepareConnected_Err
On Error GoTo WriteConnected_Err
        Call Writer.WriteInt16(ServerPacketID.eConnected)
        
#If DEBUGGING = 1 Then
        Dim i As Integer
        Dim values(1 To 10) As Byte
        For i = LBound(values) To UBound(values)
         values(i) = i
        Next i
        Writer.WriteSafeArrayInt8 values
#End If

        Exit Function
WriteConnected_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteConnected", Erl)
    Exit Function
PrepareConnected_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareConnected", Erl)
End Function

''
' Writes the "Logged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoggedMessage(ByVal UserIndex As Integer, Optional ByVal newUser As Boolean = False)
    On Error Goto WriteLoggedMessage_Err
        
        On Error GoTo WriteLoggedMessage_Err
        
100     Call Writer.WriteInt16(ServerPacketID.elogged)
101     Call Writer.WriteBool(newUser)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteLoggedMessage_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLoggedMessage", Erl)
        
    Exit Sub
WriteLoggedMessage_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteLoggedMessage", Erl)
End Sub

Public Sub WriteHora(ByVal UserIndex As Integer)
    On Error Goto WriteHora_Err
        
        On Error GoTo WriteHora_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageHora())
        
        Exit Sub

WriteHora_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteHora", Erl)
        
    Exit Sub
WriteHora_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteHora", Erl)
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
    On Error Goto WriteRemoveAllDialogs_Err
        
        On Error GoTo WriteRemoveAllDialogs_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eRemoveDialogs)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteRemoveAllDialogs_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRemoveAllDialogs", Erl)
        
    Exit Sub
WriteRemoveAllDialogs_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteRemoveAllDialogs", Erl)
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
    On Error Goto WriteRemoveCharDialog_Err
        
        On Error GoTo WriteRemoveCharDialog_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageRemoveCharDialog( _
                CharIndex))
        
        Exit Sub

WriteRemoveCharDialog_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRemoveCharDialog", Erl)
        
    Exit Sub
WriteRemoveCharDialog_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteRemoveCharDialog", Erl)
End Sub

' Writes the "NavigateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNavigateToggle(ByVal UserIndex As Integer, ByVal NewState As Boolean)
    On Error Goto WriteNavigateToggle_Err
        
        On Error GoTo WriteNavigateToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eNavigateToggle)
        Call Writer.WriteBool(NewState)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteNavigateToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNavigateToggle", Erl)
        
    Exit Sub
WriteNavigateToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteNavigateToggle", Erl)
End Sub

Public Sub WriteNadarToggle(ByVal UserIndex As Integer, _
    On Error Goto WriteNadarToggle_Err
                            ByVal Puede As Boolean, _
                            Optional ByVal esTrajeCaucho As Boolean = False)
        
        On Error GoTo WriteNadarToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eNadarToggle)
102     Call Writer.WriteBool(Puede)
104     Call Writer.WriteBool(esTrajeCaucho)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteNadarToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNadarToggle", Erl)
        
    Exit Sub
WriteNadarToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteNadarToggle", Erl)
End Sub

Public Sub WriteEquiteToggle(ByVal UserIndex As Integer)
    On Error Goto WriteEquiteToggle_Err
        
        On Error GoTo WriteEquiteToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eEquiteToggle)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteEquiteToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteEquiteToggle", Erl)
        
    Exit Sub
WriteEquiteToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteEquiteToggle", Erl)
End Sub

Public Sub WriteVelocidadToggle(ByVal UserIndex As Integer)
    On Error Goto WriteVelocidadToggle_Err
        
        On Error GoTo WriteVelocidadToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eVelocidadToggle)
102     Call Writer.WriteReal32(UserList(UserIndex).Char.speeding)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteVelocidadToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteVelocidadToggle", Erl)
        
    Exit Sub
WriteVelocidadToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteVelocidadToggle", Erl)
End Sub

Public Sub WriteMacroTrabajoToggle(ByVal UserIndex As Integer, ByVal Activar As Boolean)
    On Error Goto WriteMacroTrabajoToggle_Err
        
        On Error GoTo WriteMacroTrabajoToggle_Err
        

100     If Not Activar Then
102         UserList(UserIndex).flags.TargetObj = 0 ' Sacamos el targer del objeto
104         UserList(UserIndex).flags.UltimoMensaje = 0
106         UserList(UserIndex).Counters.Trabajando = 0
108         UserList(UserIndex).flags.UsandoMacro = False
110         UserList(UserIndex).Trabajo.Target_X = 0
112         UserList(UserIndex).Trabajo.Target_Y = 0
114         UserList(UserIndex).Trabajo.TargetSkill = 0
            UserList(UserIndex).Trabajo.Cantidad = 0
            UserList(UserIndex).Trabajo.Item = 0
        Else
116         UserList(UserIndex).flags.UsandoMacro = True
        End If

118     Call Writer.WriteInt16(ServerPacketID.eMacroTrabajoToggle)
120     Call Writer.WriteBool(Activar)
122     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteMacroTrabajoToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMacroTrabajoToggle", Erl)
        
    Exit Sub
WriteMacroTrabajoToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteMacroTrabajoToggle", Erl)
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDisconnect(ByVal UserIndex As Integer, _
    On Error Goto WriteDisconnect_Err
                           Optional ByVal FullLogout As Boolean = False)
        
        On Error GoTo WriteDisconnect_Err
        
100     Call ClearAndSaveUser(UserIndex)
102     UserList(UserIndex).flags.YaGuardo = True

110     Call Writer.WriteInt16(ServerPacketID.eDisconnect)
        Call Writer.WriteBool(FullLogout)
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteDisconnect_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDisconnect", Erl)
        
    Exit Sub
WriteDisconnect_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteDisconnect", Erl)
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
    On Error Goto WriteCommerceEnd_Err
        
        On Error GoTo WriteCommerceEnd_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCommerceEnd)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCommerceEnd_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCommerceEnd", Erl)
        
    Exit Sub
WriteCommerceEnd_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCommerceEnd", Erl)
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankEnd(ByVal UserIndex As Integer)
    On Error Goto WriteBankEnd_Err
        
        On Error GoTo WriteBankEnd_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBankEnd)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteBankEnd_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBankEnd", Erl)
        
    Exit Sub
WriteBankEnd_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteBankEnd", Erl)
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
    On Error Goto WriteCommerceInit_Err
        
        On Error GoTo WriteCommerceInit_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCommerceInit)
102     Call Writer.WriteString8(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Name)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCommerceInit_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCommerceInit", Erl)
        
    Exit Sub
WriteCommerceInit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCommerceInit", Erl)
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankInit(ByVal UserIndex As Integer)
    On Error Goto WriteBankInit_Err
        
        On Error GoTo WriteBankInit_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBankInit)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteBankInit_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBankInit", Erl)
        
    Exit Sub
WriteBankInit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteBankInit", Erl)
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
    On Error Goto WriteUserCommerceInit_Err
        
        On Error GoTo WriteUserCommerceInit_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUserCommerceInit)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUserCommerceInit_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserCommerceInit", Erl)
        
    Exit Sub
WriteUserCommerceInit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUserCommerceInit", Erl)
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
    On Error Goto WriteUserCommerceEnd_Err
        
        On Error GoTo WriteUserCommerceEnd_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUserCommerceEnd)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUserCommerceEnd_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserCommerceEnd", Erl)
        
    Exit Sub
WriteUserCommerceEnd_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUserCommerceEnd", Erl)
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)
    On Error Goto WriteShowBlacksmithForm_Err
        
        On Error GoTo WriteShowBlacksmithForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowBlacksmithForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowBlacksmithForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowBlacksmithForm", Erl)
        
    Exit Sub
WriteShowBlacksmithForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowBlacksmithForm", Erl)
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)
    On Error Goto WriteShowCarpenterForm_Err
        
        On Error GoTo WriteShowCarpenterForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowCarpenterForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowCarpenterForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowCarpenterForm", Erl)
        
    Exit Sub
WriteShowCarpenterForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowCarpenterForm", Erl)
End Sub

Public Sub WriteShowAlquimiaForm(ByVal UserIndex As Integer)
    On Error Goto WriteShowAlquimiaForm_Err
        
        On Error GoTo WriteShowAlquimiaForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowAlquimiaForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowAlquimiaForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowAlquimiaForm", Erl)
        
    Exit Sub
WriteShowAlquimiaForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowAlquimiaForm", Erl)
End Sub

Public Sub WriteShowSastreForm(ByVal UserIndex As Integer)
    On Error Goto WriteShowSastreForm_Err
        
        On Error GoTo WriteShowSastreForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowSastreForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowSastreForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowSastreForm", Erl)
        
    Exit Sub
WriteShowSastreForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowSastreForm", Erl)
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)
    On Error Goto WriteNPCKillUser_Err
        
        On Error GoTo WriteNPCKillUser_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eNPCKillUser)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteNPCKillUser_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNPCKillUser", Erl)
        
    Exit Sub
WriteNPCKillUser_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteNPCKillUser", Erl)
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub Write_BlockedWithShieldUser(ByVal UserIndex As Integer)
    On Error Goto Write_BlockedWithShieldUser_Err
        
        On Error GoTo Write_BlockedWithShieldUser_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBlockedWithShieldUser)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

Write_BlockedWithShieldUser_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.Write_BlockedWithShieldUser", Erl)
        
    Exit Sub
Write_BlockedWithShieldUser_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.Write_BlockedWithShieldUser", Erl)
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub Write_BlockedWithShieldOther(ByVal UserIndex As Integer)
    On Error Goto Write_BlockedWithShieldOther_Err
        
        On Error GoTo Write_BlockedWithShieldOther_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBlockedWithShieldOther)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

Write_BlockedWithShieldOther_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.Write_BlockedWithShieldOther", Erl)
        
    Exit Sub
Write_BlockedWithShieldOther_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.Write_BlockedWithShieldOther", Erl)
End Sub


''
' Writes the "SafeModeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)
    On Error Goto WriteSafeModeOn_Err
        
        On Error GoTo WriteSafeModeOn_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eSafeModeOn)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteSafeModeOn_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSafeModeOn", Erl)
        
    Exit Sub
WriteSafeModeOn_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteSafeModeOn", Erl)
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)
    On Error Goto WriteSafeModeOff_Err
        
        On Error GoTo WriteSafeModeOff_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eSafeModeOff)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteSafeModeOff_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSafeModeOff", Erl)
        
    Exit Sub
WriteSafeModeOff_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteSafeModeOff", Erl)
End Sub

''
' Writes the "PartySafeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePartySafeOn(ByVal UserIndex As Integer)
    On Error Goto WritePartySafeOn_Err
        
        On Error GoTo WritePartySafeOn_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePartySafeOn)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WritePartySafeOn_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePartySafeOn", Erl)
        
    Exit Sub
WritePartySafeOn_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePartySafeOn", Erl)
End Sub

''
' Writes the "PartySafeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePartySafeOff(ByVal UserIndex As Integer)
    On Error Goto WritePartySafeOff_Err
        
        On Error GoTo WritePartySafeOff_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePartySafeOff)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WritePartySafeOff_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePartySafeOff", Erl)
        
    Exit Sub
WritePartySafeOff_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePartySafeOff", Erl)
End Sub

Public Sub WriteClanSeguro(ByVal UserIndex As Integer, ByVal estado As Boolean)
    On Error Goto WriteClanSeguro_Err
        
        On Error GoTo WriteClanSeguro_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eClanSeguro)
102     Call Writer.WriteBool(estado)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteClanSeguro_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteClanSeguro", Erl)
        
    Exit Sub
WriteClanSeguro_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteClanSeguro", Erl)
End Sub

Public Sub WriteSeguroResu(ByVal UserIndex As Integer, ByVal estado As Boolean)
    On Error Goto WriteSeguroResu_Err
        
        On Error GoTo WriteSeguroResu_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eSeguroResu)
102     Call Writer.WriteBool(estado)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteSeguroResu_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSeguroResu", Erl)
        
    Exit Sub
WriteSeguroResu_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteSeguroResu", Erl)
End Sub
Public Sub WriteLegionarySecure(ByVal UserIndex As Integer, ByVal Estado As Boolean)
    On Error Goto WriteLegionarySecure_Err
        
        On Error GoTo WriteLegionarySecure_Err
        
             Call Writer.WriteInt16(ServerPacketID.eLegionarySecure)
             Call Writer.WriteBool(Estado)
             Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteLegionarySecure_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLegionarySecure", Erl)
        
    Exit Sub
WriteLegionarySecure_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteLegionarySecure", Erl)
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)
    On Error Goto WriteCantUseWhileMeditating_Err
        
        On Error GoTo WriteCantUseWhileMeditating_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCantUseWhileMeditating)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCantUseWhileMeditating_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCantUseWhileMeditating", Erl)
        
    Exit Sub
WriteCantUseWhileMeditating_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCantUseWhileMeditating", Erl)
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateSta_Err
        
        On Error GoTo WriteUpdateSta_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUpdateSta)
102     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
104     Call modSendData.SendData(ToIndex, UserIndex)

        
        Exit Sub

WriteUpdateSta_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateSta", Erl)
        
    Exit Sub
WriteUpdateSta_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateSta", Erl)
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateMana_Err
        
        On Error GoTo WriteUpdateMana_Err
        
100     Call SendData(SendTarget.ToAdminsYDioses, UserList(userindex).GuildIndex, _
                PrepareMessageCharUpdateMAN(userindex))
10     Call SendData(SendTarget.ToClanArea, UserList(userindex).GuildIndex, _
                PrepareMessageCharUpdateMAN(UserIndex))
102     Call Writer.WriteInt16(ServerPacketID.eUpdateMana)
104     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
106     Call modSendData.SendData(ToIndex, UserIndex)

        
        Exit Sub

WriteUpdateMana_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateMana", Erl)
        
    Exit Sub
WriteUpdateMana_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateMana", Erl)
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateHP_Err
        'Call SendData(SendTarget.ToDiosesYclan, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
        
        On Error GoTo WriteUpdateHP_Err
        
100     Call SendData(SendTarget.ToAdminsYDioses, UserList(userindex).GuildIndex, _
                PrepareMessageCharUpdateHP(userindex))
101     Call SendData(SendTarget.ToClanArea, UserList(userindex).GuildIndex, _
                PrepareMessageCharUpdateHP(UserIndex))
        Call SendData(SendTarget.ToGroupButIndex, UserIndex, _
                PrepareMessageCharUpdateHP(UserIndex))
102     Call Writer.WriteInt16(ServerPacketID.eUpdateHP)
104     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
        Call Writer.WriteInt32(UserList(UserIndex).Stats.Shield)

106     Call modSendData.SendData(ToIndex, UserIndex)

        
        Exit Sub

WriteUpdateHP_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateHP", Erl)
        
    Exit Sub
WriteUpdateHP_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateHP", Erl)
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateGold_Err
        
        On Error GoTo WriteUpdateGold_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUpdateGold)
102     Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
103     Call Writer.WriteInt32(SvrConfig.GetValue("OroPorNivelBilletera"))
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUpdateGold_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateGold", Erl)
        
    Exit Sub
WriteUpdateGold_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateGold", Erl)
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateExp_Err
        
        On Error GoTo WriteUpdateExp_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUpdateExp)
102     Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUpdateExp_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateExp", Erl)
        
    Exit Sub
WriteUpdateExp_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateExp", Erl)
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer)
    On Error Goto WriteChangeMap_Err
        
        On Error GoTo WriteChangeMap_Err
        
100     Call Writer.WriteInt16(ServerPacketID.echangeMap)
102     Call Writer.WriteInt16(Map)
104     Call Writer.WriteInt16(MapInfo(Map).MapResource)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteChangeMap_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeMap", Erl)
        
    Exit Sub
WriteChangeMap_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteChangeMap", Erl)
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePosUpdate(ByVal UserIndex As Integer)
    On Error Goto WritePosUpdate_Err
        
        On Error GoTo WritePosUpdate_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePosUpdate)
102     Call Writer.WriteInt8(UserList(UserIndex).Pos.X)
104     Call Writer.WriteInt8(UserList(UserIndex).Pos.Y)
106     Call modSendData.SendData(ToIndex, UserIndex)
                
        If IsValidUserRef(UserList(UserIndex).flags.GMMeSigue) Then
            Call WritePosUpdateCharIndex(UserList(UserIndex).flags.GMMeSigue.ArrayIndex, UserList(UserIndex).pos.X, UserList(UserIndex).pos.y, UserList(UserIndex).Char.charindex)
        End If
        
        Exit Sub

WritePosUpdate_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePosUpdate", Erl)
        
    Exit Sub
WritePosUpdate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePosUpdate", Erl)
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePosUpdateCharIndex(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal charindex As Integer)
    On Error Goto WritePosUpdateCharIndex_Err
        
        On Error GoTo WritePosUpdateCharIndex_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePosUpdateUserChar)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
105     Call Writer.WriteInt16(charindex)
106     Call modSendData.SendData(ToIndex, UserIndex)

        
        Exit Sub

WritePosUpdateCharIndex_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePosUpdateCharIndex", Erl)
        
    Exit Sub
WritePosUpdateCharIndex_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePosUpdateCharIndex", Erl)
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePosUpdateChar(ByVal UserIndex As Integer, ByVal X As Byte, ByVal y As Byte, ByVal charindex As Integer)
    On Error Goto WritePosUpdateChar_Err
        
        On Error GoTo WritePosUpdateChar_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePosUpdateChar)
105     Call Writer.WriteInt16(charindex)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(y)
106     Call modSendData.SendData(ToIndex, UserIndex)

        
        Exit Sub

WritePosUpdateChar_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePosUpdateChar", Erl)
        
    Exit Sub
WritePosUpdateChar_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePosUpdateChar", Erl)
End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, _
    On Error Goto WriteNPCHitUser_Err
                           ByVal Target As e_PartesCuerpo, _
                           ByVal damage As Integer)
        
        On Error GoTo WriteNPCHitUser_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eNPCHitUser)
102     Call Writer.WriteInt8(Target)
104     Call Writer.WriteInt16(damage)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteNPCHitUser_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNPCHitUser", Erl)
        
    Exit Sub
WriteNPCHitUser_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteNPCHitUser", Erl)
End Sub

 
''
' Writes the "UserHittedByUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackerChar Char index of the user hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserHittedByUser(ByVal UserIndex As Integer, _
    On Error Goto WriteUserHittedByUser_Err
                                 ByVal Target As e_PartesCuerpo, _
                                 ByVal attackerChar As Integer, _
                                 ByVal damage As Integer)
        
        On Error GoTo WriteUserHittedByUser_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUserHittedByUser)
102     Call Writer.WriteInt16(attackerChar)
104     Call Writer.WriteInt8(Target)
106     Call Writer.WriteInt16(damage)
108     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUserHittedByUser_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserHittedByUser", Erl)
        
    Exit Sub
WriteUserHittedByUser_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUserHittedByUser", Erl)
End Sub

''
' Writes the "UserHittedUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackedChar Char index of the user hitted.
' @param    damage The number of HP lost by the oponent hitted.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserHittedUser(ByVal UserIndex As Integer, _
    On Error Goto WriteUserHittedUser_Err
                               ByVal Target As e_PartesCuerpo, _
                               ByVal attackedChar As Integer, _
                               ByVal damage As Integer)
        
        On Error GoTo WriteUserHittedUser_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUserHittedUser)
102     Call Writer.WriteInt16(attackedChar)
104     Call Writer.WriteInt8(Target)
106     Call Writer.WriteInt16(damage)
108     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUserHittedUser_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserHittedUser", Erl)
        
    Exit Sub
WriteUserHittedUser_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUserHittedUser", Erl)
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChatOverHead(ByVal UserIndex As Integer, _
    On Error Goto WriteChatOverHead_Err
                             ByVal chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal Color As Long)
        
        On Error GoTo WriteChatOverHead_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageChatOverHead(chat, _
                charindex, Color, , UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))
        
        Exit Sub

WriteChatOverHead_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChatOverHead", Erl)
        
    Exit Sub
WriteChatOverHead_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteChatOverHead", Erl)
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLocaleChatOverHead(ByVal UserIndex As Integer, _
    On Error Goto WriteLocaleChatOverHead_Err
                             ByVal ChatId As Integer, _
                             ByVal Params As String, _
                             ByVal charindex As Integer, _
                             ByVal Color As Long)
        
        On Error GoTo WriteLocaleChatOverHead_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareLocaleChatOverHead(ChatId, Params, _
                charindex, Color, , UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        
        Exit Sub

WriteLocaleChatOverHead_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLocaleChatOverHead", Erl)
        
    Exit Sub
WriteLocaleChatOverHead_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteLocaleChatOverHead", Erl)
End Sub
Public Function PrepareLocalizedChatOverHead(ByVal msgId As Integer, _
    On Error Goto PrepareLocalizedChatOverHead_Err
                                             ByVal charIndex As Integer, _
                                             ByVal color As Long, _
                                             ParamArray args() As Variant) As String
    Dim finalText As String
    Dim i As Long

    finalText = "LOCMSG*" & msgId & "*"

    For i = LBound(args) To UBound(args)
        If i > LBound(args) Then finalText = finalText & "¬"
        finalText = finalText & CStr(args(i))
    Next

    PrepareLocalizedChatOverHead = PrepareMessageChatOverHead(finalText, charIndex, color)
    Exit Function
PrepareLocalizedChatOverHead_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareLocalizedChatOverHead", Erl)
End Function



Public Sub WriteTextOverChar(ByVal UserIndex As Integer, _
    On Error Goto WriteTextOverChar_Err
                             ByVal chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal Color As Long)
        
        On Error GoTo WriteTextOverChar_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextOverChar(chat, _
                CharIndex, Color))
        
        Exit Sub

WriteTextOverChar_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTextOverChar", Erl)
        
    Exit Sub
WriteTextOverChar_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteTextOverChar", Erl)
End Sub

Public Sub WriteTextOverTile(ByVal UserIndex As Integer, _
    On Error Goto WriteTextOverTile_Err
                             ByVal chat As String, _
                             ByVal X As Integer, _
                             ByVal Y As Integer, _
                             ByVal Color As Long)
        
        On Error GoTo WriteTextOverTile_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextOverTile(chat, X, _
                Y, Color))
        
        Exit Sub

WriteTextOverTile_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTextOverTile", Erl)
        
    Exit Sub
WriteTextOverTile_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteTextOverTile", Erl)
End Sub

Public Sub WriteTextCharDrop(ByVal UserIndex As Integer, _
    On Error Goto WriteTextCharDrop_Err
                             ByVal chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal Color As Long)
        
        On Error GoTo WriteTextCharDrop_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextCharDrop(chat, _
                CharIndex, Color))
        
        Exit Sub

WriteTextCharDrop_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTextCharDrop", Erl)
        
    Exit Sub
WriteTextCharDrop_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteTextCharDrop", Erl)
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, _
    On Error Goto WriteConsoleMsg_Err
                           ByVal chat As String, _
                           Optional ByVal FontIndex As e_FontTypeNames = FONTTYPE_INFO)
        
        On Error GoTo WriteConsoleMsg_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageConsoleMsg(chat, _
                FontIndex))
        
        Exit Sub

WriteConsoleMsg_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteConsoleMsg", Erl)
        
    Exit Sub
WriteConsoleMsg_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteConsoleMsg", Erl)
End Sub

Public Sub WriteLocaleMsg(ByVal UserIndex As Integer, _
    On Error Goto WriteLocaleMsg_Err
                          ByVal ID As Integer, _
                          ByVal FontIndex As e_FontTypeNames, _
                          Optional ByVal strExtra As String = vbNullString)
        
        On Error GoTo WriteLocaleMsg_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageLocaleMsg(ID, strExtra, _
                FontIndex))
        
        Exit Sub

WriteLocaleMsg_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLocaleMsg", Erl)
        
    Exit Sub
WriteLocaleMsg_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteLocaleMsg", Erl)
End Sub







''
' Writes the "GuildChat" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildChat(ByVal UserIndex As Integer, _
    On Error Goto WriteGuildChat_Err
                          ByVal chat As String, _
                          ByVal Status As Byte)
    
    On Error GoTo WriteGuildChat_Err
    
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageGuildChat(chat, Status))
    
    Exit Sub

WriteGuildChat_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildChat", Erl)
    
    Exit Sub
WriteGuildChat_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteGuildChat", Erl)
End Sub


''
' Writes the "ShowMessageBox" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal MessageId As Integer, Optional ByVal strExtra As String = vbNullString)
    On Error Goto WriteShowMessageBox_Err
    On Error GoTo WriteShowMessageBox_Err

    Call Writer.WriteInt16(ServerPacketID.eShowMessageBox)
    Call Writer.WriteInt16(MessageId)
    Call Writer.WriteString8(strExtra) ' Enviás los valores dinámicos si hay

    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub

WriteShowMessageBox_Err:
    Call Writer.Clear

    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowMessageBox", Erl)

    Exit Sub
WriteShowMessageBox_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowMessageBox", Erl)
End Sub


Public Function PrepareShowMessageBox(ByVal Message As String)
    On Error Goto PrepareShowMessageBox_Err
        On Error GoTo WriteShowMessageBox_Err
100     Call Writer.WriteInt16(ServerPacketID.eShowMessageBox)
102     Call Writer.WriteString8(Message)
        Exit Function
WriteShowMessageBox_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowMessageBox", Erl)
    Exit Function
PrepareShowMessageBox_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareShowMessageBox", Erl)
End Function

Public Sub WriteMostrarCuenta(ByVal UserIndex As Integer)
    On Error Goto WriteMostrarCuenta_Err
        
        On Error GoTo WriteMostrarCuenta_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eMostrarCuenta)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteMostrarCuenta_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMostrarCuenta", Erl)
        
    Exit Sub
WriteMostrarCuenta_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteMostrarCuenta", Erl)
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
    On Error Goto WriteUserIndexInServer_Err
        
        On Error GoTo WriteUserIndexInServer_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUserIndexInServer)
102     Call Writer.WriteInt16(UserIndex)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUserIndexInServer_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserIndexInServer", Erl)
        
    Exit Sub
WriteUserIndexInServer_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUserIndexInServer", Erl)
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
    On Error Goto WriteUserCharIndexInServer_Err
        
        On Error GoTo WriteUserCharIndexInServer_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUserCharIndexInServer)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUserCharIndexInServer_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserCharIndexInServer", Erl)
        
    Exit Sub
WriteUserCharIndexInServer_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUserCharIndexInServer", Erl)
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    cart index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal Heading As e_Heading, ByVal charindex As Integer, _
    On Error Goto WriteCharacterCreate_Err
                                ByVal x As Byte, ByVal y As Byte, ByVal weapon As Integer, ByVal shield As Integer, ByVal Cart As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, ByVal name As String, ByVal Status As Byte, ByVal privileges As Byte, ByVal ParticulaFx As Byte, _
                                ByVal Head_Aura As String, ByVal Arma_Aura As String, ByVal Body_Aura As String, ByVal DM_Aura As String, ByVal RM_Aura As String, _
                                ByVal Otra_Aura As String, ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Byte, ByVal appear As Byte, _
                                ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, _
                                ByVal UserMaxHp As Long, ByVal UserMinMAN As Long, ByVal UserMaxMAN As Long, ByVal Simbolo As Byte, Optional ByVal Idle As Boolean = False, _
                                Optional ByVal Navegando As Boolean = False, Optional ByVal tipoUsuario As e_TipoUsuario = 0, _
                                Optional ByVal TeamCaptura As Byte = 0, Optional ByVal TieneBandera As Byte = 0, Optional ByVal AnimAtaque1 As Integer = 0)
        
        On Error GoTo WriteCharacterCreate_Err
        
100 Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterCreate(Body, Head, _
            Heading, charindex, x, y, weapon, shield, Cart, FX, FXLoops, helmet, name, Status, _
            privileges, ParticulaFx, Head_Aura, Arma_Aura, Body_Aura, DM_Aura, RM_Aura, _
            Otra_Aura, Escudo_Aura, speeding, EsNPC, appear, group_index, _
            clan_index, clan_nivel, UserMinHp, UserMaxHp, UserMinMAN, UserMaxMAN, Simbolo, _
            Idle, Navegando, tipoUsuario, TeamCaptura, TieneBandera, AnimAtaque1))
        
        Exit Sub

WriteCharacterCreate_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterCreate", Erl)
        
    Exit Sub
WriteCharacterCreate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCharacterCreate", Erl)
End Sub

Public Sub WriteCharacterUpdateFlag(ByVal UserIndex As Integer, ByVal Flag As Byte, ByVal charindex As Integer)
    On Error Goto WriteCharacterUpdateFlag_Err
   On Error GoTo WriteCharacterUpdateFlag_Err
        
100 Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageUpdateFlag(Flag, charindex))
        
        Exit Sub

WriteCharacterUpdateFlag_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterUpdateFlag", Erl)
        
    Exit Sub
WriteCharacterUpdateFlag_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCharacterUpdateFlag", Erl)
End Sub


Public Sub WriteForceCharMove(ByVal UserIndex As Integer, ByVal Direccion As e_Heading)
    On Error Goto WriteForceCharMove_Err
        
        On Error GoTo WriteForceCharMove_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageForceCharMove(Direccion))
        
        Exit Sub

WriteForceCharMove_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteForceCharMove", Erl)
        
    Exit Sub
WriteForceCharMove_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteForceCharMove", Erl)
End Sub

Public Sub WriteForceCharMoveSiguiendo(ByVal UserIndex As Integer, ByVal Direccion As e_Heading)
    On Error Goto WriteForceCharMoveSiguiendo_Err
        
        On Error GoTo WriteForceCharMoveSiguiendo_Err
        
98      Call Writer.WriteInt16(ServerPacketID.eForceCharMoveSiguiendo)
100     Call Writer.WriteInt8(Direccion)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteForceCharMoveSiguiendo_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteForceCharMoveSiguiendo", Erl)
        
    Exit Sub
WriteForceCharMoveSiguiendo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteForceCharMoveSiguiendo", Erl)
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCharacterChange(ByVal UserIndex As Integer, _
    On Error Goto WriteCharacterChange_Err
                                ByVal Body As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As e_Heading, _
                                ByVal CharIndex As Integer, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal Cart As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                Optional ByVal Idle As Boolean = False, _
                                Optional ByVal Navegando As Boolean = False)
        
        On Error GoTo WriteCharacterChange_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterChange(Body, _
                head, Heading, charindex, weapon, shield, Cart, FX, FXLoops, helmet, Idle, _
                Navegando))
        
        Exit Sub

WriteCharacterChange_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterChange", Erl)
        
    Exit Sub
WriteCharacterChange_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCharacterChange", Erl)
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteObjectCreate(ByVal UserIndex As Integer, _
    On Error Goto WriteObjectCreate_Err
                             ByVal ObjIndex As Integer, _
                             ByVal amount As Integer, _
                             ByVal X As Byte, _
                             ByVal Y As Byte)
        
        On Error GoTo WriteObjectCreate_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageObjectCreate(ObjIndex, _
                amount, x, y, ObjData(ObjIndex).ElementalTags))
        
        Exit Sub

WriteObjectCreate_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteObjectCreate", Erl)
        
    Exit Sub
WriteObjectCreate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteObjectCreate", Erl)
End Sub

Public Sub WriteUpdateTrapState(ByVal UserIndex As Integer, State As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error Goto WriteUpdateTrapState_Err
On Error GoTo WriteUpdateTrapState_Err
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareTrapUpdate(State, x, y))
        Exit Sub

WriteUpdateTrapState_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateTrapState", Erl)
    Exit Sub
WriteUpdateTrapState_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateTrapState", Erl)
End Sub

Public Sub WriteParticleFloorCreate(ByVal UserIndex As Integer, _
    On Error Goto WriteParticleFloorCreate_Err
                                    ByVal Particula As Integer, _
                                    ByVal ParticulaTime As Integer, _
                                    ByVal Map As Integer, _
                                    ByVal X As Byte, _
                                    ByVal Y As Byte)
        
        On Error GoTo WriteParticleFloorCreate_Err
        

100     If Particula = 0 Then
102         Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageParticleFXToFloor( _
                    X, Y, Particula, ParticulaTime))
        End If

        
        Exit Sub

WriteParticleFloorCreate_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteParticleFloorCreate", Erl)
        
    Exit Sub
WriteParticleFloorCreate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteParticleFloorCreate", Erl)
End Sub

Public Sub WriteLightFloorCreate(ByVal UserIndex As Integer, _
    On Error Goto WriteLightFloorCreate_Err
                                 ByVal LuzColor As Long, _
                                 ByVal Rango As Byte, _
                                 ByVal Map As Integer, _
                                 ByVal X As Byte, _
                                 ByVal Y As Byte)
        
        On Error GoTo WriteLightFloorCreate_Err
        
100     MapData(Map, X, Y).Luz.Color = LuzColor
102     MapData(Map, X, Y).Luz.Rango = Rango

104     If Rango = 0 Then
106         Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageLightFXToFloor(X, _
                    Y, LuzColor, Rango))
        End If

        
        Exit Sub

WriteLightFloorCreate_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLightFloorCreate", Erl)
        
    Exit Sub
WriteLightFloorCreate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteLightFloorCreate", Erl)
End Sub

Public Sub WriteFxPiso(ByVal UserIndex As Integer, _
    On Error Goto WriteFxPiso_Err
                       ByVal GrhIndex As Integer, _
                       ByVal X As Byte, _
                       ByVal Y As Byte)
        
        On Error GoTo WriteFxPiso_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageFxPiso(GrhIndex, X, Y))
        
        Exit Sub

WriteFxPiso_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFxPiso", Erl)
        
    Exit Sub
WriteFxPiso_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteFxPiso", Erl)
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
    On Error Goto WriteObjectDelete_Err
        
        On Error GoTo WriteObjectDelete_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageObjectDelete(X, Y))
        
        Exit Sub

WriteObjectDelete_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteObjectDelete", Erl)
        
    Exit Sub
WriteObjectDelete_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteObjectDelete", Erl)
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub Write_BlockPosition(ByVal UserIndex As Integer, _
    On Error Goto Write_BlockPosition_Err
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByVal Blocked As Byte)
        
        On Error GoTo Write_BlockPosition_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBlockPosition)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt8(Blocked)
108     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

Write_BlockPosition_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.Write_BlockPosition", Erl)
        
    Exit Sub
Write_BlockPosition_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.Write_BlockPosition", Erl)
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePlayMidi(ByVal UserIndex As Integer, _
    On Error Goto WritePlayMidi_Err
                         ByVal midi As Byte, _
                         Optional ByVal loops As Integer = -1)
        
        On Error GoTo WritePlayMidi_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePlayMidi(midi, loops))
        
        Exit Sub

WritePlayMidi_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePlayMidi", Erl)
        
    Exit Sub
WritePlayMidi_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePlayMidi", Erl)
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePlayWave(ByVal UserIndex As Integer, _
    On Error Goto WritePlayWave_Err
                         ByVal wave As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte, _
                         Optional ByVal CancelLastWave As Byte = 0, _
                         Optional ByVal Localize As Byte = 0)
                         
        On Error GoTo WritePlayWave_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePlayWave(wave, x, y, CancelLastWave, Localize))
        
        Exit Sub

WritePlayWave_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePlayWave", Erl)
        
    Exit Sub
WritePlayWave_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePlayWave", Erl)
End Sub
Public Sub WritePlayWaveStep(ByVal UserIndex As Integer, _
    On Error Goto WritePlayWaveStep_Err
                         ByVal CharIndex As Integer, _
                         ByVal grh As Long, _
                         ByVal grh2 As Long, _
                         ByVal distance As Byte, _
                         ByVal balance As Integer, _
                         ByVal step As Boolean)
        
        On Error GoTo WritePlayWaveStep_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePlayWaveStep)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt32(grh)
106     Call Writer.WriteInt32(grh2)
108     Call Writer.WriteInt8(distance)
109     Call Writer.WriteInt16(balance)
110     Call Writer.WriteBool(step)
132     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WritePlayWaveStep_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePlayWaveStep", Erl)
        
    Exit Sub
WritePlayWaveStep_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePlayWaveStep", Erl)
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)
    On Error Goto WriteGuildList_Err
        
        On Error GoTo WriteGuildList_Err
        

        Dim Tmp As String

        Dim i   As Long

100     Call Writer.WriteInt16(ServerPacketID.eguildList)

        ' Prepare guild name's list
102     For i = LBound(guildList()) To UBound(guildList())
104         Tmp = Tmp & guildList(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteGuildList_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildList", Erl)
        
    Exit Sub
WriteGuildList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteGuildList", Erl)
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAreaChanged(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
    On Error Goto WriteAreaChanged_Err
        
        On Error GoTo WriteAreaChanged_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eAreaChanged)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteAreaChanged_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAreaChanged", Erl)
        
    Exit Sub
WriteAreaChanged_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteAreaChanged", Erl)
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePauseToggle(ByVal UserIndex As Integer)
    On Error Goto WritePauseToggle_Err
        
        On Error GoTo WritePauseToggle_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePauseToggle())
        
        Exit Sub

WritePauseToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePauseToggle", Erl)
        
    Exit Sub
WritePauseToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePauseToggle", Erl)
End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRainToggle(ByVal UserIndex As Integer)
    On Error Goto WriteRainToggle_Err
        
        On Error GoTo WriteRainToggle_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageRainToggle())
        
        Exit Sub

WriteRainToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRainToggle", Erl)
        
    Exit Sub
WriteRainToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteRainToggle", Erl)
End Sub

Public Sub WriteNubesToggle(ByVal UserIndex As Integer)
    On Error Goto WriteNubesToggle_Err
        
        On Error GoTo WriteNubesToggle_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageNieblandoToggle( _
                IntensidadDeNubes))
        
        Exit Sub

WriteNubesToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNubesToggle", Erl)
        
    Exit Sub
WriteNubesToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteNubesToggle", Erl)
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreateFX(ByVal UserIndex As Integer, _
    On Error Goto WriteCreateFX_Err
                         ByVal CharIndex As Integer, _
                         ByVal FX As Integer, _
                         ByVal FXLoops As Integer)

        'Writes the "CreateFX" message to the given user's outgoing data buffer


        On Error GoTo WriteCreateFX_Err

100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCreateFX(CharIndex, FX, _
                FXLoops, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))

        Exit Sub

WriteCreateFX_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCreateFX", Erl)

    Exit Sub
WriteCreateFX_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCreateFX", Erl)
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateUserStats_Err

        'Writes the "UpdateUserStats" message to the given user's outgoing data buffer

        On Error GoTo WriteUpdateUserStats_Err

100     Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
                PrepareMessageCharUpdateHP(UserIndex))
102     Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
                PrepareMessageCharUpdateMAN(UserIndex))
104     Call Writer.WriteInt16(ServerPacketID.eUpdateUserStats)
106     Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxHp)
108     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
109     Call Writer.WriteInt32(UserList(UserIndex).Stats.Shield)
110     Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxMAN)
112     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
114     Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxSta)
116     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
118     Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
119     Call Writer.WriteInt32(SvrConfig.GetValue("OroPorNivelBilletera"))
120     Call Writer.WriteInt8(UserList(UserIndex).Stats.ELV)
122     Call Writer.WriteInt32(ExpLevelUp(UserList(UserIndex).Stats.ELV))
124     Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
126     Call Writer.WriteInt8(UserList(UserIndex).clase)
128     Call modSendData.SendData(ToIndex, UserIndex)


        Exit Sub

WriteUpdateUserStats_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateUserStats", Erl)

    Exit Sub
WriteUpdateUserStats_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateUserStats", Erl)
End Sub

Public Sub WriteUpdateUserKey(ByVal UserIndex As Integer, _
    On Error Goto WriteUpdateUserKey_Err
                              ByVal Slot As Integer, _
                              ByVal Llave As Integer)
        
        On Error GoTo WriteUpdateUserKey_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUpdateUserKey)
102     Call Writer.WriteInt16(Slot)
104     Call Writer.WriteInt16(Llave)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUpdateUserKey_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateUserKey", Erl)
        
    Exit Sub
WriteUpdateUserKey_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateUserKey", Erl)
End Sub

' Actualiza el indicador de daño mágico
Public Sub WriteUpdateDM(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateDM_Err
        
        On Error GoTo WriteUpdateDM_Err
        

        Dim Valor As Integer

100     With UserList(UserIndex).Invent

            ' % daño mágico del arma
102         If .WeaponEqpObjIndex > 0 Then
104             Valor = Valor + ObjData(.WeaponEqpObjIndex).MagicDamageBonus
            End If

            ' % daño mágico del anillo
106         If .DañoMagicoEqpObjIndex > 0 Then
108             Valor = Valor + ObjData(.DañoMagicoEqpObjIndex).MagicDamageBonus
            End If

110         Call Writer.WriteInt16(ServerPacketID.eUpdateDM)
112         Call Writer.WriteInt16(Valor)
        End With

114     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUpdateDM_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateDM", Erl)
        
    Exit Sub
WriteUpdateDM_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateDM", Erl)
End Sub

' Actualiza el indicador de resistencia mágica
Public Sub WriteUpdateRM(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateRM_Err
        
        On Error GoTo WriteUpdateRM_Err
        

        Dim Valor As Integer

100     With UserList(UserIndex).Invent

            ' Resistencia mágica de la armadura
102         If .ArmourEqpObjIndex > 0 Then
104             Valor = Valor + ObjData(.ArmourEqpObjIndex).ResistenciaMagica
            End If

            ' Resistencia mágica del anillo
106         If .ResistenciaEqpObjIndex > 0 Then
108             Valor = Valor + ObjData(.ResistenciaEqpObjIndex).ResistenciaMagica
            End If

            ' Resistencia mágica del escudo
110         If .EscudoEqpObjIndex > 0 Then
112             Valor = Valor + ObjData(.EscudoEqpObjIndex).ResistenciaMagica
            End If

            ' Resistencia mágica del casco
114         If .CascoEqpObjIndex > 0 Then
116             Valor = Valor + ObjData(.CascoEqpObjIndex).ResistenciaMagica
            End If

118         Valor = Valor + 100 * ModClase(UserList(UserIndex).clase).ResistenciaMagica
120         Call Writer.WriteInt16(ServerPacketID.eUpdateRM)
122         Call Writer.WriteInt16(Valor)
        End With

124     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUpdateRM_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateRM", Erl)
        
    Exit Sub
WriteUpdateRM_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateRM", Erl)
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As e_Skill, Optional ByVal CasteaArea As Boolean = False, Optional ByVal Radio As Byte = 0)
    On Error Goto WriteWorkRequestTarget_Err
        
        On Error GoTo WriteWorkRequestTarget_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eWorkRequestTarget)
102     Call Writer.WriteInt8(Skill)
        Call Writer.WriteBool(CasteaArea)
        Call Writer.WriteInt8(Radio)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteWorkRequestTarget_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteWorkRequestTarget", Erl)
        
    Exit Sub
WriteWorkRequestTarget_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteWorkRequestTarget", Erl)
End Sub

Public Sub WriteInventoryUnlockSlots(ByVal UserIndex As Integer)
    On Error Goto WriteInventoryUnlockSlots_Err
On Error GoTo WriteInventoryUnlockSlots_Err
        With UserList(UserIndex)
            If .Stats.tipoUsuario <> tNormal Then
                Call Writer.WriteInt16(ServerPacketID.eInventoryUnlockSlots)
                Select Case .Stats.tipoUsuario
                    Case tLeyenda
                        Call Writer.WriteInt8(3)
                    Case tHeroe
                        Call Writer.WriteInt8(2)
                    Case tAventurero
                        Call Writer.WriteInt8(1)
                End Select
                Call modSendData.SendData(ToIndex, UserIndex)
            End If
        End With
        Exit Sub

WriteInventoryUnlockSlots_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteInventoryUnlockSlots", Erl)
    Exit Sub
WriteInventoryUnlockSlots_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteInventoryUnlockSlots", Erl)
End Sub

Public Sub WriteIntervals(ByVal UserIndex As Integer)
    On Error Goto WriteIntervals_Err
        
        On Error GoTo WriteIntervals_Err
        

100     With UserList(UserIndex)
102         Call Writer.WriteInt16(ServerPacketID.eIntervals)
104         Call Writer.WriteInt32(.Intervals.Arco)
106         Call Writer.WriteInt32(.Intervals.Caminar)
108         Call Writer.WriteInt32(.Intervals.Golpe)
110         Call Writer.WriteInt32(.Intervals.GolpeMagia)
112         Call Writer.WriteInt32(.Intervals.Magia)
114         Call Writer.WriteInt32(.Intervals.MagiaGolpe)
116         Call Writer.WriteInt32(.Intervals.GolpeUsar)
118         Call Writer.WriteInt32(.Intervals.TrabajarExtraer)
120         Call Writer.WriteInt32(.Intervals.TrabajarConstruir)
122         Call Writer.WriteInt32(.Intervals.UsarU)
124         Call Writer.WriteInt32(.Intervals.UsarClic)
126         Call Writer.WriteInt32(IntervaloTirar)
        End With

128     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteIntervals_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteIntervals", Erl)
        
    Exit Sub
WriteIntervals_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteIntervals", Erl)
End Sub

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error Goto WriteChangeInventorySlot_Err
        
        On Error GoTo WriteChangeInventorySlot_Err
        

        Dim ObjIndex    As Integer
        Dim NaturalElementalTags As Long
        Dim PodraUsarlo As Byte

100     Call Writer.WriteInt16(ServerPacketID.eChangeInventorySlot)
102     Call Writer.WriteInt8(Slot)
104     ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex

106     If ObjIndex > 0 Then
108         PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)
            NaturalElementalTags = ObjData(UserList(UserIndex).invent.Object(Slot).ObjIndex).ElementalTags
        End If

110     Call Writer.WriteInt16(ObjIndex)
112     Call Writer.WriteInt16(UserList(UserIndex).Invent.Object(Slot).amount)
114     Call Writer.WriteBool(UserList(UserIndex).Invent.Object(Slot).Equipped)
116     Call Writer.WriteReal32(SalePrice(ObjIndex))
118     Call Writer.WriteInt8(PodraUsarlo)
        Call Writer.WriteInt32(UserList(UserIndex).invent.Object(Slot).ElementalTags Or NaturalElementalTags)
        If ObjIndex > 0 Then
119         Call Writer.WriteBool(IsSet(ObjData(ObjIndex).ObjFlags, e_ObjFlags.e_Bindable))
        Else
            Call Writer.WriteBool(False)
        End If
120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteChangeInventorySlot_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeInventorySlot", Erl)
        
    Exit Sub
WriteChangeInventorySlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteChangeInventorySlot", Erl)
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error Goto WriteChangeBankSlot_Err
        
        On Error GoTo WriteChangeBankSlot_Err
        

        Dim ObjIndex    As Integer

        Dim Valor       As Long
        Dim NaturalElementalTags As Long
        Dim PodraUsarlo As Byte

100     Call Writer.WriteInt16(ServerPacketID.eChangeBankSlot)
102     Call Writer.WriteInt8(Slot)
104     ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
108     If ObjIndex > 0 Then
110         Valor = ObjData(ObjIndex).Valor
112         PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)
114         NaturalElementalTags = ObjData(ObjIndex).ElementalTags
        Else
        End If
        
       Call Writer.WriteInt16(ObjIndex)
        Call Writer.WriteInt32(UserList(UserIndex).BancoInvent.Object(Slot).ElementalTags Or NaturalElementalTags)


        Call Writer.WriteInt16(UserList(UserIndex).BancoInvent.Object(Slot).amount)
116     Call Writer.WriteInt32(Valor)
118     Call Writer.WriteInt8(PodraUsarlo)
120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteChangeBankSlot_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeBankSlot", Erl)
        
    Exit Sub
WriteChangeBankSlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteChangeBankSlot", Erl)
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
    On Error Goto WriteChangeSpellSlot_Err
        
        On Error GoTo WriteChangeSpellSlot_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eChangeSpellSlot)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(UserList(UserIndex).Stats.UserHechizos(Slot))

106     If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
108         Call Writer.WriteInt16(UserList(UserIndex).Stats.UserHechizos(Slot))
            Call Writer.WriteBool(IsSet(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).SpellRequirementMask, e_SpellRequirementMask.eIsBindable))
        Else
110         Call Writer.WriteInt16(-1)
            Call Writer.WriteBool(False)
        End If
        
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteChangeSpellSlot_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeSpellSlot", Erl)
        
    Exit Sub
WriteChangeSpellSlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteChangeSpellSlot", Erl)
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAttributes(ByVal UserIndex As Integer)
    On Error Goto WriteAttributes_Err
        
        On Error GoTo WriteAttributes_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eAtributes)
102     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Fuerza))
104     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Agilidad))
106     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
                e_Atributos.Inteligencia))
108     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
                e_Atributos.Constitucion))
110     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Carisma))
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteAttributes_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAttributes", Erl)
        
    Exit Sub
WriteAttributes_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteAttributes", Erl)
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)
    On Error Goto WriteBlacksmithWeapons_Err
        
        On Error GoTo WriteBlacksmithWeapons_Err
        

        Dim i              As Long

        Dim validIndexes() As Integer

        Dim Count          As Integer

100     ReDim validIndexes(1 To UBound(ArmasHerrero()))
102     Call Writer.WriteInt16(ServerPacketID.eBlacksmithWeapons)

104     For i = 1 To UBound(ArmasHerrero())

            ' Can the user create this object? If so add it to the list....
106         If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills( _
                    e_Skill.Herreria) Then
108             Count = Count + 1
110             validIndexes(Count) = i
            End If

112     Next i

        ' Write the number of objects in the list
114     Call Writer.WriteInt16(Count)

        ' Write the needed data of each object
116     For i = 1 To Count
120         Call Writer.WriteInt16(ArmasHerrero(validIndexes(i)))
128     Next i

130     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteBlacksmithWeapons_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlacksmithWeapons", Erl)
        
    Exit Sub
WriteBlacksmithWeapons_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteBlacksmithWeapons", Erl)
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)
    On Error Goto WriteBlacksmithArmors_Err
        
        On Error GoTo WriteBlacksmithArmors_Err
        

        Dim i              As Long

        Dim validIndexes() As Integer

        Dim Count          As Integer

100     ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
102     Call Writer.WriteInt16(ServerPacketID.eBlacksmithArmors)

104     For i = 1 To UBound(ArmadurasHerrero())

            ' Can the user create this object? If so add it to the list....
106         If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList( _
                    UserIndex).Stats.UserSkills(e_Skill.Herreria) / ModHerreria(UserList( _
                    UserIndex).clase), 0) Then
108             Count = Count + 1
110             validIndexes(Count) = i
            End If

112     Next i

        ' Write the number of objects in the list
114     Call Writer.WriteInt16(Count)

        ' Write the needed data of each object
116     For i = 1 To Count
128         Call Writer.WriteInt16(ArmadurasHerrero(validIndexes(i)))
130     Next i

132     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteBlacksmithArmors_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlacksmithArmors", Erl)
        
    Exit Sub
WriteBlacksmithArmors_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteBlacksmithArmors", Erl)
End Sub

Public Sub WriteBlacksmithElementalRunes(ByVal UserIndex As Integer)
    On Error Goto WriteBlacksmithElementalRunes_Err
        
        On Error GoTo WriteBlacksmithElementalRunes_Err
        

        Dim i              As Long

        Dim validIndexes() As Integer

        Dim Count          As Integer

100     ReDim validIndexes(1 To UBound(BlackSmithElementalRunes()))
102     Call Writer.WriteInt16(ServerPacketID.eBlacksmithExtraObjects)

104     For i = 1 To UBound(BlackSmithElementalRunes())

            ' Can the user create this object? If so add it to the list....
106         If ObjData(BlackSmithElementalRunes(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills( _
                    e_Skill.Herreria) Then
108             Count = Count + 1
110             validIndexes(Count) = i
            End If

112     Next i

        ' Write the number of objects in the list
114     Call Writer.WriteInt16(Count)

        ' Write the needed data of each object
116     For i = 1 To Count
120         Call Writer.WriteInt16(BlackSmithElementalRunes(validIndexes(i)))
128     Next i

130     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteBlacksmithElementalRunes_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlacksmithElementalRunes", Erl)
        
    Exit Sub
WriteBlacksmithElementalRunes_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteBlacksmithElementalRunes", Erl)
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)
    On Error Goto WriteCarpenterObjects_Err
        
        On Error GoTo WriteCarpenterObjects_Err
        

        Dim i              As Long

        Dim validIndexes() As Integer

        Dim Count          As Byte

100     ReDim validIndexes(1 To UBound(ObjCarpintero()))
102     Call Writer.WriteInt16(ServerPacketID.eCarpenterObjects)

104     For i = 1 To UBound(ObjCarpintero())

            ' Can the user create this object? If so add it to the list....
106         If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList( _
                    UserIndex).Stats.UserSkills(e_Skill.Carpinteria) Then

108             If i = 1 Then Debug.Print UserList(UserIndex).Stats.UserSkills( _
                        e_Skill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase)
110             Count = Count + 1
112             validIndexes(Count) = i
            End If

114     Next i

        ' Write the number of objects in the list
116     Call Writer.WriteInt8(Count)

        ' Write the needed data of each object
118     For i = 1 To Count
120         Call Writer.WriteInt16(ObjCarpintero(validIndexes(i)))
            'Call Writer.WriteInt16(obj.Madera)
            'Call Writer.WriteInt32(obj.GrhIndex)
            ' Ladder 07/07/2014   Ahora se envia el grafico de los objetos
122     Next i

124     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCarpenterObjects_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCarpenterObjects", Erl)
        
    Exit Sub
WriteCarpenterObjects_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCarpenterObjects", Erl)
End Sub

Public Sub WriteAlquimistaObjects(ByVal UserIndex As Integer)
    On Error Goto WriteAlquimistaObjects_Err
        
        On Error GoTo WriteAlquimistaObjects_Err
        

        Dim i              As Long

        Dim validIndexes() As Integer

        Dim Count          As Integer

100     ReDim validIndexes(1 To UBound(ObjAlquimista()))
102     Call Writer.WriteInt16(ServerPacketID.eAlquimistaObj)

104     For i = 1 To UBound(ObjAlquimista())

            ' Can the user create this object? If so add it to the list....
106         If ObjData(ObjAlquimista(i)).SkPociones <= UserList(UserIndex).Stats.UserSkills( _
                    e_Skill.Alquimia) \ ModAlquimia(UserList(UserIndex).clase) Then
108             Count = Count + 1
110             validIndexes(Count) = i
            End If

112     Next i

        ' Write the number of objects in the list
114     Call Writer.WriteInt16(Count)

        ' Write the needed data of each object
116     For i = 1 To Count
118         Call Writer.WriteInt16(ObjAlquimista(validIndexes(i)))
120     Next i

122     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteAlquimistaObjects_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAlquimistaObjects", Erl)
        
    Exit Sub
WriteAlquimistaObjects_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteAlquimistaObjects", Erl)
End Sub

Public Sub WriteSastreObjects(ByVal UserIndex As Integer)
    On Error Goto WriteSastreObjects_Err
        
        On Error GoTo WriteSastreObjects_Err
        

        Dim i              As Long

        Dim validIndexes() As Integer

        Dim Count          As Integer

100     ReDim validIndexes(1 To UBound(ObjSastre()))
102     Call Writer.WriteInt16(ServerPacketID.eSastreObj)

104     For i = 1 To UBound(ObjSastre())

            ' Can the user create this object? If so add it to the list....
106         If ObjData(ObjSastre(i)).SkSastreria <= UserList(UserIndex).Stats.UserSkills( _
                    e_Skill.Sastreria) Then
108             Count = Count + 1
110             validIndexes(Count) = i
            End If

112     Next i

        ' Write the number of objects in the list
114     Call Writer.WriteInt16(Count)

        ' Write the needed data of each object
116     For i = 1 To Count
118         Call Writer.WriteInt16(ObjSastre(validIndexes(i)))
120     Next i

122     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteSastreObjects_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSastreObjects", Erl)
        
    Exit Sub
WriteSastreObjects_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteSastreObjects", Erl)
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRestOK(ByVal UserIndex As Integer)
    On Error Goto WriteRestOK_Err
        
        On Error GoTo WriteRestOK_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eRestOK)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteRestOK_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRestOK", Erl)
        
    Exit Sub
WriteRestOK_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteRestOK", Erl)
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)
    On Error Goto WriteErrorMsg_Err

        'Writes the "ErrorMsg" message to the given user's outgoing data buffer
        
        On Error GoTo WriteErrorMsg_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageErrorMsg(Message))
        
        Exit Sub

WriteErrorMsg_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteErrorMsg", Erl)
        
    Exit Sub
WriteErrorMsg_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteErrorMsg", Erl)
End Sub

''
' Writes the "Blind" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlind(ByVal UserIndex As Integer)
    On Error Goto WriteBlind_Err
        
        On Error GoTo WriteBlind_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBlind)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteBlind_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlind", Erl)
        
    Exit Sub
WriteBlind_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteBlind", Erl)
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDumb(ByVal UserIndex As Integer)
    On Error Goto WriteDumb_Err
        
        On Error GoTo WriteDumb_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eDumb)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteDumb_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDumb", Erl)
        
    Exit Sub
WriteDumb_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteDumb", Erl)
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
'Optimizacion de protocolo por Ladder
Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error Goto WriteShowSignal_Err
        
        On Error GoTo WriteShowSignal_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowSignal)
102     Call Writer.WriteInt16(ObjIndex)
104     Call Writer.WriteInt16(ObjData(ObjIndex).GrhSecundario)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowSignal_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowSignal", Erl)
        
    Exit Sub
WriteShowSignal_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowSignal", Erl)
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, _
    On Error Goto WriteChangeNPCInventorySlot_Err
                                       ByVal Slot As Byte, _
                                       ByRef obj As t_Obj, _
                                       ByVal price As Single)
        
        On Error GoTo WriteChangeNPCInventorySlot_Err
        

        Dim PodraUsarlo As Byte

100     If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
102         PodraUsarlo = PuedeUsarObjeto(UserIndex, obj.ObjIndex)
        End If

104     Call Writer.WriteInt16(ServerPacketID.eChangeNPCInventorySlot)
106     Call Writer.WriteInt8(Slot)
108     Call Writer.WriteInt16(obj.ObjIndex)
110     Call Writer.WriteInt16(obj.amount)
112     Call Writer.WriteReal32(price)
        Call Writer.WriteInt32(obj.ElementalTags)
114     Call Writer.WriteInt8(PodraUsarlo)
116     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteChangeNPCInventorySlot_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeNPCInventorySlot", Erl)
        
    Exit Sub
WriteChangeNPCInventorySlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteChangeNPCInventorySlot", Erl)
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateHungerAndThirst_Err
        
        On Error GoTo WriteUpdateHungerAndThirst_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUpdateHungerAndThirst)
102     Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxAGU)
104     Call Writer.WriteInt8(UserList(UserIndex).Stats.MinAGU)
106     Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxHam)
108     Call Writer.WriteInt8(UserList(UserIndex).Stats.MinHam)
110     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUpdateHungerAndThirst_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateHungerAndThirst", Erl)
        
    Exit Sub
WriteUpdateHungerAndThirst_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateHungerAndThirst", Erl)
End Sub

Public Sub WriteLight(ByVal UserIndex As Integer, ByVal Map As Integer)
    On Error Goto WriteLight_Err
        
        On Error GoTo WriteLight_Err
        
100     Call Writer.WriteInt16(ServerPacketID.elight)
102     Call Writer.WriteString8(MapInfo(Map).base_light)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteLight_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLight", Erl)
        
    Exit Sub
WriteLight_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteLight", Erl)
End Sub

Public Sub WriteFlashScreen(ByVal UserIndex As Integer, _
    On Error Goto WriteFlashScreen_Err
                            ByVal Color As Long, _
                            ByVal Time As Long, _
                            Optional ByVal Ignorar As Boolean = False)
        
        On Error GoTo WriteFlashScreen_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eFlashScreen)
102     Call Writer.WriteInt32(Color)
104     Call Writer.WriteInt32(Time)
106     Call Writer.WriteBool(Ignorar)
108     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteFlashScreen_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFlashScreen", Erl)
        
    Exit Sub
WriteFlashScreen_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteFlashScreen", Erl)
End Sub

Public Sub WriteFYA(ByVal UserIndex As Integer)
    On Error Goto WriteFYA_Err
        
        On Error GoTo WriteFYA_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eFYA)
102     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(1))
104     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(2))
106     Call Writer.WriteInt16(UserList(UserIndex).flags.DuracionEfecto)
108     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteFYA_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFYA", Erl)
        
    Exit Sub
WriteFYA_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteFYA", Erl)
End Sub

Public Sub WriteCerrarleCliente(ByVal UserIndex As Integer)
    On Error Goto WriteCerrarleCliente_Err
        
        On Error GoTo WriteCerrarleCliente_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCerrarleCliente)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCerrarleCliente_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCerrarleCliente", Erl)
        
    Exit Sub
WriteCerrarleCliente_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCerrarleCliente", Erl)
End Sub


Public Sub WriteContadores(ByVal UserIndex As Integer)
    On Error Goto WriteContadores_Err
 
        On Error GoTo WriteContadores_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eContadores)
102     Call Writer.WriteInt16(UserList(UserIndex).Counters.Invisibilidad)
110     Call Writer.WriteInt16(UserList(UserIndex).flags.DuracionEfecto)
        
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteContadores_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteContadores", Erl)
        
        
    Exit Sub
WriteContadores_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteContadores", Erl)
End Sub

Public Sub WriteShowPapiro(ByVal UserIndex As Integer)
    On Error Goto WriteShowPapiro_Err
    On Error GoTo WriteShowPapiro_Err
100     Call Writer.WriteInt16(ServerPacketID.eShowPapiro)
112     Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub

WriteShowPapiro_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowPapiro", Erl)
    Exit Sub
WriteShowPapiro_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowPapiro", Erl)
End Sub

Public Sub WriteUpdateCdType(ByVal UserIndex As Integer, ByVal cdType As Byte)
    On Error Goto WriteUpdateCdType_Err
On Error GoTo WriteUpdateCdType_Err
100     Call Writer.WriteInt16(ServerPacketID.eUpdateCooldownType)
110     Call Writer.WriteInt8(cdType)
112     Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub

WriteUpdateCdType_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateCdType", Erl)
    Exit Sub
WriteUpdateCdType_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateCdType", Erl)
End Sub

Public Sub WritePrivilegios(ByVal UserIndex As Integer)
    On Error Goto WritePrivilegios_Err

        
        On Error GoTo WritePrivilegios_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePrivilegios)
        
        If UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero) Then
            Call Writer.WriteBool(True)
        Else
            Call Writer.WriteBool(False)
        End If
        
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WritePrivilegios_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePrivilegios", Erl)
        
    Exit Sub
WritePrivilegios_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePrivilegios", Erl)
End Sub

Public Sub WriteBindKeys(ByVal UserIndex As Integer)
    On Error Goto WriteBindKeys_Err
        
        On Error GoTo WriteBindKeys_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBindKeys)
102     Call Writer.WriteInt8(UserList(UserIndex).ChatCombate)
104     Call Writer.WriteInt8(UserList(UserIndex).ChatGlobal)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteBindKeys_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBindKeys", Erl)
        
    Exit Sub
WriteBindKeys_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteBindKeys", Erl)
End Sub
Public Sub WriteNotificarClienteSeguido(ByVal UserIndex As Integer, ByVal siguiendo As Byte)
    On Error Goto WriteNotificarClienteSeguido_Err
    
        
        On Error GoTo WriteNotificarClienteSeguido_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eNotificarClienteSeguido)
102     Call Writer.WriteInt8(siguiendo)
120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteNotificarClienteSeguido_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNotificarClienteSeguido", Erl)
        
    Exit Sub
WriteNotificarClienteSeguido_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteNotificarClienteSeguido", Erl)
End Sub
Public Sub WriteRecievePosSeguimiento(ByVal UserIndex As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
    On Error Goto WriteRecievePosSeguimiento_Err
    
        
        On Error GoTo WriteNotificarClienteSeguido_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eRecievePosSeguimiento)
102     Call Writer.WriteInt16(PosX)
103     Call Writer.WriteInt16(PosY)
120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteNotificarClienteSeguido_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNotificarClienteSeguido", Erl)
        
    Exit Sub
WriteRecievePosSeguimiento_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteRecievePosSeguimiento", Erl)
End Sub
Public Sub WriteGetInventarioHechizos(ByVal UserIndex As Integer, ByVal Value As Byte, ByVal hechiSel As Byte, ByVal scrollSel As Byte)
    On Error Goto WriteGetInventarioHechizos_Err
    
        
        On Error GoTo GetInventarioHechizos_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eGetInventarioHechizos)
101     Call Writer.WriteInt8(Value)
        Call Writer.WriteInt8(hechiSel)
        Call Writer.WriteInt8(scrollSel)
120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

GetInventarioHechizos_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.GetInventarioHechizos", Erl)
        
    Exit Sub
WriteGetInventarioHechizos_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteGetInventarioHechizos", Erl)
End Sub

Public Sub WriteNofiticarClienteCasteo(ByVal UserIndex As Integer, ByVal Value As Byte)
    On Error Goto WriteNofiticarClienteCasteo_Err

        
        On Error GoTo NofiticarClienteCasteo_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eNotificarClienteCasteo)
101     Call Writer.WriteInt8(Value)
120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

NofiticarClienteCasteo_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.NofiticarClienteCasteo", Erl)
        
    Exit Sub
WriteNofiticarClienteCasteo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteNofiticarClienteCasteo", Erl)
End Sub

Public Sub WriteCancelarSeguimiento(ByVal UserIndex As Integer)
    On Error Goto WriteCancelarSeguimiento_Err
    
        
        On Error GoTo WriteCancelarSeguimiento_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCancelarSeguimiento)
120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCancelarSeguimiento_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCancelarSeguimiento", Erl)
    Exit Sub
WriteCancelarSeguimiento_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCancelarSeguimiento", Erl)
End Sub
Public Sub WriteSendFollowingCharindex(ByVal UserIndex As Integer, ByVal charindex As Integer)
    On Error Goto WriteSendFollowingCharindex_Err

        
        On Error GoTo WriteSendFollowingCharindex_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eSendFollowingCharIndex)
102     Call Writer.WriteInt16(charindex)
        
120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteSendFollowingCharindex_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSendFollowingCharindex", Erl)
    Exit Sub
WriteSendFollowingCharindex_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteSendFollowingCharindex", Erl)
End Sub
''
' Writes the "MiniStats" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMiniStats(ByVal UserIndex As Integer)
    On Error Goto WriteMiniStats_Err
        
        On Error GoTo WriteMiniStats_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eMiniStats)
102     Call Writer.WriteInt32(UserList(UserIndex).Faccion.ciudadanosMatados)
104     Call Writer.WriteInt32(UserList(UserIndex).Faccion.CriminalesMatados)
106     Call Writer.WriteInt8(UserList(UserIndex).Faccion.Status)
108     Call Writer.WriteInt16(UserList(UserIndex).Stats.NPCsMuertos)
110     Call Writer.WriteInt8(UserList(UserIndex).clase)
112     Call Writer.WriteInt32(UserList(UserIndex).Counters.Pena)
114     Call Writer.WriteInt32(UserList(UserIndex).flags.VecesQueMoriste)
116     Call Writer.WriteInt8(UserList(UserIndex).genero)
115     Call Writer.WriteInt32(UserList(UserIndex).Stats.PuntosPesca)
118     Call Writer.WriteInt8(UserList(UserIndex).raza)
120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteMiniStats_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMiniStats", Erl)
        
    Exit Sub
WriteMiniStats_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteMiniStats", Erl)
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data .incomingData.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
    On Error Goto WriteLevelUp_Err
        
        On Error GoTo WriteLevelUp_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eLevelUp)
102     Call Writer.WriteInt16(skillPoints)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteLevelUp_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLevelUp", Erl)
        
    Exit Sub
WriteLevelUp_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteLevelUp", Erl)
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data .incomingData.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, _
    On Error Goto WriteAddForumMsg_Err
                            ByVal title As String, _
                            ByVal Message As String)
        
        On Error GoTo WriteAddForumMsg_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eAddForumMsg)
102     Call Writer.WriteString8(title)
104     Call Writer.WriteString8(Message)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteAddForumMsg_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAddForumMsg", Erl)
        
    Exit Sub
WriteAddForumMsg_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteAddForumMsg", Erl)
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
    On Error Goto WriteShowForumForm_Err
        
        On Error GoTo WriteShowForumForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowForumForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowForumForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowForumForm", Erl)
        
    Exit Sub
WriteShowForumForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowForumForm", Erl)
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    TargetIndex The user turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSetInvisible(ByVal UserIndex As Integer, _
    On Error Goto WriteSetInvisible_Err
                             ByVal TargetIndex As Integer, _
                             ByVal invisible As Boolean)
        
        On Error GoTo WriteSetInvisible_Err
        
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageSetInvisible(UserList(TargetIndex).Char.charindex, _
                invisible, UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y))
        
        Exit Sub

WriteSetInvisible_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSetInvisible", Erl)
        
    Exit Sub
WriteSetInvisible_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteSetInvisible", Erl)
End Sub


''
' Writes the "MeditateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
    On Error Goto WriteMeditateToggle_Err
        
        On Error GoTo WriteMeditateToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eMeditateToggle)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteMeditateToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMeditateToggle", Erl)
        
    Exit Sub
WriteMeditateToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteMeditateToggle", Erl)
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
    On Error Goto WriteBlindNoMore_Err
        
        On Error GoTo WriteBlindNoMore_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBlindNoMore)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteBlindNoMore_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlindNoMore", Erl)
        
    Exit Sub
WriteBlindNoMore_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteBlindNoMore", Erl)
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
    On Error Goto WriteDumbNoMore_Err
        
        On Error GoTo WriteDumbNoMore_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eDumbNoMore)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteDumbNoMore_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDumbNoMore", Erl)
        
    Exit Sub
WriteDumbNoMore_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteDumbNoMore", Erl)
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSendSkills(ByVal UserIndex As Integer)
    On Error Goto WriteSendSkills_Err
        
        On Error GoTo WriteSendSkills_Err
        

        Dim i As Long

100     Call Writer.WriteInt16(ServerPacketID.eSendSkills)

102     For i = 1 To NUMSKILLS
104         Call Writer.WriteInt8(UserList(UserIndex).Stats.UserSkills(i))
106     Next i

108     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteSendSkills_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSendSkills", Erl)
        
    Exit Sub
WriteSendSkills_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteSendSkills", Erl)
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    On Error Goto WriteTrainerCreatureList_Err
        
        On Error GoTo WriteTrainerCreatureList_Err
        

        Dim i   As Long

        Dim str As String

100     Call Writer.WriteInt16(ServerPacketID.eTrainerCreatureList)

102     For i = 1 To NpcList(NpcIndex).NroCriaturas
104         str = str & NpcList(NpcIndex).Criaturas(i).NpcName & SEPARATOR
106     Next i

108     If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
110     Call Writer.WriteString8(str)
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteTrainerCreatureList_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTrainerCreatureList", Erl)
        
    Exit Sub
WriteTrainerCreatureList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteTrainerCreatureList", Erl)
End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildNews(ByVal UserIndex As Integer, _
    On Error Goto WriteGuildNews_Err
                          ByVal guildNews As String, _
                          ByRef guildList() As String, _
                          ByRef MemberList() As Long, _
                          ByVal ClanNivel As Byte, _
                          ByVal ExpAcu As Integer, _
                          ByVal ExpNe As Integer)
        
        On Error GoTo WriteGuildNews_Err
        

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt16(ServerPacketID.eguildNews)
102     Call Writer.WriteString8(guildNews)

        ' Prepare guild name's list
104     For i = LBound(guildList()) To UBound(guildList())
106         Tmp = Tmp & guildList(i) & SEPARATOR
108     Next i

110     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
112     Call Writer.WriteString8(Tmp)
        ' Prepare guild member's list
114     Tmp = vbNullString

116     For i = LBound(MemberList()) To UBound(MemberList())
118         Tmp = Tmp & GetUserName(MemberList(i)) & SEPARATOR
120     Next i

122     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
124     Call Writer.WriteString8(Tmp)
126     Call Writer.WriteInt8(ClanNivel)
128     Call Writer.WriteInt16(ExpAcu)
130     Call Writer.WriteInt16(ExpNe)
132     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteGuildNews_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildNews", Erl)
        
    Exit Sub
WriteGuildNews_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteGuildNews", Erl)
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)
    On Error Goto WriteOfferDetails_Err
        
        On Error GoTo WriteOfferDetails_Err
        

100     Call Writer.WriteInt16(ServerPacketID.eOfferDetails)
102     Call Writer.WriteString8(details)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteOfferDetails_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteOfferDetails", Erl)
        
    Exit Sub
WriteOfferDetails_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteOfferDetails", Erl)
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
    On Error Goto WriteAlianceProposalsList_Err
        
        On Error GoTo WriteAlianceProposalsList_Err
        

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt16(ServerPacketID.eAlianceProposalsList)

        ' Prepare guild's list
102     For i = LBound(guilds()) To UBound(guilds())
104         Tmp = Tmp & guilds(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteAlianceProposalsList_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAlianceProposalsList", Erl)
        
    Exit Sub
WriteAlianceProposalsList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteAlianceProposalsList", Erl)
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
    On Error Goto WritePeaceProposalsList_Err
        
        On Error GoTo WritePeaceProposalsList_Err
        

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt16(ServerPacketID.ePeaceProposalsList)

        ' Prepare guilds' list
102     For i = LBound(guilds()) To UBound(guilds())
104         Tmp = Tmp & guilds(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WritePeaceProposalsList_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePeaceProposalsList", Erl)
        
    Exit Sub
WritePeaceProposalsList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePeaceProposalsList", Erl)
End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal CharName As String, _
    On Error Goto WriteCharacterInfo_Err
        ByVal race As e_Raza, ByVal Class As e_Class, ByVal gender As e_Genero, ByVal _
        level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal previousPetitions As String, _
        ByVal currentGuild As String, ByVal previousGuilds As String, ByVal _
        RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As _
        Long, ByVal criminalsKilled As Long)
        
        On Error GoTo WriteCharacterInfo_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCharacterInfo)
102     Call Writer.WriteInt8(gender)
104     Call Writer.WriteString8(CharName)
106     Call Writer.WriteInt8(race)
108     Call Writer.WriteInt8(Class)
110     Call Writer.WriteInt8(level)
112     Call Writer.WriteInt32(gold)
114     Call Writer.WriteInt32(bank)
116     Call Writer.WriteString8(previousPetitions)
118     Call Writer.WriteString8(currentGuild)
120     Call Writer.WriteString8(previousGuilds)
122     Call Writer.WriteBool(RoyalArmy)
124     Call Writer.WriteBool(CaosLegion)
126     Call Writer.WriteInt32(citicensKilled)
128     Call Writer.WriteInt32(criminalsKilled)
130     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCharacterInfo_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterInfo", Erl)
        
    Exit Sub
WriteCharacterInfo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCharacterInfo", Erl)
End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, _
    On Error Goto WriteGuildLeaderInfo_Err
                                ByRef guildList() As String, _
                                ByRef MemberList() As Long, _
                                ByVal guildNews As String, _
                                ByRef joinRequests() As String, _
                                ByVal NivelDeClan As Byte, _
                                ByVal ExpActual As Integer, _
                                ByVal ExpNecesaria As Integer)
        
        On Error GoTo WriteGuildLeaderInfo_Err
        

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt16(ServerPacketID.eGuildLeaderInfo)

        ' Prepare guild name's list
102     For i = LBound(guildList()) To UBound(guildList())
104         Tmp = Tmp & guildList(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
        ' Prepare guild member's list
112     Tmp = vbNullString

114     For i = LBound(MemberList()) To UBound(MemberList())
116         Tmp = Tmp & GetUserName(MemberList(i)) & SEPARATOR
118     Next i

120     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
122     Call Writer.WriteString8(Tmp)
        ' Store guild news
124     Call Writer.WriteString8(guildNews)
        ' Prepare the join request's list
126     Tmp = vbNullString

128     For i = LBound(joinRequests()) To UBound(joinRequests())
130         Tmp = Tmp & joinRequests(i) & SEPARATOR
132     Next i

134     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
136     Call Writer.WriteString8(Tmp)
138     Call Writer.WriteInt8(NivelDeClan)
140     Call Writer.WriteInt16(ExpActual)
142     Call Writer.WriteInt16(ExpNecesaria)
144     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteGuildLeaderInfo_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildLeaderInfo", Erl)
        
    Exit Sub
WriteGuildLeaderInfo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteGuildLeaderInfo", Erl)
End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildDetails(ByVal UserIndex As Integer, _
    On Error Goto WriteGuildDetails_Err
                             ByVal GuildName As String, _
                             ByVal founder As String, _
                             ByVal foundationDate As String, _
                             ByVal leader As String, _
                             ByVal memberCount As Integer, _
                             ByVal alignment As String, _
                             ByVal guildDesc As String, _
                             ByVal NivelDeClan As Byte)
        
        On Error GoTo WriteGuildDetails_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eGuildDetails)
102     Call Writer.WriteString8(GuildName)
104     Call Writer.WriteString8(founder)
106     Call Writer.WriteString8(foundationDate)
108     Call Writer.WriteString8(leader)
110     Call Writer.WriteInt16(memberCount)
112     Call Writer.WriteString8(alignment)
114     Call Writer.WriteString8(guildDesc)
116     Call Writer.WriteInt8(NivelDeClan)
118     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteGuildDetails_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildDetails", Erl)
        
    Exit Sub
WriteGuildDetails_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteGuildDetails", Erl)
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)
    On Error Goto WriteShowGuildFundationForm_Err
        
        On Error GoTo WriteShowGuildFundationForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowGuildFundationForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowGuildFundationForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowGuildFundationForm", Erl)
        
    Exit Sub
WriteShowGuildFundationForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowGuildFundationForm", Erl)
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteParalizeOK(ByVal UserIndex As Integer)
    On Error Goto WriteParalizeOK_Err
        
        On Error GoTo WriteParalizeOK_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eParalizeOK)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteParalizeOK_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteParalizeOK", Erl)
        
    Exit Sub
WriteParalizeOK_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteParalizeOK", Erl)
End Sub

Public Sub WriteStunStart(ByVal userIndex As Integer, ByVal duration As Integer)
    On Error Goto WriteStunStart_Err
    On Error GoTo WriteStunStart_Err
100     Call Writer.WriteInt16(ServerPacketID.eStunStart)
        Call Writer.WriteInt16(duration)
102     Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteStunStart_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteStunStart", Erl)
    Exit Sub
WriteStunStart_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteStunStart", Erl)
End Sub

Public Sub WriteInmovilizaOK(ByVal UserIndex As Integer)
    On Error Goto WriteInmovilizaOK_Err
        
        On Error GoTo WriteInmovilizaOK_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eInmovilizadoOK)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteInmovilizaOK_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteInmovilizaOK", Erl)
        
    Exit Sub
WriteInmovilizaOK_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteInmovilizaOK", Erl)
End Sub

Public Sub WriteStopped(ByVal UserIndex As Integer, ByVal Stopped As Boolean)
    On Error Goto WriteStopped_Err
        
        On Error GoTo WriteStopped_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eStopped)
102     Call Writer.WriteBool(Stopped)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteStopped_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteStopped", Erl)
        
    Exit Sub
WriteStopped_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteStopped", Erl)
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
    On Error Goto WriteShowUserRequest_Err
        
        On Error GoTo WriteShowUserRequest_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowUserRequest)
102     Call Writer.WriteString8(details)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowUserRequest_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowUserRequest", Erl)
        
    Exit Sub
WriteShowUserRequest_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowUserRequest", Erl)
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    Amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, _
    On Error Goto WriteChangeUserTradeSlot_Err
                                    ByRef itemsAenviar() As t_Obj, _
                                    ByVal gold As Long, _
                                    ByVal miOferta As Boolean)
        
        On Error GoTo WriteChangeUserTradeSlot_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eChangeUserTradeSlot)
102     Call Writer.WriteBool(miOferta)
104     Call Writer.WriteInt32(gold)

        Dim i As Long

106     For i = 1 To UBound(itemsAenviar)
108         Call Writer.WriteInt16(itemsAenviar(i).ObjIndex)

110         If itemsAenviar(i).ObjIndex = 0 Then
112             Call Writer.WriteString8("")
            Else
114             Call Writer.WriteString8(ObjData(itemsAenviar(i).ObjIndex).Name)
            End If

116         If itemsAenviar(i).ObjIndex = 0 Then
118             Call Writer.WriteInt32(0)
            Else
120             Call Writer.WriteInt32(ObjData(itemsAenviar(i).ObjIndex).GrhIndex)
            End If

122         Call Writer.WriteInt32(itemsAenviar(i).amount)
            Call Writer.WriteInt32(itemsAenviar(i).ElementalTags)
124     Next i

126     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteChangeUserTradeSlot_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeUserTradeSlot", Erl)
        
    Exit Sub
WriteChangeUserTradeSlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteChangeUserTradeSlot", Erl)
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByVal ListaCompleta As Boolean)
    On Error Goto WriteSpawnList_Err
        
        On Error GoTo WriteSpawnList_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eSpawnListt)
102     Call Writer.WriteBool(ListaCompleta)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteSpawnList_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSpawnList", Erl)
        
    Exit Sub
WriteSpawnList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteSpawnList", Erl)
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)
    On Error Goto WriteShowSOSForm_Err
        
        On Error GoTo WriteShowSOSForm_Err
        

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt16(ServerPacketID.eShowSOSForm)

102     For i = 1 To Ayuda.Longitud
104         Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
106     Next i

108     If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowSOSForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowSOSForm", Erl)
        
    Exit Sub
WriteShowSOSForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowSOSForm", Erl)
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, _
    On Error Goto WriteShowMOTDEditionForm_Err
                                    ByVal currentMOTD As String)
        
        On Error GoTo WriteShowMOTDEditionForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowMOTDEditionForm)
102     Call Writer.WriteString8(currentMOTD)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowMOTDEditionForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowMOTDEditionForm", Erl)
        
    Exit Sub
WriteShowMOTDEditionForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowMOTDEditionForm", Erl)
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
    On Error Goto WriteShowGMPanelForm_Err
        
        On Error GoTo WriteShowGMPanelForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowGMPanelForm)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.Head)
104     Call Writer.WriteInt16(UserList(UserIndex).Char.Body)
106     Call Writer.WriteInt16(UserList(UserIndex).Char.CascoAnim)
108     Call Writer.WriteInt16(UserList(UserIndex).Char.WeaponAnim)
110     Call Writer.WriteInt16(UserList(UserIndex).Char.ShieldAnim)
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowGMPanelForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowGMPanelForm", Erl)
        
    Exit Sub
WriteShowGMPanelForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowGMPanelForm", Erl)
End Sub

Public Sub WriteShowFundarClanForm(ByVal UserIndex As Integer)
    On Error Goto WriteShowFundarClanForm_Err
        
        On Error GoTo WriteShowFundarClanForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowFundarClanForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowFundarClanForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowFundarClanForm", Erl)
        
    Exit Sub
WriteShowFundarClanForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowFundarClanForm", Erl)
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserNameList(ByVal UserIndex As Integer, _
    On Error Goto WriteUserNameList_Err
                             ByRef userNamesList() As String, _
                             ByVal cant As Integer)
        
        On Error GoTo WriteUserNameList_Err
        

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt16(ServerPacketID.eUserNameList)

        ' Prepare user's names list
102     For i = 1 To cant
104         Tmp = Tmp & userNamesList(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUserNameList_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserNameList", Erl)
        
    Exit Sub
WriteUserNameList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUserNameList", Erl)
End Sub


Public Sub WriteGoliathInit(ByVal UserIndex As Integer)
    On Error Goto WriteGoliathInit_Err
        
        On Error GoTo WriteGoliathInit_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eGoliath)
102     Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
104     Call Writer.WriteInt8(UserList(UserIndex).BancoInvent.NroItems)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteGoliathInit_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGoliathInit", Erl)
        
    Exit Sub
WriteGoliathInit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteGoliathInit", Erl)
End Sub
Public Sub WritePelearConPezEspecial(ByVal UserIndex As Integer)
    On Error Goto WritePelearConPezEspecial_Err
            
        On Error GoTo WritePelearConPezEspecial_Err
        
        
100     Call Writer.WriteInt16(ServerPacketID.ePelearConPezEspecial)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WritePelearConPezEspecial_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePelearConPezEspecial", Erl)
        
    Exit Sub
WritePelearConPezEspecial_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePelearConPezEspecial", Erl)
End Sub

Public Sub WriteUpdateBankGld(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateBankGld_Err
        
        On Error GoTo WriteUpdateBankGld_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUpdateBankGld)
102     Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUpdateBankGld_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateBankGld", Erl)
        
    Exit Sub
WriteUpdateBankGld_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateBankGld", Erl)
End Sub

Public Sub WriteShowFrmLogear(ByVal UserIndex As Integer)
    On Error Goto WriteShowFrmLogear_Err
        
        On Error GoTo WriteShowFrmLogear_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowFrmLogear)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowFrmLogear_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowFrmLogear", Erl)
        
    Exit Sub
WriteShowFrmLogear_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowFrmLogear", Erl)
End Sub

Public Sub WriteShowFrmMapa(ByVal UserIndex As Integer)
    On Error Goto WriteShowFrmMapa_Err
        
        On Error GoTo WriteShowFrmMapa_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowFrmMapa)
102     Call Writer.WriteInt16(SvrConfig.GetValue("ExpMult"))
104     Call Writer.WriteInt16(SvrConfig.GetValue("GoldMult"))
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteShowFrmMapa_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowFrmMapa", Erl)
        
    Exit Sub
WriteShowFrmMapa_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShowFrmMapa", Erl)
End Sub

Public Sub WritePreguntaBox(ByVal UserIndex As Integer, ByVal MsgID As Integer, Optional ByVal Param As String = vbNullString)
    On Error Goto WritePreguntaBox_Err
    On Error GoTo WritePreguntaBox_Err

    Call Writer.WriteInt16(ServerPacketID.eShowPregunta)
    Call Writer.WriteInt16(MsgID)           ' Enviar el ID
    Call Writer.WriteString8(Param)          ' Enviar el parámetro (puede ser vacío)
    Call modSendData.SendData(ToIndex, UserIndex)
    
    Exit Sub

WritePreguntaBox_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePreguntaBox", Erl)
    Exit Sub
WritePreguntaBox_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WritePreguntaBox", Erl)
End Sub



Public Sub WriteDatosGrupo(ByVal UserIndex As Integer)
    On Error Goto WriteDatosGrupo_Err
        
        On Error GoTo WriteDatosGrupo_Err
        

        Dim i As Byte

100     With UserList(UserIndex)
102         Call Writer.WriteInt16(ServerPacketID.eDatosGrupo)
104         Call Writer.WriteBool(.Grupo.EnGrupo)

106         If .Grupo.EnGrupo = True Then
108             Call Writer.WriteInt8(UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros)

110             If .Grupo.Lider.ArrayIndex = userIndex Then

112                 For i = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros

114                     If i = 1 Then
116                         Call Writer.WriteString8(UserList(.Grupo.Miembros(i).ArrayIndex).name & _
                                    "(Líder)")
                        Else
118                         Call Writer.WriteString8(UserList(.Grupo.Miembros(i).ArrayIndex).name)
                        End If

120                 Next i

                Else

122                 For i = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros

124                     If i = 1 Then
126                         Call Writer.WriteString8(UserList(UserList( _
                                    .Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex).name & "(Líder)")
                        Else
128                         Call Writer.WriteString8(UserList(UserList( _
                                    .Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex).name)
                        End If

130                 Next i

                End If
            End If

        End With

132     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteDatosGrupo_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDatosGrupo", Erl)
        
    Exit Sub
WriteDatosGrupo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteDatosGrupo", Erl)
End Sub

Public Sub WriteUbicacion(ByVal UserIndex As Integer, _
    On Error Goto WriteUbicacion_Err
                          ByVal Miembro As Byte, _
                          ByVal GPS As Integer)
        
        On Error GoTo WriteUbicacion_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eubicacion)
102     Call Writer.WriteInt8(Miembro)

104     If GPS > 0 Then
106         Call Writer.WriteInt8(UserList(GPS).Pos.X)
108         Call Writer.WriteInt8(UserList(GPS).Pos.Y)
110         Call Writer.WriteInt16(UserList(GPS).Pos.Map)
        Else
112         Call Writer.WriteInt8(0)
114         Call Writer.WriteInt8(0)
116         Call Writer.WriteInt16(0)
        End If

118     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUbicacion_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUbicacion", Erl)
        
    Exit Sub
WriteUbicacion_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUbicacion", Erl)
End Sub

Public Sub WriteViajarForm(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    On Error Goto WriteViajarForm_Err
        
        On Error GoTo WriteViajarForm_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eViajarForm)

        Dim destinos As Byte

        Dim i        As Byte

102     destinos = NpcList(NpcIndex).NumDestinos
104     Call Writer.WriteInt8(destinos)

106     For i = 1 To destinos
108         Call Writer.WriteString8(NpcList(NpcIndex).Dest(i))
110     Next i

112     Call Writer.WriteInt8(NpcList(NpcIndex).Interface)
114     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteViajarForm_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteViajarForm", Erl)
        
    Exit Sub
WriteViajarForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteViajarForm", Erl)
End Sub

Public Sub WriteQuestDetails(ByVal UserIndex As Integer, _
    On Error Goto WriteQuestDetails_Err
                             ByVal QuestIndex As Integer, _
                             Optional QuestSlot As Byte = 0)
        
        On Error GoTo WriteQuestDetails_Err
        

        Dim i As Integer

        'ID del paquete
100     Call Writer.WriteInt16(ServerPacketID.eQuestDetails)
        'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptí todavía (1 para el primer caso y 0 para el segundo)
102     Call Writer.WriteInt8(IIf(QuestSlot, 1, 0))
        'Enviamos nombre, descripción y nivel requerido de la quest
        'Call Writer.WriteString8(QuestList(QuestIndex).Nombre)
        'Call Writer.WriteString8(QuestList(QuestIndex).Desc)
104     Call Writer.WriteInt16(QuestIndex)
106     Call Writer.WriteInt8(QuestList(QuestIndex).RequiredLevel)
108     Call Writer.WriteInt16(QuestList(QuestIndex).RequiredQuest)
    
        'Enviamos la cantidad de npcs requeridos
110     Call Writer.WriteInt8(QuestList(QuestIndex).RequiredNPCs)

112     If QuestList(QuestIndex).RequiredNPCs Then

            'Si hay npcs entonces enviamos la lista
114         For i = 1 To QuestList(QuestIndex).RequiredNPCs
116             Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).amount)
118             Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)

                'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
120             If QuestSlot Then
122                 Call Writer.WriteInt16(UserList(UserIndex).QuestStats.Quests( _
                            QuestSlot).NPCsKilled(i))
                End If
124         Next i
        End If

        'Enviamos la cantidad de objs requeridos
126     Call Writer.WriteInt8(QuestList(QuestIndex).RequiredOBJs)
128     If QuestList(QuestIndex).RequiredOBJs Then
            'Si hay objs entonces enviamos la lista
130         For i = 1 To QuestList(QuestIndex).RequiredOBJs
132             Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).amount)
134             Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex)
                'escribe si tiene ese objeto en el inventario y que cantidad
136             Call Writer.WriteInt16(get_object_amount_from_inventory(userindex, QuestList( _
                        QuestIndex).RequiredOBJ(i).ObjIndex))
                ' Call Writer.WriteInt16(0)
138         Next i
        End If
139     Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.SkillType)
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.RequiredValue)

        'Enviamos la recompensa de oro y experiencia.
140     Call Writer.WriteInt32((QuestList(QuestIndex).RewardGLD * SvrConfig.GetValue("GoldMult")))
142     Call Writer.WriteInt32((QuestList(QuestIndex).RewardEXP * SvrConfig.GetValue("ExpMult")))
        'Enviamos la cantidad de objs de recompensa
144     Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)

146     If QuestList(QuestIndex).RewardOBJs Then
            'si hay objs entonces enviamos la lista
148         For i = 1 To QuestList(QuestIndex).RewardOBJs
150             Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).amount)
152             Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
154         Next i
        End If
        
        Call Writer.WriteInt8(QuestList(QuestIndex).RewardSpellCount)
        For i = 1 To QuestList(QuestIndex).RewardSpellCount
            Writer.WriteInt16 (QuestList(QuestIndex).RewardSpellList(i))
        Next i
        
156     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteQuestDetails_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteQuestDetails", Erl)
        
    Exit Sub
WriteQuestDetails_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteQuestDetails", Erl)
End Sub
 
Public Sub WriteQuestListSend(ByVal UserIndex As Integer)
    On Error Goto WriteQuestListSend_Err
        
        On Error GoTo WriteQuestListSend_Err
        

        Dim i       As Integer

        Dim tmpStr  As String

        Dim tmpByte As Byte

100     With UserList(UserIndex)
102         Call Writer.WriteInt16(ServerPacketID.eQuestListSend)

104         For i = 1 To MAXUSERQUESTS

106             If .QuestStats.Quests(i).QuestIndex Then
108                 tmpByte = tmpByte + 1
110                 tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).nombre & ";"
                End If

112         Next i

            'Escribimos la cantidad de quests
114         Call Writer.WriteInt8(tmpByte)

            'Escribimos la lista de quests (sacamos el íltimo caracter)
116         If tmpByte Then
118             Call Writer.WriteString8(Left$(tmpStr, Len(tmpStr) - 1))
            End If

        End With

120     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteQuestListSend_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteQuestListSend", Erl)
        
    Exit Sub
WriteQuestListSend_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteQuestListSend", Erl)
End Sub

Public Sub WriteNpcQuestListSend(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    On Error Goto WriteNpcQuestListSend_Err
        
        On Error GoTo WriteNpcQuestListSend_Err
        

        Dim i          As Integer

        Dim j          As Integer

        Dim QuestIndex As Integer

100     Call Writer.WriteInt16(ServerPacketID.eNpcQuestListSend)
102     Call Writer.WriteInt8(NpcList(NpcIndex).NumQuest) 'Escribimos primero cuantas quest tiene el NPC

104     For j = 1 To NpcList(NpcIndex).NumQuest
106         QuestIndex = NpcList(NpcIndex).QuestNumber(j)
108         Call Writer.WriteInt16(QuestIndex)
110         Call Writer.WriteInt8(QuestList(QuestIndex).RequiredLevel)
112         Call Writer.WriteInt16(QuestList(QuestIndex).RequiredQuest)
            Call Writer.WriteInt8(QuestList(QuestIndex).RequiredClass)
            Call Writer.WriteInt8(QuestList(QuestIndex).LimitLevel)
            'Enviamos la cantidad de npcs requeridos
114         Call Writer.WriteInt8(QuestList(QuestIndex).RequiredNPCs)

116         If QuestList(QuestIndex).RequiredNPCs Then
                'Si hay npcs entonces enviamos la lista
118             For i = 1 To QuestList(QuestIndex).RequiredNPCs
120                 Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).amount)
122                 Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)
                    'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
                    'If QuestSlot Then
                    ' Call Writer.WriteInt16(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
                    ' End If
124             Next i
            End If
            'Enviamos la cantidad de objs requeridos
126         Call Writer.WriteInt8(QuestList(QuestIndex).RequiredOBJs)
128         If QuestList(QuestIndex).RequiredOBJs Then
                'Si hay objs entonces enviamos la lista
130             For i = 1 To QuestList(QuestIndex).RequiredOBJs
132                 Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).amount)
134                 Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex)
136             Next i
            End If
            Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSpellCount)
            If QuestList(QuestIndex).RequiredSpellCount > 0 Then
                For i = 1 To QuestList(QuestIndex).RequiredSpellCount
                    Call Writer.WriteInt16(QuestList(QuestIndex).RequiredSpellList(i))
                Next i
            End If
137         Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.SkillType)
            Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.RequiredValue)
            'Enviamos la recompensa de oro y experiencia.
138         Call Writer.WriteInt32(QuestList(QuestIndex).RewardGLD * SvrConfig.GetValue("GoldMult"))
140         Call Writer.WriteInt32(QuestList(QuestIndex).RewardEXP * SvrConfig.GetValue("ExpMult"))
            'Enviamos la cantidad de objs de recompensa
142         Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)

144         If QuestList(QuestIndex).RewardOBJs Then
                'si hay objs entonces enviamos la lista
146             For i = 1 To QuestList(QuestIndex).RewardOBJs
148                 Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).amount)
150                 Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
152             Next i
            End If
            Call Writer.WriteInt8(QuestList(QuestIndex).RewardSpellCount)
            For i = 1 To QuestList(QuestIndex).RewardSpellCount
                Call Writer.WriteInt16(QuestList(QuestIndex).RewardSpellList(i))
            Next i

            'Enviamos el estado de la QUEST
            '0 Disponible
            '1 EN CURSO
            '2 REALIZADA
            '3 no puede hacerla
            Dim PuedeHacerla As Boolean

            'La tiene aceptada el usuario?
154         If TieneQuest(UserIndex, QuestIndex) Then
156             Call Writer.WriteInt8(1)
            Else

158             If UserDoneQuest(UserIndex, QuestIndex) Then
160                 Call Writer.WriteInt8(2)
                Else
162                 PuedeHacerla = True

164                 If QuestList(QuestIndex).RequiredQuest > 0 Then
166                     If Not UserDoneQuest(UserIndex, QuestList( _
                                QuestIndex).RequiredQuest) Then
168                         PuedeHacerla = False
                        End If
                    End If

170                 If UserList(UserIndex).Stats.ELV < QuestList(QuestIndex).RequiredLevel _
                            Then
172                     PuedeHacerla = False
                    End If
                    
                    'Si el personaje es nivel mayor al limite no puede hacerla
                    If UserList(UserIndex).Stats.ELV > QuestList(QuestIndex).LimitLevel _
                            Then
                         PuedeHacerla = False
                    End If
                    
174                 If PuedeHacerla Then
176                     Call Writer.WriteInt8(0)
                    Else
178                     Call Writer.WriteInt8(3)
                    End If
                End If
            End If

180     Next j

182     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteNpcQuestListSend_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNpcQuestListSend", Erl)
        
    Exit Sub
WriteNpcQuestListSend_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteNpcQuestListSend", Erl)
End Sub

Sub WriteCommerceRecieveChatMessage(ByVal UserIndex As Integer, ByVal Message As String)
    On Error Goto WriteCommerceRecieveChatMessage_Err
        
        On Error GoTo WriteCommerceRecieveChatMessage_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCommerceRecieveChatMessage)
102     Call Writer.WriteString8(Message)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCommerceRecieveChatMessage_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCommerceRecieveChatMessage", Erl)
        
    Exit Sub
WriteCommerceRecieveChatMessage_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCommerceRecieveChatMessage", Erl)
End Sub

Sub WriteInvasionInfo(ByVal UserIndex As Integer, _
    On Error Goto WriteInvasionInfo_Err
                      ByVal Invasion As Integer, _
                      ByVal PorcentajeVida As Byte, _
                      ByVal PorcentajeTiempo As Byte)
        
        On Error GoTo WriteInvasionInfo_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eInvasionInfo)
102     Call Writer.WriteInt8(Invasion)
104     Call Writer.WriteInt8(PorcentajeVida)
106     Call Writer.WriteInt8(PorcentajeTiempo)
108     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteInvasionInfo_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteInvasionInfo", Erl)
        
    Exit Sub
WriteInvasionInfo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteInvasionInfo", Erl)
End Sub

Sub WriteOpenCrafting(ByVal UserIndex As Integer, ByVal Tipo As Byte)
    On Error Goto WriteOpenCrafting_Err
        
        On Error GoTo WriteOpenCrafting_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eOpenCrafting)
102     Call Writer.WriteInt8(Tipo)
104     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteOpenCrafting_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteOpenCrafting", Erl)
        
    Exit Sub
WriteOpenCrafting_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteOpenCrafting", Erl)
End Sub

Sub WriteCraftingItem(ByVal UserIndex As Integer, _
    On Error Goto WriteCraftingItem_Err
                      ByVal Slot As Byte, _
                      ByVal ObjIndex As Integer)
        
        On Error GoTo WriteCraftingItem_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCraftingItem)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(ObjIndex)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCraftingItem_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCraftingItem", Erl)
        
    Exit Sub
WriteCraftingItem_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCraftingItem", Erl)
End Sub

Sub WriteCraftingCatalyst(ByVal UserIndex As Integer, _
    On Error Goto WriteCraftingCatalyst_Err
                          ByVal ObjIndex As Integer, _
                          ByVal amount As Integer, _
                          ByVal Porcentaje As Byte)
        
        On Error GoTo WriteCraftingCatalyst_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCraftingCatalyst)
102     Call Writer.WriteInt16(ObjIndex)
104     Call Writer.WriteInt16(amount)
106     Call Writer.WriteInt8(Porcentaje)
108     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCraftingCatalyst_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCraftingCatalyst", Erl)
        
    Exit Sub
WriteCraftingCatalyst_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCraftingCatalyst", Erl)
End Sub

Sub WriteCraftingResult(ByVal UserIndex As Integer, _
    On Error Goto WriteCraftingResult_Err
                        ByVal Result As Integer, _
                        Optional ByVal Porcentaje As Byte = 0, _
                        Optional ByVal Precio As Long = 0)
        
        On Error GoTo WriteCraftingResult_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCraftingResult)
102     Call Writer.WriteInt16(Result)

104     If Result <> 0 Then
106         Call Writer.WriteInt8(Porcentaje)
108         Call Writer.WriteInt32(Precio)
        End If

110     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteCraftingResult_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCraftingResult", Erl)
        
    Exit Sub
WriteCraftingResult_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteCraftingResult", Erl)
End Sub


Public Sub WriteUpdateNPCSimbolo(ByVal UserIndex As Integer, _
    On Error Goto WriteUpdateNPCSimbolo_Err
                                 ByVal NpcIndex As Integer, _
                                 ByVal Simbolo As Byte)
        
        On Error GoTo WriteUpdateNPCSimbolo_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUpdateNPCSimbolo)
102     Call Writer.WriteInt16(NpcList(NpcIndex).Char.CharIndex)
104     Call Writer.WriteInt8(Simbolo)
106     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteUpdateNPCSimbolo_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateNPCSimbolo", Erl)
        
    Exit Sub
WriteUpdateNPCSimbolo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateNPCSimbolo", Erl)
End Sub

Public Sub WriteGuardNotice(ByVal UserIndex As Integer)
    On Error Goto WriteGuardNotice_Err
        
        On Error GoTo WriteGuardNotice_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eGuardNotice)
102     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteGuardNotice_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuardNotice", Erl)
        
    Exit Sub
WriteGuardNotice_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteGuardNotice", Erl)
End Sub

' \Begin: [Prepares]
Public Function PrepareMessageCharSwing(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageCharSwing_Err
                                        Optional ByVal FX As Boolean = True, _
                                        Optional ByVal ShowText As Boolean = True, _
                                        Optional ByVal NotificoTexto As Boolean = True)
        
        On Error GoTo PrepareMessageCharSwing_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCharSwing)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteBool(FX)
106     Call Writer.WriteBool(ShowText)
107     Call Writer.WriteBool(NotificoTexto)
        
        Exit Function

PrepareMessageCharSwing_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharSwing", Erl)
        
    Exit Function
PrepareMessageCharSwing_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageCharSwing", Erl)
End Function

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.
Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageSetInvisible_Err
                                           ByVal invisible As Boolean, _
                                           Optional ByVal X As Byte = 0, _
                                           Optional ByVal y As Byte = 0)
        
        On Error GoTo PrepareMessageSetInvisible_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eSetInvisible)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteBool(invisible)
105     Call Writer.WriteInt8(X)
106     Call Writer.WriteInt8(y)
        
        Exit Function

PrepareMessageSetInvisible_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageSetInvisible", Erl)
        
    Exit Function
PrepareMessageSetInvisible_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageSetInvisible", Erl)
End Function

Public Function PrepareLocaleChatOverHead(ByVal chat As Integer, _
    On Error Goto PrepareLocaleChatOverHead_Err
                                         ByVal Params As String, _
                                         ByVal charindex As Integer, _
                                         ByVal Color As Long, _
                                         Optional ByVal EsSpell As Boolean = False, _
                                         Optional ByVal x As Byte = 0, _
                                         Optional ByVal y As Byte = 0, _
                                         Optional ByVal RequiredMinDisplayTime As Integer = 0, _
                                         Optional ByVal MaxDisplayTime As Integer = 0)
    On Error GoTo PrepareMessageChatOverHead_Err

106     Call Writer.WriteInt16(ServerPacketID.eLocaleChatOverHead)
108     Call Writer.WriteInt16(chat)
109     Call Writer.WriteString8(Params)
110     Call Writer.WriteInt16(charindex)
118     Call Writer.WriteInt32(Color)
119     Call Writer.WriteBool(EsSpell)
        Call Writer.WriteInt8(x)
        Call Writer.WriteInt8(y)
        Call Writer.WriteInt16(RequiredMinDisplayTime)
        Call Writer.WriteInt16(MaxDisplayTime)
        Exit Function

PrepareMessageChatOverHead_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageChatOverHead", Erl)
    Exit Function
PrepareLocaleChatOverHead_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareLocaleChatOverHead", Erl)
End Function
''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.
Public Function PrepareMessageChatOverHead(ByVal chat As String, _
    On Error Goto PrepareMessageChatOverHead_Err
                                           ByVal CharIndex As Integer, _
                                           ByVal Color As Long, _
                                           Optional ByVal EsSpell As Boolean = False, _
                                           Optional ByVal X As Byte = 0, _
                                           Optional ByVal y As Byte = 0, _
                                           Optional ByVal RequiredMinDisplayTime As Integer = 0, _
                                           Optional ByVal MaxDisplayTime As Integer = 0)
    On Error GoTo PrepareMessageChatOverHead_Err
106     Call Writer.WriteInt16(ServerPacketID.eChatOverHead)
108     Call Writer.WriteString8(chat)
110     Call Writer.WriteInt16(CharIndex)
118     Call Writer.WriteInt32(Color)
119     Call Writer.WriteBool(EsSpell)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(y)
        Call Writer.WriteInt16(RequiredMinDisplayTime)
        Call Writer.WriteInt16(MaxDisplayTime)
        Exit Function

PrepareMessageChatOverHead_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageChatOverHead", Erl)
        
    Exit Function
PrepareMessageChatOverHead_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageChatOverHead", Erl)
End Function

Public Function PrepareMessageTextOverChar(ByVal chat As String, _
    On Error Goto PrepareMessageTextOverChar_Err
                                           ByVal CharIndex As Integer, _
                                           ByVal Color As Long)
        
        On Error GoTo PrepareMessageTextOverChar_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eTextOverChar)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt16(CharIndex)
106     Call Writer.WriteInt32(Color)
        
        Exit Function

PrepareMessageTextOverChar_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageTextOverChar", Erl)
        
    Exit Function
PrepareMessageTextOverChar_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageTextOverChar", Erl)
End Function

Public Function PrepareMessageTextCharDrop(ByVal chat As String, _
    On Error Goto PrepareMessageTextCharDrop_Err
                                           ByVal CharIndex As Integer, _
                                           ByVal Color As Long, _
                                           Optional ByVal Duration As Integer = 1300, _
                                           Optional ByVal Animated As Boolean = True)
        
        On Error GoTo PrepareMessageTextCharDrop_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eTextCharDrop)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt16(CharIndex)
106     Call Writer.WriteInt32(Color)
110     Call Writer.WriteInt16(Duration)
114     Call Writer.WriteBool(Animated)
        
        Exit Function

PrepareMessageTextCharDrop_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageTextCharDrop", Erl)
        
    Exit Function
PrepareMessageTextCharDrop_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageTextCharDrop", Erl)
End Function

Public Function PrepareMessageTextOverTile(ByVal chat As String, _
    On Error Goto PrepareMessageTextOverTile_Err
                                           ByVal X As Integer, _
                                           ByVal Y As Integer, _
                                           ByVal Color As Long, _
                                           Optional ByVal Duration As Integer = 1300, _
                                           Optional ByVal OffsetY As Integer = 0, _
                                           Optional ByVal Animated As Boolean = True)
        
        On Error GoTo PrepareMessageTextOverTile_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eTextOverTile)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt16(X)
106     Call Writer.WriteInt16(Y)
108     Call Writer.WriteInt32(Color)
110     Call Writer.WriteInt16(Duration)
112     Call Writer.WriteInt16(OffsetY)
114     Call Writer.WriteBool(Animated)
        
        Exit Function

PrepareMessageTextOverTile_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageTextOverTile", Erl)
        
    Exit Function
PrepareMessageTextOverTile_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageTextOverTile", Erl)
End Function


Public Function PrepareConsoleCharText(ByVal chat As String, ByVal Color As Long, ByVal sourceName As String, _
    On Error Goto PrepareConsoleCharText_Err
                                       ByVal sourceStatus As Integer, ByVal Privileges As Integer)
On Error GoTo PrepareConsoleCharText_Err
100     Call Writer.WriteInt16(ServerPacketID.eConsoleCharText)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt32(Color)
106     Call Writer.WriteString8(sourceName)
108     Call Writer.WriteInt16(sourceStatus)
112     Call Writer.WriteInt16(Privileges)
        Exit Function
PrepareConsoleCharText_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareConsoleCharText", Erl)
    Exit Function
PrepareConsoleCharText_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareConsoleCharText", Erl)
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageConsoleMsg(ByVal chat As String, _
    On Error Goto PrepareMessageConsoleMsg_Err
                                         ByVal FontIndex As e_FontTypeNames)
        
        On Error GoTo PrepareMessageConsoleMsg_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eConsoleMsg)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt8(FontIndex)
        
        Exit Function

PrepareMessageConsoleMsg_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageConsoleMsg", Erl)
        
    Exit Function
PrepareMessageConsoleMsg_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageConsoleMsg", Erl)
End Function
Public Function PrepareFactionMessageConsole(ByVal factionLabel As String, ByVal chat As String, ByVal FontIndex As e_FontTypeNames)
    On Error Goto PrepareFactionMessageConsole_Err
    On Error GoTo PrepareFactionMessageConsole_Err
             
        Call Writer.WriteInt16(ServerPacketID.eConsoleFactionMessage)
        Call Writer.WriteString8(chat)
        Call Writer.WriteInt8(FontIndex)
        Call Writer.WriteString8(factionLabel)
        
        Exit Function

PrepareFactionMessageConsole_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareFactionMessageConsole", Erl)
    Exit Function
PrepareFactionMessageConsole_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareFactionMessageConsole", Erl)
End Function
Public Function PrepareMessageLocaleMsg(ByVal ID As Integer, _
    On Error Goto PrepareMessageLocaleMsg_Err
                                        ByVal chat As String, _
                                        ByVal FontIndex As e_FontTypeNames)
    
    On Error GoTo PrepareMessageLocaleMsg_Err
    
100     Call Writer.WriteInt16(ServerPacketID.eLocaleMsg)
102     Call Writer.WriteInt16(ID)
104     Call Writer.WriteString8(chat)
106     Call Writer.WriteInt8(FontIndex)
    
    Exit Function

PrepareMessageLocaleMsg_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageLocaleMsg", Erl)
    
    Exit Function
PrepareMessageLocaleMsg_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageLocaleMsg", Erl)
End Function


''
' Prepares the "CharAtaca" message and returns it.
'
Public Function PrepareMessageCharAtaca(ByVal charindex As Integer, ByVal attackerIndex As Integer, ByVal danio As Long, ByVal AnimAttack As Integer)
    On Error Goto PrepareMessageCharAtaca_Err
        
        On Error GoTo PrepareMessageCharAtaca_Err
        
        
100     Call Writer.WriteInt16(ServerPacketID.eCharAtaca)
102     Call Writer.WriteInt16(charindex)
104     Call Writer.WriteInt16(attackerIndex)
106     Call Writer.WriteInt32(danio)
108     Call Writer.WriteInt16(AnimAttack)

        
        Exit Function

PrepareMessageCharAtaca_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharAtaca", Erl)
        
    Exit Function
PrepareMessageCharAtaca_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageCharAtaca", Erl)
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageCreateFX_Err
                                       ByVal FX As Integer, _
                                       ByVal FXLoops As Integer, _
                                       Optional ByVal X As Byte = 0, _
                                       Optional ByVal y As Byte = 0)
        
        On Error GoTo PrepareMessageCreateFX_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCreateFX)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(FX)
106     Call Writer.WriteInt16(FXLoops)
107     Call Writer.WriteInt8(X)
108     Call Writer.WriteInt8(y)
        
        Exit Function

PrepareMessageCreateFX_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCreateFX", Erl)
        
    Exit Function
PrepareMessageCreateFX_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageCreateFX", Erl)
End Function

Public Function PrepareMessageMeditateToggle(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageMeditateToggle_Err
                                             ByVal FX As Integer, _
                                             Optional ByVal X As Byte = 0, _
                                             Optional ByVal y As Byte = 0)
        
        On Error GoTo PrepareMessageMeditateToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eMeditateToggle)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(FX)
105     Call Writer.WriteInt8(X)
106     Call Writer.WriteInt8(y)
        
        Exit Function

PrepareMessageMeditateToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageMeditateToggle", Erl)
        
    Exit Function
PrepareMessageMeditateToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageMeditateToggle", Erl)
End Function

Public Function PrepareMessageParticleFX(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageParticleFX_Err
                                         ByVal Particula As Integer, _
                                         ByVal Time As Long, _
                                         ByVal Remove As Boolean, _
                                         Optional ByVal grh As Long = 0, _
                                         Optional ByVal X As Byte = 0, _
                                         Optional ByVal y As Byte = 0)
        
        On Error GoTo PrepareMessageParticleFX_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eParticleFX)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(Particula)
106     Call Writer.WriteInt32(Time)
108     Call Writer.WriteBool(Remove)
110     Call Writer.WriteInt32(grh)
        Call Writer.WriteInt8(X)
        Call Writer.WriteInt8(y)
        
        Exit Function

PrepareMessageParticleFX_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFX", Erl)
        
    Exit Function
PrepareMessageParticleFX_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageParticleFX", Erl)
End Function

Public Function PrepareMessageParticleFXWithDestino(ByVal Emisor As Integer, _
    On Error Goto PrepareMessageParticleFXWithDestino_Err
                                                    ByVal Receptor As Integer, _
                                                    ByVal ParticulaViaje As Integer, _
                                                    ByVal ParticulaFinal As Integer, _
                                                    ByVal Time As Long, _
                                                    ByVal wav As Integer, _
                                                    ByVal FX As Integer, _
                                                    Optional ByVal X As Byte = 0, _
                                                    Optional ByVal y As Byte = 0)
        
        On Error GoTo PrepareMessageParticleFXWithDestino_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eParticleFXWithDestino)
102     Call Writer.WriteInt16(Emisor)
104     Call Writer.WriteInt16(Receptor)
106     Call Writer.WriteInt16(ParticulaViaje)
108     Call Writer.WriteInt16(ParticulaFinal)
110     Call Writer.WriteInt32(Time)
112     Call Writer.WriteInt16(wav)
114     Call Writer.WriteInt16(FX)
115     Call Writer.WriteInt8(X)
116     Call Writer.WriteInt8(y)
    
        
        Exit Function

PrepareMessageParticleFXWithDestino_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFXWithDestino", Erl)
        
    Exit Function
PrepareMessageParticleFXWithDestino_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageParticleFXWithDestino", Erl)
End Function

Public Function PrepareMessageParticleFXWithDestinoXY(ByVal Emisor As Integer, _
    On Error Goto PrepareMessageParticleFXWithDestinoXY_Err
                                                      ByVal ParticulaViaje As Integer, _
                                                      ByVal ParticulaFinal As Integer, _
                                                      ByVal Time As Long, _
                                                      ByVal wav As Integer, _
                                                      ByVal FX As Integer, _
                                                      ByVal X As Byte, _
                                                      ByVal Y As Byte)
        
        On Error GoTo PrepareMessageParticleFXWithDestinoXY_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eParticleFXWithDestinoXY)
102     Call Writer.WriteInt16(Emisor)
104     Call Writer.WriteInt16(ParticulaViaje)
106     Call Writer.WriteInt16(ParticulaFinal)
108     Call Writer.WriteInt32(Time)
110     Call Writer.WriteInt16(wav)
112     Call Writer.WriteInt16(FX)
114     Call Writer.WriteInt8(X)
116     Call Writer.WriteInt8(Y)
        
        Exit Function

PrepareMessageParticleFXWithDestinoXY_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFXWithDestinoXY", Erl)
        
    Exit Function
PrepareMessageParticleFXWithDestinoXY_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageParticleFXWithDestinoXY", Erl)
End Function

Public Function PrepareMessageAuraToChar(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageAuraToChar_Err
                                         ByVal Aura As String, _
                                         ByVal Remove As Boolean, _
                                         ByVal Tipo As Byte)
        
        On Error GoTo PrepareMessageAuraToChar_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eAuraToChar)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteString8(Aura)
106     Call Writer.WriteBool(Remove)
108     Call Writer.WriteInt8(Tipo)
        
        Exit Function

PrepareMessageAuraToChar_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageAuraToChar", Erl)
        
    Exit Function
PrepareMessageAuraToChar_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageAuraToChar", Erl)
End Function

Public Function PrepareMessageSpeedingACT(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageSpeedingACT_Err
                                          ByVal speeding As Single)
        
        On Error GoTo PrepareMessageSpeedingACT_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eSpeedToChar)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteReal32(speeding)
        
        Exit Function

PrepareMessageSpeedingACT_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageSpeedingACT", Erl)
        
    Exit Function
PrepareMessageSpeedingACT_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageSpeedingACT", Erl)
End Function

Public Function PrepareMessageParticleFXToFloor(ByVal X As Byte, _
    On Error Goto PrepareMessageParticleFXToFloor_Err
                                                ByVal Y As Byte, _
                                                ByVal Particula As Integer, _
                                                ByVal Time As Long)
        
        On Error GoTo PrepareMessageParticleFXToFloor_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eParticleFXToFloor)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt16(Particula)
108     Call Writer.WriteInt32(Time)
        
        Exit Function

PrepareMessageParticleFXToFloor_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFXToFloor", Erl)
        
    Exit Function
PrepareMessageParticleFXToFloor_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageParticleFXToFloor", Erl)
End Function

Public Function PrepareMessageLightFXToFloor(ByVal X As Byte, _
    On Error Goto PrepareMessageLightFXToFloor_Err
                                             ByVal Y As Byte, _
                                             ByVal LuzColor As Long, _
                                             ByVal Rango As Byte)
        
        On Error GoTo PrepareMessageLightFXToFloor_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eLightToFloor)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt32(LuzColor)
108     Call Writer.WriteInt8(Rango)
        
        Exit Function

PrepareMessageLightFXToFloor_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageLightFXToFloor", Erl)
        
    Exit Function
PrepareMessageLightFXToFloor_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageLightFXToFloor", Erl)
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessagePlayWave(ByVal wave As Integer, _
    On Error Goto PrepareMessagePlayWave_Err
                                       ByVal X As Byte, _
                                       ByVal Y As Byte, _
                                       Optional ByVal CancelLastWave As Byte = False, _
                                       Optional ByVal Localize As Byte = 0)
        
        On Error GoTo PrepareMessagePlayWave_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePlayWave)
102     Call Writer.WriteInt16(wave)
104     Call Writer.WriteInt8(X)
106     Call Writer.WriteInt8(Y)
108     Call Writer.WriteInt8(CancelLastWave)
        Call Writer.WriteInt8(Localize)
        Exit Function

PrepareMessagePlayWave_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessagePlayWave", Erl)
        
    Exit Function
PrepareMessagePlayWave_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessagePlayWave", Erl)
End Function


Public Function PrepareMessageUbicacionLlamada(ByVal Mapa As Integer, _
    On Error Goto PrepareMessageUbicacionLlamada_Err
                                               ByVal X As Byte, _
                                               ByVal Y As Byte)
        
        On Error GoTo PrepareMessageUbicacionLlamada_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePosLLamadaDeClan)
102     Call Writer.WriteInt16(Mapa)
104     Call Writer.WriteInt8(X)
106     Call Writer.WriteInt8(Y)
        
        Exit Function

PrepareMessageUbicacionLlamada_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageUbicacionLlamada", Erl)
        
    Exit Function
PrepareMessageUbicacionLlamada_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageUbicacionLlamada", Erl)
End Function

Public Function PrepareMessageCharUpdateHP(ByVal UserIndex As Integer)
    On Error Goto PrepareMessageCharUpdateHP_Err
        
        On Error GoTo PrepareMessageCharUpdateHP_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCharUpdateHP)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
104     Call Writer.WriteInt32(UserList(UserIndex).Stats.MinHp)
106     Call Writer.WriteInt32(UserList(UserIndex).Stats.MaxHp)
        Call Writer.WriteInt32(UserList(UserIndex).Stats.Shield)
        
        Exit Function

PrepareMessageCharUpdateHP_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharUpdateHP", Erl)
        
    Exit Function
PrepareMessageCharUpdateHP_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageCharUpdateHP", Erl)
End Function

Public Function PrepareMessageCharUpdateMAN(ByVal UserIndex As Integer)
    On Error Goto PrepareMessageCharUpdateMAN_Err
        
        On Error GoTo PrepareMessageCharUpdateMAN_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCharUpdateMAN)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
104     Call Writer.WriteInt32(UserList(UserIndex).Stats.MinMAN)
106     Call Writer.WriteInt32(UserList(UserIndex).Stats.MaxMAN)
        
        Exit Function

PrepareMessageCharUpdateMAN_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharUpdateMAN", Erl)
        
    Exit Function
PrepareMessageCharUpdateMAN_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageCharUpdateMAN", Erl)
End Function

Public Function PrepareMessageNpcUpdateHP(ByVal NpcIndex As Integer)
    On Error Goto PrepareMessageNpcUpdateHP_Err
        
        On Error GoTo PrepareMessageNpcUpdateHP_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCharUpdateHP)
102     Call Writer.WriteInt16(NpcList(NpcIndex).Char.CharIndex)
104     Call Writer.WriteInt32(NpcList(NpcIndex).Stats.MinHp)
106     Call Writer.WriteInt32(NpcList(NpcIndex).Stats.MaxHp)
        Call Writer.WriteInt32(NpcList(NpcIndex).Stats.Shield)
        
        Exit Function

PrepareMessageNpcUpdateHP_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageNpcUpdateHP", Erl)
        
    Exit Function
PrepareMessageNpcUpdateHP_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageNpcUpdateHP", Erl)
End Function

Public Function PrepareMessageArmaMov(ByVal charindex As Integer, Optional ByVal isRanged As Byte = 0)
    On Error Goto PrepareMessageArmaMov_Err
        
        On Error GoTo PrepareMessageArmaMov_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eArmaMov)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt8(isRanged)
        
        Exit Function

PrepareMessageArmaMov_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageArmaMov", Erl)
        
    Exit Function
PrepareMessageArmaMov_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageArmaMov", Erl)
End Function

Public Function PrepareCreateProjectile(ByVal startX As Byte, ByVal startY As Byte, ByVal targetX As Byte, ByVal targetY As Byte, ByVal ProjectileType As Byte)
    On Error Goto PrepareCreateProjectile_Err
    
        On Error GoTo PrepareCreateProjectile_Err
        
        Call Writer.WriteInt16(ServerPacketID.eCreateProjectile)
        Call Writer.WriteInt8(startX)
        Call Writer.WriteInt8(startY)
        Call Writer.WriteInt8(targetX)
        Call Writer.WriteInt8(targetY)
        Call Writer.WriteInt8(ProjectileType)
        
        Exit Function

PrepareCreateProjectile_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareCreateProjectile", Erl)
        
    Exit Function
PrepareCreateProjectile_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareCreateProjectile", Erl)
End Function

Public Function PrepareMessageEscudoMov(ByVal CharIndex As Integer)
    On Error Goto PrepareMessageEscudoMov_Err
        
        On Error GoTo PrepareMessageEscudoMov_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eEscudoMov)
102     Call Writer.WriteInt16(CharIndex)
        
        Exit Function

PrepareMessageEscudoMov_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageEscudoMov", Erl)
        
    Exit Function
PrepareMessageEscudoMov_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageEscudoMov", Erl)
End Function

Public Function PrepareMessageFlashScreen(ByVal Color As Long, _
    On Error Goto PrepareMessageFlashScreen_Err
                                          ByVal Duracion As Long, _
                                          Optional ByVal Ignorar As Boolean = False)
        
        On Error GoTo PrepareMessageFlashScreen_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eFlashScreen)
102     Call Writer.WriteInt32(Color)
104     Call Writer.WriteInt32(Duracion)
106     Call Writer.WriteBool(Ignorar)
        
        Exit Function

PrepareMessageFlashScreen_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageFlashScreen", Erl)
        
    Exit Function
PrepareMessageFlashScreen_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageFlashScreen", Erl)
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageGuildChat(ByVal chat As String, ByVal Status As Byte)
    On Error Goto PrepareMessageGuildChat_Err
        
        On Error GoTo PrepareMessageGuildChat_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eGuildChat)
102     Call Writer.WriteInt8(Status)
104     Call Writer.WriteString8(chat)
        
        Exit Function

PrepareMessageGuildChat_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageGuildChat", Erl)
        
    Exit Function
PrepareMessageGuildChat_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageGuildChat", Erl)
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageShowMessageBox(ByVal chat As String)
    On Error Goto PrepareMessageShowMessageBox_Err
        
        On Error GoTo PrepareMessageShowMessageBox_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eShowMessageBox)
102     Call Writer.WriteString8(chat)
        
        Exit Function

PrepareMessageShowMessageBox_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageShowMessageBox", Erl)
        
    Exit Function
PrepareMessageShowMessageBox_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageShowMessageBox", Erl)
End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessagePlayMidi(ByVal midi As Byte, _
    On Error Goto PrepareMessagePlayMidi_Err
                                       Optional ByVal loops As Integer = -1)
        
        On Error GoTo PrepareMessagePlayMidi_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePlayMIDI)
102     Call Writer.WriteInt8(midi)
104     Call Writer.WriteInt16(loops)
        
        Exit Function

PrepareMessagePlayMidi_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessagePlayMidi", Erl)
        
    Exit Function
PrepareMessagePlayMidi_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessagePlayMidi", Erl)
End Function

Public Function PrepareMessageOnlineUser(ByVal UserOnline As Integer)
    On Error Goto PrepareMessageOnlineUser_Err
        
        On Error GoTo PrepareMessageOnlineUser_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUserOnline)
102     Call Writer.WriteInt16(UserOnline)
        
        Exit Function

PrepareMessageOnlineUser_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageOnlineUser", Erl)
        
    Exit Function
PrepareMessageOnlineUser_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageOnlineUser", Erl)
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessagePauseToggle()
    On Error Goto PrepareMessagePauseToggle_Err
        
        On Error GoTo PrepareMessagePauseToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.ePauseToggle)
        
        Exit Function

PrepareMessagePauseToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessagePauseToggle", Erl)
        
    Exit Function
PrepareMessagePauseToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessagePauseToggle", Erl)
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageRainToggle()
    On Error Goto PrepareMessageRainToggle_Err
        
        On Error GoTo PrepareMessageRainToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eRainToggle)
        Call Writer.WriteBool(Lloviendo)
        
        Exit Function

PrepareMessageRainToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageRainToggle", Erl)
        
    Exit Function
PrepareMessageRainToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageRainToggle", Erl)
End Function

Public Function PrepareMessageHora()
    On Error Goto PrepareMessageHora_Err
        
        On Error GoTo PrepareMessageHora_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eHora)
102     Call Writer.WriteInt32(CLng((GetTickCount() - HoraMundo) Mod CLng(SvrConfig.GetValue("DayLength"))))
104     Call Writer.WriteInt32(CLng(SvrConfig.GetValue("DayLength")))
        
        Exit Function

PrepareMessageHora_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageHora", Erl)
        
    Exit Function
PrepareMessageHora_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageHora", Erl)
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte)
    On Error Goto PrepareMessageObjectDelete_Err
        
        On Error GoTo PrepareMessageObjectDelete_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eObjectDelete)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
        
        Exit Function

PrepareMessageObjectDelete_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageObjectDelete", Erl)
        
    Exit Function
PrepareMessageObjectDelete_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageObjectDelete", Erl)
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessage_BlockPosition(ByVal X As Byte, _
    On Error Goto PrepareMessage_BlockPosition_Err
                                            ByVal Y As Byte, _
                                            ByVal Blocked As Byte)
        
        On Error GoTo PrepareMessage_BlockPosition_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBlockPosition)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt8(Blocked)
        
        Exit Function

PrepareMessage_BlockPosition_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessage_BlockPosition", Erl)
        
    Exit Function
PrepareMessage_BlockPosition_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessage_BlockPosition", Erl)
End Function

Public Function PrepareTrapUpdate(ByVal State As Byte, ByVal x As Byte, ByVal y As Byte)
    On Error Goto PrepareTrapUpdate_Err
    On Error GoTo PrepareTrapUpdate_Err
100     Call Writer.WriteInt16(ServerPacketID.eUpdateTrap)
102     Call Writer.WriteInt8(State)
104     Call Writer.WriteInt8(x)
106     Call Writer.WriteInt8(y)

        Exit Function
PrepareTrapUpdate_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareTrapUpdate", Erl)
    Exit Function
PrepareTrapUpdate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareTrapUpdate", Erl)
End Function

Public Function PrepareUpdateGroupInfo(ByVal UserIndex As Integer)
    On Error Goto PrepareUpdateGroupInfo_Err
    On Error GoTo PrepareTrapUpdate_Err
100 Call Writer.WriteInt16(ServerPacketID.eUpdateGroupInfo)
    If IsValidUserRef(UserList(UserIndex).Grupo.Lider) Then
        With UserList(UserList(UserIndex).Grupo.Lider.ArrayIndex).Grupo
            Dim i As Integer
            Writer.WriteInt8 (.CantidadMiembros)
            For i = 1 To .CantidadMiembros
                Writer.WriteString8 (UserList(.Miembros(i).ArrayIndex).name)
                Writer.WriteInt16 (UserList(.Miembros(i).ArrayIndex).Char.charindex)
                Writer.WriteInt16 (UserList(.Miembros(i).ArrayIndex).Char.head)
                Writer.WriteInt16 (UserList(.Miembros(i).ArrayIndex).Stats.MinHp)
                Writer.WriteInt16 (UserList(.Miembros(i).ArrayIndex).Stats.MaxHp)
            Next i
        End With
    Else
        Writer.WriteInt8 (0)
    End If
    Exit Function
PrepareTrapUpdate_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareTrapUpdate", Erl)
    Exit Function
PrepareUpdateGroupInfo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareUpdateGroupInfo", Erl)
End Function
''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
'Optimizacion por Ladder
Public Function PrepareMessageObjectCreate(ByVal ObjIndex As Integer, _
    On Error Goto PrepareMessageObjectCreate_Err
                                           ByVal amount As Integer, _
                                           ByVal X As Byte, _
                                           ByVal y As Byte, _
                                           Optional ByVal ElementalTags As Long = e_ElementalTags.Normal)
        On Error GoTo PrepareMessageObjectCreate_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eObjectCreate)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt16(ObjIndex)
108     Call Writer.WriteInt16(amount)
        Call Writer.WriteInt32(ElementalTags)
        
        Exit Function

PrepareMessageObjectCreate_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageObjectCreate", Erl)
        
    Exit Function
PrepareMessageObjectCreate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageObjectCreate", Erl)
End Function

Public Function PrepareMessageFxPiso(ByVal GrhIndex As Integer, _
    On Error Goto PrepareMessageFxPiso_Err
                                     ByVal X As Byte, _
                                     ByVal Y As Byte)
        
        On Error GoTo PrepareMessageFxPiso_Err
        
100     Call Writer.WriteInt16(ServerPacketID.efxpiso)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt16(GrhIndex)
        
        Exit Function

PrepareMessageFxPiso_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageFxPiso", Erl)
        
    Exit Function
PrepareMessageFxPiso_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageFxPiso", Erl)
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageCharacterRemove(ByVal dbgid As Integer, ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageCharacterRemove_Err
                                              ByVal Desvanecido As Boolean, _
                                              Optional ByVal FueWarp As Boolean = False)
        
        On Error GoTo PrepareMessageCharacterRemove_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCharacterRemove)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteBool(Desvanecido)
106     Call Writer.WriteBool(FueWarp)
        
        Exit Function

PrepareMessageCharacterRemove_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterRemove", Erl)
        
    Exit Function
PrepareMessageCharacterRemove_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageCharacterRemove", Erl)
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer)
    On Error Goto PrepareMessageRemoveCharDialog_Err
        
        On Error GoTo PrepareMessageRemoveCharDialog_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eRemoveCharDialog)
102     Call Writer.WriteInt16(CharIndex)
        
        Exit Function

PrepareMessageRemoveCharDialog_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageRemoveCharDialog", Erl)
        
    Exit Function
PrepareMessageRemoveCharDialog_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageRemoveCharDialog", Erl)
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data .incomingData.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal head As Integer, ByVal Heading As e_Heading, _
    On Error Goto PrepareMessageCharacterCreate_Err
                                              ByVal charindex As Integer, ByVal x As Byte, ByVal y As Byte, ByVal weapon As Integer, _
                                              ByVal shield As Integer, ByVal Cart As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, _
                                              ByVal name As String, ByVal Status As Byte, ByVal privileges As Byte, ByVal ParticulaFx As Byte, ByVal Head_Aura As String, _
                                              ByVal Arma_Aura As String, ByVal Body_Aura As String, ByVal DM_Aura As String, ByVal RM_Aura As String, _
                                              ByVal Otra_Aura As String, ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Byte, _
                                              ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, _
                                              ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal UserMinMAN As Long, ByVal UserMaxMAN As Long, ByVal Simbolo As Byte, ByVal Idle As Boolean, _
                                              ByVal Navegando As Boolean, ByVal tipoUsuario As e_TipoUsuario, Optional ByVal TeamCaptura As Byte = 0, Optional ByVal TieneBandera As Byte = 0, Optional ByVal AnimAtaque1 As Integer = 0)
        
        On Error GoTo PrepareMessageCharacterCreate_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCharacterCreate)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(Body)
106     Call Writer.WriteInt16(Head)
108     Call Writer.WriteInt8(Heading)
110     Call Writer.WriteInt8(X)
112     Call Writer.WriteInt8(Y)
114     Call Writer.WriteInt16(weapon)
116     Call Writer.WriteInt16(shield)
118     Call Writer.WriteInt16(helmet)
119     Call Writer.WriteInt16(Cart)
120     Call Writer.WriteInt16(FX)
122     Call Writer.WriteInt16(FXLoops)
124     Call Writer.WriteString8(Name)
126     Call Writer.WriteInt8(Status)
128     Call Writer.WriteInt8(privileges)
130     Call Writer.WriteInt8(ParticulaFx)
132     Call Writer.WriteString8(Head_Aura)
134     Call Writer.WriteString8(Arma_Aura)
136     Call Writer.WriteString8(Body_Aura)
138     Call Writer.WriteString8(DM_Aura)
140     Call Writer.WriteString8(RM_Aura)
142     Call Writer.WriteString8(Otra_Aura)
144     Call Writer.WriteString8(Escudo_Aura)
146     Call Writer.WriteReal32(speeding)
148     Call Writer.WriteInt8(EsNPC)
150     Call Writer.WriteInt8(appear)
152     Call Writer.WriteInt16(group_index)
154     Call Writer.WriteInt16(clan_index)
156     Call Writer.WriteInt8(clan_nivel)
158     Call Writer.WriteInt32(UserMinHp)
160     Call Writer.WriteInt32(UserMaxHp)
162     Call Writer.WriteInt32(UserMinMAN)
164     Call Writer.WriteInt32(UserMaxMAN)
166     Call Writer.WriteInt8(Simbolo)
        Dim flags As Byte
        flags = 0
        If Idle Then flags = flags Or &O1 ' 00000001
        If Navegando Then flags = flags Or &O2
        Call Writer.WriteInt8(flags)
172     Call Writer.WriteInt8(tipoUsuario)
173     Call Writer.WriteInt8(TeamCaptura)
174     Call Writer.WriteInt8(TieneBandera)
175     Call Writer.WriteInt16(AnimAtaque1)
        
        Exit Function

PrepareMessageCharacterCreate_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterCreate", Erl)
        
    Exit Function
PrepareMessageCharacterCreate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageCharacterCreate", Erl)
End Function


''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageCharacterChange(ByVal Body As Integer, _
    On Error Goto PrepareMessageCharacterChange_Err
                                              ByVal Head As Integer, _
                                              ByVal Heading As e_Heading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal Cart As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal Idle As Boolean, _
                                              ByVal Navegando As Boolean)

        On Error GoTo PrepareMessageCharacterChange_Err

100     Call Writer.WriteInt16(ServerPacketID.eCharacterChange)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(Body)
106     Call Writer.WriteInt16(Head)
108     Call Writer.WriteInt8(Heading)
110     Call Writer.WriteInt16(weapon)
112     Call Writer.WriteInt16(shield)
114     Call Writer.WriteInt16(helmet)
116     Call Writer.WriteInt16(Cart)
118     Call Writer.WriteInt16(FX)
120     Call Writer.WriteInt16(FXLoops)
        Dim flags As Byte
        flags = 0
122     If Idle Then flags = flags Or &O1
124     If Navegando Then flags = flags Or &O2
126     Call Writer.WriteInt8(flags)
        Exit Function

PrepareMessageCharacterChange_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterChange", Erl)
        
    Exit Function
PrepareMessageCharacterChange_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageCharacterChange", Erl)
End Function


''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageUpdateFlag(ByVal Flag As Byte, ByVal charindex As Integer)
    On Error Goto PrepareMessageUpdateFlag_Err

        On Error GoTo PrepareMessageUpdateFlag_Err

100     Call Writer.WriteInt16(ServerPacketID.eUpdateFlag)
        Call Writer.WriteInt16(charindex)
102     Call Writer.WriteInt8(Flag)
        Exit Function

PrepareMessageUpdateFlag_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageUpdateFlag", Erl)
        
    Exit Function
PrepareMessageUpdateFlag_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageUpdateFlag", Erl)
End Function
''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageCharacterMove_Err
                                            ByVal X As Byte, _
                                            ByVal Y As Byte)
        
        On Error GoTo PrepareMessageCharacterMove_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eCharacterMove)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt8(X)
106     Call Writer.WriteInt8(Y)
        
        Exit Function

PrepareMessageCharacterMove_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterMove", Erl)
        
    Exit Function
PrepareMessageCharacterMove_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageCharacterMove", Erl)
End Function

Public Function PrepareCharacterTranslate(ByVal CharIndexm As Integer, ByVal NewX As Byte, ByVal NewY As Byte, ByVal TranslationTime As Long)
    On Error Goto PrepareCharacterTranslate_Err
    On Error GoTo PrepareMessageCharacterMove_Err
        Call Writer.WriteInt16(ServerPacketID.eCharacterTranslate)
        Call Writer.WriteInt16(CharIndexm)
        Call Writer.WriteInt8(NewX)
        Call Writer.WriteInt8(NewY)
        Call Writer.WriteInt32(TranslationTime)
    Exit Function
PrepareMessageCharacterMove_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterMove", Erl)
    Exit Function
PrepareCharacterTranslate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareCharacterTranslate", Erl)
End Function


Public Function PrepareMessageForceCharMove(ByVal Direccion As e_Heading)
    On Error Goto PrepareMessageForceCharMove_Err
        
        On Error GoTo PrepareMessageForceCharMove_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eForceCharMove)
102     Call Writer.WriteInt8(Direccion)
        
        Exit Function

PrepareMessageForceCharMove_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageForceCharMove", Erl)
        
    Exit Function
PrepareMessageForceCharMove_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageForceCharMove", Erl)
End Function

Public Function PrepareMessageForceCharMoveSiguiendo(ByVal Direccion As e_Heading)
    On Error Goto PrepareMessageForceCharMoveSiguiendo_Err
        
        On Error GoTo PrepareMessageForceCharMoveSiguiendo_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eForceCharMoveSiguiendo)
102     Call Writer.WriteInt8(Direccion)
        
        Exit Function

PrepareMessageForceCharMoveSiguiendo_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageForceCharMoveSiguiendo", Erl)
        
    Exit Function
PrepareMessageForceCharMoveSiguiendo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageForceCharMoveSiguiendo", Erl)
End Function
''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, _
    On Error Goto PrepareMessageUpdateTagAndStatus_Err
                                                 Status As Byte, _
                                                 Tag As String)
        
        On Error GoTo PrepareMessageUpdateTagAndStatus_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eUpdateTagAndStatus)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
104     Call Writer.WriteInt8(Status)
106     Call Writer.WriteString8(Tag)
108     Call Writer.WriteInt16(UserList(userIndex).Grupo.Lider.ArrayIndex)
        
        Exit Function

PrepareMessageUpdateTagAndStatus_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageUpdateTagAndStatus", Erl)
        
    Exit Function
PrepareMessageUpdateTagAndStatus_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageUpdateTagAndStatus", Erl)
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageErrorMsg(ByVal Message As String)
    On Error Goto PrepareMessageErrorMsg_Err
        
        On Error GoTo PrepareMessageErrorMsg_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eErrorMsg)
102     Call Writer.WriteString8(Message)
        
        Exit Function

PrepareMessageErrorMsg_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageErrorMsg", Erl)
        
    Exit Function
PrepareMessageErrorMsg_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageErrorMsg", Erl)
End Function

Public Function PrepareMessageBarFx(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageBarFx_Err
                                    ByVal BarTime As Integer, _
                                    ByVal BarAccion As Byte)
        
        On Error GoTo PrepareMessageBarFx_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eBarFx)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(BarTime)
106     Call Writer.WriteInt8(BarAccion)
        
        Exit Function

PrepareMessageBarFx_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageBarFx", Erl)
        
    Exit Function
PrepareMessageBarFx_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageBarFx", Erl)
End Function

Public Function PrepareMessageNieblandoToggle(ByVal IntensidadMax As Byte)
    On Error Goto PrepareMessageNieblandoToggle_Err
        
        On Error GoTo PrepareMessageNieblandoToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eNieblaToggle)
102     Call Writer.WriteInt8(IntensidadMax)
        
        Exit Function

PrepareMessageNieblandoToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageNieblandoToggle", Erl)
        
    Exit Function
PrepareMessageNieblandoToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageNieblandoToggle", Erl)
End Function

Public Function PrepareMessageNevarToggle()
    On Error Goto PrepareMessageNevarToggle_Err
        
        On Error GoTo PrepareMessageNevarToggle_Err
        
100     Call Writer.WriteInt16(ServerPacketID.eNieveToggle)
        Call Writer.WriteBool(Nebando)
        
        Exit Function

PrepareMessageNevarToggle_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageNevarToggle", Erl)
        
    Exit Function
PrepareMessageNevarToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageNevarToggle", Erl)
End Function

Public Function PrepareMessageDoAnimation(ByVal CharIndex As Integer, _
    On Error Goto PrepareMessageDoAnimation_Err
                                          ByVal Animation As Integer)

        On Error GoTo PrepareMessageDoAnimation_Err

100     Call Writer.WriteInt16(ServerPacketID.eDoAnimation)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(Animation)

        Exit Function

PrepareMessageDoAnimation_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageDoAnimation", Erl)
    Exit Function
PrepareMessageDoAnimation_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareMessageDoAnimation", Erl)
End Function

'Public Function WritePescarEspecial(ByVal ObjIndex As Integer)

'        On Error GoTo PescarEspecial_Err
'100     Call Writer.WriteInt16(ServerPacketID.PescarEspecial)
'        Call Writer.WriteInt16(ObjIndex)

'PescarEspecial_Err:
'        Call Writer.Clear
'        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PescarEspecial", Erl)
'End Function
Public Sub writeAnswerReset(ByVal UserIndex As Integer)
    On Error Goto writeAnswerReset_Err
    On Error GoTo writeAnswerReset_Err

    Call Writer.WriteInt16(ServerPacketID.eAnswerReset)

182     Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
writeAnswerReset_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.writeAnswerReset", Erl)
        
    Exit Sub
writeAnswerReset_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.writeAnswerReset", Erl)
End Sub

Public Sub WriteShopInit(ByVal UserIndex As Integer)
    On Error Goto WriteShopInit_Err
    On Error GoTo WriteShopInit_Err
    Dim i As Long, cant_obj_shop As Integer
    Call Writer.WriteInt16(ServerPacketID.eShopInit)
    cant_obj_shop = UBound(ObjShop)
    Call Writer.WriteInt16(cant_obj_shop)
    Call LoadPatronCreditsFromDB(UserIndex)
    Call Writer.WriteInt32(UserList(userindex).Stats.Creditos)
    
    'Envío todos los objetos.
    For i = 1 To cant_obj_shop
        Call Writer.WriteInt32(ObjShop(i).ObjNum)
        Call Writer.WriteInt32(ObjShop(i).valor)
        Call Writer.WriteString8(ObjShop(i).Name)
    Next i
    
182 Call modSendData.SendData(ToIndex, UserIndex)
   Exit Sub
WriteShopInit_Err:
     Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShopInit", Erl)
    Exit Sub
WriteShopInit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShopInit", Erl)
End Sub


Public Sub WriteShopPjsInit(ByVal UserIndex As Integer)
    On Error Goto WriteShopPjsInit_Err
    On Error GoTo WriteShopPjsInit_Err
   
    Call Writer.WriteInt16(ServerPacketID.eShopPjsInit)
    
182 Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShopPjsInit_Err:
     Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShopPjsInit", Erl)
    Exit Sub
WriteShopPjsInit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteShopPjsInit", Erl)
End Sub

Public Sub writeUpdateShopClienteCredits(ByVal userindex As Integer)
    On Error Goto writeUpdateShopClienteCredits_Err
On Error GoTo WriteUpdateShopClienteCredits_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateShopClienteCredits)
    Call Writer.WriteInt32(UserList(userindex).Stats.Creditos)
182 Call modSendData.SendData(ToIndex, userindex)
    Exit Sub
writeUpdateShopClienteCredits_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description + " UI: " + UserIndex, "Argentum20Server.Protocol_Writes.writeUpdateShopClienteCredits", Erl)
    Exit Sub
writeUpdateShopClienteCredits_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.writeUpdateShopClienteCredits", Erl)
End Sub

Public Sub WriteSendSkillCdUpdate(ByVal UserIndex As Integer, ByVal SkillTypeId As Integer, ByVal SkillId As Long, ByVal TimeLeft As Long, ByVal TotalTime As Long, ByVal SkillType As e_EffectType, Optional ByVal Stacks As Integer = 1)
    On Error Goto WriteSendSkillCdUpdate_Err
On Error GoTo WriteSendSkillCdUpdate_Err
    Call Writer.WriteInt16(ServerPacketID.eSendSkillCdUpdate)
    Call Writer.WriteInt16(SkillTypeId)
    Call Writer.WriteInt32(SkillId)
    Call Writer.WriteInt32(TimeLeft)
    Call Writer.WriteInt32(TotalTime)
    Call Writer.WriteInt8(ConvertToClientBuff(SkillType))
    Call Writer.WriteInt16(Stacks)
182 Call modSendData.SendData(ToIndex, userindex)
    Exit Sub
WriteSendSkillCdUpdate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description + " UI: " + UserIndex, "Argentum20Server.Protocol_Writes.writeUpdateShopClienteCredits", Erl)
    Exit Sub
WriteSendSkillCdUpdate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteSendSkillCdUpdate", Erl)
End Sub

Public Sub WriteObjQuestSend(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal Slot As Byte)
    On Error Goto WriteObjQuestSend_Err
        
        On Error GoTo WriteNpcQuestListSend_Err
        
        Dim i As Integer

100     Call Writer.WriteInt16(ServerPacketID.eObjQuestListSend)
102     Call Writer.WriteInt16(QuestIndex) 'Escribimos primero cuantas quest tiene el NPC

110     Call Writer.WriteInt8(QuestList(QuestIndex).RequiredLevel)
112     Call Writer.WriteInt16(QuestList(QuestIndex).RequiredQuest)
            'Enviamos la cantidad de npcs requeridos
114         Call Writer.WriteInt8(QuestList(QuestIndex).RequiredNPCs)

116     If QuestList(QuestIndex).RequiredNPCs Then
                'Si hay npcs entonces enviamos la lista
118         For i = 1 To QuestList(QuestIndex).RequiredNPCs
120             Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).amount)
122             Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)
124         Next i
        End If

            'Enviamos la cantidad de objs requeridos
126     Call Writer.WriteInt8(QuestList(QuestIndex).RequiredOBJs)

128     If QuestList(QuestIndex).RequiredOBJs Then
            'Si hay objs entonces enviamos la lista
130     For i = 1 To QuestList(QuestIndex).RequiredOBJs
132         Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).amount)
134         Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex)
136     Next i
        End If
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSpellCount)
        If QuestList(QuestIndex).RequiredSpellCount > 0 Then
            For i = 1 To QuestList(QuestIndex).RequiredSpellCount
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredSpellList(i))
            Next i
        End If
137     Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.SkillType)
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.RequiredValue)
            'Enviamos la recompensa de oro y experiencia.
138     Call Writer.WriteInt32(QuestList(QuestIndex).RewardGLD * SvrConfig.GetValue("GoldMult"))
140     Call Writer.WriteInt32(QuestList(QuestIndex).RewardEXP * SvrConfig.GetValue("ExpMult"))
            'Enviamos la cantidad de objs de recompensa
142     Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)

144     If QuestList(QuestIndex).RewardOBJs Then

                'si hay objs entonces enviamos la lista
146         For i = 1 To QuestList(QuestIndex).RewardOBJs
148             Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).amount)
150             Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
152         Next i

        End If

            'Enviamos el estado de la QUEST
            '0 Disponible
            '1 EN CURSO
            '2 REALIZADA
            '3 no puede hacerla
            Dim PuedeHacerla As Boolean

            'La tiene aceptada el usuario?
154         If TieneQuest(UserIndex, QuestIndex) Then
156             Call Writer.WriteInt8(1)
            Else

158             If UserDoneQuest(UserIndex, QuestIndex) Then
160                 Call Writer.WriteInt8(2)
                Else
162                 PuedeHacerla = True

164                 If QuestList(QuestIndex).RequiredQuest > 0 Then
166                     If Not UserDoneQuest(UserIndex, QuestList( _
                                QuestIndex).RequiredQuest) Then
168                         PuedeHacerla = False
                        End If
                    End If

170                 If UserList(UserIndex).Stats.ELV < QuestList(QuestIndex).RequiredLevel _
                            Then
172                     PuedeHacerla = False
                    End If

174                 If PuedeHacerla Then
176                     Call Writer.WriteInt8(0)
                    Else
178                     Call Writer.WriteInt8(3)
                    End If
                End If
            End If
        UserList(UserIndex).flags.QuestNumber = QuestIndex
        UserList(UserIndex).flags.QuestItemSlot = Slot
        UserList(UserIndex).flags.QuestOpenByObj = True
182     Call modSendData.SendData(ToIndex, UserIndex)
        
        Exit Sub

WriteNpcQuestListSend_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNpcQuestListSend for quest: " & QuestIndex, Erl)
        
    Exit Sub
WriteObjQuestSend_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteObjQuestSend", Erl)
End Sub

Public Sub WriteDebugLogResponse(ByVal UserIndex As Integer, ByVal debugType, ByRef args() As String, ByVal argc As Integer)
    On Error Goto WriteDebugLogResponse_Err
    On Error GoTo WriteDebugLogResponse_Err:
    Call Writer.WriteInt16(ServerPacketID.eDebugDataResponse)
    
    If debugType = 0 Then
        Dim messageList() As String
        messageList = GetLastMessages()
        Dim messageCount As Integer: messageCount = UBound(messageList)
        
        Call Writer.WriteInt16(messageCount + 1)
        Call Writer.WriteString8("remote errors:")
        Dim i As Integer
        For i = 1 To messageCount
            Call Writer.WriteString8(messageList(i))
        Next i
    ElseIf debugType = 1 Then
        'TODO- debug
        Dim tIndex As Integer: tIndex = NameIndex(Args(0)).ArrayIndex
        If tIndex > 0 Then
            Call Writer.WriteInt16(2)
            Call Writer.WriteString8("remote DEBUG: " & " user name: " & args(0))
            With UserList(tIndex)
                Dim timeSinceLastReset As Long
                timeSinceLastReset = GetTickCount() - Mapping(.ConnectionDetails.ConnID).TimeLastReset
                Call Writer.WriteString8("validConnection: " & .ConnectionDetails.ConnIDValida & " connectionID: " & .ConnectionDetails.ConnID & " UserIndex: " & tIndex & " charNmae" & .name & " UserLogged state: " & .flags.UserLogged & ", time since last message: " & timeSinceLastReset & " timeout setting: " & DisconnectTimeout)
            End With
        Else
            Call Writer.WriteInt16(1)
        Call Writer.WriteString8("DEBUG: failed to find user: " & args(0))
        End If
    ElseIf debugType = 2 Then
            Call Writer.WriteInt16(1)
            Call Writer.WriteString8("remote DEBUG: avialable user slots: " & GetAvailableUserSlot & ", LastUser: " & LastUser)
    End If
    
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteDebugLogResponse_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDebugLogResponse", Erl)
    Exit Sub
WriteDebugLogResponse_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteDebugLogResponse", Erl)
End Sub

Public Function PrepareUpdateCharValue(ByVal Charindex As Integer, _
    On Error Goto PrepareUpdateCharValue_Err
                                       ByVal CharValueType As e_CharValue, _
                                       ByVal NewValue As Long)

        On Error GoTo PrepareMessageDoAnimation_Err

100     Call Writer.WriteInt16(ServerPacketID.eUpdateCharValue)
102     Call Writer.WriteInt16(Charindex)
104     Call Writer.WriteInt16(CharValueType)
106     Call Writer.WriteInt32(NewValue)

        Exit Function

PrepareMessageDoAnimation_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageDoAnimation", Erl)
    Exit Function
PrepareUpdateCharValue_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareUpdateCharValue", Erl)
End Function

Public Function PrepareActiveToggles()
    On Error Goto PrepareActiveToggles_Err
    On Error GoTo PrepareActiveToggles_Err
100     Call Writer.WriteInt16(ServerPacketID.eSendClientToggles)
        Dim ActiveToggles() As String
        Dim ActiveToggleCount As Integer
        ActiveToggles = GetActiveToggles(ActiveToggleCount)
        Call Writer.WriteInt16(ActiveToggleCount)
        Dim i As Integer
        For i = 0 To ActiveToggleCount - 1
            Call Writer.WriteString8(ActiveToggles(i))
        Next i
        Exit Function
PrepareActiveToggles_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareActiveToggles", Erl)
    Exit Function
PrepareActiveToggles_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.PrepareActiveToggles", Erl)
End Function

Public Sub WriteAntiCheatMessage(ByVal UserIndex As Integer, ByVal Data As Long, ByVal DataSize As Long)
    On Error Goto WriteAntiCheatMessage_Err
    On Error GoTo WriteAntiCheatMessage_Err
        Dim Buffer() As Byte
        ReDim Buffer(0 To (DataSize - 1)) As Byte
        CopyMemory Buffer(0), ByVal Data, DataSize
        Call Writer.WriteInt16(ServerPacketID.eAntiCheatMessage)
        Call Writer.WriteSafeArrayInt8(Buffer)
        Call modSendData.SendData(ToIndex, UserIndex)
        Exit Sub
WriteAntiCheatMessage_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAntiCheatMessage", Erl)
    Exit Sub
WriteAntiCheatMessage_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteAntiCheatMessage", Erl)
End Sub

Public Sub WriteAntiCheatStartSeassion(ByVal UserIndex As Integer)
    On Error Goto WriteAntiCheatStartSeassion_Err
    On Error GoTo WriteAntiStartSeassion_Err
        Call Writer.WriteInt16(ServerPacketID.eAntiCheatStartSession)
        Call modSendData.SendData(ToIndex, UserIndex)
        Exit Sub
WriteAntiStartSeassion_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAntiStartSeassion", Erl)
    Exit Sub
WriteAntiCheatStartSeassion_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteAntiCheatStartSeassion", Erl)
End Sub

Public Sub WriteUpdateLobbyList(ByVal UserIndex As Integer)
    On Error Goto WriteUpdateLobbyList_Err
On Error GoTo WriteUpdateLobbyList_Err
    Dim IdList() As Integer
    Dim OpenLobbyCount As Integer
    OpenLobbyCount = GetOpenLobbyList(IdList)
    Dim i As Integer
    Call Writer.WriteInt16(ServerPacketID.eReportLobbyList)
    Call Writer.WriteInt16(OpenLobbyCount)
    For i = 0 To OpenLobbyCount - 1
        Call Writer.WriteInt16(IdList(i))
        With LobbyList(IdList(i))
            Call Writer.WriteString8(.Description)
            If .Scenario Is Nothing Then
                Call Writer.WriteString8("")
            Else
                Call Writer.WriteString8(.Scenario.GetScenarioName())
            End If
            Call Writer.WriteInt16(.MinLevel)
            Call Writer.WriteInt16(.MaxLevel)
            Call Writer.WriteInt16(.MinPlayers)
            Call Writer.WriteInt16(.MaxPlayers)
            Call Writer.WriteInt16(.RegisteredPlayers)
            Call Writer.WriteInt16(.TeamSize)
            Call Writer.WriteInt16(.TeamType)
            Call Writer.WriteInt32(.InscriptionPrice)
            Call Writer.WriteInt8(IIf(Len(.Password) > 0, 1, 0))
        End With
    Next i
    
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateLobbyList_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareActiveToggles", Erl)
    Exit Sub
WriteUpdateLobbyList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol_Writes.WriteUpdateLobbyList", Erl)
End Sub
