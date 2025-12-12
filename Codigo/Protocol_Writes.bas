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
    Private Writer As Network.Writer

Public Sub InitializeAuxiliaryBuffer()
    Set Writer = New Network.Writer
End Sub
    
Public Function GetWriterBuffer() As Network.Writer
    Set GetWriterBuffer = Writer
End Function

#Else
    Public Writer As New clsNetWriter

#End If
#If PYMMO = 0 Then
    Public Sub WriteAccountCharacterList(ByVal UserIndex As Integer, ByRef Personajes() As t_PersonajeCuenta, ByVal count As Long)
        On Error GoTo WriteAccountCharacterList_Err
        Call Writer.WriteInt16(ServerPacketID.eAccountCharacterList)
        Call Writer.WriteInt(count)
        Dim i As Long
        For i = 1 To count
            With Personajes(i)
                Call Writer.WriteString8(.nombre)
                Call Writer.WriteInt(.cuerpo)
                Call Writer.WriteInt(.Cabeza)
                Call Writer.WriteInt(.clase)
                Call Writer.WriteInt(.Mapa)
                Call Writer.WriteInt(.PosX)
                Call Writer.WriteInt(.PosY)
                Call Writer.WriteInt(.nivel)
                Call Writer.WriteInt(.Status)
                Call Writer.WriteInt(.Casco)
                Call Writer.WriteInt(.Escudo)
                Call Writer.WriteInt(.Arma)
                Call Writer.WriteInt(.BackPack)
            End With
        Next i
        Call modSendData.SendData(ToIndex, UserIndex)
        Exit Sub
WriteAccountCharacterList_Err:
        Call Writer.Clear
        Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAccountCharacterList", Erl)
    End Sub

#End If
' \Begin: [Writes]
Public Function PrepareConnected()
    On Error GoTo WriteConnected_Err
    Call Writer.WriteInt16(ServerPacketID.eConnected)
    #If DEBUGGING = 1 Then
        Dim i               As Integer
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
End Function

''
' Writes the "Logged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoggedMessage(ByVal UserIndex As Integer, Optional ByVal newUser As Boolean = False)
    On Error GoTo WriteLoggedMessage_Err
    Call Writer.WriteInt16(ServerPacketID.elogged)
    Call Writer.WriteBool(newUser)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteLoggedMessage_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLoggedMessage", Erl)
End Sub

Public Sub WriteHora(ByVal UserIndex As Integer)
    On Error GoTo WriteHora_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageHora())
    Exit Sub
WriteHora_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteHora", Erl)
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
    On Error GoTo WriteRemoveAllDialogs_Err
    Call Writer.WriteInt16(ServerPacketID.eRemoveDialogs)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteRemoveAllDialogs_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRemoveAllDialogs", Erl)
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal charindex As Integer)
    On Error GoTo WriteRemoveCharDialog_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageRemoveCharDialog(charindex))
    Exit Sub
WriteRemoveCharDialog_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRemoveCharDialog", Erl)
End Sub

' Writes the "NavigateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNavigateToggle(ByVal UserIndex As Integer, ByVal NewState As Boolean)
    On Error GoTo WriteNavigateToggle_Err
    Call Writer.WriteInt16(ServerPacketID.eNavigateToggle)
    Call Writer.WriteBool(NewState)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteNavigateToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNavigateToggle", Erl)
End Sub

Public Sub WriteNadarToggle(ByVal UserIndex As Integer, ByVal Puede As Boolean, Optional ByVal esTrajeCaucho As Boolean = False)
    On Error GoTo WriteNadarToggle_Err
    Call Writer.WriteInt16(ServerPacketID.eNadarToggle)
    Call Writer.WriteBool(Puede)
    Call Writer.WriteBool(esTrajeCaucho)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteNadarToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNadarToggle", Erl)
End Sub

Public Sub WriteEquiteToggle(ByVal UserIndex As Integer)
    On Error GoTo WriteEquiteToggle_Err
    Call Writer.WriteInt16(ServerPacketID.eEquiteToggle)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteEquiteToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteEquiteToggle", Erl)
End Sub

Public Sub WriteVelocidadToggle(ByVal UserIndex As Integer)
    On Error GoTo WriteVelocidadToggle_Err
    Call Writer.WriteInt16(ServerPacketID.eVelocidadToggle)
    Call Writer.WriteReal32(UserList(UserIndex).Char.speeding)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteVelocidadToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteVelocidadToggle", Erl)
End Sub

Public Sub WriteMacroTrabajoToggle(ByVal UserIndex As Integer, ByVal Activar As Boolean)
    On Error GoTo WriteMacroTrabajoToggle_Err
    If Not Activar Then
        UserList(UserIndex).flags.TargetObj = 0 ' Sacamos el targer del objeto
        UserList(UserIndex).flags.UltimoMensaje = 0
        UserList(UserIndex).Counters.Trabajando = 0
        UserList(UserIndex).flags.UsandoMacro = False
        UserList(UserIndex).Trabajo.Target_X = 0
        UserList(UserIndex).Trabajo.Target_Y = 0
        UserList(UserIndex).Trabajo.TargetSkill = 0
        UserList(UserIndex).Trabajo.Cantidad = 0
        UserList(UserIndex).Trabajo.Item = 0
    Else
        UserList(UserIndex).flags.UsandoMacro = True
    End If
    Call Writer.WriteInt16(ServerPacketID.eMacroTrabajoToggle)
    Call Writer.WriteBool(Activar)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteMacroTrabajoToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMacroTrabajoToggle", Erl)
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDisconnect(ByVal UserIndex As Integer, Optional ByVal FullLogout As Boolean = False)
    On Error GoTo WriteDisconnect_Err
    Call ClearAndSaveUser(UserIndex)
    UserList(UserIndex).flags.YaGuardo = True
    Call Writer.WriteInt16(ServerPacketID.eDisconnect)
    Call Writer.WriteBool(FullLogout)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteDisconnect_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDisconnect", Erl)
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
    On Error GoTo WriteCommerceEnd_Err
    Call Writer.WriteInt16(ServerPacketID.eCommerceEnd)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCommerceEnd_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCommerceEnd", Erl)
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankEnd(ByVal UserIndex As Integer)
    On Error GoTo WriteBankEnd_Err
    Call Writer.WriteInt16(ServerPacketID.eBankEnd)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteBankEnd_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBankEnd", Erl)
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
    On Error GoTo WriteCommerceInit_Err
    Call Writer.WriteInt16(ServerPacketID.eCommerceInit)
    Call Writer.WriteString8(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).name)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCommerceInit_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCommerceInit", Erl)
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankInit(ByVal UserIndex As Integer)
    On Error GoTo WriteBankInit_Err
    Call Writer.WriteInt16(ServerPacketID.eBankInit)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteBankInit_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBankInit", Erl)
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
    On Error GoTo WriteUserCommerceInit_Err
    Call Writer.WriteInt16(ServerPacketID.eUserCommerceInit)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUserCommerceInit_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserCommerceInit", Erl)
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
    On Error GoTo WriteUserCommerceEnd_Err
    Call Writer.WriteInt16(ServerPacketID.eUserCommerceEnd)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUserCommerceEnd_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserCommerceEnd", Erl)
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)
    On Error GoTo WriteShowBlacksmithForm_Err
    Call Writer.WriteInt16(ServerPacketID.eShowBlacksmithForm)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowBlacksmithForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowBlacksmithForm", Erl)
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)
    On Error GoTo WriteShowCarpenterForm_Err
    Call Writer.WriteInt16(ServerPacketID.eShowCarpenterForm)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowCarpenterForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowCarpenterForm", Erl)
End Sub

Public Sub WriteShowAlquimiaForm(ByVal UserIndex As Integer)
    On Error GoTo WriteShowAlquimiaForm_Err
    Call Writer.WriteInt16(ServerPacketID.eShowAlquimiaForm)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowAlquimiaForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowAlquimiaForm", Erl)
End Sub

Public Sub WriteShowSastreForm(ByVal UserIndex As Integer)
    On Error GoTo WriteShowSastreForm_Err
    Call Writer.WriteInt16(ServerPacketID.eShowSastreForm)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowSastreForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowSastreForm", Erl)
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)
    On Error GoTo WriteNPCKillUser_Err
    Call Writer.WriteInt16(ServerPacketID.eNPCKillUser)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteNPCKillUser_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNPCKillUser", Erl)
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub Write_BlockedWithShieldUser(ByVal UserIndex As Integer)
    On Error GoTo Write_BlockedWithShieldUser_Err
    Call Writer.WriteInt16(ServerPacketID.eBlockedWithShieldUser)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
Write_BlockedWithShieldUser_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.Write_BlockedWithShieldUser", Erl)
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub Write_BlockedWithShieldOther(ByVal UserIndex As Integer)
    On Error GoTo Write_BlockedWithShieldOther_Err
    Call Writer.WriteInt16(ServerPacketID.eBlockedWithShieldOther)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
Write_BlockedWithShieldOther_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.Write_BlockedWithShieldOther", Erl)
End Sub

''
' Writes the "SafeModeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)
    On Error GoTo WriteSafeModeOn_Err
    Call Writer.WriteInt16(ServerPacketID.eSafeModeOn)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteSafeModeOn_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSafeModeOn", Erl)
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)
    On Error GoTo WriteSafeModeOff_Err
    Call Writer.WriteInt16(ServerPacketID.eSafeModeOff)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteSafeModeOff_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSafeModeOff", Erl)
End Sub

''
' Writes the "PartySafeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePartySafeOn(ByVal UserIndex As Integer)
    On Error GoTo WritePartySafeOn_Err
    Call Writer.WriteInt16(ServerPacketID.ePartySafeOn)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePartySafeOn_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePartySafeOn", Erl)
End Sub

''
' Writes the "PartySafeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePartySafeOff(ByVal UserIndex As Integer)
    On Error GoTo WritePartySafeOff_Err
    Call Writer.WriteInt16(ServerPacketID.ePartySafeOff)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePartySafeOff_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePartySafeOff", Erl)
End Sub

Public Sub WriteClanSeguro(ByVal UserIndex As Integer, ByVal Estado As Boolean)
    On Error GoTo WriteClanSeguro_Err
    Call Writer.WriteInt16(ServerPacketID.eClanSeguro)
    Call Writer.WriteBool(Estado)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteClanSeguro_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteClanSeguro", Erl)
End Sub

Public Sub WriteSeguroResu(ByVal UserIndex As Integer, ByVal Estado As Boolean)
    On Error GoTo WriteSeguroResu_Err
    Call Writer.WriteInt16(ServerPacketID.eSeguroResu)
    Call Writer.WriteBool(Estado)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteSeguroResu_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSeguroResu", Erl)
End Sub

Public Sub WriteLegionarySecure(ByVal UserIndex As Integer, ByVal Estado As Boolean)
    On Error GoTo WriteLegionarySecure_Err
    Call Writer.WriteInt16(ServerPacketID.eLegionarySecure)
    Call Writer.WriteBool(Estado)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteLegionarySecure_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLegionarySecure", Erl)
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)
    On Error GoTo WriteCantUseWhileMeditating_Err
    Call Writer.WriteInt16(ServerPacketID.eCantUseWhileMeditating)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCantUseWhileMeditating_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCantUseWhileMeditating", Erl)
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
    On Error GoTo WriteUpdateSta_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateSta)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateSta_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateSta", Erl)
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
    On Error GoTo WriteUpdateMana_Err
    Call SendData(SendTarget.ToAdminsYDioses, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateMAN(UserIndex))
    Call SendData(SendTarget.ToClanArea, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateMAN(UserIndex))
    Call Writer.WriteInt16(ServerPacketID.eUpdateMana)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateMana_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateMana", Erl)
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
    'Call SendData(SendTarget.ToDiosesYclan, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
    On Error GoTo WriteUpdateHP_Err
    Call SendData(SendTarget.ToAdminsYDioses, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToClanArea, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToGroupButIndex, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call Writer.WriteInt16(ServerPacketID.eUpdateHP)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.shield)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateHP_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateHP", Erl)
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
    On Error GoTo WriteUpdateGold_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateGold)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
    Call Writer.WriteInt32(SvrConfig.GetValue("OroPorNivelBilletera"))
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateGold_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateGold", Erl)
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
    On Error GoTo WriteUpdateExp_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateExp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateExp_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateExp", Erl)
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer)
    On Error GoTo WriteChangeMap_Err
    Call Writer.WriteInt16(ServerPacketID.eChangeMap)
    Call Writer.WriteInt16(Map)
    Call Writer.WriteInt16(MapInfo(Map).MapResource)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteChangeMap_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeMap", Erl)
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePosUpdate(ByVal UserIndex As Integer)
    On Error GoTo WritePosUpdate_Err
    Call Writer.WriteInt16(ServerPacketID.ePosUpdate)
    Call Writer.WriteInt8(UserList(UserIndex).pos.x)
    Call Writer.WriteInt8(UserList(UserIndex).pos.y)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePosUpdate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePosUpdate", Erl)
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePosUpdateCharIndex(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte, ByVal charindex As Integer)
    On Error GoTo WritePosUpdateCharIndex_Err
    Call Writer.WriteInt16(ServerPacketID.ePosUpdateUserChar)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt16(charindex)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePosUpdateCharIndex_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePosUpdateCharIndex", Erl)
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePosUpdateChar(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte, ByVal charindex As Integer)
    On Error GoTo WritePosUpdateChar_Err
    Call Writer.WriteInt16(ServerPacketID.ePosUpdateChar)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePosUpdateChar_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePosUpdateChar", Erl)
End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, ByVal Target As e_PartesCuerpo, ByVal Damage As Integer)
    On Error GoTo WriteNPCHitUser_Err
    Call Writer.WriteInt16(ServerPacketID.eNPCHitUser)
    Call Writer.WriteInt8(Target)
    Call Writer.WriteInt16(Damage)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteNPCHitUser_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNPCHitUser", Erl)
End Sub

''
' Writes the "UserHittedByUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackerChar Char index of the user hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserHittedByUser(ByVal UserIndex As Integer, ByVal Target As e_PartesCuerpo, ByVal attackerChar As Integer, ByVal Damage As Integer)
    On Error GoTo WriteUserHittedByUser_Err
    Call Writer.WriteInt16(ServerPacketID.eUserHittedByUser)
    Call Writer.WriteInt16(attackerChar)
    Call Writer.WriteInt8(Target)
    Call Writer.WriteInt16(Damage)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUserHittedByUser_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserHittedByUser", Erl)
End Sub

''
' Writes the "UserHittedUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackedChar Char index of the user hitted.
' @param    damage The number of HP lost by the oponent hitted.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserHittedUser(ByVal UserIndex As Integer, ByVal Target As e_PartesCuerpo, ByVal attackedChar As Integer, ByVal Damage As Integer)
    On Error GoTo WriteUserHittedUser_Err
    Call Writer.WriteInt16(ServerPacketID.eUserHittedUser)
    Call Writer.WriteInt16(attackedChar)
    Call Writer.WriteInt8(Target)
    Call Writer.WriteInt16(Damage)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUserHittedUser_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserHittedUser", Erl)
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal charindex As Integer, ByVal Color As Long)
    On Error GoTo WriteChatOverHead_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageChatOverHead(chat, charindex, Color, , UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
    Exit Sub
WriteChatOverHead_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChatOverHead", Erl)
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLocaleChatOverHead(ByVal UserIndex As Integer, ByVal ChatId As Integer, ByVal Params As String, ByVal charindex As Integer, ByVal Color As Long)
    On Error GoTo WriteLocaleChatOverHead_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareLocaleChatOverHead(ChatId, Params, charindex, Color, , UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
    Exit Sub
WriteLocaleChatOverHead_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLocaleChatOverHead", Erl)
End Sub

Public Function PrepareLocalizedChatOverHead(ByVal MsgID As Integer, ByVal charindex As Integer, ByVal Color As Long, ParamArray Args() As Variant) As String
    Dim finalText As String
    Dim i         As Long
    finalText = "LOCMSG*" & MsgID & "*"
    For i = LBound(Args) To UBound(Args)
        If i > LBound(Args) Then finalText = finalText & "¬"
        finalText = finalText & CStr(Args(i))
    Next
    PrepareLocalizedChatOverHead = PrepareMessageChatOverHead(finalText, charindex, Color)
End Function

Public Sub WriteTextOverChar(ByVal UserIndex As Integer, ByVal chat As String, ByVal charindex As Integer, ByVal Color As Long)
    On Error GoTo WriteTextOverChar_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextOverChar(chat, charindex, Color))
    Exit Sub
WriteTextOverChar_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTextOverChar", Erl)
End Sub

Public Sub WriteTextOverTile(ByVal UserIndex As Integer, ByVal chat As String, ByVal x As Integer, ByVal y As Integer, ByVal Color As Long)
    On Error GoTo WriteTextOverTile_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextOverTile(chat, x, y, Color))
    Exit Sub
WriteTextOverTile_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTextOverTile", Erl)
End Sub

Public Sub WriteTextCharDrop(ByVal UserIndex As Integer, ByVal chat As String, ByVal charindex As Integer, ByVal Color As Long)
    On Error GoTo WriteTextCharDrop_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextCharDrop(chat, charindex, Color))
    Exit Sub
WriteTextCharDrop_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTextCharDrop", Erl)
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal chat As String, Optional ByVal FontIndex As e_FontTypeNames = FONTTYPE_INFO)
    On Error GoTo WriteConsoleMsg_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageConsoleMsg(chat, FontIndex))
    Exit Sub
WriteConsoleMsg_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteConsoleMsg", Erl)
End Sub

Public Sub WriteLocaleMsg(ByVal UserIndex As Integer, ByVal Id As Integer, ByVal FontIndex As e_FontTypeNames, Optional ByVal strExtra As String = vbNullString)
    On Error GoTo WriteLocaleMsg_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageLocaleMsg(Id, strExtra, FontIndex))
    Exit Sub
WriteLocaleMsg_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLocaleMsg", Erl)
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal chat As String, ByVal Status As Byte)
    On Error GoTo WriteGuildChat_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageGuildChat(chat, Status))
    Exit Sub
WriteGuildChat_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildChat", Erl)
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal MessageId As Integer, Optional ByVal strExtra As String = vbNullString)
    On Error GoTo WriteShowMessageBox_Err
    Call Writer.WriteInt16(ServerPacketID.eShowMessageBox)
    Call Writer.WriteInt16(MessageId)
    Call Writer.WriteString8(strExtra) ' Enviás los valores dinámicos si hay
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowMessageBox_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowMessageBox", Erl)
End Sub

Public Function PrepareShowMessageBox(ByVal MessageId As Integer, Optional ByVal strExtra As String = vbNullString)
    On Error GoTo WriteShowMessageBox_Err
    Call Writer.WriteInt16(ServerPacketID.eShowMessageBox)
    Call Writer.WriteInt16(MessageId)
    Call Writer.WriteString8(strExtra)
    Exit Function
WriteShowMessageBox_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowMessageBox", Erl)
End Function

Public Sub WriteMostrarCuenta(ByVal UserIndex As Integer)
    On Error GoTo WriteMostrarCuenta_Err
    Call Writer.WriteInt16(ServerPacketID.eMostrarCuenta)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteMostrarCuenta_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMostrarCuenta", Erl)
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
    On Error GoTo WriteUserIndexInServer_Err
    Call Writer.WriteInt16(ServerPacketID.eUserIndexInServer)
    Call Writer.WriteInt16(UserIndex)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUserIndexInServer_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserIndexInServer", Erl)
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
    On Error GoTo WriteUserCharIndexInServer_Err
    Call Writer.WriteInt16(ServerPacketID.eUserCharIndexInServer)
    Call Writer.WriteInt16(UserList(UserIndex).Char.charindex)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUserCharIndexInServer_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserCharIndexInServer", Erl)
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
Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, _
                                ByVal body As Integer, _
                                ByVal head As Integer, _
                                ByVal Heading As e_Heading, _
                                ByVal charindex As Integer, _
                                ByVal x As Byte, _
                                ByVal y As Byte, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal Cart As Integer, _
                                ByVal BackPack As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByVal name As String, _
                                ByVal Status As Byte, _
                                ByVal privileges As Byte, _
                                ByVal ParticulaFx As Byte, _
                                ByVal Head_Aura As String, _
                                ByVal Arma_Aura As String, _
                                ByVal Body_Aura As String, _
                                ByVal DM_Aura As String, _
                                ByVal RM_Aura As String, _
                                ByVal Otra_Aura As String, _
                                ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal UserMinMAN As Long, ByVal UserMaxMAN As Long, ByVal Simbolo As Byte, Optional ByVal Idle As Boolean = False, Optional ByVal Navegando As Boolean = False, Optional ByVal tipoUsuario As e_TipoUsuario = 0, Optional ByVal TeamCaptura As Byte = 0, Optional ByVal TieneBandera As Byte = 0, Optional ByVal AnimAtaque1 As Integer = 0)
    On Error GoTo WriteCharacterCreate_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterCreate(body, head, Heading, charindex, x, y, weapon, shield, Cart, BackPack, FX, FXLoops, helmet, name, _
            Status, privileges, ParticulaFx, Head_Aura, Arma_Aura, Body_Aura, DM_Aura, RM_Aura, Otra_Aura, Escudo_Aura, speeding, EsNPC, appear, group_index, clan_index, _
            clan_nivel, UserMinHp, UserMaxHp, UserMinMAN, UserMaxMAN, Simbolo, Idle, Navegando, tipoUsuario, TeamCaptura, TieneBandera, AnimAtaque1))
    Exit Sub
WriteCharacterCreate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterCreate", Erl)
End Sub

Public Sub WriteCharacterUpdateFlag(ByVal UserIndex As Integer, ByVal Flag As Byte, ByVal charindex As Integer)
    On Error GoTo WriteCharacterUpdateFlag_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageUpdateFlag(Flag, charindex))
    Exit Sub
WriteCharacterUpdateFlag_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterUpdateFlag", Erl)
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex As Integer, ByVal Direccion As e_Heading)
    On Error GoTo WriteForceCharMove_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageForceCharMove(Direccion))
    Exit Sub
WriteForceCharMove_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteForceCharMove", Erl)
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
                                ByVal body As Integer, _
                                ByVal head As Integer, _
                                ByVal Heading As e_Heading, _
                                ByVal charindex As Integer, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal Cart As Integer, _
                                ByVal BackPack As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                Optional ByVal Idle As Boolean = False, _
                                Optional ByVal Navegando As Boolean = False)
    On Error GoTo WriteCharacterChange_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterChange(body, head, Heading, charindex, weapon, shield, Cart, BackPack, FX, FXLoops, helmet, Idle, _
            Navegando))
    Exit Sub
WriteCharacterChange_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterChange", Erl)
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal amount As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo WriteObjectCreate_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageObjectCreate(ObjIndex, amount, x, y, ObjData(ObjIndex).ElementalTags))
    Exit Sub
WriteObjectCreate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteObjectCreate", Erl)
End Sub

Public Sub WriteUpdateTrapState(ByVal UserIndex As Integer, State As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo WriteUpdateTrapState_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareTrapUpdate(State, x, y))
    Exit Sub
WriteUpdateTrapState_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateTrapState", Erl)
End Sub

Public Sub WriteParticleFloorCreate(ByVal UserIndex As Integer, ByVal Particula As Integer, ByVal ParticulaTime As Integer, ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo WriteParticleFloorCreate_Err
    If Particula = 0 Then
        Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageParticleFXToFloor(x, y, Particula, ParticulaTime))
    End If
    Exit Sub
WriteParticleFloorCreate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteParticleFloorCreate", Erl)
End Sub

Public Sub WriteLightFloorCreate(ByVal UserIndex As Integer, ByVal LuzColor As Long, ByVal Rango As Byte, ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo WriteLightFloorCreate_Err
    MapData(Map, x, y).Luz.Color = LuzColor
    MapData(Map, x, y).Luz.Rango = Rango
    If Rango = 0 Then
        Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageLightFXToFloor(x, y, LuzColor, Rango))
    End If
    Exit Sub
WriteLightFloorCreate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLightFloorCreate", Erl)
End Sub

Public Sub WriteFxPiso(ByVal UserIndex As Integer, ByVal GrhIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo WriteFxPiso_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageFxPiso(GrhIndex, x, y))
    Exit Sub
WriteFxPiso_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFxPiso", Erl)
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo WriteObjectDelete_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageObjectDelete(x, y))
    Exit Sub
WriteObjectDelete_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteObjectDelete", Erl)
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub Write_BlockPosition(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte, ByVal Blocked As Byte)
    On Error GoTo Write_BlockPosition_Err
    Call Writer.WriteInt16(ServerPacketID.eBlockPosition)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt8(Blocked)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
Write_BlockPosition_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.Write_BlockPosition", Erl)
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)
    On Error GoTo WritePlayMidi_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePlayMidi(midi, loops))
    Exit Sub
WritePlayMidi_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePlayMidi", Erl)
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
                         ByVal wave As Integer, _
                         ByVal x As Byte, _
                         ByVal y As Byte, _
                         Optional ByVal CancelLastWave As Byte = 0, _
                         Optional ByVal Localize As Byte = 0)
    On Error GoTo WritePlayWave_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePlayWave(wave, x, y, CancelLastWave, Localize))
    Exit Sub
WritePlayWave_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePlayWave", Erl)
End Sub

Public Sub WritePlayWaveStep(ByVal UserIndex As Integer, _
                             ByVal charindex As Integer, _
                             ByVal grh As Long, _
                             ByVal grh2 As Long, _
                             ByVal Distance As Byte, _
                             ByVal balance As Integer, _
                             ByVal step As Boolean)
    On Error GoTo WritePlayWaveStep_Err
    Call Writer.WriteInt16(ServerPacketID.ePlayWaveStep)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt32(grh)
    Call Writer.WriteInt32(grh2)
    Call Writer.WriteInt8(Distance)
    Call Writer.WriteInt16(balance)
    Call Writer.WriteBool(step)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePlayWaveStep_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePlayWaveStep", Erl)
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)
    On Error GoTo WriteGuildList_Err
    Dim Tmp As String
    Dim i   As Long
    Call Writer.WriteInt16(ServerPacketID.eguildList)
    ' Prepare guild name's list
    For i = LBound(guildList()) To UBound(guildList())
        Tmp = Tmp & guildList(i) & SEPARATOR
    Next i
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteGuildList_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildList", Erl)
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAreaChanged(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo WriteAreaChanged_Err
    Call Writer.WriteInt16(ServerPacketID.eAreaChanged)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteAreaChanged_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAreaChanged", Erl)
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
''
' Writes the "RainToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRainToggle(ByVal UserIndex As Integer)
    On Error GoTo WriteRainToggle_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageRainToggle())
    Exit Sub
WriteRainToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRainToggle", Erl)
End Sub

Public Sub WriteNubesToggle(ByVal UserIndex As Integer)
    On Error GoTo WriteNubesToggle_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageNieblandoToggle(IntensidadDeNubes))
    Exit Sub
WriteNubesToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNubesToggle", Erl)
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal charindex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
    'Writes the "CreateFX" message to the given user's outgoing data buffer
    On Error GoTo WriteCreateFX_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCreateFX(charindex, FX, FXLoops, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
    Exit Sub
WriteCreateFX_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCreateFX", Erl)
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
    'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
    On Error GoTo WriteUpdateUserStats_Err
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateMAN(UserIndex))
    Call Writer.WriteInt16(ServerPacketID.eUpdateUserStats)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxHp)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.shield)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxMAN)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxSta)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
    Call Writer.WriteInt32(SvrConfig.GetValue("OroPorNivelBilletera"))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.ELV)
    Call Writer.WriteInt32(ExpLevelUp(UserList(UserIndex).Stats.ELV))
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    Call Writer.WriteInt8(UserList(UserIndex).clase)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateUserStats_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateUserStats", Erl)
End Sub

Public Sub WriteUpdateUserKey(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Llave As Integer)
    On Error GoTo WriteUpdateUserKey_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateUserKey)
    Call Writer.WriteInt16(Slot)
    Call Writer.WriteInt16(Llave)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateUserKey_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateUserKey", Erl)
End Sub

' Actualiza el indicador de daño mágico
Public Sub WriteUpdateDM(ByVal UserIndex As Integer)
    On Error GoTo WriteUpdateDM_Err
    Dim Valor As Integer
    With UserList(UserIndex).invent
        ' % daño mágico del arma
        If .EquippedWeaponObjIndex > 0 Then
            Valor = Valor + ObjData(.EquippedWeaponObjIndex).MagicDamageBonus
        End If
        ' % daño mágico del anillo
        If .EquippedRingAccesoryObjIndex > 0 Then
            Valor = Valor + ObjData(.EquippedRingAccesoryObjIndex).MagicDamageBonus
        End If
        Call Writer.WriteInt16(ServerPacketID.eUpdateDM)
        Call Writer.WriteInt16(Valor)
    End With
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateDM_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateDM", Erl)
End Sub

' Actualiza el indicador de resistencia mágica
Public Sub WriteUpdateRM(ByVal UserIndex As Integer)
    On Error GoTo WriteUpdateRM_Err
    Dim Valor As Integer
    With UserList(UserIndex).invent
        ' Resistencia mágica de la armadura
        If .EquippedArmorObjIndex > 0 Then
            Valor = Valor + ObjData(.EquippedArmorObjIndex).ResistenciaMagica
        End If
        ' Resistencia mágica del anillo
        If .EquippedRingAccesoryObjIndex > 0 Then
            Valor = Valor + ObjData(.EquippedRingAccesoryObjIndex).ResistenciaMagica
        End If
        ' Resistencia mágica del escudo
        If .EquippedShieldObjIndex > 0 Then
            Valor = Valor + ObjData(.EquippedShieldObjIndex).ResistenciaMagica
        End If
        ' Resistencia mágica del casco
        If .EquippedHelmetObjIndex > 0 Then
            Valor = Valor + ObjData(.EquippedHelmetObjIndex).ResistenciaMagica
        End If
        Valor = Valor + 100 * ModClase(UserList(UserIndex).clase).ResistenciaMagica
        Call Writer.WriteInt16(ServerPacketID.eUpdateRM)
        Call Writer.WriteInt16(Valor)
    End With
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateRM_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateRM", Erl)
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As e_Skill, Optional ByVal CasteaArea As Boolean = False, Optional ByVal Radio As Byte = 0)
    On Error GoTo WriteWorkRequestTarget_Err
    Call Writer.WriteInt16(ServerPacketID.eWorkRequestTarget)
    Call Writer.WriteInt8(Skill)
    Call Writer.WriteBool(CasteaArea)
    Call Writer.WriteInt8(Radio)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteWorkRequestTarget_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteWorkRequestTarget", Erl)
End Sub

Public Sub WriteInventoryUnlockSlots(ByVal UserIndex As Integer)
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
End Sub

Public Sub WriteIntervals(ByVal UserIndex As Integer)
    On Error GoTo WriteIntervals_Err
    With UserList(UserIndex)
        Call Writer.WriteInt16(ServerPacketID.eIntervals)
        Call Writer.WriteInt32(.Intervals.Arco)
        Call Writer.WriteInt32(.Intervals.Caminar)
        Call Writer.WriteInt32(.Intervals.Golpe)
        Call Writer.WriteInt32(.Intervals.GolpeMagia)
        Call Writer.WriteInt32(.Intervals.Magia)
        Call Writer.WriteInt32(.Intervals.MagiaGolpe)
        Call Writer.WriteInt32(.Intervals.GolpeUsar)
        Call Writer.WriteInt32(.Intervals.TrabajarExtraer)
        Call Writer.WriteInt32(.Intervals.TrabajarConstruir)
        Call Writer.WriteInt32(.Intervals.UsarU)
        Call Writer.WriteInt32(.Intervals.UsarClic)
        Call Writer.WriteInt32(IntervaloTirar)
    End With
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteIntervals_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteIntervals", Erl)
End Sub

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo WriteChangeInventorySlot_Err
    Dim ObjIndex             As Integer
    Dim NaturalElementalTags As Long
    Dim PodraUsarlo          As Byte
    Call Writer.WriteInt16(ServerPacketID.eChangeInventorySlot)
    Call Writer.WriteInt8(Slot)
    ObjIndex = UserList(UserIndex).invent.Object(Slot).ObjIndex
    If ObjIndex > 0 Then
        PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)
        NaturalElementalTags = ObjData(UserList(UserIndex).invent.Object(Slot).ObjIndex).ElementalTags
    End If
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt16(UserList(UserIndex).invent.Object(Slot).amount)
    Call Writer.WriteBool(UserList(UserIndex).invent.Object(Slot).Equipped)
    Call Writer.WriteReal32(SalePrice(ObjIndex))
    Call Writer.WriteInt8(PodraUsarlo)
    Call Writer.WriteInt32(UserList(UserIndex).invent.Object(Slot).ElementalTags Or NaturalElementalTags)
    If ObjIndex > 0 Then
        Call Writer.WriteBool(IsSet(ObjData(ObjIndex).ObjFlags, e_ObjFlags.e_Bindable))
    Else
        Call Writer.WriteBool(False)
    End If
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteChangeInventorySlot_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeInventorySlot", Erl)
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo WriteChangeBankSlot_Err
    Dim ObjIndex             As Integer
    Dim Valor                As Long
    Dim NaturalElementalTags As Long
    Dim PodraUsarlo          As Byte
    Call Writer.WriteInt16(ServerPacketID.eChangeBankSlot)
    Call Writer.WriteInt8(Slot)
    ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
    If ObjIndex > 0 Then
        Valor = ObjData(ObjIndex).Valor
        PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)
        NaturalElementalTags = ObjData(ObjIndex).ElementalTags
    Else
    End If
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt32(UserList(UserIndex).BancoInvent.Object(Slot).ElementalTags Or NaturalElementalTags)
    Call Writer.WriteInt16(UserList(UserIndex).BancoInvent.Object(Slot).amount)
    Call Writer.WriteInt32(Valor)
    Call Writer.WriteInt8(PodraUsarlo)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteChangeBankSlot_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeBankSlot", Erl)
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
    On Error GoTo WriteChangeSpellSlot_Err
    Call Writer.WriteInt16(ServerPacketID.eChangeSpellSlot)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.UserHechizos(Slot))
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call Writer.WriteInt16(UserList(UserIndex).Stats.UserHechizos(Slot))
        Call Writer.WriteBool(IsSet(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).SpellRequirementMask, e_SpellRequirementMask.eIsBindable))
    Else
        Call Writer.WriteInt16(-1)
        Call Writer.WriteBool(False)
    End If
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteChangeSpellSlot_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeSpellSlot", Erl)
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAttributes(ByVal UserIndex As Integer)
    On Error GoTo WriteAttributes_Err
    Call Writer.WriteInt16(ServerPacketID.eAtributes)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Fuerza))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Agilidad))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Inteligencia))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Constitucion))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Carisma))
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteAttributes_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAttributes", Erl)
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)
    On Error GoTo WriteBlacksmithWeapons_Err
    Dim i              As Long
    Dim validIndexes() As Integer
    Dim count          As Integer
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    Call Writer.WriteInt16(ServerPacketID.eBlacksmithWeapons)
    For i = 1 To UBound(ArmasHerrero())
        ' Can the user create this object? If so add it to the list....
        If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(e_Skill.Herreria) Then
            count = count + 1
            validIndexes(count) = i
        End If
    Next i
    ' Write the number of objects in the list
    Call Writer.WriteInt16(count)
    ' Write the needed data of each object
    For i = 1 To count
        Call Writer.WriteInt16(ArmasHerrero(validIndexes(i)))
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteBlacksmithWeapons_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlacksmithWeapons", Erl)
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)
    On Error GoTo WriteBlacksmithArmors_Err
    Dim i              As Long
    Dim validIndexes() As Integer
    Dim count          As Integer
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    Call Writer.WriteInt16(ServerPacketID.eBlacksmithArmors)
    For i = 1 To UBound(ArmadurasHerrero())
        ' Can the user create this object? If so add it to the list....
        If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(e_Skill.Herreria) / ModHerreria(UserList(UserIndex).clase), 0) Then
            count = count + 1
            validIndexes(count) = i
        End If
    Next i
    ' Write the number of objects in the list
    Call Writer.WriteInt16(count)
    ' Write the needed data of each object
    For i = 1 To count
        Call Writer.WriteInt16(ArmadurasHerrero(validIndexes(i)))
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteBlacksmithArmors_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlacksmithArmors", Erl)
End Sub

Public Sub WriteBlacksmithElementalRunes(ByVal UserIndex As Integer)
    On Error GoTo WriteBlacksmithElementalRunes_Err
    Dim i              As Long
    Dim validIndexes() As Integer
    Dim count          As Integer
    ReDim validIndexes(1 To UBound(BlackSmithElementalRunes()))
    Call Writer.WriteInt16(ServerPacketID.eBlacksmithExtraObjects)
    For i = 1 To UBound(BlackSmithElementalRunes())
        ' Can the user create this object? If so add it to the list....
        If ObjData(BlackSmithElementalRunes(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(e_Skill.Herreria) Then
            count = count + 1
            validIndexes(count) = i
        End If
    Next i
    ' Write the number of objects in the list
    Call Writer.WriteInt16(count)
    ' Write the needed data of each object
    For i = 1 To count
        Call Writer.WriteInt16(BlackSmithElementalRunes(validIndexes(i)))
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteBlacksmithElementalRunes_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlacksmithElementalRunes", Erl)
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)
    On Error GoTo WriteCarpenterObjects_Err
    Dim i              As Long
    Dim validIndexes() As Integer
    Dim count          As Byte
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    Call Writer.WriteInt16(ServerPacketID.eCarpenterObjects)
    For i = 1 To UBound(ObjCarpintero())
        ' Can the user create this object? If so add it to the list....
        If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(e_Skill.Carpinteria) Then
            If i = 1 Then Debug.Print UserList(UserIndex).Stats.UserSkills(e_Skill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase)
            count = count + 1
            validIndexes(count) = i
        End If
    Next i
    ' Write the number of objects in the list
    Call Writer.WriteInt8(count)
    ' Write the needed data of each object
    For i = 1 To count
        Call Writer.WriteInt16(ObjCarpintero(validIndexes(i)))
        'Call Writer.WriteInt16(obj.Madera)
        'Call Writer.WriteInt32(obj.GrhIndex)
        ' Ladder 07/07/2014   Ahora se envia el grafico de los objetos
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCarpenterObjects_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCarpenterObjects", Erl)
End Sub

Public Sub WriteAlquimistaObjects(ByVal UserIndex As Integer)
    On Error GoTo WriteAlquimistaObjects_Err
    Dim i              As Long
    Dim validIndexes() As Integer
    Dim count          As Integer
    ReDim validIndexes(1 To UBound(ObjAlquimista()))
    Call Writer.WriteInt16(ServerPacketID.eAlquimistaObj)
    For i = 1 To UBound(ObjAlquimista())
        ' Can the user create this object? If so add it to the list....
        If ObjData(ObjAlquimista(i)).SkPociones <= UserList(UserIndex).Stats.UserSkills(e_Skill.Alquimia) \ ModAlquimia(UserList(UserIndex).clase) Then
            count = count + 1
            validIndexes(count) = i
        End If
    Next i
    ' Write the number of objects in the list
    Call Writer.WriteInt16(count)
    ' Write the needed data of each object
    For i = 1 To count
        Call Writer.WriteInt16(ObjAlquimista(validIndexes(i)))
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteAlquimistaObjects_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAlquimistaObjects", Erl)
End Sub

Public Sub WriteSastreObjects(ByVal UserIndex As Integer)
    On Error GoTo WriteSastreObjects_Err
    Dim i              As Long
    Dim validIndexes() As Integer
    Dim count          As Integer
    ReDim validIndexes(1 To UBound(ObjSastre()))
    Call Writer.WriteInt16(ServerPacketID.eSastreObj)
    For i = 1 To UBound(ObjSastre())
        ' Can the user create this object? If so add it to the list....
        If ObjData(ObjSastre(i)).SkSastreria <= UserList(UserIndex).Stats.UserSkills(e_Skill.Sastreria) Then
            count = count + 1
            validIndexes(count) = i
        End If
    Next i
    ' Write the number of objects in the list
    Call Writer.WriteInt16(count)
    ' Write the needed data of each object
    For i = 1 To count
        Call Writer.WriteInt16(ObjSastre(validIndexes(i)))
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteSastreObjects_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSastreObjects", Erl)
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRestOK(ByVal UserIndex As Integer)
    On Error GoTo WriteRestOK_Err
    Call Writer.WriteInt16(ServerPacketID.eRestOK)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteRestOK_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRestOK", Erl)
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)
    'Writes the "ErrorMsg" message to the given user's outgoing data buffer
    On Error GoTo WriteErrorMsg_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageErrorMsg(Message))
    Exit Sub
WriteErrorMsg_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteErrorMsg", Erl)
End Sub

''
' Writes the "Blind" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlind(ByVal UserIndex As Integer)
    On Error GoTo WriteBlind_Err
    Call Writer.WriteInt16(ServerPacketID.eBlind)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteBlind_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlind", Erl)
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDumb(ByVal UserIndex As Integer)
    On Error GoTo WriteDumb_Err
    Call Writer.WriteInt16(ServerPacketID.eDumb)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteDumb_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDumb", Erl)
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
'Optimizacion de protocolo por Ladder
Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo WriteShowSignal_Err
    Call Writer.WriteInt16(ServerPacketID.eShowSignal)
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt16(ObjData(ObjIndex).GrhSecundario)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowSignal_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowSignal", Erl)
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef obj As t_Obj, ByVal price As Single)
    On Error GoTo WriteChangeNPCInventorySlot_Err
    Dim PodraUsarlo As Byte
    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        PodraUsarlo = PuedeUsarObjeto(UserIndex, obj.ObjIndex)
    End If
    Call Writer.WriteInt16(ServerPacketID.eChangeNPCInventorySlot)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(obj.ObjIndex)
    Call Writer.WriteInt16(obj.amount)
    Call Writer.WriteReal32(price)
    Call Writer.WriteInt32(obj.ElementalTags)
    Call Writer.WriteInt8(PodraUsarlo)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteChangeNPCInventorySlot_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeNPCInventorySlot", Erl)
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
    On Error GoTo WriteUpdateHungerAndThirst_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateHungerAndThirst)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxAGU)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MinAGU)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxHam)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MinHam)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateHungerAndThirst_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateHungerAndThirst", Erl)
End Sub

Public Sub WriteLight(ByVal UserIndex As Integer, ByVal Map As Integer)
    On Error GoTo WriteLight_Err
    Call Writer.WriteInt16(ServerPacketID.eLight)
    Call Writer.WriteString8(MapInfo(Map).base_light)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteLight_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLight", Erl)
End Sub

Public Sub WriteFlashScreen(ByVal UserIndex As Integer, ByVal Color As Long, ByVal Time As Long, Optional ByVal Ignorar As Boolean = False)
    On Error GoTo WriteFlashScreen_Err
    Call Writer.WriteInt16(ServerPacketID.eFlashScreen)
    Call Writer.WriteInt32(Color)
    Call Writer.WriteInt32(Time)
    Call Writer.WriteBool(Ignorar)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteFlashScreen_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFlashScreen", Erl)
End Sub

Public Sub WriteFYA(ByVal UserIndex As Integer)
    On Error GoTo WriteFYA_Err
    Call Writer.WriteInt16(ServerPacketID.eFYA)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(1))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(2))
    Call Writer.WriteInt16(UserList(UserIndex).flags.DuracionEfecto)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteFYA_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFYA", Erl)
End Sub

Public Sub WriteCerrarleCliente(ByVal UserIndex As Integer)
    On Error GoTo WriteCerrarleCliente_Err
    Call Writer.WriteInt16(ServerPacketID.eCerrarleCliente)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCerrarleCliente_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCerrarleCliente", Erl)
End Sub

Public Sub WriteContadores(ByVal UserIndex As Integer)
    On Error GoTo WriteContadores_Err
    Call Writer.WriteInt16(ServerPacketID.eContadores)
    Call Writer.WriteInt16(UserList(UserIndex).Counters.Invisibilidad)
    Call Writer.WriteInt16(UserList(UserIndex).flags.DuracionEfecto)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteContadores_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteContadores", Erl)
End Sub

Public Sub WriteShowPapiro(ByVal UserIndex As Integer)
    On Error GoTo WriteShowPapiro_Err
    Call Writer.WriteInt16(ServerPacketID.eShowPapiro)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowPapiro_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowPapiro", Erl)
End Sub

Public Sub WriteUpdateCdType(ByVal UserIndex As Integer, ByVal cdType As Byte)
    On Error GoTo WriteUpdateCdType_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateCooldownType)
    Call Writer.WriteInt8(cdType)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateCdType_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateCdType", Erl)
End Sub

Public Sub WritePrivilegios(ByVal UserIndex As Integer)
    On Error GoTo WritePrivilegios_Err
    Call Writer.WriteInt16(ServerPacketID.ePrivilegios)
    If UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero) Then
        Call Writer.WriteBool(True)
    Else
        Call Writer.WriteBool(False)
    End If
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePrivilegios_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePrivilegios", Erl)
End Sub

Public Sub WriteBindKeys(ByVal UserIndex As Integer)
    On Error GoTo WriteBindKeys_Err
    Call Writer.WriteInt16(ServerPacketID.eBindKeys)
    Call Writer.WriteInt8(UserList(UserIndex).ChatCombate)
    Call Writer.WriteInt8(UserList(UserIndex).ChatGlobal)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteBindKeys_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBindKeys", Erl)
End Sub


Public Sub WriteGetInventarioHechizos(ByVal UserIndex As Integer, ByVal value As Byte, ByVal hechiSel As Byte, ByVal scrollSel As Byte)
    On Error GoTo GetInventarioHechizos_Err
    Call Writer.WriteInt16(ServerPacketID.eGetInventarioHechizos)
    Call Writer.WriteInt8(value)
    Call Writer.WriteInt8(hechiSel)
    Call Writer.WriteInt8(scrollSel)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
GetInventarioHechizos_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.GetInventarioHechizos", Erl)
End Sub

Public Sub WriteNofiticarClienteCasteo(ByVal UserIndex As Integer, ByVal value As Byte)
    On Error GoTo NofiticarClienteCasteo_Err
    Call Writer.WriteInt16(ServerPacketID.eNotificarClienteCasteo)
    Call Writer.WriteInt8(value)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
NofiticarClienteCasteo_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.NofiticarClienteCasteo", Erl)
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMiniStats(ByVal UserIndex As Integer)
    On Error GoTo WriteMiniStats_Err
    Call Writer.WriteInt16(ServerPacketID.eMiniStats)
    Call Writer.WriteInt32(UserList(UserIndex).Faccion.ciudadanosMatados)
    Call Writer.WriteInt32(UserList(UserIndex).Faccion.CriminalesMatados)
    Call Writer.WriteInt8(UserList(UserIndex).Faccion.Status)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.NPCsMuertos)
    Call Writer.WriteInt8(UserList(UserIndex).clase)
    Call Writer.WriteInt32(UserList(UserIndex).Counters.Pena)
    Call Writer.WriteInt32(UserList(UserIndex).flags.VecesQueMoriste)
    Call Writer.WriteInt8(UserList(UserIndex).genero)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.PuntosPesca)
    Call Writer.WriteInt8(UserList(UserIndex).raza)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteMiniStats_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMiniStats", Erl)
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data .incomingData.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
    On Error GoTo WriteLevelUp_Err
    Call Writer.WriteInt16(ServerPacketID.eLevelUp)
    Call Writer.WriteInt16(skillPoints)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteLevelUp_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLevelUp", Erl)
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data .incomingData.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, ByVal title As String, ByVal Message As String)
    On Error GoTo WriteAddForumMsg_Err
    Call Writer.WriteInt16(ServerPacketID.eAddForumMsg)
    Call Writer.WriteString8(title)
    Call Writer.WriteString8(Message)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteAddForumMsg_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAddForumMsg", Erl)
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
    On Error GoTo WriteShowForumForm_Err
    Call Writer.WriteInt16(ServerPacketID.eShowForumForm)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowForumForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowForumForm", Erl)
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    TargetIndex The user turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal TargetIndex As Integer, ByVal invisible As Boolean)
    On Error GoTo WriteSetInvisible_Err
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageSetInvisible(UserList(TargetIndex).Char.charindex, invisible, UserList(TargetIndex).pos.x, UserList( _
            TargetIndex).pos.y))
    Exit Sub
WriteSetInvisible_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSetInvisible", Erl)
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
    On Error GoTo WriteMeditateToggle_Err
    Call Writer.WriteInt16(ServerPacketID.eMeditateToggle)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteMeditateToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMeditateToggle", Erl)
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
    On Error GoTo WriteBlindNoMore_Err
    Call Writer.WriteInt16(ServerPacketID.eBlindNoMore)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteBlindNoMore_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlindNoMore", Erl)
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
    On Error GoTo WriteDumbNoMore_Err
    Call Writer.WriteInt16(ServerPacketID.eDumbNoMore)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteDumbNoMore_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDumbNoMore", Erl)
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSendSkills(ByVal UserIndex As Integer)
    On Error GoTo WriteSendSkills_Err
    Dim i As Long
    Call Writer.WriteInt16(ServerPacketID.eSendSkills)
    For i = 1 To NUMSKILLS
        Call Writer.WriteInt8(UserList(UserIndex).Stats.UserSkills(i))
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteSendSkills_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSendSkills", Erl)
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    On Error GoTo WriteTrainerCreatureList_Err
    Dim i   As Long
    Dim str As String
    Call Writer.WriteInt16(ServerPacketID.eTrainerCreatureList)
    For i = 1 To NpcList(NpcIndex).NroCriaturas
        str = str & NpcList(NpcIndex).Criaturas(i).NpcName & SEPARATOR
    Next i
    If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
    Call Writer.WriteString8(str)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteTrainerCreatureList_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTrainerCreatureList", Erl)
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
                          ByVal guildNews As String, _
                          ByRef guildList() As String, _
                          ByRef MemberList() As Long, _
                          ByVal ClanNivel As Byte, _
                          ByVal ExpAcu As Integer, _
                          ByVal ExpNe As Integer)
    On Error GoTo WriteGuildNews_Err
    Dim i   As Long
    Dim Tmp As String
    Call Writer.WriteInt16(ServerPacketID.eguildNews)
    Call Writer.WriteString8(guildNews)
    ' Prepare guild name's list
    For i = LBound(guildList()) To UBound(guildList())
        Tmp = Tmp & guildList(i) & SEPARATOR
    Next i
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    ' Prepare guild member's list
    Tmp = vbNullString
    For i = LBound(MemberList()) To UBound(MemberList())
        Tmp = Tmp & GetUserName(MemberList(i)) & SEPARATOR
    Next i
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call Writer.WriteInt8(ClanNivel)
    Call Writer.WriteInt16(ExpAcu)
    Call Writer.WriteInt16(ExpNe)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteGuildNews_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildNews", Erl)
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)
    On Error GoTo WriteOfferDetails_Err
    Call Writer.WriteInt16(ServerPacketID.eOfferDetails)
    Call Writer.WriteString8(details)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteOfferDetails_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteOfferDetails", Erl)
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
    On Error GoTo WriteAlianceProposalsList_Err
    Dim i   As Long
    Dim Tmp As String
    Call Writer.WriteInt16(ServerPacketID.eAlianceProposalsList)
    ' Prepare guild's list
    For i = LBound(guilds()) To UBound(guilds())
        Tmp = Tmp & guilds(i) & SEPARATOR
    Next i
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteAlianceProposalsList_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAlianceProposalsList", Erl)
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
    On Error GoTo WritePeaceProposalsList_Err
    Dim i   As Long
    Dim Tmp As String
    Call Writer.WriteInt16(ServerPacketID.ePeaceProposalsList)
    ' Prepare guilds' list
    For i = LBound(guilds()) To UBound(guilds())
        Tmp = Tmp & guilds(i) & SEPARATOR
    Next i
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePeaceProposalsList_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePeaceProposalsList", Erl)
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
Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal CharName As String, ByVal race As e_Raza, ByVal Class As e_Class, ByVal gender As e_Genero, ByVal level As Byte, _
        ByVal gold As Long, ByVal bank As Long, ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, _
        ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)
    On Error GoTo WriteCharacterInfo_Err
    Call Writer.WriteInt16(ServerPacketID.eCharacterInfo)
    Call Writer.WriteInt8(gender)
    Call Writer.WriteString8(CharName)
    Call Writer.WriteInt8(race)
    Call Writer.WriteInt8(Class)
    Call Writer.WriteInt8(level)
    Call Writer.WriteInt32(gold)
    Call Writer.WriteInt32(bank)
    Call Writer.WriteString8(previousPetitions)
    Call Writer.WriteString8(currentGuild)
    Call Writer.WriteString8(previousGuilds)
    Call Writer.WriteBool(RoyalArmy)
    Call Writer.WriteBool(CaosLegion)
    Call Writer.WriteInt32(citicensKilled)
    Call Writer.WriteInt32(criminalsKilled)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCharacterInfo_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterInfo", Erl)
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
    Call Writer.WriteInt16(ServerPacketID.eGuildLeaderInfo)
    ' Prepare guild name's list
    For i = LBound(guildList()) To UBound(guildList())
        Tmp = Tmp & guildList(i) & SEPARATOR
    Next i
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    ' Prepare guild member's list
    Tmp = vbNullString
    For i = LBound(MemberList()) To UBound(MemberList())
        Tmp = Tmp & GetUserName(MemberList(i)) & SEPARATOR
    Next i
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    ' Store guild news
    Call Writer.WriteString8(guildNews)
    ' Prepare the join request's list
    Tmp = vbNullString
    For i = LBound(joinRequests()) To UBound(joinRequests())
        Tmp = Tmp & joinRequests(i) & SEPARATOR
    Next i
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call Writer.WriteInt8(NivelDeClan)
    Call Writer.WriteInt16(ExpActual)
    Call Writer.WriteInt16(ExpNecesaria)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteGuildLeaderInfo_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildLeaderInfo", Erl)
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
                             ByVal GuildName As String, _
                             ByVal founder As String, _
                             ByVal foundationDate As String, _
                             ByVal leader As String, _
                             ByVal memberCount As Integer, _
                             ByVal alignment As String, _
                             ByVal guildDesc As String, _
                             ByVal NivelDeClan As Byte)
    On Error GoTo WriteGuildDetails_Err
    Call Writer.WriteInt16(ServerPacketID.eGuildDetails)
    Call Writer.WriteString8(GuildName)
    Call Writer.WriteString8(founder)
    Call Writer.WriteString8(foundationDate)
    Call Writer.WriteString8(leader)
    Call Writer.WriteInt16(memberCount)
    Call Writer.WriteString8(alignment)
    Call Writer.WriteString8(guildDesc)
    Call Writer.WriteInt8(NivelDeClan)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteGuildDetails_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildDetails", Erl)
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)
    On Error GoTo WriteShowGuildFundationForm_Err
    Call Writer.WriteInt16(ServerPacketID.eShowGuildFundationForm)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowGuildFundationForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowGuildFundationForm", Erl)
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteParalizeOK(ByVal UserIndex As Integer)
    On Error GoTo WriteParalizeOK_Err
    Call Writer.WriteInt16(ServerPacketID.eParalizeOK)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteParalizeOK_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteParalizeOK", Erl)
End Sub

Public Sub WriteStunStart(ByVal UserIndex As Integer, ByVal Duration As Integer)
    On Error GoTo WriteStunStart_Err
    Call Writer.WriteInt16(ServerPacketID.eStunStart)
    Call Writer.WriteInt16(Duration)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteStunStart_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteStunStart", Erl)
End Sub

Public Sub WriteInmovilizaOK(ByVal UserIndex As Integer)
    On Error GoTo WriteInmovilizaOK_Err
    Call Writer.WriteInt16(ServerPacketID.eInmovilizadoOK)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteInmovilizaOK_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteInmovilizaOK", Erl)
End Sub

Public Sub WriteStopped(ByVal UserIndex As Integer, ByVal Stopped As Boolean)
    On Error GoTo WriteStopped_Err
    Call Writer.WriteInt16(ServerPacketID.eStopped)
    Call Writer.WriteBool(Stopped)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteStopped_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteStopped", Erl)
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
    On Error GoTo WriteShowUserRequest_Err
    Call Writer.WriteInt16(ServerPacketID.eShowUserRequest)
    Call Writer.WriteString8(details)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowUserRequest_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowUserRequest", Erl)
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    Amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByRef itemsAenviar() As t_Obj, ByVal gold As Long, ByVal miOferta As Boolean)
    On Error GoTo WriteChangeUserTradeSlot_Err
    Call Writer.WriteInt16(ServerPacketID.eChangeUserTradeSlot)
    Call Writer.WriteBool(miOferta)
    Call Writer.WriteInt32(gold)
    Dim i As Long
    For i = 1 To UBound(itemsAenviar)
        Call Writer.WriteInt16(itemsAenviar(i).ObjIndex)
        If itemsAenviar(i).ObjIndex = 0 Then
            Call Writer.WriteString8("")
        Else
            Call Writer.WriteString8(ObjData(itemsAenviar(i).ObjIndex).name)
        End If
        If itemsAenviar(i).ObjIndex = 0 Then
            Call Writer.WriteInt32(0)
        Else
            Call Writer.WriteInt32(ObjData(itemsAenviar(i).ObjIndex).GrhIndex)
        End If
        Call Writer.WriteInt32(itemsAenviar(i).amount)
        Call Writer.WriteInt32(itemsAenviar(i).ElementalTags)
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteChangeUserTradeSlot_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeUserTradeSlot", Erl)
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByVal ListaCompleta As Boolean)
    On Error GoTo WriteSpawnList_Err
    Call Writer.WriteInt16(ServerPacketID.eSpawnListt)
    Call Writer.WriteBool(ListaCompleta)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteSpawnList_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSpawnList", Erl)
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)
    On Error GoTo WriteShowSOSForm_Err
    Dim i   As Long
    Dim Tmp As String
    Call Writer.WriteInt16(ServerPacketID.eShowSOSForm)
    For i = 1 To Ayuda.Longitud
        Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
    Next i
    If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowSOSForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowSOSForm", Erl)
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)
    On Error GoTo WriteShowMOTDEditionForm_Err
    Call Writer.WriteInt16(ServerPacketID.eShowMOTDEditionForm)
    Call Writer.WriteString8(currentMOTD)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowMOTDEditionForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowMOTDEditionForm", Erl)
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
    On Error GoTo WriteShowGMPanelForm_Err
    Call Writer.WriteInt16(ServerPacketID.eShowGMPanelForm)
    Call Writer.WriteInt16(UserList(UserIndex).Char.head)
    Call Writer.WriteInt16(UserList(UserIndex).Char.body)
    Call Writer.WriteInt16(UserList(UserIndex).Char.CascoAnim)
    Call Writer.WriteInt16(UserList(UserIndex).Char.WeaponAnim)
    Call Writer.WriteInt16(UserList(UserIndex).Char.ShieldAnim)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowGMPanelForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowGMPanelForm", Erl)
End Sub

Public Sub WriteShowFundarClanForm(ByVal UserIndex As Integer)
    On Error GoTo WriteShowFundarClanForm_Err
    Call Writer.WriteInt16(ServerPacketID.eShowFundarClanForm)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowFundarClanForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowFundarClanForm", Erl)
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
    On Error GoTo WriteUserNameList_Err
    Dim i   As Long
    Dim Tmp As String
    Call Writer.WriteInt16(ServerPacketID.eUserNameList)
    ' Prepare user's names list
    For i = 1 To cant
        Tmp = Tmp & userNamesList(i) & SEPARATOR
    Next i
    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUserNameList_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserNameList", Erl)
End Sub

Public Sub WriteGoliathInit(ByVal UserIndex As Integer)
    On Error GoTo WriteGoliathInit_Err
    Call Writer.WriteInt16(ServerPacketID.eGoliath)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
    Call Writer.WriteInt8(UserList(UserIndex).BancoInvent.NroItems)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteGoliathInit_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGoliathInit", Erl)
End Sub

Public Sub WritePelearConPezEspecial(ByVal UserIndex As Integer)
    On Error GoTo WritePelearConPezEspecial_Err
    Call Writer.WriteInt16(ServerPacketID.ePelearConPezEspecial)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePelearConPezEspecial_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePelearConPezEspecial", Erl)
End Sub

Public Sub WriteUpdateBankGld(ByVal UserIndex As Integer)
    On Error GoTo WriteUpdateBankGld_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateBankGld)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateBankGld_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateBankGld", Erl)
End Sub

Public Sub WriteShowFrmLogear(ByVal UserIndex As Integer)
    On Error GoTo WriteShowFrmLogear_Err
    Call Writer.WriteInt16(ServerPacketID.eShowFrmLogear)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowFrmLogear_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowFrmLogear", Erl)
End Sub

Public Sub WriteShowFrmMapa(ByVal UserIndex As Integer)
    On Error GoTo WriteShowFrmMapa_Err
    Call Writer.WriteInt16(ServerPacketID.eShowFrmMapa)
    Call Writer.WriteInt16(SvrConfig.GetValue("ExpMult"))
    Call Writer.WriteInt16(SvrConfig.GetValue("GoldMult"))
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShowFrmMapa_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowFrmMapa", Erl)
End Sub

Public Sub WritePreguntaBox(ByVal UserIndex As Integer, ByVal MsgID As Integer, Optional ByVal Param As String = vbNullString)
    On Error GoTo WritePreguntaBox_Err
    Call Writer.WriteInt16(ServerPacketID.eShowPregunta)
    Call Writer.WriteInt16(MsgID)           ' Enviar el ID
    Call Writer.WriteString8(Param)          ' Enviar el parámetro (puede ser vacío)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WritePreguntaBox_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePreguntaBox", Erl)
End Sub

Public Sub WriteDatosGrupo(ByVal UserIndex As Integer)
    On Error GoTo WriteDatosGrupo_Err
    Dim i As Byte
    With UserList(UserIndex)
        Call Writer.WriteInt16(ServerPacketID.eDatosGrupo)
        Call Writer.WriteBool(.Grupo.EnGrupo)
        If .Grupo.EnGrupo = True Then
            Call Writer.WriteInt8(UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros)
            If .Grupo.Lider.ArrayIndex = UserIndex Then
                For i = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
                    If i = 1 Then
                        Call Writer.WriteString8(UserList(.Grupo.Miembros(i).ArrayIndex).name & "(Líder)")
                    Else
                        Call Writer.WriteString8(UserList(.Grupo.Miembros(i).ArrayIndex).name)
                    End If
                Next i
            Else
                For i = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
                    If i = 1 Then
                        Call Writer.WriteString8(UserList(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex).name & "(Líder)")
                    Else
                        Call Writer.WriteString8(UserList(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex).name)
                    End If
                Next i
            End If
        End If
    End With
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteDatosGrupo_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDatosGrupo", Erl)
End Sub

Public Sub WriteUbicacion(ByVal UserIndex As Integer, ByVal Miembro As Byte, ByVal GPS As Integer)
    On Error GoTo WriteUbicacion_Err
    Call Writer.WriteInt16(ServerPacketID.eubicacion)
    Call Writer.WriteInt8(Miembro)
    If GPS > 0 Then
        Call Writer.WriteInt8(UserList(GPS).pos.x)
        Call Writer.WriteInt8(UserList(GPS).pos.y)
        Call Writer.WriteInt16(UserList(GPS).pos.Map)
    Else
        Call Writer.WriteInt8(0)
        Call Writer.WriteInt8(0)
        Call Writer.WriteInt16(0)
    End If
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUbicacion_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUbicacion", Erl)
End Sub

Public Sub WriteViajarForm(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    On Error GoTo WriteViajarForm_Err
    Call Writer.WriteInt16(ServerPacketID.eViajarForm)
    Dim destinos As Byte
    Dim i        As Byte
    destinos = NpcList(NpcIndex).NumDestinos
    Call Writer.WriteInt8(destinos)
    For i = 1 To destinos
        Call Writer.WriteString8(NpcList(NpcIndex).dest(i))
    Next i
    Call Writer.WriteInt8(NpcList(NpcIndex).Interface)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteViajarForm_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteViajarForm", Erl)
End Sub

Public Sub WriteQuestDetails(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, Optional QuestSlot As Byte = 0)
    On Error GoTo WriteQuestDetails_Err
    Dim i As Integer
    'ID del paquete
    Call Writer.WriteInt16(ServerPacketID.eQuestDetails)
    'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptí todavía (1 para el primer caso y 0 para el segundo)
    Call Writer.WriteInt8(IIf(QuestSlot, 1, 0))
    'Enviamos nombre, descripción y nivel requerido de la quest
    'Call Writer.WriteString8(QuestList(QuestIndex).Nombre)
    'Call Writer.WriteString8(QuestList(QuestIndex).Desc)
    Call Writer.WriteInt16(QuestIndex)
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredLevel)
    Call Writer.WriteInt16(QuestList(QuestIndex).RequiredQuest)
    'Enviamos la cantidad de npcs requeridos
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredNPCs)
    If QuestList(QuestIndex).RequiredNPCs Then
        'Si hay npcs entonces enviamos la lista
        For i = 1 To QuestList(QuestIndex).RequiredNPCs
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).amount)
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)
            'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
            If QuestSlot Then
                Call Writer.WriteInt16(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
            End If
        Next i
    End If
    'Enviamos la cantidad de objs requeridos
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredOBJs)
    If QuestList(QuestIndex).RequiredOBJs Then
        'Si hay objs entonces enviamos la lista
        For i = 1 To QuestList(QuestIndex).RequiredOBJs
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).amount)
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex)
            'escribe si tiene ese objeto en el inventario y que cantidad
            Call Writer.WriteInt16(get_object_amount_from_inventory(UserIndex, QuestList(QuestIndex).RequiredOBJ(i).ObjIndex))
            ' Call Writer.WriteInt16(0)
        Next i
    End If
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.SkillType)
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.RequiredValue)
    'Enviamos la recompensa de oro y experiencia.
    Call Writer.WriteInt32((QuestList(QuestIndex).RewardGLD * SvrConfig.GetValue("GoldMult")))
    Call Writer.WriteInt32((QuestList(QuestIndex).RewardEXP * SvrConfig.GetValue("ExpMult")))
    'Enviamos la cantidad de objs de recompensa
    Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)
    If QuestList(QuestIndex).RewardOBJs Then
        'si hay objs entonces enviamos la lista
        For i = 1 To QuestList(QuestIndex).RewardOBJs
            Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).amount)
            Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
        Next i
    End If
    Call Writer.WriteInt8(QuestList(QuestIndex).RewardSpellCount)
    For i = 1 To QuestList(QuestIndex).RewardSpellCount
        Writer.WriteInt16 (QuestList(QuestIndex).RewardSpellList(i))
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteQuestDetails_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteQuestDetails", Erl)
End Sub
 
Public Sub WriteQuestListSend(ByVal UserIndex As Integer)
    On Error GoTo WriteQuestListSend_Err
    Dim i       As Integer
    Dim tmpStr  As String
    Dim tmpByte As Byte
    With UserList(UserIndex)
        Call Writer.WriteInt16(ServerPacketID.eQuestListSend)
        For i = 1 To MAXUSERQUESTS
            If .QuestStats.Quests(i).QuestIndex Then
                tmpByte = tmpByte + 1
                tmpStr = tmpStr & .QuestStats.Quests(i).QuestIndex & ";"
            End If
        Next i
        'Escribimos la cantidad de quests
        Call Writer.WriteInt8(tmpByte)
        'Escribimos la lista de quests (sacamos el íltimo caracter)
        If tmpByte Then
            Call Writer.WriteString8(Left$(tmpStr, Len(tmpStr) - 1))
        End If
    End With
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteQuestListSend_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteQuestListSend", Erl)
End Sub

Public Sub WriteNpcQuestListSend(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    On Error GoTo WriteNpcQuestListSend_Err
    Dim i          As Integer
    Dim j          As Integer
    Dim QuestIndex As Integer
    Call Writer.WriteInt16(ServerPacketID.eNpcQuestListSend)
    Call Writer.WriteInt8(NpcList(NpcIndex).NumQuest) 'Escribimos primero cuantas quest tiene el NPC
    For j = 1 To NpcList(NpcIndex).NumQuest
        QuestIndex = NpcList(NpcIndex).QuestNumber(j)
        Call Writer.WriteInt16(QuestIndex)
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredLevel)
        Call Writer.WriteInt16(QuestList(QuestIndex).RequiredQuest)
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredClass)
        Call Writer.WriteInt8(QuestList(QuestIndex).LimitLevel)
        'Enviamos la cantidad de npcs requeridos
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredNPCs)
        If QuestList(QuestIndex).RequiredNPCs Then
            'Si hay npcs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).amount)
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)
                'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
                'If QuestSlot Then
                ' Call Writer.WriteInt16(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
                ' End If
            Next i
        End If
        'Enviamos la cantidad de objs requeridos
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredOBJs)
        If QuestList(QuestIndex).RequiredOBJs Then
            'Si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredOBJs
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).amount)
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex)
            Next i
        End If
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSpellCount)
        If QuestList(QuestIndex).RequiredSpellCount > 0 Then
            For i = 1 To QuestList(QuestIndex).RequiredSpellCount
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredSpellList(i))
            Next i
        End If
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.SkillType)
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.RequiredValue)
        'Enviamos la recompensa de oro y experiencia.
        Call Writer.WriteInt32(QuestList(QuestIndex).RewardGLD * SvrConfig.GetValue("GoldMult"))
        Call Writer.WriteInt32(QuestList(QuestIndex).RewardEXP * SvrConfig.GetValue("ExpMult"))
        'Enviamos la cantidad de objs de recompensa
        Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)
        If QuestList(QuestIndex).RewardOBJs Then
            'si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RewardOBJs
                Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).amount)
                Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
            Next i
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
        If TieneQuest(UserIndex, QuestIndex) Then
            Call Writer.WriteInt8(1)
        Else
            If UserDoneQuest(UserIndex, QuestIndex) Then
                Call Writer.WriteInt8(2)
            Else
                PuedeHacerla = True
                If QuestList(QuestIndex).RequiredQuest > 0 Then
                    If Not UserDoneQuest(UserIndex, QuestList(QuestIndex).RequiredQuest) Then
                        PuedeHacerla = False
                    End If
                End If
                If QuestList(QuestIndex).RequiredLevel > 0 Then
                    If UserList(UserIndex).Stats.ELV < QuestList(QuestIndex).RequiredLevel Then
                        PuedeHacerla = False
                    End If
                End If
                'Si el personaje es nivel mayor al limite no puede hacerla
                If QuestList(QuestIndex).LimitLevel > 0 Then
                    If UserList(UserIndex).Stats.ELV > QuestList(QuestIndex).LimitLevel Then
                        PuedeHacerla = False
                    End If
                End If
                If PuedeHacerla Then
                    Call Writer.WriteInt8(0)
                Else
                    Call Writer.WriteInt8(3)
                End If
            End If
        End If
    Next j
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteNpcQuestListSend_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNpcQuestListSend", Erl)
End Sub

Sub WriteCommerceRecieveChatMessage(ByVal UserIndex As Integer, ByVal Message As String)
    On Error GoTo WriteCommerceRecieveChatMessage_Err
    Call Writer.WriteInt16(ServerPacketID.eCommerceRecieveChatMessage)
    Call Writer.WriteString8(Message)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCommerceRecieveChatMessage_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCommerceRecieveChatMessage", Erl)
End Sub

Sub WriteInvasionInfo(ByVal UserIndex As Integer, ByVal Invasion As Integer, ByVal PorcentajeVida As Byte, ByVal PorcentajeTiempo As Byte)
    On Error GoTo WriteInvasionInfo_Err
    Call Writer.WriteInt16(ServerPacketID.eInvasionInfo)
    Call Writer.WriteInt8(Invasion)
    Call Writer.WriteInt8(PorcentajeVida)
    Call Writer.WriteInt8(PorcentajeTiempo)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteInvasionInfo_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteInvasionInfo", Erl)
End Sub

Sub WriteOpenCrafting(ByVal UserIndex As Integer, ByVal Tipo As Byte)
    On Error GoTo WriteOpenCrafting_Err
    Call Writer.WriteInt16(ServerPacketID.eOpenCrafting)
    Call Writer.WriteInt8(Tipo)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteOpenCrafting_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteOpenCrafting", Erl)
End Sub

Sub WriteCraftingItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ObjIndex As Integer)
    On Error GoTo WriteCraftingItem_Err
    Call Writer.WriteInt16(ServerPacketID.eCraftingItem)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(ObjIndex)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCraftingItem_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCraftingItem", Erl)
End Sub

Sub WriteCraftingCatalyst(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal amount As Integer, ByVal Porcentaje As Byte)
    On Error GoTo WriteCraftingCatalyst_Err
    Call Writer.WriteInt16(ServerPacketID.eCraftingCatalyst)
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt16(amount)
    Call Writer.WriteInt8(Porcentaje)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCraftingCatalyst_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCraftingCatalyst", Erl)
End Sub

Sub WriteCraftingResult(ByVal UserIndex As Integer, ByVal Result As Integer, Optional ByVal Porcentaje As Byte = 0, Optional ByVal precio As Long = 0)
    On Error GoTo WriteCraftingResult_Err
    Call Writer.WriteInt16(ServerPacketID.eCraftingResult)
    Call Writer.WriteInt16(Result)
    If Result <> 0 Then
        Call Writer.WriteInt8(Porcentaje)
        Call Writer.WriteInt32(precio)
    End If
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteCraftingResult_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCraftingResult", Erl)
End Sub

Public Sub WriteUpdateNPCSimbolo(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Simbolo As Byte)
    On Error GoTo WriteUpdateNPCSimbolo_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateNPCSimbolo)
    Call Writer.WriteInt16(NpcList(NpcIndex).Char.charindex)
    Call Writer.WriteInt8(Simbolo)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteUpdateNPCSimbolo_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateNPCSimbolo", Erl)
End Sub

Public Sub WriteGuardNotice(ByVal UserIndex As Integer)
    On Error GoTo WriteGuardNotice_Err
    Call Writer.WriteInt16(ServerPacketID.eGuardNotice)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteGuardNotice_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuardNotice", Erl)
End Sub

' \Begin: [Prepares]
Public Function PrepareMessageCharSwing(ByVal charindex As Integer, _
                                        Optional ByVal FX As Boolean = True, _
                                        Optional ByVal ShowText As Boolean = True, _
                                        Optional ByVal NotificoTexto As Boolean = True)
    On Error GoTo PrepareMessageCharSwing_Err
    Call Writer.WriteInt16(ServerPacketID.eCharSwing)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteBool(FX)
    Call Writer.WriteBool(ShowText)
    Call Writer.WriteBool(NotificoTexto)
    Exit Function
PrepareMessageCharSwing_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharSwing", Erl)
End Function

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.
Public Function PrepareMessageSetInvisible(ByVal charindex As Integer, ByVal invisible As Boolean, Optional ByVal x As Byte = 0, Optional ByVal y As Byte = 0)
    On Error GoTo PrepareMessageSetInvisible_Err
    Call Writer.WriteInt16(ServerPacketID.eSetInvisible)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteBool(invisible)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareMessageSetInvisible_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageSetInvisible", Erl)
End Function

Public Function PrepareLocaleChatOverHead(ByVal chat As Integer, _
                                          ByVal Params As String, _
                                          ByVal charindex As Integer, _
                                          ByVal Color As Long, _
                                          Optional ByVal EsSpell As Boolean = False, _
                                          Optional ByVal x As Byte = 0, _
                                          Optional ByVal y As Byte = 0, _
                                          Optional ByVal RequiredMinDisplayTime As Integer = 0, _
                                          Optional ByVal MaxDisplayTime As Integer = 0)
    On Error GoTo PrepareMessageChatOverHead_Err
    Call Writer.WriteInt16(ServerPacketID.eLocaleChatOverHead)
    Call Writer.WriteInt16(chat)
    Call Writer.WriteString8(Params)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt32(Color)
    Call Writer.WriteBool(EsSpell)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt16(RequiredMinDisplayTime)
    Call Writer.WriteInt16(MaxDisplayTime)
    Exit Function
PrepareMessageChatOverHead_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageChatOverHead", Erl)
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
                                           ByVal charindex As Integer, _
                                           ByVal Color As Long, _
                                           Optional ByVal EsSpell As Boolean = False, _
                                           Optional ByVal x As Byte = 0, _
                                           Optional ByVal y As Byte = 0, _
                                           Optional ByVal RequiredMinDisplayTime As Integer = 0, _
                                           Optional ByVal MaxDisplayTime As Integer = 0)
    On Error GoTo PrepareMessageChatOverHead_Err
    Call Writer.WriteInt16(ServerPacketID.eChatOverHead)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt32(Color)
    Call Writer.WriteBool(EsSpell)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt16(RequiredMinDisplayTime)
    Call Writer.WriteInt16(MaxDisplayTime)
    Exit Function
PrepareMessageChatOverHead_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageChatOverHead", Erl)
End Function

Public Function PrepareMessageTextOverChar(ByVal chat As String, ByVal charindex As Integer, ByVal Color As Long)
    On Error GoTo PrepareMessageTextOverChar_Err
    Call Writer.WriteInt16(ServerPacketID.eTextOverChar)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt32(Color)
    Exit Function
PrepareMessageTextOverChar_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageTextOverChar", Erl)
End Function

Public Function PrepareMessageTextCharDrop(ByVal chat As String, _
                                           ByVal charindex As Integer, _
                                           ByVal Color As Long, _
                                           Optional ByVal Duration As Integer = 1300, _
                                           Optional ByVal Animated As Boolean = True)
    On Error GoTo PrepareMessageTextCharDrop_Err
    Call Writer.WriteInt16(ServerPacketID.eTextCharDrop)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt32(Color)
    Call Writer.WriteInt16(Duration)
    Call Writer.WriteBool(Animated)
    Exit Function
PrepareMessageTextCharDrop_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageTextCharDrop", Erl)
End Function

Public Function PrepareMessageTextOverTile(ByVal chat As String, _
                                           ByVal x As Integer, _
                                           ByVal y As Integer, _
                                           ByVal Color As Long, _
                                           Optional ByVal Duration As Integer = 1300, _
                                           Optional ByVal OffsetY As Integer = 0, _
                                           Optional ByVal Animated As Boolean = True)
    On Error GoTo PrepareMessageTextOverTile_Err
    Call Writer.WriteInt16(ServerPacketID.eTextOverTile)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt16(x)
    Call Writer.WriteInt16(y)
    Call Writer.WriteInt32(Color)
    Call Writer.WriteInt16(Duration)
    Call Writer.WriteInt16(OffsetY)
    Call Writer.WriteBool(Animated)
    Exit Function
PrepareMessageTextOverTile_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageTextOverTile", Erl)
End Function

Public Function PrepareConsoleCharText(ByVal chat As String, ByVal Color As Long, ByVal sourceName As String, ByVal sourceStatus As Integer, ByVal privileges As Integer)
    On Error GoTo PrepareConsoleCharText_Err
    Call Writer.WriteInt16(ServerPacketID.eConsoleCharText)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt32(Color)
    Call Writer.WriteString8(sourceName)
    Call Writer.WriteInt16(sourceStatus)
    Call Writer.WriteInt16(privileges)
    Exit Function
PrepareConsoleCharText_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareConsoleCharText", Erl)
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageConsoleMsg(ByVal chat As String, ByVal FontIndex As e_FontTypeNames)
    On Error GoTo PrepareMessageConsoleMsg_Err
    Call Writer.WriteInt16(ServerPacketID.eConsoleMsg)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt8(FontIndex)
    Exit Function
PrepareMessageConsoleMsg_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageConsoleMsg", Erl)
End Function

Public Function PrepareFactionMessageConsole(ByVal factionLabel As String, ByVal chat As String, ByVal FontIndex As e_FontTypeNames)
    On Error GoTo PrepareFactionMessageConsole_Err
    Call Writer.WriteInt16(ServerPacketID.eConsoleFactionMessage)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt8(FontIndex)
    Call Writer.WriteString8(factionLabel)
    Exit Function
PrepareFactionMessageConsole_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareFactionMessageConsole", Erl)
End Function

Public Function PrepareMessageLocaleMsg(ByVal Id As Integer, ByVal chat As String, ByVal FontIndex As e_FontTypeNames)
    On Error GoTo PrepareMessageLocaleMsg_Err
    Call Writer.WriteInt16(ServerPacketID.eLocaleMsg)
    Call Writer.WriteInt16(Id)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt8(FontIndex)
    Exit Function
PrepareMessageLocaleMsg_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageLocaleMsg", Erl)
End Function

''
' Prepares the "CharAtaca" message and returns it.
'
Public Function PrepareMessageCharAtaca(ByVal charindex As Integer, ByVal attackerIndex As Integer, ByVal danio As Long, ByVal AnimAttack As Integer)
    On Error GoTo PrepareMessageCharAtaca_Err
    Call Writer.WriteInt16(ServerPacketID.eCharAtaca)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(attackerIndex)
    Call Writer.WriteInt32(danio)
    Call Writer.WriteInt16(AnimAttack)
    Exit Function
PrepareMessageCharAtaca_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharAtaca", Erl)
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
Public Function PrepareMessageCreateFX(ByVal charindex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, Optional ByVal x As Byte = 0, Optional ByVal y As Byte = 0)
    On Error GoTo PrepareMessageCreateFX_Err
    Call Writer.WriteInt16(ServerPacketID.eCreateFX)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt16(FXLoops)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareMessageCreateFX_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCreateFX", Erl)
End Function

Public Function PrepareMessageMeditateToggle(ByVal charindex As Integer, ByVal FX As Integer, Optional ByVal x As Byte = 0, Optional ByVal y As Byte = 0)
    On Error GoTo PrepareMessageMeditateToggle_Err
    Call Writer.WriteInt16(ServerPacketID.eMeditateToggle)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareMessageMeditateToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageMeditateToggle", Erl)
End Function

Public Function PrepareMessageParticleFX(ByVal charindex As Integer, _
                                         ByVal Particula As Integer, _
                                         ByVal Time As Long, _
                                         ByVal Remove As Boolean, _
                                         Optional ByVal grh As Long = 0, _
                                         Optional ByVal x As Byte = 0, _
                                         Optional ByVal y As Byte = 0)
    On Error GoTo PrepareMessageParticleFX_Err
    Call Writer.WriteInt16(ServerPacketID.eParticleFX)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(Particula)
    Call Writer.WriteInt32(Time)
    Call Writer.WriteBool(Remove)
    Call Writer.WriteInt32(grh)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareMessageParticleFX_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFX", Erl)
End Function

Public Function PrepareMessageParticleFXWithDestino(ByVal Emisor As Integer, _
                                                    ByVal Receptor As Integer, _
                                                    ByVal ParticulaViaje As Integer, _
                                                    ByVal ParticulaFinal As Integer, _
                                                    ByVal Time As Long, _
                                                    ByVal wav As Integer, _
                                                    ByVal FX As Integer, _
                                                    Optional ByVal x As Byte = 0, _
                                                    Optional ByVal y As Byte = 0)
    On Error GoTo PrepareMessageParticleFXWithDestino_Err
    Call Writer.WriteInt16(ServerPacketID.eParticleFXWithDestino)
    Call Writer.WriteInt16(Emisor)
    Call Writer.WriteInt16(Receptor)
    Call Writer.WriteInt16(ParticulaViaje)
    Call Writer.WriteInt16(ParticulaFinal)
    Call Writer.WriteInt32(Time)
    Call Writer.WriteInt16(wav)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareMessageParticleFXWithDestino_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFXWithDestino", Erl)
End Function

Public Function PrepareMessageParticleFXWithDestinoXY(ByVal Emisor As Integer, _
                                                      ByVal ParticulaViaje As Integer, _
                                                      ByVal ParticulaFinal As Integer, _
                                                      ByVal Time As Long, _
                                                      ByVal wav As Integer, _
                                                      ByVal FX As Integer, _
                                                      ByVal x As Byte, _
                                                      ByVal y As Byte)
    On Error GoTo PrepareMessageParticleFXWithDestinoXY_Err
    Call Writer.WriteInt16(ServerPacketID.eParticleFXWithDestinoXY)
    Call Writer.WriteInt16(Emisor)
    Call Writer.WriteInt16(ParticulaViaje)
    Call Writer.WriteInt16(ParticulaFinal)
    Call Writer.WriteInt32(Time)
    Call Writer.WriteInt16(wav)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareMessageParticleFXWithDestinoXY_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFXWithDestinoXY", Erl)
End Function

Public Function PrepareMessageAuraToChar(ByVal charindex As Integer, ByVal Aura As String, ByVal Remove As Boolean, ByVal Tipo As Byte)
    On Error GoTo PrepareMessageAuraToChar_Err
    Call Writer.WriteInt16(ServerPacketID.eAuraToChar)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteString8(Aura)
    Call Writer.WriteBool(Remove)
    Call Writer.WriteInt8(Tipo)
    Exit Function
PrepareMessageAuraToChar_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageAuraToChar", Erl)
End Function

Public Function PrepareMessageSpeedingACT(ByVal charindex As Integer, ByVal speeding As Single)
    On Error GoTo PrepareMessageSpeedingACT_Err
    Call Writer.WriteInt16(ServerPacketID.eSpeedToChar)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteReal32(speeding)
    Exit Function
PrepareMessageSpeedingACT_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageSpeedingACT", Erl)
End Function

Public Function PrepareMessageParticleFXToFloor(ByVal x As Byte, ByVal y As Byte, ByVal Particula As Integer, ByVal Time As Long)
    On Error GoTo PrepareMessageParticleFXToFloor_Err
    Call Writer.WriteInt16(ServerPacketID.eParticleFXToFloor)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt16(Particula)
    Call Writer.WriteInt32(Time)
    Exit Function
PrepareMessageParticleFXToFloor_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFXToFloor", Erl)
End Function

Public Function PrepareMessageLightFXToFloor(ByVal x As Byte, ByVal y As Byte, ByVal LuzColor As Long, ByVal Rango As Byte)
    On Error GoTo PrepareMessageLightFXToFloor_Err
    Call Writer.WriteInt16(ServerPacketID.eLightToFloor)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt32(LuzColor)
    Call Writer.WriteInt8(Rango)
    Exit Function
PrepareMessageLightFXToFloor_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageLightFXToFloor", Erl)
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessagePlayWave(ByVal wave As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal CancelLastWave As Byte = False, Optional ByVal Localize As Byte = 0)
    On Error GoTo PrepareMessagePlayWave_Err
    Call Writer.WriteInt16(ServerPacketID.ePlayWave)
    Call Writer.WriteInt16(wave)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt8(CancelLastWave)
    Call Writer.WriteInt8(Localize)
    Exit Function
PrepareMessagePlayWave_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessagePlayWave", Erl)
End Function

Public Function PrepareMessageUbicacionLlamada(ByVal Mapa As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo PrepareMessageUbicacionLlamada_Err
    Call Writer.WriteInt16(ServerPacketID.ePosLLamadaDeClan)
    Call Writer.WriteInt16(Mapa)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareMessageUbicacionLlamada_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageUbicacionLlamada", Erl)
End Function

Public Function PrepareMessageCharUpdateHP(ByVal UserIndex As Integer)
    On Error GoTo PrepareMessageCharUpdateHP_Err
    Call Writer.WriteInt16(ServerPacketID.eCharUpdateHP)
    Call Writer.WriteInt16(UserList(UserIndex).Char.charindex)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.MinHp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.MaxHp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.shield)
    Exit Function
PrepareMessageCharUpdateHP_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharUpdateHP", Erl)
End Function

Public Function PrepareMessageCharUpdateMAN(ByVal UserIndex As Integer)
    On Error GoTo PrepareMessageCharUpdateMAN_Err
    Call Writer.WriteInt16(ServerPacketID.eCharUpdateMAN)
    Call Writer.WriteInt16(UserList(UserIndex).Char.charindex)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.MinMAN)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.MaxMAN)
    Exit Function
PrepareMessageCharUpdateMAN_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharUpdateMAN", Erl)
End Function

Public Function PrepareMessageNpcUpdateHP(ByVal NpcIndex As Integer)
    On Error GoTo PrepareMessageNpcUpdateHP_Err
    Call Writer.WriteInt16(ServerPacketID.eCharUpdateHP)
    Call Writer.WriteInt16(NpcList(NpcIndex).Char.charindex)
    Call Writer.WriteInt32(NpcList(NpcIndex).Stats.MinHp)
    Call Writer.WriteInt32(NpcList(NpcIndex).Stats.MaxHp)
    Call Writer.WriteInt32(NpcList(NpcIndex).Stats.shield)
    Exit Function
PrepareMessageNpcUpdateHP_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageNpcUpdateHP", Erl)
End Function

Public Function PrepareMessageArmaMov(ByVal charindex As Integer, Optional ByVal isRanged As Byte = 0)
    On Error GoTo PrepareMessageArmaMov_Err
    Call Writer.WriteInt16(ServerPacketID.eArmaMov)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt8(isRanged)
    Exit Function
PrepareMessageArmaMov_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageArmaMov", Erl)
End Function

Public Function PrepareCreateProjectile(ByVal startX As Byte, ByVal startY As Byte, ByVal TargetX As Byte, ByVal TargetY As Byte, ByVal ProjectileType As Byte)
    On Error GoTo PrepareCreateProjectile_Err
    Call Writer.WriteInt16(ServerPacketID.eCreateProjectile)
    Call Writer.WriteInt8(startX)
    Call Writer.WriteInt8(startY)
    Call Writer.WriteInt8(TargetX)
    Call Writer.WriteInt8(TargetY)
    Call Writer.WriteInt8(ProjectileType)
    Exit Function
PrepareCreateProjectile_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareCreateProjectile", Erl)
End Function

Public Function PrepareMessageEscudoMov(ByVal charindex As Integer)
    On Error GoTo PrepareMessageEscudoMov_Err
    Call Writer.WriteInt16(ServerPacketID.eEscudoMov)
    Call Writer.WriteInt16(charindex)
    Exit Function
PrepareMessageEscudoMov_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageEscudoMov", Erl)
End Function

Public Function PrepareMessageFlashScreen(ByVal Color As Long, ByVal Duracion As Long, Optional ByVal Ignorar As Boolean = False)
    On Error GoTo PrepareMessageFlashScreen_Err
    Call Writer.WriteInt16(ServerPacketID.eFlashScreen)
    Call Writer.WriteInt32(Color)
    Call Writer.WriteInt32(Duracion)
    Call Writer.WriteBool(Ignorar)
    Exit Function
PrepareMessageFlashScreen_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageFlashScreen", Erl)
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageGuildChat(ByVal chat As String, ByVal Status As Byte)
    On Error GoTo PrepareMessageGuildChat_Err
    Call Writer.WriteInt16(ServerPacketID.eGuildChat)
    Call Writer.WriteInt8(Status)
    Call Writer.WriteString8(chat)
    Exit Function
PrepareMessageGuildChat_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageGuildChat", Erl)
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageShowMessageBox(ByVal chat As String)
    On Error GoTo PrepareMessageShowMessageBox_Err
    Call Writer.WriteInt16(ServerPacketID.eShowMessageBox)
    Call Writer.WriteString8(chat)
    Exit Function
PrepareMessageShowMessageBox_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageShowMessageBox", Erl)
End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1)
    On Error GoTo PrepareMessagePlayMidi_Err
    Call Writer.WriteInt16(ServerPacketID.ePlayMIDI)
    Call Writer.WriteInt8(midi)
    Call Writer.WriteInt16(loops)
    Exit Function
PrepareMessagePlayMidi_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessagePlayMidi", Erl)
End Function

Public Function PrepareMessageOnlineUser(ByVal UserOnline As Integer)
    On Error GoTo PrepareMessageOnlineUser_Err
    Call Writer.WriteInt16(ServerPacketID.eUserOnline)
    Call Writer.WriteInt16(UserOnline)
    Exit Function
PrepareMessageOnlineUser_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageOnlineUser", Erl)
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessagePauseToggle()
    On Error GoTo PrepareMessagePauseToggle_Err
    Call Writer.WriteInt16(ServerPacketID.ePauseToggle)
    Exit Function
PrepareMessagePauseToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessagePauseToggle", Erl)
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageRainToggle()
    On Error GoTo PrepareMessageRainToggle_Err
    Call Writer.WriteInt16(ServerPacketID.eRainToggle)
    Call Writer.WriteBool(Lloviendo)
    Exit Function
PrepareMessageRainToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageRainToggle", Erl)
End Function

Public Function PrepareMessageHora()
    On Error GoTo PrepareMessageHora_Err
    Dim dayLen As Long
    dayLen = CLng(SvrConfig.GetValue("DayLength"))
    If dayLen <= 0 Then dayLen = 1
    ' Ensure WorldTime is initialized at server start, e.g. WorldTime_Init dayLen, 0
    Dim t As Long, d As Long
    WorldTime_PrepareHora t, d
    Call Writer.WriteInt16(ServerPacketID.ehora)
    Call Writer.WriteInt32(t)   ' elapsed 0..dayLen-1
    Call Writer.WriteInt32(d)   ' dayLen
    Exit Function
PrepareMessageHora_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageHora", Erl)
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageObjectDelete(ByVal x As Byte, ByVal y As Byte)
    On Error GoTo PrepareMessageObjectDelete_Err
    Call Writer.WriteInt16(ServerPacketID.eObjectDelete)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareMessageObjectDelete_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageObjectDelete", Erl)
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessage_BlockPosition(ByVal x As Byte, ByVal y As Byte, ByVal Blocked As Byte)
    On Error GoTo PrepareMessage_BlockPosition_Err
    Call Writer.WriteInt16(ServerPacketID.eBlockPosition)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt8(Blocked)
    Exit Function
PrepareMessage_BlockPosition_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessage_BlockPosition", Erl)
End Function

Public Function PrepareTrapUpdate(ByVal State As Byte, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo PrepareTrapUpdate_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateTrap)
    Call Writer.WriteInt8(State)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareTrapUpdate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareTrapUpdate", Erl)
End Function

Public Function PrepareUpdateGroupInfo(ByVal UserIndex As Integer)
    On Error GoTo PrepareTrapUpdate_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateGroupInfo)
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
                                           ByVal amount As Integer, _
                                           ByVal x As Byte, _
                                           ByVal y As Byte, _
                                           Optional ByVal ElementalTags As Long = e_ElementalTags.Normal)
    On Error GoTo PrepareMessageObjectCreate_Err
    Call Writer.WriteInt16(ServerPacketID.eObjectCreate)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt16(amount)
    Call Writer.WriteInt32(ElementalTags)
    Exit Function
PrepareMessageObjectCreate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageObjectCreate", Erl)
End Function

Public Function PrepareMessageFxPiso(ByVal GrhIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo PrepareMessageFxPiso_Err
    Call Writer.WriteInt16(ServerPacketID.efxpiso)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt16(GrhIndex)
    Exit Function
PrepareMessageFxPiso_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageFxPiso", Erl)
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageCharacterRemove(ByVal dbgid As Integer, ByVal charindex As Integer, ByVal Desvanecido As Boolean, Optional ByVal FueWarp As Boolean = False)
    On Error GoTo PrepareMessageCharacterRemove_Err
    Call Writer.WriteInt16(ServerPacketID.eCharacterRemove)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteBool(Desvanecido)
    Call Writer.WriteBool(FueWarp)
    Exit Function
PrepareMessageCharacterRemove_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterRemove", Erl)
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageRemoveCharDialog(ByVal charindex As Integer)
    On Error GoTo PrepareMessageRemoveCharDialog_Err
    Call Writer.WriteInt16(ServerPacketID.eRemoveCharDialog)
    Call Writer.WriteInt16(charindex)
    Exit Function
PrepareMessageRemoveCharDialog_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageRemoveCharDialog", Erl)
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
Public Function PrepareMessageCharacterCreate(ByVal body As Integer, _
                                              ByVal head As Integer, _
                                              ByVal Heading As e_Heading, _
                                              ByVal charindex As Integer, _
                                              ByVal x As Byte, _
                                              ByVal y As Byte, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal Cart As Integer, _
                                              ByVal BackPack As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal name As String, _
                                              ByVal Status As Byte, _
                                              ByVal privileges As Byte, _
                                              ByVal ParticulaFx As Byte, _
                                              ByVal Head_Aura As String, _
                                              ByVal Arma_Aura As String, _
                                              ByVal Body_Aura As String, _
                                              ByVal DM_Aura As String, _
                                              ByVal RM_Aura As String, _
                                              ByVal Otra_Aura As String, _
                                              ByVal Escudo_Aura As String, _
                                              ByVal speeding As Single, ByVal EsNPC As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal UserMinMAN As Long, ByVal UserMaxMAN As Long, ByVal Simbolo As Byte, ByVal Idle As Boolean, ByVal Navegando As Boolean, ByVal tipoUsuario As e_TipoUsuario, Optional ByVal TeamCaptura As Byte = 0, Optional ByVal TieneBandera As Byte = 0, Optional ByVal AnimAtaque1 As Integer = 0)
    On Error GoTo PrepareMessageCharacterCreate_Err
    Call Writer.WriteInt16(ServerPacketID.eCharacterCreate)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(body)
    Call Writer.WriteInt16(head)
    Call Writer.WriteInt8(Heading)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt16(weapon)
    Call Writer.WriteInt16(shield)
    Call Writer.WriteInt16(helmet)
    Call Writer.WriteInt16(Cart)
    Call Writer.WriteInt16(BackPack)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt16(FXLoops)
    Call Writer.WriteString8(name)
    Call Writer.WriteInt8(Status)
    Call Writer.WriteInt8(privileges)
    Call Writer.WriteInt8(ParticulaFx)
    Call Writer.WriteString8(Head_Aura)
    Call Writer.WriteString8(Arma_Aura)
    Call Writer.WriteString8(Body_Aura)
    Call Writer.WriteString8(DM_Aura)
    Call Writer.WriteString8(RM_Aura)
    Call Writer.WriteString8(Otra_Aura)
    Call Writer.WriteString8(Escudo_Aura)
    Call Writer.WriteReal32(speeding)
    Call Writer.WriteInt8(EsNPC)
    Call Writer.WriteInt8(appear)
    Call Writer.WriteInt16(group_index)
    Call Writer.WriteInt16(clan_index)
    Call Writer.WriteInt8(clan_nivel)
    Call Writer.WriteInt32(UserMinHp)
    Call Writer.WriteInt32(UserMaxHp)
    Call Writer.WriteInt32(UserMinMAN)
    Call Writer.WriteInt32(UserMaxMAN)
    Call Writer.WriteInt8(Simbolo)
    Dim flags As Byte
    flags = 0
    If Idle Then flags = flags Or &O1 ' 00000001
    If Navegando Then flags = flags Or &O2
    Call Writer.WriteInt8(flags)
    Call Writer.WriteInt8(tipoUsuario)
    Call Writer.WriteInt8(TeamCaptura)
    Call Writer.WriteInt8(TieneBandera)
    Call Writer.WriteInt16(AnimAtaque1)
    Exit Function
PrepareMessageCharacterCreate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterCreate", Erl)
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
Public Function PrepareMessageCharacterChange(ByVal body As Integer, _
                                              ByVal head As Integer, _
                                              ByVal Heading As e_Heading, _
                                              ByVal charindex As Integer, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal Cart As Integer, _
                                              ByVal BackPack As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal Idle As Boolean, _
                                              ByVal Navegando As Boolean)
    On Error GoTo PrepareMessageCharacterChange_Err
    Call Writer.WriteInt16(ServerPacketID.eCharacterChange)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(body)
    Call Writer.WriteInt16(head)
    Call Writer.WriteInt8(Heading)
    Call Writer.WriteInt16(weapon)
    Call Writer.WriteInt16(shield)
    Call Writer.WriteInt16(helmet)
    Call Writer.WriteInt16(Cart)
    Call Writer.WriteInt16(BackPack)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt16(FXLoops)
    Dim flags As Byte
    flags = 0
    If Idle Then flags = flags Or &O1
    If Navegando Then flags = flags Or &O2
    Call Writer.WriteInt8(flags)
    Exit Function
PrepareMessageCharacterChange_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterChange", Erl)
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
    On Error GoTo PrepareMessageUpdateFlag_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateFlag)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt8(Flag)
    Exit Function
PrepareMessageUpdateFlag_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageUpdateFlag", Erl)
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageCharacterMove(ByVal charindex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo PrepareMessageCharacterMove_Err
    Call Writer.WriteInt16(ServerPacketID.eCharacterMove)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Exit Function
PrepareMessageCharacterMove_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterMove", Erl)
End Function

Public Function PrepareCharacterTranslate(ByVal CharIndexm As Integer, ByVal NewX As Byte, ByVal NewY As Byte, ByVal TranslationTime As Long)
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
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As e_Heading)
    On Error GoTo PrepareMessageForceCharMove_Err
    Call Writer.WriteInt16(ServerPacketID.eForceCharMove)
    Call Writer.WriteInt8(Direccion)
    Exit Function
PrepareMessageForceCharMove_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageForceCharMove", Erl)
End Function


''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, Status As Byte, Tag As String)
    On Error GoTo PrepareMessageUpdateTagAndStatus_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateTagAndStatus)
    Call Writer.WriteInt16(UserList(UserIndex).Char.charindex)
    Call Writer.WriteInt8(Status)
    Call Writer.WriteString8(Tag)
    Call Writer.WriteInt16(UserList(UserIndex).Grupo.Lider.ArrayIndex)
    Exit Function
PrepareMessageUpdateTagAndStatus_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageUpdateTagAndStatus", Erl)
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageErrorMsg(ByVal Message As String)
    On Error GoTo PrepareMessageErrorMsg_Err
    Call Writer.WriteInt16(ServerPacketID.eErrorMsg)
    Call Writer.WriteString8(Message)
    Exit Function
PrepareMessageErrorMsg_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageErrorMsg", Erl)
End Function

Public Function PrepareMessageBarFx(ByVal charindex As Integer, ByVal BarTime As Integer, ByVal BarAccion As Byte)
    On Error GoTo PrepareMessageBarFx_Err
    Call Writer.WriteInt16(ServerPacketID.eBarFx)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(BarTime)
    Call Writer.WriteInt8(BarAccion)
    Exit Function
PrepareMessageBarFx_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageBarFx", Erl)
End Function

Public Function PrepareMessageNieblandoToggle(ByVal IntensidadMax As Byte)
    On Error GoTo PrepareMessageNieblandoToggle_Err
    Call Writer.WriteInt16(ServerPacketID.eNieblaToggle)
    Call Writer.WriteInt8(IntensidadMax)
    Exit Function
PrepareMessageNieblandoToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageNieblandoToggle", Erl)
End Function

Public Function PrepareMessageNevarToggle()
    On Error GoTo PrepareMessageNevarToggle_Err
    Call Writer.WriteInt16(ServerPacketID.eNieveToggle)
    Call Writer.WriteBool(Nebando)
    Exit Function
PrepareMessageNevarToggle_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageNevarToggle", Erl)
End Function

Public Function PrepareMessageDoAnimation(ByVal charindex As Integer, ByVal Animation As Integer)
    On Error GoTo PrepareMessageDoAnimation_Err
    Call Writer.WriteInt16(ServerPacketID.eDoAnimation)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(Animation)
    Exit Function
PrepareMessageDoAnimation_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageDoAnimation", Erl)
End Function



Public Sub WriteShopInit(ByVal UserIndex As Integer)
    On Error GoTo WriteShopInit_Err
    Dim i As Long, cant_obj_shop As Integer
    Call Writer.WriteInt16(ServerPacketID.eShopInit)
    cant_obj_shop = UBound(ObjShop)
    Call Writer.WriteInt16(cant_obj_shop)
    Call LoadPatronCreditsFromDB(UserIndex)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Creditos)
    'Envío todos los objetos.
    For i = 1 To cant_obj_shop
        Call Writer.WriteInt32(ObjShop(i).ObjNum)
        Call Writer.WriteInt32(ObjShop(i).Valor)
        Call Writer.WriteString8(ObjShop(i).name)
    Next i
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShopInit_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShopInit", Erl)
End Sub

Public Sub WriteShopPjsInit(ByVal UserIndex As Integer)
    On Error GoTo WriteShopPjsInit_Err
    Call Writer.WriteInt16(ServerPacketID.eShopPjsInit)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteShopPjsInit_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShopPjsInit", Erl)
End Sub

Public Sub writeUpdateShopClienteCredits(ByVal UserIndex As Integer)
    On Error GoTo writeUpdateShopClienteCredits_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateShopClienteCredits)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Creditos)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
writeUpdateShopClienteCredits_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description + " UI: " + UserIndex, "Argentum20Server.Protocol_Writes.writeUpdateShopClienteCredits", Erl)
End Sub

Public Sub WriteSendSkillCdUpdate(ByVal UserIndex As Integer, _
                                  ByVal SkillTypeId As Integer, _
                                  ByVal SkillId As Long, _
                                  ByVal TimeLeft As Long, _
                                  ByVal TotalTime As Long, _
                                  ByVal SkillType As e_EffectType, _
                                  Optional ByVal Stacks As Integer = 1)
    On Error GoTo WriteSendSkillCdUpdate_Err
    Call Writer.WriteInt16(ServerPacketID.eSendSkillCdUpdate)
    Call Writer.WriteInt16(SkillTypeId)
    Call Writer.WriteInt32(SkillId)
    Call Writer.WriteInt32(TimeLeft)
    Call Writer.WriteInt32(TotalTime)
    Call Writer.WriteInt8(ConvertToClientBuff(SkillType))
    Call Writer.WriteInt16(Stacks)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteSendSkillCdUpdate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description + " UI: " + UserIndex, "Argentum20Server.Protocol_Writes.writeUpdateShopClienteCredits", Erl)
End Sub

Public Sub WriteObjQuestSend(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal Slot As Byte)
    On Error GoTo WriteNpcQuestListSend_Err
    Dim i As Integer
    Call Writer.WriteInt16(ServerPacketID.eObjQuestListSend)
    Call Writer.WriteInt16(QuestIndex) 'Escribimos primero cuantas quest tiene el NPC
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredLevel)
    Call Writer.WriteInt16(QuestList(QuestIndex).RequiredQuest)
    'Enviamos la cantidad de npcs requeridos
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredNPCs)
    If QuestList(QuestIndex).RequiredNPCs Then
        'Si hay npcs entonces enviamos la lista
        For i = 1 To QuestList(QuestIndex).RequiredNPCs
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).amount)
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)
        Next i
    End If
    'Enviamos la cantidad de objs requeridos
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredOBJs)
    If QuestList(QuestIndex).RequiredOBJs Then
        'Si hay objs entonces enviamos la lista
        For i = 1 To QuestList(QuestIndex).RequiredOBJs
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).amount)
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex)
        Next i
    End If
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSpellCount)
    If QuestList(QuestIndex).RequiredSpellCount > 0 Then
        For i = 1 To QuestList(QuestIndex).RequiredSpellCount
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredSpellList(i))
        Next i
    End If
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.SkillType)
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredSkill.RequiredValue)
    'Enviamos la recompensa de oro y experiencia.
    Call Writer.WriteInt32(QuestList(QuestIndex).RewardGLD * SvrConfig.GetValue("GoldMult"))
    Call Writer.WriteInt32(QuestList(QuestIndex).RewardEXP * SvrConfig.GetValue("ExpMult"))
    'Enviamos la cantidad de objs de recompensa
    Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)
    If QuestList(QuestIndex).RewardOBJs Then
        'si hay objs entonces enviamos la lista
        For i = 1 To QuestList(QuestIndex).RewardOBJs
            Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).amount)
            Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
        Next i
    End If
    'Enviamos el estado de la QUEST
    '0 Disponible
    '1 EN CURSO
    '2 REALIZADA
    '3 no puede hacerla
    Dim PuedeHacerla As Boolean
    'La tiene aceptada el usuario?
    If TieneQuest(UserIndex, QuestIndex) Then
        Call Writer.WriteInt8(1)
    Else
        If UserDoneQuest(UserIndex, QuestIndex) Then
            Call Writer.WriteInt8(2)
        Else
            PuedeHacerla = True
            If QuestList(QuestIndex).RequiredQuest > 0 Then
                If Not UserDoneQuest(UserIndex, QuestList(QuestIndex).RequiredQuest) Then
                    PuedeHacerla = False
                End If
            End If
            If UserList(UserIndex).Stats.ELV < QuestList(QuestIndex).RequiredLevel Then
                PuedeHacerla = False
            End If
            If PuedeHacerla Then
                Call Writer.WriteInt8(0)
            Else
                Call Writer.WriteInt8(3)
            End If
        End If
    End If
    UserList(UserIndex).flags.QuestNumber = QuestIndex
    UserList(UserIndex).flags.QuestItemSlot = Slot
    UserList(UserIndex).flags.QuestOpenByObj = True
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteNpcQuestListSend_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNpcQuestListSend for quest: " & QuestIndex, Erl)
End Sub

Public Sub WriteDebugLogResponse(ByVal UserIndex As Integer, ByVal debugType, ByRef Args() As String, ByVal argc As Integer)
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
            Call Writer.WriteString8("remote DEBUG: " & " user name: " & Args(0))
            With UserList(tIndex)
                Dim timeSinceLastReset As Long
                timeSinceLastReset = CLng(TicksElapsed(Mapping(.ConnectionDetails.ConnID).TimeLastReset, GetTickCountRaw()))
                Call Writer.WriteString8("validConnection: " & .ConnectionDetails.ConnIDValida & " connectionID: " & .ConnectionDetails.ConnID & " UserIndex: " & tIndex & _
                        " charNmae" & .name & " UserLogged state: " & .flags.UserLogged & ", time since last message: " & timeSinceLastReset & " timeout setting: " & _
                        DisconnectTimeout)
            End With
        Else
            Call Writer.WriteInt16(1)
            Call Writer.WriteString8("DEBUG: failed to find user: " & Args(0))
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
End Sub

Public Function PrepareUpdateCharValue(ByVal charindex As Integer, ByVal CharValueType As e_CharValue, ByVal NewValue As Long)
    On Error GoTo PrepareMessageDoAnimation_Err
    Call Writer.WriteInt16(ServerPacketID.eUpdateCharValue)
    Call Writer.WriteInt16(charindex)
    Call Writer.WriteInt16(CharValueType)
    Call Writer.WriteInt32(NewValue)
    Exit Function
PrepareMessageDoAnimation_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageDoAnimation", Erl)
End Function

Public Function PrepareActiveToggles()
    On Error GoTo PrepareActiveToggles_Err
    Call Writer.WriteInt16(ServerPacketID.eSendClientToggles)
    Dim ActiveToggles()   As String
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
End Function

Public Sub WriteAntiCheatMessage(ByVal UserIndex As Integer, ByVal data As Long, ByVal DataSize As Long)
    On Error GoTo WriteAntiCheatMessage_Err
    Dim Buffer() As Byte
    ReDim Buffer(0 To (DataSize - 1)) As Byte
    CopyMemory Buffer(0), ByVal data, DataSize
    Call Writer.WriteInt16(ServerPacketID.eAntiCheatMessage)
    Call Writer.WriteSafeArrayInt8(Buffer)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteAntiCheatMessage_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAntiCheatMessage", Erl)
End Sub

Public Sub WriteAntiCheatStartSeassion(ByVal UserIndex As Integer)
    On Error GoTo WriteAntiStartSeassion_Err
    Call Writer.WriteInt16(ServerPacketID.eAntiCheatStartSession)
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
WriteAntiStartSeassion_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAntiStartSeassion", Erl)
End Sub

Public Sub WriteUpdateLobbyList(ByVal UserIndex As Integer)
    On Error GoTo WriteUpdateLobbyList_Err
    Dim IdList()       As Integer
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
End Sub

Public Sub WriteChangeSkinSlot(ByVal UserIndex As Integer, ByVal TypeSkin As e_OBJType, ByVal Slot As Byte)
    With UserList(UserIndex)
        Call Writer.WriteInt16(ServerPacketID.eChangeSkinSlot)
        Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt16(.Invent_Skins.Object(Slot).ObjIndex)
        'Enviamos si está equipada la skin o no
        Call Writer.WriteBool(.Invent_Skins.Object(Slot).Equipped)
        'Si hay algún error de dateo, no bugeamos el inventario.
        If .Invent_Skins.Object(Slot).ObjIndex > 0 And Not .Invent_Skins.Object(Slot).Deleted Then
            Call Writer.WriteInt32(ObjData(.Invent_Skins.Object(Slot).ObjIndex).GrhIndex)
            Call Writer.WriteInt8(ObjData(.Invent_Skins.Object(Slot).ObjIndex).OBJType)
            Call Writer.WriteString8(ObjData(.Invent_Skins.Object(Slot).ObjIndex).name)
        Else
            Call Writer.WriteInt32(0)
            Call Writer.WriteInt16(0)
            Call Writer.WriteString8(vbNullString)
        End If
        Call modSendData.SendData(SendTarget.ToIndex, UserIndex)
    End With
End Sub
Public Sub WriteGuildConfig(ByVal UserIndex As Integer)
    On Error GoTo WriteGuildConfig_Err

    Call Writer.WriteInt16(ServerPacketID.eGuildConfig)
    Call Writer.WriteInt8(RequiredGuildLevelCallSupport)
    Call Writer.WriteInt8(RequiredGuildLevelSeeInvisible)
    Call Writer.WriteInt8(RequiredGuildLevelSafe)
    Call Writer.WriteInt8(RequiredGuildLevelShowHPBar)
    Call Writer.WriteInt8(MAX_LEVEL_GUILD)
    
    Dim i As Byte
    For i = 1 To MAX_LEVEL_GUILD
        Call Writer.WriteInt8(MembersByLevel(i))
    Next i
    
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub

WriteGuildConfig_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildConfig", Erl)
End Sub
Public Sub WriteShowPickUpObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal amount As Integer)
    On Error GoTo WriteShowPickUpObj_Err

    Call Writer.WriteInt16(ServerPacketID.eShowPickUpObj)
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt16(amount)
    
    Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub

WriteShowPickUpObj_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowPickUpObj", Erl)
End Sub

Public Sub WriteJailCounterToUser(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call Writer.WriteInt16(ServerPacketID.eJailTimeAndPenaltyReason)
        Call Writer.WriteInt32(.Counters.Pena)
        Call Writer.WriteString16(.LastJailPenaltyDescription)
        Call modSendData.SendData(SendTarget.ToIndex, UserIndex)
    End With
End Sub

