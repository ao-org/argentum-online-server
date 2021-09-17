Attribute VB_Name = "Protocol_Writes"
'CSEH: ErrReport
Option Explicit

Private Writer  As Network.Writer

Public Sub InitializeAuxiliaryBuffer()
100     Set Writer = New Network.Writer
End Sub
    
Public Function GetWriterBuffer() As Network.Writer
100     Set GetWriterBuffer = Writer
End Function

' \Begin: [Writes]

Public Sub WriteConnected(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteConnected_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Connected)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteConnected_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteConnected", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Logged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteLoggedMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.logged)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteLoggedMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLoggedMessage", Erl)
        '</EhFooter>
End Sub

Public Sub WriteHora(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteHora_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageHora())
        '<EhFooter>
        Exit Sub

WriteHora_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteHora", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteRemoveAllDialogs_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.RemoveDialogs)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteRemoveAllDialogs_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRemoveAllDialogs", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteRemoveCharDialog_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageRemoveCharDialog( _
                CharIndex))
        '<EhFooter>
        Exit Sub

WriteRemoveCharDialog_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRemoveCharDialog", Erl)
        '</EhFooter>
End Sub

' Writes the "NavigateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteNavigateToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.NavigateToggle)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteNavigateToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNavigateToggle", Erl)
        '</EhFooter>
End Sub

Public Sub WriteNadarToggle(ByVal UserIndex As Integer, _
                            ByVal Puede As Boolean, _
                            Optional ByVal esTrajeCaucho As Boolean = False)
        '<EhHeader>
        On Error GoTo WriteNadarToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.NadarToggle)
102     Call Writer.WriteBool(Puede)
104     Call Writer.WriteBool(esTrajeCaucho)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteNadarToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNadarToggle", Erl)
        '</EhFooter>
End Sub

Public Sub WriteEquiteToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteEquiteToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.EquiteToggle)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteEquiteToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteEquiteToggle", Erl)
        '</EhFooter>
End Sub

Public Sub WriteVelocidadToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteVelocidadToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.VelocidadToggle)
102     Call Writer.WriteReal32(UserList(UserIndex).Char.speeding)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteVelocidadToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteVelocidadToggle", Erl)
        '</EhFooter>
End Sub

Public Sub WriteMacroTrabajoToggle(ByVal UserIndex As Integer, ByVal Activar As Boolean)
        '<EhHeader>
        On Error GoTo WriteMacroTrabajoToggle_Err
        '</EhHeader>

100     If Not Activar Then
102         UserList(UserIndex).flags.TargetObj = 0 ' Sacamos el targer del objeto
104         UserList(UserIndex).flags.UltimoMensaje = 0
106         UserList(UserIndex).Counters.Trabajando = 0
108         UserList(UserIndex).flags.UsandoMacro = False
110         UserList(UserIndex).Trabajo.Target_X = 0
112         UserList(UserIndex).Trabajo.Target_Y = 0
114         UserList(UserIndex).Trabajo.TargetSkill = 0
        Else
116         UserList(UserIndex).flags.UsandoMacro = True
        End If

118     Call Writer.WriteInt(ServerPacketID.MacroTrabajoToggle)
120     Call Writer.WriteBool(Activar)
122     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteMacroTrabajoToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMacroTrabajoToggle", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDisconnect(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteDisconnect_Err
        '</EhHeader>

110     Call Writer.WriteInt(ServerPacketID.Disconnect)
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteDisconnect_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDisconnect", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteCommerceEnd_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CommerceEnd)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteCommerceEnd_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCommerceEnd", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankEnd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBankEnd_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.BankEnd)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteBankEnd_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBankEnd", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteCommerceInit_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CommerceInit)
102     Call Writer.WriteString8(NpcList(UserList(UserIndex).flags.TargetNPC).Name)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteCommerceInit_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCommerceInit", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankInit(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBankInit_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.BankInit)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteBankInit_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBankInit", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserCommerceInit_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UserCommerceInit)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUserCommerceInit_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserCommerceInit", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserCommerceEnd_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UserCommerceEnd)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUserCommerceEnd_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserCommerceEnd", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowBlacksmithForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowBlacksmithForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowBlacksmithForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowBlacksmithForm", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowCarpenterForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowCarpenterForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowCarpenterForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowCarpenterForm", Erl)
        '</EhFooter>
End Sub

Public Sub WriteShowAlquimiaForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowAlquimiaForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowAlquimiaForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowAlquimiaForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowAlquimiaForm", Erl)
        '</EhFooter>
End Sub

Public Sub WriteShowSastreForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowSastreForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowSastreForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowSastreForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowSastreForm", Erl)
        '</EhFooter>
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteNPCKillUser_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.NPCKillUser)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteNPCKillUser_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNPCKillUser", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub Write_BlockedWithShieldUser(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Write_BlockedWithShieldUser_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.BlockedWithShieldUser)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

Write_BlockedWithShieldUser_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.Write_BlockedWithShieldUser", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub Write_BlockedWithShieldOther(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo Write_BlockedWithShieldOther_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.BlockedWithShieldOther)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

Write_BlockedWithShieldOther_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.Write_BlockedWithShieldOther", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CharSwing" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCharSwing(ByVal UserIndex As Integer, _
                          ByVal CharIndex As Integer, _
                          Optional ByVal FX As Boolean = True, _
                          Optional ByVal ShowText As Boolean = True)
        '<EhHeader>
        On Error GoTo WriteCharSwing_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharSwing(CharIndex, _
                FX, ShowText))
        '<EhFooter>
        Exit Sub

WriteCharSwing_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharSwing", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SafeModeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteSafeModeOn_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SafeModeOn)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteSafeModeOn_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSafeModeOn", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteSafeModeOff_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SafeModeOff)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteSafeModeOff_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSafeModeOff", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PartySafeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePartySafeOn(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WritePartySafeOn_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.PartySafeOn)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WritePartySafeOn_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePartySafeOn", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PartySafeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePartySafeOff(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WritePartySafeOff_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.PartySafeOff)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WritePartySafeOff_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePartySafeOff", Erl)
        '</EhFooter>
End Sub

Public Sub WriteClanSeguro(ByVal UserIndex As Integer, ByVal estado As Boolean)
        '<EhHeader>
        On Error GoTo WriteClanSeguro_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ClanSeguro)
102     Call Writer.WriteBool(estado)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteClanSeguro_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteClanSeguro", Erl)
        '</EhFooter>
End Sub

Public Sub WriteSeguroResu(ByVal UserIndex As Integer, ByVal estado As Boolean)
        '<EhHeader>
        On Error GoTo WriteSeguroResu_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SeguroResu)
102     Call Writer.WriteBool(estado)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteSeguroResu_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSeguroResu", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteCantUseWhileMeditating_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CantUseWhileMeditating)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteCantUseWhileMeditating_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCantUseWhileMeditating", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateSta_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateSta)
102     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateSta_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateSta", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateMana_Err
        '</EhHeader>
100     Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
                PrepareMessageCharUpdateMAN(UserIndex))
102     Call Writer.WriteInt(ServerPacketID.UpdateMana)
104     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateMana_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateMana", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
        'Call SendData(SendTarget.ToDiosesYclan, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
        '<EhHeader>
        On Error GoTo WriteUpdateHP_Err
        '</EhHeader>
100     Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
                PrepareMessageCharUpdateHP(UserIndex))
102     Call Writer.WriteInt(ServerPacketID.UpdateHP)
104     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateHP_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateHP", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateGold_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateGold)
102     Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateGold_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateGold", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateExp_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateExp)
102     Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateExp_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateExp", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer)
        '<EhHeader>
        On Error GoTo WriteChangeMap_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.changeMap)
102     Call Writer.WriteInt16(Map)
104     Call Writer.WriteInt16(0)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteChangeMap_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeMap", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePosUpdate(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WritePosUpdate_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.PosUpdate)
102     Call Writer.WriteInt8(UserList(UserIndex).Pos.X)
104     Call Writer.WriteInt8(UserList(UserIndex).Pos.Y)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WritePosUpdate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePosUpdate", Erl)
        '</EhFooter>
End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, _
                           ByVal Target As e_PartesCuerpo, _
                           ByVal damage As Integer)
        '<EhHeader>
        On Error GoTo WriteNPCHitUser_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.NPCHitUser)
102     Call Writer.WriteInt8(Target)
104     Call Writer.WriteInt16(damage)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteNPCHitUser_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNPCHitUser", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserHitNPC(ByVal UserIndex As Integer, ByVal damage As Long)
        '<EhHeader>
        On Error GoTo WriteUserHitNPC_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UserHitNPC)
102     Call Writer.WriteInt32(damage)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUserHitNPC_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserHitNPC", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserAttackedSwing(ByVal UserIndex As Integer, _
                                  ByVal AttackerIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserAttackedSwing_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UserAttackedSwing)
102     Call Writer.WriteInt16(UserList(AttackerIndex).Char.CharIndex)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUserAttackedSwing_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserAttackedSwing", Erl)
        '</EhFooter>
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
                                 ByVal Target As e_PartesCuerpo, _
                                 ByVal attackerChar As Integer, _
                                 ByVal damage As Integer)
        '<EhHeader>
        On Error GoTo WriteUserHittedByUser_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UserHittedByUser)
102     Call Writer.WriteInt16(attackerChar)
104     Call Writer.WriteInt8(Target)
106     Call Writer.WriteInt16(damage)
108     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUserHittedByUser_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserHittedByUser", Erl)
        '</EhFooter>
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
                               ByVal Target As e_PartesCuerpo, _
                               ByVal attackedChar As Integer, _
                               ByVal damage As Integer)
        '<EhHeader>
        On Error GoTo WriteUserHittedUser_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UserHittedUser)
102     Call Writer.WriteInt16(attackedChar)
104     Call Writer.WriteInt8(Target)
106     Call Writer.WriteInt16(damage)
108     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUserHittedUser_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserHittedUser", Erl)
        '</EhFooter>
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
                             ByVal chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal Color As Long)
        '<EhHeader>
        On Error GoTo WriteChatOverHead_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageChatOverHead(chat, _
                CharIndex, Color))
        '<EhFooter>
        Exit Sub

WriteChatOverHead_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChatOverHead", Erl)
        '</EhFooter>
End Sub

Public Sub WriteTextOverChar(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal Color As Long)
        '<EhHeader>
        On Error GoTo WriteTextOverChar_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextOverChar(chat, _
                CharIndex, Color))
        '<EhFooter>
        Exit Sub

WriteTextOverChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTextOverChar", Erl)
        '</EhFooter>
End Sub

Public Sub WriteTextOverTile(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal X As Integer, _
                             ByVal Y As Integer, _
                             ByVal Color As Long)
        '<EhHeader>
        On Error GoTo WriteTextOverTile_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextOverTile(chat, X, _
                Y, Color))
        '<EhFooter>
        Exit Sub

WriteTextOverTile_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTextOverTile", Erl)
        '</EhFooter>
End Sub

Public Sub WriteTextCharDrop(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal Color As Long)
        '<EhHeader>
        On Error GoTo WriteTextCharDrop_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextCharDrop(chat, _
                CharIndex, Color))
        '<EhFooter>
        Exit Sub

WriteTextCharDrop_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTextCharDrop", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, _
                           ByVal chat As String, _
                           ByVal FontIndex As e_FontTypeNames)
        '<EhHeader>
        On Error GoTo WriteConsoleMsg_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageConsoleMsg(chat, _
                FontIndex))
        '<EhFooter>
        Exit Sub

WriteConsoleMsg_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteConsoleMsg", Erl)
        '</EhFooter>
End Sub

Public Sub WriteLocaleMsg(ByVal UserIndex As Integer, _
                          ByVal ID As Integer, _
                          ByVal FontIndex As e_FontTypeNames, _
                          Optional ByVal strExtra As String = vbNullString)
        '<EhHeader>
        On Error GoTo WriteLocaleMsg_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageLocaleMsg(ID, strExtra, _
                FontIndex))
        '<EhFooter>
        Exit Sub

WriteLocaleMsg_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLocaleMsg", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildChat(ByVal UserIndex As Integer, _
                          ByVal chat As String, _
                          ByVal Status As Byte)
        '<EhHeader>
        On Error GoTo WriteGuildChat_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageGuildChat(chat, Status))
        '<EhFooter>
        Exit Sub

WriteGuildChat_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildChat", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String)
        '<EhHeader>
        On Error GoTo WriteShowMessageBox_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowMessageBox)
102     Call Writer.WriteString8(message)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowMessageBox_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowMessageBox", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserIndexInServer_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UserIndexInServer)
102     Call Writer.WriteInt16(UserIndex)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUserIndexInServer_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserIndexInServer", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUserCharIndexInServer_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UserCharIndexInServer)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUserCharIndexInServer_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserCharIndexInServer", Erl)
        '</EhFooter>
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
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As e_Heading, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal Status As Byte, ByVal privileges As Byte, ByVal ParticulaFx As Byte, ByVal Head_Aura As String, ByVal Arma_Aura As String, ByVal Body_Aura As String, ByVal DM_Aura As String, ByVal RM_Aura As String, ByVal Otra_Aura As String, ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal UserMinMAN As Long, ByVal UserMaxMAN As Long, ByVal Simbolo As Byte, Optional ByVal Idle As Boolean = False, Optional ByVal Navegando As Boolean = False, Optional ByVal tipoUsuario As e_TipoUsuario = 0)
        '<EhHeader>
        On Error GoTo WriteCharacterCreate_Err
        '</EhHeader>
100 Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterCreate(Body, Head, _
            Heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, Name, Status, _
            privileges, ParticulaFx, Head_Aura, Arma_Aura, Body_Aura, DM_Aura, RM_Aura, _
            Otra_Aura, Escudo_Aura, speeding, EsNPC, appear, group_index, _
            clan_index, clan_nivel, UserMinHp, UserMaxHp, UserMinMAN, UserMaxMAN, Simbolo, _
            Idle, Navegando, tipoUsuario))
        '<EhFooter>
        Exit Sub

WriteCharacterCreate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterCreate", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, _
                                ByVal CharIndex As Integer, _
                                ByVal Desvanecido As Boolean)
        '<EhHeader>
        On Error GoTo WriteCharacterRemove_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterRemove( _
                CharIndex, Desvanecido))
        '<EhFooter>
        Exit Sub

WriteCharacterRemove_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterRemove", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCharacterMove(ByVal UserIndex As Integer, _
                              ByVal CharIndex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WriteCharacterMove_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterMove( _
                CharIndex, X, Y))
        '<EhFooter>
        Exit Sub

WriteCharacterMove_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterMove", Erl)
        '</EhFooter>
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex As Integer, ByVal Direccion As e_Heading)
        '<EhHeader>
        On Error GoTo WriteForceCharMove_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageForceCharMove(Direccion))
        '<EhFooter>
        Exit Sub

WriteForceCharMove_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteForceCharMove", Erl)
        '</EhFooter>
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
                                ByVal Body As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As e_Heading, _
                                ByVal CharIndex As Integer, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                Optional ByVal Idle As Boolean = False, _
                                Optional ByVal Navegando As Boolean = False)
        '<EhHeader>
        On Error GoTo WriteCharacterChange_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterChange(Body, _
                Head, Heading, CharIndex, weapon, shield, FX, FXLoops, helmet, Idle, _
                Navegando))
        '<EhFooter>
        Exit Sub

WriteCharacterChange_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterChange", Erl)
        '</EhFooter>
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
                             ByVal ObjIndex As Integer, _
                             ByVal amount As Integer, _
                             ByVal X As Byte, _
                             ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WriteObjectCreate_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageObjectCreate(ObjIndex, _
                amount, X, Y))
        '<EhFooter>
        Exit Sub

WriteObjectCreate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteObjectCreate", Erl)
        '</EhFooter>
End Sub

Public Sub WriteParticleFloorCreate(ByVal UserIndex As Integer, _
                                    ByVal Particula As Integer, _
                                    ByVal ParticulaTime As Integer, _
                                    ByVal Map As Integer, _
                                    ByVal X As Byte, _
                                    ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WriteParticleFloorCreate_Err
        '</EhHeader>

100     If Particula = 0 Then
102         Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageParticleFXToFloor( _
                    X, Y, Particula, ParticulaTime))
        End If

        '<EhFooter>
        Exit Sub

WriteParticleFloorCreate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteParticleFloorCreate", Erl)
        '</EhFooter>
End Sub

Public Sub WriteLightFloorCreate(ByVal UserIndex As Integer, _
                                 ByVal LuzColor As Long, _
                                 ByVal Rango As Byte, _
                                 ByVal Map As Integer, _
                                 ByVal X As Byte, _
                                 ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WriteLightFloorCreate_Err
        '</EhHeader>
100     MapData(Map, X, Y).Luz.Color = LuzColor
102     MapData(Map, X, Y).Luz.Rango = Rango

104     If Rango = 0 Then
106         Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageLightFXToFloor(X, _
                    Y, LuzColor, Rango))
        End If

        '<EhFooter>
        Exit Sub

WriteLightFloorCreate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLightFloorCreate", Erl)
        '</EhFooter>
End Sub

Public Sub WriteFxPiso(ByVal UserIndex As Integer, _
                       ByVal GrhIndex As Integer, _
                       ByVal X As Byte, _
                       ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WriteFxPiso_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageFxPiso(GrhIndex, X, Y))
        '<EhFooter>
        Exit Sub

WriteFxPiso_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFxPiso", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WriteObjectDelete_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageObjectDelete(X, Y))
        '<EhFooter>
        Exit Sub

WriteObjectDelete_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteObjectDelete", Erl)
        '</EhFooter>
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
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByVal Blocked As Byte)
        '<EhHeader>
        On Error GoTo Write_BlockPosition_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.BlockPosition)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt8(Blocked)
108     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

Write_BlockPosition_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.Write_BlockPosition", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePlayMidi(ByVal UserIndex As Integer, _
                         ByVal midi As Byte, _
                         Optional ByVal loops As Integer = -1)
        '<EhHeader>
        On Error GoTo WritePlayMidi_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePlayMidi(midi, loops))
        '<EhFooter>
        Exit Sub

WritePlayMidi_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePlayMidi", Erl)
        '</EhFooter>
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
                         ByVal X As Byte, _
                         ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo WritePlayWave_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePlayWave(wave, X, Y))
        '<EhFooter>
        Exit Sub

WritePlayWave_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePlayWave", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)
        '<EhHeader>
        On Error GoTo WriteGuildList_Err
        '</EhHeader>

        Dim Tmp As String

        Dim i   As Long

100     Call Writer.WriteInt(ServerPacketID.guildList)

        ' Prepare guild name's list
102     For i = LBound(guildList()) To UBound(guildList())
104         Tmp = Tmp & guildList(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteGuildList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildList", Erl)
        '</EhFooter>
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAreaChanged(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteAreaChanged_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.AreaChanged)
102     Call Writer.WriteInt8(UserList(UserIndex).Pos.X)
104     Call Writer.WriteInt8(UserList(UserIndex).Pos.Y)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteAreaChanged_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAreaChanged", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePauseToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WritePauseToggle_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePauseToggle())
        '<EhFooter>
        Exit Sub

WritePauseToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePauseToggle", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRainToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteRainToggle_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageRainToggle())
        '<EhFooter>
        Exit Sub

WriteRainToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRainToggle", Erl)
        '</EhFooter>
End Sub

Public Sub WriteNubesToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteNubesToggle_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageNieblandoToggle( _
                IntensidadDeNubes))
        '<EhFooter>
        Exit Sub

WriteNubesToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNubesToggle", Erl)
        '</EhFooter>
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
                         ByVal CharIndex As Integer, _
                         ByVal FX As Integer, _
                         ByVal FXLoops As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "CreateFX" message to the given user's outgoing data buffer
        '***************************************************
        '<EhHeader>
        On Error GoTo WriteCreateFX_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCreateFX(CharIndex, FX, _
                FXLoops))
        '<EhFooter>
        Exit Sub

WriteCreateFX_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCreateFX", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
        '***************************************************
        '<EhHeader>
        On Error GoTo WriteUpdateUserStats_Err
        '</EhHeader>
100     Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
                PrepareMessageCharUpdateHP(UserIndex))
102     Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
                PrepareMessageCharUpdateMAN(UserIndex))
104     Call Writer.WriteInt(ServerPacketID.UpdateUserStats)
106     Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxHp)
108     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
110     Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxMAN)
112     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
114     Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxSta)
116     Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
118     Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
120     Call Writer.WriteInt8(UserList(UserIndex).Stats.ELV)
122     Call Writer.WriteInt32(ExpLevelUp(UserList(UserIndex).Stats.ELV))
124     Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
126     Call Writer.WriteInt8(UserList(UserIndex).clase)
128     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateUserStats_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateUserStats", Erl)
        '</EhFooter>
End Sub

Public Sub WriteUpdateUserKey(ByVal UserIndex As Integer, _
                              ByVal Slot As Integer, _
                              ByVal Llave As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateUserKey_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateUserKey)
102     Call Writer.WriteInt16(Slot)
104     Call Writer.WriteInt16(Llave)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateUserKey_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateUserKey", Erl)
        '</EhFooter>
End Sub

' Actualiza el indicador de daño mágico
Public Sub WriteUpdateDM(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateDM_Err
        '</EhHeader>

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

110         Call Writer.WriteInt(ServerPacketID.UpdateDM)
112         Call Writer.WriteInt16(Valor)
        End With

114     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateDM_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateDM", Erl)
        '</EhFooter>
End Sub

' Actualiza el indicador de resistencia mágica
Public Sub WriteUpdateRM(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateRM_Err
        '</EhHeader>

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
120         Call Writer.WriteInt(ServerPacketID.UpdateRM)
122         Call Writer.WriteInt16(Valor)
        End With

124     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateRM_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateRM", Erl)
        '</EhFooter>
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As e_Skill)
        '<EhHeader>
        On Error GoTo WriteWorkRequestTarget_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.WorkRequestTarget)
102     Call Writer.WriteInt8(Skill)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteWorkRequestTarget_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteWorkRequestTarget", Erl)
        '</EhFooter>
End Sub

' Writes the "InventoryUnlockSlots" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteInventoryUnlockSlots(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteInventoryUnlockSlots_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.InventoryUnlockSlots)
102     Call Writer.WriteInt8(UserList(UserIndex).Stats.InventLevel)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteInventoryUnlockSlots_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteInventoryUnlockSlots", Erl)
        '</EhFooter>
End Sub

Public Sub WriteIntervals(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteIntervals_Err
        '</EhHeader>

100     With UserList(UserIndex)
102         Call Writer.WriteInt(ServerPacketID.Intervals)
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
        '<EhFooter>
        Exit Sub

WriteIntervals_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteIntervals", Erl)
        '</EhFooter>
End Sub

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteChangeInventorySlot_Err
        '</EhHeader>

        Dim ObjIndex    As Integer

        Dim PodraUsarlo As Byte

100     Call Writer.WriteInt(ServerPacketID.ChangeInventorySlot)
102     Call Writer.WriteInt8(Slot)
104     ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex

106     If ObjIndex > 0 Then
108         PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)
        End If

110     Call Writer.WriteInt16(ObjIndex)
112     Call Writer.WriteInt16(UserList(UserIndex).Invent.Object(Slot).amount)
114     Call Writer.WriteBool(UserList(UserIndex).Invent.Object(Slot).Equipped)
116     Call Writer.WriteReal32(SalePrice(ObjIndex))
118     Call Writer.WriteInt8(PodraUsarlo)
120     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteChangeInventorySlot_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeInventorySlot", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteChangeBankSlot_Err
        '</EhHeader>

        Dim ObjIndex    As Integer

        Dim Valor       As Long

        Dim PodraUsarlo As Byte

100     Call Writer.WriteInt(ServerPacketID.ChangeBankSlot)
102     Call Writer.WriteInt8(Slot)
104     ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
106     Call Writer.WriteInt16(ObjIndex)

108     If ObjIndex > 0 Then
110         Valor = ObjData(ObjIndex).Valor
112         PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)
        End If

114     Call Writer.WriteInt16(UserList(UserIndex).BancoInvent.Object(Slot).amount)
116     Call Writer.WriteInt32(Valor)
118     Call Writer.WriteInt8(PodraUsarlo)
120     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteChangeBankSlot_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeBankSlot", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
        '<EhHeader>
        On Error GoTo WriteChangeSpellSlot_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ChangeSpellSlot)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(UserList(UserIndex).Stats.UserHechizos(Slot))

106     If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
108         Call Writer.WriteInt8(UserList(UserIndex).Stats.UserHechizos(Slot))
        Else
110         Call Writer.WriteInt8(255)
        End If

112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteChangeSpellSlot_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeSpellSlot", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAttributes(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteAttributes_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Atributes)
102     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Fuerza))
104     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Agilidad))
106     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
                e_Atributos.Inteligencia))
108     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
                e_Atributos.Constitucion))
110     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Carisma))
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteAttributes_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAttributes", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBlacksmithWeapons_Err
        '</EhHeader>

        Dim i              As Long

        Dim obj            As t_ObjData

        Dim validIndexes() As Integer

        Dim Count          As Integer

100     ReDim validIndexes(1 To UBound(ArmasHerrero()))
102     Call Writer.WriteInt(ServerPacketID.BlacksmithWeapons)

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
118         obj = ObjData(ArmasHerrero(validIndexes(i)))
            'Call Writer.WriteString8(obj.Index)
120         Call Writer.WriteInt16(ArmasHerrero(validIndexes(i)))
122         Call Writer.WriteInt16(obj.LingH)
124         Call Writer.WriteInt16(obj.LingP)
126         Call Writer.WriteInt16(obj.LingO)
128     Next i

130     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteBlacksmithWeapons_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlacksmithWeapons", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBlacksmithArmors_Err
        '</EhHeader>

        Dim i              As Long

        Dim obj            As t_ObjData

        Dim validIndexes() As Integer

        Dim Count          As Integer

100     ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
102     Call Writer.WriteInt(ServerPacketID.BlacksmithArmors)

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
118         obj = ObjData(ArmadurasHerrero(validIndexes(i)))
120         Call Writer.WriteString8(obj.Name)
122         Call Writer.WriteInt16(obj.LingH)
124         Call Writer.WriteInt16(obj.LingP)
126         Call Writer.WriteInt16(obj.LingO)
128         Call Writer.WriteInt16(ArmadurasHerrero(validIndexes(i)))
130     Next i

132     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteBlacksmithArmors_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlacksmithArmors", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteCarpenterObjects_Err
        '</EhHeader>

        Dim i              As Long

        Dim validIndexes() As Integer

        Dim Count          As Byte

100     ReDim validIndexes(1 To UBound(ObjCarpintero()))
102     Call Writer.WriteInt(ServerPacketID.CarpenterObjects)

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
        '<EhFooter>
        Exit Sub

WriteCarpenterObjects_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCarpenterObjects", Erl)
        '</EhFooter>
End Sub

Public Sub WriteAlquimistaObjects(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteAlquimistaObjects_Err
        '</EhHeader>

        Dim i              As Long

        Dim validIndexes() As Integer

        Dim Count          As Integer

100     ReDim validIndexes(1 To UBound(ObjAlquimista()))
102     Call Writer.WriteInt(ServerPacketID.AlquimistaObj)

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
        '<EhFooter>
        Exit Sub

WriteAlquimistaObjects_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAlquimistaObjects", Erl)
        '</EhFooter>
End Sub

Public Sub WriteSastreObjects(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteSastreObjects_Err
        '</EhHeader>

        Dim i              As Long

        Dim validIndexes() As Integer

        Dim Count          As Integer

100     ReDim validIndexes(1 To UBound(ObjSastre()))
102     Call Writer.WriteInt(ServerPacketID.SastreObj)

104     For i = 1 To UBound(ObjSastre())

            ' Can the user create this object? If so add it to the list....
106         If ObjData(ObjSastre(i)).SkMAGOria <= UserList(UserIndex).Stats.UserSkills( _
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
        '<EhFooter>
        Exit Sub

WriteSastreObjects_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSastreObjects", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRestOK(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteRestOK_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.RestOK)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteRestOK_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteRestOK", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal message As String)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "ErrorMsg" message to the given user's outgoing data buffer
        '***************************************************
        '<EhHeader>
        On Error GoTo WriteErrorMsg_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageErrorMsg(message))
        '<EhFooter>
        Exit Sub

WriteErrorMsg_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteErrorMsg", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Blind" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlind(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBlind_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Blind)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteBlind_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlind", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDumb(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteDumb_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Dumb)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteDumb_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDumb", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
'Optimizacion de protocolo por Ladder
Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowSignal_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowSignal)
102     Call Writer.WriteInt16(ObjIndex)
104     Call Writer.WriteInt16(ObjData(ObjIndex).GrhSecundario)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowSignal_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowSignal", Erl)
        '</EhFooter>
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
                                       ByVal Slot As Byte, _
                                       ByRef obj As t_Obj, _
                                       ByVal price As Single)
        '<EhHeader>
        On Error GoTo WriteChangeNPCInventorySlot_Err
        '</EhHeader>

        Dim PodraUsarlo As Byte

100     If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
102         PodraUsarlo = PuedeUsarObjeto(UserIndex, obj.ObjIndex)
        End If

104     Call Writer.WriteInt(ServerPacketID.ChangeNPCInventorySlot)
106     Call Writer.WriteInt8(Slot)
108     Call Writer.WriteInt16(obj.ObjIndex)
110     Call Writer.WriteInt16(obj.amount)
112     Call Writer.WriteReal32(price)
114     Call Writer.WriteInt8(PodraUsarlo)
116     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteChangeNPCInventorySlot_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeNPCInventorySlot", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteUpdateHungerAndThirst_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateHungerAndThirst)
102     Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxAGU)
104     Call Writer.WriteInt8(UserList(UserIndex).Stats.MinAGU)
106     Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxHam)
108     Call Writer.WriteInt8(UserList(UserIndex).Stats.MinHam)
110     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateHungerAndThirst_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateHungerAndThirst", Erl)
        '</EhFooter>
End Sub

Public Sub WriteLight(ByVal UserIndex As Integer, ByVal Map As Integer)
        '<EhHeader>
        On Error GoTo WriteLight_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.light)
102     Call Writer.WriteString8(MapInfo(Map).base_light)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteLight_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLight", Erl)
        '</EhFooter>
End Sub

Public Sub WriteFlashScreen(ByVal UserIndex As Integer, _
                            ByVal Color As Long, _
                            ByVal Time As Long, _
                            Optional ByVal Ignorar As Boolean = False)
        '<EhHeader>
        On Error GoTo WriteFlashScreen_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.FlashScreen)
102     Call Writer.WriteInt32(Color)
104     Call Writer.WriteInt32(Time)
106     Call Writer.WriteBool(Ignorar)
108     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteFlashScreen_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFlashScreen", Erl)
        '</EhFooter>
End Sub

Public Sub WriteFYA(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteFYA_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.FYA)
102     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(1))
104     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(2))
106     Call Writer.WriteInt16(UserList(UserIndex).flags.DuracionEfecto)
108     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteFYA_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFYA", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCerrarleCliente(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteCerrarleCliente_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CerrarleCliente)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteCerrarleCliente_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCerrarleCliente", Erl)
        '</EhFooter>
End Sub

Public Sub WriteOxigeno(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteOxigeno_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Oxigeno)
102     Call Writer.WriteInt16(UserList(UserIndex).Counters.Oxigeno)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteOxigeno_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteOxigeno", Erl)
        '</EhFooter>
End Sub

Public Sub WriteContadores(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteContadores_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Contadores)
102     Call Writer.WriteInt16(UserList(UserIndex).Counters.Invisibilidad)

104     If UserList(UserIndex).flags.NecesitaOxigeno Then
106         Call Writer.WriteInt16(UserList(UserIndex).Counters.Oxigeno)
        Else
108         Call Writer.WriteInt16(0)
        End If

110     Call Writer.WriteInt16(UserList(UserIndex).flags.DuracionEfecto)
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteContadores_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteContadores", Erl)
        '</EhFooter>
End Sub

Public Sub WriteBindKeys(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBindKeys_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.BindKeys)
102     Call Writer.WriteInt8(UserList(UserIndex).ChatCombate)
104     Call Writer.WriteInt8(UserList(UserIndex).ChatGlobal)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteBindKeys_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBindKeys", Erl)
        '</EhFooter>
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMiniStats(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteMiniStats_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.MiniStats)
102     Call Writer.WriteInt32(UserList(UserIndex).Faccion.ciudadanosMatados)
104     Call Writer.WriteInt32(UserList(UserIndex).Faccion.CriminalesMatados)
106     Call Writer.WriteInt8(UserList(UserIndex).Faccion.Status)
108     Call Writer.WriteInt16(UserList(UserIndex).Stats.NPCsMuertos)
110     Call Writer.WriteInt8(UserList(UserIndex).clase)
112     Call Writer.WriteInt32(UserList(UserIndex).Counters.Pena)
114     Call Writer.WriteInt32(UserList(UserIndex).flags.VecesQueMoriste)
116     Call Writer.WriteInt8(UserList(UserIndex).genero)
118     Call Writer.WriteInt8(UserList(UserIndex).raza)
120     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteMiniStats_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMiniStats", Erl)
        '</EhFooter>
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data .incomingData.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
        '<EhHeader>
        On Error GoTo WriteLevelUp_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.LevelUp)
102     Call Writer.WriteInt16(skillPoints)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteLevelUp_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteLevelUp", Erl)
        '</EhFooter>
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data .incomingData.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, _
                            ByVal title As String, _
                            ByVal message As String)
        '<EhHeader>
        On Error GoTo WriteAddForumMsg_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.AddForumMsg)
102     Call Writer.WriteString8(title)
104     Call Writer.WriteString8(message)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteAddForumMsg_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAddForumMsg", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowForumForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowForumForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowForumForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowForumForm", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSetInvisible(ByVal UserIndex As Integer, _
                             ByVal CharIndex As Integer, _
                             ByVal invisible As Boolean)
        '<EhHeader>
        On Error GoTo WriteSetInvisible_Err
        '</EhHeader>
100     Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageSetInvisible(CharIndex, _
                invisible))
        '<EhFooter>
        Exit Sub

WriteSetInvisible_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSetInvisible", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.
''
' Writes the "DiceRoll" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDiceRoll(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteDiceRoll_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.DiceRoll)
        ' TODO: SACAR ESTE PAQUETE USAR EL DE ATRIBUTOS
102     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Fuerza))
104     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Agilidad))
106     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
                e_Atributos.Inteligencia))
108     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
                e_Atributos.Constitucion))
110     Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(e_Atributos.Carisma))
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteDiceRoll_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDiceRoll", Erl)
        '</EhFooter>
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteMeditateToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.MeditateToggle)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteMeditateToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteMeditateToggle", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteBlindNoMore_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.BlindNoMore)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteBlindNoMore_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteBlindNoMore", Erl)
        '</EhFooter>
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteDumbNoMore_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.DumbNoMore)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteDumbNoMore_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDumbNoMore", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSendSkills(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteSendSkills_Err
        '</EhHeader>

        Dim i As Long

100     Call Writer.WriteInt(ServerPacketID.SendSkills)

102     For i = 1 To NUMSKILLS
104         Call Writer.WriteInt8(UserList(UserIndex).Stats.UserSkills(i))
106     Next i

108     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteSendSkills_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSendSkills", Erl)
        '</EhFooter>
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteTrainerCreatureList_Err
        '</EhHeader>

        Dim i   As Long

        Dim str As String

100     Call Writer.WriteInt(ServerPacketID.TrainerCreatureList)

102     For i = 1 To NpcList(NpcIndex).NroCriaturas
104         str = str & NpcList(NpcIndex).Criaturas(i).NpcName & SEPARATOR
106     Next i

108     If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
110     Call Writer.WriteString8(str)
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteTrainerCreatureList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteTrainerCreatureList", Erl)
        '</EhFooter>
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
                          ByRef MemberList() As String, _
                          ByVal ClanNivel As Byte, _
                          ByVal ExpAcu As Integer, _
                          ByVal ExpNe As Integer)
        '<EhHeader>
        On Error GoTo WriteGuildNews_Err
        '</EhHeader>

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt(ServerPacketID.guildNews)
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
118         Tmp = Tmp & MemberList(i) & SEPARATOR
120     Next i

122     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
124     Call Writer.WriteString8(Tmp)
126     Call Writer.WriteInt8(ClanNivel)
128     Call Writer.WriteInt16(ExpAcu)
130     Call Writer.WriteInt16(ExpNe)
132     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteGuildNews_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildNews", Erl)
        '</EhFooter>
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)
        '<EhHeader>
        On Error GoTo WriteOfferDetails_Err
        '</EhHeader>

100     Call Writer.WriteInt(ServerPacketID.OfferDetails)
102     Call Writer.WriteString8(details)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteOfferDetails_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteOfferDetails", Erl)
        '</EhFooter>
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
        '<EhHeader>
        On Error GoTo WriteAlianceProposalsList_Err
        '</EhHeader>

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt(ServerPacketID.AlianceProposalsList)

        ' Prepare guild's list
102     For i = LBound(guilds()) To UBound(guilds())
104         Tmp = Tmp & guilds(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteAlianceProposalsList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteAlianceProposalsList", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)
        '<EhHeader>
        On Error GoTo WritePeaceProposalsList_Err
        '</EhHeader>

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt(ServerPacketID.PeaceProposalsList)

        ' Prepare guilds' list
102     For i = LBound(guilds()) To UBound(guilds())
104         Tmp = Tmp & guilds(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WritePeaceProposalsList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePeaceProposalsList", Erl)
        '</EhFooter>
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
        ByVal race As e_Raza, ByVal Class As e_Class, ByVal gender As e_Genero, ByVal _
        level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal previousPetitions As String, _
        ByVal currentGuild As String, ByVal previousGuilds As String, ByVal _
        RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As _
        Long, ByVal criminalsKilled As Long)
        '<EhHeader>
        On Error GoTo WriteCharacterInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CharacterInfo)
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
        '<EhFooter>
        Exit Sub

WriteCharacterInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCharacterInfo", Erl)
        '</EhFooter>
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
                                ByRef MemberList() As String, _
                                ByVal guildNews As String, _
                                ByRef joinRequests() As String, _
                                ByVal NivelDeClan As Byte, _
                                ByVal ExpActual As Integer, _
                                ByVal ExpNecesaria As Integer)
        '<EhHeader>
        On Error GoTo WriteGuildLeaderInfo_Err
        '</EhHeader>

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt(ServerPacketID.GuildLeaderInfo)

        ' Prepare guild name's list
102     For i = LBound(guildList()) To UBound(guildList())
104         Tmp = Tmp & guildList(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
        ' Prepare guild member's list
112     Tmp = vbNullString

114     For i = LBound(MemberList()) To UBound(MemberList())
116         Tmp = Tmp & MemberList(i) & SEPARATOR
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
        '<EhFooter>
        Exit Sub

WriteGuildLeaderInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildLeaderInfo", Erl)
        '</EhFooter>
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
        '<EhHeader>
        On Error GoTo WriteGuildDetails_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.GuildDetails)
102     Call Writer.WriteString8(GuildName)
104     Call Writer.WriteString8(founder)
106     Call Writer.WriteString8(foundationDate)
108     Call Writer.WriteString8(leader)
110     Call Writer.WriteInt16(memberCount)
112     Call Writer.WriteString8(alignment)
114     Call Writer.WriteString8(guildDesc)
116     Call Writer.WriteInt8(NivelDeClan)
118     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteGuildDetails_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuildDetails", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowGuildFundationForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowGuildFundationForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowGuildFundationForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowGuildFundationForm", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteParalizeOK(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteParalizeOK_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ParalizeOK)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteParalizeOK_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteParalizeOK", Erl)
        '</EhFooter>
End Sub

Public Sub WriteInmovilizaOK(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteInmovilizaOK_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.InmovilizadoOK)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteInmovilizaOK_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteInmovilizaOK", Erl)
        '</EhFooter>
End Sub

Public Sub WriteStopped(ByVal UserIndex As Integer, ByVal Stopped As Boolean)
        '<EhHeader>
        On Error GoTo WriteStopped_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Stopped)
102     Call Writer.WriteBool(Stopped)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteStopped_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteStopped", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
        '<EhHeader>
        On Error GoTo WriteShowUserRequest_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowUserRequest)
102     Call Writer.WriteString8(details)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowUserRequest_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowUserRequest", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    Amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, _
                                    ByRef itemsAenviar() As t_Obj, _
                                    ByVal gold As Long, _
                                    ByVal miOferta As Boolean)
        '<EhHeader>
        On Error GoTo WriteChangeUserTradeSlot_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ChangeUserTradeSlot)
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
124     Next i

126     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteChangeUserTradeSlot_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteChangeUserTradeSlot", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByVal ListaCompleta As Boolean)
        '<EhHeader>
        On Error GoTo WriteSpawnList_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SpawnListt)
102     Call Writer.WriteBool(ListaCompleta)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteSpawnList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteSpawnList", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowSOSForm_Err
        '</EhHeader>

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt(ServerPacketID.ShowSOSForm)

102     For i = 1 To Ayuda.Longitud
104         Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
106     Next i

108     If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowSOSForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowSOSForm", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, _
                                    ByVal currentMOTD As String)
        '<EhHeader>
        On Error GoTo WriteShowMOTDEditionForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowMOTDEditionForm)
102     Call Writer.WriteString8(currentMOTD)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowMOTDEditionForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowMOTDEditionForm", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowGMPanelForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowGMPanelForm)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.Head)
104     Call Writer.WriteInt16(UserList(UserIndex).Char.Body)
106     Call Writer.WriteInt16(UserList(UserIndex).Char.CascoAnim)
108     Call Writer.WriteInt16(UserList(UserIndex).Char.WeaponAnim)
110     Call Writer.WriteInt16(UserList(UserIndex).Char.ShieldAnim)
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowGMPanelForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowGMPanelForm", Erl)
        '</EhFooter>
End Sub

Public Sub WriteShowFundarClanForm(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowFundarClanForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowFundarClanForm)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowFundarClanForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowFundarClanForm", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserNameList(ByVal UserIndex As Integer, _
                             ByRef userNamesList() As String, _
                             ByVal cant As Integer)
        '<EhHeader>
        On Error GoTo WriteUserNameList_Err
        '</EhHeader>

        Dim i   As Long

        Dim Tmp As String

100     Call Writer.WriteInt(ServerPacketID.UserNameList)

        ' Prepare user's names list
102     For i = 1 To cant
104         Tmp = Tmp & userNamesList(i) & SEPARATOR
106     Next i

108     If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
110     Call Writer.WriteString8(Tmp)
112     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUserNameList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUserNameList", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Pong" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePong(ByVal UserIndex As Integer, ByVal Time As Long)
        '<EhHeader>
        On Error GoTo WritePong_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Pong)
102     Call Writer.WriteInt32(Time)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WritePong_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePong", Erl)
        '</EhFooter>
End Sub

Public Sub WritePersonajesDeCuenta(ByVal UserIndex As Integer, Personajes As Object)
        '<EhHeader>
        On Error GoTo WritePersonajesDeCuenta_Err
        '</EhHeader>

134     Call Writer.WriteInt(ServerPacketID.PersonajesDeCuenta)
136     Call Writer.WriteInt8(Personajes.Count)

        Dim Status As Byte, Cabeza As Integer

        Dim Personaje As Variant

138     For Each Personaje In Personajes
            Status = Personaje.Item("status")
            If EsRolesMaster(Personaje.Item("name")) Then
                Status = 3
            ElseIf EsConsejero(Personaje.Item("name")) Then
                Status = 4
            ElseIf EsSemiDios(Personaje.Item("name")) Then
                Status = 5
            ElseIf EsDios(Personaje.Item("name")) Then
                Status = 6
            ElseIf EsAdmin(Personaje.Item("name")) Then
                Status = 7
            End If
            
            Cabeza = Personaje.Item("head_id")
            If Personaje.Item("min_hp") = 0 Or Personaje.Item("is_sailing") Then
                Cabeza = 0
            End If

140         Call Writer.WriteString8(Personaje.Item("name"))
142         Call Writer.WriteInt8(Personaje.Item("level"))
144         Call Writer.WriteInt16(Personaje.Item("pos_map"))
146         Call Writer.WriteInt16(Personaje.Item("pos_x"))
148         Call Writer.WriteInt16(Personaje.Item("pos_y"))
150         Call Writer.WriteInt16(Personaje.Item("body_id"))
152         Call Writer.WriteInt16(Cabeza)
154         Call Writer.WriteInt8(Status)
156         Call Writer.WriteInt8(Personaje.Item("class_id"))
158         Call Writer.WriteInt16(Personaje.Item("helmet_id"))
160         Call Writer.WriteInt16(Personaje.Item("shield_id"))
162         Call Writer.WriteInt16(Personaje.Item("weapon_id"))
164         Call Writer.WriteString8(modGuilds.GuildName(Personaje.Item("guild_index")))
166     Next

168     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WritePersonajesDeCuenta_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePersonajesDeCuenta", Erl)
        '</EhFooter>
End Sub

Public Sub WriteGoliathInit(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteGoliathInit_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Goliath)
102     Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
104     Call Writer.WriteInt8(UserList(UserIndex).BancoInvent.NroItems)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteGoliathInit_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGoliathInit", Erl)
        '</EhFooter>
End Sub

Public Sub WriteShowFrmLogear(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowFrmLogear_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowFrmLogear)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowFrmLogear_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowFrmLogear", Erl)
        '</EhFooter>
End Sub

Public Sub WriteShowFrmMapa(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteShowFrmMapa_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowFrmMapa)
102     Call Writer.WriteInt16(ExpMult)
104     Call Writer.WriteInt16(OroMult)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteShowFrmMapa_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteShowFrmMapa", Erl)
        '</EhFooter>
End Sub

Public Sub WriteFamiliar(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteFamiliar_Err
        '</EhHeader>

100     With UserList(UserIndex)
102         Call Writer.WriteInt(ServerPacketID.Familiar)
104         Call Writer.WriteInt8(.Familiar.Existe)
106         Call Writer.WriteInt8(.Familiar.Muerto)
108         Call Writer.WriteString8(.Familiar.nombre)
110         Call Writer.WriteInt32(.Familiar.Exp)
112         Call Writer.WriteInt32(.Familiar.ELU)
114         Call Writer.WriteInt8(.Familiar.nivel)
116         Call Writer.WriteInt16(.Familiar.MinHp)
118         Call Writer.WriteInt16(.Familiar.MaxHp)
120         Call Writer.WriteInt16(.Familiar.MinHIT)
122         Call Writer.WriteInt16(.Familiar.MaxHit)
        End With

124     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteFamiliar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteFamiliar", Erl)
        '</EhFooter>
End Sub

Public Sub WritePreguntaBox(ByVal UserIndex As Integer, ByVal message As String)
        '<EhHeader>
        On Error GoTo WritePreguntaBox_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowPregunta)
102     Call Writer.WriteString8(message)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WritePreguntaBox_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WritePreguntaBox", Erl)
        '</EhFooter>
End Sub

Public Sub WriteDatosGrupo(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteDatosGrupo_Err
        '</EhHeader>

        Dim i As Byte

100     With UserList(UserIndex)
102         Call Writer.WriteInt(ServerPacketID.DatosGrupo)
104         Call Writer.WriteBool(.Grupo.EnGrupo)

106         If .Grupo.EnGrupo = True Then
108             Call Writer.WriteInt8(UserList(.Grupo.Lider).Grupo.CantidadMiembros)

                'Call Writer.WriteInt8(UserList(.Grupo.Lider).name)
110             If .Grupo.Lider = UserIndex Then

112                 For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros

114                     If i = 1 Then
116                         Call Writer.WriteString8(UserList(.Grupo.Miembros(i)).Name & _
                                    "(Líder)")
                        Else
118                         Call Writer.WriteString8(UserList(.Grupo.Miembros(i)).Name)
                        End If

120                 Next i

                Else

122                 For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros

124                     If i = 1 Then
126                         Call Writer.WriteString8(UserList(UserList( _
                                    .Grupo.Lider).Grupo.Miembros(i)).Name & "(Líder)")
                        Else
128                         Call Writer.WriteString8(UserList(UserList( _
                                    .Grupo.Lider).Grupo.Miembros(i)).Name)
                        End If

130                 Next i

                End If
            End If

        End With

132     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteDatosGrupo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteDatosGrupo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteUbicacion(ByVal UserIndex As Integer, _
                          ByVal Miembro As Byte, _
                          ByVal GPS As Integer)
        '<EhHeader>
        On Error GoTo WriteUbicacion_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ubicacion)
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
        '<EhFooter>
        Exit Sub

WriteUbicacion_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUbicacion", Erl)
        '</EhFooter>
End Sub

Public Sub WriteViajarForm(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteViajarForm_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ViajarForm)

        Dim destinos As Byte

        Dim i        As Byte

102     destinos = NpcList(NpcIndex).NumDestinos
104     Call Writer.WriteInt8(destinos)

106     For i = 1 To destinos
108         Call Writer.WriteString8(NpcList(NpcIndex).Dest(i))
110     Next i

112     Call Writer.WriteInt8(NpcList(NpcIndex).Interface)
114     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteViajarForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteViajarForm", Erl)
        '</EhFooter>
End Sub

Public Sub WriteQuestDetails(ByVal UserIndex As Integer, _
                             ByVal QuestIndex As Integer, _
                             Optional QuestSlot As Byte = 0)
        '<EhHeader>
        On Error GoTo WriteQuestDetails_Err
        '</EhHeader>

        Dim i As Integer

        'ID del paquete
100     Call Writer.WriteInt(ServerPacketID.QuestDetails)
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
136             Call Writer.WriteInt16(CantidadObjEnInv(UserIndex, QuestList( _
                        QuestIndex).RequiredOBJ(i).ObjIndex))
                ' Call Writer.WriteInt16(0)
138         Next i

        End If

        'Enviamos la recompensa de oro y experiencia.
140     Call Writer.WriteInt32((QuestList(QuestIndex).RewardGLD * OroMult))
142     Call Writer.WriteInt32((QuestList(QuestIndex).RewardEXP * ExpMult))
        'Enviamos la cantidad de objs de recompensa
144     Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)

146     If QuestList(QuestIndex).RewardOBJs Then

            'si hay objs entonces enviamos la lista
148         For i = 1 To QuestList(QuestIndex).RewardOBJs
150             Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).amount)
152             Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
154         Next i

        End If

156     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteQuestDetails_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteQuestDetails", Erl)
        '</EhFooter>
End Sub
 
Public Sub WriteQuestListSend(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteQuestListSend_Err
        '</EhHeader>

        Dim i       As Integer

        Dim tmpStr  As String

        Dim tmpByte As Byte

100     With UserList(UserIndex)
102         Call Writer.WriteInt(ServerPacketID.QuestListSend)

104         For i = 1 To MAXUSERQUESTS

106             If .QuestStats.Quests(i).QuestIndex Then
108                 tmpByte = tmpByte + 1
110                 tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).nombre & "-"
                End If

112         Next i

            'Escribimos la cantidad de quests
114         Call Writer.WriteInt8(tmpByte)

            'Escribimos la lista de quests (sacamos el último caracter)
116         If tmpByte Then
118             Call Writer.WriteString8(Left$(tmpStr, Len(tmpStr) - 1))
            End If

        End With

120     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteQuestListSend_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteQuestListSend", Erl)
        '</EhFooter>
End Sub

Public Sub WriteNpcQuestListSend(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteNpcQuestListSend_Err
        '</EhHeader>

        Dim i          As Integer

        Dim j          As Integer

        Dim QuestIndex As Integer

100     Call Writer.WriteInt(ServerPacketID.NpcQuestListSend)
102     Call Writer.WriteInt8(NpcList(NpcIndex).NumQuest) 'Escribimos primero cuantas quest tiene el NPC

104     For j = 1 To NpcList(NpcIndex).NumQuest
106         QuestIndex = NpcList(NpcIndex).QuestNumber(j)
108         Call Writer.WriteInt16(QuestIndex)
110         Call Writer.WriteInt8(QuestList(QuestIndex).RequiredLevel)
112         Call Writer.WriteInt16(QuestList(QuestIndex).RequiredQuest)
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

            'Enviamos la recompensa de oro y experiencia.
138         Call Writer.WriteInt32(QuestList(QuestIndex).RewardGLD * OroMult)
140         Call Writer.WriteInt32(QuestList(QuestIndex).RewardEXP * ExpMult)
            'Enviamos la cantidad de objs de recompensa
142         Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)

144         If QuestList(QuestIndex).RewardOBJs Then

                'si hay objs entonces enviamos la lista
146             For i = 1 To QuestList(QuestIndex).RewardOBJs
148                 Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).amount)
150                 Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
152             Next i

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

180     Next j

182     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteNpcQuestListSend_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNpcQuestListSend", Erl)
        '</EhFooter>
End Sub
Public Sub writeAnswerReset(ByVal UserIndex As Integer)
    On Error GoTo writeAnswerReset_Err
    
    Call Writer.WriteInt(ServerPacketID.AnswerReset)
    
182     Call modSendData.SendData(ToIndex, UserIndex)
    Exit Sub
writeAnswerReset_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.writeAnswerReset", Erl)
        '</EhFooter>
End Sub
Public Sub WriteObjQuestSend(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteNpcQuestListSend_Err
        '</EhHeader>
        Dim i As Integer

100     Call Writer.WriteInt(ServerPacketID.ObjQuestListSend)
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

            'Enviamos la recompensa de oro y experiencia.
138     Call Writer.WriteInt32(QuestList(QuestIndex).RewardGLD * OroMult)
140     Call Writer.WriteInt32(QuestList(QuestIndex).RewardEXP * ExpMult)
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

182     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteNpcQuestListSend_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteNpcQuestListSend", Erl)
        '</EhFooter>
End Sub
Sub WriteCommerceRecieveChatMessage(ByVal UserIndex As Integer, ByVal message As String)
        '<EhHeader>
        On Error GoTo WriteCommerceRecieveChatMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CommerceRecieveChatMessage)
102     Call Writer.WriteString8(message)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteCommerceRecieveChatMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCommerceRecieveChatMessage", Erl)
        '</EhFooter>
End Sub

Sub WriteInvasionInfo(ByVal UserIndex As Integer, _
                      ByVal Invasion As Integer, _
                      ByVal PorcentajeVida As Byte, _
                      ByVal PorcentajeTiempo As Byte)
        '<EhHeader>
        On Error GoTo WriteInvasionInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.InvasionInfo)
102     Call Writer.WriteInt8(Invasion)
104     Call Writer.WriteInt8(PorcentajeVida)
106     Call Writer.WriteInt8(PorcentajeTiempo)
108     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteInvasionInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteInvasionInfo", Erl)
        '</EhFooter>
End Sub

Sub WriteOpenCrafting(ByVal UserIndex As Integer, ByVal Tipo As Byte)
        '<EhHeader>
        On Error GoTo WriteOpenCrafting_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.OpenCrafting)
102     Call Writer.WriteInt8(Tipo)
104     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteOpenCrafting_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteOpenCrafting", Erl)
        '</EhFooter>
End Sub

Sub WriteCraftingItem(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal ObjIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteCraftingItem_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CraftingItem)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(ObjIndex)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteCraftingItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCraftingItem", Erl)
        '</EhFooter>
End Sub
Sub writeSendMousePosToClient(ByVal userindex As Integer, ByVal x As Integer, ByVal y As Integer)
  '<EhHeader>
        On Error GoTo writeSendMousePosToClient_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SendMousePos)
102     Call Writer.WriteInt16(x)
104     Call Writer.WriteInt16(y)
106     Call modSendData.SendData(ToIndex, 1)
        '<EhFooter>
        Exit Sub

writeSendMousePosToClient_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.writeSendMousePosToClient", Erl)
        '</EhFooter>
End Sub

Sub WriteCraftingCatalyst(ByVal UserIndex As Integer, _
                          ByVal ObjIndex As Integer, _
                          ByVal amount As Integer, _
                          ByVal Porcentaje As Byte)
        '<EhHeader>
        On Error GoTo WriteCraftingCatalyst_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CraftingCatalyst)
102     Call Writer.WriteInt16(ObjIndex)
104     Call Writer.WriteInt16(amount)
106     Call Writer.WriteInt8(Porcentaje)
108     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteCraftingCatalyst_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCraftingCatalyst", Erl)
        '</EhFooter>
End Sub

Sub WriteCraftingResult(ByVal UserIndex As Integer, _
                        ByVal Result As Integer, _
                        Optional ByVal Porcentaje As Byte = 0, _
                        Optional ByVal Precio As Long = 0)
        '<EhHeader>
        On Error GoTo WriteCraftingResult_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CraftingResult)
102     Call Writer.WriteInt16(Result)

104     If Result <> 0 Then
106         Call Writer.WriteInt8(Porcentaje)
108         Call Writer.WriteInt32(Precio)
        End If

110     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteCraftingResult_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteCraftingResult", Erl)
        '</EhFooter>
End Sub

Sub WriteForceUpdate(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteForceUpdate_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ForceUpdate)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteForceUpdate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteForceUpdate", Erl)
        '</EhFooter>
End Sub

Public Sub WriteUpdateNPCSimbolo(ByVal UserIndex As Integer, _
                                 ByVal NpcIndex As Integer, _
                                 ByVal Simbolo As Byte)
        '<EhHeader>
        On Error GoTo WriteUpdateNPCSimbolo_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateNPCSimbolo)
102     Call Writer.WriteInt16(NpcList(NpcIndex).Char.CharIndex)
104     Call Writer.WriteInt8(Simbolo)
106     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteUpdateNPCSimbolo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteUpdateNPCSimbolo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteGuardNotice(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteGuardNotice_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.GuardNotice)
102     Call modSendData.SendData(ToIndex, UserIndex)
        '<EhFooter>
        Exit Sub

WriteGuardNotice_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.WriteGuardNotice", Erl)
        '</EhFooter>
End Sub

' \Begin: [Prepares]
Public Function PrepareMessageCharSwing(ByVal CharIndex As Integer, _
                                        Optional ByVal FX As Boolean = True, _
                                        Optional ByVal ShowText As Boolean = True)
        '<EhHeader>
        On Error GoTo PrepareMessageCharSwing_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CharSwing)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteBool(FX)
106     Call Writer.WriteBool(ShowText)
        '<EhFooter>
        Exit Function

PrepareMessageCharSwing_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharSwing", Erl)
        '</EhFooter>
End Function

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.
Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, _
                                           ByVal invisible As Boolean)
        '<EhHeader>
        On Error GoTo PrepareMessageSetInvisible_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SetInvisible)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteBool(invisible)
        '<EhFooter>
        Exit Function

PrepareMessageSetInvisible_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageSetInvisible", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageSetEscribiendo(ByVal CharIndex As Integer, _
                                             ByVal Escribiendo As Boolean)
        '<EhHeader>
        On Error GoTo PrepareMessageSetEscribiendo_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SetEscribiendo)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteBool(Escribiendo)
        '<EhFooter>
        Exit Function

PrepareMessageSetEscribiendo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageSetEscribiendo", Erl)
        '</EhFooter>
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
                                           ByVal CharIndex As Integer, _
                                           ByVal Color As Long)
        '<EhHeader>
        On Error GoTo PrepareMessageChatOverHead_Err
        '</EhHeader>

        Dim R As Long, G As Long, B As Long

100     B = (Color And 16711680) / 65536
102     G = (Color And 65280) / 256
104     R = Color And 255
106     Call Writer.WriteInt(ServerPacketID.ChatOverHead)
108     Call Writer.WriteString8(chat)
110     Call Writer.WriteInt16(CharIndex)
        ' Write rgb channels and save one byte from long :D
112     Call Writer.WriteInt8(R)
114     Call Writer.WriteInt8(G)
116     Call Writer.WriteInt8(B)
118     Call Writer.WriteInt32(Color)
        '<EhFooter>
        Exit Function

PrepareMessageChatOverHead_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageChatOverHead", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageTextOverChar(ByVal chat As String, _
                                           ByVal CharIndex As Integer, _
                                           ByVal Color As Long)
        '<EhHeader>
        On Error GoTo PrepareMessageTextOverChar_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.TextOverChar)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt16(CharIndex)
106     Call Writer.WriteInt32(Color)
        '<EhFooter>
        Exit Function

PrepareMessageTextOverChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageTextOverChar", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageTextCharDrop(ByVal chat As String, _
                                           ByVal CharIndex As Integer, _
                                           ByVal Color As Long)
        '<EhHeader>
        On Error GoTo PrepareMessageTextCharDrop_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.TextCharDrop)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt16(CharIndex)
106     Call Writer.WriteInt32(Color)
        '<EhFooter>
        Exit Function

PrepareMessageTextCharDrop_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageTextCharDrop", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageTextOverTile(ByVal chat As String, _
                                           ByVal X As Integer, _
                                           ByVal Y As Integer, _
                                           ByVal Color As Long)
        '<EhHeader>
        On Error GoTo PrepareMessageTextOverTile_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.TextOverTile)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt16(X)
106     Call Writer.WriteInt16(Y)
108     Call Writer.WriteInt32(Color)
        '<EhFooter>
        Exit Function

PrepareMessageTextOverTile_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageTextOverTile", Erl)
        '</EhFooter>
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageConsoleMsg(ByVal chat As String, _
                                         ByVal FontIndex As e_FontTypeNames)
        '<EhHeader>
        On Error GoTo PrepareMessageConsoleMsg_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ConsoleMsg)
102     Call Writer.WriteString8(chat)
104     Call Writer.WriteInt8(FontIndex)
        '<EhFooter>
        Exit Function

PrepareMessageConsoleMsg_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageConsoleMsg", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageLocaleMsg(ByVal ID As Integer, _
                                        ByVal chat As String, _
                                        ByVal FontIndex As e_FontTypeNames)
        '<EhHeader>
        On Error GoTo PrepareMessageLocaleMsg_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.LocaleMsg)
102     Call Writer.WriteInt16(ID)
104     Call Writer.WriteString8(chat)
106     Call Writer.WriteInt8(FontIndex)
        '<EhFooter>
        Exit Function

PrepareMessageLocaleMsg_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageLocaleMsg", Erl)
        '</EhFooter>
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
                                       ByVal FX As Integer, _
                                       ByVal FXLoops As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageCreateFX_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CreateFX)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(FX)
106     Call Writer.WriteInt16(FXLoops)
        '<EhFooter>
        Exit Function

PrepareMessageCreateFX_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCreateFX", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageMeditateToggle(ByVal CharIndex As Integer, _
                                             ByVal FX As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageMeditateToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.MeditateToggle)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(FX)
        '<EhFooter>
        Exit Function

PrepareMessageMeditateToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageMeditateToggle", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageParticleFX(ByVal CharIndex As Integer, _
                                         ByVal Particula As Integer, _
                                         ByVal Time As Long, _
                                         ByVal Remove As Boolean, _
                                         Optional ByVal grh As Long = 0)
        '<EhHeader>
        On Error GoTo PrepareMessageParticleFX_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ParticleFX)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(Particula)
106     Call Writer.WriteInt32(Time)
108     Call Writer.WriteBool(Remove)
110     Call Writer.WriteInt32(grh)
        '<EhFooter>
        Exit Function

PrepareMessageParticleFX_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFX", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageParticleFXWithDestino(ByVal Emisor As Integer, _
                                                    ByVal Receptor As Integer, _
                                                    ByVal ParticulaViaje As Integer, _
                                                    ByVal ParticulaFinal As Integer, _
                                                    ByVal Time As Long, _
                                                    ByVal wav As Integer, _
                                                    ByVal FX As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageParticleFXWithDestino_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ParticleFXWithDestino)
102     Call Writer.WriteInt16(Emisor)
104     Call Writer.WriteInt16(Receptor)
106     Call Writer.WriteInt16(ParticulaViaje)
108     Call Writer.WriteInt16(ParticulaFinal)
110     Call Writer.WriteInt32(Time)
112     Call Writer.WriteInt16(wav)
114     Call Writer.WriteInt16(FX)
        '<EhFooter>
        Exit Function

PrepareMessageParticleFXWithDestino_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFXWithDestino", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageParticleFXWithDestinoXY(ByVal Emisor As Integer, _
                                                      ByVal ParticulaViaje As Integer, _
                                                      ByVal ParticulaFinal As Integer, _
                                                      ByVal Time As Long, _
                                                      ByVal wav As Integer, _
                                                      ByVal FX As Integer, _
                                                      ByVal X As Byte, _
                                                      ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageParticleFXWithDestinoXY_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ParticleFXWithDestinoXY)
102     Call Writer.WriteInt16(Emisor)
104     Call Writer.WriteInt16(ParticulaViaje)
106     Call Writer.WriteInt16(ParticulaFinal)
108     Call Writer.WriteInt32(Time)
110     Call Writer.WriteInt16(wav)
112     Call Writer.WriteInt16(FX)
114     Call Writer.WriteInt8(X)
116     Call Writer.WriteInt8(Y)
        '<EhFooter>
        Exit Function

PrepareMessageParticleFXWithDestinoXY_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFXWithDestinoXY", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageAuraToChar(ByVal CharIndex As Integer, _
                                         ByVal Aura As String, _
                                         ByVal Remove As Boolean, _
                                         ByVal Tipo As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageAuraToChar_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.AuraToChar)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteString8(Aura)
106     Call Writer.WriteBool(Remove)
108     Call Writer.WriteInt8(Tipo)
        '<EhFooter>
        Exit Function

PrepareMessageAuraToChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageAuraToChar", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageSpeedingACT(ByVal CharIndex As Integer, _
                                          ByVal speeding As Single)
        '<EhHeader>
        On Error GoTo PrepareMessageSpeedingACT_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.SpeedToChar)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteReal32(speeding)
        '<EhFooter>
        Exit Function

PrepareMessageSpeedingACT_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageSpeedingACT", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageParticleFXToFloor(ByVal X As Byte, _
                                                ByVal Y As Byte, _
                                                ByVal Particula As Integer, _
                                                ByVal Time As Long)
        '<EhHeader>
        On Error GoTo PrepareMessageParticleFXToFloor_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ParticleFXToFloor)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt16(Particula)
108     Call Writer.WriteInt32(Time)
        '<EhFooter>
        Exit Function

PrepareMessageParticleFXToFloor_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageParticleFXToFloor", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageLightFXToFloor(ByVal X As Byte, _
                                             ByVal Y As Byte, _
                                             ByVal LuzColor As Long, _
                                             ByVal Rango As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageLightFXToFloor_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.LightToFloor)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt32(LuzColor)
108     Call Writer.WriteInt8(Rango)
        '<EhFooter>
        Exit Function

PrepareMessageLightFXToFloor_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageLightFXToFloor", Erl)
        '</EhFooter>
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
                                       ByVal X As Byte, _
                                       ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessagePlayWave_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.PlayWave)
102     Call Writer.WriteInt16(wave)
104     Call Writer.WriteInt8(X)
106     Call Writer.WriteInt8(Y)
        '<EhFooter>
        Exit Function

PrepareMessagePlayWave_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessagePlayWave", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageUbicacionLlamada(ByVal Mapa As Integer, _
                                               ByVal X As Byte, _
                                               ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageUbicacionLlamada_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.PosLLamadaDeClan)
102     Call Writer.WriteInt16(Mapa)
104     Call Writer.WriteInt8(X)
106     Call Writer.WriteInt8(Y)
        '<EhFooter>
        Exit Function

PrepareMessageUbicacionLlamada_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageUbicacionLlamada", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageCharUpdateHP(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageCharUpdateHP_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CharUpdateHP)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
104     Call Writer.WriteInt32(UserList(UserIndex).Stats.MinHp)
106     Call Writer.WriteInt32(UserList(UserIndex).Stats.MaxHp)
        '<EhFooter>
        Exit Function

PrepareMessageCharUpdateHP_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharUpdateHP", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageCharUpdateMAN(ByVal UserIndex As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageCharUpdateMAN_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CharUpdateMAN)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
104     Call Writer.WriteInt32(UserList(UserIndex).Stats.MinMAN)
106     Call Writer.WriteInt32(UserList(UserIndex).Stats.MaxMAN)
        '<EhFooter>
        Exit Function

PrepareMessageCharUpdateMAN_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharUpdateMAN", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageNpcUpdateHP(ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageNpcUpdateHP_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CharUpdateHP)
102     Call Writer.WriteInt16(NpcList(NpcIndex).Char.CharIndex)
104     Call Writer.WriteInt32(NpcList(NpcIndex).Stats.MinHp)
106     Call Writer.WriteInt32(NpcList(NpcIndex).Stats.MaxHp)
        '<EhFooter>
        Exit Function

PrepareMessageNpcUpdateHP_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageNpcUpdateHP", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageArmaMov(ByVal CharIndex As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageArmaMov_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ArmaMov)
102     Call Writer.WriteInt16(CharIndex)
        '<EhFooter>
        Exit Function

PrepareMessageArmaMov_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageArmaMov", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageEscudoMov(ByVal CharIndex As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageEscudoMov_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.EscudoMov)
102     Call Writer.WriteInt16(CharIndex)
        '<EhFooter>
        Exit Function

PrepareMessageEscudoMov_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageEscudoMov", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageFlashScreen(ByVal Color As Long, _
                                          ByVal Duracion As Long, _
                                          Optional ByVal Ignorar As Boolean = False)
        '<EhHeader>
        On Error GoTo PrepareMessageFlashScreen_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.FlashScreen)
102     Call Writer.WriteInt32(Color)
104     Call Writer.WriteInt32(Duracion)
106     Call Writer.WriteBool(Ignorar)
        '<EhFooter>
        Exit Function

PrepareMessageFlashScreen_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageFlashScreen", Erl)
        '</EhFooter>
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageGuildChat(ByVal chat As String, ByVal Status As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageGuildChat_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.GuildChat)
102     Call Writer.WriteInt8(Status)
104     Call Writer.WriteString8(chat)
        '<EhFooter>
        Exit Function

PrepareMessageGuildChat_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageGuildChat", Erl)
        '</EhFooter>
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageShowMessageBox(ByVal chat As String)
        '<EhHeader>
        On Error GoTo PrepareMessageShowMessageBox_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ShowMessageBox)
102     Call Writer.WriteString8(chat)
        '<EhFooter>
        Exit Function

PrepareMessageShowMessageBox_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageShowMessageBox", Erl)
        '</EhFooter>
End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessagePlayMidi(ByVal midi As Byte, _
                                       Optional ByVal loops As Integer = -1)
        '<EhHeader>
        On Error GoTo PrepareMessagePlayMidi_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.PlayMIDI)
102     Call Writer.WriteInt8(midi)
104     Call Writer.WriteInt16(loops)
        '<EhFooter>
        Exit Function

PrepareMessagePlayMidi_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessagePlayMidi", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageOnlineUser(ByVal UserOnline As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageOnlineUser_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UserOnline)
102     Call Writer.WriteInt16(UserOnline)
        '<EhFooter>
        Exit Function

PrepareMessageOnlineUser_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageOnlineUser", Erl)
        '</EhFooter>
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessagePauseToggle()
        '<EhHeader>
        On Error GoTo PrepareMessagePauseToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.PauseToggle)
        '<EhFooter>
        Exit Function

PrepareMessagePauseToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessagePauseToggle", Erl)
        '</EhFooter>
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageRainToggle()
        '<EhHeader>
        On Error GoTo PrepareMessageRainToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.RainToggle)
        '<EhFooter>
        Exit Function

PrepareMessageRainToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageRainToggle", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageHora()
        '<EhHeader>
        On Error GoTo PrepareMessageHora_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.Hora)
102     Call Writer.WriteInt32((GetTickCount() - HoraMundo) Mod DuracionDia)
104     Call Writer.WriteInt32(DuracionDia)
        '<EhFooter>
        Exit Function

PrepareMessageHora_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageHora", Erl)
        '</EhFooter>
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageObjectDelete_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ObjectDelete)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
        '<EhFooter>
        Exit Function

PrepareMessageObjectDelete_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageObjectDelete", Erl)
        '</EhFooter>
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
                                            ByVal Y As Byte, _
                                            ByVal Blocked As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessage_BlockPosition_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.BlockPosition)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt8(Blocked)
        '<EhFooter>
        Exit Function

PrepareMessage_BlockPosition_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessage_BlockPosition", Erl)
        '</EhFooter>
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
                                           ByVal X As Byte, _
                                           ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageObjectCreate_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ObjectCreate)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt16(ObjIndex)
108     Call Writer.WriteInt16(amount)
        '<EhFooter>
        Exit Function

PrepareMessageObjectCreate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageObjectCreate", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageFxPiso(ByVal GrhIndex As Integer, _
                                     ByVal X As Byte, _
                                     ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageFxPiso_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.fxpiso)
102     Call Writer.WriteInt8(X)
104     Call Writer.WriteInt8(Y)
106     Call Writer.WriteInt16(GrhIndex)
        '<EhFooter>
        Exit Function

PrepareMessageFxPiso_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageFxPiso", Erl)
        '</EhFooter>
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer, _
                                              ByVal Desvanecido As Boolean)
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterRemove_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CharacterRemove)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteBool(Desvanecido)
        '<EhFooter>
        Exit Function

PrepareMessageCharacterRemove_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterRemove", Erl)
        '</EhFooter>
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer)
        '<EhHeader>
        On Error GoTo PrepareMessageRemoveCharDialog_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.RemoveCharDialog)
102     Call Writer.WriteInt16(CharIndex)
        '<EhFooter>
        Exit Function

PrepareMessageRemoveCharDialog_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageRemoveCharDialog", Erl)
        '</EhFooter>
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
Public Function PrepareMessageCharacterCreate(ByVal Body As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal Heading As e_Heading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal X As Byte, _
                                              ByVal Y As Byte, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal Name As String, _
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
                                              ByVal speeding As Single, _
                                              ByVal EsNPC As Byte, _
                                              ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal UserMinMAN As Long, ByVal UserMaxMAN As Long, ByVal Simbolo As Byte, ByVal Idle As Boolean, ByVal Navegando As Boolean, ByVal tipoUsuario As e_TipoUsuario)
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterCreate_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CharacterCreate)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(Body)
106     Call Writer.WriteInt16(Head)
108     Call Writer.WriteInt8(Heading)
110     Call Writer.WriteInt8(X)
112     Call Writer.WriteInt8(Y)
114     Call Writer.WriteInt16(weapon)
116     Call Writer.WriteInt16(shield)
118     Call Writer.WriteInt16(helmet)
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
168     Call Writer.WriteBool(Idle)
170     Call Writer.WriteBool(Navegando)
172     Call Writer.WriteInt16(tipoUsuario)
        '<EhFooter>
        Exit Function

PrepareMessageCharacterCreate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterCreate", Erl)
        '</EhFooter>
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
                                              ByVal Head As Integer, _
                                              ByVal Heading As e_Heading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal Idle As Boolean, _
                                              ByVal Navegando As Boolean)
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterChange_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CharacterChange)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(Body)
106     Call Writer.WriteInt16(Head)
108     Call Writer.WriteInt8(Heading)
110     Call Writer.WriteInt16(weapon)
112     Call Writer.WriteInt16(shield)
114     Call Writer.WriteInt16(helmet)
116     Call Writer.WriteInt16(FX)
118     Call Writer.WriteInt16(FXLoops)
120     Call Writer.WriteBool(Idle)
122     Call Writer.WriteBool(Navegando)
        '<EhFooter>
        Exit Function

PrepareMessageCharacterChange_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterChange", Erl)
        '</EhFooter>
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
                                            ByVal X As Byte, _
                                            ByVal Y As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageCharacterMove_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.CharacterMove)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt8(X)
106     Call Writer.WriteInt8(Y)
        '<EhFooter>
        Exit Function

PrepareMessageCharacterMove_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageCharacterMove", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As e_Heading)
        '<EhHeader>
        On Error GoTo PrepareMessageForceCharMove_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ForceCharMove)
102     Call Writer.WriteInt8(Direccion)
        '<EhFooter>
        Exit Function

PrepareMessageForceCharMove_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageForceCharMove", Erl)
        '</EhFooter>
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
                                                 Status As Byte, _
                                                 Tag As String)
        '<EhHeader>
        On Error GoTo PrepareMessageUpdateTagAndStatus_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.UpdateTagAndStatus)
102     Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
104     Call Writer.WriteInt8(Status)
106     Call Writer.WriteString8(Tag)
108     Call Writer.WriteInt16(UserList(UserIndex).Grupo.Lider)
        '<EhFooter>
        Exit Function

PrepareMessageUpdateTagAndStatus_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageUpdateTagAndStatus", Erl)
        '</EhFooter>
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageErrorMsg(ByVal message As String)
        '<EhHeader>
        On Error GoTo PrepareMessageErrorMsg_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.ErrorMsg)
102     Call Writer.WriteString8(message)
        '<EhFooter>
        Exit Function

PrepareMessageErrorMsg_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageErrorMsg", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageBarFx(ByVal CharIndex As Integer, _
                                    ByVal BarTime As Integer, _
                                    ByVal BarAccion As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageBarFx_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.BarFx)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(BarTime)
106     Call Writer.WriteInt8(BarAccion)
        '<EhFooter>
        Exit Function

PrepareMessageBarFx_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageBarFx", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageNieblandoToggle(ByVal IntensidadMax As Byte)
        '<EhHeader>
        On Error GoTo PrepareMessageNieblandoToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.NieblaToggle)
102     Call Writer.WriteInt8(IntensidadMax)
        '<EhFooter>
        Exit Function

PrepareMessageNieblandoToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageNieblandoToggle", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageNevarToggle()
        '<EhHeader>
        On Error GoTo PrepareMessageNevarToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt(ServerPacketID.NieveToggle)
        '<EhFooter>
        Exit Function

PrepareMessageNevarToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageNevarToggle", Erl)
        '</EhFooter>
End Function

Public Function PrepareMessageDoAnimation(ByVal CharIndex As Integer, _
                                          ByVal Animation As Integer)

        On Error GoTo PrepareMessageDoAnimation_Err

100     Call Writer.WriteInt(ServerPacketID.DoAnimation)
102     Call Writer.WriteInt16(CharIndex)
104     Call Writer.WriteInt16(Animation)

        Exit Function

PrepareMessageDoAnimation_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PrepareMessageDoAnimation", Erl)
End Function


'Public Function WritePescarEspecial(ByVal ObjIndex As Integer)

'        On Error GoTo PescarEspecial_Err
'100     Call Writer.WriteInt(ServerPacketID.PescarEspecial)
'        Call Writer.WriteInt16(ObjIndex)

'PescarEspecial_Err:
'        Call Writer.Clear
'        Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Protocol_Writes.PescarEspecial", Erl)
'End Function


