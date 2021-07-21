Attribute VB_Name = "Protocol_Writes"
Option Explicit

Private Writer  As Network.Writer

Public Sub InitializeAuxiliaryBuffer()
    Set Writer = New Network.Writer
End Sub
    
Public Function GetWriterBuffer() As Network.Writer
    Set GetWriterBuffer = Writer
End Function

' \Begin: [Writes]

Public Sub WriteConnected(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.Connected)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "Logged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.logged)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteHora(ByVal UserIndex As Integer)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageHora())
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.RemoveDialogs)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageRemoveCharDialog( _
            CharIndex))
End Sub

' Writes the "NavigateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.NavigateToggle)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteNadarToggle(ByVal UserIndex As Integer, _
                            ByVal Puede As Boolean, _
                            Optional ByVal esTrajeCaucho As Boolean = False)
    Call Writer.WriteInt(ServerPacketID.NadarToggle)
    Call Writer.WriteBool(Puede)
    Call Writer.WriteBool(esTrajeCaucho)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteEquiteToggle(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.EquiteToggle)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteVelocidadToggle(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.VelocidadToggle)
    Call Writer.WriteReal32(UserList(UserIndex).Char.speeding)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteMacroTrabajoToggle(ByVal UserIndex As Integer, ByVal Activar As Boolean)

    If Not Activar Then
        UserList(UserIndex).flags.TargetObj = 0 ' Sacamos el targer del objeto
        UserList(UserIndex).flags.UltimoMensaje = 0
        UserList(UserIndex).Counters.Trabajando = 0
        UserList(UserIndex).flags.UsandoMacro = False
        UserList(UserIndex).trabajo.Target_X = 0
        UserList(UserIndex).trabajo.Target_Y = 0
        UserList(UserIndex).trabajo.TargetSkill = 0
    Else
        UserList(UserIndex).flags.UsandoMacro = True
    End If

    Call Writer.WriteInt(ServerPacketID.MacroTrabajoToggle)
    Call Writer.WriteBool(Activar)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDisconnect(ByVal UserIndex As Integer, _
                           Optional ByVal FullLogout As Boolean = False)
    Call ClearAndSaveUser(UserIndex)
    UserList(UserIndex).flags.YaGuardo = True

    If Not FullLogout Then
        Call WritePersonajesDeCuenta(UserIndex)
        Call WriteMostrarCuenta(UserIndex)
    End If

    Call Writer.WriteInt(ServerPacketID.Disconnect)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.CommerceEnd)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankEnd(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.BankEnd)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.CommerceInit)
    Call Writer.WriteString8(NpcList(UserList(UserIndex).flags.TargetNPC).Name)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankInit(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.BankInit)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.UserCommerceInit)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.UserCommerceEnd)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowBlacksmithForm)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowCarpenterForm)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteShowAlquimiaForm(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowAlquimiaForm)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteShowSastreForm(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowSastreForm)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.NPCKillUser)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlockedWithShieldUser(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.BlockedWithShieldUser)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlockedWithShieldOther(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.BlockedWithShieldOther)
    Call modSendData.SendData(ToIndex, UserIndex)
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharSwing(CharIndex, _
            FX, ShowText))
End Sub

''
' Writes the "SafeModeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.SafeModeOn)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.SafeModeOff)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "PartySafeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePartySafeOn(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.PartySafeOn)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "PartySafeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePartySafeOff(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.PartySafeOff)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteClanSeguro(ByVal UserIndex As Integer, ByVal estado As Boolean)
    Call Writer.WriteInt(ServerPacketID.ClanSeguro)
    Call Writer.WriteBool(estado)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteSeguroResu(ByVal UserIndex As Integer, ByVal estado As Boolean)
    Call Writer.WriteInt(ServerPacketID.SeguroResu)
    Call Writer.WriteBool(estado)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.CantUseWhileMeditating)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.UpdateSta)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
            PrepareMessageCharUpdateMAN(UserIndex))
    Call Writer.WriteInt(ServerPacketID.UpdateMana)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
    'Call SendData(SendTarget.ToDiosesYclan, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
            PrepareMessageCharUpdateHP(UserIndex))
    Call Writer.WriteInt(ServerPacketID.UpdateHP)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.UpdateGold)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.UpdateExp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer)
    Call Writer.WriteInt(ServerPacketID.changeMap)
    Call Writer.WriteInt16(Map)
    Call Writer.WriteInt16(0)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePosUpdate(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.PosUpdate)
    Call Writer.WriteInt8(UserList(UserIndex).Pos.X)
    Call Writer.WriteInt8(UserList(UserIndex).Pos.Y)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, _
                           ByVal Target As PartesCuerpo, _
                           ByVal damage As Integer)
    Call Writer.WriteInt(ServerPacketID.NPCHitUser)
    Call Writer.WriteInt8(Target)
    Call Writer.WriteInt16(damage)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserHitNPC(ByVal UserIndex As Integer, ByVal damage As Long)
    Call Writer.WriteInt(ServerPacketID.UserHitNPC)
    Call Writer.WriteInt32(damage)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserAttackedSwing(ByVal UserIndex As Integer, _
                                  ByVal AttackerIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.UserAttackedSwing)
    Call Writer.WriteInt16(UserList(AttackerIndex).Char.CharIndex)
    Call modSendData.SendData(ToIndex, UserIndex)
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
                                 ByVal Target As PartesCuerpo, _
                                 ByVal attackerChar As Integer, _
                                 ByVal damage As Integer)
    Call Writer.WriteInt(ServerPacketID.UserHittedByUser)
    Call Writer.WriteInt16(attackerChar)
    Call Writer.WriteInt8(Target)
    Call Writer.WriteInt16(damage)
    Call modSendData.SendData(ToIndex, UserIndex)
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
                               ByVal Target As PartesCuerpo, _
                               ByVal attackedChar As Integer, _
                               ByVal damage As Integer)
    Call Writer.WriteInt(ServerPacketID.UserHittedUser)
    Call Writer.WriteInt16(attackedChar)
    Call Writer.WriteInt8(Target)
    Call Writer.WriteInt16(damage)
    Call modSendData.SendData(ToIndex, UserIndex)
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageChatOverHead(chat, _
            CharIndex, Color))
End Sub

Public Sub WriteTextOverChar(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal Color As Long)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextOverChar(chat, _
            CharIndex, Color))
End Sub

Public Sub WriteTextOverTile(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal X As Integer, _
                             ByVal Y As Integer, _
                             ByVal Color As Long)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextOverTile(chat, X, _
            Y, Color))
End Sub

Public Sub WriteTextCharDrop(ByVal UserIndex As Integer, _
                             ByVal chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal Color As Long)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextCharDrop(chat, _
            CharIndex, Color))
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
                           ByVal FontIndex As FontTypeNames)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageConsoleMsg(chat, _
            FontIndex))
End Sub

Public Sub WriteLocaleMsg(ByVal UserIndex As Integer, _
                          ByVal ID As Integer, _
                          ByVal FontIndex As FontTypeNames, _
                          Optional ByVal strExtra As String = vbNullString)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageLocaleMsg(ID, strExtra, _
            FontIndex))
End Sub

Public Sub WriteListaCorreo(ByVal UserIndex As Integer, ByVal Actualizar As Boolean)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageListaCorreo(UserIndex, _
            Actualizar))
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageGuildChat(chat, Status))
End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String)
    Call Writer.WriteInt(ServerPacketID.ShowMessageBox)
    Call Writer.WriteString8(Message)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteMostrarCuenta(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.MostrarCuenta)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.UserIndexInServer)
    Call Writer.WriteInt16(UserIndex)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.UserCharIndexInServer)
    Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
    Call modSendData.SendData(ToIndex, UserIndex)
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
Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal Status As Byte, ByVal privileges As Byte, ByVal ParticulaFx As Byte, ByVal Head_Aura As String, ByVal Arma_Aura As String, ByVal Body_Aura As String, ByVal DM_Aura As String, ByVal RM_Aura As String, ByVal Otra_Aura As String, ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Byte, ByVal donador As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal UserMinMAN As Long, ByVal UserMaxMAN As Long, ByVal Simbolo As Byte, Optional ByVal Idle As Boolean = False, Optional ByVal Navegando As Boolean = False)
Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterCreate(Body, Head, _
        Heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, Name, Status, _
        privileges, ParticulaFx, Head_Aura, Arma_Aura, Body_Aura, DM_Aura, RM_Aura, _
        Otra_Aura, Escudo_Aura, speeding, EsNPC, donador, appear, group_index, _
        clan_index, clan_nivel, UserMinHp, UserMaxHp, UserMinMAN, UserMaxMAN, Simbolo, _
        Idle, Navegando))
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterRemove( _
            CharIndex, Desvanecido))
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterMove( _
            CharIndex, X, Y))
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageForceCharMove(Direccion))
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
                                ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                Optional ByVal Idle As Boolean = False, _
                                Optional ByVal Navegando As Boolean = False)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCharacterChange(Body, _
            Head, Heading, CharIndex, weapon, shield, FX, FXLoops, helmet, Idle, _
            Navegando))
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageObjectCreate(ObjIndex, _
            amount, X, Y))
End Sub

Public Sub WriteParticleFloorCreate(ByVal UserIndex As Integer, _
                                    ByVal Particula As Integer, _
                                    ByVal ParticulaTime As Integer, _
                                    ByVal Map As Integer, _
                                    ByVal X As Byte, _
                                    ByVal Y As Byte)

    If Particula = 0 Then
        Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageParticleFXToFloor( _
                X, Y, Particula, ParticulaTime))
    End If

End Sub

Public Sub WriteLightFloorCreate(ByVal UserIndex As Integer, _
                                 ByVal LuzColor As Long, _
                                 ByVal Rango As Byte, _
                                 ByVal Map As Integer, _
                                 ByVal X As Byte, _
                                 ByVal Y As Byte)
    MapData(Map, X, Y).Luz.Color = LuzColor
    MapData(Map, X, Y).Luz.Rango = Rango

    If Rango = 0 Then
        Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageLightFXToFloor(X, _
                Y, LuzColor, Rango))
    End If

End Sub

Public Sub WriteFxPiso(ByVal UserIndex As Integer, _
                       ByVal GrhIndex As Integer, _
                       ByVal X As Byte, _
                       ByVal Y As Byte)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageFxPiso(GrhIndex, X, Y))
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageObjectDelete(X, Y))
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlockPosition(ByVal UserIndex As Integer, _
                              ByVal X As Byte, _
                              ByVal Y As Byte, _
                              ByVal Blocked As Byte)
    Call Writer.WriteInt(ServerPacketID.BlockPosition)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteInt8(Blocked)
    Call modSendData.SendData(ToIndex, UserIndex)
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePlayMidi(midi, loops))
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePlayWave(wave, X, Y))
End Sub

''
' Writes the "GuildList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)

    Dim Tmp As String

    Dim I   As Long

    Call Writer.WriteInt(ServerPacketID.guildList)

    ' Prepare guild name's list
    For I = LBound(guildList()) To UBound(guildList())
        Tmp = Tmp & guildList(I) & SEPARATOR
    Next I

    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAreaChanged(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.AreaChanged)
    Call Writer.WriteInt8(UserList(UserIndex).Pos.X)
    Call Writer.WriteInt8(UserList(UserIndex).Pos.Y)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePauseToggle(ByVal UserIndex As Integer)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessagePauseToggle())
End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRainToggle(ByVal UserIndex As Integer)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageRainToggle())
End Sub

Public Sub WriteNubesToggle(ByVal UserIndex As Integer)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageNieblandoToggle( _
            IntensidadDeNubes))
End Sub

Public Sub WriteTrofeoToggleOn(ByVal UserIndex As Integer)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTrofeoToggleOn())
End Sub

Public Sub WriteTrofeoToggleOff(ByVal UserIndex As Integer)
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTrofeoToggleOff())
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageCreateFX(CharIndex, FX, _
            FXLoops))
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
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
            PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, _
            PrepareMessageCharUpdateMAN(UserIndex))
    Call Writer.WriteInt(ServerPacketID.UpdateUserStats)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxHp)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinHp)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxMAN)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinMAN)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MaxSta)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.MinSta)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.GLD)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.ELV)
    Call Writer.WriteInt32(ExpLevelUp(UserList(UserIndex).Stats.ELV))
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Exp)
    Call Writer.WriteInt8(UserList(UserIndex).clase)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteUpdateUserKey(ByVal UserIndex As Integer, _
                              ByVal Slot As Integer, _
                              ByVal Llave As Integer)
    Call Writer.WriteInt(ServerPacketID.UpdateUserKey)
    Call Writer.WriteInt16(Slot)
    Call Writer.WriteInt16(Llave)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

' Actualiza el indicador de daño mágico
Public Sub WriteUpdateDM(ByVal UserIndex As Integer)

    Dim Valor As Integer

    With UserList(UserIndex).Invent

        ' % daño mágico del arma
        If .WeaponEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.WeaponEqpObjIndex).MagicDamageBonus
        End If

        ' % daño mágico del anillo
        If .DañoMagicoEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.DañoMagicoEqpObjIndex).MagicDamageBonus
        End If

        Call Writer.WriteInt(ServerPacketID.UpdateDM)
        Call Writer.WriteInt16(Valor)
    End With

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

' Actualiza el indicador de resistencia mágica
Public Sub WriteUpdateRM(ByVal UserIndex As Integer)

    Dim Valor As Integer

    With UserList(UserIndex).Invent

        ' Resistencia mágica de la armadura
        If .ArmourEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.ArmourEqpObjIndex).ResistenciaMagica
        End If

        ' Resistencia mágica del anillo
        If .ResistenciaEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.ResistenciaEqpObjIndex).ResistenciaMagica
        End If

        ' Resistencia mágica del escudo
        If .EscudoEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.EscudoEqpObjIndex).ResistenciaMagica
        End If

        ' Resistencia mágica del casco
        If .CascoEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.CascoEqpObjIndex).ResistenciaMagica
        End If

        Valor = Valor + 100 * ModClase(UserList(UserIndex).clase).ResistenciaMagica
        Call Writer.WriteInt(ServerPacketID.UpdateRM)
        Call Writer.WriteInt16(Valor)
    End With

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)
    Call Writer.WriteInt(ServerPacketID.WorkRequestTarget)
    Call Writer.WriteInt8(Skill)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

' Writes the "InventoryUnlockSlots" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteInventoryUnlockSlots(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.InventoryUnlockSlots)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.InventLevel)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteIntervals(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Call Writer.WriteInt(ServerPacketID.Intervals)
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
End Sub

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

    Dim ObjIndex    As Integer

    Dim PodraUsarlo As Byte

    Call Writer.WriteInt(ServerPacketID.ChangeInventorySlot)
    Call Writer.WriteInt8(Slot)
    ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex

    If ObjIndex > 0 Then
        PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)
    End If

    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt16(UserList(UserIndex).Invent.Object(Slot).amount)
    Call Writer.WriteBool(UserList(UserIndex).Invent.Object(Slot).Equipped)
    Call Writer.WriteReal32(SalePrice(ObjIndex))
    Call Writer.WriteInt8(PodraUsarlo)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

    Dim ObjIndex    As Integer

    Dim Valor       As Long

    Dim PodraUsarlo As Byte

    Call Writer.WriteInt(ServerPacketID.ChangeBankSlot)
    Call Writer.WriteInt8(Slot)
    ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
    Call Writer.WriteInt16(ObjIndex)

    If ObjIndex > 0 Then
        Valor = ObjData(ObjIndex).Valor
        PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)
    End If

    Call Writer.WriteInt16(UserList(UserIndex).BancoInvent.Object(Slot).amount)
    Call Writer.WriteInt32(Valor)
    Call Writer.WriteInt8(PodraUsarlo)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
    Call Writer.WriteInt(ServerPacketID.ChangeSpellSlot)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.UserHechizos(Slot))

    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call Writer.WriteInt8(UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call Writer.WriteInt8(255)
    End If

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAttributes(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.Atributes)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
            eAtributos.Inteligencia))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
            eAtributos.Constitucion))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer

    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    Call Writer.WriteInt(ServerPacketID.BlacksmithWeapons)

    For I = 1 To UBound(ArmasHerrero())

        ' Can the user create this object? If so add it to the list....
        If ObjData(ArmasHerrero(I)).SkHerreria <= UserList(UserIndex).Stats.UserSkills( _
                eSkill.Herreria) Then
            Count = Count + 1
            validIndexes(Count) = I
        End If

    Next I

    ' Write the number of objects in the list
    Call Writer.WriteInt16(Count)

    ' Write the needed data of each object
    For I = 1 To Count
        obj = ObjData(ArmasHerrero(validIndexes(I)))
        'Call Writer.WriteString8(obj.Index)
        Call Writer.WriteInt16(ArmasHerrero(validIndexes(I)))
        Call Writer.WriteInt16(obj.LingH)
        Call Writer.WriteInt16(obj.LingP)
        Call Writer.WriteInt16(obj.LingO)
    Next I

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer

    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    Call Writer.WriteInt(ServerPacketID.BlacksmithArmors)

    For I = 1 To UBound(ArmadurasHerrero())

        ' Can the user create this object? If so add it to the list....
        If ObjData(ArmadurasHerrero(I)).SkHerreria <= Round(UserList( _
                UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreria(UserList( _
                UserIndex).clase), 0) Then
            Count = Count + 1
            validIndexes(Count) = I
        End If

    Next I

    ' Write the number of objects in the list
    Call Writer.WriteInt16(Count)

    ' Write the needed data of each object
    For I = 1 To Count
        obj = ObjData(ArmadurasHerrero(validIndexes(I)))
        Call Writer.WriteString8(obj.Name)
        Call Writer.WriteInt16(obj.LingH)
        Call Writer.WriteInt16(obj.LingP)
        Call Writer.WriteInt16(obj.LingO)
        Call Writer.WriteInt16(ArmadurasHerrero(validIndexes(I)))
    Next I

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)

    Dim i              As Long

    Dim validIndexes() As Integer

    Dim Count          As Byte

    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    Call Writer.WriteInt(ServerPacketID.CarpenterObjects)

    For I = 1 To UBound(ObjCarpintero())

        ' Can the user create this object? If so add it to the list....
        If ObjData(ObjCarpintero(I)).SkCarpinteria <= UserList( _
                UserIndex).Stats.UserSkills(eSkill.Carpinteria) Then

            If I = 1 Then Debug.Print UserList(UserIndex).Stats.UserSkills( _
                    eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase)
            Count = Count + 1
            validIndexes(Count) = I
        End If

    Next I

    ' Write the number of objects in the list
    Call Writer.WriteInt8(Count)

    ' Write the needed data of each object
    For I = 1 To Count
        Call Writer.WriteInt16(ObjCarpintero(validIndexes(I)))
        'Call Writer.WriteInt16(obj.Madera)
        'Call Writer.WriteInt32(obj.GrhIndex)
        ' Ladder 07/07/2014   Ahora se envia el grafico de los objetos
    Next I

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteAlquimistaObjects(ByVal UserIndex As Integer)

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer

    ReDim validIndexes(1 To UBound(ObjAlquimista()))
    Call Writer.WriteInt(ServerPacketID.AlquimistaObj)

    For I = 1 To UBound(ObjAlquimista())

        ' Can the user create this object? If so add it to the list....
        If ObjData(ObjAlquimista(I)).SkPociones <= UserList(UserIndex).Stats.UserSkills( _
                eSkill.Alquimia) \ ModAlquimia(UserList(UserIndex).clase) Then
            Count = Count + 1
            validIndexes(Count) = I
        End If

    Next I

    ' Write the number of objects in the list
    Call Writer.WriteInt16(Count)

    ' Write the needed data of each object
    For I = 1 To Count
        Call Writer.WriteInt16(ObjAlquimista(validIndexes(I)))
    Next I

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteSastreObjects(ByVal UserIndex As Integer)

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer

    ReDim validIndexes(1 To UBound(ObjSastre()))
    Call Writer.WriteInt(ServerPacketID.SastreObj)

    For I = 1 To UBound(ObjSastre())

        ' Can the user create this object? If so add it to the list....
        If ObjData(ObjSastre(I)).SkMAGOria <= UserList(UserIndex).Stats.UserSkills( _
                eSkill.Sastreria) Then
            Count = Count + 1
            validIndexes(Count) = I
        End If

    Next I

    ' Write the number of objects in the list
    Call Writer.WriteInt16(Count)

    ' Write the needed data of each object
    For I = 1 To Count
        Call Writer.WriteInt16(ObjSastre(validIndexes(I)))
    Next I

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRestOK(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.RestOK)
    Call modSendData.SendData(ToIndex, UserIndex)
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageErrorMsg(Message))
End Sub

''
' Writes the "Blind" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlind(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.Blind)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDumb(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.Dumb)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
'Optimizacion de protocolo por Ladder
Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowSignal)
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt16(ObjData(ObjIndex).GrhSecundario)
    Call modSendData.SendData(ToIndex, UserIndex)
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
                                       ByRef obj As obj, _
                                       ByVal price As Single)

    Dim ObjInfo     As ObjData

    Dim PodraUsarlo As Byte

    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(obj.ObjIndex)
        PodraUsarlo = PuedeUsarObjeto(UserIndex, obj.ObjIndex)
    End If

    Call Writer.WriteInt(ServerPacketID.ChangeNPCInventorySlot)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(obj.ObjIndex)
    Call Writer.WriteInt16(obj.amount)
    Call Writer.WriteReal32(price)
    Call Writer.WriteInt8(PodraUsarlo)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.UpdateHungerAndThirst)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxAGU)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MinAGU)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MaxHam)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.MinHam)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteLight(ByVal UserIndex As Integer, ByVal Map As Integer)
    Call Writer.WriteInt(ServerPacketID.light)
    Call Writer.WriteString8(MapInfo(Map).base_light)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteFlashScreen(ByVal UserIndex As Integer, _
                            ByVal Color As Long, _
                            ByVal Time As Long, _
                            Optional ByVal Ignorar As Boolean = False)
    Call Writer.WriteInt(ServerPacketID.FlashScreen)
    Call Writer.WriteInt32(Color)
    Call Writer.WriteInt32(Time)
    Call Writer.WriteBool(Ignorar)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteFYA(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.FYA)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(1))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(2))
    Call Writer.WriteInt16(UserList(UserIndex).flags.DuracionEfecto)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteCerrarleCliente(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.CerrarleCliente)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteOxigeno(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.Oxigeno)
    Call Writer.WriteInt16(UserList(UserIndex).Counters.Oxigeno)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteContadores(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.Contadores)
    Call Writer.WriteInt16(UserList(UserIndex).Counters.Invisibilidad)
    Call Writer.WriteInt16(UserList(UserIndex).Counters.ScrollExperiencia)
    Call Writer.WriteInt16(UserList(UserIndex).Counters.ScrollOro)

    If UserList(UserIndex).flags.NecesitaOxigeno Then
        Call Writer.WriteInt16(UserList(UserIndex).Counters.Oxigeno)
    Else
        Call Writer.WriteInt16(0)
    End If

    Call Writer.WriteInt16(UserList(UserIndex).flags.DuracionEfecto)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteBindKeys(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.BindKeys)
    Call Writer.WriteInt8(UserList(UserIndex).ChatCombate)
    Call Writer.WriteInt8(UserList(UserIndex).ChatGlobal)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMiniStats(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.MiniStats)
    Call Writer.WriteInt32(UserList(UserIndex).Faccion.ciudadanosMatados)
    Call Writer.WriteInt32(UserList(UserIndex).Faccion.CriminalesMatados)
    Call Writer.WriteInt8(UserList(UserIndex).Faccion.Status)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.NPCsMuertos)
    Call Writer.WriteInt8(UserList(UserIndex).clase)
    Call Writer.WriteInt32(UserList(UserIndex).Counters.Pena)
    Call Writer.WriteInt32(UserList(UserIndex).flags.VecesQueMoriste)
    Call Writer.WriteInt8(UserList(UserIndex).genero)
    Call Writer.WriteInt8(UserList(UserIndex).raza)
    Call Writer.WriteInt8(UserList(UserIndex).donador.activo)
    Call Writer.WriteInt32(CreditosDonadorCheck(UserList(UserIndex).Cuenta))
    'ARREGLANDO
    Call Writer.WriteInt16(DiasDonadorCheck(UserList(UserIndex).Cuenta))
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data .incomingData.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
    Call Writer.WriteInt(ServerPacketID.LevelUp)
    Call Writer.WriteInt16(skillPoints)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data .incomingData.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, _
                            ByVal title As String, _
                            ByVal Message As String)
    Call Writer.WriteInt(ServerPacketID.AddForumMsg)
    Call Writer.WriteString8(title)
    Call Writer.WriteString8(Message)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowForumForm)
    Call modSendData.SendData(ToIndex, UserIndex)
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
    Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageSetInvisible(CharIndex, _
            invisible))
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
    Call Writer.WriteInt(ServerPacketID.DiceRoll)
    ' TODO: SACAR ESTE PAQUETE USAR EL DE ATRIBUTOS
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
            eAtributos.Inteligencia))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos( _
            eAtributos.Constitucion))
    Call Writer.WriteInt8(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.MeditateToggle)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.BlindNoMore)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.DumbNoMore)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSendSkills(ByVal UserIndex As Integer)

    Dim i As Long

    Call Writer.WriteInt(ServerPacketID.SendSkills)

    For I = 1 To NUMSKILLS
        Call Writer.WriteInt8(UserList(UserIndex).Stats.UserSkills(I))
    Next I

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    Dim i   As Long

    Dim str As String

    Call Writer.WriteInt(ServerPacketID.TrainerCreatureList)

    For I = 1 To NpcList(NpcIndex).NroCriaturas
        str = str & NpcList(NpcIndex).Criaturas(I).NpcName & SEPARATOR
    Next I

    If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
    Call Writer.WriteString8(str)
    Call modSendData.SendData(ToIndex, UserIndex)
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

    Dim i   As Long

    Dim Tmp As String

    Call Writer.WriteInt(ServerPacketID.guildNews)
    Call Writer.WriteString8(guildNews)

    ' Prepare guild name's list
    For I = LBound(guildList()) To UBound(guildList())
        Tmp = Tmp & guildList(I) & SEPARATOR
    Next I

    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    ' Prepare guild member's list
    Tmp = vbNullString

    For I = LBound(MemberList()) To UBound(MemberList())
        Tmp = Tmp & MemberList(I) & SEPARATOR
    Next I

    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call Writer.WriteInt8(ClanNivel)
    Call Writer.WriteInt16(ExpAcu)
    Call Writer.WriteInt16(ExpNe)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)

    Dim i As Long

    Call Writer.WriteInt(ServerPacketID.OfferDetails)
    Call Writer.WriteString8(details)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    Dim i   As Long

    Dim Tmp As String

    Call Writer.WriteInt(ServerPacketID.AlianceProposalsList)

    ' Prepare guild's list
    For I = LBound(guilds()) To UBound(guilds())
        Tmp = Tmp & guilds(I) & SEPARATOR
    Next I

    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    Dim i   As Long

    Dim Tmp As String

    Call Writer.WriteInt(ServerPacketID.PeaceProposalsList)

    ' Prepare guilds' list
    For I = LBound(guilds()) To UBound(guilds())
        Tmp = Tmp & guilds(I) & SEPARATOR
    Next I

    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
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
        ByVal race As eRaza, ByVal Class As eClass, ByVal gender As eGenero, ByVal _
        level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal previousPetitions _
        As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal _
        RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As _
        Long, ByVal criminalsKilled As Long)
    Call Writer.WriteInt(ServerPacketID.CharacterInfo)
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

    Dim i   As Long

    Dim Tmp As String

    Call Writer.WriteInt(ServerPacketID.GuildLeaderInfo)

    ' Prepare guild name's list
    For I = LBound(guildList()) To UBound(guildList())
        Tmp = Tmp & guildList(I) & SEPARATOR
    Next I

    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    ' Prepare guild member's list
    Tmp = vbNullString

    For I = LBound(MemberList()) To UBound(MemberList())
        Tmp = Tmp & MemberList(I) & SEPARATOR
    Next I

    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    ' Store guild news
    Call Writer.WriteString8(guildNews)
    ' Prepare the join request's list
    Tmp = vbNullString

    For I = LBound(joinRequests()) To UBound(joinRequests())
        Tmp = Tmp & joinRequests(I) & SEPARATOR
    Next I

    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call Writer.WriteInt8(NivelDeClan)
    Call Writer.WriteInt16(ExpActual)
    Call Writer.WriteInt16(ExpNecesaria)
    Call modSendData.SendData(ToIndex, UserIndex)
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
                             ByVal NivelDeClan As Byte, _
                             ByVal ExpActual As Integer, _
                             ByVal ExpNecesaria As Integer)
    Call Writer.WriteInt(ServerPacketID.GuildDetails)
    Call Writer.WriteString8(GuildName)
    Call Writer.WriteString8(founder)
    Call Writer.WriteString8(foundationDate)
    Call Writer.WriteString8(leader)
    Call Writer.WriteInt16(memberCount)
    Call Writer.WriteString8(alignment)
    Call Writer.WriteString8(guildDesc)
    Call Writer.WriteInt8(NivelDeClan)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowGuildFundationForm)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteParalizeOK(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ParalizeOK)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteInmovilizaOK(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.InmovilizadoOK)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteStopped(ByVal UserIndex As Integer, ByVal Stopped As Boolean)
    Call Writer.WriteInt(ServerPacketID.Stopped)
    Call Writer.WriteBool(Stopped)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
    Call Writer.WriteInt(ServerPacketID.ShowUserRequest)
    Call Writer.WriteString8(details)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    Amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, _
                                    ByRef itemsAenviar() As obj, _
                                    ByVal gold As Long, _
                                    ByVal miOferta As Boolean)
    Call Writer.WriteInt(ServerPacketID.ChangeUserTradeSlot)
    Call Writer.WriteBool(miOferta)
    Call Writer.WriteInt32(gold)

    Dim I As Long

    For I = 1 To UBound(itemsAenviar)
        Call Writer.WriteInt16(itemsAenviar(I).ObjIndex)

        If itemsAenviar(I).ObjIndex = 0 Then
            Call Writer.WriteString8("")
        Else
            Call Writer.WriteString8(ObjData(itemsAenviar(I).ObjIndex).Name)
        End If

        If itemsAenviar(I).ObjIndex = 0 Then
            Call Writer.WriteInt32(0)
        Else
            Call Writer.WriteInt32(ObjData(itemsAenviar(I).ObjIndex).GrhIndex)
        End If

        Call Writer.WriteInt32(itemsAenviar(I).amount)
    Next I

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByVal ListaCompleta As Boolean)
    Call Writer.WriteInt(ServerPacketID.SpawnListt)
    Call Writer.WriteBool(ListaCompleta)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)

    Dim i   As Long

    Dim Tmp As String

    Call Writer.WriteInt(ServerPacketID.ShowSOSForm)

    For I = 1 To Ayuda.Longitud
        Tmp = Tmp & Ayuda.VerElemento(I) & SEPARATOR
    Next I

    If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, _
                                    ByVal currentMOTD As String)
    Call Writer.WriteInt(ServerPacketID.ShowMOTDEditionForm)
    Call Writer.WriteString8(currentMOTD)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowGMPanelForm)
    Call Writer.WriteInt16(UserList(UserIndex).Char.Head)
    Call Writer.WriteInt16(UserList(UserIndex).Char.Body)
    Call Writer.WriteInt16(UserList(UserIndex).Char.CascoAnim)
    Call Writer.WriteInt16(UserList(UserIndex).Char.WeaponAnim)
    Call Writer.WriteInt16(UserList(UserIndex).Char.ShieldAnim)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteShowFundarClanForm(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowFundarClanForm)
    Call modSendData.SendData(ToIndex, UserIndex)
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

    Dim i   As Long

    Dim Tmp As String

    Call Writer.WriteInt(ServerPacketID.UserNameList)

    ' Prepare user's names list
    For I = 1 To cant
        Tmp = Tmp & userNamesList(I) & SEPARATOR
    Next I

    If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
    Call Writer.WriteString8(Tmp)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

''
' Writes the "Pong" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePong(ByVal UserIndex As Integer, ByVal Time As Long)
    Call Writer.WriteInt(ServerPacketID.Pong)
    Call Writer.WriteInt32(Time)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WritePersonajesDeCuenta(ByVal UserIndex As Integer)

    Dim UserCuenta                     As String

    Dim CantPersonajes                 As Byte

    Dim Personaje(1 To MAX_PERSONAJES) As PersonajeCuenta

    Dim donador                        As Boolean

    Dim i                              As Byte

    UserCuenta = UserList(UserIndex).Cuenta
    donador = DonadorCheck(UserCuenta)

    If Database_Enabled Then
        CantPersonajes = GetPersonajesCuentaDatabase(UserList(UserIndex).AccountID, _
                Personaje)
    Else
        CantPersonajes = ObtenerCantidadDePersonajes(UserCuenta)

        For i = 1 To CantPersonajes
            Personaje(i).nombre = ObtenerNombrePJ(UserCuenta, i)
            Personaje(i).Cabeza = ObtenerCabeza(Personaje(i).nombre)
            Personaje(i).clase = ObtenerClase(Personaje(i).nombre)
            Personaje(i).cuerpo = ObtenerCuerpo(Personaje(i).nombre)
            Personaje(i).Mapa = ReadField(1, ObtenerMapa(Personaje(i).nombre), Asc("-"))
            Personaje(i).nivel = ObtenerNivel(Personaje(i).nombre)
            Personaje(i).Status = ObtenerCriminal(Personaje(i).nombre)
            Personaje(i).Casco = ObtenerCasco(Personaje(i).nombre)
            Personaje(i).Escudo = ObtenerEscudo(Personaje(i).nombre)
            Personaje(i).Arma = ObtenerArma(Personaje(i).nombre)
            Personaje(i).ClanIndex = GetUserGuildIndexCharfile(Personaje(i).nombre)
        Next i

    End If

    Call Writer.WriteInt(ServerPacketID.PersonajesDeCuenta)
    Call Writer.WriteInt8(CantPersonajes)

    For I = 1 To CantPersonajes
        Call Writer.WriteString8(Personaje(I).nombre)
        Call Writer.WriteInt8(Personaje(I).nivel)
        Call Writer.WriteInt16(Personaje(I).Mapa)
        Call Writer.WriteInt16(Personaje(I).posX)
        Call Writer.WriteInt16(Personaje(I).posY)
        Call Writer.WriteInt16(Personaje(I).cuerpo)
        Call Writer.WriteInt16(Personaje(I).Cabeza)
        Call Writer.WriteInt8(Personaje(I).Status)
        Call Writer.WriteInt8(Personaje(I).clase)
        Call Writer.WriteInt16(Personaje(I).Casco)
        Call Writer.WriteInt16(Personaje(I).Escudo)
        Call Writer.WriteInt16(Personaje(I).Arma)
        Call Writer.WriteString8(modGuilds.GuildName(Personaje(I).ClanIndex))
    Next I

    Call Writer.WriteInt8(IIf(donador, 1, 0))
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteGoliathInit(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.Goliath)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.Banco)
    Call Writer.WriteInt8(UserList(UserIndex).BancoInvent.NroItems)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteShowFrmLogear(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowFrmLogear)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteShowFrmMapa(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ShowFrmMapa)

    If UserList(UserIndex).donador.activo = 1 Then
        Call Writer.WriteInt16(ExpMult * UserList(UserIndex).flags.ScrollExp * 1.1)
    Else
        Call Writer.WriteInt16(ExpMult * UserList(UserIndex).flags.ScrollExp)
    End If

    Call Writer.WriteInt16(OroMult * UserList(UserIndex).flags.ScrollOro)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteFamiliar(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        Call Writer.WriteInt(ServerPacketID.Familiar)
        Call Writer.WriteInt8(.Familiar.Existe)
        Call Writer.WriteInt8(.Familiar.Muerto)
        Call Writer.WriteString8(.Familiar.nombre)
        Call Writer.WriteInt32(.Familiar.Exp)
        Call Writer.WriteInt32(.Familiar.ELU)
        Call Writer.WriteInt8(.Familiar.nivel)
        Call Writer.WriteInt16(.Familiar.MinHp)
        Call Writer.WriteInt16(.Familiar.MaxHp)
        Call Writer.WriteInt16(.Familiar.MinHIT)
        Call Writer.WriteInt16(.Familiar.MaxHit)
    End With

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteRecompensas(ByVal UserIndex As Integer)

    On Error GoTo 0

    Dim a, b, c As Byte

    b = UserList(UserIndex).UserLogros + 1
    a = UserList(UserIndex).NPcLogros + 1
    c = UserList(UserIndex).LevelLogros + 1
    Call Writer.WriteInt(ServerPacketID.Logros)
    'Logros NPC
    Call Writer.WriteString8(NPcLogros(a).nombre)
    Call Writer.WriteString8(NPcLogros(a).Desc)
    Call Writer.WriteInt16(NPcLogros(a).cant)
    Call Writer.WriteInt8(NPcLogros(a).TipoRecompensa)

    If NPcLogros(a).TipoRecompensa = 1 Then
        Call Writer.WriteString8(NPcLogros(a).ObjRecompensa)
    End If

    If NPcLogros(a).TipoRecompensa = 2 Then
        Call Writer.WriteInt32(NPcLogros(a).OroRecompensa)
    End If

    If NPcLogros(a).TipoRecompensa = 3 Then
        Call Writer.WriteInt32(NPcLogros(a).ExpRecompensa)
    End If

    If NPcLogros(a).TipoRecompensa = 4 Then
        Call Writer.WriteInt8(NPcLogros(a).HechizoRecompensa)
    End If

    Call Writer.WriteInt16(UserList(UserIndex).Stats.NPCsMuertos)

    If UserList(UserIndex).Stats.NPCsMuertos >= NPcLogros(a).cant Then
        Call Writer.WriteBool(True)
    Else
        Call Writer.WriteBool(False)
    End If

    'Logros User
    Call Writer.WriteString8(UserLogros(b).nombre)
    Call Writer.WriteString8(UserLogros(b).Desc)
    Call Writer.WriteInt16(UserLogros(b).cant)
    Call Writer.WriteInt16(UserLogros(b).TipoRecompensa)
    Call Writer.WriteInt16(UserList(UserIndex).Stats.UsuariosMatados)

    If UserLogros(a).TipoRecompensa = 1 Then
        Call Writer.WriteString8(UserLogros(b).ObjRecompensa)
    End If

    If UserLogros(a).TipoRecompensa = 2 Then
        Call Writer.WriteInt32(UserLogros(b).OroRecompensa)
    End If

    If UserLogros(a).TipoRecompensa = 3 Then
        Call Writer.WriteInt32(UserLogros(b).ExpRecompensa)
    End If

    If UserLogros(a).TipoRecompensa = 4 Then
        Call Writer.WriteInt8(UserLogros(b).HechizoRecompensa)
    End If

    If UserList(UserIndex).Stats.UsuariosMatados >= UserLogros(b).cant Then
        Call Writer.WriteBool(True)
    Else
        Call Writer.WriteBool(False)
    End If

    'Nivel User
    Call Writer.WriteString8(LevelLogros(c).nombre)
    Call Writer.WriteString8(LevelLogros(c).Desc)
    Call Writer.WriteInt16(LevelLogros(c).cant)
    Call Writer.WriteInt16(LevelLogros(c).TipoRecompensa)
    Call Writer.WriteInt8(UserList(UserIndex).Stats.ELV)

    If LevelLogros(c).TipoRecompensa = 1 Then
        Call Writer.WriteString8(LevelLogros(c).ObjRecompensa)
    End If

    If LevelLogros(c).TipoRecompensa = 2 Then
        Call Writer.WriteInt32(LevelLogros(c).OroRecompensa)
    End If

    If LevelLogros(c).TipoRecompensa = 3 Then
        Call Writer.WriteInt32(LevelLogros(c).ExpRecompensa)
    End If

    If LevelLogros(c).TipoRecompensa = 4 Then
        Call Writer.WriteInt8(LevelLogros(c).HechizoRecompensa)
    End If

    If UserList(UserIndex).Stats.ELV >= LevelLogros(c).cant Then
        Call Writer.WriteBool(True)
    Else
        Call Writer.WriteBool(False)
    End If

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WritePreguntaBox(ByVal UserIndex As Integer, ByVal message As String)
    Call Writer.WriteInt(ServerPacketID.ShowPregunta)
    Call Writer.WriteString8(Message)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteDatosGrupo(ByVal UserIndex As Integer)

    Dim i As Byte

    With UserList(UserIndex)
        Call Writer.WriteInt(ServerPacketID.DatosGrupo)
        Call Writer.WriteBool(.Grupo.EnGrupo)

        If .Grupo.EnGrupo = True Then
            Call Writer.WriteInt8(UserList(.Grupo.Lider).Grupo.CantidadMiembros)

            'Call Writer.WriteInt8(UserList(.Grupo.Lider).name)
            If .Grupo.Lider = UserIndex Then

                For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros

                    If i = 1 Then
                        Call Writer.WriteString8(UserList(.Grupo.Miembros(I)).Name & _
                                "(Líder)")
                    Else
                        Call Writer.WriteString8(UserList(.Grupo.Miembros(I)).Name)
                    End If

                Next i

            Else

                For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros

                    If i = 1 Then
                        Call Writer.WriteString8(UserList(UserList( _
                                .Grupo.Lider).Grupo.Miembros(I)).Name & "(Líder)")
                    Else
                        Call Writer.WriteString8(UserList(UserList( _
                                .Grupo.Lider).Grupo.Miembros(I)).Name)
                    End If

                Next i

            End If
        End If

    End With

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteUbicacion(ByVal UserIndex As Integer, _
                          ByVal Miembro As Byte, _
                          ByVal GPS As Integer)
    Call Writer.WriteInt(ServerPacketID.ubicacion)
    Call Writer.WriteInt8(Miembro)

    If GPS > 0 Then
        Call Writer.WriteInt8(UserList(GPS).Pos.X)
        Call Writer.WriteInt8(UserList(GPS).Pos.Y)
        Call Writer.WriteInt16(UserList(GPS).Pos.Map)
    Else
        Call Writer.WriteInt8(0)
        Call Writer.WriteInt8(0)
        Call Writer.WriteInt16(0)
    End If

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteCorreoPicOn(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.CorreoPicOn)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteShop(ByVal UserIndex As Integer)

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer

    ReDim validIndexes(1 To UBound(ObjDonador()))
    Call Writer.WriteInt(ServerPacketID.DonadorObj)

    For I = 1 To UBound(ObjDonador())
        Count = Count + 1
        validIndexes(Count) = I
    Next I

    ' Write the number of objects in the list
    Call Writer.WriteInt16(Count)

    ' Write the needed data of each object
    For I = 1 To Count
        Call Writer.WriteInt16(ObjDonador(validIndexes(I)).ObjIndex)
        Call Writer.WriteInt16(ObjDonador(validIndexes(I)).Valor)
    Next I

    Call Writer.WriteInt32(CreditosDonadorCheck(UserList(UserIndex).Cuenta))
    Call Writer.WriteInt16(DiasDonadorCheck(UserList(UserIndex).Cuenta))
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteRanking(ByVal UserIndex As Integer)

    Dim i As Byte

    Call Writer.WriteInt(ServerPacketID.Ranking)

    For I = 1 To 10
        Call Writer.WriteString8(Rankings(1).user(I).Nick)
        Call Writer.WriteInt16(Rankings(1).user(I).Value)
    Next I

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteActShop(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ActShop)
    Call Writer.WriteInt32(CreditosDonadorCheck(UserList(UserIndex).Cuenta))
    Call Writer.WriteInt16(DiasDonadorCheck(UserList(UserIndex).Cuenta))
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteViajarForm(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ViajarForm)

    Dim destinos As Byte

    Dim I        As Byte

    destinos = NpcList(NpcIndex).NumDestinos
    Call Writer.WriteInt8(destinos)

    For I = 1 To destinos
        Call Writer.WriteString8(NpcList(NpcIndex).Dest(I))
    Next I

    Call Writer.WriteInt8(NpcList(NpcIndex).Interface)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteQuestDetails(ByVal UserIndex As Integer, _
                             ByVal QuestIndex As Integer, _
                             Optional QuestSlot As Byte = 0)

    Dim i As Integer

    'ID del paquete
    Call Writer.WriteInt(ServerPacketID.QuestDetails)
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
        For I = 1 To QuestList(QuestIndex).RequiredNPCs
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(I).amount)
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(I).NpcIndex)

            'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
            If QuestSlot Then
                Call Writer.WriteInt16(UserList(UserIndex).QuestStats.Quests( _
                        QuestSlot).NPCsKilled(I))
            End If

        Next I

    End If

    'Enviamos la cantidad de objs requeridos
    Call Writer.WriteInt8(QuestList(QuestIndex).RequiredOBJs)

    If QuestList(QuestIndex).RequiredOBJs Then

        'Si hay objs entonces enviamos la lista
        For I = 1 To QuestList(QuestIndex).RequiredOBJs
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(I).amount)
            Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(I).ObjIndex)
            'escribe si tiene ese objeto en el inventario y que cantidad
            Call Writer.WriteInt16(CantidadObjEnInv(UserIndex, QuestList( _
                    QuestIndex).RequiredOBJ(I).ObjIndex))
            ' Call Writer.WriteInt16(0)
        Next I

    End If

    'Enviamos la recompensa de oro y experiencia.
    Call Writer.WriteInt32((QuestList(QuestIndex).RewardGLD * OroMult))
    Call Writer.WriteInt32((QuestList(QuestIndex).RewardEXP * ExpMult))
    'Enviamos la cantidad de objs de recompensa
    Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)

    If QuestList(QuestIndex).RewardOBJs Then

        'si hay objs entonces enviamos la lista
        For I = 1 To QuestList(QuestIndex).RewardOBJs
            Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(I).amount)
            Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(I).ObjIndex)
        Next I

    End If

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub
 
Public Sub WriteQuestListSend(ByVal UserIndex As Integer)

    Dim i       As Integer

    Dim tmpStr  As String

    Dim tmpByte As Byte

    With UserList(UserIndex)
        Call Writer.WriteInt(ServerPacketID.QuestListSend)

        For i = 1 To MAXUSERQUESTS

            If .QuestStats.Quests(i).QuestIndex Then
                tmpByte = tmpByte + 1
                tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).nombre & "-"
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
End Sub

Public Sub WriteNpcQuestListSend(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    Dim I          As Integer

    Dim j          As Integer

    Dim tmpStr     As String

    Dim tmpByte    As Byte

    Dim QuestIndex As Integer

    Call Writer.WriteInt(ServerPacketID.NpcQuestListSend)
    Call Writer.WriteInt8(NpcList(NpcIndex).NumQuest) 'Escribimos primero cuantas quest tiene el NPC

    For j = 1 To NpcList(NpcIndex).NumQuest
        QuestIndex = NpcList(NpcIndex).QuestNumber(j)
        Call Writer.WriteInt16(QuestIndex)
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredLevel)
        Call Writer.WriteInt16(QuestList(QuestIndex).RequiredQuest)
        'Enviamos la cantidad de npcs requeridos
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredNPCs)

        If QuestList(QuestIndex).RequiredNPCs Then

            'Si hay npcs entonces enviamos la lista
            For I = 1 To QuestList(QuestIndex).RequiredNPCs
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(I).amount)
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredNPC(I).NpcIndex)
                'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
                'If QuestSlot Then
                ' Call Writer.WriteInt16(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))
                ' End If
            Next I

        End If

        'Enviamos la cantidad de objs requeridos
        Call Writer.WriteInt8(QuestList(QuestIndex).RequiredOBJs)

        If QuestList(QuestIndex).RequiredOBJs Then

            'Si hay objs entonces enviamos la lista
            For I = 1 To QuestList(QuestIndex).RequiredOBJs
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(I).amount)
                Call Writer.WriteInt16(QuestList(QuestIndex).RequiredOBJ(I).ObjIndex)
            Next I

        End If

        'Enviamos la recompensa de oro y experiencia.
        Call Writer.WriteInt32(QuestList(QuestIndex).RewardGLD * OroMult)
        Call Writer.WriteInt32(QuestList(QuestIndex).RewardEXP * ExpMult)
        'Enviamos la cantidad de objs de recompensa
        Call Writer.WriteInt8(QuestList(QuestIndex).RewardOBJs)

        If QuestList(QuestIndex).RewardOBJs Then

            'si hay objs entonces enviamos la lista
            For I = 1 To QuestList(QuestIndex).RewardOBJs
                Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(I).amount)
                Call Writer.WriteInt16(QuestList(QuestIndex).RewardOBJ(I).ObjIndex)
            Next I

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
                    If Not UserDoneQuest(UserIndex, QuestList( _
                            QuestIndex).RequiredQuest) Then
                        PuedeHacerla = False
                    End If
                End If

                If UserList(UserIndex).Stats.ELV < QuestList(QuestIndex).RequiredLevel _
                        Then
                    PuedeHacerla = False
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
End Sub

Public Sub WriteTolerancia0(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.Tolerancia0)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Sub WriteCommerceRecieveChatMessage(ByVal UserIndex As Integer, ByVal message As String)
    Call Writer.WriteInt(ServerPacketID.CommerceRecieveChatMessage)
    Call Writer.WriteString8(Message)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Sub WriteInvasionInfo(ByVal UserIndex As Integer, _
                      ByVal Invasion As Integer, _
                      ByVal PorcentajeVida As Byte, _
                      ByVal PorcentajeTiempo As Byte)
    Call Writer.WriteInt(ServerPacketID.InvasionInfo)
    Call Writer.WriteInt8(Invasion)
    Call Writer.WriteInt8(PorcentajeVida)
    Call Writer.WriteInt8(PorcentajeTiempo)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Sub WriteOpenCrafting(ByVal UserIndex As Integer, ByVal Tipo As Byte)
    Call Writer.WriteInt(ServerPacketID.OpenCrafting)
    Call Writer.WriteInt8(Tipo)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Sub WriteCraftingItem(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal ObjIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.CraftingItem)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(ObjIndex)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Sub WriteCraftingCatalyst(ByVal UserIndex As Integer, _
                          ByVal ObjIndex As Integer, _
                          ByVal amount As Integer, _
                          ByVal Porcentaje As Byte)
    Call Writer.WriteInt(ServerPacketID.CraftingCatalyst)
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt16(amount)
    Call Writer.WriteInt8(Porcentaje)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Sub WriteCraftingResult(ByVal UserIndex As Integer, _
                        ByVal Result As Integer, _
                        Optional ByVal Porcentaje As Byte, _
                        Optional ByVal Precio As Long)
    Call Writer.WriteInt(ServerPacketID.CraftingResult)
    Call Writer.WriteInt16(Result)

    If Result <> 0 Then
        Call Writer.WriteInt8(Porcentaje)
        Call Writer.WriteInt32(Precio)
    End If

    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Sub WriteForceUpdate(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ForceUpdate)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteUpdateNPCSimbolo(ByVal UserIndex As Integer, _
                                 ByVal NpcIndex As Integer, _
                                 ByVal Simbolo As Byte)
    Call Writer.WriteInt(ServerPacketID.UpdateNPCSimbolo)
    Call Writer.WriteInt16(NpcList(NpcIndex).Char.CharIndex)
    Call Writer.WriteInt8(Simbolo)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

Public Sub WriteGuardNotice(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.GuardNotice)
    Call modSendData.SendData(ToIndex, UserIndex)
End Sub

' \Begin: [Prepares]
Public Function PrepareMessageCharSwing(ByVal CharIndex As Integer, _
                                        Optional ByVal FX As Boolean = True, _
                                        Optional ByVal ShowText As Boolean = True)
    Call Writer.WriteInt(ServerPacketID.CharSwing)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteBool(FX)
    Call Writer.WriteBool(ShowText)
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
    Call Writer.WriteInt(ServerPacketID.SetInvisible)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteBool(invisible)
End Function

Public Function PrepareMessageSetEscribiendo(ByVal CharIndex As Integer, _
                                             ByVal Escribiendo As Boolean)
    Call Writer.WriteInt(ServerPacketID.SetEscribiendo)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteBool(Escribiendo)
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
                                           ByVal Color As Long, _
                                           Optional ByVal Name As String = "")

    Dim R, g, b As Byte

    b = (Color And 16711680) / 65536
    g = (Color And 65280) / 256
    R = Color And 255
    Call Writer.WriteInt(ServerPacketID.ChatOverHead)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt16(CharIndex)
    ' Write rgb channels and save one byte from long :D
    Call Writer.WriteInt8(R)
    Call Writer.WriteInt8(g)
    Call Writer.WriteInt8(b)
    Call Writer.WriteInt32(Color)
End Function

Public Function PrepareMessageTextOverChar(ByVal chat As String, _
                                           ByVal CharIndex As Integer, _
                                           ByVal Color As Long)
    Call Writer.WriteInt(ServerPacketID.TextOverChar)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt32(Color)
End Function

Public Function PrepareMessageTextCharDrop(ByVal chat As String, _
                                           ByVal CharIndex As Integer, _
                                           ByVal Color As Long)
    Call Writer.WriteInt(ServerPacketID.TextCharDrop)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt32(Color)
End Function

Public Function PrepareMessageTextOverTile(ByVal chat As String, _
                                           ByVal X As Integer, _
                                           ByVal Y As Integer, _
                                           ByVal Color As Long)
    Call Writer.WriteInt(ServerPacketID.TextOverTile)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt16(X)
    Call Writer.WriteInt16(Y)
    Call Writer.WriteInt32(Color)
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageConsoleMsg(ByVal chat As String, _
                                         ByVal FontIndex As FontTypeNames)
    Call Writer.WriteInt(ServerPacketID.ConsoleMsg)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt8(FontIndex)
End Function

Public Function PrepareMessageLocaleMsg(ByVal ID As Integer, _
                                        ByVal chat As String, _
                                        ByVal FontIndex As FontTypeNames)
    Call Writer.WriteInt(ServerPacketID.LocaleMsg)
    Call Writer.WriteInt16(ID)
    Call Writer.WriteString8(chat)
    Call Writer.WriteInt8(FontIndex)
End Function

Public Function PrepareMessageListaCorreo(ByVal UserIndex As Integer, _
                                          ByVal Actualizar As Boolean)

    Dim cant As Byte

    Dim I    As Byte

    cant = UserList(UserIndex).Correo.CantCorreo
    UserList(UserIndex).Correo.NoLeidos = 0
    Call Writer.WriteInt(ServerPacketID.ListaCorreo)
    Call Writer.WriteInt8(cant)

    If cant > 0 Then

        For I = 1 To cant
            Call Writer.WriteString8(UserList(UserIndex).Correo.Mensaje(I).Remitente)
            Call Writer.WriteString8(UserList(UserIndex).Correo.Mensaje(I).Mensaje)
            Call Writer.WriteInt8(UserList(UserIndex).Correo.Mensaje(I).ItemCount)
            Call Writer.WriteString8(UserList(UserIndex).Correo.Mensaje(I).Item)
            Call Writer.WriteInt8(UserList(UserIndex).Correo.Mensaje(I).Leido)
            Call Writer.WriteString8(UserList(UserIndex).Correo.Mensaje(I).Fecha)
            'Call ReadMessageCorreo(UserIndex, i)
        Next I

    End If

    Call Writer.WriteBool(Actualizar)
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
    Call Writer.WriteInt(ServerPacketID.CreateFX)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt16(FXLoops)
End Function

Public Function PrepareMessageMeditateToggle(ByVal CharIndex As Integer, _
                                             ByVal FX As Integer)
    Call Writer.WriteInt(ServerPacketID.MeditateToggle)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt16(FX)
End Function

Public Function PrepareMessageParticleFX(ByVal CharIndex As Integer, _
                                         ByVal Particula As Integer, _
                                         ByVal Time As Long, _
                                         ByVal Remove As Boolean, _
                                         Optional ByVal grh As Long = 0)
    Call Writer.WriteInt(ServerPacketID.ParticleFX)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt16(Particula)
    Call Writer.WriteInt32(Time)
    Call Writer.WriteBool(Remove)
    Call Writer.WriteInt32(grh)
End Function

Public Function PrepareMessageParticleFXWithDestino(ByVal Emisor As Integer, _
                                                    ByVal Receptor As Integer, _
                                                    ByVal ParticulaViaje As Integer, _
                                                    ByVal ParticulaFinal As Integer, _
                                                    ByVal Time As Long, _
                                                    ByVal wav As Integer, _
                                                    ByVal FX As Integer)
    Call Writer.WriteInt(ServerPacketID.ParticleFXWithDestino)
    Call Writer.WriteInt16(Emisor)
    Call Writer.WriteInt16(Receptor)
    Call Writer.WriteInt16(ParticulaViaje)
    Call Writer.WriteInt16(ParticulaFinal)
    Call Writer.WriteInt32(Time)
    Call Writer.WriteInt16(wav)
    Call Writer.WriteInt16(FX)
End Function

Public Function PrepareMessageParticleFXWithDestinoXY(ByVal Emisor As Integer, _
                                                      ByVal ParticulaViaje As Integer, _
                                                      ByVal ParticulaFinal As Integer, _
                                                      ByVal Time As Long, _
                                                      ByVal wav As Integer, _
                                                      ByVal FX As Integer, _
                                                      ByVal X As Byte, _
                                                      ByVal Y As Byte)
    Call Writer.WriteInt(ServerPacketID.ParticleFXWithDestinoXY)
    Call Writer.WriteInt16(Emisor)
    Call Writer.WriteInt16(ParticulaViaje)
    Call Writer.WriteInt16(ParticulaFinal)
    Call Writer.WriteInt32(Time)
    Call Writer.WriteInt16(wav)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
End Function

Public Function PrepareMessageAuraToChar(ByVal CharIndex As Integer, _
                                         ByVal Aura As String, _
                                         ByVal Remove As Boolean, _
                                         ByVal Tipo As Byte)
    Call Writer.WriteInt(ServerPacketID.AuraToChar)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteString8(Aura)
    Call Writer.WriteBool(Remove)
    Call Writer.WriteInt8(Tipo)
End Function

Public Function PrepareMessageSpeedingACT(ByVal CharIndex As Integer, _
                                          ByVal speeding As Single)
    Call Writer.WriteInt(ServerPacketID.SpeedToChar)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteReal32(speeding)
End Function

Public Function PrepareMessageParticleFXToFloor(ByVal X As Byte, _
                                                ByVal Y As Byte, _
                                                ByVal Particula As Integer, _
                                                ByVal Time As Long)
    Call Writer.WriteInt(ServerPacketID.ParticleFXToFloor)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteInt16(Particula)
    Call Writer.WriteInt32(Time)
End Function

Public Function PrepareMessageLightFXToFloor(ByVal X As Byte, _
                                             ByVal Y As Byte, _
                                             ByVal LuzColor As Long, _
                                             ByVal Rango As Byte)
    Call Writer.WriteInt(ServerPacketID.LightToFloor)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteInt32(LuzColor)
    Call Writer.WriteInt8(Rango)
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
    Call Writer.WriteInt(ServerPacketID.PlayWave)
    Call Writer.WriteInt16(wave)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
End Function

Public Function PrepareMessageUbicacionLlamada(ByVal Mapa As Integer, _
                                               ByVal X As Byte, _
                                               ByVal Y As Byte)
    Call Writer.WriteInt(ServerPacketID.PosLLamadaDeClan)
    Call Writer.WriteInt16(Mapa)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
End Function

Public Function PrepareMessageCharUpdateHP(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.CharUpdateHP)
    Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.MinHp)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.MaxHp)
End Function

Public Function PrepareMessageCharUpdateMAN(ByVal UserIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.CharUpdateMAN)
    Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.MinMAN)
    Call Writer.WriteInt32(UserList(UserIndex).Stats.MaxMAN)
End Function

Public Function PrepareMessageNpcUpdateHP(ByVal NpcIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.CharUpdateHP)
    Call Writer.WriteInt16(NpcList(NpcIndex).Char.CharIndex)
    Call Writer.WriteInt32(NpcList(NpcIndex).Stats.MinHp)
    Call Writer.WriteInt32(NpcList(NpcIndex).Stats.MaxHp)
End Function

Public Function PrepareMessageArmaMov(ByVal CharIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.ArmaMov)
    Call Writer.WriteInt16(CharIndex)
End Function

Public Function PrepareMessageEscudoMov(ByVal CharIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.EscudoMov)
    Call Writer.WriteInt16(CharIndex)
End Function

Public Function PrepareMessageFlashScreen(ByVal Color As Long, _
                                          ByVal Duracion As Long, _
                                          Optional ByVal Ignorar As Boolean = False)
    Call Writer.WriteInt(ServerPacketID.FlashScreen)
    Call Writer.WriteInt32(Color)
    Call Writer.WriteInt32(Duracion)
    Call Writer.WriteBool(Ignorar)
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageGuildChat(ByVal chat As String, ByVal Status As Byte)
    Call Writer.WriteInt(ServerPacketID.GuildChat)
    Call Writer.WriteInt8(Status)
    Call Writer.WriteString8(chat)
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageShowMessageBox(ByVal chat As String)
    Call Writer.WriteInt(ServerPacketID.ShowMessageBox)
    Call Writer.WriteString8(chat)
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
    Call Writer.WriteInt(ServerPacketID.PlayMIDI)
    Call Writer.WriteInt8(midi)
    Call Writer.WriteInt16(loops)
End Function

Public Function PrepareMessageOnlineUser(ByVal UserOnline As Integer)
    Call Writer.WriteInt(ServerPacketID.UserOnline)
    Call Writer.WriteInt16(UserOnline)
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessagePauseToggle()
    Call Writer.WriteInt(ServerPacketID.PauseToggle)
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageRainToggle()
    Call Writer.WriteInt(ServerPacketID.RainToggle)
End Function

Public Function PrepareMessageTrofeoToggleOn()
    Call Writer.WriteInt(ServerPacketID.TrofeoToggleOn)
End Function

Public Function PrepareMessageTrofeoToggleOff()
    Call Writer.WriteInt(ServerPacketID.TrofeoToggleOff)
End Function

Public Function PrepareMessageHora()
    Call Writer.WriteInt(ServerPacketID.Hora)
    Call Writer.WriteInt32((GetTickCount() - HoraMundo) Mod DuracionDia)
    Call Writer.WriteInt32(DuracionDia)
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte)
    Call Writer.WriteInt(ServerPacketID.ObjectDelete)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageBlockPosition(ByVal X As Byte, _
                                            ByVal Y As Byte, _
                                            ByVal Blocked As Byte)
    Call Writer.WriteInt(ServerPacketID.BlockPosition)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteInt8(Blocked)
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
    Call Writer.WriteInt(ServerPacketID.ObjectCreate)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteInt16(ObjIndex)
    Call Writer.WriteInt16(amount)
End Function

Public Function PrepareMessageFxPiso(ByVal GrhIndex As Integer, _
                                     ByVal X As Byte, _
                                     ByVal Y As Byte)
    Call Writer.WriteInt(ServerPacketID.fxpiso)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteInt16(GrhIndex)
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer, _
                                              ByVal Desvanecido As Boolean)
    Call Writer.WriteInt(ServerPacketID.CharacterRemove)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteBool(Desvanecido)
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer)
    Call Writer.WriteInt(ServerPacketID.RemoveCharDialog)
    Call Writer.WriteInt16(CharIndex)
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
                                              ByVal Heading As eHeading, _
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
                                              ByVal donador As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal UserMinMAN As Long, ByVal UserMaxMAN As Long, ByVal Simbolo As Byte, ByVal Idle As Boolean, ByVal Navegando As Boolean)
    Call Writer.WriteInt(ServerPacketID.CharacterCreate)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt16(Body)
    Call Writer.WriteInt16(Head)
    Call Writer.WriteInt8(Heading)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
    Call Writer.WriteInt16(weapon)
    Call Writer.WriteInt16(shield)
    Call Writer.WriteInt16(helmet)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt16(FXLoops)
    Call Writer.WriteString8(Name)
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
    Call Writer.WriteInt8(donador)
    Call Writer.WriteInt8(appear)
    Call Writer.WriteInt16(group_index)
    Call Writer.WriteInt16(clan_index)
    Call Writer.WriteInt8(clan_nivel)
    Call Writer.WriteInt32(UserMinHp)
    Call Writer.WriteInt32(UserMaxHp)
    Call Writer.WriteInt32(UserMinMAN)
    Call Writer.WriteInt32(UserMaxMAN)
    Call Writer.WriteInt8(Simbolo)
    Call Writer.WriteBool(Idle)
    Call Writer.WriteBool(Navegando)
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
                                              ByVal Heading As eHeading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal Idle As Boolean, _
                                              ByVal Navegando As Boolean)
    Call Writer.WriteInt(ServerPacketID.CharacterChange)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt16(Body)
    Call Writer.WriteInt16(Head)
    Call Writer.WriteInt8(Heading)
    Call Writer.WriteInt16(weapon)
    Call Writer.WriteInt16(shield)
    Call Writer.WriteInt16(helmet)
    Call Writer.WriteInt16(FX)
    Call Writer.WriteInt16(FXLoops)
    Call Writer.WriteBool(Idle)
    Call Writer.WriteBool(Navegando)
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
    Call Writer.WriteInt(ServerPacketID.CharacterMove)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt8(X)
    Call Writer.WriteInt8(Y)
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading)
    Call Writer.WriteInt(ServerPacketID.ForceCharMove)
    Call Writer.WriteInt8(Direccion)
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
    Call Writer.WriteInt(ServerPacketID.UpdateTagAndStatus)
    Call Writer.WriteInt16(UserList(UserIndex).Char.CharIndex)
    Call Writer.WriteInt8(Status)
    Call Writer.WriteString8(Tag)
    Call Writer.WriteInt16(UserList(UserIndex).Grupo.Lider)
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Function PrepareMessageErrorMsg(ByVal Message As String)
    Call Writer.WriteInt(ServerPacketID.ErrorMsg)
    Call Writer.WriteString8(Message)
End Function

Public Function PrepareMessageBarFx(ByVal CharIndex As Integer, _
                                    ByVal BarTime As Integer, _
                                    ByVal BarAccion As Byte)
    Call Writer.WriteInt(ServerPacketID.BarFx)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt16(BarTime)
    Call Writer.WriteInt8(BarAccion)
End Function

Public Function PrepareMessageNieblandoToggle(ByVal IntensidadMax As Byte)
    Call Writer.WriteInt(ServerPacketID.NieblaToggle)
    Call Writer.WriteInt8(IntensidadMax)
End Function

Public Function PrepareMessageNevarToggle()
    Call Writer.WriteInt(ServerPacketID.NieveToggle)
End Function

Public Function PrepareMessageDoAnimation(ByVal CharIndex As Integer, _
                                          ByVal Animation As Integer)
    Call Writer.WriteInt(ServerPacketID.DoAnimation)
    Call Writer.WriteInt16(CharIndex)
    Call Writer.WriteInt16(Animation)
End Function

' \End: Prepares
