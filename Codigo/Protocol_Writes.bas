Attribute VB_Name = "Protocol_Writes"
Option Explicit

''
' Writes the "Logged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.logged)
        Call .EndPacket
    
    End With
    
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteHora(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData

        Call .WritePrepared(PrepareMessageHora())
        Call .EndPacket
    
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.RemoveDialogs)
        Call .EndPacket
    
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageRemoveCharDialog(CharIndex))
        Call .EndPacket
        
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

' Writes the "NavigateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.NavigateToggle)
        Call .EndPacket
    
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteNadarToggle(ByVal UserIndex As Integer, ByVal Puede As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.NadarToggle)
        Call .WriteBoolean(Puede)
        
        Call .EndPacket
        
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If
    
End Sub

Public Sub WriteEquiteToggle(ByVal UserIndex As Integer)
        
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.EquiteToggle)
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteVelocidadToggle(ByVal UserIndex As Integer)
        
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.VelocidadToggle)
        Call .WriteSingle(UserList(UserIndex).Char.speeding)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteMacroTrabajoToggle(ByVal UserIndex As Integer, ByVal Activar As Boolean)

    If Not Activar Then
    
        UserList(UserIndex).flags.TargetObj = 0 ' Sacamos el targer del objeto
        UserList(UserIndex).flags.UltimoMensaje = 0
        UserList(UserIndex).Counters.Trabajando = 0
        UserList(UserIndex).flags.UsandoMacro = False
       
    Else
    
        UserList(UserIndex).flags.UsandoMacro = True

    End If

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.MacroTrabajoToggle)
        Call .WriteBoolean(Activar)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Disconnect" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    Call SaveUser(UserIndex, True)
    UserList(UserIndex).flags.YaGuardo = True

    Call WritePersonajesDeCuenta(UserIndex)
    Call WriteMostrarCuenta(UserIndex)
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.Disconnect)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.CommerceEnd)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.BankEnd)
        Call .EndPacket
    
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.CommerceInit)
        Call .WriteASCIIString(NpcList(UserList(UserIndex).flags.TargetNPC).Name)
        Call .EndPacket
    
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.BankInit)
        Call .EndPacket
    
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.UserCommerceInit)
        Call .EndPacket
    
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.UserCommerceEnd)
        Call .EndPacket
    
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ShowBlacksmithForm)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ShowCarpenterForm)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteShowAlquimiaForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ShowAlquimiaForm)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteShowSastreForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ShowSastreForm)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCKillUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(UserIndex).outgoingData.WriteID(ServerPacketID.NPCKillUser)
    Call UserList(UserIndex).outgoingData.EndPacket
    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.BlockedWithShieldUser)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.BlockedWithShieldOther)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharSwing" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharSwing(ByVal UserIndex As Integer, ByVal CharIndex As Integer, Optional ByVal FX As Boolean = True, Optional ByVal ShowText As Boolean = True)

    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageCharSwing(CharIndex, FX, ShowText))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SafeModeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeModeOn" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.SafeModeOn)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeModeOff" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.SafeModeOff)
        Call .EndPacket
    
    End With
    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PartySafeOn" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySafeOn(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Rapsodius
    'Last Modification: 10/10/07
    'Writes the "PartySafeOn" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.PartySafeOn)
        Call .EndPacket
    
    End With
    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PartySafeOff" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySafeOff(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Rapsodius
    'Last Modification: 10/10/07
    'Writes the "PartySafeOff" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.PartySafeOff)
        Call .EndPacket
    
    End With
    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteClanSeguro(ByVal UserIndex As Integer, ByVal estado As Boolean)

    '***************************************************
    'Author: Rapsodius
    'Last Modification: 10/10/07
    'Writes the "PartySafeOff" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ClanSeguro)
        Call .WriteBoolean(estado)
        
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteSeguroResu(ByVal UserIndex As Integer, ByVal estado As Boolean)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.SeguroResu)
        Call .WriteBoolean(estado)
        
        Call .EndPacket
        
    End With
    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.CantUseWhileMeditating)
        
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateMAN(UserIndex))

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)

    'Call SendData(SendTarget.ToDiosesYclan, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateHP(UserIndex))

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateGold" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateExp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMap" message to the given user's outgoing data buffer
    '***************************************************

    Dim Version As Integer

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.changeMap)
        Call .WriteInteger(Map)
        Call .WriteInteger(Version)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PosUpdate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal damage As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCHitUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.NPCHitUser)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHitNPC(ByVal UserIndex As Integer, ByVal damage As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserHitNPC" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UserHitNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call .WriteLong(damage)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserAttackedSwing(ByVal UserIndex As Integer, ByVal AttackerIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserAttackedSwing" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UserAttackedSwing)
        Call .WriteInteger(UserList(AttackerIndex).Char.CharIndex)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserHittedByUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackerChar Char index of the user hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedByUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackerChar As Integer, ByVal damage As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserHittedByUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UserHittedByUser)
        Call .WriteInteger(attackerChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserHittedUser" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackedChar Char index of the user hitted.
' @param    damage The number of HP lost by the oponent hitted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackedChar As Integer, ByVal damage As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserHittedUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UserHittedUser)
        Call .WriteInteger(attackedChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal Color As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
 
    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageChatOverHead(chat, CharIndex, Color))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteTextOverChar(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal Color As Long)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageTextOverChar(chat, CharIndex, Color))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteTextOverTile(ByVal UserIndex As Integer, ByVal chat As String, ByVal X As Integer, ByVal Y As Integer, ByVal Color As Long)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageTextOverTile(chat, X, Y, Color))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteTextCharDrop(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal Color As Long)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageTextCharDrop(chat, CharIndex, Color))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageConsoleMsg(chat, FontIndex))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteLocaleMsg(ByVal UserIndex As Integer, ByVal ID As Integer, ByVal FontIndex As FontTypeNames, Optional ByVal strExtra As String = vbNullString)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
        
    If UserIndex = 0 Then Exit Sub

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageLocaleMsg(ID, strExtra, FontIndex))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteListaCorreo(ByVal UserIndex As Integer, ByVal Actualizar As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageListaCorreo(UserIndex, Actualizar))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal chat As String, ByVal Status As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildChat" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageGuildChat(chat, Status))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteMostrarCuenta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.MostrarCuenta)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageCharacterCreate(Body, Head, Heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, Name, Status, privileges, ParticulaFx, Head_Aura, Arma_Aura, Body_Aura, DM_Aura, RM_Aura, Otra_Aura, Escudo_Aura, speeding, EsNPC, donador, appear, group_index, clan_index, clan_nivel, UserMinHp, UserMaxHp, UserMinMAN, UserMaxMAN, Simbolo, Idle, Navegando))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Desvanecido As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterRemove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageCharacterRemove(CharIndex, Desvanecido))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageCharacterMove(CharIndex, X, Y))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Writes the "ForceCharMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageForceCharMove(Direccion))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

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

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, Optional ByVal Idle As Boolean = False, Optional ByVal Navegando As Boolean = False)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterChange" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageCharacterChange(Body, Head, Heading, CharIndex, weapon, shield, FX, FXLoops, helmet, Idle, Navegando))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal amount As Integer, ByVal X As Byte, ByVal Y As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectCreate" message to the given user's outgoing data buffer
    '***************************************************

    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageObjectCreate(ObjIndex, amount, X, Y))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteParticleFloorCreate(ByVal UserIndex As Integer, ByVal Particula As Integer, ByVal ParticulaTime As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

    On Error GoTo ErrHandler
  
    If Particula = 0 Then Exit Sub
    
    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageParticleFXToFloor(X, Y, Particula, ParticulaTime))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteLightFloorCreate(ByVal UserIndex As Integer, ByVal LuzColor As Long, ByVal Rango As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

    On Error GoTo ErrHandler
     
    MapData(Map, X, Y).Luz.Color = LuzColor
    MapData(Map, X, Y).Luz.Rango = Rango

    If Rango = 0 Then Exit Sub

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageLightFXToFloor(X, Y, LuzColor, Rango))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteFxPiso(ByVal UserIndex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageFxPiso(GrhIndex, X, Y))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectDelete" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageObjectDelete(X, Y))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockPosition" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.BlockPosition)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteByte(Blocked)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PlayMidi" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessagePlayMidi(midi, loops))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessagePlayWave(wave, X, Y))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim Tmp As String
    Dim i   As Long
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AreaChanged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.AreaChanged)
        
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PauseToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessagePauseToggle())
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageRainToggle())
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteNubesToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageNieblandoToggle(IntensidadDeNubes))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteTrofeoToggleOn(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageTrofeoToggleOn())
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteTrofeoToggleOff(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageTrofeoToggleOff())
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateFX" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

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
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateMAN(UserIndex))

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateUserStats)
        
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(ExpLevelUp(UserList(UserIndex).Stats.ELV))
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
        Call .WriteByte(UserList(UserIndex).clase)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteUpdateUserKey(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Llave As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateUserKey)
        
        Call .WriteInteger(Slot)
        Call .WriteInteger(Llave)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

' Actualiza el indicador de daño mágico
Public Sub WriteUpdateDM(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
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

    End With

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateDM)
        
        Call .WriteInteger(Valor)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

' Actualiza el indicador de resistencia mágica
Public Sub WriteUpdateRM(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
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

    End With

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateRM)
        
        Call .WriteInteger(Valor)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.WorkRequestTarget)
        
        Call .WriteByte(Skill)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

' Writes the "InventoryUnlockSlots" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInventoryUnlockSlots(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Ruthnar
    'Last Modification: 30/09/20
    'Writes the "WriteInventoryUnlockSlots" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.InventoryUnlockSlots)
        
        Call .WriteByte(UserList(UserIndex).Stats.InventLevel)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteIntervals(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex)
        Call .outgoingData.WriteID(ServerPacketID.Intervals)
        
        Call .outgoingData.WriteLong(.Intervals.Arco)
        Call .outgoingData.WriteLong(.Intervals.Caminar)
        Call .outgoingData.WriteLong(.Intervals.Golpe)
        Call .outgoingData.WriteLong(.Intervals.GolpeMagia)
        Call .outgoingData.WriteLong(.Intervals.Magia)
        Call .outgoingData.WriteLong(.Intervals.MagiaGolpe)
        Call .outgoingData.WriteLong(.Intervals.GolpeUsar)
        Call .outgoingData.WriteLong(.Intervals.TrabajarExtraer)
        Call .outgoingData.WriteLong(.Intervals.TrabajarConstruir)
        Call .outgoingData.WriteLong(.Intervals.UsarU)
        Call .outgoingData.WriteLong(.Intervals.UsarClic)
        Call .outgoingData.WriteLong(IntervaloTirar)
        
        Call .outgoingData.EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 3/12/09
    'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
    '3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
    '***************************************************

    On Error GoTo ErrHandler

    Dim ObjIndex    As Integer
    Dim PodraUsarlo As Byte

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ChangeInventorySlot)
        
        Call .WriteByte(Slot)
                
        ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex

        If ObjIndex > 0 Then
            PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)

        End If
    
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(Slot).amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(Slot).Equipped)
        Call .WriteSingle(SalePrice(ObjIndex))
        Call .WriteByte(PodraUsarlo)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
        
    Dim ObjIndex    As Integer
    Dim Valor       As Long
    Dim PodraUsarlo As Byte

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ChangeBankSlot)
        
        Call .WriteByte(Slot)

        ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            Valor = ObjData(ObjIndex).Valor
            PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)

        End If

        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(Slot).amount)
        Call .WriteLong(Valor)
        Call .WriteByte(PodraUsarlo)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ChangeSpellSlot)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(Slot))
        
        If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
            Call .WriteByte(UserList(UserIndex).Stats.UserHechizos(Slot))
        Else
            Call .WriteByte("255")
        End If
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Atributes" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Atributes" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.Atributes)
        
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
    'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            obj = ObjData(ArmasHerrero(validIndexes(i)))
            'Call .WriteASCIIString(obj.Index)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
        Next i
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
    'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreria(UserList(UserIndex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(obj.Name)
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
        Next i
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long
    Dim validIndexes() As Integer
    Dim Count          As Byte
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) Then
                If i = 1 Then Debug.Print UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase)
                Count = Count + 1
                validIndexes(Count) = i

            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteByte(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
            'Call .WriteInteger(obj.Madera)
            'Call .WriteLong(obj.GrhIndex)
            ' Ladder 07/07/2014   Ahora se envia el grafico de los objetos
        Next i
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteAlquimistaObjects(ByVal UserIndex As Integer)

    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    ReDim validIndexes(1 To UBound(ObjAlquimista()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.AlquimistaObj)
        
        For i = 1 To UBound(ObjAlquimista())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjAlquimista(i)).SkPociones <= UserList(UserIndex).Stats.UserSkills(eSkill.Alquimia) \ ModAlquimia(UserList(UserIndex).clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Call .WriteInteger(ObjAlquimista(validIndexes(i)))
        Next i
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteSastreObjects(ByVal UserIndex As Integer)

    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long
    Dim obj            As ObjData
    Dim validIndexes() As Integer
    Dim Count          As Integer
    ReDim validIndexes(1 To UBound(ObjSastre()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.SastreObj)
        
        For i = 1 To UBound(ObjSastre())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjSastre(i)).SkMAGOria <= UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Call .WriteInteger(ObjSastre(validIndexes(i)))
        Next i
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RestOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RestOK" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.RestOK)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

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
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageErrorMsg(message))
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Blind" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Blind" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.Blind)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Dumb" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Dumb" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.Dumb)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
'Optimizacion de protocolo por Ladder

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSignal" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowSignal)
        
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef obj As obj, ByVal price As Single)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Last Modified by: Nicolas Ezequiel Bouhid (NicoNZ)
    'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim ObjInfo     As ObjData
    Dim PodraUsarlo As Byte
    
    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(obj.ObjIndex)
        PodraUsarlo = PuedeUsarObjeto(UserIndex, obj.ObjIndex)
    End If
        
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ChangeNPCInventorySlot)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(obj.ObjIndex)
        Call .WriteInteger(obj.amount)
        Call .WriteSingle(price)
        Call .WriteByte(PodraUsarlo)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateHungerAndThirst)
        
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteLight(ByVal UserIndex As Integer, ByVal Map As Integer)

    On Error GoTo ErrHandler

    Dim light As String
        light = MapInfo(Map).base_light

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.light)
        
        Call .WriteASCIIString(light)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteFlashScreen(ByVal UserIndex As Integer, ByVal Color As Long, ByVal Time As Long, Optional ByVal Ignorar As Boolean = False)

    On Error GoTo ErrHandler
 
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.FlashScreen)
        
        Call .WriteLong(Color)
        Call .WriteLong(Time)
        Call .WriteBoolean(Ignorar)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteFYA(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.FYA)
        
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(1))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(2))
        Call .WriteInteger(UserList(UserIndex).flags.DuracionEfecto)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteCerrarleCliente(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.CerrarleCliente)
        Call .EndPacket
        
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteOxigeno(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.Oxigeno)
        
        Call .WriteInteger(UserList(UserIndex).Counters.Oxigeno)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteContadores(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.Contadores)
        
        Call .WriteInteger(UserList(UserIndex).Counters.Invisibilidad)
        Call .WriteInteger(UserList(UserIndex).Counters.ScrollExperiencia)
        Call .WriteInteger(UserList(UserIndex).Counters.ScrollOro)

        If UserList(UserIndex).flags.NecesitaOxigeno Then
            Call .WriteInteger(UserList(UserIndex).Counters.Oxigeno)
        Else
            Call .WriteInteger(0)

        End If

        Call .WriteInteger(UserList(UserIndex).flags.DuracionEfecto)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteBindKeys(ByVal UserIndex As Integer)

    '***************************************************
    'Envia los macros al cliente!
    'Por Ladder
    '23/09/2014
    'Flor te amo!
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.BindKeys)
        
        Call .WriteByte(UserList(UserIndex).ChatCombate)
        Call .WriteByte(UserList(UserIndex).ChatGlobal)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MiniStats" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(UserIndex).Faccion.ciudadanosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.CriminalesMatados)
        Call .WriteByte(UserList(UserIndex).Faccion.Status)
        
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(UserIndex).clase)
        Call .WriteLong(UserList(UserIndex).Counters.Pena)
        
        'Ladder 31/07/08  Envio mas estadisticas :P
        Call .WriteLong(UserList(UserIndex).flags.VecesQueMoriste)
        Call .WriteByte(UserList(UserIndex).genero)
        Call .WriteByte(UserList(UserIndex).raza)
        
        Call .WriteByte(UserList(UserIndex).donador.activo)
        Call .WriteLong(CreditosDonadorCheck(UserList(UserIndex).Cuenta))
        'ARREGLANDO
        
        Call .WriteInteger(DiasDonadorCheck(UserList(UserIndex).Cuenta))
         
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data .incomingData.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LevelUp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.LevelUp)
        
        Call .WriteInteger(skillPoints)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data .incomingData.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, ByVal title As String, ByVal message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AddForumMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.AddForumMsg)
        
        Call .WriteASCIIString(title)
        Call .WriteASCIIString(message)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowForumForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ShowForumForm)
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetInvisible" message to the given user's outgoing data buffer
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WritePrepared(PrepareMessageSetInvisible(CharIndex, invisible))
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DiceRoll" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.DiceRoll)
        
        ' TODO: SACAR ESTE PAQUETE USAR EL DE ATRIBUTOS
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MeditateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.MeditateToggle)
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlindNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.BlindNoMore)
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumbNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.DumbNoMore)
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SendSkills" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.SendSkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserList(UserIndex).Stats.UserSkills(i))
        Next i
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To NpcList(NpcIndex).NroCriaturas
            str = str & NpcList(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal UserIndex As Integer, ByVal guildNews As String, ByRef guildList() As String, ByRef MemberList() As String, ByVal ClanNivel As Byte, ByVal ExpAcu As Integer, ByVal ExpNe As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNews" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.guildNews)
        
        Call .WriteASCIIString(guildNews)

        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
          
        Call .WriteASCIIString(Tmp)
        Call .WriteByte(ClanNivel)
        Call .WriteInteger(ExpAcu)
        Call .WriteInteger(ExpNe)
        
        Call .EndPacket
    End With

    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OfferDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

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

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal CharName As String, ByVal race As eRaza, ByVal Class As eClass, ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.CharacterInfo)
        Call .WriteByte(gender)
        
        Call .WriteASCIIString(CharName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        
        Call .WriteByte(level)
        Call .WriteLong(gold)
        Call .WriteLong(bank)
        
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

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

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, ByRef guildList() As String, ByRef MemberList() As String, ByVal guildNews As String, ByRef joinRequests() As String, ByVal NivelDeClan As Byte, ByVal ExpActual As Integer, ByVal ExpNecesaria As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString

        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Call .WriteByte(NivelDeClan)
        
        Call .WriteInteger(ExpActual)
        Call .WriteInteger(ExpNecesaria)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

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

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, ByVal leader As String, ByVal memberCount As Integer, ByVal alignment As String, ByVal guildDesc As String, ByVal NivelDeClan As Byte, ByVal ExpActual As Integer, ByVal ExpNecesaria As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i    As Long
    Dim temp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteInteger(memberCount)
        Call .WriteASCIIString(alignment)
        Call .WriteASCIIString(guildDesc)
        Call .WriteByte(NivelDeClan)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ShowGuildFundationForm)
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/12/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Writes the "ParalizeOK" message to the given user's outgoing data buffer
    'And updates user position
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ParalizeOK)
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteInmovilizaOK(ByVal UserIndex As Integer)

    '***************************************************
    'Inmovilizar
    'Por Ladder
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.InmovilizadoOK)
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteStopped(ByVal UserIndex As Integer, ByVal Stopped As Boolean)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.Stopped)
        
        Call .WriteBoolean(Stopped)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)
        
        Call .EndPacket
    End With

    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    Amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByRef itemsAenviar() As obj, ByVal gold As Long, ByVal miOferta As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.ChangeUserTradeSlot)
        Call .WriteBoolean(miOferta)
        Call .WriteLong(gold)
            
        Dim i As Long
        For i = 1 To UBound(itemsAenviar)
            Call .WriteInteger(itemsAenviar(i).ObjIndex)

            If itemsAenviar(i).ObjIndex = 0 Then
                Call .WriteASCIIString("")
            Else
                Call .WriteASCIIString(ObjData(itemsAenviar(i).ObjIndex).Name)

            End If
                
            If itemsAenviar(i).ObjIndex = 0 Then
                Call .WriteLong(0)
            Else
                Call .WriteLong(ObjData(itemsAenviar(i).ObjIndex).GrhIndex)

            End If
                
            Call .WriteLong(itemsAenviar(i).amount)
        Next i
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.SpawnListt)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & i & SEPARATOR
            
        Next i
     
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowGMPanelForm)
        
        Call .WriteInteger(UserList(UserIndex).Char.Head)
        Call .WriteInteger(UserList(UserIndex).Char.Body)
        Call .WriteInteger(UserList(UserIndex).Char.CascoAnim)
        Call .WriteInteger(UserList(UserIndex).Char.WeaponAnim)
        Call .WriteInteger(UserList(UserIndex).Char.ShieldAnim)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteShowFundarClanForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ShowFundarClanForm)
        Call .EndPacket
        
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06 NIGO:
    'Writes the "UserNameList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Pong" message to the given user's outgoing data .incomingData.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer, ByVal Time As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Pong" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.Pong)
        
        Call .WriteLong(Time)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub


Public Sub WritePersonajesDeCuenta(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    
    Dim UserCuenta                     As String
    Dim CantPersonajes                 As Byte
    Dim Personaje(1 To MAX_PERSONAJES) As PersonajeCuenta
    Dim donador                        As Boolean
    Dim i                              As Byte
    
    UserCuenta = UserList(UserIndex).Cuenta
    
    donador = DonadorCheck(UserCuenta)

    If Database_Enabled Then
        CantPersonajes = GetPersonajesCuentaDatabase(UserList(UserIndex).AccountID, Personaje)
    
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
    
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.PersonajesDeCuenta)
        Call .WriteByte(CantPersonajes)
            
        For i = 1 To CantPersonajes
            Call .WriteASCIIString(Personaje(i).nombre)
            Call .WriteByte(Personaje(i).nivel)
            Call .WriteInteger(Personaje(i).Mapa)
            Call .WriteInteger(Personaje(i).posX)
            Call .WriteInteger(Personaje(i).posY)
            Call .WriteInteger(Personaje(i).cuerpo)
            Call .WriteInteger(Personaje(i).Cabeza)
            Call .WriteByte(Personaje(i).Status)
            Call .WriteByte(Personaje(i).clase)
            Call .WriteInteger(Personaje(i).Casco)
            Call .WriteInteger(Personaje(i).Escudo)
            Call .WriteInteger(Personaje(i).Arma)
            Call .WriteASCIIString(modGuilds.GuildName(Personaje(i).ClanIndex))
        Next i
            
        Call .WriteByte(IIf(donador, 1, 0))
        
        Call .EndPacket
    End With
    
    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteGoliathInit(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.Goliath)
        
        Call .WriteLong(UserList(UserIndex).Stats.Banco)
        Call .WriteByte(UserList(UserIndex).BancoInvent.NroItems)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteShowFrmLogear(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowFrmLogear)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteShowFrmMapa(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowFrmMapa)
        
        If UserList(UserIndex).donador.activo = 1 Then
            Call .WriteInteger(ExpMult * UserList(UserIndex).flags.ScrollExp * 1.1)
        Else
            Call .WriteInteger(ExpMult * UserList(UserIndex).flags.ScrollExp)
        End If

        Call .WriteInteger(OroMult * UserList(UserIndex).flags.ScrollOro)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteFamiliar(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        Call .outgoingData.WriteID(ServerPacketID.Familiar)
        
        Call .outgoingData.WriteByte(.Familiar.Existe)
        Call .outgoingData.WriteByte(.Familiar.Muerto)
        Call .outgoingData.WriteASCIIString(.Familiar.nombre)
        Call .outgoingData.WriteLong(.Familiar.Exp)
        Call .outgoingData.WriteLong(.Familiar.ELU)
        Call .outgoingData.WriteByte(.Familiar.nivel)
        Call .outgoingData.WriteInteger(.Familiar.MinHp)
        Call .outgoingData.WriteInteger(.Familiar.MaxHp)
        Call .outgoingData.WriteInteger(.Familiar.MinHIT)
        Call .outgoingData.WriteInteger(.Familiar.MaxHit)
        
        Call .outgoingData.EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteRecompensas(ByVal UserIndex As Integer)
        
    On Error GoTo WriteRecompensas_Err

    '***************************************************
    'Envia las recompensas al cliente!
    'Por Ladder
    '22/04/2015
    'Flor te amo!
    '***************************************************

    With UserList(UserIndex).outgoingData
    
        Dim a, b, c As Byte
 
        b = UserList(UserIndex).UserLogros + 1
        a = UserList(UserIndex).NPcLogros + 1
        c = UserList(UserIndex).LevelLogros + 1
        
        Call .WriteID(ServerPacketID.Logros)
        
        'Logros NPC
        Call .WriteASCIIString(NPcLogros(a).nombre)
        Call .WriteASCIIString(NPcLogros(a).Desc)
        Call .WriteInteger(NPcLogros(a).cant)
        Call .WriteByte(NPcLogros(a).TipoRecompensa)
        
        If NPcLogros(a).TipoRecompensa = 1 Then
            Call .WriteASCIIString(NPcLogros(a).ObjRecompensa)
        End If

        If NPcLogros(a).TipoRecompensa = 2 Then
            Call .WriteLong(NPcLogros(a).OroRecompensa)
        End If

        If NPcLogros(a).TipoRecompensa = 3 Then
            Call .WriteLong(NPcLogros(a).ExpRecompensa)
        End If
        
        If NPcLogros(a).TipoRecompensa = 4 Then
            Call .WriteByte(NPcLogros(a).HechizoRecompensa)
        End If
        
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
        If UserList(UserIndex).Stats.NPCsMuertos >= NPcLogros(a).cant Then
            Call .WriteBoolean(True)
        Else
            Call .WriteBoolean(False)
        End If
        
        'Logros User
        Call .WriteASCIIString(UserLogros(b).nombre)
        Call .WriteASCIIString(UserLogros(b).Desc)
        Call .WriteInteger(UserLogros(b).cant)
        Call .WriteInteger(UserLogros(b).TipoRecompensa)
        Call .WriteInteger(UserList(UserIndex).Stats.UsuariosMatados)

        If UserLogros(a).TipoRecompensa = 1 Then
            Call .WriteASCIIString(UserLogros(b).ObjRecompensa)
        End If
        
        If UserLogros(a).TipoRecompensa = 2 Then
            Call .WriteLong(UserLogros(b).OroRecompensa)

        End If

        If UserLogros(a).TipoRecompensa = 3 Then
            Call .WriteLong(UserLogros(b).ExpRecompensa)

        End If
        
        If UserLogros(a).TipoRecompensa = 4 Then
            Call .WriteByte(UserLogros(b).HechizoRecompensa)

        End If

        If UserList(UserIndex).Stats.UsuariosMatados >= UserLogros(b).cant Then
            Call .WriteBoolean(True)
        Else
            Call .WriteBoolean(False)

        End If

        'Nivel User
        Call .WriteASCIIString(LevelLogros(c).nombre)
        Call .WriteASCIIString(LevelLogros(c).Desc)
        Call .WriteInteger(LevelLogros(c).cant)
        Call .WriteInteger(LevelLogros(c).TipoRecompensa)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)

        If LevelLogros(c).TipoRecompensa = 1 Then
            Call .WriteASCIIString(LevelLogros(c).ObjRecompensa)

        End If
        
        If LevelLogros(c).TipoRecompensa = 2 Then
            Call .WriteLong(LevelLogros(c).OroRecompensa)

        End If

        If LevelLogros(c).TipoRecompensa = 3 Then
            Call .WriteLong(LevelLogros(c).ExpRecompensa)

        End If
        
        If LevelLogros(c).TipoRecompensa = 4 Then
            Call .WriteByte(LevelLogros(c).HechizoRecompensa)

        End If

        If UserList(UserIndex).Stats.ELV >= LevelLogros(c).cant Then
            Call .WriteBoolean(True)
        Else
            Call .WriteBoolean(False)

        End If
        
        Call .EndPacket
    End With

    Exit Sub

WriteRecompensas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRecompensas", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
        
End Sub

Public Sub WritePreguntaBox(ByVal UserIndex As Integer, ByVal message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowPregunta)
        
        Call .WriteASCIIString(message)
        
        Call UserList(UserIndex).outgoingData.EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteDatosGrupo(ByVal UserIndex As Integer)

    Dim i As Byte

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        Call .outgoingData.WriteID(ServerPacketID.DatosGrupo)
        Call .outgoingData.WriteBoolean(.Grupo.EnGrupo)
        
        If .Grupo.EnGrupo = True Then
            Call .outgoingData.WriteByte(UserList(.Grupo.Lider).Grupo.CantidadMiembros)
            'Call .outgoingData.WriteByte(UserList(.Grupo.Lider).name)
   
            If .Grupo.Lider = UserIndex Then
             
                For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros

                    If i = 1 Then
                        Call .outgoingData.WriteASCIIString(UserList(.Grupo.Miembros(i)).Name & "(Líder)")
                    Else
                        Call .outgoingData.WriteASCIIString(UserList(.Grupo.Miembros(i)).Name)

                    End If

                Next i

            Else
          
                For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros
                
                    If i = 1 Then
                        Call .outgoingData.WriteASCIIString(UserList(UserList(.Grupo.Lider).Grupo.Miembros(i)).Name & "(Líder)")
                    Else
                        Call .outgoingData.WriteASCIIString(UserList(UserList(.Grupo.Lider).Grupo.Miembros(i)).Name)

                    End If

                Next i

            End If

        End If
   
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub


Public Sub WriteUbicacion(ByVal UserIndex As Integer, ByVal Miembro As Byte, ByVal GPS As Integer)

    Dim i   As Byte
    Dim X   As Byte
    Dim Y   As Byte
    Dim Map As Integer

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ubicacion)
        Call .WriteByte(Miembro)

        If GPS > 0 Then
        
            Call .WriteByte(UserList(GPS).Pos.X)
            Call .WriteByte(UserList(GPS).Pos.Y)
            Call .WriteInteger(UserList(GPS).Pos.Map)
        Else
            Call .WriteByte(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)

        End If
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteCorreoPicOn(ByVal UserIndex As Integer)

    '***************************************************
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.CorreoPicOn)
        Call .EndPacket

    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteShop(ByVal UserIndex As Integer)

    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ObjDonador()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.DonadorObj)
        
        For i = 1 To UBound(ObjDonador())
            Count = Count + 1
            validIndexes(Count) = i
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Call .WriteInteger(ObjDonador(validIndexes(i)).ObjIndex)
            Call .WriteInteger(ObjDonador(validIndexes(i)).Valor)
        Next i
        
        Call .WriteLong(CreditosDonadorCheck(UserList(UserIndex).Cuenta))
        Call .WriteInteger(DiasDonadorCheck(UserList(UserIndex).Cuenta))
        
        Call .EndPacket
    End With

    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteRanking(ByVal UserIndex As Integer)

    '***************************************************
    On Error GoTo ErrHandler

    Dim i As Byte
    
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.Ranking)

        For i = 1 To 10
            Call .WriteASCIIString(Rankings(1).user(i).Nick)
            Call .WriteInteger(Rankings(1).user(i).Value)
        Next i
        
        Call .EndPacket
    End With

    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteActShop(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ActShop)
        
        Call .WriteLong(CreditosDonadorCheck(UserList(UserIndex).Cuenta))
        Call .WriteInteger(DiasDonadorCheck(UserList(UserIndex).Cuenta))
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteViajarForm(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.ViajarForm)
        
        Dim destinos As Byte
        Dim i        As Byte

        destinos = NpcList(NpcIndex).NumDestinos
        
        Call .WriteByte(destinos)
        
        For i = 1 To destinos
            Call .WriteASCIIString(NpcList(NpcIndex).Dest(i))
        Next i
        
        Call .WriteByte(NpcList(NpcIndex).Interface)
        
        Call .EndPacket
    End With

    
    Exit Sub
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteQuestDetails(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, Optional QuestSlot As Byte = 0)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envía el paquete QuestDetails y la información correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
 
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
    
        'ID del paquete
        Call .WriteID(ServerPacketID.QuestDetails)
        
        'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptí todavía (1 para el primer caso y 0 para el segundo)
        Call .WriteByte(IIf(QuestSlot, 1, 0))
        
        'Enviamos nombre, descripción y nivel requerido de la quest
        'Call .WriteASCIIString(QuestList(QuestIndex).Nombre)
        'Call .WriteASCIIString(QuestList(QuestIndex).Desc)
        Call .WriteInteger(QuestIndex)
        Call .WriteByte(QuestList(QuestIndex).RequiredLevel)
        
        Call .WriteInteger(QuestList(QuestIndex).RequiredQuest)
        
        'Enviamos la cantidad de npcs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)

        If QuestList(QuestIndex).RequiredNPCs Then

            'Si hay npcs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).amount)
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)

                'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
                If QuestSlot Then
                    Call .WriteInteger(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))

                End If

            Next i

        End If
        
        'Enviamos la cantidad de objs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredOBJs)

        If QuestList(QuestIndex).RequiredOBJs Then

            'Si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredOBJs
                Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).amount)
                Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex)
                
                'escribe si tiene ese objeto en el inventario y que cantidad
                Call .WriteInteger(CantidadObjEnInv(UserIndex, QuestList(QuestIndex).RequiredOBJ(i).ObjIndex))
                ' Call .WriteInteger(0)
                
            Next i

        End If
    
        'Enviamos la recompensa de oro y experiencia.
        Call .WriteLong((QuestList(QuestIndex).RewardGLD * OroMult))
        Call .WriteLong((QuestList(QuestIndex).RewardEXP * ExpMult))
        
        'Enviamos la cantidad de objs de recompensa
        Call .WriteByte(QuestList(QuestIndex).RewardOBJs)

        If QuestList(QuestIndex).RewardOBJs Then

            'si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RewardOBJs
                Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).amount)
                Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
            Next i

        End If
        
        Call .EndPacket
    End With

    Exit Sub
 
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
 
Public Sub WriteQuestListSend(ByVal UserIndex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envía el paquete QuestList y la información correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i       As Integer
    Dim tmpStr  As String
    Dim tmpByte As Byte
 
    On Error GoTo ErrHandler
 
    With UserList(UserIndex)
        Call .outgoingData.WriteID(ServerPacketID.QuestListSend)
    
        For i = 1 To MAXUSERQUESTS

            If .QuestStats.Quests(i).QuestIndex Then
                tmpByte = tmpByte + 1
                tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).nombre & "-"

            End If

        Next i
        
        'Escribimos la cantidad de quests
        Call .outgoingData.WriteByte(tmpByte)
        
        'Escribimos la lista de quests (sacamos el íltimo caracter)
        If tmpByte Then
            Call .outgoingData.WriteASCIIString(Left$(tmpStr, Len(tmpStr) - 1))
        End If
        
        Call .outgoingData.EndPacket
    End With

    Exit Sub
 
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteNpcQuestListSend(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envía el paquete QuestList y la información correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i       As Integer
    Dim j       As Integer
    Dim tmpStr  As String
    Dim tmpByte As Byte
 
    On Error GoTo ErrHandler
    
    Dim QuestIndex As Integer
 
    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.NpcQuestListSend)
        
        Call .WriteByte(NpcList(NpcIndex).NumQuest) 'Escribimos primero cuantas quest tiene el NPC
    
        For j = 1 To NpcList(NpcIndex).NumQuest
        
            QuestIndex = NpcList(NpcIndex).QuestNumber(j)
            
            Call .WriteInteger(QuestIndex)
            Call .WriteByte(QuestList(QuestIndex).RequiredLevel)
        
            Call .WriteInteger(QuestList(QuestIndex).RequiredQuest)
        
            'Enviamos la cantidad de npcs requeridos
            Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)

            If QuestList(QuestIndex).RequiredNPCs Then

                'Si hay npcs entonces enviamos la lista
                For i = 1 To QuestList(QuestIndex).RequiredNPCs
                    Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).amount)
                    Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)

                    'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
                    'If QuestSlot Then
                    ' Call .WriteInteger(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))

                    ' End If

                Next i

            End If
        
            'Enviamos la cantidad de objs requeridos
            Call .WriteByte(QuestList(QuestIndex).RequiredOBJs)

            If QuestList(QuestIndex).RequiredOBJs Then

                'Si hay objs entonces enviamos la lista
                For i = 1 To QuestList(QuestIndex).RequiredOBJs
                    Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).amount)
                    Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex)
                Next i

            End If
    
            'Enviamos la recompensa de oro y experiencia.
            Call .WriteLong(QuestList(QuestIndex).RewardGLD * OroMult)
            Call .WriteLong(QuestList(QuestIndex).RewardEXP * ExpMult)
        
            'Enviamos la cantidad de objs de recompensa
            Call .WriteByte(QuestList(QuestIndex).RewardOBJs)

            If QuestList(QuestIndex).RewardOBJs Then

                'si hay objs entonces enviamos la lista
                For i = 1 To QuestList(QuestIndex).RewardOBJs
                    Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).amount)
                    Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
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
                Call .WriteByte(1)
            Else

                If UserDoneQuest(UserIndex, QuestIndex) Then
                    Call .WriteByte(2)
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
                        Call .WriteByte(0)
                    Else
                        Call .WriteByte(3)

                    End If
                
                End If
                
            End If

        Next j
        
        Call .EndPacket
        
    End With

    Exit Sub
 
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteRequestScreenShot(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
    
    
    With UserList(UserIndex).outgoingData
    
        Call .WriteID(ServerPacketID.RequestScreenShot)
        Call .EndPacket
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub


Public Sub WriteShowScreenShot(ByVal UserIndex As Integer, Name As String)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.ShowScreenShot)
        
        Call .WriteASCIIString(Name)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub


Public Sub WriteScreenShotData(ByVal UserIndex As Integer, Buffer As clsByteQueue, ByVal Offset As Long, ByVal Size As Long)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.ScreenShotData)

        Call .WriteSubBuffer(Buffer, Offset, Size)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteTolerancia0(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.Tolerancia0)
        Call .EndPacket
        
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteRedundancia(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.Redundancia)
        
        Call .WriteByte(UserList(UserIndex).Redundance)
        
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Sub WriteCommerceRecieveChatMessage(ByVal UserIndex As Integer, ByVal message As String)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.CommerceRecieveChatMessage)
        Call .WriteASCIIString(message)
        Call .EndPacket
    End With

    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Sub WriteInvasionInfo(ByVal UserIndex As Integer, ByVal Invasion As Integer, ByVal PorcentajeVida As Byte, ByVal PorcentajeTiempo As Byte)
    
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.InvasionInfo)
        
        Call .WriteByte(Invasion)
        Call .WriteByte(PorcentajeVida)
        Call .WriteByte(PorcentajeTiempo)
        
        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If
    
End Sub

Sub WriteOpenCrafting(ByVal UserIndex As Integer, ByVal Tipo As Byte)
    
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.OpenCrafting)
        Call .WriteByte(Tipo)

        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
    
End Sub

Sub WriteCraftingItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ObjIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.CraftingItem)
        Call .WriteByte(Slot)
        Call .WriteInteger(ObjIndex)

        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
    
End Sub

Sub WriteCraftingCatalyst(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal amount As Integer, ByVal Porcentaje As Byte)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.CraftingCatalyst)
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(amount)
        Call .WriteByte(Porcentaje)

        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
    
End Sub

Sub WriteCraftingResult(ByVal UserIndex As Integer, ByVal Result As Integer, Optional ByVal Porcentaje As Byte, Optional ByVal Precio As Long)
    
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.CraftingResult)
        Call .WriteInteger(Result)
        
        If Result <> 0 Then
            Call .WriteByte(Porcentaje)
            Call .WriteLong(Precio)
        End If

        Call .EndPacket
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
    
End Sub
