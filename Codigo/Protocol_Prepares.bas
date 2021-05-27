Attribute VB_Name = "Protocol_Prepares"
Option Explicit

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer              As New clsByteQueue

Public Function PrepareMessageCharSwing(ByVal CharIndex As Integer, Optional ByVal FX As Boolean = True, Optional ByVal ShowText As Boolean = True) As t_DataBuffer

    '***************************************************
    With auxiliarBuffer
    
        Call .WriteID(ServerPacketID.CharSwing)
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(FX)
        Call .WriteBoolean(ShowText)
        
        Call .EndPacket
        
        PrepareMessageCharSwing = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "SetInvisible" message and returns it.
    '***************************************************
    'Call WriteContadores(UserIndex)
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        Call .EndPacket
        
        PrepareMessageSetInvisible = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageSetEscribiendo(ByVal CharIndex As Integer, ByVal Escribiendo As Boolean) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "SetInvisible" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.SetEscribiendo)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Escribiendo)
        Call .EndPacket
        
        PrepareMessageSetEscribiendo = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal chat As String, ByVal CharIndex As Integer, ByVal Color As Long, Optional ByVal name As String = "") As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ChatOverHead" message and returns it.
    '***************************************************
    Dim R, g, b As Byte

    b = (Color And 16711680) / 65536
    g = (Color And 65280) / 256
    R = Color And 255

    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ChatOverHead)
        
        Call .WriteASCIIString(chat)
        Call .WriteInteger(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(R)
        Call .WriteByte(g)
        Call .WriteByte(b)
        Call .WriteLong(Color)
        
        Call .EndPacket
        
        PrepareMessageChatOverHead = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageTextOverChar(ByVal chat As String, ByVal CharIndex As Integer, ByVal Color As Long) As t_DataBuffer

    '***************************************************
    With auxiliarBuffer
    
        Call .WriteID(ServerPacketID.TextOverChar)
        
        Call .WriteASCIIString(chat)
        Call .WriteInteger(CharIndex)
        Call .WriteLong(Color)
        
        Call .EndPacket
        
        PrepareMessageTextOverChar = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

Public Function PrepareMessageTextCharDrop(ByVal chat As String, ByVal CharIndex As Integer, ByVal Color As Long) As t_DataBuffer
        
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.TextCharDrop)
        
        Call .WriteASCIIString(chat)
        Call .WriteInteger(CharIndex)
        Call .WriteLong(Color)
        
        Call .EndPacket
        PrepareMessageTextCharDrop = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

Public Function PrepareMessageTextOverTile(ByVal chat As String, ByVal X As Integer, ByVal Y As Integer, ByVal Color As Long) As t_DataBuffer
        
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.TextOverTile)
        
        Call .WriteASCIIString(chat)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteLong(Color)
        
        Call .EndPacket
        
        PrepareMessageTextOverTile = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal chat As String, ByVal FontIndex As FontTypeNames) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ConsoleMsg" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ConsoleMsg)
        
        Call .WriteASCIIString(chat)
        Call .WriteByte(FontIndex)
        
        Call .EndPacket
        
        PrepareMessageConsoleMsg = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

Public Function PrepareMessageLocaleMsg(ByVal ID As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ConsoleMsg" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.LocaleMsg)
        
        Call .WriteInteger(ID)
        Call .WriteASCIIString(chat)
        Call .WriteByte(FontIndex)
        
        Call .EndPacket
        PrepareMessageLocaleMsg = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

Public Function PrepareMessageListaCorreo(ByVal UserIndex As Integer, ByVal Actualizar As Boolean) As t_DataBuffer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ConsoleMsg" message and returns it.
    '***************************************************

    Dim cant As Byte
    Dim i    As Byte

    cant = UserList(UserIndex).Correo.CantCorreo
    UserList(UserIndex).Correo.NoLeidos = 0

    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ListaCorreo)
        Call .WriteByte(cant)

        If cant > 0 Then

            For i = 1 To cant
                Call .WriteASCIIString(UserList(UserIndex).Correo.Mensaje(i).Remitente)
                Call .WriteASCIIString(UserList(UserIndex).Correo.Mensaje(i).Mensaje)
                Call .WriteByte(UserList(UserIndex).Correo.Mensaje(i).ItemCount)
                Call .WriteASCIIString(UserList(UserIndex).Correo.Mensaje(i).Item)

                Call .WriteByte(UserList(UserIndex).Correo.Mensaje(i).Leido)
                Call .WriteASCIIString(UserList(UserIndex).Correo.Mensaje(i).Fecha)
                'Call ReadMessageCorreo(UserIndex, i)
            Next i

        End If

        Call .WriteBoolean(Actualizar)
        
        Call .EndPacket
        
        PrepareMessageListaCorreo = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
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

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.CreateFX)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        Call .EndPacket
        
        PrepareMessageCreateFX = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

Public Function PrepareMessageMeditateToggle(ByVal CharIndex As Integer, ByVal FX As Integer) As t_DataBuffer
    '***************************************************

    With auxiliarBuffer
        Call .WriteID(ServerPacketID.MeditateToggle)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        
        Call .EndPacket
        PrepareMessageMeditateToggle = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageParticleFX(ByVal CharIndex As Integer, ByVal Particula As Integer, ByVal Time As Long, ByVal Remove As Boolean) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ParticleFX)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(Particula)
        Call .WriteLong(Time)
        Call .WriteBoolean(Remove)
        
        Call .EndPacket
        
        PrepareMessageParticleFX = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

Public Function PrepareMessageParticleFXWithDestino(ByVal Emisor As Integer, ByVal Receptor As Integer, ByVal ParticulaViaje As Integer, ByVal ParticulaFinal As Integer, ByVal Time As Long, ByVal wav As Integer, ByVal FX As Integer) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ParticleFXWithDestino)
        
        Call .WriteInteger(Emisor)
        Call .WriteInteger(Receptor)
        Call .WriteInteger(ParticulaViaje)
        Call .WriteInteger(ParticulaFinal)
        Call .WriteLong(Time)
        Call .WriteInteger(wav)
        Call .WriteInteger(FX)
        
        Call .EndPacket
        
        PrepareMessageParticleFXWithDestino = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

Public Function PrepareMessageParticleFXWithDestinoXY(ByVal Emisor As Integer, ByVal ParticulaViaje As Integer, ByVal ParticulaFinal As Integer, ByVal Time As Long, ByVal wav As Integer, ByVal FX As Integer, ByVal X As Byte, ByVal Y As Byte) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ParticleFXWithDestinoXY)
        
        Call .WriteInteger(Emisor)
        Call .WriteInteger(ParticulaViaje)
        Call .WriteInteger(ParticulaFinal)
        Call .WriteLong(Time)
        Call .WriteInteger(wav)
        Call .WriteInteger(FX)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .EndPacket
        
        PrepareMessageParticleFXWithDestinoXY = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageAuraToChar(ByVal CharIndex As Integer, ByVal Aura As String, ByVal Remove As Boolean, ByVal Tipo As Byte) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.AuraToChar)
        
        Call .WriteInteger(CharIndex)
        Call .WriteASCIIString(Aura)
        Call .WriteBoolean(Remove)
        Call .WriteByte(Tipo)
        
        Call .EndPacket
        
        PrepareMessageAuraToChar = ConvertDataBuffer(.Length, .ReadAll)

    End With
        
End Function

Public Function PrepareMessageSpeedingACT(ByVal CharIndex As Integer, ByVal speeding As Single) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.SpeedToChar)
        
        Call .WriteInteger(CharIndex)
        Call .WriteSingle(speeding)
        
        Call .EndPacket
        
        PrepareMessageSpeedingACT = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageParticleFXToFloor(ByVal X As Byte, ByVal Y As Byte, ByVal Particula As Integer, ByVal Time As Long) As t_DataBuffer

    '***************************************************
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ParticleFXToFloor)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(Particula)
        Call .WriteLong(Time)
        
        Call .EndPacket
        
        PrepareMessageParticleFXToFloor = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageLightFXToFloor(ByVal X As Byte, ByVal Y As Byte, ByVal LuzColor As Long, ByVal Rango As Byte) As t_DataBuffer

    '***************************************************
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.LightToFloor)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteLong(LuzColor)
        Call .WriteByte(Rango)
        
        Call .EndPacket
        
        PrepareMessageLightFXToFloor = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Integer, ByVal X As Byte, ByVal Y As Byte) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.PlayWave)
        
        Call .WriteInteger(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .EndPacket
        
        PrepareMessagePlayWave = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageUbicacionLlamada(ByVal Mapa As Integer, ByVal X As Byte, ByVal Y As Byte) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.PosLLamadaDeClan)
        
        Call .WriteInteger(Mapa)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .EndPacket
        
        PrepareMessageUbicacionLlamada = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageCharUpdateHP(ByVal UserIndex As Integer) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.CharUpdateHP)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHp)
        
        Call .EndPacket
        
        PrepareMessageCharUpdateHP = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageNpcUpdateHP(ByVal NpcIndex As Integer) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.CharUpdateHP)
        
        Call .WriteInteger(NpcList(NpcIndex).Char.CharIndex)
        Call .WriteInteger(NpcList(NpcIndex).Stats.MinHp)
        Call .WriteInteger(NpcList(NpcIndex).Stats.MaxHp)
        
        Call .EndPacket
        
        PrepareMessageCharUpdateHP = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageArmaMov(ByVal CharIndex As Integer) As t_DataBuffer

    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ArmaMov)
        
        Call .WriteInteger(CharIndex)
        
        Call .EndPacket
        PrepareMessageArmaMov = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageEscudoMov(ByVal CharIndex As Integer) As t_DataBuffer

    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.EscudoMov)
        
        Call .WriteInteger(CharIndex)
        
        Call .EndPacket
        
        PrepareMessageEscudoMov = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageFlashScreen(ByVal Color As Long, ByVal Duracion As Long, Optional ByVal Ignorar As Boolean = False) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.FlashScreen)
        
        Call .WriteLong(Color)
        Call .WriteLong(Duracion)
        Call .WriteBoolean(Ignorar)
        
        Call .EndPacket
        
        PrepareMessageFlashScreen = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal chat As String, ByVal status As Byte) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "GuildChat" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.GuildChat)
        Call .WriteByte(status)
        Call .WriteASCIIString(chat)
        
        Call .EndPacket
        
        PrepareMessageGuildChat = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As t_DataBuffer

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Prepares the "ShowMessageBox" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ShowMessageBox)
        
        Call .WriteASCIIString(chat)
        
        Call .EndPacket
        
        PrepareMessageShowMessageBox = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "GuildChat" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.PlayMIDI)
        
        Call .WriteByte(midi)
        Call .WriteInteger(loops)
        
        Call .EndPacket
        
        PrepareMessagePlayMidi = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageOnlineUser(ByVal UserOnline As Integer) As t_DataBuffer

    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.UserOnline)
        
        Call .WriteInteger(UserOnline)
        
        Call .EndPacket
        
        PrepareMessageOnlineUser = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PauseToggle" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.PauseToggle)
        Call .EndPacket
        
        PrepareMessagePauseToggle = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "RainToggle" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.RainToggle)
        Call .EndPacket
        
        PrepareMessageRainToggle = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageTrofeoToggleOn() As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "TrofeoToggle" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.TrofeoToggleOn)
        Call .EndPacket
        
        PrepareMessageTrofeoToggleOn = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageTrofeoToggleOff() As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "TrofeoToggle" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.TrofeoToggleOff)
        Call .EndPacket
        
        PrepareMessageTrofeoToggleOff = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageHora() As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "RainToggle" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.Hora)
        
        Call .WriteLong((GetTickCount() - HoraMundo) Mod DuracionDia)
        Call .WriteLong(DuracionDia)
        
        Call .EndPacket
        
        PrepareMessageHora = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ObjectDelete" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ObjectDelete)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .EndPacket
        
        PrepareMessageObjectDelete = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Byte) As t_DataBuffer

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Prepares the "BlockPosition" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.BlockPosition)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteByte(Blocked)
        
        Call .EndPacket
        
        PrepareMessageBlockPosition = ConvertDataBuffer(.Length, .ReadAll)

    End With

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
Public Function PrepareMessageObjectCreate(ByVal ObjIndex As Integer, ByVal amount As Integer, ByVal X As Byte, ByVal Y As Byte) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'prepares the "ObjectCreate" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ObjectCreate)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(amount)
        
        Call .EndPacket
        
        PrepareMessageObjectCreate = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageFxPiso(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'prepares the "ObjectCreate" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.fxpiso)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(GrhIndex)
        
        Call .EndPacket
        
        PrepareMessageFxPiso = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer, ByVal Desvanecido As Boolean) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterRemove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.CharacterRemove)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(Desvanecido)
        
        Call .EndPacket
        
        PrepareMessageCharacterRemove = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        Call .EndPacket
        PrepareMessageRemoveCharDialog = ConvertDataBuffer(.Length, .ReadAll)

    End With

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

Public Function PrepareMessageCharacterCreate(ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal Status As Byte, ByVal privileges As Byte, ByVal ParticulaFx As Byte, ByVal Head_Aura As String, ByVal Arma_Aura As String, ByVal Body_Aura As String, ByVal DM_Aura As String, ByVal RM_Aura As String, ByVal Otra_Aura As String, ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Byte, ByVal donador As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal Simbolo As Byte, ByVal Idle As Boolean, ByVal Navegando As Boolean) As t_DataBuffer
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterCreate" message and returns it
    '***************************************************

    With auxiliarBuffer
        Call .WriteID(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(Body)
        Call .WriteInteger(Head)
        Call .WriteByte(Heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(name)
        Call .WriteByte(Status)
        Call .WriteByte(privileges)
        Call .WriteByte(ParticulaFx)
        Call .WriteASCIIString(Head_Aura)
        Call .WriteASCIIString(Arma_Aura)
        Call .WriteASCIIString(Body_Aura)
        Call .WriteASCIIString(DM_Aura)
        Call .WriteASCIIString(RM_Aura)
        Call .WriteASCIIString(Otra_Aura)
        Call .WriteASCIIString(Escudo_Aura)
        Call .WriteSingle(speeding)
        Call .WriteByte(EsNPC)
        Call .WriteByte(donador)
        Call .WriteByte(appear)
        Call .WriteInteger(group_index)
        Call .WriteInteger(clan_index)
        Call .WriteByte(clan_nivel)
        Call .WriteLong(UserMinHp)
        Call .WriteLong(UserMaxHp)
        Call .WriteByte(Simbolo)
        Call .WriteBoolean(Idle)
        Call .WriteBoolean(Navegando)
        Call .EndPacket
        PrepareMessageCharacterCreate = ConvertDataBuffer(.Length, .ReadAll)

    End With

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

Public Function PrepareMessageCharacterChange(ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Idle As Boolean, ByVal Navegando As Boolean) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterChange" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(Body)
        Call .WriteInteger(Head)
        Call .WriteByte(Heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteBoolean(Idle)
        Call .WriteBoolean(Navegando)
        
        Call .EndPacket
        
        PrepareMessageCharacterChange = ConvertDataBuffer(.Length, .ReadAll)

    End With
 
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterMove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .EndPacket
        PrepareMessageCharacterMove = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As t_DataBuffer

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Prepares the "ForceCharMove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        Call .EndPacket
        PrepareMessageForceCharMove = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, Status As Byte, Tag As String) As t_DataBuffer
  
    '***************************************************
    'Author: Alejandro Salvo (Salvito)
    'Last Modification: 04/07/07
    'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    'Prepares the "UpdateTagAndStatus" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteByte(Status)
        Call .WriteASCIIString(Tag)
        Call .WriteInteger(UserList(UserIndex).Grupo.Lider)
        
        Call .EndPacket
        
        PrepareMessageUpdateTagAndStatus = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Sub WriteUpdateNPCSimbolo(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Simbolo As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData
        Call .WriteID(ServerPacketID.UpdateNPCSimbolo)
        
        Call .WriteInteger(NpcList(NpcIndex).Char.CharIndex)
        Call .WriteByte(Simbolo)
        
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
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As t_DataBuffer

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ErrorMsg" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(message)
        Call .EndPacket
        
        PrepareMessageErrorMsg = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageBarFx(ByVal CharIndex As Integer, ByVal BarTime As Integer, ByVal BarAccion As Byte) As t_DataBuffer

    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 20/10/2014
    'Envia el Efecto de Barra
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.BarFx)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(BarTime)
        Call .WriteByte(BarAccion)
        
        Call .EndPacket
        
        PrepareMessageBarFx = ConvertDataBuffer(.Length, .ReadAll)

    End With
  
End Function

Public Function PrepareMessageNieblandoToggle(ByVal IntensidadMax As Byte) As t_DataBuffer

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.NieblaToggle)
        Call .WriteByte(IntensidadMax)
        Call .EndPacket
        
        PrepareMessageNieblandoToggle = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function

Public Function PrepareMessageNevarToggle() As t_DataBuffer

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    With auxiliarBuffer
        Call .WriteID(ServerPacketID.NieveToggle)
        Call .EndPacket
        PrepareMessageNevarToggle = ConvertDataBuffer(.Length, .ReadAll)

    End With

End Function
