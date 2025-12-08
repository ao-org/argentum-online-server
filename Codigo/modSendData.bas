Attribute VB_Name = "modSendData"
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

Public Enum SendTarget
    ToAll = 1
    ToIndex
    toMap
    ToPCArea
    ToPCAliveArea
    ToPCAreaButGMs
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToNPCAliveArea
    ToNPCDeadArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToPCAliveAreaButIndex
    ToAdminAreaButIndex
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToSuperiores
    ToSuperioresArea
    ToPCDeadArea
    ToPCDeadAreaButIndex
    ToAdminsYDioses
    ToJugadoresCaptura
    ToGroup
    ToGroupButIndex
End Enum

Public Sub SendToConnection(ByVal ConnectionID, Optional Args As Variant)
    On Error GoTo SendToConnection_Err
    #If DIRECT_PLAY = 0 Then
        Dim Writer As Network.Writer
        Set Writer = Protocol_Writes.GetWriterBuffer()
    #Else
        Dim Writer As clsNetWriter
        Set Writer = Protocol_Writes.Writer
    #End If
    Call modNetwork.SendToConnection(ConnectionID, Writer)
    Writer.Clear
    Exit Sub
SendToConnection_Err:
    Call TraceError(Err.Number, Err.Description, "modSendData.SendToConnection", Erl)
    Call Writer.Clear
End Sub

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, Optional Args As Variant, Optional ByVal ValidateInvi As Boolean = False)
    On Error GoTo SendData_Err
    #If DIRECT_PLAY = 0 Then
        Dim Buffer As Network.Writer
        Set Buffer = Protocol_Writes.GetWriterBuffer()
    #Else
        Dim Buffer As clsNetWriter
        Set Buffer = Protocol_Writes.Writer
    #End If
    Dim LoopC As Long
    Dim Map   As Integer
    Select Case sndRoute
        Case SendTarget.ToIndex
            Debug.Assert sndIndex >= LBound(UserList) And sndIndex <= UBound(UserList)
            With UserList(sndIndex)
                If (.ConnectionDetails.ConnIDValida) Then
                    Call modNetwork.Send(sndIndex, Buffer)
                End If
            End With
        Case SendTarget.ToPCArea
            Debug.Assert sndIndex >= LBound(UserList) And sndIndex <= UBound(UserList)
            Call SendToUserArea(sndIndex, Buffer, ValidateInvi)
        Case SendTarget.ToPCAliveArea
            Debug.Assert sndIndex >= LBound(UserList) And sndIndex <= UBound(UserList)
            Call SendToUserAliveArea(sndIndex, Buffer, ValidateInvi)
        Case SendTarget.ToPCAreaButGMs
            Debug.Assert sndIndex >= LBound(UserList) And sndIndex <= UBound(UserList)
            Call SendToUserAreaButGMs(sndIndex, Buffer)
        Case SendTarget.ToPCDeadArea
            Call SendToPCDeadArea(sndIndex, Buffer)
        Case SendTarget.ToPCDeadAreaButIndex
            Call SendToPCDeadAreaButIndex(sndIndex, Buffer)
        Case SendTarget.ToAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnectionDetails.ConnIDValida Then
                    If UserList(LoopC).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero) Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToAdminsYDioses
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnectionDetails.ConnIDValida Then
                    If UserList(LoopC).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios) Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToJugadoresCaptura
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnectionDetails.ConnIDValida Then
                    If UserList(LoopC).flags.jugando_captura = 1 Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToSuperiores
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnectionDetails.ConnIDValida Then
                    If CompararPrivilegiosUser(LoopC, sndIndex) > 0 Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToSuperioresArea
            Call SendToSuperioresArea(sndIndex, Buffer)
        Case SendTarget.ToAll
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnectionDetails.ConnIDValida Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToAllButIndex
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnectionDetails.ConnIDValida) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.toMap
            Call SendToMap(sndIndex, Buffer)
        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, Buffer)
        Case SendTarget.ToGuildMembers
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                    Call modNetwork.Send(LoopC, Buffer)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, Buffer, ValidateInvi)
        Case SendTarget.ToPCAliveAreaButIndex
            Call SendToUserAliveAreaButindex(sndIndex, Buffer, ValidateInvi)
        Case SendTarget.ToAdminAreaButIndex
            Call SendToAdminAreaButIndex(sndIndex, Buffer)
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, Buffer)
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, Buffer)
        Case SendTarget.ToNPCAliveArea
            Call SendToNpcAliveArea(sndIndex, Buffer)
        Case SendTarget.ToDiosesYclan
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                    Call modNetwork.Send(LoopC, Buffer)
                End If
                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            While LoopC > 0
                If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                    Call modNetwork.Send(LoopC, Buffer)
                End If
                LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            Wend
        Case SendTarget.ToConsejo
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                    If UserList(LoopC).Faccion.Status = e_Facciones.consejo Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToConsejoCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                    If UserList(LoopC).Faccion.Status = e_Facciones.concilio Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToRolesMasters
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                    If UserList(LoopC).flags.Privilegios And e_PlayerType.RoleMaster Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToRealYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                    If UserList(LoopC).Faccion.Status = e_Facciones.Armada Or (UserList(LoopC).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or _
                            e_PlayerType.SemiDios Or e_PlayerType.Consejero)) <> 0 Or UserList(LoopC).Faccion.Status = e_Facciones.consejo Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToCaosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                    If UserList(LoopC).Faccion.Status = e_Facciones.Caos Or (UserList(LoopC).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or _
                            e_PlayerType.SemiDios Or e_PlayerType.Consejero)) <> 0 Or UserList(LoopC).Faccion.Status = e_Facciones.concilio Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                End If
            Next LoopC
        Case SendTarget.ToGroup
            Call SendToGroup(sndIndex, Buffer)
        Case SendTarget.ToGroupButIndex
            Call SendToGroupButIndex(sndIndex, Buffer)
    End Select
SendData_Err:
    Call Buffer.Clear
    If (Err.Number <> 0) Then
        Call TraceError(Err.Number, Err.Description, "modSendData.SendData", Erl)
    End If
End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToUserAliveArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer, Optional ByVal ValidateInvi As Boolean = False)
    #Else
        Private Sub SendToUserAliveArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter, Optional ByVal ValidateInvi As Boolean = False)
        #End If
        On Error GoTo SendToUserArea_Err
        Dim LoopC      As Long
        Dim tempIndex  As Integer
        Dim Map        As Integer
        Dim AreaX      As Integer
        Dim AreaY      As Integer
        Dim enviaDatos As Boolean
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        If UserList(tempIndex).flags.Muerto = 0 Or MapInfo(UserList(tempIndex).pos.Map).Seguro = 1 Or (UserList(UserIndex).GuildIndex > 0 And UserList( _
                                UserIndex).GuildIndex = UserList(tempIndex).GuildIndex) Or IsSet(UserList(UserIndex).flags.StatusMask, e_StatusMask.eTalkToDead) Then
                            enviaDatos = True
                            If Not EsGM(tempIndex) Then
                                If UserList(UserIndex).flags.invisible + UserList(UserIndex).flags.Oculto > 0 And ValidateInvi And Not (UserList(tempIndex).GuildIndex > 0 And _
                                        UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex And modGuilds.NivelDeClan(UserList(tempIndex).GuildIndex) >= RequiredGuildLevelSeeInvisible) And _
                                        UserList(UserIndex).flags.Navegando = 0 Then
                                    If Distancia(UserList(UserIndex).pos, UserList(tempIndex).pos) > DISTANCIA_ENVIO_DATOS And UserList(UserIndex).Counters.timeFx + UserList( _
                                            UserIndex).Counters.timeChat = 0 Then
                                        enviaDatos = False
                                    End If
                                End If
                            End If
                            If enviaDatos Then
                                Call modNetwork.Send(tempIndex, Buffer)
                            End If
                        End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToUserArea_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserArea", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer, Optional ByVal ValidateInvi As Boolean)
    #Else
        Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter, Optional ByVal ValidateInvi As Boolean)
        #End If
        On Error GoTo SendToUserArea_Err
        Dim LoopC      As Long
        Dim tempIndex  As Integer
        Dim Map        As Integer
        Dim AreaX      As Integer
        Dim AreaY      As Integer
        Dim enviaDatos As Boolean
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        enviaDatos = True
                        If Not EsGM(tempIndex) Then
                            If UserList(UserIndex).flags.invisible + UserList(UserIndex).flags.Oculto > 0 And ValidateInvi Then
                                If Distancia(UserList(UserIndex).pos, UserList(tempIndex).pos) > DISTANCIA_ENVIO_DATOS And UserList(UserIndex).Counters.timeFx + UserList( _
                                        UserIndex).Counters.timeChat = 0 Then
                                    enviaDatos = False
                                End If
                            End If
                        End If
                        If enviaDatos Then
                            Call modNetwork.Send(tempIndex, Buffer)
                        End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToUserArea_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserArea", Erl)
    End Sub


#If DIRECT_PLAY = 0 Then
    Private Sub SendToPCDeadArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToPCDeadArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToUserArea_Err
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        ' Envio a los que estan MUERTOS y a los GMs cercanos.
                        If UserList(tempIndex).flags.Muerto = 1 Or EsGM(tempIndex) Or IsSet(UserList(tempIndex).flags.StatusMask, e_StatusMask.eTalkToDead) Then
                            Call modNetwork.Send(tempIndex, Buffer)
                        End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToUserArea_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToPCDeadArea", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToPCDeadAreaButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToPCDeadAreaButIndex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToUserArea_Err
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        ' Envio a los que estan MUERTOS y a los GMs cercanos.
                        If tempIndex <> UserIndex Then
                            If UserList(tempIndex).flags.Muerto = 1 Or IsSet(UserList(tempIndex).flags.StatusMask, e_StatusMask.eTalkToDead) Then
                                Call modNetwork.Send(tempIndex, Buffer)
                            End If
                        End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToUserArea_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToPCDeadArea", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToSuperioresArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToSuperioresArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToSuperioresArea_Err
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
            If TempInt Then  'Esta en el area?
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                If TempInt Then
                    If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        If CompararPrivilegiosUser(UserIndex, tempIndex) < 0 Then
                            Call modNetwork.Send(tempIndex, Buffer)
                        End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToSuperioresArea_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToSuperioresArea", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer, Optional ByVal ValidateInvi As Boolean = False)
    #Else
        Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter, Optional ByVal ValidateInvi As Boolean = False)
        #End If
        On Error GoTo SendToUserAreaButindex_Err
        Dim LoopC      As Long
        Dim TempInt    As Integer
        Dim tempIndex  As Integer
        Dim Map        As Integer
        Dim AreaX      As Integer
        Dim AreaY      As Integer
        Dim enviaDatos As Boolean
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
            If TempInt Then  'Esta en el area?
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                If TempInt Then
                    If tempIndex <> UserIndex Then
                        If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                            enviaDatos = True
                            If Not EsGM(tempIndex) Then
                                If UserList(UserIndex).flags.invisible + UserList(UserIndex).flags.Oculto > 0 And ValidateInvi Then
                                    If Distancia(UserList(UserIndex).pos, UserList(tempIndex).pos) > DISTANCIA_ENVIO_DATOS And UserList(UserIndex).Counters.timeFx + UserList( _
                                            UserIndex).Counters.timeChat = 0 Then
                                        enviaDatos = False
                                    End If
                                End If
                            End If
                            If enviaDatos Then
                                Call modNetwork.Send(tempIndex, Buffer)
                            End If
                        End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToUserAreaButindex_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAreaButindex", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Function CanSendToUser(ByRef SourceUser As t_User, _
                                   ByRef TargetUser As t_User, _
                                   ByVal TargetIndex As Integer, _
                                   ByRef Buffer As Network.Writer, _
                                   ByVal ValidateInvi As Boolean) As Boolean
    #Else
        Private Function CanSendToUser(ByRef SourceUser As t_User, ByRef TargetUser As t_User, ByVal TargetIndex As Integer, ByRef Buffer As clsNetWriter, ByVal ValidateInvi As _
                Boolean) As Boolean
        #End If
        If (TargetUser.AreasInfo.AreaReciveX And SourceUser.AreasInfo.AreaPerteneceX) = 0 Then Exit Function
        If (TargetUser.AreasInfo.AreaReciveY And SourceUser.AreasInfo.AreaPerteneceY) = 0 Then Exit Function
        If Not TargetUser.ConnectionDetails.ConnIDValida Then Exit Function
        If Not (TargetUser.flags.Muerto = 0 Or MapInfo(TargetUser.pos.Map).Seguro = 1 Or (SourceUser.GuildIndex > 0 And SourceUser.GuildIndex = TargetUser.GuildIndex) Or IsSet( _
                TargetUser.flags.StatusMask, e_StatusMask.eTalkToDead) Or IsSet(SourceUser.flags.StatusMask, e_StatusMask.eTalkToDead)) Then Exit Function
        If Not EsGM(TargetIndex) Then
            If SourceUser.flags.invisible + SourceUser.flags.Oculto > 0 And ValidateInvi And Not CheckGuildSend(SourceUser, TargetUser) And SourceUser.flags.Navegando = 0 Then
                If Distancia(SourceUser.pos, TargetUser.pos) > DISTANCIA_ENVIO_DATOS And SourceUser.Counters.timeFx + SourceUser.Counters.timeChat = 0 Then
                    Exit Function
                End If
            End If
        End If
        CanSendToUser = True
    End Function

Public Function CheckGuildSend(ByRef SourceUser As t_User, ByRef TargetUser As t_User) As Boolean
    CheckGuildSend = False
    If SourceUser.GuildIndex = 0 Then Exit Function
    If SourceUser.GuildIndex <> TargetUser.GuildIndex Then Exit Function
    If modGuilds.NivelDeClan(TargetUser.GuildIndex) < RequiredGuildLevelSeeInvisible Then
        CheckGuildSend = SourceUser.Counters.timeGuildChat > 0
        Exit Function
    End If
    CheckGuildSend = True
End Function

#If DIRECT_PLAY = 0 Then
    Private Sub SendToUserAliveAreaButindex(ByVal UserIndex As Integer, ByRef Buffer As Network.Writer, Optional ByVal ValidateInvi As Boolean = False)
    #Else
        Private Sub SendToUserAliveAreaButindex(ByVal UserIndex As Integer, ByRef Buffer As clsNetWriter, Optional ByVal ValidateInvi As Boolean = False)
        #End If
        On Error GoTo SendToUserAliveAreaButindex_Err
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        If Not MapaValido(Map) Then Exit Sub
        With UserList(UserIndex)
            For LoopC = 1 To ConnGroups(Map).CountEntrys
                tempIndex = ConnGroups(Map).UserEntrys(LoopC)
                If tempIndex <> UserIndex Then
                    If CanSendToUser(UserList(UserIndex), UserList(tempIndex), tempIndex, Buffer, ValidateInvi) Then
                        Call modNetwork.Send(tempIndex, Buffer)
                    End If
                End If
            Next LoopC
        End With
        Exit Sub
SendToUserAliveAreaButindex_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAliveAreaButindex", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToAdminAreaButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToAdminAreaButIndex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToUserAreaButindex_Err
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
            If TempInt Then  'Esta en el area?
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                If TempInt Then
                    If tempIndex <> UserIndex And EsGM(tempIndex) Then
                        If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                            If CompararPrivilegios(UserList(tempIndex).flags.Privilegios, UserList(UserIndex).flags.Privilegios) >= 0 Then
                                Call modNetwork.Send(tempIndex, Buffer)
                            End If
                        End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToUserAreaButindex_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToAdminAreaButIndex", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToUserAreaButGMs(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToUserAreaButGMs(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToUserAreaButindex_Err
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
            If TempInt Then  'Esta en el area?
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                If TempInt Then
                    If Not EsGM(tempIndex) Then
                        If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                            If CompararPrivilegios(UserList(UserIndex).flags.Privilegios, UserList(tempIndex).flags.Privilegios) >= 0 Then
                                Call modNetwork.Send(tempIndex, Buffer)
                            End If
                        End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToUserAreaButindex_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAreaButindex", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToUserGuildArea_Err
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        If UserList(UserIndex).GuildIndex = 0 Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    If UserList(tempIndex).ConnectionDetails.ConnIDValida And UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex Then
                        Call modNetwork.Send(tempIndex, Buffer)
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToUserGuildArea_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserGuildArea", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToNpcArea_Err
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        If NpcIndex = 0 Then Exit Sub
        Map = NpcList(NpcIndex).pos.Map
        AreaX = NpcList(NpcIndex).AreasInfo.AreaPerteneceX
        AreaY = NpcList(NpcIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
            If TempInt Then  'Esta en el area?
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                If TempInt Then
                    If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        Call modNetwork.Send(tempIndex, Buffer)
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToNpcArea_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToNpcArea", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToNpcAliveArea(ByVal NpcIndex As Long, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToNpcAliveArea(ByVal NpcIndex As Long, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToNpcArea_Err
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        If NpcIndex = 0 Then Exit Sub
        Map = NpcList(NpcIndex).pos.Map
        AreaX = NpcList(NpcIndex).AreasInfo.AreaPerteneceX
        AreaY = NpcList(NpcIndex).AreasInfo.AreaPerteneceY
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
            If TempInt Then  'Esta en el area?
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                If TempInt Then
                    If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        If UserList(tempIndex).flags.Muerto = 0 Then
                            Call modNetwork.Send(tempIndex, Buffer)
                        End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
SendToNpcArea_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToNpcArea", Erl)
    End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ParamArray Args() As Variant)
    On Error GoTo SendToAreaByPos_Err
    Dim LoopC     As Long
    Dim TempInt   As Integer
    Dim tempIndex As Integer
    AreaX = 2 ^ (AreaX \ 12)
    AreaY = 2 ^ (AreaY \ 12)
    If Not MapaValido(Map) Then Exit Sub
    #If DIRECT_PLAY = 0 Then
        Dim Buffer As Network.Writer
        Set Buffer = Protocol_Writes.GetWriterBuffer()
    #Else
        Dim Buffer As clsNetWriter
        Set Buffer = Protocol_Writes.Writer
    #End If
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                    Call modNetwork.Send(tempIndex, Buffer)
                End If
            End If
        End If
    Next LoopC
SendToAreaByPos_Err:
    Call Buffer.Clear
    If (Err.Number <> 0) Then
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToAreaByPos", Erl)
    End If
End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToMap(ByVal Map As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToMap(ByVal Map As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToMap_Err
        Dim LoopC     As Long
        Dim tempIndex As Integer
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                Call modNetwork.Send(tempIndex, Buffer)
            End If
        Next LoopC
        Exit Sub
SendToMap_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToMap", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToMapButIndex_Err
        Dim LoopC     As Long
        Dim Map       As Integer
        Dim tempIndex As Integer
        If UserIndex = 0 Then Exit Sub
        Map = UserList(UserIndex).pos.Map
        If Not MapaValido(Map) Then Exit Sub
        For LoopC = 1 To ConnGroups(Map).CountEntrys
            tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            If tempIndex <> UserIndex And UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                Call modNetwork.Send(tempIndex, Buffer)
            End If
        Next LoopC
        Exit Sub
SendToMapButIndex_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToMapButIndex", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToGroup(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToGroup(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToGroup_Err
        Dim LoopC As Long
        If UserIndex = 0 Then Exit Sub
        If Not UserList(UserIndex).Grupo.EnGrupo Then Exit Sub
        With UserList(UserList(UserIndex).Grupo.Lider.ArrayIndex).Grupo
            For LoopC = 1 To .CantidadMiembros
                If IsValidUserRef(.Miembros(LoopC)) Then
                    Call modNetwork.Send(.Miembros(LoopC).ArrayIndex, Buffer)
                End If
            Next LoopC
        End With
        Exit Sub
SendToGroup_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToGroup", Erl)
    End Sub

#If DIRECT_PLAY = 0 Then
    Private Sub SendToGroupButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
    #Else
        Private Sub SendToGroupButIndex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
        #End If
        On Error GoTo SendToGroupButIndex_Err
        Dim LoopC As Long
        If UserIndex = 0 Then Exit Sub
        If Not UserList(UserIndex).Grupo.EnGrupo Then Exit Sub
        With UserList(UserList(UserIndex).Grupo.Lider.ArrayIndex).Grupo
            For LoopC = 1 To .CantidadMiembros
                If IsValidUserRef(.Miembros(LoopC)) And .Miembros(LoopC).ArrayIndex <> UserIndex Then
                    Call modNetwork.Send(.Miembros(LoopC).ArrayIndex, Buffer)
                End If
            Next LoopC
        End With
        Exit Sub
SendToGroupButIndex_Err:
        Call TraceError(Err.Number, Err.Description, "modSendData.SendToGroupButIndex", Erl)
    End Sub
