Attribute VB_Name = "Extra"
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
Public m_NameIndex As New Dictionary

Public Function esCiudad(ByVal Map As Integer) As Boolean
    Dim i As Byte
    For i = 0 To UBound(TotalMapasCiudades)
        If TotalMapasCiudades(i) = Map Then
            esCiudad = True
            Exit Function
        End If
    Next i
End Function

Public Sub AgregarAConsola(ByVal Text As String)
    On Error GoTo AgregarAConsola_Err
    frmMain.List1.AddItem (Text)
    Exit Sub
AgregarAConsola_Err:
    Call TraceError(Err.Number, Err.Description, "ModLadder.AgregarAConsola", Erl)
End Sub

Public Function NameIndex(ByRef username As String) As t_UserReference
    Dim UserRef As t_UserReference
    Dim key     As String
    Dim wrapper As clsUserRefWrapper
    key = UCase$(Replace(username, "+", " "))
    If m_NameIndex.Exists(key) Then
        Set wrapper = m_NameIndex(key)
        UserRef.ArrayIndex = wrapper.PlayerIndex
        UserRef.VersionId = wrapper.VersionId
    Else
        Call SetUserRef(UserRef, 0)
    End If
    NameIndex = UserRef
End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByRef x As Byte, ByRef y As Byte)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 26/03/2009
    'Search for a Legal pos for the user who is being teleported.
    '***************************************************
    On Error GoTo FindLegalPos_Err
    If MapData(Map, x, y).UserIndex <> 0 Or MapData(Map, x, y).NpcIndex <> 0 Then
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(Map, x, y).UserIndex = UserIndex Then Exit Sub
        Dim FoundPlace     As Boolean
        Dim tX             As Long
        Dim tY             As Long
        Dim Rango          As Long
        Dim OtherUserIndex As Integer
        For Rango = 0 To 5
            For tY = y - Rango To y + Rango
                For tX = x - Rango To x + Rango
                    'Reviso que no haya User ni NPC
                    If MapData(Map, tX, tY).UserIndex = 0 And MapData(Map, tX, tY).NpcIndex = 0 Then
                        If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        Exit For
                    End If
                Next tX
                If FoundPlace Then Exit For
            Next tY
            If FoundPlace Then Exit For
        Next Rango
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            x = tX
            y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(Map, x, y).UserIndex
            If OtherUserIndex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If IsValidUserRef(UserList(OtherUserIndex).ComUsu.DestUsu) Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu.ArrayIndex)
                        'Msg1104= Comercio cancelado. El otro usuario se ha desconectado.
                        Call WriteLocaleMsg(UserList(OtherUserIndex).ComUsu.DestUsu.ArrayIndex, "1104", e_FontTypeNames.FONTTYPE_TALK)
                    End If
                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteShowMessageBox(OtherUserIndex, 1758, vbNullString) 'Msg1758=Alguien se ha conectado donde te encontrabas, por favor reconéctate...
                    End If
                End If
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If
    Exit Sub
FindLegalPos_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.FindLegalPos", Erl)
End Sub

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
    On Error GoTo EsNewbie_Err
    If UserIndex > 0 Then EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
    Exit Function
EsNewbie_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.EsNewbie", Erl)
End Function

Public Function esCiudadano(ByVal UserIndex As Integer) As Boolean
    On Error GoTo esCiudadano_Err
    If UserIndex > 0 Then esCiudadano = Status(UserIndex) = e_Facciones.Ciudadano
    Exit Function
esCiudadano_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.esCiudadano", Erl)
End Function

Public Function esCriminal(ByVal UserIndex As Integer) As Boolean
    On Error GoTo esCriminal_Err
    If UserIndex > 0 Then esCriminal = Status(UserIndex) = e_Facciones.Criminal
    Exit Function
esCriminal_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.esCriminal", Erl)
End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
    On Error GoTo esArmada_Err
    If UserIndex > 0 Then esArmada = (UserList(UserIndex).Faccion.Status = e_Facciones.Armada Or UserList(UserIndex).Faccion.Status = e_Facciones.consejo)
    Exit Function
esArmada_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.esArmada", Erl)
End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean
    On Error GoTo esCaos_Err
    If UserIndex > 0 Then esCaos = (UserList(UserIndex).Faccion.Status = e_Facciones.Caos Or UserList(UserIndex).Faccion.Status = e_Facciones.concilio)
    Exit Function
esCaos_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.esCaos", Erl)
End Function

Public Function FactionCanAttackFaction(ByVal AttackerFaction As e_Facciones, ByVal TargetFaction As e_Facciones)
    Select Case AttackerFaction
        Case e_Facciones.Ciudadano, e_Facciones.Armada, e_Facciones.consejo
            If TargetFaction = e_Facciones.Armada Or TargetFaction = e_Facciones.Ciudadano Or TargetFaction = e_Facciones.consejo Then
                Exit Function
            End If
        Case e_Facciones.Caos, e_Facciones.concilio
            If TargetFaction = e_Facciones.Caos Or TargetFaction = e_Facciones.concilio Then
                Exit Function
            End If
    End Select
    FactionCanAttackFaction = True
End Function

Public Function FactionCanHelpFaction(ByVal SourceFaction As e_Facciones, ByVal TargetFaction As e_Facciones) As e_InteractionResult
    Select Case SourceFaction
        Case e_Facciones.Ciudadano, e_Facciones.Armada, e_Facciones.consejo
            If TargetFaction = e_Facciones.Caos Or TargetFaction = e_Facciones.concilio Then
                FactionCanHelpFaction = eOposingFaction
                Exit Function
            ElseIf TargetFaction = e_Facciones.Criminal Then
                FactionCanHelpFaction = eCantHelpCriminal
                Exit Function
            End If
        Case e_Facciones.Caos, e_Facciones.concilio
            If TargetFaction = e_Facciones.Armada Or TargetFaction = e_Facciones.consejo Or TargetFaction = e_Facciones.Ciudadano Then
                FactionCanHelpFaction = eOposingFaction
                Exit Function
            End If
        Case Else
    End Select
    FactionCanHelpFaction = eInteractionOk
End Function

Public Function EsGM(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************
    On Error GoTo EsGM_Err
    If UserIndex = 0 Then
        EsGM = False
        Exit Function
    End If
    EsGM = (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) <> 0
    Exit Function
EsGM_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.EsGM", Erl)
End Function

Private Function CheckMapRestrictions(ByVal UserIndex As Integer, ByVal Map As Integer) As Boolean
    With UserList(UserIndex)
        If EsGM(UserIndex) Then
            CheckMapRestrictions = True
            Exit Function
        End If
        If MapInfo(Map).Newbie And Not EsNewbie(UserIndex) Then
            If .flags.UltimoMensaje <> MSG_MAP_NEWBIE_ONLY Then
                ' Msg771=Sólo los newbies pueden entrar a este mapa.
                Call WriteLocaleMsg(UserIndex, MSG_MAP_NEWBIE_ONLY, e_FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = MSG_MAP_NEWBIE_ONLY
            End If
            Exit Function
        End If
        If MapInfo(Map).NoPKs And (Status(UserIndex) = 0 Or Status(UserIndex) = 2) Then
            If .flags.UltimoMensaje <> MSG_MAP_ONLY_CITIZENS Then
                ' Msg772=Sólo los ciudadanos pueden entrar a este mapa.
                Call WriteLocaleMsg(UserIndex, MSG_MAP_ONLY_CITIZENS, e_FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = MSG_MAP_ONLY_CITIZENS
            End If
            Exit Function
        End If
        If MapInfo(Map).NoCiudadanos And (Status(UserIndex) = 1 Or Status(UserIndex) = 3) Then
            If .flags.UltimoMensaje <> MSG_MAP_ONLY_CRIMINALS Then
                ' Msg773=Sólo los criminales pueden entrar a este mapa.
                Call WriteLocaleMsg(UserIndex, MSG_MAP_ONLY_CRIMINALS, e_FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = MSG_MAP_ONLY_CRIMINALS
            End If
            Exit Function
        End If
        If MapInfo(Map).SoloClanes And .GuildIndex <= 0 Then
            If .flags.UltimoMensaje <> MSG_MAP_REQUIRES_CLAN Then
                ' Msg774=Necesitas pertenecer a un clan para entrar a este mapa.
                Call WriteLocaleMsg(UserIndex, MSG_MAP_REQUIRES_CLAN, e_FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = MSG_MAP_REQUIRES_CLAN
            End If
            Exit Function
        End If
        If MapInfo(Map).MinLevel <> 0 And .Stats.ELV < MapInfo(Map).MinLevel Then
            If .flags.UltimoMensaje <> MSG_MAP_MIN_LEVEL Then
                'Msg1108= Necesitas ser al menos nivel ¬1
                Call WriteLocaleMsg(UserIndex, MSG_MAP_MIN_LEVEL, e_FontTypeNames.FONTTYPE_INFO, MapInfo(Map).MinLevel)
                .flags.UltimoMensaje = MSG_MAP_MIN_LEVEL
            End If
            Exit Function
        End If
        If MapInfo(Map).MaxLevel <> 0 And .Stats.ELV >= MapInfo(Map).MaxLevel Then
            If .flags.UltimoMensaje <> MSG_MAP_MAX_LEVEL Then
                'Msg1109= Sólo los personajes inferiores a nivel ¬1
                Call WriteLocaleMsg(UserIndex, MSG_MAP_MAX_LEVEL, e_FontTypeNames.FONTTYPE_INFO, MapInfo(Map).MaxLevel)
                .flags.UltimoMensaje = MSG_MAP_MAX_LEVEL
            End If
            Exit Function
        End If
        If MapInfo(Map).OnlyGroups And Not .Grupo.EnGrupo Then
            If .flags.UltimoMensaje <> MSG_MAP_REQUIRES_GROUP Then
                ' Msg775=Necesitas pertenecer a un grupo para entrar a este mapa.
                Call WriteLocaleMsg(UserIndex, MSG_MAP_REQUIRES_GROUP, e_FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = MSG_MAP_REQUIRES_GROUP
            End If
            Exit Function
        End If
        If MapInfo(Map).OnlyPatreon And Not (.Stats.tipoUsuario = tAventurero Or .Stats.tipoUsuario = tHeroe Or .Stats.tipoUsuario = tLeyenda) Then
            If .flags.UltimoMensaje <> MSG_MAP_REQUIRES_PATREON Then
                ' Msg776=Necesitas ser Patreon para entrar a este mapa.
                Call WriteLocaleMsg(UserIndex, MSG_MAP_REQUIRES_PATREON, e_FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = MSG_MAP_REQUIRES_PATREON
            End If
            Exit Function
        End If
        CheckMapRestrictions = True
    End With
End Function

Public Function GetTransportNextIndex(ByVal Map As Integer, ByVal PosX As Byte, ByVal PosY As Byte) As Integer
    Dim i As Integer
    With MapInfo(Map)
        For i = 0 To UBound(.TransportNetwork)
            If .TransportNetwork(i).TileX = PosX And .TransportNetwork(i).TileY = PosY Then
                GetTransportNextIndex = i
                Exit Function
            End If
        Next i
    End With
    GetTransportNextIndex = -1
End Function

Public Function GetExitTransport(ByVal Map As Integer, ByVal ExcludeIndex As Integer) As Integer
    Dim Output As Integer
    Do
        Output = RandomNumber(0, UBound(MapInfo(Map).TransportNetwork))
    Loop While Output = ExcludeIndex
    GetExitTransport = Output
End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 23/01/2007
    'Handles the Map passage of Users. Allows the existance
    'of exclusive maps for Newbies, Royal Army and Caos Legion members
    'and enables GMs to enter every map without restriction.
    'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
    '***************************************************
    On Error GoTo ErrHandler
    Dim nPos       As t_WorldPos
    Dim EsTeleport As Boolean
    Dim TelepRadio As Byte
    Dim aN         As Integer
    Dim destPos    As t_WorldPos
    With UserList(UserIndex)
        'Controla las salidas
        If InMapBounds(Map, x, y) Then
            If MapData(Map, x, y).trigger = e_Trigger.TRANSFER_ONLY_DEAD Then
                If .flags.Muerto <> 1 Then Exit Sub  ' si está vivo, no teletransportar
            End If
            If MapData(Map, x, y).trigger = AUTORESU Then
                Call ResucitarOCurar(UserIndex)
            End If
            If MapData(Map, x, y).ObjInfo.ObjIndex > 0 Then
                EsTeleport = ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).OBJType = e_OBJType.otTeleport
            End If
            If Not MapData(Map, x, y).Trap Is Nothing Then
                Call ModMap.ActivateTrap(UserIndex, eUser, Map, x, y)
            End If
            If EsTeleport Then
                If ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).Subtipo = e_TeleportSubType.eTransportNetwork Then
                    Dim StartTransportIndex As Integer
                    Dim ExitPortal          As Integer
                    StartTransportIndex = GetTransportNextIndex(Map, x, y)
                    If .LastTransportNetwork.Map = Map And .LastTransportNetwork.ExitIndex = StartTransportIndex Then
                        ExitPortal = .LastTransportNetwork.StartIdex
                    Else
                        ExitPortal = GetExitTransport(Map, StartTransportIndex)
                    End If
                    destPos = MapData(Map, MapInfo(Map).TransportNetwork(ExitPortal).TileX, MapInfo(Map).TransportNetwork(ExitPortal).TileY).TileExit
                    If destPos.Map > 0 And destPos.Map <= NumMaps Then
                        .LastTransportNetwork.Map = Map
                        .LastTransportNetwork.StartIdex = StartTransportIndex
                        .LastTransportNetwork.ExitIndex = ExitPortal
                        Call WarpUserChar(UserIndex, destPos.Map, destPos.x, destPos.y, EsTeleport)
                    Else
                        Call LogError("Invalid teleport at map: " & Map & "(" & x & ", " & y & ")")
                    End If
                    Exit Sub
                End If
            End If
            If (MapData(Map, x, y).TileExit.Map > 0) And (MapData(Map, x, y).TileExit.Map <= NumMaps) Then
                '  Restricciones de mapas
                If CheckMapRestrictions(UserIndex, MapData(Map, x, y).TileExit.Map) Then
                    If EsMapaInterdimensional(MapData(Map, x, y).TileExit.Map) And Not EsMapaInterdimensional(.pos.Map) Then
                        .flags.ReturnPos = .pos
                    End If
                    destPos.Map = MapData(Map, x, y).TileExit.Map
                    If EsTeleport Then
                        destPos.x = RandomNumber(MapData(Map, x, y).TileExit.x - ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).Radio, MapData(Map, x, y).TileExit.x + ObjData( _
                                MapData(Map, x, y).ObjInfo.ObjIndex).Radio)
                        destPos.y = RandomNumber(MapData(Map, x, y).TileExit.y - ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).Radio, MapData(Map, x, y).TileExit.y + ObjData( _
                                MapData(Map, x, y).ObjInfo.ObjIndex).Radio)
                    Else
                        destPos.x = MapData(Map, x, y).TileExit.x
                        destPos.y = MapData(Map, x, y).TileExit.y
                    End If
                    If .flags.Navegando Then
                        Call ClosestLegalPos(destPos, nPos, True)
                    Else
                        Call ClosestLegalPos(destPos, nPos)
                    End If
                    If nPos.x <> 0 And nPos.y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, EsTeleport)
                    End If
                End If
                'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                Call ClearAttackerNpc(UserIndex)
            ElseIf MapData(Map, x, y).TileExit.Map < 0 Then
                If .flags.ReturnPos.Map <> 0 Then
                    If LegalPos(.flags.ReturnPos.Map, .flags.ReturnPos.x, .flags.ReturnPos.y, .flags.Navegando = 1, , , False) Then
                        Call WarpUserChar(UserIndex, .flags.ReturnPos.Map, .flags.ReturnPos.x, .flags.ReturnPos.y, False)
                    Else
                        Call ClosestLegalPos(.flags.ReturnPos, nPos)
                        If nPos.x <> 0 And nPos.y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, EsTeleport)
                        End If
                    End If
                    'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                    Call ClearAttackerNpc(UserIndex)
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, ".DotileEvents", Erl)
End Sub

Public Sub ClearAttackerNpc(ByVal UserIndex As Integer)
    On Error GoTo ClearAttackerNpc_err
    With UserList(UserIndex)
        Dim aN As Integer
        If Not IsValidNpcRef(.flags.AtacadoPorNpc) Then
            Call ClearNpcRef(.flags.AtacadoPorNpc)
        Else
            aN = .flags.AtacadoPorNpc.ArrayIndex
            If aN > 0 Then
                If IsValidUserRef(NpcList(aN).TargetUser) And NpcList(aN).TargetUser.ArrayIndex = UserIndex Then
                    Call SetMovement(aN, NpcList(aN).flags.OldMovement)
                    NpcList(aN).Hostile = NpcList(aN).flags.OldHostil
                    NpcList(aN).flags.AttackedBy = vbNullString
                    Call SetUserRef(NpcList(aN).TargetUser, 0)
                End If
            End If
        End If
        If Not IsValidNpcRef(.flags.NPCAtacado) Then
            Call ClearNpcRef(.flags.NPCAtacado)
        Else
            aN = .flags.NPCAtacado.ArrayIndex
            If aN > 0 Then
                If NpcList(aN).flags.AttackedFirstBy = .name Then
                    NpcList(aN).flags.AttackedFirstBy = vbNullString
                End If
            End If
        End If
        Call ClearNpcRef(.flags.AtacadoPorNpc)
        Call ClearNpcRef(.flags.NPCAtacado)
    End With
    Exit Sub
ClearAttackerNpc_err:
    Call TraceError(Err.Number, Err.Description, "Extra.ClearAttackerNpc failed to clear userindex" & UserIndex & "(" & UserList(UserIndex).VersionId & ")", Erl)
End Sub

Function InRangoVision(ByVal UserIndex As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo InRangoVision_Err
    If x > UserList(UserIndex).pos.x - MinXBorder And x < UserList(UserIndex).pos.x + MinXBorder Then
        If y > UserList(UserIndex).pos.y - MinYBorder And y < UserList(UserIndex).pos.y + MinYBorder Then
            InRangoVision = True
            Exit Function
        End If
    End If
    InRangoVision = False
    Exit Function
InRangoVision_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.InRangoVision", Erl)
End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo InRangoVisionNPC_Err
    If x > NpcList(NpcIndex).pos.x - MinXBorder And x < NpcList(NpcIndex).pos.x + MinXBorder Then
        If y > NpcList(NpcIndex).pos.y - MinYBorder And y < NpcList(NpcIndex).pos.y + MinYBorder Then
            InRangoVisionNPC = True
            Exit Function
        End If
    End If
    InRangoVisionNPC = False
    Exit Function
InRangoVisionNPC_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.InRangoVisionNPC", Erl)
End Function

Function InMapBounds(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo InMapBounds_Err
    If (Map <= 0 Or Map > NumMaps) Or x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True
    End If
    Exit Function
InMapBounds_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.InMapBounds", Erl)
End Function

Public Function TileRequiresPatreon(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo TileRequiresPatreon_Err
    If Not InMapBounds(Map, x, y) Then Exit Function
    TileRequiresPatreon = MapData(Map, x, y).trigger = e_Trigger.ONLY_PATREON_TILE
    Exit Function
TileRequiresPatreon_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.TileRequiresPatreon", Erl)
End Function

Function ClosestLegalPosNPC(ByVal NpcIndex As Integer, ByVal MaxRange As Integer, Optional ByVal IgnoreUsers As Boolean, Optional ByVal IgnoreDeadUsers As Boolean) As t_WorldPos
    On Error GoTo ErrHandler
    Dim LoopC As Integer
    Dim tX    As Integer
    Dim tY    As Integer
    With NpcList(NpcIndex)
        Do
            tY = .pos.y - LoopC
            For tX = .pos.x - LoopC To .pos.x + LoopC
                If ValidNPCSpawnPos(ClosestLegalPosNPC, .pos.Map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers, IgnoreDeadUsers) Then
                    Exit Function
                End If
            Next
            tX = .pos.x - LoopC
            For tY = .pos.y - LoopC + 1 To .pos.y + LoopC - 1
                If ValidNPCSpawnPos(ClosestLegalPosNPC, .pos.Map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers, IgnoreDeadUsers) Then
                    Exit Function
                End If
            Next
            tX = .pos.x + LoopC
            For tY = .pos.y - LoopC + 1 To .pos.y + LoopC - 1
                If ValidNPCSpawnPos(ClosestLegalPosNPC, .pos.Map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers, IgnoreDeadUsers) Then
                    Exit Function
                End If
            Next
            tY = .pos.y + LoopC
            For tX = .pos.x - LoopC To .pos.x + LoopC
                If ValidNPCSpawnPos(ClosestLegalPosNPC, .pos.Map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers, IgnoreDeadUsers) Then
                    Exit Function
                End If
            Next
            LoopC = LoopC + 1
        Loop While LoopC <= MaxRange
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Extra.ClosestLegalPosNPC")
End Function

Private Function ValidNPCSpawnPos(OutPos As t_WorldPos, _
                                  ByVal Map As Integer, _
                                  ByVal x As Integer, _
                                  ByVal y As Integer, _
                                  ByVal AguaValida As Boolean, _
                                  ByVal TierraValida As Boolean, _
                                  ByVal IgnoreUsers As Boolean, _
                                  ByVal IgnoreDeadUsers As Boolean) As Boolean
    If LegalPos(Map, x, y, AguaValida, TierraValida, , False) Then
        If TestSpawnTrigger(Map, x, y) Then
            If Not HayPCarea(Map, x, y, IgnoreDeadUsers) Or IgnoreUsers Then
                ValidNPCSpawnPos = True
                OutPos.Map = Map
                OutPos.x = x
                OutPos.y = y
                Exit Function
            End If
        End If
    End If
End Function

Sub ClosestLegalPos(pos As t_WorldPos, ByRef nPos As t_WorldPos, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True)
    '*****************************************************************
    'Author: Unknown (original version)
    'Last Modification: 24/01/2007 (ToxicWaste)
    'Encuentra la posicion legal mas cercana y la guarda en nPos
    '*****************************************************************
    On Error GoTo ClosestLegalPos_Err
    Dim Notfound As Boolean
    Dim LoopC    As Integer
    Dim tX       As Integer
    Dim tY       As Integer
    nPos.Map = pos.Map
    Do While Not LegalPos(pos.Map, nPos.x, nPos.y, PuedeAgua, PuedeTierra, , False)
        If LoopC > 12 Then
            Notfound = True
            Exit Do
        End If
        For tY = pos.y - LoopC To pos.y + LoopC
            For tX = pos.x - LoopC To pos.x + LoopC
                If LegalPos(nPos.Map, tX, tY, PuedeAgua, PuedeTierra, , False) Then
                    nPos.x = tX
                    nPos.y = tY
                    Exit Sub
                End If
            Next tX
        Next tY
        LoopC = LoopC + 1
    Loop
    If Notfound = True Then
        nPos.x = 0
        nPos.y = 0
    End If
    Exit Sub
ClosestLegalPos_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.ClosestLegalPos", Erl)
End Sub

Sub ClosestStablePos(pos As t_WorldPos, ByRef nPos As t_WorldPos)
    '*****************************************************************
    'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
    '*****************************************************************
    On Error GoTo ClosestStablePos_Err
    Dim Notfound As Boolean
    Dim LoopC    As Integer
    Dim tX       As Integer
    Dim tY       As Integer
    nPos.Map = pos.Map
    Do While Not LegalPos(pos.Map, nPos.x, nPos.y)
        If LoopC > 12 Then
            Notfound = True
            Exit Do
        End If
        For tY = pos.y - LoopC To pos.y + LoopC
            For tX = pos.x - LoopC To pos.x + LoopC
                If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                    nPos.x = tX
                    nPos.y = tY
                    '¿Hay objeto?
                    tX = pos.x + LoopC
                    tY = pos.y + LoopC
                End If
            Next tX
        Next tY
        LoopC = LoopC + 1
    Loop
    If Notfound = True Then
        nPos.x = 0
        nPos.y = 0
    End If
    Exit Sub
ClosestStablePos_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.ClosestStablePos", Erl)
End Sub

Function IP_Index(ByVal inIP As String) As Integer
    On Error GoTo IP_Index_Err
    Dim UserIndex As Integer
    '¿Nombre valido?
    If LenB(inIP) = 0 Then
        IP_Index = 0
        Exit Function
    End If
    UserIndex = 1
    Do Until UserList(UserIndex).ConnectionDetails.IP = inIP
        UserIndex = UserIndex + 1
        If UserIndex > MaxUsers Then
            IP_Index = 0
            Exit Function
        End If
    Loop
    IP_Index = UserIndex
    Exit Function
    Exit Function
IP_Index_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.IP_Index", Erl)
End Function

Sub HeadtoPos(ByVal head As e_Heading, ByRef pos As t_WorldPos)
    On Error GoTo HeadtoPos_Err
    '*****************************************************************
    'Toma una posicion y se mueve hacia donde esta perfilado
    '*****************************************************************
    Dim x  As Integer
    Dim y  As Integer
    Dim nX As Integer
    Dim nY As Integer
    x = pos.x
    y = pos.y
    If head = e_Heading.NORTH Then
        nX = x
        nY = y - 1
    End If
    If head = e_Heading.SOUTH Then
        nX = x
        nY = y + 1
    End If
    If head = e_Heading.EAST Then
        nX = x + 1
        nY = y
    End If
    If head = e_Heading.WEST Then
        nX = x - 1
        nY = y
    End If
    'Devuelve valores
    pos.x = nX
    pos.y = nY
    Exit Sub
HeadtoPos_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.HeadtoPos", Erl)
End Sub

'Returns the front left positon from current heading
Public Sub GetHeadingLeft(ByVal head As e_Heading, ByRef pos As t_WorldPos)
    Dim x, y, nX, nY As Integer
    x = pos.x
    y = pos.y
    If head = e_Heading.NORTH Then
        nX = x - 1
        nY = y - 1
    End If
    If head = e_Heading.SOUTH Then
        nX = x + 1
        nY = y + 1
    End If
    If head = e_Heading.EAST Then
        nX = x + 1
        nY = y - 1
    End If
    If head = e_Heading.WEST Then
        nX = x - 1
        nY = y + 1
    End If
    'Devuelve valores
    pos.x = nX
    pos.y = nY
End Sub

'Returns the front right positon from current heading
Sub GetHeadingRight(ByVal head As e_Heading, ByRef pos As t_WorldPos)
    Dim x, y, nX, nY As Integer
    x = pos.x
    y = pos.y
    If head = e_Heading.NORTH Then
        nX = x + 1
        nY = y - 1
    End If
    If head = e_Heading.SOUTH Then
        nX = x - 1
        nY = y + 1
    End If
    If head = e_Heading.EAST Then
        nX = x + 1
        nY = y + 1
    End If
    If head = e_Heading.WEST Then
        nX = x - 1
        nY = y - 1
    End If
    'Devuelve valores
    pos.x = nX
    pos.y = nY
End Sub

' Autor: WyroX - 20/01/2021
' Retorna el heading recibo como parámetro pero rotado, según el valor R.
' Si R es 1, rota en sentido horario. Si R es -1, en sentido antihorario.
Function Rotate_Heading(ByVal Heading As e_Heading, ByVal r As Integer) As e_Heading
    Rotate_Heading = (Heading + r + 3) Mod 4 + 1
End Function

Function LegalPos(ByVal Map As Integer, _
                  ByVal x As Integer, _
                  ByVal y As Integer, _
                  Optional ByVal PuedeAgua As Boolean = False, _
                  Optional ByVal PuedeTierra As Boolean = True, _
                  Optional ByVal Montado As Boolean = False, _
                  Optional ByVal PuedeTraslado As Boolean = True, _
                  Optional ByVal PuedeBloqueoParcial As Boolean = True) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 23/01/2007
    'Checks if the position is Legal.
    '***************************************************
    '¿Es un mapa valido?
    On Error GoTo LegalPos_Err
    If Map <= 0 Or Map > NumMaps Then Exit Function
    If x < MinXBorder Or x > MaxXBorder Then Exit Function
    If y < MinYBorder Or y > MaxYBorder Then Exit Function
    With MapData(Map, x, y)
        If .NpcIndex <> 0 Then Exit Function
        If .UserIndex <> 0 Then Exit Function
        If Not PuedeTraslado Then
            If .TileExit.Map > 0 Then Exit Function
        End If
        If Not PuedeAgua Then
            If (.Blocked And FLAG_AGUA) <> 0 Then Exit Function
        End If
        If Not PuedeTierra Then
            If (.Blocked And FLAG_AGUA) = 0 Then Exit Function
        End If
        If PuedeBloqueoParcial Then
            If (.Blocked And e_Block.ALL_SIDES) = e_Block.ALL_SIDES Then Exit Function
        Else
            If (.Blocked And e_Block.ALL_SIDES) > 0 Then Exit Function
        End If
    End With
    LegalPos = True
    Exit Function
LegalPos_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.LegalPos", Erl)
End Function

Function LegalPosDestrabar(ByVal Map As Integer, _
                           ByVal x As Integer, _
                           ByVal y As Integer, _
                           Optional ByVal PuedeAgua As Boolean = False, _
                           Optional ByVal PuedeTierra As Boolean = True, _
                           Optional ByVal Montado As Boolean = False, _
                           Optional ByVal PuedeTraslado As Boolean = True, _
                           Optional ByVal PuedeBloqueoParcial As Boolean = True) As Boolean
    On Error GoTo LegalPosDestrabar_Err
    If Map <= 0 Or Map > NumMaps Then Exit Function
    If x < MinXBorder Or x > MaxXBorder Then Exit Function
    If y < MinYBorder Or y > MaxYBorder Then Exit Function
    With MapData(Map, x, y)
        If .NpcIndex <> 0 Then Exit Function
        If Not PuedeTraslado Then
            If .TileExit.Map > 0 Then Exit Function
        End If
        If Not PuedeAgua Then
            If (.Blocked And FLAG_AGUA) <> 0 Then Exit Function
        End If
        If Not PuedeTierra Then
            If (.Blocked And FLAG_AGUA) = 0 Then Exit Function
        End If
        If PuedeBloqueoParcial Then
            If (.Blocked And e_Block.ALL_SIDES) = e_Block.ALL_SIDES Then Exit Function
        Else
            If (.Blocked And e_Block.ALL_SIDES) > 0 Then Exit Function
        End If
    End With
    LegalPosDestrabar = True
    Exit Function
LegalPosDestrabar_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.LegalPosDestrabar", Erl)
End Function

Function LegalWalk(ByVal Map As Integer, _
                   ByVal x As Integer, _
                   ByVal y As Integer, _
                   ByVal Heading As e_Heading, _
                   Optional ByVal PuedeAgua As Boolean = False, _
                   Optional ByVal PuedeTierra As Boolean = True, _
                   Optional ByVal Montado As Boolean = False, _
                   Optional ByVal PuedeTraslado As Boolean = True, _
                   Optional ByVal WalkerIndex As Integer) As Boolean
    On Error GoTo LegalWalk_Err
    If Map <= 0 Or Map > NumMaps Then Exit Function
    If x < MinXBorder Or x > MaxXBorder Then Exit Function
    If y < MinYBorder Or y > MaxYBorder Then Exit Function
    With MapData(Map, x, y)
        If .NpcIndex <> 0 Then
            If Not IsSet(NpcList(.NpcIndex).flags.StatusMask, e_StatusMask.eDontBlockTile) Then Exit Function
        End If
        If .UserIndex <> 0 Then
            If UserList(.UserIndex).flags.AdminInvisible = 0 And UserList(.UserIndex).flags.Muerto = 0 Then
                Exit Function
            End If
        End If
        If Not PuedeTraslado Then
            If .TileExit.Map > 0 Then Exit Function
        End If
        If Not PuedeAgua Then
            If (.Blocked And FLAG_AGUA) <> 0 And Not .trigger = e_Trigger.VALIDOPUENTE Then Exit Function
        End If
        If Not PuedeTierra Then
            If (.Blocked And FLAG_AGUA) = 0 Then Exit Function
        End If
        If .trigger = WORKERONLY Then
            If Not UserList(WalkerIndex).clase = Trabajador Then Exit Function
        End If
        If WalkerIndex <> 0 Then
            If TileRequiresPatreon(Map, x, y) Then
                If Not EsGM(WalkerIndex) And Not IsPatreon(WalkerIndex) Then
                    If UserList(WalkerIndex).flags.UltimoMensaje <> MSG_TILE_REQUIRES_PATREON Then
                        Call WriteLocaleMsg(WalkerIndex, MSG_TILE_REQUIRES_PATREON, e_FontTypeNames.FONTTYPE_INFO)
                        UserList(WalkerIndex).flags.UltimoMensaje = MSG_TILE_REQUIRES_PATREON
                    End If
                    Exit Function
                End If
            End If
        End If
        If (.Blocked And 2 ^ (Heading - 1)) <> 0 Then Exit Function
    End With
    LegalWalk = True
    Exit Function
LegalWalk_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.LegalWalk", Erl)
End Function

Function LegalPosNPC(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean
    On Error GoTo LegalPosNPC_Err
    If (Map <= 0 Or Map > NumMaps) Or (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
        LegalPosNPC = False
    ElseIf MapData(Map, x, y).TileExit.Map > 0 Then
        LegalPosNPC = False
    Else
        If AguaValida = 0 Then
            LegalPosNPC = (MapData(Map, x, y).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES And (MapData(Map, x, y).UserIndex = 0) And (MapData(Map, x, y).NpcIndex = 0) _
                    And (MapData(Map, x, y).trigger <> e_Trigger.POSINVALIDA Or IsPet) And (MapData(Map, x, y).Blocked And FLAG_AGUA) = 0
        Else
            LegalPosNPC = (MapData(Map, x, y).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES And (MapData(Map, x, y).UserIndex = 0) And (MapData(Map, x, y).NpcIndex = 0) _
                    And (MapData(Map, x, y).trigger <> e_Trigger.POSINVALIDA Or IsPet)
        End If
    End If
    Exit Function
LegalPosNPC_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.LegalPosNPC", Erl)
End Function

Function LegalWalkNPC(ByVal Map As Integer, _
                      ByVal x As Integer, _
                      ByVal y As Integer, _
                      ByVal Heading As e_Heading, _
                      Optional ByVal PuedeAgua As Boolean = False, _
                      Optional ByVal PuedeTierra As Boolean = True, _
                      Optional ByVal IgnoraInvalida As Boolean = False, _
                      Optional ByVal PuedePisar As Boolean, _
                      Optional ByVal esGuardia As Boolean = False) As Boolean
    ' Reescrito por WyroX
    On Error GoTo LegalWalkNPC_Err
    If (Map <= 0 Or Map > NumMaps) Or (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
        Exit Function
    End If
    With MapData(Map, x, y)
        If .TileExit.Map Then Exit Function
        If Not PuedeAgua Then
            If .Blocked And FLAG_AGUA Then
                Exit Function
            End If
        End If
        If Not PuedeTierra Then
            If (.Blocked And FLAG_AGUA) = 0 Then
                Exit Function
            End If
        End If
        If Not PuedePisar Then
            If .UserIndex Then
                If UserList(.UserIndex).flags.AdminInvisible = 0 And UserList(.UserIndex).flags.Muerto = 0 Then
                    Exit Function
                End If
            ElseIf .NpcIndex Then
                Exit Function
            End If
        End If
        If Not IgnoraInvalida Then
            If .trigger = e_Trigger.POSINVALIDA Then
                Exit Function
            End If
        End If
        If Not esGuardia Then
            If .Blocked And 2 ^ (Heading - 1) Then
                Exit Function
            End If
        Else
            If (.Blocked And 2 ^ (Heading - 1)) And Not HayPuerta(Map, x + 1, y) And Not HayPuerta(Map, x, y) And Not HayPuerta(Map, x + 1, y - 1) And Not HayPuerta(Map, x, y - _
                    1) Then Exit Function
        End If
    End With
    LegalWalkNPC = True
    Exit Function
LegalWalkNPC_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.LegalWalkNPC", Erl)
End Function

Sub SendHelp(ByVal Index As Integer)
    On Error GoTo SendHelp_Err
    Dim NumHelpLines As Integer
    Dim LoopC        As Integer
    NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))
    For LoopC = 1 To NumHelpLines
        Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), e_FontTypeNames.FONTTYPE_INFO)
    Next LoopC
    Exit Sub
SendHelp_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.SendHelp", Erl)
End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    On Error GoTo Expresar_Err
    If NpcList(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, NpcList(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(NpcList(NpcIndex).Expresiones(randomi), NpcList(NpcIndex).Char.charindex, vbWhite))
    End If
    Exit Sub
Expresar_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.Expresar", Erl)
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo LookatTile_Err
    'Responde al click del usuario sobre el mapa
    Dim FoundChar       As Byte
    Dim FoundSomething  As Byte
    Dim TempCharIndex   As Integer
    Dim extraStrings    As String
    Dim ft              As e_FontTypeNames
    Dim Statuses        As Long
    Dim FactionStatuses As Long
    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(UserIndex).pos.y - y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).pos.x - x) > RANGO_VISION_X) Then
        Exit Sub
    End If
    '¿Posicion valida?
    If InMapBounds(Map, x, y) Then
        UserList(UserIndex).flags.TargetMap = Map
        UserList(UserIndex).flags.TargetX = x
        UserList(UserIndex).flags.TargetY = y
        '¿Es un obj?
        If MapData(Map, x, y).ObjInfo.ObjIndex > 0 Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = x
            UserList(UserIndex).flags.TargetObjY = y
            FoundSomething = 1
        ElseIf MapData(Map, x + 1, y).ObjInfo.ObjIndex > 0 Then
            'Informa el nombre
            If ObjData(MapData(Map, x + 1, y).ObjInfo.ObjIndex).OBJType = e_OBJType.otDoors Then
                UserList(UserIndex).flags.TargetObjMap = Map
                UserList(UserIndex).flags.TargetObjX = x + 1
                UserList(UserIndex).flags.TargetObjY = y
                FoundSomething = 1
            End If
        ElseIf MapData(Map, x + 1, y + 1).ObjInfo.ObjIndex > 0 Then
            If ObjData(MapData(Map, x + 1, y + 1).ObjInfo.ObjIndex).OBJType = e_OBJType.otDoors Then
                'Informa el nombre
                UserList(UserIndex).flags.TargetObjMap = Map
                UserList(UserIndex).flags.TargetObjX = x + 1
                UserList(UserIndex).flags.TargetObjY = y + 1
                FoundSomething = 1
            End If
        ElseIf MapData(Map, x, y + 1).ObjInfo.ObjIndex > 0 Then
            If ObjData(MapData(Map, x, y + 1).ObjInfo.ObjIndex).OBJType = e_OBJType.otDoors Then
                'Informa el nombre
                UserList(UserIndex).flags.TargetObjMap = Map
                UserList(UserIndex).flags.TargetObjX = x
                UserList(UserIndex).flags.TargetObjY = y + 1
                FoundSomething = 1
            End If
        End If
        If FoundSomething = 1 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
            If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then
                Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "* - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList( _
                        UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount & " *" & (MapData(Map, x, y).ObjInfo.ElementalTags Or ObjData(MapData( _
                        Map, x, y).ObjInfo.ObjIndex).ElementalTags) & "*" & "", e_FontTypeNames.FONTTYPE_INFO)
            Else
                If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = e_OBJType.otOreDeposit Then
                    Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
                    Call WriteLocaleMsg(UserIndex, 1618, e_FontTypeNames.FONTTYPE_INFO, ObjData(UserList(UserIndex).flags.TargetObj).name & "¬" & (MapData(Map, UserList( _
                            UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount))   'Msg1618=¬1 - (Minerales disponibles: ¬2)
                ElseIf ObjData(UserList(UserIndex).flags.TargetObj).OBJType = e_OBJType.otTrees Then
                    Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
                    Call WriteLocaleMsg(UserIndex, 1619, e_FontTypeNames.FONTTYPE_INFO, ObjData(UserList(UserIndex).flags.TargetObj).name & "¬" & (MapData(Map, UserList( _
                            UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount)) 'Msg1619=¬1 - (Recursos disponibles: ¬2)
                ElseIf ObjData(UserList(UserIndex).flags.TargetObj).OBJType = e_OBJType.otTeleport Then
                    If MapData(Map, x, y).TileExit.Map > 0 Then
                        If LenB(MapInfo(MapData(Map, x, y).TileExit.Map).map_name) <> 0 Then
                            Call WriteLocaleMsg(UserIndex, 1620, e_FontTypeNames.FONTTYPE_INFO, MapInfo(MapData(Map, x, y).TileExit.Map).map_name) 'Msg1620=Portal a ¬1
                        Else
                            'Msg492=Portal a un mapa desconocido...
                            Call WriteLocaleMsg(UserIndex, 492, e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
                        'Msg493=Portal hacia la quinta dimensión
                        Call WriteLocaleMsg(UserIndex, 493, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "*", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        Dim yy As Byte
        If y + 1 <= YMaxMapSize Then
            'it's a character?
            If MapData(Map, x, y + 1).UserIndex > 0 Then
                TempCharIndex = MapData(Map, x, y + 1).UserIndex
                yy = y + 1
                FoundChar = 1
            End If
            'it's an npc?
            If MapData(Map, x, y + 1).NpcIndex > 0 Then
                TempCharIndex = MapData(Map, x, y + 1).NpcIndex
                FoundChar = 2
            End If
        End If
        If FoundChar = 0 Then
            'it's a character?
            If MapData(Map, x, y).UserIndex > 0 Then
                TempCharIndex = MapData(Map, x, y).UserIndex
                yy = y
                FoundChar = 1
            End If
            'it's an npc?
            If MapData(Map, x, y).NpcIndex > 0 Then
                TempCharIndex = MapData(Map, x, y).NpcIndex
                FoundChar = 2
            End If
        End If
        'Reaccion al personaje
        If FoundChar = 1 Then '  ¿Encontro un Usuario?
            If UserList(TempCharIndex).pos.x <> x Or UserList(TempCharIndex).pos.y <> yy Then
                'Mapblock bugeado
                MapData(Map, x, yy).UserIndex = 0
                Call TraceError(999, "Hay un mapblock con un usuario y el usuario tiene otra posicion. Index: " & TempCharIndex & ", user: " & UserList(TempCharIndex).name & _
                        ", mapblock: " & Map & "," & x & "," & y & ", userpos: " & UserList(TempCharIndex).pos.Map & "," & UserList(TempCharIndex).pos.x & "," & UserList( _
                        TempCharIndex).pos.y, "LookatTile")
            End If
            If UserList(TempCharIndex).flags.AdminInvisible = 0 Or CompararPrivilegiosUser(UserIndex, TempCharIndex) >= 0 Then
                If UserList(TempCharIndex).showName Or CompararPrivilegiosUser(UserIndex, TempCharIndex) >= 0 Then
                    extraStrings = PrepareUserStatusEffectMsgsForPlayers(TempCharIndex, UserIndex, Statuses, FactionStatuses, ft)
                    If LenB(extraStrings) > 0 Then
                        'if im not dead or (i have guild and the target is a guildmate) or im clicking myself
                        If UserList(UserIndex).flags.Muerto = 0 Or (UserList(UserIndex).GuildIndex > 0 And UserList(UserIndex).GuildIndex = UserList(TempCharIndex).GuildIndex) _
                                Or UserIndex = TempCharIndex Then
                            Call WriteLocaleMsg(UserIndex, 1105, ft, extraStrings & "¬" & Statuses & "¬" & FactionStatuses)
                        End If
                    End If
                    FoundSomething = 1
                    Call SetUserRef(UserList(UserIndex).flags.TargetUser, TempCharIndex)
                    Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
                    UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
                End If
            End If
        End If
        If FoundChar = 2 Then '¿Encontro un NPC?
            Dim estatus       As String
            Dim NpcStatusMask As Long
            If Len(NpcList(TempCharIndex).Desc) > 1 Then
                '  Hacemos que se detenga a hablar un momento :P
                If NpcList(TempCharIndex).Movement = Caminata Then
                    NpcList(TempCharIndex).Contadores.IntervaloMovimiento = AddMod32(GetTickCountRaw(), 5000 + Len(NpcList(TempCharIndex).Desc) * 50) ' 5 segundos + 1 segundo cada 20 caracteres
                End If
                If UserList(UserIndex).flags.Muerto = 0 Or (UserList(UserIndex).flags.Muerto = 1 And NpcList(TempCharIndex).npcType = e_NPCType.Revividor) Then
                    If NpcList(TempCharIndex).npcType = e_NPCType.Quest Or NpcList(TempCharIndex).npcType = e_NPCType.Banquero Or NpcList(TempCharIndex).npcType = _
                       e_NPCType.Revividor Or NpcList(TempCharIndex).npcType = e_NPCType.Comun Or NpcList(TempCharIndex).npcType = e_NPCType.Entrenador Or NpcList( _
                       TempCharIndex).npcType = e_NPCType.Gobernador Then
                        If Distance(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, NpcList(TempCharIndex).pos.x, NpcList(TempCharIndex).pos.y) < 3 Then
                            If NpcList(TempCharIndex).Movement = Caminata Then
                                NpcList(TempCharIndex).Contadores.IntervaloMovimiento = AddMod32(GetTickCountRaw(), 15000) ' 15 segundos
                            End If
                            If NpcList(TempCharIndex).SoundOpen <> 0 Then
                                Call WritePlayWave(UserIndex, NpcList(TempCharIndex).SoundOpen, NpcList(TempCharIndex).pos.x, NpcList(TempCharIndex).pos.y, 1, 1)
                            End If
                            Call WriteChatOverHead(UserIndex, "NPCDESC*" & NpcList(TempCharIndex).Numero, NpcList(TempCharIndex).Char.charindex, vbWhite)
                        End If
                    Else
                        Call WriteChatOverHead(UserIndex, "NPCDESC*" & NpcList(TempCharIndex).Numero, NpcList(TempCharIndex).Char.charindex, vbWhite)
                    End If
                End If
            ElseIf IsValidUserRef(NpcList(TempCharIndex).MaestroUser) Then
                If UserList(UserIndex).flags.Muerto = 0 Then
                    estatus = PrepareStatusMsgsForNpcs(TempCharIndex, UserIndex, NpcStatusMask)
                    Call WriteLocaleMsg(UserIndex, 1621, e_FontTypeNames.FONTTYPE_INFO, NpcList(TempCharIndex).Numero & "¬" & NpcList(TempCharIndex).flags.ElementalTags & "¬" & _
                            estatus & "¬" & UserList(NpcList(TempCharIndex).MaestroUser.ArrayIndex).name)  'Msg1621=NPC ¬1 ¬2 es mascota de ¬3
                End If
            Else
                If UserList(UserIndex).flags.Muerto = 0 Then
                    estatus = PrepareStatusMsgsForNpcs(TempCharIndex, UserIndex, NpcStatusMask)
                    Call WriteLocaleMsg(UserIndex, 1622, e_FontTypeNames.FONTTYPE_INFO, NpcList(TempCharIndex).Numero & "¬" & NpcList(TempCharIndex).flags.ElementalTags & "¬" & _
                            estatus)  'Msg1622=NPC ¬1 ¬2
                End If
            End If
            ' End If
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = NpcList(TempCharIndex).npcType
            Call SetNpcRef(UserList(UserIndex).flags.TargetNPC, TempCharIndex)
            Call SetUserRef(UserList(UserIndex).flags.TargetUser, 0)
            UserList(UserIndex).flags.TargetObj = 0
            Dim i As Long, j As Long
            For i = 1 To MAXUSERQUESTS
                With UserList(UserIndex).QuestStats.Quests(i)
                    If .QuestIndex Then
                        If QuestList(.QuestIndex).RequiredTargetNPCs Then
                            For j = 1 To QuestList(.QuestIndex).RequiredTargetNPCs
                                If QuestList(.QuestIndex).RequiredTargetNPC(j).NpcIndex = NpcList(TempCharIndex).Numero Then
                                    If QuestList(.QuestIndex).RequiredTargetNPC(j).amount > .NPCsTarget(j) Then
                                        .NPCsTarget(j) = .NPCsTarget(j) + 1
                                    End If
                                    If QuestList(.QuestIndex).RequiredTargetNPC(j).amount = .NPCsTarget(j) Then
                                        Call FinishQuest(UserIndex, .QuestIndex, i)
                                        Call WriteUpdateNPCSimbolo(UserIndex, TempCharIndex, 1)
                                        Call WriteLocaleChatOverHead(UserIndex, 1353, "", NpcList(TempCharIndex).Char.charindex, vbYellow) ' Msg1353=¡Quest Finalizada!
                                        ' Msg494=Quest Finalizada!
                                        Call WriteLocaleMsg(UserIndex, 494, e_FontTypeNames.FONTTYPE_INFOIAO)
                                    End If
                                End If
                            Next j
                        End If
                    End If
                End With
            Next i
        End If
        With UserList(UserIndex)
            If FoundChar = 0 Then
                Call ClearNpcRef(.flags.TargetNPC)
                .flags.TargetNpcTipo = e_NPCType.Comun
                Call SetUserRef(.flags.TargetUser, 0)
            End If
            '*** NO ENCONTRO NADA ***
            If FoundSomething = 0 Then
                Call ClearNpcRef(.flags.TargetNPC)
                .flags.TargetNpcTipo = e_NPCType.Comun
                Call SetUserRef(.flags.TargetUser, 0)
                .flags.TargetObj = 0
                .flags.TargetObjMap = 0
                .flags.TargetObjX = 0
                .flags.TargetObjY = 0
                If .ShowNothingInterestingMessage > 0 Then
                    'Msg1114= No ves nada interesante.
                    Call WriteLocaleMsg(UserIndex, 1114, e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End With
    Else
        With UserList(UserIndex)
            If FoundSomething = 0 Then
                Call ClearNpcRef(.flags.TargetNPC)
                .flags.TargetNpcTipo = e_NPCType.Comun
                Call SetUserRef(UserList(UserIndex).flags.TargetUser, 0)
                .flags.TargetObj = 0
                .flags.TargetObjMap = 0
                .flags.TargetObjX = 0
                .flags.TargetObjY = 0
                If .ShowNothingInterestingMessage > 0 Then
                    'Msg1106= No ves nada interesante.
                    Call WriteLocaleMsg(UserIndex, 1106, e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End With
    End If
    Exit Sub
LookatTile_Err:
    Call TraceError(Err.Number, Err.Description & " Pos: " & Map & "-" & x & "-" & y & " Usuario: " & UserList(UserIndex).name & " Extra.LookatTile", Erl)
    If FoundChar = 2 Then
        Call TraceError(Err.Number, Err.Description & " Pos: " & Map & "-" & x & "-" & y & "Npc: " & NpcList(TempCharIndex).Numero & " Usuario: " & UserList(UserIndex).name & _
                "Extra.LookatTile X LADDER", Erl)
    End If
End Sub

Function FindDirection(pos As t_WorldPos, Target As t_WorldPos) As e_Heading
    On Error GoTo FindDirection_Err
    '*****************************************************************
    'Devuelve la direccion en la cual el target se encuentra
    'desde pos, 0 si la direc es igual
    '*****************************************************************
    Dim x As Integer
    Dim y As Integer
    x = pos.x - Target.x
    y = pos.y - Target.y
    'NE
    If Sgn(x) = -1 And Sgn(y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), e_Heading.NORTH, e_Heading.EAST)
        Exit Function
    End If
    'NW
    If Sgn(x) = 1 And Sgn(y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), e_Heading.WEST, e_Heading.NORTH)
        Exit Function
    End If
    'SW
    If Sgn(x) = 1 And Sgn(y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), e_Heading.WEST, e_Heading.SOUTH)
        Exit Function
    End If
    'SE
    If Sgn(x) = -1 And Sgn(y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), e_Heading.SOUTH, e_Heading.EAST)
        Exit Function
    End If
    'Sur
    If Sgn(x) = 0 And Sgn(y) = -1 Then
        FindDirection = e_Heading.SOUTH
        Exit Function
    End If
    'norte
    If Sgn(x) = 0 And Sgn(y) = 1 Then
        FindDirection = e_Heading.NORTH
        Exit Function
    End If
    'oeste
    If Sgn(x) = 1 And Sgn(y) = 0 Then
        FindDirection = e_Heading.WEST
        Exit Function
    End If
    'este
    If Sgn(x) = -1 And Sgn(y) = 0 Then
        FindDirection = e_Heading.EAST
        Exit Function
    End If
    'misma
    If Sgn(x) = 0 And Sgn(y) = 0 Then
        FindDirection = 0
        Exit Function
    End If
    Exit Function
FindDirection_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.FindDirection", Erl)
End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean
    On Error GoTo ItemNoEsDeMapa_Err
    ItemNoEsDeMapa = ObjData(Index).OBJType <> e_OBJType.otDoors And ObjData(Index).OBJType <> e_OBJType.otSignBoards And ObjData(Index).OBJType <> e_OBJType.otTrees And ObjData( _
            Index).OBJType <> e_OBJType.otOreDeposit And ObjData(Index).OBJType <> e_OBJType.otTeleport And ObjData(Index).OBJType <> e_OBJType.otMail And ObjData(Index).OBJType _
            <> e_OBJType.otDecorations
    Exit Function
ItemNoEsDeMapa_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.ItemNoEsDeMapa", Erl)
End Function

'[/Barrin 30-11-03]
Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
    On Error GoTo MostrarCantidad_Err
    MostrarCantidad = ObjData(Index).OBJType <> e_OBJType.otDoors And ObjData(Index).OBJType <> e_OBJType.otSignBoards And ObjData(Index).OBJType <> e_OBJType.otOreDeposit And _
            ObjData(Index).OBJType <> e_OBJType.otTrees And ObjData(Index).OBJType <> e_OBJType.otMail And ObjData(Index).OBJType <> e_OBJType.otTeleport
    Exit Function
MostrarCantidad_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.MostrarCantidad", Erl)
End Function

Public Function EsObjetoFijo(ByVal OBJType As e_OBJType) As Boolean
    On Error GoTo EsObjetoFijo_Err
    EsObjetoFijo = OBJType = OBJType = e_OBJType.otSignBoards Or OBJType = e_OBJType.otTrees Or OBJType = e_OBJType.otOreDeposit Or OBJType = e_OBJType.otDecorations
    Exit Function
EsObjetoFijo_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.EsObjetoFijo", Erl)
End Function

Public Function HayPuerta(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If MapData(Map, x, y).ObjInfo.ObjIndex > 0 Then
        HayPuerta = (ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).OBJType = e_OBJType.otDoors) And ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).Cerrada = 1 And (ObjData(MapData( _
                Map, x, y).ObjInfo.ObjIndex).Llave = 0)
    End If
End Function

Public Sub CargarMapasEspeciales()
    Dim File As clsIniManager
    Set File = New clsIniManager
    Call File.Initialize(DatPath & "MapasEspeciales.dat")
    Dim Cantidad As Integer
    Cantidad = val(File.GetValue("MapasInterdimensionales", "Cantidad"))
    Dim i As Integer
    If Cantidad > 0 Then
        ReDim MapasInterdimensionales(1 To Cantidad)
        For i = 1 To Cantidad
            MapasInterdimensionales(i) = val(File.GetValue("MapasInterdimensionales", "Mapa" & i))
        Next
    Else
        ReDim MapasInterdimensionales(0)
    End If
    Cantidad = val(File.GetValue("MapasEventos", "Cantidad"))
    If Cantidad > 0 Then
        ReDim MapasEventos(1 To Cantidad)
        For i = 1 To Cantidad
            MapasEventos(i) = val(File.GetValue("MapasEventos", "Mapa" & i))
        Next
    Else
        ReDim MapasEventos(0)
    End If
    Cantidad = val(File.GetValue("MapasNoDrop", "Cantidad"))
    If Cantidad > 0 Then
        ReDim MapasNoDrop(1 To Cantidad)
        For i = 1 To Cantidad
            MapasNoDrop(i) = val(File.GetValue("MapasNoDrop", "Mapa" & i))
        Next
    Else
        ReDim MapasNoDrop(0)
    End If
    Set File = Nothing
End Sub

Public Function EsMapaEvento(ByVal destMap As Long) As Boolean
    Dim i As Long
    For i = 1 To UBound(MapasEventos)
        If MapasEventos(i) = destMap Then
            EsMapaEvento = True
            Exit Function
        End If
    Next i
    EsMapaEvento = False
End Function

Public Function EsMapaNoDrop(ByVal destMap As Long) As Boolean
    Dim i As Long
    For i = 1 To UBound(MapasNoDrop)
        If MapasNoDrop(i) = destMap Then
            EsMapaNoDrop = True
            Exit Function
        End If
    Next i
    EsMapaNoDrop = False
End Function

Public Sub resetPj(ByVal UserIndex As Integer, Optional ByVal borrarHechizos As Boolean = False)
    With UserList(UserIndex)
        .flags.Muerto = False
        .flags.Escondido = 0
        .flags.Casado = 0
        .flags.SpouseId = 0
        '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
        .Stats.SkillPts = 10
        Call WriteLevelUp(UserIndex, 10)
        Call DarCuerpo(UserIndex) 'Ladder REVISAR
        .OrigChar = .Char
        Dim i As Long
        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = 100
        Next i
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
        .Char.CartAnim = NoCart
        '  Vida inicial
        .Stats.MaxHp = .Stats.UserAtributos(e_Atributos.Constitucion)
        .Stats.MinHp = .Stats.MaxHp
        '  Maná inicial
        .Stats.MaxMAN = .Stats.UserAtributos(e_Atributos.Inteligencia) * ModClase(.clase).ManaInicial
        .Stats.MinMAN = .Stats.MaxMAN
        Dim MiInt As Integer
        MiInt = RandomNumber(1, .Stats.UserAtributos(e_Atributos.Agilidad) \ 6)
        If MiInt = 1 Then MiInt = 2
        .Stats.MaxSta = 20 * MiInt
        .Stats.MinSta = 20 * MiInt
        .Stats.MaxAGU = 100
        .Stats.MinAGU = 100
        .Stats.MaxHam = 100
        .Stats.MinHam = 100
        .flags.VecesQueMoriste = 0
        .flags.Montado = 0
        .Stats.MaxHit = 2
        .Stats.MinHIT = 1
        .Stats.Exp = 0
        .Stats.ELV = 1
        .Stats.GLD = 0
        .Stats.Banco = 0
        If .flags.TomoPocion Then
            For i = 1 To 4
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i
            Call WriteFYA(UserIndex)
        End If
        .flags.DuracionEfecto = 0
        Call VaciarInventario(UserIndex)
        Call ResetCd(UserList(UserIndex))
        Call RellenarInventario(UserIndex)
        'Agrego la poción
        Dim slot_libre As Byte
        For i = 1 To MAX_INVENTORY_SLOTS
            If .invent.Object(i).amount = 0 Then
                slot_libre = i
                Exit For
            End If
        Next i
        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            .BancoInvent.Object(i).amount = 0
            .BancoInvent.Object(i).Equipped = 0
            .BancoInvent.Object(i).ObjIndex = 0
            .BancoInvent.Object(i).ElementalTags = 0
        Next i
        .invent.Object(slot_libre).ObjIndex = POCION_RESET
        .invent.Object(slot_libre).amount = 1
        'Valores Default de facciones al Activar nuevo usuario
        Call ResetFacciones(UserIndex)
        .Faccion.Status = 1
        If borrarHechizos Then
            Call ResetUserSpells(UserIndex)
            '    Call UpdateUserHechizos(True, UserIndex, 0)
        End If
        Call UpdateUserHechizos(True, UserIndex, 0)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call UpdateUserInv(True, UserIndex, 0)
        Call Execute("update user set is_reset = 1 where id = ?;", .Id)
        Call Execute("update quest set quest_id = 0, npcs = 0, npcstarget = 0 where user_id = ?;", .Id)
        Call Execute("delete from quest_done where user_id = ?;", .Id)
        Call ResetQuestStats(UserIndex)
        Call WarpUserChar(UserIndex, .pos.Map, .pos.x, .pos.y, True)
    End With
    'Call WarpUserChar(UserIndex, 1, 55, 45, True)
End Sub

Public Sub ResucitarOCurar(ByVal UserIndex As Integer)
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call RevivirUsuario(UserIndex)
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
        UserList(UserIndex).flags.Envenenado = False
        Call WriteUpdateHP(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(20, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        UserList(UserIndex).Counters.timeFx = 3
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, 35, 1, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        ' Msg495=¡¡Hás sido resucitado!!
        Call WriteLocaleMsg(UserIndex, 495, e_FontTypeNames.FONTTYPE_INFO)
    ElseIf UserList(UserIndex).Stats.MinHp < UserList(UserIndex).Stats.MaxHp Then
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
        UserList(UserIndex).flags.Envenenado = False
        Call WriteUpdateHP(UserIndex)
        UserList(UserIndex).Counters.timeFx = 3
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, 9, 1, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(18, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        ' Msg496=¡¡Hás sido curado!!
        Call WriteLocaleMsg(UserIndex, 496, e_FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Function ByteArr2String(ByRef arr() As Byte) As String
    Dim str As String
    Dim i   As Long
    For i = 0 To UBound(arr)
        str = str + Chr$(arr(i))
    Next i
    ByteArr2String = str
End Function

Public Function PacketIdToString(ByVal PacketId As Long) As String
    Select Case PacketId
        Case 1
            PacketIdToString = "CastSpell"
            Exit Function
        Case 2
            PacketIdToString = "WorkLeftClick"
            Exit Function
        Case 3
            PacketIdToString = "LeftClick"
            Exit Function
        Case 4
            PacketIdToString = "UseItem"
            Exit Function
        Case 5
            PacketIdToString = "UseItemU"
            Exit Function
        Case 6
            PacketIdToString = "Walk"
            Exit Function
        Case 7
            PacketIdToString = "Sailing"
            Exit Function
        Case 8
            PacketIdToString = "Talk"
            Exit Function
        Case 9
            PacketIdToString = "Attack"
            Exit Function
        Case 10
            PacketIdToString = "Drop"
            Exit Function
        Case 11
            PacketIdToString = "Work"
            Exit Function
        Case 12
            PacketIdToString = "EquipItem"
            Exit Function
        Case 13
            PacketIdToString = "GuildMessage"
            Exit Function
        Case 14
            PacketIdToString = "QuestionGM"
            Exit Function
        Case 15
            PacketIdToString = "ChangeHeading"
            Exit Function
        Case 16
            PacketIdToString = "Hide"
            Exit Function
    End Select
End Function

Public Sub TimerQuestOrco()
    Dim UserIndex As Integer
    For UserIndex = 1 To LastUser
        If UserIndex > 0 Then
            With UserList(UserIndex)
                If .flags.UserLogged Then
                    Dim Prob As Long, estimatedProb As Long
                    Prob = RandomNumber(1, LastUser)
                    estimatedProb = LastUser / 3
                    If Prob < estimatedProb And Not EsGM(UserIndex) Then
                        UserList(UserIndex).Stats.MinHam = 0
                        UserList(UserIndex).Stats.MinAGU = 0
                        Call WriteUpdateHungerAndThirst(UserIndex)
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, 20, 50, False))
                    End If
                End If
            End With
        End If
    Next UserIndex
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(156, NO_3D_SOUND, NO_3D_SOUND))
End Sub

Public Function TestRequiredEquipedItem(ByRef inventory As t_Inventario, ByVal RequiredItemsFlag As Long, ByVal RequiredWeaponMask As Integer) As e_SpellRequirementMask
    If IsSet(RequiredItemsFlag, e_SpellRequirementMask.eArmor) And inventory.EquippedArmorObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eArmor
        Exit Function
    End If
    If IsSet(RequiredItemsFlag, e_SpellRequirementMask.eHelm) And inventory.EquippedHelmetObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eHelm
        Exit Function
    End If
    If IsSet(RequiredItemsFlag, e_SpellRequirementMask.eMagicItem) And inventory.EquippedAmuletAccesoryObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eMagicItem
        Exit Function
    End If
    If IsSet(RequiredItemsFlag, e_SpellRequirementMask.eProjectile) And inventory.EquippedMunitionObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eProjectile
        Exit Function
    End If
    If IsSet(RequiredItemsFlag, e_SpellRequirementMask.eShield) And inventory.EquippedShieldObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eShield
        Exit Function
    End If
    If IsSet(RequiredItemsFlag, e_SpellRequirementMask.eShip) And inventory.EquippedShipObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eShip
        Exit Function
    End If
    If IsSet(RequiredItemsFlag, e_SpellRequirementMask.eTool) And inventory.EquippedWorkingToolObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eTool
        Exit Function
    End If
    If IsSet(RequiredItemsFlag, e_SpellRequirementMask.eWeapon) Then
        If inventory.EquippedWeaponObjIndex = 0 Then
            If Not IsIntSet(RequiredWeaponMask, e_WeaponType.eFist) Then
                TestRequiredEquipedItem = e_SpellRequirementMask.eWeapon
                Exit Function
            End If
        ElseIf RequiredWeaponMask > 0 Then
            If Not IsIntSet(RequiredWeaponMask, ShiftLeft(1, ObjData(inventory.EquippedWeaponObjIndex).WeaponType)) Then
                TestRequiredEquipedItem = e_SpellRequirementMask.eWeapon
                Exit Function
            End If
        End If
    End If
    TestRequiredEquipedItem = e_SpellRequirementMask.eNone
End Function

Public Sub SendrequiredItemMessage(ByVal UserIndex As Integer, ByVal itemMask As e_SpellRequirementMask, ByVal Message As String)
    Select Case itemMask
        Case e_SpellRequirementMask.eArmor
            ' Msg497=Necesitás una armadura
            Call WriteLocaleMsg(UserIndex, 497, Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eHelm
            ' Msg564=Necesitás un casco
            Call WriteLocaleMsg(UserIndex, 564, Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eKnucle
            ' Msg565=Necesitás unos nudillos
            Call WriteLocaleMsg(UserIndex, 565, Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eMagicItem
            ' Msg566=Necesitás un objeto magico
            Call WriteLocaleMsg(UserIndex, 566, Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eProjectile
            ' Msg567=Necesitás municiones
            Call WriteLocaleMsg(UserIndex, 567, Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eShield
            ' Msg568=Necesitás un escudo
            Call WriteLocaleMsg(UserIndex, 568, Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eShip
            ' Msg569=Necesitás un barco
            Call WriteLocaleMsg(UserIndex, 569, Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eTool
            ' Msg570=Necesitás una herramienta
            Call WriteLocaleMsg(UserIndex, 570, Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eWeapon
            ' Msg571=Necesitás un arma
            Call WriteLocaleMsg(UserIndex, 571, Message, e_FontTypeNames.FONTTYPE_INFO)
    End Select
End Sub

Public Sub SendHelpInteractionMessage(ByVal UserIndex As Integer, ByVal CanHelpResult As e_InteractionResult)
    Select Case CanHelpResult
        Case eOposingFaction
            Call WriteLocaleMsg(UserIndex, MsgFactionForbidAction, e_FontTypeNames.FONTTYPE_INFO)
        Case eCantHelpCriminal
            Call WriteLocaleMsg(UserIndex, MsgDisableAttackGuardToContinue, e_FontTypeNames.FONTTYPE_INFO)
        Case eCantHelpCriminalClanRules
            Call WriteLocaleMsg(UserIndex, MsgClanForbidAction, e_FontTypeNames.FONTTYPE_INFO)
        Case Else
            Debug.Assert "Invalid message"
    End Select
End Sub

Public Sub SendAttackInteractionMessage(ByVal UserIndex As Integer, ByVal CanAttackResult As e_AttackInteractionResult)
    Select Case CanAttackResult
        Case e_AttackInteractionResult.eCanAttack
            'Do nothing
        Case e_AttackInteractionResult.eDeathAttacker
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eFightActive
            ' Msg498=No podés atacar en este momento.
            Call WriteLocaleMsg(UserIndex, 498, e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eDeathTarget
            ' Msg499=No podés atacar a un espiritu.
            Call WriteLocaleMsg(UserIndex, 499, e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eSameGroup
            ' Msg500=No podés atacar a un miembro de tu grupo.
            Call WriteLocaleMsg(UserIndex, 500, e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eTalkWithMaster
            ' Msg501=No podés atacar usuarios mientras estás en consulta.
            Call WriteLocaleMsg(UserIndex, 501, e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eAttackerIsCursed
            ' Msg502=No podés atacar usuarios mientras estás en consulta.
            Call WriteLocaleMsg(UserIndex, 502, e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eMounted
            ' Msg503=No podés atacar usando una montura.
            Call WriteLocaleMsg(UserIndex, 503, e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eSameTeam
            ' Msg504=No podes atacar un miembro de tu equipo.
            Call WriteLocaleMsg(UserIndex, 504, e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eNotEnougthPrivileges
            ' Msg505=El ser es demasiado poderoso.
            Call WriteLocaleMsg(UserIndex, 505, e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eSameClan
            ' Msg506=No podes atacar a un miembro de tu clan.
            Call WriteLocaleMsg(UserIndex, 506, e_FontTypeNames.FONTTYPE_INFOIAO)
        Case e_AttackInteractionResult.eSameFaction
            Call WriteLocaleMsg(UserIndex, MsgFacctionForbidAttack, e_FontTypeNames.FONTTYPE_INFOIAO)
        Case e_AttackInteractionResult.eRemoveSafe
            Call WriteLocaleMsg(UserIndex, MsgRemoveSafeToAttack, e_FontTypeNames.FONTTYPE_INFOIAO)
        Case e_AttackInteractionResult.eSafeArea
            ' Msg572=Esta es una zona segura, aquí no podés atacar otros usuarios.
            Call WriteLocaleMsg(UserIndex, 572, e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eCantAttackYourself
            Call WriteLocaleMsg(UserIndex, MsgCantAttackYourself, e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eAttackSameFaction
            ' Msg507=¡Atacaste un ciudadano! Te has convertido en un Criminal.
            Call WriteLocaleMsg(UserIndex, 507, e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eAttackPetSameFaction
            ' Msg508=¡Atacaste una mascota de un ciudadano! Te has convertido en un Criminal.
            Call WriteLocaleMsg(UserIndex, 508, e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eRemoveSafeCitizenNpc
            ' Msg509=Debes quitar el seguro para poder atacar la criatura que esta luchando con otro usuario.
            Call WriteLocaleMsg(UserIndex, 509, e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eAttackCitizenNpc
            ' Msg510=¡Atacaste una criatura de otro usuario! Te has convertido en un Criminal.
            Call WriteLocaleMsg(UserIndex, 510, e_FontTypeNames.FONTTYPE_WARNING)
        Case Else
            ' Msg511=Target inválido.
            Call WriteLocaleMsg(UserIndex, 511, e_FontTypeNames.FONTTYPE_INFOIAO)
    End Select
End Sub

Public Function PreferedTileForDirection(ByRef Direction As t_Vector, ByRef currentPosition As t_WorldPos) As t_WorldPos
    Dim Ret    As t_WorldPos
    Dim Normal As t_Vector
    Ret.Map = currentPosition.Map
    Normal = GetNormal(Direction)
    If Abs(Normal.x) > Abs(Normal.y) Then
        Ret.x = currentPosition.x + 1 * Sgn(Normal.x)
        Ret.y = currentPosition.y
    ElseIf Abs(Normal.x) < Abs(Normal.y) Then
        Ret.x = currentPosition.x
        Ret.y = currentPosition.y + 1 * Sgn(Normal.y)
    Else
        Ret.y = currentPosition.y + 1 * Sgn(Normal.y)
        Ret.x = currentPosition.x + 1 * Sgn(Normal.x)
    End If
    If Not LegalPos(Ret.Map, Ret.x, Ret.y, False, True) Then
        Call ClosestStablePos(Ret, Ret)
    End If
    If Ret.x = 0 And Ret.y = 0 Then
        PreferedTileForDirection = currentPosition
    Else
        PreferedTileForDirection = Ret
    End If
End Function

Public Function PrepareUserStatusEffectMsgsForPlayers(ByVal targetUserIndex As Integer, _
                                                      ByVal SourceUserIndex As Integer, _
                                                      ByRef Statuses As Long, _
                                                      ByRef FactionStatuses As Long, _
                                                      ByRef fontType As e_FontTypeNames) As String
    On Error GoTo PrepareUserStatusEffectMsgsForPlayers_Err
    Dim extraStrings As String
    With UserList(targetUserIndex)
        extraStrings = extraStrings & .name & "-"
        If Len(.Desc) > 0 Then
            extraStrings = extraStrings & .Desc & "-"
        Else
            extraStrings = extraStrings & "-"
        End If
        If .GuildIndex > 0 Then
            extraStrings = extraStrings & modGuilds.GuildName(.GuildIndex) & "-"
        Else
            extraStrings = extraStrings & "-"
        End If
        If .flags.Casado = 1 Then
            extraStrings = extraStrings & GetUserSpouse(targetUserIndex) & "-"
        Else
            extraStrings = extraStrings & "-"
        End If
        'if im am a gm and im clicking other person i have extra data
        If EsGM(SourceUserIndex) Then
            extraStrings = extraStrings & .clase & "-" & .raza & "-" & .Stats.ELV & "-" & .Stats.ELO & "-"
        Else
            extraStrings = extraStrings & "----"
        End If
        If EsNewbie(targetUserIndex) Then
            Call SetMask(Statuses, e_UsersInfoMask.Newbie)
        End If
        Dim factionRank As Byte
        If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
            factionRank = TituloReal(targetUserIndex)
            Select Case factionRank
                Case e_RoyalArmyRanks.FirstHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ArmyFirstHierarchy)
                Case e_RoyalArmyRanks.SecondHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ArmySecondHierarchy)
                Case e_RoyalArmyRanks.ThirdHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ArmyThirdHierarchy)
                Case e_RoyalArmyRanks.FourthHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ArmyFourthHierarchy)
                Case e_RoyalArmyRanks.FifthHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ArmyFifthHierarchy)
                Case e_RoyalArmyRanks.SixthHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ArmySixthHierarchy)
                Case e_RoyalArmyRanks.SeventhHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ArmySeventhHierarchy)
                Case Else
            End Select
        End If
        If .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
            factionRank = TituloCaos(targetUserIndex)
            Select Case factionRank
                Case e_ChaosArmyRanks.FirstHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ChaosFirstHierarchy)
                Case e_ChaosArmyRanks.SecondHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ChaosSecondHierarchy)
                Case e_ChaosArmyRanks.ThirdHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ChaosThirdHierarchy)
                Case e_ChaosArmyRanks.FourthHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ChaosFourthHierarchy)
                Case e_ChaosArmyRanks.FifthHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ChaosFifthHierarchy)
                Case e_ChaosArmyRanks.SixthHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ChaosSixthHierarchy)
                Case e_ChaosArmyRanks.SeventhHierarchy
                    Call SetMask(FactionStatuses, e_UsersInfoMask2.ChaosSeventhHierarchy)
                Case Else
            End Select
        End If
        If .Faccion.Status = e_Facciones.Criminal Then
            Call SetMask(FactionStatuses, e_UsersInfoMask2.Criminal)
            fontType = e_FontTypeNames.FONTTYPE_CRIMINAL
        End If
        If .Faccion.Status = e_Facciones.Caos Then
            Call SetMask(FactionStatuses, e_UsersInfoMask2.Chaotic)
            fontType = e_FontTypeNames.FONTTYPE_CRIMINAL_CAOS
        End If
        If .Faccion.Status = e_Facciones.concilio Then
            Call SetMask(FactionStatuses, e_UsersInfoMask2.ChaoticCouncil)
            fontType = e_FontTypeNames.FONTTYPE_CONSEJOCAOS
        End If
        If .Faccion.Status = e_Facciones.Ciudadano Then
            Call SetMask(FactionStatuses, e_UsersInfoMask2.Citizen)
            fontType = e_FontTypeNames.FONTTYPE_CITIZEN
        End If
        If .Faccion.Status = e_Facciones.Armada Then
            Call SetMask(FactionStatuses, e_UsersInfoMask2.Army)
            fontType = e_FontTypeNames.FONTTYPE_CITIZEN_ARMADA
        End If
        If .Faccion.Status = e_Facciones.consejo Then
            Call SetMask(FactionStatuses, e_UsersInfoMask2.RoyalCouncil)
            fontType = e_FontTypeNames.FONTTYPE_CONSEJO
        End If
        If EsGM(targetUserIndex) Then
            Select Case .flags.Privilegios
                Case e_PlayerType.Consejero
                    Call SetMask(Statuses, e_UsersInfoMask.Counselor)
                    fontType = e_FontTypeNames.FONTTYPE_GM
                Case e_PlayerType.SemiDios
                    Call SetMask(Statuses, e_UsersInfoMask.DemiGod)
                    fontType = e_FontTypeNames.FONTTYPE_GM
                Case e_PlayerType.Dios
                    Call SetMask(Statuses, e_UsersInfoMask.God)
                    fontType = e_FontTypeNames.FONTTYPE_DIOS
                Case e_PlayerType.Admin
                    Call SetMask(Statuses, e_UsersInfoMask.Admin)
                    fontType = e_FontTypeNames.FONTTYPE_DIOS
                Case e_PlayerType.RoleMaster
                    Call SetMask(Statuses, e_UsersInfoMask.RoleMaster)
                    fontType = e_FontTypeNames.FONTTYPE_GM
            End Select
        End If
        'if im clicking and i have survival skill 50 or more i see all status
        If UserList(SourceUserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 50 Then
            If .flags.Envenenado > 0 Then
                Call SetMask(Statuses, e_UsersInfoMask.Poisoned)
            End If
            If .flags.Ceguera = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.Blind)
            End If
            If .flags.Incinerado = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.Incinerated)
            End If
            If .flags.Paralizado = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.Paralized)
            End If
            If .flags.Inmovilizado = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.Inmovilized)
            End If
            If .Counters.Trabajando > 0 Then
                Call SetMask(Statuses, e_UsersInfoMask.Working)
            End If
            If .flags.invisible = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.invisible)
            End If
            If .flags.Oculto = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.Hidden)
            End If
            If .flags.Estupidez = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.Stupid)
            End If
            If .flags.Maldicion = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.Cursed)
            End If
            If .flags.Silenciado = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.Silenced)
            End If
            If .flags.Comerciando = True Then
                Call SetMask(Statuses, e_UsersInfoMask.Trading)
            End If
            If .flags.Descansar = 1 Then
                Call SetMask(Statuses, e_UsersInfoMask.Resting)
            End If
            If .flags.Meditando Then
                Call SetMask(Statuses, e_UsersInfoMask.Focusing)
            End If
            Select Case .Stats.MinHp
                Case 0
                    Call SetMask(Statuses, e_UsersInfoMask.Dead)
                    fontType = e_FontTypeNames.FONTTYPE_New_Gris
                Case Is < (.Stats.MaxHp * 0.1)
                    Call SetMask(Statuses, e_UsersInfoMask.AlmostDead)
                Case Is < (.Stats.MaxHp * 0.5)
                    Call SetMask(Statuses, e_UsersInfoMask.SeriouslyWounded)
                Case Is < (.Stats.MaxHp * 0.75)
                    Call SetMask(Statuses, e_UsersInfoMask.Wounded)
                Case Is < (.Stats.MaxHp * 0.99)
                    Call SetMask(Statuses, e_UsersInfoMask.LightlyWounded)
                Case Else
                    Call SetMask(Statuses, e_UsersInfoMask.Intact)
            End Select
        Else
            'even if i dont have survival, I have to see the dead people
            If .Stats.MinHp = 0 Then
                Call SetMask(Statuses, e_UsersInfoMask.Dead)
                fontType = e_FontTypeNames.FONTTYPE_New_Gris
            End If
        End If
        PrepareUserStatusEffectMsgsForPlayers = extraStrings
    End With
    Exit Function
PrepareUserStatusEffectMsgsForPlayers_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.PrepareUserStatusEffectMsgsForPlayers", Erl)
End Function

Public Function PrepareStatusMsgsForNpcs(ByVal TargetNpcIndex As Integer, ByVal SourceUserIndex As Integer, ByRef NpcStatusMask As Long) As String
    On Error GoTo PrepareStatusMsgsForNpcs_Err
    Dim extraStrings      As String
    Dim UserSurvivalSkill As Integer
    UserSurvivalSkill = UserList(SourceUserIndex).Stats.UserSkills(e_Skill.Supervivencia)
    With NpcList(TargetNpcIndex)
        Select Case UserSurvivalSkill
            Case Is >= 75
                extraStrings = extraStrings & .Stats.MinHp & "/" & .Stats.MaxHp & "-"
            Case Is >= 50
                extraStrings = extraStrings & Round((.Stats.MinHp / .Stats.MaxHp) * 100#, 0) & "%" & "-"
            Case Is >= 25
                Select Case .Stats.MinHp
                    Case Is < (.Stats.MaxHp * 0.1)
                        Call SetMask(NpcStatusMask, e_NpcInfoMask.AlmostDead)
                    Case Is < (.Stats.MaxHp * 0.2)
                        Call SetMask(NpcStatusMask, e_NpcInfoMask.SeriouslyWounded)
                    Case Is < (.Stats.MaxHp * 0.5)
                        Call SetMask(NpcStatusMask, e_NpcInfoMask.Wounded)
                    Case Is < (.Stats.MaxHp * 0.7)
                        Call SetMask(NpcStatusMask, e_NpcInfoMask.LightlyWounded)
                    Case Else
                        Call SetMask(NpcStatusMask, e_NpcInfoMask.Intact)
                End Select
                extraStrings = extraStrings & "-"
            Case Else
                If .Stats.MinHp < .Stats.MaxHp Then
                    Call SetMask(NpcStatusMask, e_NpcInfoMask.Wounded)
                Else
                    Call SetMask(NpcStatusMask, e_NpcInfoMask.Intact)
                End If
                extraStrings = extraStrings & "-"
        End Select
        If .flags.Envenenado > 0 Then
            Call SetMask(NpcStatusMask, e_NpcInfoMask.Poisoned)
        End If
        If .flags.Paralizado = 1 Then
            If UserSurvivalSkill >= 100 Then
                extraStrings = extraStrings & CInt(.Contadores.Paralisis / 6.5)
            End If
            Call SetMask(NpcStatusMask, e_NpcInfoMask.Paralized)
            extraStrings = extraStrings & "-"
        Else
            extraStrings = extraStrings & "-"
        End If
        If .flags.Inmovilizado = 1 Then
            If UserSurvivalSkill >= 100 Then
                extraStrings = extraStrings & CInt(.Contadores.Inmovilizado / 6.5)
            End If
            Call SetMask(NpcStatusMask, e_NpcInfoMask.Inmovilized)
            extraStrings = extraStrings & "-"
        Else
            extraStrings = extraStrings & "-"
        End If
        If GetOwnedBy(TargetNpcIndex) <> 0 Then
            Call SetMask(NpcStatusMask, e_NpcInfoMask.Fighting)
            extraStrings = extraStrings & .flags.AttackedBy & "|"
            extraStrings = extraStrings & CInt((IntervaloNpcOwner - (GlobalFrameTime - .flags.AttackedTime)) / 1000) & "-"
        Else
            extraStrings = extraStrings & "-"
        End If
        If EsGM(SourceUserIndex) Then
            extraStrings = extraStrings & TargetNpcIndex & "-"
        Else
            extraStrings = extraStrings & "-"
        End If
        PrepareStatusMsgsForNpcs = extraStrings & NpcStatusMask
    End With
    Exit Function
PrepareStatusMsgsForNpcs_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.PrepareStatusMsgsForNpcs", Erl)
End Function
