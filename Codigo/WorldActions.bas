Attribute VB_Name = "WorldActions"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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

Public Function GetMapName(ByVal Map As Long) As String
    On Error GoTo get_map_name_Err
    GetMapName = MapInfo(Map).map_name
    Exit Function
get_map_name_Err:
    Call TraceError(Err.Number, Err.Description, "WorldActions.GetMapName", Erl)
End Function

Public Function CanUseObject(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByVal writeInConsole As Boolean = False) As Byte
    On Error GoTo PuedeUsarObjeto_Err
    Dim Objeto As t_ObjData
    Dim msg    As String
    Dim Extra  As String
    Extra = vbNullString
    Objeto = ObjData(ObjIndex)
    With UserList(UserIndex)
        If EsGM(UserIndex) Then
            CanUseObject = 0
            msg = vbNullString
            Exit Function
        End If
        ' Keep the original priority: first match wins
        If Objeto.Newbie = 1 And Not EsNewbie(UserIndex) Then
            CanUseObject = 7
            msg = "679" ' Only newbies can use this item.
        ElseIf .Stats.ELV < Objeto.MinELV Then
            CanUseObject = 6
            msg = "1926" ' Need level {0}
            Extra = CStr(Objeto.MinELV)
        ElseIf .Stats.ELV > Objeto.MaxLEV And Objeto.MaxLEV > 0 Then
            CanUseObject = 6
            msg = "1982" ' Not for level {0} or higher
            Extra = CStr(Objeto.MaxLEV)
        ElseIf Not FaccionPuedeUsarItem(UserIndex, ObjIndex) And JerarquiaPuedeUsarItem(UserIndex, ObjIndex) Then
            CanUseObject = 3
            msg = "416"  ' Faction doesn't allow it.
        ElseIf Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
            CanUseObject = 2
            msg = "265"  ' Class cannot use this item.
        ElseIf Not SexoPuedeUsarItem(UserIndex, ObjIndex) Then
            CanUseObject = 1
            msg = "267"  ' Sex cannot use this item.
        ElseIf Not RazaPuedeUsarItem(UserIndex, ObjIndex) Then
            CanUseObject = 5
            msg = "266"  ' Race cannot use this item.
        ElseIf (Objeto.SkillIndex > 0) Then
            If (.Stats.UserSkills(Objeto.SkillIndex) < Objeto.SkillRequerido) Then
                CanUseObject = 4
                msg = "NEED_SKILL_POINTS" ' e.g. "Necesitas {0} puntos en {1}..."
                Extra = CStr(Objeto.SkillRequerido) & "¬" & SkillsNames(Objeto.SkillIndex)
            Else
                CanUseObject = 0
                msg = vbNullString
            End If
        Else
            CanUseObject = 0
            msg = vbNullString
        End If
        ' Only emit when we actually have a message
        If msg <> vbNullString Then
            If writeInConsole Then
                Call WriteLocaleMsg(UserIndex, msg, e_FontTypeNames.FONTTYPE_INFO, Extra)
            End If
        End If
    End With
    Exit Function
PuedeUsarObjeto_Err:
    Call TraceError(Err.Number, Err.Description, "WorldActions.CanUseObject", Erl)
End Function

Public Sub CompletePendingAction(ByVal UserIndex As Integer)
    On Error GoTo EndProgrammedAction_Err
    Dim obj  As t_ObjData
    Dim Slot As Byte
    With UserList(UserIndex)
        Select Case .Accion.TipoAccion
            Case e_AccionBarra.Runa
                obj = ObjData(.Accion.RunaObj)
                Slot = .Accion.ObjSlot
                Select Case obj.TipoRuna
                    Case e_RuneType.ReturnHome
                        Call HomeArrival(UserIndex)
                    Case e_RuneType.MesonSafePassage
                        If .pos.Map = MAP_MESON_HOSTIGADO Or .pos.Map = MAP_MESON_HOSTIGADO_TRADING_ZONE Then
                            Call WriteLocaleMsg(UserIndex, MSG_NOT_USABLE_INSIDE_MESON, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If obj.HastaMap <> MAP_MESON_HOSTIGADO Then
                            Call WriteLocaleMsg(UserIndex, MSG_INVALID_RUNE, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If Not IsValidMapPosition(obj.HastaMap, obj.HastaX, obj.HastaY) Then
                            Call WriteLocaleMsg(UserIndex, MSG_INVALID_RUNE, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        .flags.ReturnPos = .pos
                        Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                        Call WriteLocaleMsg(UserIndex, MSG_SUCCESFULLY_TELEPORTED, e_FontTypeNames.FONTTYPE_WARNING)
                    Case e_RuneType.FastTravel
                        If .pos.Map <> obj.DesdeMap Then
                            Call WriteLocaleMsg(UserIndex, MSG_INVALID_FAST_TRAVEL_MAP_ORIGIN, e_FontTypeNames.FONTTYPE_WARNING)
                        End If
                        If Not IsValidMapPosition(obj.HastaMap, obj.HastaX, obj.HastaY) Then
                            Call WriteLocaleMsg(UserIndex, MSG_INVALID_RUNE, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                        Call WriteLocaleMsg(UserIndex, MSG_SUCCESFULLY_TELEPORTED, e_FontTypeNames.FONTTYPE_WARNING)
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        Call UpdateUserInv(False, UserIndex, Slot)
                End Select
            Case e_AccionBarra.Hogar
                Call HomeArrival(UserIndex)
        End Select
        .Accion.Particula = 0
        .Accion.TipoAccion = e_AccionBarra.CancelarAccion
        .Accion.HechizoPendiente = 0
        .Accion.RunaObj = 0
        .Accion.ObjSlot = 0
        .Accion.AccionPendiente = False
    End With
    Exit Sub
EndProgrammedAction_Err:
    Call TraceError(Err.Number, Err.Description, "WorldActions.CompletePendingAction", Erl)
End Sub

Public Sub HandleWorldAction(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo HandleWorldAction_Err

    If Not IsWorldActionRequestValid(UserIndex, Map, x, y) Then Exit Sub

    If TryHandleNpcWorldAction(UserIndex, Map, x, y) Then Exit Sub
    If TryHandleObjectWorldAction(UserIndex, Map, x, y) Then Exit Sub
    If TryHandleMultiTileObjectWorldAction(UserIndex, Map, x, y) Then Exit Sub

    Exit Sub

HandleWorldAction_Err:
    Call TraceError(Err.Number, Err.Description, "WorldActions.HandleWorldAction", Erl)
End Sub

Private Function IsWorldActionRequestValid(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If UserIndex <= 0 Then Exit Function

    'Range validation (ToxicWaste)
    If (Abs(UserList(UserIndex).pos.y - y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).pos.x - x) > RANGO_VISION_X) Then
        Exit Function
    End If

    'Valid position?
    If Not InMapBounds(Map, x, y) Then Exit Function

    IsWorldActionRequestValid = True
End Function

Private Function TryHandleNpcWorldAction(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    Dim TempCharIndex As Integer

    If MapData(Map, x, y).NpcIndex <= 0 Then Exit Function

    TempCharIndex = MapData(Map, x, y).NpcIndex
    'Set the target NPC
    Call SetNpcRef(UserList(UserIndex).flags.TargetNPC, TempCharIndex)
    UserList(UserIndex).flags.TargetNpcTipo = NpcList(TempCharIndex).npcType

    Call HandleNpcInteractionByType(UserIndex, TempCharIndex)

    TryHandleNpcWorldAction = True
End Function

Private Function EnsureUserAliveForNpcInteraction(ByVal UserIndex As Integer, Optional ByVal fontType As e_FontTypeNames = e_FontTypeNames.FONTTYPE_INFO) As Boolean
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteLocaleMsg(UserIndex, MSG_MUERTO, fontType)
        Exit Function
    End If

    EnsureUserAliveForNpcInteraction = True
End Function

Private Function EnsureUserNotTrading(ByVal UserIndex As Integer) As Boolean
    If UserList(UserIndex).flags.Comerciando Then Exit Function

    EnsureUserNotTrading = True
End Function

Private Function EnsureNpcWithinDistance(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal maxDistance As Integer, Optional ByVal tooFarMsgId As Long = MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, Optional ByVal fontType As e_FontTypeNames = e_FontTypeNames.FONTTYPE_INFO) As Boolean
    If Distancia(NpcList(NpcIndex).pos, UserList(UserIndex).pos) > maxDistance Then
        Call WriteLocaleMsg(UserIndex, tooFarMsgId, fontType)
        Exit Function
    End If

    EnsureNpcWithinDistance = True
End Function

Private Sub PauseWalkingNpcForInteraction(ByVal NpcIndex As Integer, ByVal intervalMs As Long)
    If NpcList(NpcIndex).Movement = e_TipoAI.Caminata Then
        NpcList(NpcIndex).Contadores.IntervaloMovimiento = AddMod32(GetTickCountRaw(), intervalMs)
    End If
End Sub


Private Sub HandleReviverNpcInteraction(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    If Distancia(UserList(UserIndex).pos, NpcList(NpcIndex).pos) > 5 Then
        'Msg8=El sacerdote no puede curarte debido a que estas demasiado lejos.
        Call WriteLocaleMsg(UserIndex, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    '  No frenamos la caminata al curar/revivir para evitar que puedan dejar al NPC quieto a base de clicks.
    If UserList(UserIndex).pos.Map <> MAP_HOME_IN_JAIL Then
        UserList(UserIndex).flags.Envenenado = 0
        UserList(UserIndex).flags.Incinerado = 0
    End If
    'Revivimos si es necesario
    If UserList(UserIndex).flags.Muerto = 1 And (NpcList(NpcIndex).npcType = e_NPCType.Revividor Or EsNewbie(UserIndex)) Then
        ' Msg585=¿Has sido resucitado!
        Call WriteLocaleMsg(UserIndex, MSG_SIDO_RESUCITADO_585, e_FontTypeNames.FONTTYPE_INFO)
        Call RevivirUsuario(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.Resucitar, 30, False))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
    Else
        If UserList(UserIndex).pos.Map = MAP_HOME_IN_JAIL And NpcList(NpcIndex).npcType = e_NPCType.Revividor Then
            Exit Sub
        End If
        'curamos totalmente
        If UserList(UserIndex).Stats.MinHp <> UserList(UserIndex).Stats.MaxHp Then
            UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
            Call WritePlayWave(UserIndex, 117, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y)
            'Msg83=El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...¿Has sido curado!
            Call WriteLocaleMsg(UserIndex, MSG_CURA_LANZA_PALABRAS_AIRE_COMIENZAS_SENTIR_COMO_CUERPO, e_FontTypeNames.FONTTYPE_INFOIAO)
            Call WriteUpdateUserStats(UserIndex)
            If Status(UserIndex) = 4 Or Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.CurarCrimi, 100, False))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.CurarCrimi, 100, False))
            End If
        End If
    End If
End Sub

Private Sub HandleFactionRecruiterNpcInteraction(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    If UserList(UserIndex).flags.Muerto = 1 Then
        'Msg77=¡¡Estás muerto!!.
        Call WriteLocaleMsg(UserIndex, MSG_MUERTO, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If Distancia(NpcList(NpcIndex).pos, UserList(UserIndex).pos) > 3 Then
        Call WriteLocaleMsg(UserIndex, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If NpcList(NpcIndex).flags.Faccion = 0 Then
        If UserList(UserIndex).Faccion.Status <> e_Facciones.Armada And UserList(UserIndex).Faccion.Status <> e_Facciones.consejo Then
            Call EnlistarArmadaReal(UserIndex)
        Else
            Call RecompensaArmadaReal(UserIndex)
        End If
    Else
        If UserList(UserIndex).Faccion.Status <> e_Facciones.Caos And UserList(UserIndex).Faccion.Status <> e_Facciones.concilio Then
            Call EnlistarCaos(UserIndex)
        Else
            Call RecompensaCaos(UserIndex)
        End If
    End If
End Sub

Private Function GetCityName(ByVal city As e_Ciudad) As String
    Select Case city
        Case e_Ciudad.cUllathorpe
            GetCityName = "Ullathorpe"
        Case e_Ciudad.cNix
            GetCityName = "Nix"
        Case e_Ciudad.cBanderbill
            GetCityName = "Banderbill"
        Case e_Ciudad.cLindos
            GetCityName = "Lindos"
        Case e_Ciudad.cArghal
            GetCityName = "Arghal"
        Case e_Ciudad.cForgat
            GetCityName = "Forgat"
        Case e_Ciudad.cEldoria
            GetCityName = "Eldoria"
        Case e_Ciudad.cArkhein
            GetCityName = "Arkhein"
        Case e_Ciudad.cPenthar
            GetCityName = "Penthar"
        Case Else
            GetCityName = "Ullathorpe"
    End Select
End Function

Private Sub HandleGovernorNpcInteraction(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    If UserList(UserIndex).flags.Muerto = 1 Then
        'Msg77=¡¡Estás muerto!!.
        Call WriteLocaleMsg(UserIndex, MSG_MUERTO, e_FontTypeNames.FONTTYPE_INFOIAO)
        Exit Sub
    End If

    If Distancia(NpcList(NpcIndex).pos, UserList(UserIndex).pos) > 3 Then
        Call WriteLocaleMsg(UserIndex, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO)
        'Msg8=Estas demasiado lejos del gobernador.
        Exit Sub
    End If

    Dim DeDonde    As String
    Dim Gobernador As t_Npc

    Gobernador = NpcList(NpcIndex)

    If UserList(UserIndex).Hogar = Gobernador.GobernadorDe Then
        Call WriteLocaleChatOverHead(UserIndex, 1349, "", Gobernador.Char.charindex, vbWhite) ' Msg1349=Ya perteneces a esta ciudad. Gracias por ser uno más de nosotros.
        Exit Sub
    End If

    If (UserList(UserIndex).Faccion.Status = 0 Or UserList(UserIndex).Faccion.Status = 2) And Gobernador.GobernadorDe = e_Ciudad.cBanderbill Then
        Call WriteLocaleChatOverHead(UserIndex, "1350", "", Gobernador.Char.charindex, vbWhite) ' Msg1350=Aquí no aceptamos criminales.
        Exit Sub
    End If

    If (UserList(UserIndex).Faccion.Status = 3 Or UserList(UserIndex).Faccion.Status = 1) And Gobernador.GobernadorDe = e_Ciudad.cArkhein Then
        Call WriteLocaleChatOverHead(UserIndex, "1351", "", Gobernador.Char.charindex, vbWhite) ' Msg1351=¡¡Sal de aquí ciudadano asqueroso!!
        Exit Sub
    End If

    UserList(UserIndex).PosibleHogar = Gobernador.GobernadorDe
    DeDonde = GetCityName(UserList(UserIndex).PosibleHogar)
    UserList(UserIndex).flags.pregunta = 3
    Call WritePreguntaBox(UserIndex, 1592, DeDonde)
End Sub

Private Sub HandleFishingDeliveryNpcInteraction(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    If UserList(UserIndex).clase <> Trabajador Then
        'Msg2168=Solo los trabajadores pueden registrar los peces especiales.
        Call WriteLocaleMsg(UserIndex, MSG_SOLO_TRABAJADORES_PUEDEN_REGISTRAR_PECES_ESPECIALES, e_FontTypeNames.FONTTYPE_INFOIAO)
        Exit Sub
    End If
    Dim PuntosTotales       As Long
    Dim OroTotal            As Long

    Call CalculateSpecialFishingRewards(UserIndex, PuntosTotales, OroTotal)

    If PuntosTotales > 0 Then
        UserList(UserIndex).flags.pregunta = 5
        Call WritePreguntaBox(UserIndex, 1593, PuntosTotales & "¿" & PonerPuntos(OroTotal * 1.2)) 'Msg1593= Tienes un total de ¿1 puntos y ¿2 monedas de oro para reclamar, ¿Deseas aceptar?
    Else
        Dim charindexstr As Integer
        charindexstr = str$(NpcList(NpcIndex).Char.charindex)
        Call WriteLocaleChatOverHead(UserIndex, "1352", "", charindexstr, &HFFFF00) ' Msg1352=No tienes ningún trofeo de pesca para entregar.
    End If
End Sub

Private Sub CalculateSpecialFishingRewards(ByVal UserIndex As Integer, ByRef totalPoints As Long, ByRef totalGold As Long)
    Dim i                   As Integer, j As Integer
    Dim CantPecesEspeciales As Long

    CantPecesEspeciales = UBound(PecesEspeciales)
    If CantPecesEspeciales > 0 Then
        For i = 1 To MAX_INVENTORY_SLOTS
            For j = 1 To CantPecesEspeciales
                If UserList(UserIndex).invent.Object(i).ObjIndex = PecesEspeciales(j).ObjIndex Then
                    totalPoints = totalPoints + (ObjData(UserList(UserIndex).invent.Object(i).ObjIndex).PuntosPesca * UserList(UserIndex).invent.Object(i).amount)
                    totalGold = totalGold + (ObjData(UserList(UserIndex).invent.Object(i).ObjIndex).Valor * UserList(UserIndex).invent.Object(i).amount)
                End If
            Next j
        Next i
    End If
End Sub

Private Sub HandleNpcInteractionByType(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
            If NpcList(NpcIndex).Comercia = 1 Then
                If Not EnsureUserAliveForNpcInteraction(UserIndex, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                If Not EnsureUserNotTrading(UserIndex) Then Exit Sub
                If Not EnsureNpcWithinDistance(UserIndex, NpcIndex, 4, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                Call PauseWalkingNpcForInteraction(NpcIndex, 15000)
                'Iniciamos la rutina pa' comerciar.
                Call IniciarComercioNPC(UserIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.Banquero Then
                If Not EnsureUserAliveForNpcInteraction(UserIndex, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                If Not EnsureUserNotTrading(UserIndex) Then Exit Sub
                If Not EnsureNpcWithinDistance(UserIndex, NpcIndex, 4, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                'A depositar de una
                Call IniciarBanco(UserIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.Pirata Then  'VIAJES
                If Not EnsureUserAliveForNpcInteraction(UserIndex, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                If Not EnsureUserNotTrading(UserIndex) Then Exit Sub
                If Not EnsureNpcWithinDistance(UserIndex, NpcIndex, 4, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO) Then
                    'Msg1070= Estas demasiado lejos del vendedor de pasajes.
                    Call WriteLocaleMsg(UserIndex, MSG_DEMASIADO_LEJOS_VENDEDOR_PASAJES, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If NpcList(NpcIndex).SoundOpen <> 0 Then
                    Call WritePlayWave(UserIndex, NpcList(NpcIndex).SoundOpen, NO_3D_SOUND, NO_3D_SOUND, , 1)
                End If
                'A depositar de unaIniciarTransporte
                Call WriteViajarForm(UserIndex, NpcIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.Revividor Or NpcList(NpcIndex).npcType = e_NPCType.ResucitadorNewbie Then
                Call HandleReviverNpcInteraction(UserIndex, NpcIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.Subastador Then
                If Not EnsureUserAliveForNpcInteraction(UserIndex, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                If Not EnsureNpcWithinDistance(UserIndex, NpcIndex, 1, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                Call PauseWalkingNpcForInteraction(NpcIndex, 20000)
                Call IniciarSubasta(UserIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.Quest Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, MSG_MUERTO, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call WritePlayWave(UserIndex, NpcList(NpcIndex).SoundOpen, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y, 2, 1)
                Call EnviarQuest(UserIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.Enlistador Then
                Call HandleFactionRecruiterNpcInteraction(UserIndex, NpcIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.Gobernador Then
                Call HandleGovernorNpcInteraction(UserIndex, NpcIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.EntregaPesca Then
                Call HandleFishingDeliveryNpcInteraction(UserIndex, NpcIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.AO20Shop Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, MSG_MUERTO, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
                Call WriteShopInit(UserIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.AO20ShopPjs Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, MSG_MUERTO, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
                Call WriteShopPjsInit(UserIndex)
            ElseIf NpcList(NpcIndex).npcType = e_NPCType.EventMaster Then
                If Not EnsureUserAliveForNpcInteraction(UserIndex, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                If Not EnsureNpcWithinDistance(UserIndex, NpcIndex, 4, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                Call WriteUpdateLobbyList(UserIndex)
            ElseIf NpcList(NpcIndex).Craftea > 0 Then
                If Not EnsureUserAliveForNpcInteraction(UserIndex, e_FontTypeNames.FONTTYPE_INFOIAO) Then Exit Sub
                If Not EnsureNpcWithinDistance(UserIndex, NpcIndex, 3, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO) Then Exit Sub
                UserList(UserIndex).flags.Crafteando = NpcList(NpcIndex).Craftea
                Call WriteOpenCrafting(UserIndex, NpcList(NpcIndex).Craftea)
            End If
End Sub

Private Function TryHandleObjectWorldAction(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If MapData(Map, x, y).ObjInfo.ObjIndex <= 0 Then Exit Function

            UserList(UserIndex).flags.TargetObj = MapData(Map, x, y).ObjInfo.ObjIndex
            Select Case ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).OBJType
                Case e_OBJType.otDoors 'Es una puerta
                    Call HandleDoorAction(Map, x, y, UserIndex)
                Case e_OBJType.otSignBoards 'Es un cartel
                    Call HandleSignAction(Map, x, y, UserIndex)
                Case e_OBJType.otMail 'Es un cartel
                    'Call AccionParaCorreo(Map, x, Y, UserIndex)
                    ' Msg586=El correo está temporalmente deshabilitado.
                    Call WriteLocaleMsg(UserIndex, MSG_CORREO_TEMPORALMENTE_DESHABILITADO, e_FontTypeNames.FONTTYPE_EJECUCION)
                Case e_OBJType.otAnvil 'Pozos
                    Call HandleAnvilAction(Map, x, y, UserIndex)
                Case e_OBJType.otWood    'Leña
                    If MapData(Map, x, y).ObjInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
                        Call HandleCampfireTwigAction(Map, x, y, UserIndex)
                    End If
                Case Else
                    GoTo TryHandleObjectWorldAction_Handled
            End Select
TryHandleObjectWorldAction_Handled:
    TryHandleObjectWorldAction = True
End Function

Private Function TryHandleMultiTileObjectWorldAction(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    If TryHandleObjectAtOffset(UserIndex, Map, x + 1, y) Then
        TryHandleMultiTileObjectWorldAction = True
        Exit Function
    End If

    If TryHandleObjectAtOffset(UserIndex, Map, x + 1, y + 1) Then
        TryHandleMultiTileObjectWorldAction = True
        Exit Function
    End If

    If TryHandleObjectAtOffset(UserIndex, Map, x, y + 1) Then
        TryHandleMultiTileObjectWorldAction = True
        Exit Function
    End If
End Function

Private Function TryHandleObjectAtOffset(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If Not InMapBounds(Map, x, y) Then Exit Function
    If MapData(Map, x, y).ObjInfo.ObjIndex <= 0 Then Exit Function

    UserList(UserIndex).flags.TargetObj = MapData(Map, x, y).ObjInfo.ObjIndex
    Select Case ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).OBJType
        Case e_OBJType.otDoors 'Es una puerta
            Call HandleDoorAction(Map, x, y, UserIndex)
            TryHandleObjectAtOffset = True
    End Select
End Function
Sub HandleAnvilAction(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
    On Error GoTo AccionParaYunque_Err
    Dim pos As t_WorldPos
    pos.Map = Map
    pos.x = x
    pos.y = y
    If Distancia(pos, UserList(UserIndex).pos) > 2 Then
        ' Msg8=Estas demasiado lejos.
        Call WriteLocaleMsg(UserIndex, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).invent.EquippedWorkingToolObjIndex = 0 Then
        'Msg1071= Debes tener equipado un martillo de herrero para trabajar con el yunque.
        Call WriteLocaleMsg(UserIndex, MSG_DEBES_TENER_EQUIPADO_MARTILLO_HERRERO_TRABAJAR_YUNQUE, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If ObjData(UserList(UserIndex).invent.EquippedWorkingToolObjIndex).Subtipo <> e_WorkingToolSubType.SmithHammer Then
        'Msg1072= La herramienta que tienes no es la correcta, necesitas un martillo de herrero para poder trabajar.
        Call WriteLocaleMsg(UserIndex, MSG_NO_HERRAMIENTA_TIENES_CORRECTA_NECESITAS_MARTILLO_HERRERO_PODER_TRABAJAR, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    UserList(UserIndex).flags.TargetMap = Map
    UserList(UserIndex).flags.TargetX = x
    UserList(UserIndex).flags.TargetY = y
    Call EnivarArmasConstruibles(UserIndex)
    Call EnivarArmadurasConstruibles(UserIndex)
    Call SendCraftableElementRunes(UserIndex)
    Call WriteShowBlacksmithForm(UserIndex)
    Exit Sub
AccionParaYunque_Err:
    Call TraceError(Err.Number, Err.Description, "WorldActions.HandleAnvilAction", Erl)
End Sub

Sub HandleDoorAction(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte, ByVal UserIndex As Integer, Optional ByVal SinDistancia As Boolean = False)
    On Error GoTo Handler
    Dim puerta As t_ObjData 'ver ReyarB
    If Distance(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, x, y) > 2 And Not SinDistancia Then
        Call WriteLocaleMsg(UserIndex, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    puerta = ObjData(MapData(Map, x, y).ObjInfo.ObjIndex)
    If puerta.Llave = 1 And Not SinDistancia Then
        If puerta.GrhIndex = 11445 Or puerta.GrhIndex = 11444 Or puerta.GrhIndex = 59878 Or puerta.GrhIndex = 59877 Then
            'Msg1073= Al parecer, alguien cerró esta puerta. Debe haber algún interruptor por algún lado...
            Call WriteLocaleMsg(UserIndex, MSG_PARECER_ALGUIEN_CERRO_PUERTA_DEBE_HABER_ALGUN_INTERRUPTOR, e_FontTypeNames.FONTTYPE_INFO)
        Else
            'Msg1074= La puerta esta cerrada con llave.
            Call WriteLocaleMsg(UserIndex, MSG_PUERTA_CERRADA_LLAVE, e_FontTypeNames.FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    If puerta.Cerrada = 1 Then 'Abre la puerta
        MapData(Map, x, y).ObjInfo.ObjIndex = puerta.IndexAbierta
        Call BloquearPuerta(Map, x, y, False)
        If puerta.GrhIndex = 11445 Or puerta.GrhIndex = 11444 Or puerta.GrhIndex = 59878 Or puerta.GrhIndex = 59877 Then
            'Msg1075= Has abierto la compuerta del ducto.
            Call WriteLocaleMsg(UserIndex, MSG_ABIERTO_COMPUERTA_DUCTO, e_FontTypeNames.FONTTYPE_INFO)
        End If
    Else 'Cierra puerta
        MapData(Map, x, y).ObjInfo.ObjIndex = puerta.IndexCerrada
        Call BloquearPuerta(Map, x, y, True)
    End If
    If ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).Subtipo = 1 Then
        Call HandleDoorAction(Map, x - 3, y + 1, UserIndex, True)
    End If
    Call modSendData.SendToAreaByPos(Map, x, y, PrepareMessageObjectCreate(MapData(Map, x, y).ObjInfo.ObjIndex, MapData(Map, x, y).ObjInfo.amount, x, y))
    If puerta.GrhIndex = 11445 Or puerta.GrhIndex = 11444 Or puerta.GrhIndex = 59878 Or puerta.GrhIndex = 59877 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA_DUCTO, x, y))
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, x, y))
    End If
    UserList(UserIndex).flags.TargetObj = MapData(Map, x, y).ObjInfo.ObjIndex
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "WorldActions.HandleDoorAction", Erl)
End Sub

Sub HandleNpcDoorAction(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte, ByVal NpcIndex As Integer)
    On Error GoTo Handler
    Dim puerta As t_ObjData 'ver ReyarB
    puerta = ObjData(MapData(Map, x, y).ObjInfo.ObjIndex)
    If puerta.Cerrada = 1 Then 'Abre la puerta
        MapData(Map, x, y).ObjInfo.ObjIndex = puerta.IndexAbierta
        Call BloquearPuerta(Map, x, y, False)
    Else 'Cierra puerta
        MapData(Map, x, y).ObjInfo.ObjIndex = puerta.IndexCerrada
        Call BloquearPuerta(Map, x, y, True)
    End If
    Call modSendData.SendToAreaByPos(Map, x, y, PrepareMessageObjectCreate(MapData(Map, x, y).ObjInfo.ObjIndex, MapData(Map, x, y).ObjInfo.amount, x, y))
    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_PUERTA, x, y))
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "WorldActions.HandleNpcDoorAction", Erl)
End Sub

Sub HandleSignAction(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
    On Error GoTo Handler
    If Len(ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).texto) > 0 Then
        Call WriteShowSignal(UserIndex, MapData(Map, x, y).ObjInfo.ObjIndex)
    Else
        Call WriteShowPapiro(UserIndex)
    End If
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "WorldActions.HandleSignAction", Erl)
End Sub

Sub HandleCampfireTwigAction(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
    On Error GoTo Handler
    Dim Suerte As Byte
    Dim exito  As Byte
    Dim pos    As t_WorldPos
    pos.Map = Map
    pos.x = x
    pos.y = y
    With UserList(UserIndex)
        If Distancia(pos, .pos) > 2 Then
            ' Msg8=Estas demasiado lejos.
            Call WriteLocaleMsg(UserIndex, MSG_SACERDOTE_PUEDE_CURARTE_DEBIDO_DEMASIADO_LEJOS, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapInfo(Map).lluvia And Lloviendo Then
            'Msg1076= Esta lloviendo, no podés encender una fogata aquí.
            Call WriteLocaleMsg(UserIndex, MSG_NO_LLOVIENDO_PODES_ENCENDER_FOGATA_AQUI, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapData(Map, x, y).trigger = e_Trigger.ZonaSegura Or MapInfo(Map).Seguro = 1 Then
            'Msg1077= En zona segura no podés hacer fogatas.
            Call WriteLocaleMsg(UserIndex, MSG_NO_ZONA_SEGURA_PODES_HACER_FOGATAS, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapData(Map, x - 1, y).ObjInfo.ObjIndex = FOGATA Or MapData(Map, x + 1, y).ObjInfo.ObjIndex = FOGATA Or MapData(Map, x, y - 1).ObjInfo.ObjIndex = FOGATA Or MapData( _
           Map, x, y + 1).ObjInfo.ObjIndex = FOGATA Then
            'Msg1078= Debes alejarte un poco de la otra fogata.
            Call WriteLocaleMsg(UserIndex, MSG_DEBES_ALEJARTE_POCO_OTRA_FOGATA, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Select Case .Stats.UserSkills(Supervivencia)
            Case Is < 6
                Suerte = 4
            Case Is < 10
                Suerte = 3
            Case Else
                Suerte = 2
        End Select
        exito = RandomNumber(1, Suerte)
        If exito = 1 Then
            If MapInfo(.pos.Map).zone <> Ciudad Then
                Dim obj As t_Obj
                obj.ObjIndex = FOGATA
                obj.amount = 1
                'Msg1079= Has prendido la fogata.
                Call WriteLocaleMsg(UserIndex, MSG_PRENDIDO_FOGATA, e_FontTypeNames.FONTTYPE_INFO)
                Call MakeObj(obj, Map, x, y)
            Else
                'Msg1080= La ley impide realizar fogatas en las ciudades.
                Call WriteLocaleMsg(UserIndex, MSG_LEY_IMPIDE_REALIZAR_FOGATAS_CIUDADES, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            'Msg1081= No has podido hacer fuego.
            Call WriteLocaleMsg(UserIndex, MSG_NO_PODIDO_HACER_FUEGO, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Call SubirSkill(UserIndex, Supervivencia)
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "WorldActions.HandleCampfireTwigAction", Erl)
End Sub
