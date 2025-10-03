Attribute VB_Name = "Acciones"
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

Public Function get_map_name(ByVal Map As Long) As String
    On Error GoTo get_map_name_Err
    get_map_name = MapInfo(Map).map_name
    Exit Function
get_map_name_Err:
    Call TraceError(Err.Number, Err.Description, "Acciones.get_map_name", Erl)
End Function

Public Function PuedeUsarObjeto(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByVal writeInConsole As Boolean = False) As Byte
    On Error GoTo PuedeUsarObjeto_Err
    Dim Objeto As t_ObjData
    Dim Msg    As String
    Dim Extra  As String
    Dim i      As Long
    Extra = vbNullString
    Objeto = ObjData(ObjIndex)
    With UserList(UserIndex)
        If EsGM(UserIndex) Then
            PuedeUsarObjeto = 0
            Msg = vbNullString
            Exit Function
        End If
        ' Keep the original priority: first match wins
        If Objeto.Newbie = 1 And Not EsNewbie(UserIndex) Then
            PuedeUsarObjeto = 7
            Msg = "679" ' Only newbies can use this item.
        ElseIf .Stats.ELV < Objeto.MinELV Then
            PuedeUsarObjeto = 6
            Msg = "1926" ' Need level {0}
            Extra = CStr(Objeto.MinELV)
        ElseIf .Stats.ELV > Objeto.MaxLEV And Objeto.MaxLEV > 0 Then
            PuedeUsarObjeto = 6
            Msg = "1982" ' Not for level {0} or higher
            Extra = CStr(Objeto.MaxLEV)
        ElseIf Not FaccionPuedeUsarItem(UserIndex, ObjIndex) And JerarquiaPuedeUsarItem(UserIndex, ObjIndex) Then
            PuedeUsarObjeto = 3
            Msg = "416"  ' Faction doesn't allow it.
        ElseIf Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
            PuedeUsarObjeto = 2
            Msg = "265"  ' Class cannot use this item.
        ElseIf Not SexoPuedeUsarItem(UserIndex, ObjIndex) Then
            PuedeUsarObjeto = 1
            Msg = "267"  ' Sex cannot use this item.
        ElseIf Not RazaPuedeUsarItem(UserIndex, ObjIndex) Then
            PuedeUsarObjeto = 5
            Msg = "266"  ' Race cannot use this item.
        ElseIf (Objeto.SkillIndex > 0) Then
            If (.Stats.UserSkills(Objeto.SkillIndex) < Objeto.SkillRequerido) Then
                PuedeUsarObjeto = 4
                Msg = "NEED_SKILL_POINTS" ' e.g. "Necesitas {0} puntos en {1}..."
                Extra = CStr(Objeto.SkillRequerido) & "¬" & SkillsNames(Objeto.SkillIndex)
            Else
                PuedeUsarObjeto = 0
                Msg = vbNullString
            End If
        Else
            PuedeUsarObjeto = 0
            Msg = vbNullString
        End If
        ' Only emit when we actually have a message
        If Msg <> vbNullString Then
            If writeInConsole Then
                Call WriteLocaleMsg(UserIndex, Msg, e_FontTypeNames.FONTTYPE_INFO, Extra)
            End If
        End If
    End With
    Exit Function
PuedeUsarObjeto_Err:
    Call TraceError(Err.Number, Err.Description, "Acciones.PuedeUsarObjeto", Erl)
End Function

Public Sub CompletarAccionFin(ByVal UserIndex As Integer)
    On Error GoTo CompletarAccionFin_Err
    Dim obj  As t_ObjData
    Dim Slot As Byte
    Select Case UserList(UserIndex).Accion.TipoAccion
        Case e_AccionBarra.Runa
            obj = ObjData(UserList(UserIndex).Accion.RunaObj)
            Slot = UserList(UserIndex).Accion.ObjSlot
            Select Case obj.TipoRuna
                Case e_RuneType.ReturnHome 'lleva a la ciudad de origen vivo o muerto
                    Dim DeDonde As t_CityWorldPos
                    Dim Map     As Integer
                    Dim x       As Byte
                    Dim y       As Byte
                    If UserList(UserIndex).flags.Muerto = 0 Then
                        Select Case UserList(UserIndex).Hogar
                            Case e_Ciudad.cUllathorpe
                                DeDonde = CityUllathorpe
                            Case e_Ciudad.cNix
                                DeDonde = CityNix
                            Case e_Ciudad.cBanderbill
                                DeDonde = CityBanderbill
                            Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                                DeDonde = CityLindos
                            Case e_Ciudad.cArghal
                                DeDonde = CityArghal
                            Case e_Ciudad.cForgat
                                DeDonde = CityForgat
                            Case e_Ciudad.cEldoria
                                DeDonde = CityEldoria
                            Case e_Ciudad.cArkhein
                                DeDonde = CityArkhein
                            Case e_Ciudad.cPenthar
                                DeDonde = CityPenthar
                            Case Else
                                DeDonde = CityUllathorpe
                        End Select
                        Map = DeDonde.Map
                        x = DeDonde.x
                        y = DeDonde.y
                    Else
                        If MapInfo(UserList(UserIndex).pos.Map).ResuCiudad <> 0 Then
                            Select Case MapInfo(UserList(UserIndex).pos.Map).ResuCiudad
                                Case e_Ciudad.cUllathorpe
                                    DeDonde = CityUllathorpe
                                Case e_Ciudad.cNix
                                    DeDonde = CityNix
                                Case e_Ciudad.cBanderbill
                                    DeDonde = CityBanderbill
                                Case e_Ciudad.cLindos
                                    DeDonde = CityLindos
                                Case e_Ciudad.cArghal
                                    DeDonde = CityArghal
                                Case e_Ciudad.cForgat
                                    DeDonde = CityForgat
                                Case e_Ciudad.cArkhein
                                    DeDonde = CityArkhein
                                Case e_Ciudad.cEldoria
                                    DeDonde = CityEldoria
                                Case e_Ciudad.cPenthar
                                    DeDonde = CityPenthar
                                Case Else
                                    DeDonde = CityUllathorpe
                            End Select
                        Else
                            Select Case UserList(UserIndex).Hogar
                                Case e_Ciudad.cUllathorpe
                                    DeDonde = CityUllathorpe
                                Case e_Ciudad.cNix
                                    DeDonde = CityNix
                                Case e_Ciudad.cBanderbill
                                    DeDonde = CityBanderbill
                                Case e_Ciudad.cLindos
                                    DeDonde = CityLindos
                                Case e_Ciudad.cArghal
                                    DeDonde = CityArghal
                                Case e_Ciudad.cForgat
                                    DeDonde = CityForgat
                                Case e_Ciudad.cArkhein
                                    DeDonde = CityArkhein
                                Case e_Ciudad.cEldoria
                                    DeDonde = CityEldoria
                                Case e_Ciudad.cPenthar
                                    DeDonde = CityPenthar
                                Case Else
                                    DeDonde = CityUllathorpe
                            End Select
                        End If
                        Map = DeDonde.MapaResu
                        x = DeDonde.ResuX
                        y = DeDonde.ResuY
                        Dim Resu As Boolean
                        Resu = True
                    End If
                    Call FindLegalPos(UserIndex, Map, x, y)
                    Call WarpUserChar(UserIndex, Map, x, y, True)
                    'Msg1065= Has regresado a tu ciudad de origen.
                    Call WriteLocaleMsg(UserIndex, 1065, e_FontTypeNames.FONTTYPE_WARNING)
                    'Call WriteFlashScreen(UserIndex, &HA4FFFF, 150, True)
                    If UserList(UserIndex).flags.Navegando = 1 Then
                        Dim barca As t_ObjData
                        barca = ObjData(UserList(UserIndex).invent.EquippedShipObjIndex)
                        Call DoNavega(UserIndex, barca, UserList(UserIndex).invent.EquippedShipSlot)
                    End If
                    If Resu Then
                        UserList(UserIndex).Counters.TimerBarra = 5
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.Resucitar, UserList( _
                                UserIndex).Counters.TimerBarra, False))
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.charindex, UserList(UserIndex).Counters.TimerBarra, _
                                e_AccionBarra.Resucitar))
                        UserList(UserIndex).Accion.AccionPendiente = True
                        UserList(UserIndex).Accion.Particula = e_ParticleEffects.Resucitar
                        UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.Resucitar
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                        'Msg82=El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...
                        Call WriteLocaleMsg(UserIndex, 82, e_FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                    If Not Resu Then
                        UserList(UserIndex).Accion.AccionPendiente = False
                        UserList(UserIndex).Accion.Particula = 0
                        UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
                    End If
                    UserList(UserIndex).Accion.HechizoPendiente = 0
                    UserList(UserIndex).Accion.RunaObj = 0
                    UserList(UserIndex).Accion.ObjSlot = 0
                Case e_RuneType.Escape
                    Map = obj.HastaMap
                    x = obj.HastaX
                    y = obj.HastaY
                    If obj.DesdeMap = 0 Then
                        Call FindLegalPos(UserIndex, Map, x, y)
                        Call WarpUserChar(UserIndex, Map, x, y, True)
                        'Msg1066= Te has teletransportado por el mundo.
                        Call WriteLocaleMsg(UserIndex, 1066, e_FontTypeNames.FONTTYPE_WARNING)
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        Call UpdateUserInv(False, UserIndex, Slot)
                    Else
                        If UserList(UserIndex).pos.Map <> obj.DesdeMap Then
                            'Msg1067= Esta runa no puede ser usada desde aquí.
                            Call WriteLocaleMsg(UserIndex, 1067, e_FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Call UpdateUserInv(False, UserIndex, Slot)
                            Call FindLegalPos(UserIndex, Map, x, y)
                            Call WarpUserChar(UserIndex, Map, x, y, True)
                            'Msg1068= Te has teletransportado por el mundo.
                            Call WriteLocaleMsg(UserIndex, 1068, e_FontTypeNames.FONTTYPE_WARNING)
                        End If
                    End If
                    UserList(UserIndex).Accion.Particula = 0
                    UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
                    UserList(UserIndex).Accion.HechizoPendiente = 0
                    UserList(UserIndex).Accion.RunaObj = 0
                    UserList(UserIndex).Accion.ObjSlot = 0
                    UserList(UserIndex).Accion.AccionPendiente = False
                Case e_RuneType.MesonSafePassage
                    If UserList(UserIndex).pos.Map = MAP_MESON_HOSTIGADO Or UserList(UserIndex).pos.Map = MAP_MESON_HOSTIGADO_TRADING_ZONE Then
                        'mensaje de error de "no puedes usar la runa estando en el meson"
                        Call WriteLocaleMsg(UserIndex, 2081, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If obj.HastaMap <> MAP_MESON_HOSTIGADO Then
                        'mensaje de error de runa invalida, hay algo mal dateado llamar a un gm o avisar a soporte
                        Call WriteLocaleMsg(UserIndex, 2080, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    UserList(UserIndex).flags.ReturnPos = UserList(UserIndex).pos
                    Map = obj.HastaMap
                    x = obj.HastaX
                    y = obj.HastaY
                    Call WarpUserChar(UserIndex, Map, x, y, True)
                    'Msg1066= Te has teletransportado por el mundo.
                    Call WriteLocaleMsg(UserIndex, 1066, e_FontTypeNames.FONTTYPE_WARNING)
                    UserList(UserIndex).Accion.Particula = 0
                    UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
                    UserList(UserIndex).Accion.HechizoPendiente = 0
                    UserList(UserIndex).Accion.RunaObj = 0
                    UserList(UserIndex).Accion.ObjSlot = 0
                    UserList(UserIndex).Accion.AccionPendiente = False
            End Select
        Case e_AccionBarra.Hogar
            Call HomeArrival(UserIndex)
            UserList(UserIndex).Accion.AccionPendiente = False
            UserList(UserIndex).Accion.Particula = 0
            UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
        Case e_AccionBarra.Intermundia
            If UserList(UserIndex).flags.Muerto = 0 Then
                Dim uh As Integer
                Dim Mapaf, Xf, Yf As Integer
                uh = UserList(UserIndex).Accion.HechizoPendiente
                Mapaf = Hechizos(uh).TeleportXMap
                Xf = Hechizos(uh).TeleportXX
                Yf = Hechizos(uh).TeleportXY
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(uh).wav, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))  'Esta linea faltaba. Pablo (ToxicWaste)
                'Msg1069= ¡Has abierto la puerta a intermundia!
                Call WriteLocaleMsg(UserIndex, 1069, e_FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_GraphicEffects.Runa, -1, True))
                UserList(UserIndex).flags.Portal = 10
                UserList(UserIndex).flags.PortalMDestino = Mapaf
                UserList(UserIndex).flags.PortalYDestino = Xf
                UserList(UserIndex).flags.PortalXDestino = Yf
                Dim Mapa As Integer
                Mapa = UserList(UserIndex).flags.PortalM
                x = UserList(UserIndex).flags.PortalX
                y = UserList(UserIndex).flags.PortalY
                MapData(Mapa, x, y).Particula = e_ParticleEffects.HaloGreen
                MapData(Mapa, x, y).TimeParticula = -1
                MapData(Mapa, x, y).TileExit.Map = UserList(UserIndex).flags.PortalMDestino
                MapData(Mapa, x, y).TileExit.x = UserList(UserIndex).flags.PortalXDestino
                MapData(Mapa, x, y).TileExit.y = UserList(UserIndex).flags.PortalYDestino
                Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(x, y, e_ParticleEffects.HaloGreen, -1))
                Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageLightFXToFloor(x, y, &HFF80C0, 105))
            End If
            UserList(UserIndex).Accion.Particula = 0
            UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
            UserList(UserIndex).Accion.HechizoPendiente = 0
            UserList(UserIndex).Accion.RunaObj = 0
            UserList(UserIndex).Accion.ObjSlot = 0
            UserList(UserIndex).Accion.AccionPendiente = False
            '
        Case e_AccionBarra.Resucitar
            ' Msg585=¡Has sido resucitado!
            Call WriteLocaleMsg(UserIndex, 585, e_FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.Resucitar, 250, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("117", UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
            Call RevivirUsuario(UserIndex, True)
            UserList(UserIndex).Accion.Particula = 0
            UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
            UserList(UserIndex).Accion.HechizoPendiente = 0
            UserList(UserIndex).Accion.RunaObj = 0
            UserList(UserIndex).Accion.ObjSlot = 0
            UserList(UserIndex).Accion.AccionPendiente = False
    End Select
    Exit Sub
CompletarAccionFin_Err:
    Call TraceError(Err.Number, Err.Description, "Acciones.CompletarAccionFin", Erl)
End Sub

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo Accion_Err
    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(UserIndex).pos.y - y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).pos.x - x) > RANGO_VISION_X) Then
        Exit Sub
    End If
    If UserIndex <= 0 Then Exit Sub
    '¿Posicion valida?
    If InMapBounds(Map, x, y) Then
        Dim FoundChar      As Byte
        Dim FoundSomething As Byte
        Dim TempCharIndex  As Integer
        If MapData(Map, x, y).NpcIndex > 0 Then     'Acciones NPCs
            TempCharIndex = MapData(Map, x, y).NpcIndex
            'Set the target NPC
            Call SetNpcRef(UserList(UserIndex).flags.TargetNPC, TempCharIndex)
            UserList(UserIndex).flags.TargetNpcTipo = NpcList(TempCharIndex).npcType
            If NpcList(TempCharIndex).Comercia = 1 Then
                '¿Esta el user muerto? Si es asi no puede comerciar
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Is it already in commerce mode??
                If UserList(UserIndex).flags.Comerciando Then
                    Exit Sub
                End If
                If Distancia(NpcList(TempCharIndex).pos, UserList(UserIndex).pos) > 4 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If NpcList(TempCharIndex).Movement = e_TipoAI.Caminata Then
                    NpcList(TempCharIndex).Contadores.IntervaloMovimiento = AddMod32(GetTickCountRaw(), 15000)
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarComercioNPC(UserIndex)
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.Banquero Then
                '¿Esta el user muerto? Si es asi no puede comerciar
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Is it already in commerce mode??
                If UserList(UserIndex).flags.Comerciando Then
                    Exit Sub
                End If
                If Distancia(NpcList(TempCharIndex).pos, UserList(UserIndex).pos) > 4 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'A depositar de una
                Call IniciarBanco(UserIndex)
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.Pirata Then  'VIAJES
                '¿Esta el user muerto? Si es asi no puede comerciar
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Is it already in commerce mode??
                If UserList(UserIndex).flags.Comerciando Then
                    Exit Sub
                End If
                If Distancia(NpcList(TempCharIndex).pos, UserList(UserIndex).pos) > 4 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg1070= Estas demasiado lejos del vendedor de pasajes.
                    Call WriteLocaleMsg(UserIndex, 1070, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If NpcList(TempCharIndex).SoundOpen <> 0 Then
                    Call WritePlayWave(UserIndex, NpcList(TempCharIndex).SoundOpen, NO_3D_SOUND, NO_3D_SOUND, , 1)
                End If
                'A depositar de unaIniciarTransporte
                Call WriteViajarForm(UserIndex, TempCharIndex)
                Exit Sub
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.Revividor Or NpcList(TempCharIndex).npcType = e_NPCType.ResucitadorNewbie Then
                If Distancia(UserList(UserIndex).pos, NpcList(TempCharIndex).pos) > 5 Then
                    'Msg8=El sacerdote no puede curarte debido a que estas demasiado lejos.
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                '  Hacemos que se detenga a hablar un momento :P
                If NpcList(TempCharIndex).Movement = Caminata Then
                    NpcList(TempCharIndex).Contadores.IntervaloMovimiento = AddMod32(GetTickCountRaw(), 5000) ' 5 segundos
                End If
                UserList(UserIndex).flags.Envenenado = 0
                UserList(UserIndex).flags.Incinerado = 0
                'Revivimos si es necesario
                If UserList(UserIndex).flags.Muerto = 1 And (NpcList(TempCharIndex).npcType = e_NPCType.Revividor Or EsNewbie(UserIndex)) Then
                    ' Msg585=¡Has sido resucitado!
                    Call WriteLocaleMsg(UserIndex, 585, e_FontTypeNames.FONTTYPE_INFO)
                    Call RevivirUsuario(UserIndex)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.Resucitar, 30, False))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                Else
                    'curamos totalmente
                    If UserList(UserIndex).Stats.MinHp <> UserList(UserIndex).Stats.MaxHp Then
                        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
                        Call WritePlayWave(UserIndex, 117, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y)
                        'Msg83=El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...¡Has sido curado!
                        Call WriteLocaleMsg(UserIndex, 83, e_FontTypeNames.FONTTYPE_INFOIAO)
                        Call WriteUpdateUserStats(UserIndex)
                        If Status(UserIndex) = 4 Or Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.CurarCrimi, 100, False))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.CurarCrimi, 100, False))
                        End If
                    End If
                End If
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.Subastador Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Distancia(NpcList(TempCharIndex).pos, UserList(UserIndex).pos) > 1 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If NpcList(TempCharIndex).Movement = Caminata Then
                    NpcList(TempCharIndex).Contadores.IntervaloMovimiento = AddMod32(GetTickCountRaw(), 20000)
                End If
                Call IniciarSubasta(UserIndex)
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.Quest Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call WritePlayWave(UserIndex, NpcList(TempCharIndex).SoundOpen, NpcList(TempCharIndex).pos.x, NpcList(TempCharIndex).pos.y, 2, 1)
                Call EnviarQuest(UserIndex)
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.Enlistador Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Distancia(NpcList(TempCharIndex).pos, UserList(UserIndex).pos) > 3 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If NpcList(TempCharIndex).flags.Faccion = 0 Then
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
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.Gobernador Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
                If Distancia(NpcList(TempCharIndex).pos, UserList(UserIndex).pos) > 3 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg8=Estas demasiado lejos del gobernador.
                    Exit Sub
                End If
                Dim DeDonde    As String
                Dim Gobernador As t_Npc
                Gobernador = NpcList(TempCharIndex)
                If UserList(UserIndex).Hogar = Gobernador.GobernadorDe Then
                    Call WriteLocaleChatOverHead(UserIndex, 1349, "", Gobernador.Char.charindex, vbWhite) ' Msg1349=Ya perteneces a esta ciudad. Gracias por ser uno más de nosotros.
                    Exit Sub
                End If
                If UserList(UserIndex).Faccion.Status = 0 Or UserList(UserIndex).Faccion.Status = 2 Then
                    If Gobernador.GobernadorDe = e_Ciudad.cBanderbill Then
                        Call WriteLocaleChatOverHead(UserIndex, "1350", "", Gobernador.Char.charindex, vbWhite) ' Msg1350=Aquí no aceptamos criminales.
                        Exit Sub
                    End If
                End If
                If UserList(UserIndex).Faccion.Status = 3 Or UserList(UserIndex).Faccion.Status = 1 Then
                    If Gobernador.GobernadorDe = e_Ciudad.cArkhein Then
                        Call WriteLocaleChatOverHead(UserIndex, "1351", "", Gobernador.Char.charindex, vbWhite) ' Msg1351=¡¡Sal de aquí ciudadano asqueroso!!
                        Exit Sub
                    End If
                End If
                If UserList(UserIndex).Hogar <> Gobernador.GobernadorDe Then
                    UserList(UserIndex).PosibleHogar = Gobernador.GobernadorDe
                    Select Case UserList(UserIndex).PosibleHogar
                        Case e_Ciudad.cUllathorpe
                            DeDonde = "Ullathorpe"
                        Case e_Ciudad.cNix
                            DeDonde = "Nix"
                        Case e_Ciudad.cBanderbill
                            DeDonde = "Banderbill"
                        Case e_Ciudad.cLindos
                            DeDonde = "Lindos"
                        Case e_Ciudad.cArghal
                            DeDonde = "Arghal"
                        Case e_Ciudad.cForgat
                            DeDonde = "Forgat"
                        Case e_Ciudad.cEldoria
                            DeDonde = "Eldoria"
                        Case e_Ciudad.cArkhein
                            DeDonde = "Arkhein"
                        Case e_Ciudad.cPenthar
                            DeDonde = "Penthar"
                        Case Else
                            DeDonde = "Ullathorpe"
                    End Select
                    UserList(UserIndex).flags.pregunta = 3
                    Call WritePreguntaBox(UserIndex, 1592, DeDonde)
                End If
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.EntregaPesca Then
                Dim i                   As Integer, j As Integer
                Dim PuntosTotales       As Long
                Dim CantPecesEspeciales As Long
                Dim OroTotal            As Long
                CantPecesEspeciales = UBound(PecesEspeciales)
                If CantPecesEspeciales > 0 Then
                    For i = 1 To MAX_INVENTORY_SLOTS
                        For j = 1 To CantPecesEspeciales
                            If UserList(UserIndex).invent.Object(i).ObjIndex = PecesEspeciales(j).ObjIndex Then
                                PuntosTotales = PuntosTotales + (ObjData(UserList(UserIndex).invent.Object(i).ObjIndex).PuntosPesca * UserList(UserIndex).invent.Object(i).amount)
                                OroTotal = OroTotal + (ObjData(UserList(UserIndex).invent.Object(i).ObjIndex).Valor * UserList(UserIndex).invent.Object(i).amount)
                            End If
                        Next j
                    Next i
                End If
                If PuntosTotales > 0 Then
                    UserList(UserIndex).flags.pregunta = 5
                    Call WritePreguntaBox(UserIndex, 1593, PuntosTotales & "¬" & PonerPuntos(OroTotal * 1.2)) 'Msg1593= Tienes un total de ¬1 puntos y ¬2 monedas de oro para reclamar, ¿Deseas aceptar?
                Else
                    Dim charindexstr As Integer
                    charindexstr = str(NpcList(TempCharIndex).Char.charindex)
                    Call WriteLocaleChatOverHead(UserIndex, "1352", "", charindexstr, &HFFFF00) ' Msg1352=No tienes ningún trofeo de pesca para entregar.
                End If
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.AO20Shop Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
                Call WriteShopInit(UserIndex)
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.AO20ShopPjs Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
                Call WriteShopPjsInit(UserIndex)
            ElseIf NpcList(TempCharIndex).npcType = e_NPCType.EventMaster Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Distancia(NpcList(TempCharIndex).pos, UserList(UserIndex).pos) > 4 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call WriteUpdateLobbyList(UserIndex)
            ElseIf NpcList(TempCharIndex).Craftea > 0 Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
                If Distancia(NpcList(TempCharIndex).pos, UserList(UserIndex).pos) > 3 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                UserList(UserIndex).flags.Crafteando = NpcList(TempCharIndex).Craftea
                Call WriteOpenCrafting(UserIndex, NpcList(TempCharIndex).Craftea)
            End If
            '¿Es un obj?
        ElseIf MapData(Map, x, y).ObjInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, x, y).ObjInfo.ObjIndex
            Select Case ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).OBJType
                Case e_OBJType.otDoors 'Es una puerta
                    Call AccionParaPuerta(Map, x, y, UserIndex)
                Case e_OBJType.otSignBoards 'Es un cartel
                    Call AccionParaCartel(Map, x, y, UserIndex)
                Case e_OBJType.otMail 'Es un cartel
                    'Call AccionParaCorreo(Map, x, Y, UserIndex)
                    ' Msg586=El correo está temporalmente deshabilitado.
                    Call WriteLocaleMsg(UserIndex, 586, e_FontTypeNames.FONTTYPE_EJECUCION)
                Case e_OBJType.otAnvil 'Pozos
                    Call AccionParaYunque(Map, x, y, UserIndex)
                Case e_OBJType.otWood    'Leña
                    If MapData(Map, x, y).ObjInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
                        Call AccionParaRamita(Map, x, y, UserIndex)
                    End If
                Case Else
                    Exit Sub
            End Select
            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
        ElseIf MapData(Map, x + 1, y).ObjInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, x + 1, y).ObjInfo.ObjIndex
            Select Case ObjData(MapData(Map, x + 1, y).ObjInfo.ObjIndex).OBJType
                Case e_OBJType.otDoors 'Es una puerta
                    Call AccionParaPuerta(Map, x + 1, y, UserIndex)
            End Select
        ElseIf MapData(Map, x + 1, y + 1).ObjInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, x + 1, y + 1).ObjInfo.ObjIndex
            Select Case ObjData(MapData(Map, x + 1, y + 1).ObjInfo.ObjIndex).OBJType
                Case e_OBJType.otDoors 'Es una puerta
                    Call AccionParaPuerta(Map, x + 1, y + 1, UserIndex)
            End Select
        ElseIf MapData(Map, x, y + 1).ObjInfo.ObjIndex > 0 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, x, y + 1).ObjInfo.ObjIndex
            Select Case ObjData(MapData(Map, x, y + 1).ObjInfo.ObjIndex).OBJType
                Case e_OBJType.otDoors 'Es una puerta
                    Call AccionParaPuerta(Map, x, y + 1, UserIndex)
            End Select
        End If
    End If
    Exit Sub
Accion_Err:
    Call TraceError(Err.Number, Err.Description, "Acciones.Accion", Erl)
End Sub

Sub AccionParaYunque(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
    On Error GoTo AccionParaYunque_Err
    Dim pos As t_WorldPos
    pos.Map = Map
    pos.x = x
    pos.y = y
    If Distancia(pos, UserList(UserIndex).pos) > 2 Then
        ' Msg8=Estas demasiado lejos.
        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).invent.EquippedWorkingToolObjIndex = 0 Then
        'Msg1071= Debes tener equipado un martillo de herrero para trabajar con el yunque.
        Call WriteLocaleMsg(UserIndex, 1071, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If ObjData(UserList(UserIndex).invent.EquippedWorkingToolObjIndex).Subtipo <> 7 Then
        'Msg1072= La herramienta que tienes no es la correcta, necesitas un martillo de herrero para poder trabajar.
        Call WriteLocaleMsg(UserIndex, 1072, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Call EnivarArmasConstruibles(UserIndex)
    Call EnivarArmadurasConstruibles(UserIndex)
    Call SendCraftableElementRunes(UserIndex)
    Call WriteShowBlacksmithForm(UserIndex)
    Exit Sub
AccionParaYunque_Err:
    Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaYunque", Erl)
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte, ByVal UserIndex As Integer, Optional ByVal SinDistancia As Boolean)
    On Error GoTo Handler
    Dim puerta As t_ObjData 'ver ReyarB
    If Distance(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, x, y) > 2 And Not SinDistancia Then
        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    puerta = ObjData(MapData(Map, x, y).ObjInfo.ObjIndex)
    If puerta.Llave = 1 And Not SinDistancia Then
        If puerta.GrhIndex = 11445 Or puerta.GrhIndex = 11444 Or puerta.GrhIndex = 59878 Or puerta.GrhIndex = 59877 Then
            'Msg1073= Al parecer, alguien cerró esta puerta. Debe haber algún interruptor por algún lado...
            Call WriteLocaleMsg(UserIndex, 1073, e_FontTypeNames.FONTTYPE_INFO)
        Else
            'Msg1074= La puerta esta cerrada con llave.
            Call WriteLocaleMsg(UserIndex, 1074, e_FontTypeNames.FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    If puerta.Cerrada = 1 Then 'Abre la puerta
        MapData(Map, x, y).ObjInfo.ObjIndex = puerta.IndexAbierta
        Call BloquearPuerta(Map, x, y, False)
        If puerta.GrhIndex = 11445 Or puerta.GrhIndex = 11444 Or puerta.GrhIndex = 59878 Or puerta.GrhIndex = 59877 Then
            'Msg1075= Has abierto la compuerta del ducto.
            Call WriteLocaleMsg(UserIndex, 1075, e_FontTypeNames.FONTTYPE_INFO)
        End If
    Else 'Cierra puerta
        MapData(Map, x, y).ObjInfo.ObjIndex = puerta.IndexCerrada
        Call BloquearPuerta(Map, x, y, True)
    End If
    If ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).Subtipo = 1 Then
        Call AccionParaPuerta(Map, x - 3, y + 1, UserIndex, True)
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
    Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaPuerta", Erl)
End Sub

Sub AccionParaPuertaNpc(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte, ByVal NpcIndex As Integer)
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
    Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaPuertaNpc", Erl)
End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
    On Error GoTo Handler
    Dim MiObj As t_Obj
    'If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
    If Len(ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).texto) > 0 Then
        Call WriteShowSignal(UserIndex, MapData(Map, x, y).ObjInfo.ObjIndex)
    Else
        Call WriteShowPapiro(UserIndex)
    End If
    'End If
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaCartel", Erl)
End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
    On Error GoTo Handler
    Dim Suerte As Byte
    Dim exito  As Byte
    Dim raise  As Integer
    Dim pos    As t_WorldPos
    pos.Map = Map
    pos.x = x
    pos.y = y
    With UserList(UserIndex)
        If Distancia(pos, .pos) > 2 Then
            ' Msg8=Estas demasiado lejos.
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapInfo(Map).lluvia And Lloviendo Then
            'Msg1076= Esta lloviendo, no podés encender una fogata aquí.
            Call WriteLocaleMsg(UserIndex, 1076, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapData(Map, x, y).trigger = e_Trigger.ZonaSegura Or MapInfo(Map).Seguro = 1 Then
            'Msg1077= En zona segura no podés hacer fogatas.
            Call WriteLocaleMsg(UserIndex, 1077, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapData(Map, x - 1, y).ObjInfo.ObjIndex = FOGATA Or MapData(Map, x + 1, y).ObjInfo.ObjIndex = FOGATA Or MapData(Map, x, y - 1).ObjInfo.ObjIndex = FOGATA Or MapData( _
                Map, x, y + 1).ObjInfo.ObjIndex = FOGATA Then
            'Msg1078= Debes alejarte un poco de la otra fogata.
            Call WriteLocaleMsg(UserIndex, 1078, e_FontTypeNames.FONTTYPE_INFO)
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
                Call WriteLocaleMsg(UserIndex, 1079, e_FontTypeNames.FONTTYPE_INFO)
                Call MakeObj(obj, Map, x, y)
            Else
                'Msg1080= La ley impide realizar fogatas en las ciudades.
                Call WriteLocaleMsg(UserIndex, 1080, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            'Msg1081= No has podido hacer fuego.
            Call WriteLocaleMsg(UserIndex, 1081, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Call SubirSkill(UserIndex, Supervivencia)
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaRamita", Erl)
End Sub
