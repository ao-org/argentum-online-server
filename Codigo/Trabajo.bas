Attribute VB_Name = "Trabajo"
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
Public Const GOLD_OBJ_INDEX As Long = 12
Public Const FISHING_NET_FX As Long = 12
Public Const NET_INMO_DURATION = 10

Function ExpectObjectTypeAt(ByVal objectType As Integer, ByVal Map As Integer, ByVal MapX As Byte, ByVal MapY As Byte) As Boolean
    Dim ObjIndex As Integer
    ObjIndex = MapData(Map, MapX, MapY).ObjInfo.ObjIndex
    If ObjIndex = 0 Then
        ExpectObjectTypeAt = False
        Exit Function
    End If
    ExpectObjectTypeAt = ObjData(ObjIndex).OBJType = objectType
End Function

Function IsUserAtPos(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte) As Boolean
    IsUserAtPos = MapData(Map, x, y).UserIndex > 0
End Function

Function IsNpcAtPos(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte)
    IsNpcAtPos = MapData(Map, x, y).NpcIndex > 0
End Function

Sub HandleFishingNet(ByVal UserIndex As Integer)
    On Error GoTo HandleFishingNet_Err:
    With UserList(UserIndex)
        If (MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).Blocked And FLAG_AGUA) <> 0 Or MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).trigger = _
                e_Trigger.PESCAINVALIDA Then
            If Abs(.pos.x - .Trabajo.Target_X) + Abs(.pos.y - .Trabajo.Target_Y) > 8 Then
                Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                Call WriteWorkRequestTarget(UserIndex, 0)
                Exit Sub
            End If
            If MapInfo(UserList(UserIndex).pos.Map).Seguro = 1 Then
                ' Msg593=Esta prohibida la pesca masiva en las ciudades.
                Call WriteLocaleMsg(UserIndex, 593, e_FontTypeNames.FONTTYPE_INFO)
                Call WriteWorkRequestTarget(UserIndex, 0)
                Exit Sub
            End If
            If UserList(UserIndex).flags.Navegando = 0 Then
                ' Msg594=Necesitas estar sobre tu barca para utilizar la red de pesca.
                Call WriteLocaleMsg(UserIndex, 594, e_FontTypeNames.FONTTYPE_INFO)
                Call WriteWorkRequestTarget(UserIndex, 0)
                Exit Sub
            End If
            If SvrConfig.GetValue("FISHING_POOL_ID") <> MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.ObjIndex Then
                ' Msg595=Para pescar con red deberás buscar un área de pesca.
                Call WriteLocaleMsg(UserIndex, 595, e_FontTypeNames.FONTTYPE_INFO)
                Call WriteWorkRequestTarget(UserIndex, 0)
                Exit Sub
            End If
            If MapInfo(.pos.Map).zone = "DUNGEON" Then
                Call WriteLocaleMsg(UserIndex, 596, e_FontTypeNames.FONTTYPE_INFO)
                Call WriteWorkRequestTarget(UserIndex, 0)
                Exit Sub
            End If
            Call DoPescar(UserIndex, True)
        Else
            ' Msg596=Zona de pesca no Autorizada. Busca otro lugar para hacerlo.
            Call WriteLocaleMsg(UserIndex, 596, e_FontTypeNames.FONTTYPE_INFO)
            Call WriteWorkRequestTarget(UserIndex, 0)
        End If
    End With
    Exit Sub
HandleFishingNet_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.HandleFishingNet", Erl)
End Sub

Public Sub Trabajar(ByVal UserIndex As Integer, ByVal Skill As e_Skill)
    Dim DummyInt As Integer
    With UserList(UserIndex)
        Select Case Skill
            Case e_Skill.Pescar
                If .invent.EquippedWorkingToolObjIndex = 0 Then Exit Sub
                If ObjData(.invent.EquippedWorkingToolObjIndex).OBJType <> e_OBJType.otWorkingTools Then Exit Sub
                Select Case ObjData(.invent.EquippedWorkingToolObjIndex).Subtipo
                    Case e_ToolsSubtype.eFishingRod
                        If (MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).Blocked And FLAG_AGUA) <> 0 And Not MapData(.pos.Map, .pos.x, .pos.y).trigger = _
                                e_Trigger.PESCAINVALIDA Then
                            Dim isStandingOnWater As Boolean
                            Dim isAdjacentToWater As Boolean

                            isStandingOnWater = (MapData(.pos.Map, .pos.x, .pos.y).Blocked And FLAG_AGUA) <> 0
                            isAdjacentToWater = (MapData(.pos.Map, .pos.x + 1, .pos.y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.pos.Map, .pos.x, .pos.y + 1).Blocked And FLAG_AGUA) <> 0 Or (MapData( _
                                    .pos.Map, .pos.x - 1, .pos.y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.pos.Map, .pos.x, .pos.y - 1).Blocked And FLAG_AGUA) <> 0

                            If isStandingOnWater Then
                                Call WriteLocaleMsg(UserIndex, "1436", e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteMacroTrabajoToggle(UserIndex, False)
                            ElseIf isAdjacentToWater Then
                                .flags.PescandoEspecial = False
                                If UserList(UserIndex).flags.Navegando = 0 Then
                                    If MapInfo(.pos.Map).zone = "DUNGEON" Then
                                        Call WriteLocaleMsg(UserIndex, 596, e_FontTypeNames.FONTTYPE_INFO)
                                        Call WriteMacroTrabajoToggle(UserIndex, False)
                                    Else
                                        Call DoPescar(UserIndex, False)
                                    End If
                                Else
                                    Call WriteLocaleMsg(UserIndex, 1436, e_FontTypeNames.FONTTYPE_INFO)
                                    Call WriteMacroTrabajoToggle(UserIndex, False)
                                End If
                            Else
                                'Msg1021= Acércate a la costa para pescar.
                                Call WriteLocaleMsg(UserIndex, 1021, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteMacroTrabajoToggle(UserIndex, False)
                            End If
                        Else
                            ' Msg596=Zona de pesca no Autorizada. Busca otro lugar para hacerlo.
                            Call WriteLocaleMsg(UserIndex, 596, e_FontTypeNames.FONTTYPE_INFO)
                            Call WriteMacroTrabajoToggle(UserIndex, False)
                        End If
                    Case e_ToolsSubtype.eFishingNet
                        Call HandleFishingNet(UserIndex)
                End Select
            Case e_Skill.Carpinteria
                'Veo cual es la cantidad máxima que puede construir de una
                Dim cantidad_maxima As Long
                If UserList(UserIndex).clase = e_Class.Trabajador Then
                    cantidad_maxima = UserList(UserIndex).Stats.UserSkills(e_Skill.Carpinteria) / 10
                    If cantidad_maxima = 0 Then cantidad_maxima = 1
                Else
                    'Si no hace de a 1
                    cantidad_maxima = 1
                End If
                Call CarpinteroConstruirItem(UserIndex, UserList(UserIndex).Trabajo.Item, UserList(UserIndex).Trabajo.Cantidad, cantidad_maxima)
            Case e_Skill.Mineria
                If .invent.EquippedWorkingToolObjIndex = 0 Then Exit Sub
                If ObjData(.invent.EquippedWorkingToolObjIndex).OBJType <> e_OBJType.otWorkingTools Then Exit Sub
                'Check interval
                If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub
                Select Case ObjData(.invent.EquippedWorkingToolObjIndex).Subtipo
                    Case 8  ' Herramientas de Mineria - Piquete
                        'Target whatever is in the tile
                        Call LookatTile(UserIndex, .pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y)
                        DummyInt = MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.ObjIndex
                        If DummyInt > 0 Then
                            'Check distance
                            If Abs(.pos.x - .Trabajo.Target_X) + Abs(.pos.y - .Trabajo.Target_Y) > 2 Then
                                Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                                'Msg8=Estís demasiado lejos.
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                            '¡Hay un yacimiento donde clickeo?
                            If ObjData(DummyInt).OBJType = e_OBJType.otOreDeposit Then
                                ' Si el Yacimiento requiere herramienta `Dorada` y la herramienta no lo es, o vice versa.
                                ' Se usa para el yacimiento de Oro.
                                If ObjData(DummyInt).Dorada <> ObjData(.invent.EquippedWorkingToolObjIndex).Dorada Or ObjData(DummyInt).Blodium <> ObjData( _
                                        .invent.EquippedWorkingToolObjIndex).Blodium Then
                                    If ObjData(DummyInt).Blodium <> ObjData(.invent.EquippedWorkingToolObjIndex).Blodium Then
                                        ' Msg597=El pico minero especial solo puede extraer minerales del yacimiento de Blodium.
                                        Call WriteLocaleMsg(UserIndex, 597, e_FontTypeNames.FONTTYPE_INFO)
                                        Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub
                                    Else
                                        'Msg1022= El pico dorado solo puede extraer minerales del yacimiento de Oro.
                                        Call WriteLocaleMsg(UserIndex, 1022, e_FontTypeNames.FONTTYPE_INFO)
                                        Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub
                                    End If
                                End If
                                If MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount <= 0 Then
                                    ' Msg598=Este yacimiento no tiene más minerales para entregar.
                                    Call WriteLocaleMsg(UserIndex, 598, e_FontTypeNames.FONTTYPE_INFO)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Call WriteMacroTrabajoToggle(UserIndex, False)
                                    Exit Sub
                                End If
                                Call DoMineria(UserIndex, .Trabajo.Target_X, .Trabajo.Target_Y, ObjData(.invent.EquippedWorkingToolObjIndex).Dorada = 1)
                            Else
                                ' Msg599=Ahí no hay ningún yacimiento.
                                Call WriteLocaleMsg(UserIndex, 599, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                            End If
                        Else
                            ' Msg599=Ahí no hay ningún yacimiento.
                            Call WriteLocaleMsg(UserIndex, 599, e_FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)
                        End If
                End Select
            Case e_Skill.Talar
                If .invent.EquippedWorkingToolObjIndex = 0 Then Exit Sub
                If ObjData(.invent.EquippedWorkingToolObjIndex).OBJType <> e_OBJType.otWorkingTools Then Exit Sub
                'Check interval
                If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub
                Select Case ObjData(.invent.EquippedWorkingToolObjIndex).Subtipo
                    Case 6      ' Herramientas de Carpinteria - Hacha
                        DummyInt = MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.ObjIndex
                        If DummyInt > 0 Then
                            If Abs(.pos.x - .Trabajo.Target_X) + Abs(.pos.y - .Trabajo.Target_Y) > 1 Then
                                Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                                'Msg8=Estas demasiado lejos.
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                            If .pos.x = .Trabajo.Target_X And .pos.y = .Trabajo.Target_Y Then
                                ' Msg600=No podés talar desde allí.
                                Call WriteLocaleMsg(UserIndex, 600, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                            If ObjData(DummyInt).Elfico <> ObjData(.invent.EquippedWorkingToolObjIndex).Elfico Then
                                ' Msg601=Sólo puedes talar árboles elficos con un hacha élfica.
                                Call WriteLocaleMsg(UserIndex, 601, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                            If ObjData(DummyInt).Pino <> ObjData(.invent.EquippedWorkingToolObjIndex).Pino Then
                                ' Msg602=Sólo puedes talar árboles de pino nudoso con un hacha de pino.
                                Call WriteLocaleMsg(UserIndex, 602, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                            If MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount <= 0 Then
                                ' Msg603=El árbol ya no te puede entregar más leña.
                                Call WriteLocaleMsg(UserIndex, 603, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Call WriteMacroTrabajoToggle(UserIndex, False)
                                Exit Sub
                            End If
                            '¡Hay un arbol donde clickeo?
                            If ObjData(DummyInt).OBJType = e_OBJType.otTrees Then
                                Call DoTalar(UserIndex, .Trabajo.Target_X, .Trabajo.Target_Y, ObjData(.invent.EquippedWorkingToolObjIndex).Dorada = 1)
                            End If
                        Else
                            ' Msg604=No hay ningún árbol ahí.
                            Call WriteLocaleMsg(UserIndex, 604, e_FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)
                            If UserList(UserIndex).Counters.Trabajando > 1 Then
                                Call WriteMacroTrabajoToggle(UserIndex, False)
                            End If
                        End If
                End Select
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                'Check interval
                If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = e_OBJType.otForge Then
                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > UserList(UserIndex).CurrentInventorySlots Then
                            Exit Sub
                        End If
                        ''chequeamos que no se zarpe duplicando oro
                        If .invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .invent.Object(.flags.TargetObjInvSlot).amount = 0 Then
                                ' Msg605=No tienes más minerales
                                Call WriteLocaleMsg(UserIndex, 605, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                            ''FUISTE
                            Call WriteShowMessageBox(UserIndex, 1774, vbNullString) 'Msg1774=Has sido expulsado por el sistema anti cheats.
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        Call FundirMineral(UserIndex)
                    Else
                        ' Msg606=Ahí no hay ninguna fragua.
                        Call WriteLocaleMsg(UserIndex, 606, e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        If UserList(UserIndex).Counters.Trabajando > 1 Then
                            Call WriteMacroTrabajoToggle(UserIndex, False)
                        End If
                    End If
                Else
                    ' Msg606=Ahí no hay ninguna fragua.
                    Call WriteLocaleMsg(UserIndex, 606, e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(UserIndex, 0)
                    If UserList(UserIndex).Counters.Trabajando > 1 Then
                        Call WriteMacroTrabajoToggle(UserIndex, False)
                    End If
                End If
        End Select
    End With
End Sub

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
    '********************************************************
    'Autor: Nacho (Integer)
    'Last Modif: 28/01/2007
    'Chequea si ya debe mostrarse
    'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
    '********************************************************
    On Error GoTo DoPermanecerOculto_Err
    With UserList(UserIndex)
        Dim velocidadOcultarse As Integer
        velocidadOcultarse = 1
        'HarThaoS: Si tiene armadura de cazador, dependiendo skills vemos cuanto tiempo se oculta
        If .clase = e_Class.Hunter Then
            If TieneArmaduraCazador(UserIndex) Then
                Select Case .Stats.UserSkills(e_Skill.Ocultarse)
                    Case Is = 100
                        Exit Sub
                    Case Is < 100
                        velocidadOcultarse = RandomNumber(0, 1)
                End Select
            End If
        End If
        .Counters.TiempoOculto = .Counters.TiempoOculto - velocidadOcultarse
        If .Counters.TiempoOculto <= 0 Then
            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            If .flags.Navegando = 1 Then
                If .clase = e_Class.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call EquiparBarco(UserIndex)
                    ' Msg592=¡Has recuperado tu apariencia normal!
                    Call WriteLocaleMsg(UserIndex, 592, e_FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart, NoBackPack)
                    Call RefreshCharStatus(UserIndex)
                End If
            Else
                If .flags.invisible = 0 And .flags.AdminInvisible = 0 Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                    'Msg1023= ¡Has vuelto a ser visible!
                    Call WriteLocaleMsg(UserIndex, 1023, e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
    Exit Sub
DoPermanecerOculto_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.DoPermanecerOculto", Erl)
End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
    'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
    'Modifique la fórmula y ahora anda bien.
    On Error GoTo ErrHandler
    Dim Suerte As Double
    Dim res    As Integer
    Dim Skill  As Integer
    With UserList(UserIndex)
        If .flags.Navegando = 1 And .clase <> e_Class.Pirat Then
            Call WriteLocaleMsg(UserIndex, 56, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If GlobalFrameTime - .Counters.LastAttackTime < HideAfterHitTime Then
            Exit Sub
        End If
        Skill = .Stats.UserSkills(e_Skill.Ocultarse)
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        res = RandomNumber(1, 100)
        If res <= Suerte Then
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            Select Case .clase
                Case e_Class.Bandit, e_Class.Thief
                    .Counters.TiempoOculto = RandomNumber(Int(Suerte / 2.5), Int(Suerte / 2))
                Case e_Class.Hunter
                    .Counters.TiempoOculto = Int(Suerte / 2)
                Case Else
                    .Counters.TiempoOculto = Int(Suerte / 3)
            End Select
            If .flags.Navegando = 1 Then
                If .clase = e_Class.Pirat Then
                    .Char.body = iFragataFantasmal
                    .flags.Oculto = 1
                    .Counters.TiempoOculto = IntervaloOculto
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart, NoBackPack)
                    'Msg1024= ¡Te has camuflado como barco fantasma!
                    Call WriteLocaleMsg(UserIndex, 1024, e_FontTypeNames.FONTTYPE_INFO)
                    Call RefreshCharStatus(UserIndex)
                End If
            Else
                UserList(UserIndex).Counters.timeFx = 3
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                'Msg55=¡Te has escondido entre las sombras!
                Call WriteLocaleMsg(UserIndex, 55, e_FontTypeNames.FONTTYPE_INFO)
            End If
            Call SubirSkill(UserIndex, Ocultarse)
        Else
            If Not .flags.UltimoMensaje = 4 Then
                'Msg57=¡No has logrado esconderte!
                Call WriteLocaleMsg(UserIndex, 57, e_FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4
            End If
        End If
        .Counters.Ocultando = .Counters.Ocultando + 1
    End With
    Exit Sub
ErrHandler:
    Call LogError("Error en Sub DoOcultarse")
End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As t_ObjData, ByVal Slot As Integer)
    On Error GoTo DoNavega_Err
    With UserList(UserIndex)
        If .invent.EquippedShipObjIndex <> .invent.Object(Slot).ObjIndex Then
            If Not EsGM(UserIndex) Then
                Select Case Barco.Subtipo
                    Case 2  'Galera
                        If .clase <> e_Class.Trabajador And .clase <> e_Class.Pirat Then
                            'Msg1025= ¡Solo Piratas y trabajadores pueden usar galera!
                            Call WriteLocaleMsg(UserIndex, 1025, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    Case 3  'Galeón
                        If .clase <> e_Class.Pirat Then
                            'Msg1026= Solo los Piratas pueden usar Galeón!!
                            Call WriteLocaleMsg(UserIndex, 1026, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                End Select
            End If
            Dim SkillNecesario As Byte
            SkillNecesario = IIf(.clase = e_Class.Trabajador Or .clase = e_Class.Pirat, Barco.MinSkill \ 2, Barco.MinSkill)
            ' Tiene el skill necesario?
            If .Stats.UserSkills(e_Skill.Navegacion) < SkillNecesario Then
                Call WriteLocaleMsg(UserIndex, 1448, e_FontTypeNames.FONTTYPE_INFO, SkillNecesario & "¬" & IIf(Barco.Subtipo = 0, "traje", "barco"))  ' Msg1448=Necesitas al menos ¬1 puntos en navegación para poder usar este ¬2
                Exit Sub
            End If
            If .invent.EquippedShipObjIndex = 0 Then
                Call WriteNavigateToggle(UserIndex, True)
                .flags.Navegando = 1
                Call TargetUpdateTerrain(.EffectOverTime)
            End If
            .invent.EquippedShipObjIndex = .invent.Object(Slot).ObjIndex
            .invent.EquippedShipSlot = Slot
            If .flags.Montado > 0 Then
                Call DoMontar(UserIndex, ObjData(.invent.EquippedSaddleObjIndex), .invent.EquippedSaddleSlot)
            End If
            If .flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
                'Msg1027= Pierdes el efecto del mimetismo.
                Call WriteLocaleMsg(UserIndex, 1027, e_FontTypeNames.FONTTYPE_INFO)
                .Counters.Mimetismo = 0
                .flags.Mimetizado = e_EstadoMimetismo.Desactivado
                Call RefreshCharStatus(UserIndex)
            End If
            Call EquiparBarco(UserIndex)
        Else
            Call WriteNadarToggle(UserIndex, False)
            .flags.Navegando = 0
            Call WriteNavigateToggle(UserIndex, False)
            Call TargetUpdateTerrain(.EffectOverTime)
            .invent.EquippedShipObjIndex = 0
            .invent.EquippedShipSlot = 0
            If .flags.Muerto = 0 Then
                .Char.head = .OrigChar.head
                If .invent.EquippedArmorObjIndex > 0 Then
                    If .invent.EquippedArmorObjIndex > 0 And .Invent_Skins.ObjIndexArmourEquipped > 0 And .Invent_Skins.SlotBoatEquipped > 0 Then
                        If .Invent_Skins.Object(.Invent_Skins.SlotBoatEquipped).Equipped Then
                            .Char.body = ObtenerRopaje(UserIndex, ObjData(.Invent_Skins.ObjIndexArmourEquipped))
                        Else
                            .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.EquippedArmorObjIndex))
                        End If
                    Else
                        .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.EquippedArmorObjIndex))
                    End If
                Else
                    Call SetNakedBody(UserList(UserIndex))
                End If
                If .invent.EquippedShieldObjIndex > 0 Then .Char.ShieldAnim = ObjData(.invent.EquippedShieldObjIndex).ShieldAnim
                If .invent.EquippedWeaponObjIndex > 0 Then .Char.WeaponAnim = ObjData(.invent.EquippedWeaponObjIndex).WeaponAnim
                If .invent.EquippedWorkingToolObjIndex > 0 Then .Char.WeaponAnim = ObjData(.invent.EquippedWorkingToolObjIndex).WeaponAnim
                If .invent.EquippedHelmetObjIndex > 0 Then .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                If .invent.EquippedAmuletAccesoryObjIndex > 0 Then
                    If ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje > 0 Then .Char.CartAnim = ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje
                End If
            Else
                .Char.body = iCuerpoMuerto
                .Char.head = 0
                Call ClearClothes(.Char)
            End If
            Call ActualizarVelocidadDeUsuario(UserIndex)
        End If
        ' Volver visible
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 And .flags.invisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            'MSG307=Has vuelto a ser visible.
            Call WriteLocaleMsg(UserIndex, 307, e_FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        End If
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundEffects.BARCA_SOUND, .pos.x, .pos.y))
    End With
    Exit Sub
DoNavega_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.DoNavega", Erl)
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
    On Error GoTo FundirMineral_Err
    If UserList(UserIndex).clase <> e_Class.Trabajador Then
        ' Msg607=Tu clase no tiene el conocimiento suficiente para trabajar este mineral.
        Call WriteLocaleMsg(UserIndex, 607, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios) Then
        Exit Sub
    End If
    If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
        Dim SkillRequerido As Integer
        SkillRequerido = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill
        If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = e_OBJType.otMinerals And UserList(UserIndex).Stats.UserSkills(e_Skill.Mineria) >= SkillRequerido Then
            Call DoLingotes(UserIndex)
        ElseIf SkillRequerido > 100 Then
            ' Msg608=Los mortales no pueden fundir este mineral.
            Call WriteLocaleMsg(UserIndex, 608, e_FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteLocaleMsg(UserIndex, 1449, e_FontTypeNames.FONTTYPE_INFO)  ' Msg1449=No tenés conocimientos de minería suficientes para trabajar este mineral. Necesitas ¬1 puntos en minería.
        End If
    End If
    Exit Sub
FundirMineral_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.FundirMineral", Erl)
End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer, Optional ByVal ElementalTags As Long = e_ElementalTags.Normal) As Boolean
    On Error GoTo TieneObjetos_Err
    Dim i     As Long
    Dim total As Long
    If (ItemIndex = GOLD_OBJ_INDEX) Then
        total = UserList(UserIndex).Stats.GLD
    End If
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        If UserList(UserIndex).invent.Object(i).ObjIndex = ItemIndex And UserList(UserIndex).invent.Object(i).ElementalTags = ElementalTags Then
            total = total + UserList(UserIndex).invent.Object(i).amount
        End If
    Next i
    If cant <= total Then
        TieneObjetos = True
        Exit Function
    End If
    Exit Function
TieneObjetos_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.TieneObjetos", Erl)
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer, Optional ByVal ElementalTags As Long = e_ElementalTags.Normal) As Boolean
    On Error GoTo QuitarObjetos_Err
    With UserList(UserIndex)
        Dim i As Long
        For i = 1 To .CurrentInventorySlots
            If .invent.Object(i).ObjIndex = ItemIndex And .invent.Object(i).ElementalTags = ElementalTags Then
                .invent.Object(i).amount = .invent.Object(i).amount - cant
                If .invent.Object(i).amount <= 0 Then
                    If .invent.Object(i).Equipped Then
                        Call Desequipar(UserIndex, i)
                    End If
                    cant = Abs(.invent.Object(i).amount)
                    .invent.Object(i).amount = 0
                    .invent.Object(i).ObjIndex = 0
                    .invent.Object(i).ElementalTags = 0
                Else
                    cant = 0
                End If
                Call UpdateUserInv(False, UserIndex, i)
                If cant = 0 Then
                    QuitarObjetos = True
                    Exit Function
                End If
            End If
        Next i
        If (ItemIndex = GOLD_OBJ_INDEX And cant > 0) Then
            .Stats.GLD = .Stats.GLD - cant
            If (.Stats.GLD < 0) Then
                cant = Abs(.Stats.GLD)
                .Stats.GLD = 0
            End If
            Call WriteUpdateGold(UserIndex)
            If (cant = 0) Then
                QuitarObjetos = True
                Exit Function
            End If
        End If
    End With
    Exit Function
QuitarObjetos_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.QuitarObjetos", Erl)
End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    On Error GoTo HerreroQuitarMateriales_Err
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(e_Minerales.LingoteDeHierro, ObjData(ItemIndex).LingH, UserIndex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(e_Minerales.LingoteDePlata, ObjData(ItemIndex).LingP, UserIndex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(e_Minerales.LingoteDeOro, ObjData(ItemIndex).LingO, UserIndex)
    If ObjData(ItemIndex).Coal > 0 Then Call QuitarObjetos(e_Minerales.Coal, ObjData(ItemIndex).Coal, UserIndex)
    If ObjData(ItemIndex).Blodium > 0 Then Call QuitarObjetos(e_Minerales.Blodium, ObjData(ItemIndex).Blodium, UserIndex)
    If ObjData(ItemIndex).FireEssence > 0 Then Call QuitarObjetos(e_Minerales.FireEssence, ObjData(ItemIndex).FireEssence, UserIndex)
    If ObjData(ItemIndex).WaterEssence > 0 Then Call QuitarObjetos(e_Minerales.WaterEssence, ObjData(ItemIndex).WaterEssence, UserIndex)
    If ObjData(ItemIndex).EarthEssence > 0 Then Call QuitarObjetos(e_Minerales.EarthEssence, ObjData(ItemIndex).EarthEssence, UserIndex)
    If ObjData(ItemIndex).WindEssence > 0 Then Call QuitarObjetos(e_Minerales.WindEssence, ObjData(ItemIndex).WindEssence, UserIndex)
    Exit Sub
HerreroQuitarMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.HerreroQuitarMateriales", Erl)
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer, ByVal CantidadElfica As Integer, ByVal CantidadPino As Integer)
    On Error GoTo CarpinteroQuitarMateriales_Err
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Wood, Cantidad, UserIndex)
    If ObjData(ItemIndex).MaderaElfica > 0 Then Call QuitarObjetos(ElvenWood, CantidadElfica, UserIndex)
    If ObjData(ItemIndex).MaderaPino > 0 Then Call QuitarObjetos(PinoWood, CantidadPino, UserIndex)
    Exit Sub
CarpinteroQuitarMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.CarpinteroQuitarMateriales", Erl)
End Sub

Sub AlquimistaQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    On Error GoTo AlquimistaQuitarMateriales_Err
    If ObjData(ItemIndex).Raices > 0 Then Call QuitarObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex)
    If ObjData(ItemIndex).Botella > 0 Then Call QuitarObjetos(Botella, ObjData(ItemIndex).Botella, UserIndex)
    If ObjData(ItemIndex).Cuchara > 0 Then Call QuitarObjetos(Cuchara, ObjData(ItemIndex).Cuchara, UserIndex)
    If ObjData(ItemIndex).Mortero > 0 Then Call QuitarObjetos(Mortero, ObjData(ItemIndex).Mortero, UserIndex)
    If ObjData(ItemIndex).FrascoAlq > 0 Then Call QuitarObjetos(FrascoAlq, ObjData(ItemIndex).FrascoAlq, UserIndex)
    If ObjData(ItemIndex).FrascoElixir > 0 Then Call QuitarObjetos(FrascoElixir, ObjData(ItemIndex).FrascoElixir, UserIndex)
    If ObjData(ItemIndex).Dosificador > 0 Then Call QuitarObjetos(Dosificador, ObjData(ItemIndex).Dosificador, UserIndex)
    If ObjData(ItemIndex).Orquidea > 0 Then Call QuitarObjetos(Orquidea, ObjData(ItemIndex).Orquidea, UserIndex)
    If ObjData(ItemIndex).Carmesi > 0 Then Call QuitarObjetos(Carmesi, ObjData(ItemIndex).Carmesi, UserIndex)
    If ObjData(ItemIndex).HongoDeLuz > 0 Then Call QuitarObjetos(HongoDeLuz, ObjData(ItemIndex).HongoDeLuz, UserIndex)
    If ObjData(ItemIndex).Esporas > 0 Then Call QuitarObjetos(Esporas, ObjData(ItemIndex).Esporas, UserIndex)
    If ObjData(ItemIndex).Tuna > 0 Then Call QuitarObjetos(Tuna, ObjData(ItemIndex).Tuna, UserIndex)
    If ObjData(ItemIndex).Cala > 0 Then Call QuitarObjetos(Cala, ObjData(ItemIndex).Cala, UserIndex)
    If ObjData(ItemIndex).ColaDeZorro > 0 Then Call QuitarObjetos(ColaDeZorro, ObjData(ItemIndex).ColaDeZorro, UserIndex)
    If ObjData(ItemIndex).FlorOceano > 0 Then Call QuitarObjetos(FlorOceano, ObjData(ItemIndex).FlorOceano, UserIndex)
    If ObjData(ItemIndex).FlorRoja > 0 Then Call QuitarObjetos(FlorRoja, ObjData(ItemIndex).FlorRoja, UserIndex)
    If ObjData(ItemIndex).Hierva > 0 Then Call QuitarObjetos(Hierva, ObjData(ItemIndex).Hierva, UserIndex)
    If ObjData(ItemIndex).HojasDeRin > 0 Then Call QuitarObjetos(HojasDeRin, ObjData(ItemIndex).HojasDeRin, UserIndex)
    If ObjData(ItemIndex).HojasRojas > 0 Then Call QuitarObjetos(HojasRojas, ObjData(ItemIndex).HojasRojas, UserIndex)
    If ObjData(ItemIndex).SemillasPros > 0 Then Call QuitarObjetos(SemillasPros, ObjData(ItemIndex).SemillasPros, UserIndex)
    If ObjData(ItemIndex).Pimiento > 0 Then Call QuitarObjetos(Pimiento, ObjData(ItemIndex).Pimiento, UserIndex)
    Exit Sub
AlquimistaQuitarMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.AlquimistaQuitarMateriales", Erl)
End Sub

Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    On Error GoTo SastreQuitarMateriales_Err
    If ObjData(ItemIndex).PielLobo > 0 Then Call QuitarObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex)
    If ObjData(ItemIndex).PielOsoPardo > 0 Then Call QuitarObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex)
    If ObjData(ItemIndex).PielOsoPolaR > 0 Then Call QuitarObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex)
    If ObjData(ItemIndex).PielLoboNegro > 0 Then Call QuitarObjetos(PielLoboNegro, ObjData(ItemIndex).PielLoboNegro, UserIndex)
    If ObjData(ItemIndex).PielTigre > 0 Then Call QuitarObjetos(PielTigre, ObjData(ItemIndex).PielTigre, UserIndex)
    If ObjData(ItemIndex).PielTigreBengala > 0 Then Call QuitarObjetos(PielTigreBengala, ObjData(ItemIndex).PielTigreBengala, UserIndex)
    Exit Sub
SastreQuitarMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.SastreQuitarMateriales", Erl)
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Long) As Boolean
    On Error GoTo CarpinteroTieneMateriales_Err
    If ObjData(ItemIndex).Madera > 0 Then
        If Not TieneObjetos(Wood, ObjData(ItemIndex).Madera * Cantidad, UserIndex) Then
            ' Msg609=No tenés suficiente madera.
            Call WriteLocaleMsg(UserIndex, 609, e_FontTypeNames.FONTTYPE_INFO)
            CarpinteroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).MaderaElfica > 0 Then
        If Not TieneObjetos(ElvenWood, ObjData(ItemIndex).MaderaElfica * Cantidad, UserIndex) Then
            ' Msg610=No tenés suficiente madera élfica.
            Call WriteLocaleMsg(UserIndex, 610, e_FontTypeNames.FONTTYPE_INFO)
            CarpinteroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).MaderaPino > 0 Then
        If Not TieneObjetos(PinoWood, ObjData(ItemIndex).MaderaPino * Cantidad, UserIndex) Then
            ' Msg611=No tenés suficiente madera de pino nudoso.
            Call WriteLocaleMsg(UserIndex, 611, e_FontTypeNames.FONTTYPE_INFO)
            CarpinteroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    CarpinteroTieneMateriales = True
    Exit Function
CarpinteroTieneMateriales_Err:
    Call TraceError(Err.Number, Err.Description + " UI:" + UserIndex + " Item: " + ItemIndex, "Trabajo.CarpinteroTieneMateriales", Erl)
End Function

Function AlquimistaTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    On Error GoTo AlquimistaTieneMateriales_Err
    If ObjData(ItemIndex).Raices > 0 Then
        If Not TieneObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex) Then
            ' Msg612=No tenés suficientes raíces.
            Call WriteLocaleMsg(UserIndex, 612, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Botella > 0 Then
        If Not TieneObjetos(Botella, ObjData(ItemIndex).Botella, UserIndex) Then
            ' Msg613=No tenés suficientes botellas.
            Call WriteLocaleMsg(UserIndex, 613, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Cuchara > 0 Then
        If Not TieneObjetos(Cuchara, ObjData(ItemIndex).Cuchara, UserIndex) Then
            ' Msg614=No tenés suficientes cucharas.
            Call WriteLocaleMsg(UserIndex, 614, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Mortero > 0 Then
        If Not TieneObjetos(Mortero, ObjData(ItemIndex).Mortero, UserIndex) Then
            ' Msg615=No tenés suficientes morteros.
            Call WriteLocaleMsg(UserIndex, 615, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).FrascoAlq > 0 Then
        If Not TieneObjetos(FrascoAlq, ObjData(ItemIndex).FrascoAlq, UserIndex) Then
            ' Msg616=No tenés suficientes frascos de alquimistas.
            Call WriteLocaleMsg(UserIndex, 616, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).FrascoElixir > 0 Then
        If Not TieneObjetos(FrascoElixir, ObjData(ItemIndex).FrascoElixir, UserIndex) Then
            ' Msg617=No tenés suficientes frascos de elixir superior.
            Call WriteLocaleMsg(UserIndex, 617, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Dosificador > 0 Then
        If Not TieneObjetos(Dosificador, ObjData(ItemIndex).Dosificador, UserIndex) Then
            ' Msg618=No tenés suficientes dosificadores.
            Call WriteLocaleMsg(UserIndex, 618, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Orquidea > 0 Then
        If Not TieneObjetos(Orquidea, ObjData(ItemIndex).Orquidea, UserIndex) Then
            ' Msg619=No tenés suficientes orquídeas silvestres.
            Call WriteLocaleMsg(UserIndex, 619, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Carmesi > 0 Then
        If Not TieneObjetos(Carmesi, ObjData(ItemIndex).Carmesi, UserIndex) Then
            ' Msg620=No tenés suficientes raíces carmesí.
            Call WriteLocaleMsg(UserIndex, 620, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).HongoDeLuz > 0 Then
        If Not TieneObjetos(HongoDeLuz, ObjData(ItemIndex).HongoDeLuz, UserIndex) Then
            ' Msg621=No tenés suficientes hongos de luz.
            Call WriteLocaleMsg(UserIndex, 621, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Esporas > 0 Then
        If Not TieneObjetos(Esporas, ObjData(ItemIndex).Esporas, UserIndex) Then
            ' Msg622=No tenés suficientes esporas silvestres.
            Call WriteLocaleMsg(UserIndex, 622, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Tuna > 0 Then
        If Not TieneObjetos(Tuna, ObjData(ItemIndex).Tuna, UserIndex) Then
            ' Msg623=No tenés suficientes tunas silvestres.
            Call WriteLocaleMsg(UserIndex, 623, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Cala > 0 Then
        If Not TieneObjetos(Cala, ObjData(ItemIndex).Cala, UserIndex) Then
            ' Msg624=No tenés suficientes calas venenosas.
            Call WriteLocaleMsg(UserIndex, 624, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).ColaDeZorro > 0 Then
        If Not TieneObjetos(ColaDeZorro, ObjData(ItemIndex).ColaDeZorro, UserIndex) Then
            ' Msg625=No tenés suficientes colas de zorro.
            Call WriteLocaleMsg(UserIndex, 625, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).FlorOceano > 0 Then
        If Not TieneObjetos(FlorOceano, ObjData(ItemIndex).FlorOceano, UserIndex) Then
            ' Msg626=No tenés suficientes flores del óceano.
            Call WriteLocaleMsg(UserIndex, 626, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).FlorRoja > 0 Then
        If Not TieneObjetos(FlorRoja, ObjData(ItemIndex).FlorRoja, UserIndex) Then
            ' Msg627=No tenés suficientes flores rojas.
            Call WriteLocaleMsg(UserIndex, 627, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Hierva > 0 Then
        If Not TieneObjetos(Hierva, ObjData(ItemIndex).Hierva, UserIndex) Then
            ' Msg628=No tenés suficientes hierbas de sangre.
            Call WriteLocaleMsg(UserIndex, 628, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).HojasDeRin > 0 Then
        If Not TieneObjetos(HojasDeRin, ObjData(ItemIndex).HojasDeRin, UserIndex) Then
            ' Msg629=No tenés suficientes hojas de rin.
            Call WriteLocaleMsg(UserIndex, 629, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).HojasRojas > 0 Then
        If Not TieneObjetos(HojasRojas, ObjData(ItemIndex).HojasRojas, UserIndex) Then
            ' Msg630=No tenés suficientes hojas rojas.
            Call WriteLocaleMsg(UserIndex, 630, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).SemillasPros > 0 Then
        If Not TieneObjetos(SemillasPros, ObjData(ItemIndex).SemillasPros, UserIndex) Then
            ' Msg631=No tenés suficientes semillas prósperas.
            Call WriteLocaleMsg(UserIndex, 631, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Pimiento > 0 Then
        If Not TieneObjetos(Pimiento, ObjData(ItemIndex).Pimiento, UserIndex) Then
            ' Msg632=No tenés suficientes Pimientos Muerte.
            Call WriteLocaleMsg(UserIndex, 632, e_FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    AlquimistaTieneMateriales = True
    Exit Function
AlquimistaTieneMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.AlquimistaTieneMateriales", Erl)
End Function

Function SastreTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    On Error GoTo SastreTieneMateriales_Err
    If ObjData(ItemIndex).PielLobo > 0 Then
        If Not TieneObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex) Then
            ' Msg633=No tenés suficientes pieles de lobo.
            Call WriteLocaleMsg(UserIndex, 633, e_FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).PielOsoPardo > 0 Then
        If Not TieneObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex) Then
            ' Msg634=No tenés suficientes pieles de oso pardo.
            Call WriteLocaleMsg(UserIndex, 634, e_FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).PielOsoPolaR > 0 Then
        If Not TieneObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex) Then
            ' Msg635=No tenés suficientes pieles de oso polar.
            Call WriteLocaleMsg(UserIndex, 635, e_FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).PielLoboNegro > 0 Then
        If Not TieneObjetos(PielLoboNegro, ObjData(ItemIndex).PielLoboNegro, UserIndex) Then
            ' Msg636=No tenés suficientes pieles de lobo negro.
            Call WriteLocaleMsg(UserIndex, 636, e_FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).PielTigre > 0 Then
        If Not TieneObjetos(PielTigre, ObjData(ItemIndex).PielTigre, UserIndex) Then
            ' Msg637=No tenés suficientes pieles de tigre.
            Call WriteLocaleMsg(UserIndex, 637, e_FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).PielTigreBengala > 0 Then
        If Not TieneObjetos(PielTigreBengala, ObjData(ItemIndex).PielTigreBengala, UserIndex) Then
            ' Msg638=No tenés suficientes pieles de tigre de bengala.
            Call WriteLocaleMsg(UserIndex, 638, e_FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    SastreTieneMateriales = True
    Exit Function
SastreTieneMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.SastreTieneMateriales", Erl)
End Function

Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    On Error GoTo HerreroTieneMateriales_Err
    If ObjData(ItemIndex).LingH > 0 Then
        If Not TieneObjetos(e_Minerales.LingoteDeHierro, ObjData(ItemIndex).LingH, UserIndex) Then
            ' Msg639=No tenés suficientes lingotes de hierro.
            Call WriteLocaleMsg(UserIndex, 639, e_FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
        If Not TieneObjetos(e_Minerales.LingoteDePlata, ObjData(ItemIndex).LingP, UserIndex) Then
            ' Msg640=No tenés suficientes lingotes de plata.
            Call WriteLocaleMsg(UserIndex, 640, e_FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
        If Not TieneObjetos(e_Minerales.LingoteDeOro, ObjData(ItemIndex).LingO, UserIndex) Then
            ' Msg641=No tenés suficientes lingotes de oro.
            Call WriteLocaleMsg(UserIndex, 641, e_FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Coal > 0 Then
        If Not TieneObjetos(e_Minerales.Coal, ObjData(ItemIndex).Coal, UserIndex) Then
            ' Msg642=No tenés suficientes carbón.
            Call WriteLocaleMsg(UserIndex, 642, e_FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).Blodium > 0 Then
        If Not TieneObjetos(e_Minerales.Blodium, ObjData(ItemIndex).Blodium, UserIndex) Then
            Call WriteLocaleMsg(UserIndex, 2089, e_FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).FireEssence > 0 Then
        If Not TieneObjetos(e_Minerales.FireEssence, ObjData(ItemIndex).FireEssence, UserIndex) Then
            Call WriteLocaleMsg(UserIndex, 2090, e_FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).WaterEssence > 0 Then
        If Not TieneObjetos(e_Minerales.WaterEssence, ObjData(ItemIndex).WaterEssence, UserIndex) Then
            Call WriteLocaleMsg(UserIndex, 2090, e_FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).EarthEssence > 0 Then
        If Not TieneObjetos(e_Minerales.EarthEssence, ObjData(ItemIndex).EarthEssence, UserIndex) Then
            Call WriteLocaleMsg(UserIndex, 2090, e_FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    If ObjData(ItemIndex).WindEssence > 0 Then
        If Not TieneObjetos(e_Minerales.WindEssence, ObjData(ItemIndex).WindEssence, UserIndex) Then
            Call WriteLocaleMsg(UserIndex, 2090, e_FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function
        End If
    End If
    HerreroTieneMateriales = True
    Exit Function
HerreroTieneMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.HerreroTieneMateriales", Erl)
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    On Error GoTo PuedeConstruir_Err
    PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(e_Skill.Herreria) >= ObjData(ItemIndex).SkHerreria
    Exit Function
PuedeConstruir_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruir", Erl)
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
    On Error GoTo PuedeConstruirHerreria_Err
    Dim i As Long
    Select Case ObjData(ItemIndex).OBJType
        Case e_OBJType.otWeapon, e_OBJType.otArrows
            For i = 1 To UBound(ArmasHerrero)
                If ArmasHerrero(i) = ItemIndex Then
                    PuedeConstruirHerreria = True
                    Exit Function
                End If
            Next i
        Case e_OBJType.otArmor, e_OBJType.otHelmet, e_OBJType.otShield, e_OBJType.otAmulets, e_OBJType.otRingAccesory
            For i = 1 To UBound(ArmadurasHerrero)
                If ArmadurasHerrero(i) = ItemIndex Then
                    PuedeConstruirHerreria = True
                    Exit Function
                End If
            Next i
        Case e_OBJType.otElementalRune
            For i = 1 To UBound(BlackSmithElementalRunes)
                If BlackSmithElementalRunes(i) = ItemIndex Then
                    PuedeConstruirHerreria = True
                    Exit Function
                End If
            Next i
    End Select
    PuedeConstruirHerreria = False
    Exit Function
PuedeConstruirHerreria_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirHerreria", Erl)
End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    On Error GoTo HerreroConstruirItem_Err
    If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
    If Not HayLugarEnInventario(UserIndex, ItemIndex, 1) Then
        ' Msg643=No tienes suficiente espacio en el inventario.
        Call WriteLocaleMsg(UserIndex, 643, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).flags.Privilegios And (e_PlayerType.Consejero) Then
        Exit Sub
    End If
    If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) And KnowsCraftingRecipe(UserIndex, ItemIndex) Then
        Call HerreroQuitarMateriales(UserIndex, ItemIndex)
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
        Call WriteUpdateSta(UserIndex)
        ' AGREGAR FX
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, 253, 25, False, ObjData(ItemIndex).GrhIndex))
        Select Case ObjData(ItemIndex).OBJType
            Case e_OBJType.otWeapon
                Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.charindex, vbWhite)
            Case e_OBJType.otShield
                Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.charindex, vbWhite)
            Case e_OBJType.otHelmet
                Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.charindex, vbWhite)
            Case e_OBJType.otArmor
                Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.charindex, vbWhite)
            Case e_OBJType.otElementalRune
                Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.charindex, vbWhite)
            Case Else
        End Select
        Dim MiObj As t_Obj
        MiObj.amount = 1
        MiObj.ObjIndex = ItemIndex
        MiObj.ElementalTags = ObjData(ItemIndex).ElementalTags
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
        End If
        Call SubirSkill(UserIndex, e_Skill.Herreria)
        Call UpdateUserInv(True, UserIndex, 0)
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    End If
    Exit Sub
HerreroConstruirItem_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.HerreroConstruirItem", Erl)
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
    On Error GoTo PuedeConstruirCarpintero_Err
    Dim i As Long
    For i = 1 To UBound(ObjCarpintero)
        If ObjCarpintero(i) = ItemIndex Then
            PuedeConstruirCarpintero = True
            Exit Function
        End If
    Next i
    PuedeConstruirCarpintero = False
    Exit Function
PuedeConstruirCarpintero_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirCarpintero", Erl)
End Function

Public Function PuedeConstruirAlquimista(ByVal ItemIndex As Integer) As Boolean
    On Error GoTo PuedeConstruirAlquimista_Err
    Dim i As Long
    For i = 1 To UBound(ObjAlquimista)
        If ObjAlquimista(i) = ItemIndex Then
            PuedeConstruirAlquimista = True
            Exit Function
        End If
    Next i
    PuedeConstruirAlquimista = False
    Exit Function
PuedeConstruirAlquimista_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirAlquimista", Erl)
End Function

Public Function PuedeConstruirSastre(ByVal ItemIndex As Integer) As Boolean
    On Error GoTo PuedeConstruirSastre_Err
    Dim i As Long
    For i = 1 To UBound(ObjSastre)
        If ObjSastre(i) = ItemIndex Then
            PuedeConstruirSastre = True
            Exit Function
        End If
    Next i
    PuedeConstruirSastre = False
    Exit Function
PuedeConstruirSastre_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirSastre", Erl)
End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Long, ByVal cantidad_maxima As Integer)
    On Error GoTo CarpinteroConstruirItem_Err
    If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
    If UserList(UserIndex).flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios) Then
        Exit Sub
    End If
    If ItemIndex = 0 Then Exit Sub
    'Si no tiene equipado el serrucho
    If UserList(UserIndex).invent.EquippedWorkingToolObjIndex = 0 Then
        ' Antes de usar la herramienta deberias equipartela.
        Call WriteLocaleMsg(UserIndex, 376, e_FontTypeNames.FONTTYPE_INFO)
        Call WriteMacroTrabajoToggle(UserIndex, False)
        Exit Sub
    End If
    Dim cantidad_a_construir    As Long
    Dim madera_requerida        As Long
    Dim madera_elfica_requerida As Long
    Dim madera_pino_requerida   As Long
    cantidad_a_construir = IIf(UserList(UserIndex).Trabajo.Cantidad >= cantidad_maxima, cantidad_maxima, UserList(UserIndex).Trabajo.Cantidad)
    If cantidad_a_construir <= 0 Then
        Call WriteMacroTrabajoToggle(UserIndex, False)
        Exit Sub
    End If
    If CarpinteroTieneMateriales(UserIndex, ItemIndex, cantidad_a_construir) And UserList(UserIndex).Stats.UserSkills(e_Skill.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria _
            And PuedeConstruirCarpintero(ItemIndex) And ObjData(UserList(UserIndex).invent.EquippedWorkingToolObjIndex).OBJType = e_OBJType.otWorkingTools And ObjData(UserList( _
            UserIndex).invent.EquippedWorkingToolObjIndex).Subtipo = 5 Then
        If UserList(UserIndex).Stats.MinSta > 2 Then
            Call QuitarSta(UserIndex, 2)
        Else
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            'Msg93=Estás muy cansado para trabajar.
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If
        If ObjData(ItemIndex).Madera > 0 Then madera_requerida = ObjData(ItemIndex).Madera * cantidad_a_construir
        If ObjData(ItemIndex).MaderaElfica > 0 Then madera_elfica_requerida = ObjData(ItemIndex).MaderaElfica * cantidad_a_construir
        If ObjData(ItemIndex).MaderaPino > 0 Then madera_pino_requerida = ObjData(ItemIndex).MaderaPino * cantidad_a_construir
        Call CarpinteroQuitarMateriales(UserIndex, ItemIndex, madera_requerida, madera_elfica_requerida, madera_pino_requerida)
        UserList(UserIndex).Trabajo.Cantidad = UserList(UserIndex).Trabajo.Cantidad - cantidad_a_construir
        Call WriteTextCharDrop(UserIndex, "+" & cantidad_a_construir, UserList(UserIndex).Char.charindex, vbWhite)
        Dim MiObj As t_Obj
        MiObj.amount = cantidad_a_construir
        MiObj.ObjIndex = ItemIndex
        MiObj.ElementalTags = ObjData(ItemIndex).ElementalTags
        ' AGREGAR FX
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
        End If
        Call SubirSkill(UserIndex, e_Skill.Carpinteria)
        'Call UpdateUserInv(True, UserIndex, 0)
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    End If
    Exit Sub
CarpinteroConstruirItem_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.CarpinteroConstruirItem", Erl)
End Sub

Public Sub AlquimistaConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    On Error GoTo AlquimistaConstruirItem_Err
    If Not UserList(UserIndex).Stats.MinSta > 0 Then
        Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    ' === [ Validate Array Bounds Before Accessing Elements ] ===
    ' Check if UserIndex is valid
    If UserIndex < LBound(UserList) Or UserIndex > UBound(UserList) Then
        Call TraceError(1001, "UserIndex out of range: " & UserIndex, "AlquimistaConstruirItem", Erl)
        Exit Sub
    End If
    ' Check if ItemIndex is valid
    If ItemIndex < LBound(ObjData) Or ItemIndex > UBound(ObjData) Then
        Call TraceError(1002, "ItemIndex out of range: " & ItemIndex, "AlquimistaConstruirItem", Erl)
        Exit Sub
    End If
    ' Check if the equipped tool index is valid
    Dim ToolIndex As Integer
    ToolIndex = UserList(UserIndex).invent.EquippedWorkingToolObjIndex
    If ToolIndex < LBound(ObjData) Or ToolIndex > UBound(ObjData) Then
        Call TraceError(1003, "EquippedWorkingToolObjIndex out of range: " & ToolIndex, "AlquimistaConstruirItem", Erl)
        Exit Sub
    End If
    ' === [ Main Logic ] ===
    If AlquimistaTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(e_Skill.Alquimia) >= ObjData(ItemIndex).SkPociones And PuedeConstruirAlquimista( _
            ItemIndex) And ObjData(ToolIndex).OBJType = e_OBJType.otWorkingTools And ObjData(ToolIndex).Subtipo = 4 And KnowsCraftingRecipe(UserIndex, ItemIndex) Then
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 1
        Call WriteUpdateSta(UserIndex)
        ' AGREGAR FX
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, 253, 25, False, ObjData(ItemIndex).GrhIndex))
        Call AlquimistaQuitarMateriales(UserIndex, ItemIndex)
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(1152, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        Dim MiObj As t_Obj
        MiObj.amount = 1
        MiObj.ObjIndex = ItemIndex
        MiObj.ElementalTags = ObjData(ItemIndex).ElementalTags
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
        End If
        Call SubirSkill(UserIndex, e_Skill.Alquimia)
        Call UpdateUserInv(True, UserIndex, 0)
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    End If
    Exit Sub
AlquimistaConstruirItem_Err:
    Call TraceError(Err.Number, Err.Description & " | UserIndex: " & UserIndex & " | ItemIndex: " & ItemIndex, "Trabajo.AlquimistaConstruirItem", Erl)
    Resume Next ' Allow execution to continue after logging the error
End Sub

Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    On Error GoTo SastreConstruirItem_Err
    If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
    If Not UserList(UserIndex).Stats.MinSta > 0 Then
        Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If ItemIndex = 0 Then Exit Sub
    If UserList(UserIndex).invent.EquippedWorkingToolObjIndex = 0 Then
        Exit Sub
    End If
    If SastreTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(e_Skill.Sastreria) >= ObjData(ItemIndex).SkSastreria And PuedeConstruirSastre( _
            ItemIndex) And ObjData(UserList(UserIndex).invent.EquippedWorkingToolObjIndex).OBJType = e_OBJType.otWorkingTools And ObjData(UserList( _
            UserIndex).invent.EquippedWorkingToolObjIndex).Subtipo = 9 Then
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
        Call WriteUpdateSta(UserIndex)
        Call SastreQuitarMateriales(UserIndex, ItemIndex)
        Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.charindex, vbWhite)
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(63, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        Dim MiObj As t_Obj
        MiObj.amount = 1
        MiObj.ObjIndex = ItemIndex
        MiObj.ElementalTags = ObjData(ItemIndex).ElementalTags
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
        End If
        Call SubirSkill(UserIndex, e_Skill.Sastreria)
        Call UpdateUserInv(True, UserIndex, 0)
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    End If
    Exit Sub
SastreConstruirItem_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.SastreConstruirItem", Erl)
End Sub

Private Function MineralesParaLingote(ByVal Lingote As e_Minerales, ByVal cant As Byte) As Integer
    On Error GoTo MineralesParaLingote_Err
    Select Case Lingote
        Case e_Minerales.HierroCrudo
            MineralesParaLingote = 13 * cant
        Case e_Minerales.PlataCruda
            MineralesParaLingote = 25 * cant
        Case e_Minerales.OroCrudo
            MineralesParaLingote = 50 * cant
        Case Else
            MineralesParaLingote = 10000
    End Select
    Exit Function
MineralesParaLingote_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.MineralesParaLingote", Erl)
End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)
    On Error GoTo DoLingotes_Err
    Dim Slot       As Integer
    Dim obji       As Integer
    Dim cant       As Byte
    Dim necesarios As Integer
    If UserList(UserIndex).Stats.MinSta > 2 Then
        Call QuitarSta(UserIndex, 2)
    Else
        Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
        'Msg93=Estás muy cansado para excavar.
        Call WriteMacroTrabajoToggle(UserIndex, False)
        Exit Sub
    End If
    Slot = UserList(UserIndex).flags.TargetObjInvSlot
    obji = UserList(UserIndex).invent.Object(Slot).ObjIndex
    cant = RandomNumber(10, 20)
    necesarios = MineralesParaLingote(obji, cant)
    If UserList(UserIndex).invent.Object(Slot).amount < MineralesParaLingote(obji, cant) Or ObjData(obji).OBJType <> e_OBJType.otMinerals Then
        ' Msg645=No tienes suficientes minerales para hacer un lingote.
        Call WriteLocaleMsg(UserIndex, 645, e_FontTypeNames.FONTTYPE_INFO)
        Call WriteMacroTrabajoToggle(UserIndex, False)
        Exit Sub
    End If
    UserList(UserIndex).invent.Object(Slot).amount = UserList(UserIndex).invent.Object(Slot).amount - MineralesParaLingote(obji, cant)
    If UserList(UserIndex).invent.Object(Slot).amount < 1 Then
        UserList(UserIndex).invent.Object(Slot).amount = 0
        UserList(UserIndex).invent.Object(Slot).ObjIndex = 0
    End If
    Dim nPos  As t_WorldPos
    Dim MiObj As t_Obj
    MiObj.amount = cant
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
    End If
    Call UpdateUserInv(False, UserIndex, Slot)
    Call WriteTextCharDrop(UserIndex, "+" & cant, UserList(UserIndex).Char.charindex, vbWhite)
    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(41, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
    Call SubirSkill(UserIndex, e_Skill.Mineria)
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    If UserList(UserIndex).Counters.Trabajando = 1 And Not UserList(UserIndex).flags.UsandoMacro Then
        Call WriteMacroTrabajoToggle(UserIndex, True)
    End If
    Exit Sub
DoLingotes_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.DoLingotes", Erl)
End Sub

Function ModAlquimia(ByVal clase As e_Class) As Integer
    On Error GoTo ModAlquimia_Err
    Select Case clase
        Case e_Class.Druid
            ModAlquimia = 1
        Case e_Class.Trabajador
            ModAlquimia = 1
        Case Else
            ModAlquimia = 3
    End Select
    Exit Function
ModAlquimia_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ModAlquimia", Erl)
End Function

Function ModSastre(ByVal clase As e_Class) As Integer
    On Error GoTo ModSastre_Err
    Select Case clase
        Case e_Class.Trabajador
            ModSastre = 1
        Case Else
            ModSastre = 3
    End Select
    Exit Function
ModSastre_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ModSastre", Erl)
End Function

Function ModCarpinteria(ByVal clase As e_Class) As Integer
    On Error GoTo ModCarpinteria_Err
    Select Case clase
        Case e_Class.Trabajador
            ModCarpinteria = 1
        Case Else
            ModCarpinteria = 3
    End Select
    Exit Function
ModCarpinteria_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ModCarpinteria", Erl)
End Function

Function ModHerreria(ByVal clase As e_Class) As Single
    On Error GoTo ModHerreriA_Err
    Select Case clase
        Case e_Class.Trabajador
            ModHerreria = 1
        Case Else
            ModHerreria = 3
    End Select
    Exit Function
ModHerreriA_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ModHerreriA", Erl)
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer, Optional ByVal invisible As Byte = 2)
    On Error GoTo DoAdminInvisible_Err
    With UserList(UserIndex)
        If invisible = 2 Then
            .flags.AdminInvisible = IIf(.flags.AdminInvisible = 1, 0, 1)
        Else
            .flags.AdminInvisible = invisible
        End If
        If .flags.AdminInvisible = 1 Then
            .flags.invisible = 1
            .flags.Oculto = 1
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
            Call SendData(SendTarget.ToPCAreaButGMs, UserIndex, PrepareMessageCharacterRemove(2, .Char.charindex, True))
        Else
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            Call MakeUserChar(True, 0, UserIndex, .pos.Map, .pos.x, .pos.y, 1)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))
        End If
    End With
    Exit Sub
DoAdminInvisible_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.DoAdminInvisible", Erl)
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal UserIndex As Integer)
    On Error GoTo TratarDeHacerFogata_Err
    Dim Suerte    As Byte
    Dim exito     As Byte
    Dim obj       As t_Obj
    Dim posMadera As t_WorldPos
    If Not LegalPos(Map, x, y) Then Exit Sub
    With posMadera
        .Map = Map
        .x = x
        .y = y
    End With
    If MapData(Map, x, y).ObjInfo.ObjIndex <> 58 Then
        ' Msg646=Necesitas clickear sobre Leña para hacer ramitas.
        Call WriteLocaleMsg(UserIndex, 646, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If Distancia(posMadera, UserList(UserIndex).pos) > 2 Then
        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
        'Call WriteLocaleMsg(UserIndex, 1455, e_FontTypeNames.FONTTYPE_INFO)  ' Msg1455=Estás demasiado lejos para prender la fogata.
        Exit Sub
    End If
    If UserList(UserIndex).flags.Muerto = 1 Then
        ' Msg647=No podés hacer fogatas estando muerto.
        Call WriteLocaleMsg(UserIndex, 647, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If MapData(Map, x, y).ObjInfo.amount < 3 Then
        ' Msg648=Necesitas por lo menos tres troncos para hacer una fogata.
        Call WriteLocaleMsg(UserIndex, 648, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) < 6 Then
        Suerte = 3
    ElseIf UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) <= 34 Then
        Suerte = 2
    ElseIf UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 35 Then
        Suerte = 1
    End If
    exito = RandomNumber(1, Suerte)
    If exito = 1 Then
        obj.ObjIndex = FOGATA_APAG
        obj.amount = MapData(Map, x, y).ObjInfo.amount \ 3
        Call WriteLocaleMsg(UserIndex, 1456, e_FontTypeNames.FONTTYPE_INFO)  ' Msg1456=Has hecho ¬1 ramitas.
        Call MakeObj(obj, Map, x, y)
        'Seteamos la fogata como el nuevo TargetObj del user
        UserList(UserIndex).flags.TargetObj = FOGATA_APAG
    End If
    Call SubirSkill(UserIndex, Supervivencia)
    Exit Sub
TratarDeHacerFogata_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.TratarDeHacerFogata", Erl)
End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, Optional ByVal RedDePesca As Boolean = False)
    On Error GoTo ErrHandler
    Dim bonificacionPescaLvl(1 To 47) As Single
    Dim bonificacionCaña As Double
    Dim bonificacionZona  As Double
    Dim bonificacionLvl   As Double
    Dim bonificacionClase As Double
    Dim bonificacionTotal As Double
    Dim RestaStamina      As Integer
    Dim Reward            As Double
    Dim esEspecial        As Boolean
    Dim i                 As Integer
    Dim NpcIndex          As Integer
    ' Shugar - 13/8/2024
    ' Paso los poderes de las cañas al dateo de pesca.dat
    ' Paso la reducción de pesca en zona segura a balance.dat
    With UserList(UserIndex)
        RestaStamina = IIf(RedDePesca, 12, RandomNumber(2, 3))
        If .flags.Privilegios And (e_PlayerType.Consejero) Then
            Exit Sub
        End If
        If .Stats.MinSta > RestaStamina Then
            Call QuitarSta(UserIndex, RestaStamina)
        Else
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If
        If MapInfo(.pos.Map).Seguro = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageArmaMov(.Char.charindex, 0))
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(.Char.charindex, 0))
        End If
        bonificacionPescaLvl(1) = 0
        bonificacionPescaLvl(2) = 0.009
        bonificacionPescaLvl(3) = 0.015
        bonificacionPescaLvl(4) = 0.019
        bonificacionPescaLvl(5) = 0.025
        bonificacionPescaLvl(6) = 0.03
        bonificacionPescaLvl(7) = 0.035
        bonificacionPescaLvl(8) = 0.04
        bonificacionPescaLvl(9) = 0.045
        bonificacionPescaLvl(10) = 0.05
        bonificacionPescaLvl(11) = 0.06
        bonificacionPescaLvl(12) = 0.07
        bonificacionPescaLvl(13) = 0.08
        bonificacionPescaLvl(14) = 0.09
        bonificacionPescaLvl(15) = 0.1
        bonificacionPescaLvl(16) = 0.11
        bonificacionPescaLvl(17) = 0.13
        bonificacionPescaLvl(18) = 0.14
        bonificacionPescaLvl(19) = 0.16
        bonificacionPescaLvl(20) = 0.18
        bonificacionPescaLvl(21) = 0.2
        bonificacionPescaLvl(22) = 0.22
        bonificacionPescaLvl(23) = 0.24
        bonificacionPescaLvl(24) = 0.27
        bonificacionPescaLvl(25) = 0.3
        bonificacionPescaLvl(26) = 0.32
        bonificacionPescaLvl(27) = 0.35
        bonificacionPescaLvl(28) = 0.37
        bonificacionPescaLvl(29) = 0.4
        bonificacionPescaLvl(30) = 0.43
        bonificacionPescaLvl(31) = 0.47
        bonificacionPescaLvl(32) = 0.51
        bonificacionPescaLvl(33) = 0.55
        bonificacionPescaLvl(34) = 0.58
        bonificacionPescaLvl(35) = 0.62
        bonificacionPescaLvl(36) = 0.7
        bonificacionPescaLvl(37) = 0.77
        bonificacionPescaLvl(38) = 0.84
        bonificacionPescaLvl(39) = 0.92
        bonificacionPescaLvl(40) = 1#
        bonificacionPescaLvl(41) = 1.1
        bonificacionPescaLvl(42) = 1.15
        bonificacionPescaLvl(43) = 1.3
        bonificacionPescaLvl(44) = 1.5
        bonificacionPescaLvl(45) = 1.8
        bonificacionPescaLvl(46) = 2#
        bonificacionPescaLvl(47) = 2.5
        'Bonificación según el nivel
        bonificacionLvl = 1 + bonificacionPescaLvl(.Stats.ELV)
        'Bonificacion de la caña dependiendo de su poder:
        bonificacionCaña = PoderCanas(ObjData(.invent.EquippedWorkingToolObjIndex).Power) / 10
        'Bonificación total
        bonificacionTotal = bonificacionCaña * bonificacionLvl * SvrConfig.GetValue("RecoleccionMult")
        'Si es zona segura se aplica una penalización
        If MapInfo(.pos.Map).Seguro Then
            bonificacionTotal = bonificacionTotal * PorcentajePescaSegura / 100
        End If
        'Shugar: La reward ya estaba hardcodeada así...
        'no la voy a tocar, pero ahora por lo menos puede ajustarse desde dateo con la bonificación de las cañas!
        'Calculo el botin esperado por iteracción. 'La base del calculo son 8000 por hora + 20% de chances de no pescar + un +/- 10%
        Reward = (IntervaloTrabajarExtraer / 3600000) * 8000 * bonificacionTotal * 1.2 * (1 + (RandomNumber(0, 20) - 10) / 100)
        'Calculo la suerte de pescar o no pescar y aplico eso sobre el reward para promediar.
        Dim Suerte As Integer
        Dim Pesco  As Boolean
        If .Stats.UserSkills(e_Skill.Pescar) < 20 Then
            Suerte = 20
        ElseIf .Stats.UserSkills(e_Skill.Pescar) < 40 Then
            Suerte = 35
        ElseIf .Stats.UserSkills(e_Skill.Pescar) < 70 Then
            Suerte = 55
        ElseIf .Stats.UserSkills(e_Skill.Pescar) < 100 Then
            Suerte = 68
        Else
            Suerte = 80
        End If
        Pesco = RandomNumber(1, 100) <= Suerte '80% de posibilidad de pescar
        If Pesco Then
            Dim nPos     As t_WorldPos
            Dim MiObj    As t_Obj
            Dim objValue As Integer
            If IsFeatureEnabled("gain_exp_while_working") Then
                Call GiveExpWhileWorking(UserIndex, UserList(UserIndex).invent.EquippedWorkingToolObjIndex, e_JobsTypes.Fisherman)
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
            End If
            ' Shugar: al final no importa el valor del pez ya que se ajusta la cantidad...
            ' Genero el obj pez que pesqué y su cantidad
            MiObj.ObjIndex = ObtenerPezRandom(ObjData(.invent.EquippedWorkingToolObjIndex).Power)
            objValue = max(ObjData(MiObj.ObjIndex).Valor / 3, 1)
            'si esta macreando y para que esten mas atentos les mando un NPC y saco el macro de trabajar
            If MiObj.ObjIndex = (SvrConfig.GetValue("FISHING_SPECIALFISH1_ID") Or MiObj.ObjIndex = SvrConfig.GetValue("FISHING_SPECIALFISH2_ID")) And (UserList( _
                    UserIndex).pos.Map) <> SvrConfig.GetValue("FISHING_MAP_SPECIAL_FISH1_ID") Then
                MiObj.ObjIndex = SvrConfig.GetValue("FISHING_SPECIALFISH1_REMPLAZO_ID")
                If MapInfo(UserList(UserIndex).pos.Map).Seguro = 0 Then NpcIndex = SpawnNpc(SvrConfig.GetValue("NPC_WATCHMAN_ID"), .pos, True, False)
                Call WriteMacroTrabajoToggle(UserIndex, False)
            End If
            MiObj.amount = Round(Reward / objValue)
            If MiObj.amount <= 0 Then
                MiObj.amount = 1
            End If
            Dim StopWorking As Boolean
            StopWorking = False
            ' Si es insegura y es un fishing pool:
            If MapInfo(UserList(UserIndex).pos.Map).Seguro = 0 And SvrConfig.GetValue("FISHING_POOL_ID") = MapData(.pos.Map, .Trabajo.Target_X, _
                    .Trabajo.Target_Y).ObjInfo.ObjIndex Then
                ' Si se está por vaciar el fishing pool:
                If MiObj.amount > MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount Then
                    MiObj.amount = MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount
                    Call CreateFishingPool(.pos.Map)
                    Call EraseObj(MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount, .pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y)
                    ' Msg649=No hay mas peces aqui.
                    Call WriteLocaleMsg(UserIndex, 649, e_FontTypeNames.FONTTYPE_INFO)
                    StopWorking = True
                End If
                ' Resto los recursos que saqué
                MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount = MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount - MiObj.amount
            End If
            ' Verifico si el pescado es especial o no
            For i = 1 To UBound(PecesEspeciales)
                If PecesEspeciales(i).ObjIndex = MiObj.ObjIndex Then
                    esEspecial = True
                End If
            Next i
            ' Si no es especial, actualizo el UserIndex
            If Not esEspecial Then
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
                ' Si es especial, corto el macro y activo el minijuego
                ' Solo aplica a cañas, no a red de pesca
            Else
                .flags.PescandoEspecial = True
                Call WriteMacroTrabajoToggle(UserIndex, False)
                .Stats.NumObj_PezEspecial = MiObj.ObjIndex
                Call WritePelearConPezEspecial(UserIndex)
                Exit Sub
            End If
            If MiObj.ObjIndex = 0 Then Exit Sub
            ' Si no entra en el inventario dejo de pescar
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                StopWorking = True
            End If
            Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.charindex, vbWhite)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .pos.x, .pos.y))
            ' Al pescar también podés sacar cosas raras (se setean desde RecursosEspeciales.dat)
            ' Por cada drop posible
            Dim res As Long
            For i = 1 To UBound(EspecialesPesca)
                ' Tiramos al azar entre 1 y la probabilidad
                res = RandomNumber(1, IIf(RedDePesca, EspecialesPesca(i).data * 2, EspecialesPesca(i).data)) ' Red de pesca chance x2 (revisar)
                ' Si tiene suerte y le pega
                If res = 1 Then
                    MiObj.ObjIndex = EspecialesPesca(i).ObjIndex
                    MiObj.amount = 1 ' Solo un item por vez
                    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.pos, MiObj)
                    ' Le mandamos un mensaje
                    Call WriteLocaleMsg(UserIndex, 1457, e_FontTypeNames.FONTTYPE_INFO)  ' Msg1457=¡Has conseguido ¬1!
                End If
            Next
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, GRH_FALLO_PESCA))
        End If
        If MapInfo(UserList(UserIndex).pos.Map).Seguro = 0 Then
            Call SubirSkill(UserIndex, e_Skill.Pescar)
        End If
        If StopWorking Then
            Call WriteWorkRequestTarget(UserIndex, 0)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If
        .Counters.Trabajando = .Counters.Trabajando + 1
        .Counters.LastTrabajo = Int(IntervaloTrabajarExtraer / 1000)
        'Ladder 06/07/14 Activamos el macro de trabajo
        If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
            Call WriteMacroTrabajoToggle(UserIndex, True)
        End If
    End With
    Exit Sub
ErrHandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " - " & Err.Description & " Line number: " & Erl)
End Sub

''
' Try to steal an item / gold to another character
'
' @param LadronIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub DoRobar(ByVal LadronIndex As Integer, ByVal VictimaIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 05/04/2010
    'Last Modification By: ZaMa
    '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadronIndex) when the thief stoles gold. (MarKoxX)
    '27/11/2009: ZaMa - Optimizacion de codigo.
    '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
    '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
    '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
    '23/04/2010: ZaMa - No se puede robar mas sin energia.
    '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
    '*************************************************
    On Error GoTo ErrHandler
    Dim OtroUserIndex As Integer
    If UserList(LadronIndex).flags.Privilegios And (e_PlayerType.Consejero) Then Exit Sub
    If MapInfo(UserList(VictimaIndex).pos.Map).Seguro = 1 Then Exit Sub
    If Not UserMod.CanMove(UserList(VictimaIndex).flags, UserList(VictimaIndex).Counters) Then
        'Msg1028= No podes robarle a objetivos inmovilizados.
        Call WriteLocaleMsg(LadronIndex, "1028", e_FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    If UserList(VictimaIndex).flags.EnConsulta Then
        'Msg1029= ¡No puedes robar a usuarios en consulta!
        Call WriteLocaleMsg(LadronIndex, "1029", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Dim Penable As Boolean
    With UserList(LadronIndex)
        If esCiudadano(LadronIndex) Then
            If (.flags.Seguro) Then
                'Msg1030= Debes quitarte el seguro para robarle a un ciudadano o a un miembro del Ejército Real
                Call WriteLocaleMsg(LadronIndex, "1030", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        ElseIf esArmada(LadronIndex) Then ' Armada robando a armada or ciudadano?
            If (esCiudadano(VictimaIndex) Or esArmada(VictimaIndex)) Then
                'Msg1031= Los miembros del Ejército Real no tienen permitido robarle a ciudadanos o a otros miembros del Ejército Real
                Call WriteLocaleMsg(LadronIndex, "1031", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        ElseIf esCaos(LadronIndex) Then ' Caos robando a caos?
            If (esCaos(VictimaIndex)) Then
                'Msg1032= No puedes robar a otros miembros de la Legión Oscura.
                Call WriteLocaleMsg(LadronIndex, "1032", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        End If
        'Me fijo si el ladrón tiene clan
        If .GuildIndex > 0 Then
            'Si tiene clan me fijo si su clan es de alineación ciudadana
            If esCiudadano(LadronIndex) And GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                If PersonajeEsLeader(.Id) Then
                    'Msg1033= No puedes robar siendo lider de un clan ciudadano.
                    Call WriteLocaleMsg(LadronIndex, "1033", e_FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If
        End If
        If TriggerZonaPelea(LadronIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .genero = e_Genero.Hombre Then
                'Msg1034= Estás muy cansado para robar.
                Call WriteLocaleMsg(LadronIndex, "1034", e_FontTypeNames.FONTTYPE_INFO)
            Else
                'Msg1035= Estás muy cansada para robar.
                Call WriteLocaleMsg(LadronIndex, "1035", e_FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        If .GuildIndex > 0 Then
            If .flags.SeguroClan And NivelDeClan(.GuildIndex) >= 3 Then
                If .GuildIndex = UserList(VictimaIndex).GuildIndex Then
                    'Msg1036= No podes robarle a un miembro de tu clan.
                    Call WriteLocaleMsg(LadronIndex, "1036", e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
            End If
        End If
        ' Quito energia
        Call QuitarSta(LadronIndex, 15)
        If UserList(VictimaIndex).flags.Privilegios And e_PlayerType.User Then
            Dim Probabilidad As Byte
            Dim res          As Integer
            Dim RobarSkill   As Byte
            RobarSkill = .Stats.UserSkills(e_Skill.Robar)
            If (RobarSkill > 0 And RobarSkill < 10) Then
                Probabilidad = 1
            ElseIf (RobarSkill >= 10 And RobarSkill <= 20) Then
                Probabilidad = 5
            ElseIf (RobarSkill >= 20 And RobarSkill <= 30) Then
                Probabilidad = 10
            ElseIf (RobarSkill >= 30 And RobarSkill <= 40) Then
                Probabilidad = 15
            ElseIf (RobarSkill >= 40 And RobarSkill <= 50) Then
                Probabilidad = 25
            ElseIf (RobarSkill >= 50 And RobarSkill <= 60) Then
                Probabilidad = 35
            ElseIf (RobarSkill >= 60 And RobarSkill <= 70) Then
                Probabilidad = 40
            ElseIf (RobarSkill >= 70 And RobarSkill <= 80) Then
                Probabilidad = 55
            ElseIf (RobarSkill >= 80 And RobarSkill <= 90) Then
                Probabilidad = 70
            ElseIf (RobarSkill >= 90 And RobarSkill < 100) Then
                Probabilidad = 80
            ElseIf (RobarSkill = 100) Then
                Probabilidad = 90
            End If
            If (RandomNumber(1, 100) < Probabilidad) Then 'Exito robo
                If UserList(VictimaIndex).flags.Comerciando Then
                    OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu.ArrayIndex
                    If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                        'Msg1037= Comercio cancelado, ¡te están robando!
                        Call WriteLocaleMsg(VictimaIndex, "1037", e_FontTypeNames.FONTTYPE_TALK)
                        'Msg1038= Comercio cancelado, al otro usuario le robaron.
                        Call WriteLocaleMsg(OtroUserIndex, "1038", e_FontTypeNames.FONTTYPE_TALK)
                        Call LimpiarComercioSeguro(VictimaIndex)
                    End If
                End If
                If (RandomNumber(1, 50) < 25) And (.clase = e_Class.Thief) Then '50% de robar items
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadronIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadronIndex, PrepareMessageLocaleMsg(1867, UserList(VictimaIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1867=¬1 no tiene objetos.
                    End If
                Else '50% de robar oro
                    If UserList(VictimaIndex).Stats.GLD > 0 Then
                        Dim n     As Long
                        Dim Extra As Single
                        ' Multiplicador extra por niveles
                        Select Case .Stats.ELV
                            Case Is < 13
                                Extra = 1
                            Case Is < 25
                                Extra = 1.1
                            Case Is < 35
                                Extra = 1.2
                            Case Is >= 35 And .Stats.ELV <= 40
                                Extra = 1.3
                            Case Is >= 41 And .Stats.ELV < 45
                                Extra = 1.4
                            Case Is >= 45 And .Stats.ELV <= 46
                                Extra = 1.5
                            Case Is = 47
                                Extra = 5
                        End Select
                        If .clase = e_Class.Thief Then
                            'Si no tiene puestos los guantes de hurto roba un 50% menos.
                            If .invent.EquippedWeaponObjIndex > 0 Then
                                If ObjData(.invent.EquippedWeaponObjIndex).Subtipo = 5 Then
                                    n = RandomNumber(.Stats.ELV * 50 * Extra, .Stats.ELV * 100 * Extra) * SvrConfig.GetValue("GoldMult")
                                Else
                                    n = RandomNumber(.Stats.ELV * 25 * Extra, .Stats.ELV * 50 * Extra) * SvrConfig.GetValue("GoldMult")
                                End If
                            Else
                                n = RandomNumber(.Stats.ELV * 25 * Extra, .Stats.ELV * 50 * Extra) * SvrConfig.GetValue("GoldMult")
                            End If
                        Else
                            n = RandomNumber(1, 100) * SvrConfig.GetValue("GoldMult")
                        End If
                        If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                        Dim prevGold As Long: prevGold = UserList(VictimaIndex).Stats.GLD
                        UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                        Dim ProtectedGold As Long
                        ProtectedGold = SvrConfig.GetValue("OroPorNivelBilletera") * UserList(VictimaIndex).Stats.ELV
                        If prevGold >= ProtectedGold And UserList(VictimaIndex).Stats.GLD < ProtectedGold Then
                            n = prevGold - ProtectedGold
                            UserList(VictimaIndex).Stats.GLD = ProtectedGold
                        End If
                        .Stats.GLD = .Stats.GLD + n
                        If .Stats.GLD > MAXORO Then .Stats.GLD = MAXORO
                        Call WriteLocaleMsg(LadronIndex, "1458", e_FontTypeNames.FONTTYPE_New_Rojo_Salmon, PonerPuntos(n) & "¬" & UserList(VictimaIndex).name) ' Msg1458=Le has robado ¬1 monedas de oro a ¬2
                        Call WriteLocaleMsg(VictimaIndex, "1530", e_FontTypeNames.FONTTYPE_New_Rojo_Salmon, UserList(LadronIndex).name & "¬" & PonerPuntos(n)) 'Msg1530=¬1 te ha robado ¬2 monedas de oro.
                        Call WriteUpdateGold(LadronIndex) 'Le actualizamos la billetera al ladron
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                    Else
                        Call WriteConsoleMsg(LadronIndex, PrepareMessageLocaleMsg(1868, UserList(VictimaIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1868=¬1 no tiene oro.
                    End If
                End If
                Call SubirSkill(LadronIndex, e_Skill.Robar)
            Else
                'Msg1039= ¡No has logrado robar nada!
                Call WriteLocaleMsg(LadronIndex, "1039", e_FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(VictimaIndex, "1459", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1459=¡¬1 ha intentado robarte!
                Call SubirSkill(LadronIndex, e_Skill.Robar)
            End If
            If Status(LadronIndex) = Ciudadano Then Call VolverCriminal(LadronIndex)
        End If
    End With
    Exit Sub
ErrHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.Description)
End Sub

Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
    ' Agregué los barcos
    ' Agrego poción negra
    ' Esta funcion determina qué objetos son robables.
    On Error GoTo ObjEsRobable_Err
    Dim OI As Integer
    OI = UserList(VictimaIndex).invent.Object(Slot).ObjIndex
    ObjEsRobable = ObjData(OI).OBJType <> e_OBJType.otKeys And ObjData(OI).OBJType <> e_OBJType.otShips And ObjData(OI).OBJType <> e_OBJType.otSaddles And ObjData(OI).OBJType <> _
            e_OBJType.otRecallStones And ObjData(OI).ObjDonador = 0 And ObjData(OI).Instransferible = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And UserList( _
            VictimaIndex).invent.Object(Slot).Equipped = 0
    Exit Function
ObjEsRobable_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ObjEsRobable", Erl)
End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Private Sub RobarObjeto(ByVal LadronIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010
    '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
    '***************************************************
    On Error GoTo RobarObjeto_Err
    Dim Flag As Boolean
    Dim i    As Integer
    Flag = False
    With UserList(VictimaIndex)
        If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final del inventario?
            i = 1
            Do While Not Flag And i <= .CurrentInventorySlots
                'Hay objeto en este slot?
                If .invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then Flag = True
                    End If
                End If
                If Not Flag Then i = i + 1
            Loop
        Else
            i = .CurrentInventorySlots
            Do While Not Flag And i > 0
                'Hay objeto en este slot?
                If .invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then Flag = True
                    End If
                End If
                If Not Flag Then i = i - 1
            Loop
        End If
        If Flag Then
            Dim MiObj     As t_Obj
            Dim num       As Integer
            Dim ObjAmount As Integer
            ObjAmount = .invent.Object(i).amount
            'Cantidad al azar entre el 3 y el 6% del total, con minimo 1.
            num = MaximoInt(1, RandomNumber(ObjAmount * 0.03, ObjAmount * 0.06))
            MiObj.amount = num
            MiObj.ObjIndex = .invent.Object(i).ObjIndex
            .invent.Object(i).amount = ObjAmount - num
            If .invent.Object(i).amount <= 0 Then
                Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
            End If
            Call UpdateUserInv(False, VictimaIndex, CByte(i))
            If Not MeterItemEnInventario(LadronIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(LadronIndex).pos, MiObj)
            End If
            If UserList(LadronIndex).clase = e_Class.Thief Then
                Call WriteLocaleMsg(LadronIndex, "1460", e_FontTypeNames.FONTTYPE_New_Rojo_Salmon, MiObj.amount & "¬" & ObjData(MiObj.ObjIndex).name)  ' Msg1460=Has robado ¬1 ¬2
                Call WriteLocaleMsg(VictimaIndex, "1531", e_FontTypeNames.FONTTYPE_New_Rojo_Salmon, UserList(LadronIndex).name & "¬" & MiObj.amount & "¬" & ObjData( _
                        MiObj.ObjIndex).name) 'Msg1531=¬1 te ha robado ¬2 ¬3.
            Else
                Call WriteLocaleMsg(LadronIndex, "1461", e_FontTypeNames.FONTTYPE_New_Rojo_Salmon, MiObj.amount & "¬" & ObjData(MiObj.ObjIndex).name)  ' Msg1461=Has hurtado ¬1 ¬2
            End If
        Else
            'Msg1040= No has logrado robar ningun objeto.
            Call WriteLocaleMsg(LadronIndex, "1040", e_FontTypeNames.FONTTYPE_INFO)
        End If
        'If exiting, cancel de quien es robado
        Call CancelExit(VictimaIndex)
    End With
    Exit Sub
RobarObjeto_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.RobarObjeto", Erl)
End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
    On Error GoTo QuitarSta_Err
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    If UserList(UserIndex).Stats.MinSta = 0 Then Exit Sub
    Call WriteUpdateSta(UserIndex)
    Exit Sub
QuitarSta_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.QuitarSta", Erl)
End Sub

Public Sub DoRaices(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo ErrHandler
    Dim Suerte As Integer
    Dim res    As Integer
    With UserList(UserIndex)
        If .flags.Privilegios And (e_PlayerType.Consejero) Then
            Exit Sub
        End If
        If .Stats.MinSta > 2 Then
            Call QuitarSta(UserIndex, 2)
        Else
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            ' Msg650=Estás muy cansado para obtener raices.
            Call WriteLocaleMsg(UserIndex, 650, e_FontTypeNames.FONTTYPE_INFO)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If
        Dim Skill As Integer
        Skill = .Stats.UserSkills(e_Skill.Alquimia)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        res = RandomNumber(1, Suerte)
        If res < 7 Then
            Dim nPos  As t_WorldPos
            Dim MiObj As t_Obj
            Call ActualizarRecurso(.pos.Map, x, y)
            MiObj.amount = RandomNumber(5, 7)
            MiObj.amount = Round(MiObj.amount * 2.5 * SvrConfig.GetValue("RecoleccionMult"))
            MiObj.ObjIndex = Raices
            MapData(.pos.Map, x, y).ObjInfo.amount = MapData(.pos.Map, x, y).ObjInfo.amount - MiObj.amount
            If MapData(.pos.Map, x, y).ObjInfo.amount < 0 Then
                MapData(.pos.Map, x, y).ObjInfo.amount = 0
            End If
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.pos, MiObj)
            End If
            Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.charindex, vbWhite)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .pos.x, .pos.y))
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .pos.x, .pos.y))
        End If
        Call SubirSkill(UserIndex, e_Skill.Alquimia)
        .Counters.Trabajando = .Counters.Trabajando + 1
        .Counters.LastTrabajo = Int(IntervaloTrabajarExtraer / 1000)
        If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
            Call WriteMacroTrabajoToggle(UserIndex, True)
        End If
    End With
    Exit Sub
ErrHandler:
    Call LogError("Error en DoRaices")
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal ObjetoDorado As Boolean = False)
    On Error GoTo ErrHandler
    Dim Suerte As Integer
    Dim res    As Integer
    With UserList(UserIndex)
        If .flags.Privilegios And (e_PlayerType.Consejero) Then
            Exit Sub
        End If
        'EsfuerzoTalarLeñador = 1
        If .Stats.MinSta > 5 Then
            Call QuitarSta(UserIndex, 5)
        Else
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If
        Dim Skill As Integer
        Skill = .Stats.UserSkills(e_Skill.Talar)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        'HarThaoS: Le agrego más dificultad al talar en zona segura.  37% probabilidad de fallo en segura vs 16% en insegura
        res = RandomNumber(1, IIf(MapInfo(UserList(UserIndex).pos.Map).Seguro = 1, Suerte + 4, Suerte))
        'ReyarB: aumento chances solamente si es el arbol de pino nudoso.
        If ObjData(MapData(.pos.Map, x, y).ObjInfo.ObjIndex).Pino = 1 Then
            res = 1
            Suerte = 100
        End If
        '118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
        If res < 6 Then
            Dim nPos  As t_WorldPos
            Dim MiObj As t_Obj
            Call ActualizarRecurso(.pos.Map, x, y)
            MapData(.pos.Map, x, y).ObjInfo.data = GetTickCountRaw() ' Ultimo uso
            If .clase = Trabajador Then
                MiObj.amount = GetExtractResourceForLevel(.Stats.ELV)
            Else
                MiObj.amount = RandomNumber(1, 2)
            End If
            MiObj.amount = MiObj.amount * SvrConfig.GetValue("RecoleccionMult")
            If ObjData(MapData(.pos.Map, x, y).ObjInfo.ObjIndex).Elfico = 1 Then
                MiObj.ObjIndex = ElvenWood
            ElseIf ObjData(MapData(.pos.Map, x, y).ObjInfo.ObjIndex).Pino = 1 Then
                MiObj.ObjIndex = PinoWood
            Else
                MiObj.ObjIndex = Wood
            End If
            If MiObj.amount > MapData(.pos.Map, x, y).ObjInfo.amount Then
                MiObj.amount = MapData(.pos.Map, x, y).ObjInfo.amount
            End If
            MapData(.pos.Map, x, y).ObjInfo.amount = MapData(.pos.Map, x, y).ObjInfo.amount - MiObj.amount
            ' AGREGAR FX
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.pos, MiObj)
            End If
            Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.charindex, vbWhite)
            If MapInfo(.pos.Map).Seguro = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(SND_TALAR, .pos.x, .pos.y))
            Else
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .pos.x, .pos.y))
            End If
            ' Al talar también podés dropear cosas raras (se setean desde RecursosEspeciales.dat)
            Dim i As Integer
            ' Por cada drop posible
            For i = 1 To UBound(EspecialesTala)
                ' Tiramos al azar entre 1 y la probabilidad
                res = RandomNumber(1, EspecialesTala(i).data)
                ' Si tiene suerte y le pega
                If res = 1 Then
                    MiObj.ObjIndex = EspecialesTala(i).ObjIndex
                    MiObj.amount = 1 ' Solo un item por vez
                    ' Tiro siempre el item al piso, me parece más rolero, como que cae del árbol :P
                    Call TirarItemAlPiso(.pos, MiObj)
                End If
            Next i
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(64, .pos.x, .pos.y))
        End If
        Call SubirSkill(UserIndex, e_Skill.Talar)
        .Counters.Trabajando = .Counters.Trabajando + 1
        .Counters.LastTrabajo = Int(IntervaloTrabajarExtraer / 1000)
        If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
            Call WriteMacroTrabajoToggle(UserIndex, True)
        End If
    End With
    Exit Sub
ErrHandler:
    Call LogError("Error en DoTalar")
End Sub

Public Sub DoMineria(ByVal UserIndex As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal ObjetoDorado As Boolean = False)
    On Error GoTo ErrHandler
    Dim Suerte     As Integer
    Dim res        As Integer
    Dim Metal      As Integer
    Dim Yacimiento As t_ObjData
    With UserList(UserIndex)
        If .flags.Privilegios And (e_PlayerType.Consejero) Then
            Exit Sub
        End If
        If .Stats.MinSta > 5 Then
            Call QuitarSta(UserIndex, 5)
        Else
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If
        Dim Skill As Integer
        Skill = .Stats.UserSkills(e_Skill.Mineria)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        'HarThaoS: Le agrego más dificultad al talar en zona segura.  37% probabilidad de fallo en segura vs 16% en insegura
        res = RandomNumber(1, IIf(MapInfo(UserList(UserIndex).pos.Map).Seguro = 1, Suerte + 2, Suerte))
        '118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
        'ReyarB: aumento chances solamente si es mineria de blodium.
        If ObjData(MapData(.pos.Map, x, y).ObjInfo.ObjIndex).MineralIndex = 3787 Then
            res = 1
            Suerte = 100
        End If
        If res <= 5 Then
            Dim MiObj As t_Obj
            Dim nPos  As t_WorldPos
            Call ActualizarRecurso(.pos.Map, x, y)
            MapData(.pos.Map, x, y).ObjInfo.data = GetTickCountRaw() ' Ultimo uso
            Yacimiento = ObjData(MapData(.pos.Map, x, y).ObjInfo.ObjIndex)
            MiObj.ObjIndex = Yacimiento.MineralIndex
            If .clase = Trabajador Then
                MiObj.amount = GetExtractResourceForLevel(.Stats.ELV)
            Else
                MiObj.amount = RandomNumber(1, 2)
            End If
            MiObj.amount = MiObj.amount * SvrConfig.GetValue("RecoleccionMult")
            If MiObj.amount > MapData(.pos.Map, x, y).ObjInfo.amount Then
                MiObj.amount = MapData(.pos.Map, x, y).ObjInfo.amount
            End If
            MapData(.pos.Map, x, y).ObjInfo.amount = MapData(.pos.Map, x, y).ObjInfo.amount - MiObj.amount
            If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.pos, MiObj)
            ' AGREGAR FX
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
            Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.charindex, vbWhite)
            ' Msg651=¡Has extraído algunos minerales!
            Call WriteLocaleMsg(UserIndex, 651, e_FontTypeNames.FONTTYPE_INFO)
            If MapInfo(.pos.Map).Seguro = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(15, .pos.x, .pos.y))
            Else
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(15, .pos.x, .pos.y))
            End If
            ' Al minar también puede dropear una gema
            Dim i As Integer
            ' Por cada drop posible
            For i = 1 To Yacimiento.CantItem
                ' Tiramos al azar entre 1 y la probabilidad
                res = RandomNumber(1, Yacimiento.Item(i).amount)
                ' Si tiene suerte y le pega
                If res = 1 Then
                    ' Se lo metemos al inventario (o lo tiramos al piso)
                    MiObj.ObjIndex = Yacimiento.Item(i).ObjIndex
                    MiObj.amount = 1 ' Solo una gema por vez
                    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.pos, MiObj)
                    ' Le mandamos un mensaje
                    Call WriteLocaleMsg(UserIndex, 1465, e_FontTypeNames.FONTTYPE_INFO)  ' Msg1465=¡Has conseguido ¬1!
                End If
            Next
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(2185, .pos.x, .pos.y))
        End If
        Call SubirSkill(UserIndex, e_Skill.Mineria)
        .Counters.Trabajando = .Counters.Trabajando + 1
        .Counters.LastTrabajo = Int(IntervaloTrabajarExtraer / 1000)
        If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
            Call WriteMacroTrabajoToggle(UserIndex, True)
        End If
    End With
    Exit Sub
ErrHandler:
    Call LogError("Error en Sub DoMineria")
End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
    On Error GoTo DoMeditar_Err
    Dim Mana As Long
    With UserList(UserIndex)
        .Counters.TimerMeditar = .Counters.TimerMeditar + 1
        .Counters.TiempoInicioMeditar = .Counters.TiempoInicioMeditar + 1
        If .Counters.TimerMeditar >= IntervaloMeditar And .Counters.TiempoInicioMeditar > 20 Then
            If e_Class.Bard And .invent.EquippedRingAccesoryObjIndex = CommonLuteIndex Then
                Mana = Porcentaje(.Stats.MaxMAN, Porcentaje(PorcentajeRecuperoMana, RecoveryMana + .Stats.UserSkills(e_Skill.Meditar) * MultiplierManaxSkills)) + ManaCommonLute
            ElseIf e_Class.Bard And .invent.EquippedRingAccesoryObjIndex = MagicLuteIndex Then
                Mana = Porcentaje(.Stats.MaxMAN, Porcentaje(PorcentajeRecuperoMana, RecoveryMana + .Stats.UserSkills(e_Skill.Meditar) * MultiplierManaxSkills)) + ManaMagicLute
            ElseIf e_Class.Bard And .invent.EquippedRingAccesoryObjIndex = ElvenLuteIndex Then
                Mana = Porcentaje(.Stats.MaxMAN, Porcentaje(PorcentajeRecuperoMana, RecoveryMana + .Stats.UserSkills(e_Skill.Meditar) * MultiplierManaxSkills)) + ManaElvenLute
            Else
                Mana = Porcentaje(.Stats.MaxMAN, Porcentaje(PorcentajeRecuperoMana, RecoveryMana + .Stats.UserSkills(e_Skill.Meditar) * MultiplierManaxSkills))
            End If
            If Mana <= 0 Then Mana = 1
            If .Stats.MinMAN + Mana >= .Stats.MaxMAN Then
                .Stats.MinMAN = .Stats.MaxMAN
                .flags.Meditando = False
                .Char.FX = 0
                Call WriteUpdateMana(UserIndex)
                Call SubirSkill(UserIndex, Meditar)
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
            Else
                .Stats.MinMAN = .Stats.MinMAN + Mana
                Call WriteUpdateMana(UserIndex)
                Call SubirSkill(UserIndex, Meditar)
            End If
            .Counters.TimerMeditar = 0
        End If
    End With
    Exit Sub
DoMeditar_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.DoMeditar", Erl)
End Sub

Public Sub DoMontar(ByVal UserIndex As Integer, ByRef Montura As t_ObjData, ByVal Slot As Integer)
    On Error GoTo DoMontar_Err
    With UserList(UserIndex)
        If PuedeUsarObjeto(UserIndex, .invent.Object(Slot).ObjIndex, True) > 0 Then
            Exit Sub
        End If
        If .flags.Montado = 0 And .Counters.EnCombate > 0 Then
            Call WriteLocaleMsg(UserIndex, 1466, e_FontTypeNames.FONTTYPE_INFOBOLD, .Counters.EnCombate)  ' Msg1466=Estás en combate, debes aguardar ¬1 segundo(s) para montar...
            Exit Sub
        End If
        If .flags.EnReto Then
            ' Msg652=No podés montar en un reto.
            Call WriteLocaleMsg(UserIndex, 652, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Montado = 0 And (MapData(.pos.Map, .pos.x, .pos.y).trigger > 10) Then
            ' Msg653=No podés montar aquí.
            Call WriteLocaleMsg(UserIndex, 653, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
            ' Msg654=Pierdes el efecto del mimetismo.
            Call WriteLocaleMsg(UserIndex, 654, e_FontTypeNames.FONTTYPE_INFO)
            .Counters.Mimetismo = 0
            .flags.Mimetizado = e_EstadoMimetismo.Desactivado
            Call RefreshCharStatus(UserIndex)
        End If
        ' Si está oculto o invisible, hago que pueda montar pero se haga visible
        If (.flags.Oculto = 1 Or .flags.invisible = 1) And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .flags.invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.DisabledInvisibility = 0
            Call WriteLocaleMsg(UserIndex, 307, e_FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        End If
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
        End If
        If .flags.Montado = 1 And .invent.EquippedSaddleObjIndex > 0 Then
            If ObjData(.invent.EquippedSaddleObjIndex).ResistenciaMagica > 0 Then
                Call UpdateUserInv(False, UserIndex, .invent.EquippedSaddleSlot)
            End If
        End If
        .invent.EquippedSaddleObjIndex = .invent.Object(Slot).ObjIndex
        .invent.EquippedSaddleSlot = Slot
        If .flags.Montado = 0 Then
            .Char.body = Montura.Ropaje
            .Char.head = .OrigChar.head
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = .Char.CascoAnim
            .Char.CartAnim = NoCart
            .flags.Montado = 1
            Call TargetUpdateTerrain(.EffectOverTime)
        Else
            .flags.Montado = 0
            .Char.head = .OrigChar.head
            Call TargetUpdateTerrain(.EffectOverTime)
            If .invent.EquippedArmorObjIndex > 0 Then
                .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.EquippedArmorObjIndex))
            Else
                Call SetNakedBody(UserList(UserIndex))
            End If
            If .invent.EquippedShieldObjIndex > 0 Then .Char.ShieldAnim = ObjData(.invent.EquippedShieldObjIndex).ShieldAnim
            If .invent.EquippedWeaponObjIndex > 0 Then .Char.WeaponAnim = ObjData(.invent.EquippedWeaponObjIndex).WeaponAnim
            If .invent.EquippedHelmetObjIndex > 0 Then .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
            If .invent.EquippedAmuletAccesoryObjIndex > 0 Then
                If ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje > 0 Then .Char.CartAnim = ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje
            End If
        End If
        Call ActualizarVelocidadDeUsuario(UserIndex)
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteEquiteToggle(UserIndex)
    End With
    Exit Sub
DoMontar_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.DoMontar", Erl)
End Sub

Public Sub ActualizarRecurso(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo ActualizarRecurso_Err
    Dim ObjIndex As Integer
    ObjIndex = MapData(Map, x, y).ObjInfo.ObjIndex
    Dim TiempoActual As Long
    TiempoActual = GetTickCountRaw()
    ' Data = Ultimo uso
    Dim lastUse As Long
    lastUse = MapData(Map, x, y).ObjInfo.data
    If lastUse <> &H7FFFFFFF Then
        Dim elapsedMs As Double
        elapsedMs = TicksElapsed(lastUse, TiempoActual)
        If elapsedMs / 1000# > ObjData(ObjIndex).TiempoRegenerar Then
            MapData(Map, x, y).ObjInfo.amount = ObjData(ObjIndex).VidaUtil
            MapData(Map, x, y).ObjInfo.data = &H7FFFFFFF   ' Ultimo uso = Max Long
        End If
    End If
    Exit Sub
ActualizarRecurso_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ActualizarRecurso", Erl)
End Sub

Public Function ObtenerPezRandom(ByVal PoderCania As Integer) As Long
    On Error GoTo ObtenerPezRandom_Err

    Dim PesoMinimo As Long
    Dim PesoMaximo As Long
    Dim ValorGenerado As Long
    Dim PezIndex As Long

    ' Aseguramos que PoderCania esté dentro del rango válido del array.
    PoderCania = Clamp(PoderCania, LBound(PesoPeces), UBound(PesoPeces))
    
    ' PesoMaximo: suma de pesos acumulados de todos los peces que puede pescar esta caña
    PesoMaximo = PesoPeces(PoderCania)
    
    ' Esto asegura que el aleatorio solo considere los peces que pertenecen al Power actual
    If PoderCania > LBound(PesoPeces) Then
        PesoMinimo = PesoPeces(PoderCania - 1)
    Else
        PesoMinimo = 0
    End If

    ' Generamos un valor aleatorio solo dentro del rango correspondiente
    If PesoMaximo <= PesoMinimo Then
        ValorGenerado = RandomNumber(0, PesoMaximo - 1)
    Else
        ValorGenerado = RandomNumber(PesoMinimo, PesoMaximo - 1)
    End If

    ' Obtenemos el pez correspondiente
    PezIndex = BinarySearchPeces(ValorGenerado) ' BinarySearchPeces() espera un valor en el mismo espacio acumulado que PesoPeces().
    ObtenerPezRandom = Peces(PezIndex).ObjIndex

    Exit Function

ObtenerPezRandom_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ObtenerPezRandom", Erl)
End Function

Function ModDomar(ByVal clase As e_Class) As Integer
    On Error GoTo ModDomar_Err
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case clase
        Case e_Class.Druid
            ModDomar = 6
        Case e_Class.Hunter
            ModDomar = 6
        Case e_Class.Cleric
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
    Exit Function
ModDomar_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ModDomar", Erl)
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
    On Error GoTo FreeMascotaIndex_Err
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/03/09
    '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
    '***************************************************
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
    FreeMascotaIndex = -1
    Exit Function
FreeMascotaIndex_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.FreeMascotaIndex", Erl)
End Function

Private Function HayEspacioMascotas(ByVal UserIndex As Integer) As Boolean
    HayEspacioMascotas = (FreeMascotaIndex(UserIndex) > 0)
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    On Error GoTo ErrHandler
    Dim puntosDomar As Integer
    Dim CanStay     As Boolean
    Dim petType     As Integer
    Dim NroPets     As Integer
    If IsValidUserRef(NpcList(NpcIndex).MaestroUser) And NpcList(NpcIndex).MaestroUser.ArrayIndex = UserIndex Then
        ' Msg655=Ya domaste a esa criatura.
        Call WriteLocaleMsg(UserIndex, 655, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    With UserList(UserIndex)
        If .flags.Privilegios And e_PlayerType.Consejero Then Exit Sub
        If .NroMascotas < MAXMASCOTAS And HayEspacioMascotas(UserIndex) Then
            If IsValidNpcRef(NpcList(NpcIndex).MaestroNPC) > 0 Or IsValidUserRef(NpcList(NpcIndex).MaestroUser) Then
                ' Msg656=La criatura ya tiene amo.
                Call WriteLocaleMsg(UserIndex, 656, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            puntosDomar = CInt(.Stats.UserAtributos(e_Atributos.Carisma)) * CInt(.Stats.UserSkills(e_Skill.Domar))
            If .clase = e_Class.Druid Then
                puntosDomar = puntosDomar / 6 'original es 6
            Else
                puntosDomar = puntosDomar / 118 'para que solo el druida dome
            End If
            'No tiene nivel suficiente?
            If NpcList(NpcIndex).MinTameLevel > .Stats.ELV Then
                ' Msg1321=Debes ser nivel ¬1 o superior para domar esta criatura.
                Call WriteLocaleMsg(UserIndex, 1321, e_FontTypeNames.FONTTYPE_INFO, NpcList(NpcIndex).MinTameLevel)
                Exit Sub
            End If
            If NpcList(NpcIndex).flags.Domable <= puntosDomar And RandomNumber(1, 5) = 1 Then
                Dim Index As Integer
                .NroMascotas = .NroMascotas + 1
                Index = FreeMascotaIndex(UserIndex)
                Call SetNpcRef(.MascotasIndex(Index), NpcIndex)
                .MascotasType(Index) = NpcList(NpcIndex).Numero
                Call SetUserRef(NpcList(NpcIndex).MaestroUser, UserIndex)
                .flags.ModificoMascotas = True
                Call FollowAmo(NpcIndex)
                Call ReSpawnNpc(NpcList(NpcIndex))
                ' Msg657=La criatura te ha aceptado como su amo.
                Call WriteLocaleMsg(UserIndex, 657, e_FontTypeNames.FONTTYPE_INFO)
                ' Es zona segura?
                If MapInfo(.pos.Map).NoMascotas = 1 Then
                    petType = NpcList(NpcIndex).Numero
                    NroPets = .NroMascotas
                    Call QuitarNPC(NpcIndex, eNewPet)
                    .MascotasType(Index) = petType
                    .NroMascotas = NroPets
                    ' Msg658=No se permiten mascotas en zona segura. estas te esperaran afuera.
                    Call WriteLocaleMsg(UserIndex, 658, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If Not .flags.UltimoMensaje = 5 Then
                    ' Msg659=No has logrado domar la criatura.
                    Call WriteLocaleMsg(UserIndex, 659, e_FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 5
                End If
            End If
            Call SubirSkill(UserIndex, e_Skill.Domar)
        Else
            ' Msg660=No puedes controlar mas criaturas.
            Call WriteLocaleMsg(UserIndex, 660, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.Description)
End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    On Error GoTo PuedeDomarMascota_Err
    '***************************************************
    'Author: ZaMa
    'This function checks how many NPCs of the same type have
    'been tamed by the user.
    'Returns True if that amount is less than two.
    '***************************************************
    Dim i           As Long
    Dim numMascotas As Long
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) = NpcList(NpcIndex).Numero Then
            numMascotas = numMascotas + 1
        End If
    Next i
    If numMascotas <= 1 Then PuedeDomarMascota = True
    Exit Function
PuedeDomarMascota_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeDomarMascota", Erl)
End Function

Public Function EntregarPezEspecial(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .flags.PescandoEspecial Then
            Dim obj As t_Obj
            obj.amount = 1
            obj.ObjIndex = .Stats.NumObj_PezEspecial
            If Not MeterItemEnInventario(UserIndex, obj) Then
                .Stats.NumObj_PezEspecial = 0
                .flags.PescandoEspecial = False
                Exit Function
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(obj.ObjIndex).GrhIndex))
            'Msg922=Felicitaciones has pescado un pez de gran porte ( " & ObjData(obj.ObjIndex).name & " )
            Call WriteLocaleMsg(UserIndex, 922, e_FontTypeNames.FONTTYPE_FIGHT, ObjData(obj.ObjIndex).name)
            .Stats.NumObj_PezEspecial = 0
            .flags.PescandoEspecial = False
        End If
    End With
End Function

Public Sub FishOrThrowNet(ByVal UserIndex As Integer)
    On Error GoTo FishOrThrowNet_Err:
    With UserList(UserIndex)
        If ObjData(.invent.EquippedWorkingToolObjIndex).OBJType <> e_OBJType.otWorkingTools Then Exit Sub
        If ObjData(.invent.EquippedWorkingToolObjIndex).Subtipo = e_ToolsSubtype.eFishingNet Then
            If MapInfo(.pos.Map).Seguro = 1 Or Not ExpectObjectTypeAt(e_OBJType.otFishingPool, .pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y) Then
                If IsValidUserRef(.flags.TargetUser) Or IsValidNpcRef(.flags.TargetNPC) Then
                    ThrowNetToTarget (UserIndex)
                    Call WriteWorkRequestTarget(UserIndex, 0)
                    Exit Sub
                End If
            End If
        End If
        Call Trabajar(UserIndex, e_Skill.Pescar)
    End With
    Exit Sub
FishOrThrowNet_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.FishOrThrowNet", Erl)
End Sub

Sub ThrowNetToTarget(ByVal UserIndex As Integer)
    On Error GoTo ThrowNetToTarget_Err:
    With UserList(UserIndex)
        If .invent.EquippedWorkingToolObjIndex = 0 Then Exit Sub
        If ObjData(.invent.EquippedWorkingToolObjIndex).OBJType <> e_OBJType.otWorkingTools Then Exit Sub
        If ObjData(.invent.EquippedWorkingToolObjIndex).Subtipo <> e_ToolsSubtype.eFishingNet Then Exit Sub
        'If it's outside range log it and exit
        If Abs(.pos.x - .Trabajo.Target_X) > RANGO_VISION_X Or Abs(.pos.y - .Trabajo.Target_Y) > RANGO_VISION_Y Then
            Call LogSecurity("Ataque fuera de rango de " & .name & "(" & .pos.Map & "/" & .pos.x & "/" & .pos.y & ") ip: " & .ConnectionDetails.IP & " a la posicion (" & _
                    .pos.Map & "/" & .Trabajo.Target_X & "/" & .Trabajo.Target_Y & ")")
            Exit Sub
        End If
        'Check bow's interval
        If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
        'Check attack-spell interval
        If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
        'Check Magic interval
        If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
        'check item cd
        Dim ThrowNet As Boolean
        ThrowNet = False
        If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then
            Dim tU As Integer
            tU = UserList(UserIndex).flags.TargetUser.ArrayIndex
            If UserIndex = tU Then
                Call WriteLocaleMsg(UserIndex, MsgCantAttackYourself, e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            If IsSet(UserList(tU).flags.StatusMask, eCCInmunity) Then
                Call WriteLocaleMsg(UserIndex, MsgCCInunity, e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            If Not UserMod.CanMove(UserList(tU).flags, UserList(tU).Counters) Then
                ' Msg661=No podes inmovilizar un objetivo que no puede moverse.
                Call WriteLocaleMsg(UserIndex, 661, e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            UserList(tU).Counters.Inmovilizado = NET_INMO_DURATION
            If UserList(tU).flags.Inmovilizado = 0 Then
                UserList(tU).flags.Inmovilizado = 1
                Call SendData(SendTarget.ToPCAliveArea, tU, PrepareMessageCreateFX(UserList(tU).Char.charindex, FISHING_NET_FX, 0, UserList(tU).pos.x, UserList(tU).pos.y))
                Call WriteInmovilizaOK(tU)
                Call WritePosUpdate(tU)
                ThrowNet = True
            End If
            Call SetUserRef(UserList(UserIndex).flags.TargetUser, 0)
        ElseIf IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
            Dim NpcIndex As Integer
            NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
            If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
                Dim UserAttackInteractionResult As t_AttackInteractionResult
                UserAttackInteractionResult = UserCanAttackNpc(UserIndex, NpcIndex)
                Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
                If UserAttackInteractionResult.CanAttack Then
                    If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
                Else
                    Exit Sub
                End If
                Call NPCAtacado(NpcIndex, UserIndex)
                NpcList(NpcIndex).flags.Inmovilizado = 1
                NpcList(NpcIndex).Contadores.Inmovilizado = (NET_INMO_DURATION * 6.5) * 6
                NpcList(NpcIndex).flags.Paralizado = 0
                NpcList(NpcIndex).Contadores.Paralisis = 0
                Call AnimacionIdle(NpcIndex, True)
                ThrowNet = True
                Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageFxPiso(FISHING_NET_FX, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y))
                Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
            Else
                Call WriteLocaleMsg(UserIndex, MSgNpcInmuneToEffect, e_FontTypeNames.FONTTYPE_INFOIAO)
            End If
        End If
        If ThrowNet Then
            Call UpdateCd(UserIndex, ObjData(.invent.EquippedWorkingToolObjIndex).cdType)
            Call QuitarUserInvItem(UserIndex, .invent.EquippedWorkingToolSlot, 1)
            Call UpdateUserInv(True, UserIndex, .invent.EquippedWorkingToolSlot)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, .Trabajo.Target_X, _
                    .Trabajo.Target_Y, 3))
        End If
    End With
    Exit Sub
ThrowNetToTarget_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ThrowNetToTarget", Erl)
End Sub

Public Function GetExtractResourceForLevel(ByVal level As Integer) As Integer
    Dim upper As Long
    Dim lower As Long
    lower = Int(CDbl(level + 0.000001) / 3.6)
    upper = Int(CDbl(level + 0.000001) / 2)
    GetExtractResourceForLevel = RandomNumber(lower, upper)
End Function

Public Function GiveExpWhileWorking(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal JobType As Byte)
    On Error GoTo GiveExpWhileWorking_Err:
    Dim tmpExp As Byte
    Select Case JobType
        Case e_JobsTypes.Miner
            tmpExp = SvrConfig.GetValue("MiningExp")
        Case e_JobsTypes.Woodcutter
            tmpExp = SvrConfig.GetValue("FellingExp")
        Case e_JobsTypes.Blacksmith
            tmpExp = SvrConfig.GetValue("ForgingExp")
        Case e_JobsTypes.Carpenter
            tmpExp = SvrConfig.GetValue("CarpentryExp")
        Case e_JobsTypes.Woodcutter
            tmpExp = SvrConfig.GetValue("FellingExp")
        Case e_JobsTypes.Fisherman
            If ObjData(ItemIndex).Power >= 2 Then
                tmpExp = SvrConfig.GetValue("FishingExp")
            End If
        Case e_JobsTypes.Alchemist
            tmpExp = SvrConfig.GetValue("MixingExp")
        Case Else
            tmpExp = SvrConfig.GetValue("ElseExp")
    End Select
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + tmpExp
    Exit Function
GiveExpWhileWorking_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.GiveExpWhileWorking", Erl)
End Function

Public Function KnowsCraftingRecipe(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    KnowsCraftingRecipe = True
    Dim hIndex As Integer
    hIndex = ObjData(ItemIndex).Hechizo
    'item doesnt require recipe
    If hIndex = 0 Then
        Exit Function
    End If
    If Not TieneHechizo(hIndex, UserIndex) Then
        'Msg644=Lamentablemente no aprendiste la receta para crear este item.
        Call WriteLocaleMsg(UserIndex, 644, e_FontTypeNames.FONTTYPE_INFOBOLD)
        KnowsCraftingRecipe = False
        Exit Function
    End If
End Function
