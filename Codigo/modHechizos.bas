Attribute VB_Name = "modHechizos"
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
Private Const FLAUTA_ELFICA As Long = 40

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer, Optional ByVal IgnoreVisibilityCheck As Boolean = False)
    On Error GoTo NpcLanzaSpellSobreUser_Err
    Dim Damage    As Integer
    Dim DamageStr As String
    If Spell = 0 Then Exit Sub
    Dim IsAlive As Boolean
    IsAlive = True
    With UserList(UserIndex)
        If .flags.Muerto Then Exit Sub
        '¿NPC puede ver a través de la invisibilidad?
        If Not IgnoreVisibilityCheck Then
            If .flags.invisible = 1 Or .flags.Oculto = 1 Or .flags.Inmunidad = 1 Then Exit Sub
        End If
        NpcList(NpcIndex).Contadores.IntervaloLanzarHechizo = GetTickCountRaw()
        If Hechizos(Spell).Tipo = uPhysicalSkill Then
            If Not HandlePhysicalSkill(NpcIndex, eNpc, UserIndex, eUser, Spell, IsAlive) Then
                Exit Sub
            End If
        End If
        Call InfoHechizoDeNpcSobreUser(NpcIndex, UserIndex, Spell)
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.eDoHeal) Then
            Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
            Damage = Damage * NPCs.GetMagicHealingBonus(NpcList(NpcIndex))
            Damage = Damage * UserMod.GetSelfHealingBonus(UserList(UserIndex))
            If Damage > 0 Then
                Call UserMod.DoDamageOrHeal(UserIndex, NpcIndex, eNpc, Damage, e_DamageSourceType.e_magic, Spell)
                DamageStr = PonerPuntos(Damage)
                Call WriteLocaleMsg(UserIndex, 32, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name & "¬" & DamageStr)
            End If
        ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.eDoDamage) Then
            .Counters.EnCombate = IntervaloEnCombate
            Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
            Damage = Damage * (1 + NpcList(NpcIndex).Stats.MagicBonus)
            ' Si el hechizo no ignora la RM
            If Hechizos(Spell).AntiRm = 0 Then
                Dim PorcentajeRM As Integer
                PorcentajeRM = GetUserMRForNpc(UserIndex)
                ' Resto el porcentaje total
                Damage = Damage - Porcentaje(Damage, PorcentajeRM)
            End If
            Damage = Damage * NPCs.GetMagicDamageModifier(NpcList(NpcIndex))
            Damage = Damage * UserMod.GetMagicDamageReduction(UserList(UserIndex))
            If Damage < 0 Then Damage = 0
            IsAlive = UserMod.DoDamageOrHeal(UserIndex, NpcIndex, eNpc, -Damage, e_DamageSourceType.e_magic, Spell) = eStillAlive
            DamageStr = PonerPuntos(Damage)
            Call WriteLocaleMsg(UserIndex, 1627, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name & "¬" & DamageStr) 'Msg1627=¬1 te ha quitado ¬2 puntos de vida.
            Call SubirSkill(UserIndex, Resistencia)
            If NpcList(NpcIndex).Char.CastAnimation > 0 Then Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageCharAtaca(NpcList(NpcIndex).Char.charindex, _
                    .Char.charindex, DamageStr))
            
        End If
        If IsAlive Then
            Dim Effect As IBaseEffectOverTime
            If Hechizos(Spell).EotId > 0 Then
                Set Effect = FindEffectOnTarget(NpcIndex, .EffectOverTime, Hechizos(Spell).EotId)
                If Effect Is Nothing Then
                    Call CreateEffect(NpcIndex, eNpc, UserIndex, eUser, Hechizos(Spell).EotId)
                Else
                    Call Effect.Reset(NpcIndex, eNpc, Hechizos(Spell).EotId)
                End If
            End If
        End If
        'Mana
        If Hechizos(Spell).SubeMana = 1 Then
            Damage = RandomNumber(Hechizos(Spell).MinMana, Hechizos(Spell).MaxMana)
            .Stats.MinMAN = MinimoInt(.Stats.MinMAN + Damage, .Stats.MaxMAN)
            Call WriteUpdateMana(UserIndex)
            Call WriteLocaleMsg(UserIndex, 1628, e_FontTypeNames.FONTTYPE_INFO, NpcList(NpcIndex).name & "¬" & Damage) 'Msg1628=¬1 te ha restaurado ¬2 puntos de maná.
        ElseIf Hechizos(Spell).SubeMana = 2 Then
            Damage = RandomNumber(Hechizos(Spell).MinMana, Hechizos(Spell).MaxMana)
            .Stats.MinMAN = MaximoInt(.Stats.MinMAN - Damage, 0)
            Call WriteUpdateMana(UserIndex)
            Call WriteLocaleMsg(UserIndex, 1629, e_FontTypeNames.FONTTYPE_INFO, NpcList(NpcIndex).name & "¬" & Damage) 'Msg1629=¬1 te ha quitado ¬2 puntos de maná.
        End If
        If Hechizos(Spell).SubeAgilidad = 1 Then
            Damage = RandomNumber(Hechizos(Spell).MinAgilidad, Hechizos(Spell).MaxAgilidad)
            .flags.TomoPocion = True
            .flags.DuracionEfecto = Hechizos(Spell).Duration
            .Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(.Stats.UserAtributos(e_Atributos.Agilidad) + Damage, .Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)
            Call WriteFYA(UserIndex)
        ElseIf Hechizos(Spell).SubeAgilidad = 2 Then
            Damage = RandomNumber(Hechizos(Spell).MinAgilidad, Hechizos(Spell).MaxAgilidad)
            .flags.TomoPocion = True
            .flags.DuracionEfecto = Hechizos(Spell).Duration
            .Stats.UserAtributos(e_Atributos.Agilidad) = MaximoInt(MINATRIBUTOS, .Stats.UserAtributos(e_Atributos.Agilidad) - Damage)
            Call WriteFYA(UserIndex)
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveDebuff) Then
            Dim NegativeEffect As IBaseEffectOverTime
            Set NegativeEffect = EffectsOverTime.FindEffectOfTypeOnTarget(.EffectOverTime, eDebuff)
            If Not NegativeEffect Is Nothing Then
                NegativeEffect.RemoveMe = True
                Exit Sub
            End If
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.StealBuff) Then
            Dim TargetBuff As IBaseEffectOverTime
            Set TargetBuff = EffectsOverTime.FindEffectOfTypeOnTarget(.EffectOverTime, eBuff)
            If Not TargetBuff Is Nothing Then
                Call EffectsOverTime.ChangeOwner(UserIndex, eUser, NpcIndex, eNpc, TargetBuff)
            End If
        End If
        If Hechizos(Spell).SubeFuerza = 1 Then
            Damage = RandomNumber(Hechizos(Spell).MinFuerza, Hechizos(Spell).MaxFuerza)
            .flags.TomoPocion = True
            .flags.DuracionEfecto = Hechizos(Spell).Duration
            .Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(.Stats.UserAtributos(e_Atributos.Fuerza) + Damage, .Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)
            Call WriteFYA(UserIndex)
        ElseIf Hechizos(Spell).SubeFuerza = 2 Then
            Damage = RandomNumber(Hechizos(Spell).MinFuerza, Hechizos(Spell).MaxFuerza)
            .flags.TomoPocion = True
            .flags.DuracionEfecto = Hechizos(Spell).Duration
            .Stats.UserAtributos(e_Atributos.Fuerza) = MaximoInt(MINATRIBUTOS, .Stats.UserAtributos(e_Atributos.Fuerza) - Damage)
            Call WriteFYA(UserIndex)
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.Paralize) Then
            If .flags.Paralizado = 0 Then
                .flags.Paralizado = 1
                .Counters.Paralisis = Hechizos(Spell).Duration / 2
                Call WriteParalizeOK(UserIndex)
                Call WritePosUpdate(UserIndex)
            End If
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.Immobilize) Then
            If .flags.Inmovilizado = 0 Then
                .flags.Inmovilizado = 1
                .Counters.Inmovilizado = Hechizos(Spell).Duration / 2
                Call WriteInmovilizaOK(UserIndex)
                Call WritePosUpdate(UserIndex)
            End If
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveParalysis) Then
            If .flags.Paralizado > 0 Then
                .flags.Paralizado = 0
                .Counters.Paralisis = 0
                Call WriteParalizeOK(UserIndex)
            End If
            If .flags.Inmovilizado > 0 Then
                .flags.Inmovilizado = 0
                .Counters.Inmovilizado = 0
                Call WriteInmovilizaOK(UserIndex)
            End If
            Call WritePosUpdate(UserIndex)
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.Incinerate) Then
            If .flags.Incinerado = 0 Then
                .flags.Incinerado = 1
                .Counters.Incineracion = Hechizos(Spell).Duration
                Call WriteLocaleMsg(UserIndex, 1630, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name) 'Msg1630=Has sido incinerado por ¬1.
            End If
        End If
        If Hechizos(Spell).Envenena > 0 Then
            If .flags.Envenenado = 0 Then
                .flags.Envenenado = Hechizos(Spell).Envenena
                .Counters.Veneno = Hechizos(Spell).Duration
                Call WriteLocaleMsg(UserIndex, 1631, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name) 'Msg1631=Has sido envenenado por ¬1.
            End If
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveInvisibility) Then
            Call UserMod.RemoveInvisibility(UserIndex)
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.Dumb) Then
            If .flags.Estupidez = 0 Then
                .flags.Estupidez = IsSet(Hechizos(Spell).Effects, e_SpellEffects.Dumb)
                .Counters.Estupidez = Hechizos(Spell).Duration
                Call WriteLocaleMsg(UserIndex, 1632, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name) 'Msg1632=Has sido estupidizado por ¬1.
                Call WriteDumb(UserIndex)
            End If
        ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveDumb) Then
            If .flags.Estupidez > 0 Then
                .flags.Estupidez = 0
                .Counters.Estupidez = 0
                Call WriteLocaleMsg(UserIndex, 1633, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name) 'Msg1633=¬1 te removió la estupidez.
                Call WriteDumbNoMore(UserIndex)
            End If
        End If
        If Hechizos(Spell).velocidad > 0 Then
            If .Counters.velocidad = 0 Then
                .flags.VelocidadHechizada = Hechizos(Spell).velocidad
                .Counters.velocidad = Hechizos(Spell).Duration
                Call ActualizarVelocidadDeUsuario(UserIndex)
            End If
        End If
        If NpcList(NpcIndex).Char.CastAnimation > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(NpcList(NpcIndex).Char.charindex, NpcList(NpcIndex).Char.CastAnimation))
        ElseIf NpcList(NpcIndex).Char.WeaponAnim > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageArmaMov(NpcList(NpcIndex).Char.charindex, 0))
        End If
    End With
    With NpcList(NpcIndex)
        If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eDisplayCastMessage) Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageChatOverHead("PMAG*" & Spell, .Char.charindex, vbCyan, True, .pos.x, .pos.y, _
                    RequiredSpellDisplayTime, MaxInvisibleSpellDisplayTime))
        End If
        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteConsoleMsg(UserIndex, "HecMSGA*" & Spell & "*" & .name, e_FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With
    Exit Sub
NpcLanzaSpellSobreUser_Err:
    Call TraceError(Err.Number, Err.Description & " Hechizo: " & Spell, "modHechizos.NpcLanzaSpellSobreUser", Erl)
End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
    On Error GoTo NpcLanzaSpellSobreNpc_Err
    Dim Damage    As Integer
    Dim DamageStr As String
    Dim IsAlive   As Boolean
    IsAlive = True
    If Hechizos(Spell).Tipo = uPhysicalSkill Then
        If Not HandlePhysicalSkill(NpcIndex, eNpc, TargetNPC, eNpc, Spell, IsAlive) Then
            Exit Sub
        End If
    End If
    With NpcList(TargetNPC)
        .Contadores.IntervaloLanzarHechizo = GetTickCountRaw()
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.eDoHeal) Then ' Cura
            Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
            Damage = Damage * NPCs.GetMagicHealingBonus(NpcList(NpcIndex))
            Damage = Damage * NPCs.GetSelfHealingBonus(NpcList(TargetNPC))
            DamageStr = PonerPuntos(Damage)
            If Hechizos(Spell).wav > 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .pos.x, .pos.y))
            End If
            If Hechizos(Spell).FXgrh > 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
            End If
            If Damage > 0 Then
                Call SendData(SendTarget.ToPCAliveArea, TargetNPC, PrepareMessageTextCharDrop(DamageStr, .Char.charindex, vbGreen))
            End If
            Call NPCs.DoDamageOrHeal(TargetNPC, NpcIndex, eNpc, Damage, e_DamageSourceType.e_magic, Spell)
            Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageNpcUpdateHP(TargetNPC))
        ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.eDoDamage) Then
            Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
            Damage = Damage * NPCs.GetMagicDamageModifier(NpcList(NpcIndex))
            Damage = Damage * NPCs.GetMagicDamageReduction(NpcList(TargetNPC))
            Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .pos.x, .pos.y))
            Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
            IsAlive = NPCs.DoDamageOrHeal(TargetNPC, NpcIndex, eNpc, -Damage, e_DamageSourceType.e_magic, Spell) = eStillAlive
            If .npcType = DummyTarget Then
                Call DummyTargetAttacked(TargetNPC)
            End If
        ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.Paralize) Then
            If .flags.Paralizado = 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .pos.x, .pos.y))
                Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
                .flags.Paralizado = 1
                .Contadores.Paralisis = Hechizos(Spell).Duration / 2
            End If
        ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.Immobilize) Then
            If .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .pos.x, .pos.y))
                Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
                .flags.Inmovilizado = 1
                .Contadores.Inmovilizado = Hechizos(Spell).Duration / 2
            End If
        ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveParalysis) Then
            If .flags.Paralizado + .flags.Inmovilizado > 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .pos.x, .pos.y))
                Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
                .flags.Paralizado = 0
                .Contadores.Paralisis = 0
                .flags.Inmovilizado = 0
                .Contadores.Inmovilizado = 0
            End If
        ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.Incinerate) Then
            If .flags.Incinerado = 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .pos.x, .pos.y))
                If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
                    Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageParticleFX(.Char.charindex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))
                End If
                .flags.Incinerado = 1
            End If
        End If
        If IsAlive Then
            Dim Effect As IBaseEffectOverTime
            If Hechizos(Spell).EotId > 0 Then
                Set Effect = FindEffectOnTarget(NpcIndex, NpcList(TargetNPC).EffectOverTime, Hechizos(Spell).EotId)
                If Effect Is Nothing Then
                    Call CreateEffect(NpcIndex, eNpc, TargetNPC, eNpc, Hechizos(Spell).EotId)
                Else
                    Call Effect.Reset(NpcIndex, eNpc, Hechizos(Spell).EotId)
                End If
            End If
        End If
    End With
    With NpcList(NpcIndex)
        If .Char.CastAnimation > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(.Char.charindex, .Char.CastAnimation))
        ElseIf .Char.WeaponAnim > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageArmaMov(.Char.charindex, 0))
        End If
        If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eDisplayCastMessage) Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageChatOverHead("PMAG*" & Spell, .Char.charindex, vbCyan, True, .pos.x, .pos.y, _
                    RequiredSpellDisplayTime, MaxInvisibleSpellDisplayTime))
        End If
    End With
    Exit Sub
NpcLanzaSpellSobreNpc_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.NpcLanzaSpellSobreNpc", Erl)
End Sub

Public Sub NpcLanzaSpellSobreArea(ByVal NpcIndex As Integer, ByVal SpellIndex As Integer)
    On Error GoTo NpcLanzaSpellSobreArea_Err
    Dim afectaUsers    As Boolean
    Dim afectaNPCs     As Boolean
    Dim TargetMap      As t_MapBlock
    Dim PosCasteadaX   As Integer
    Dim PosCasteadaY   As Integer
    Dim x              As Long
    Dim y              As Long
    Dim mitadAreaRadio As Integer
    NpcList(NpcIndex).Contadores.IntervaloLanzarHechizo = GetTickCountRaw()
    With Hechizos(SpellIndex)
        afectaUsers = (.AreaAfecta = 1 Or .AreaAfecta = 3)
        afectaNPCs = (.AreaAfecta = 2 Or .AreaAfecta = 3)
        mitadAreaRadio = CInt(.AreaRadio / 2)
        If IsValidUserRef(NpcList(NpcIndex).TargetUser) Then
            PosCasteadaX = UserList(NpcList(NpcIndex).TargetUser.ArrayIndex).pos.x + RandomNumber(-2, 2)
            PosCasteadaY = UserList(NpcList(NpcIndex).TargetUser.ArrayIndex).pos.y + RandomNumber(-2, 2)
        Else
            PosCasteadaX = NpcList(NpcIndex).pos.x + RandomNumber(-2, 2)
            PosCasteadaY = NpcList(NpcIndex).pos.y + RandomNumber(-1, 2)
        End If
        For x = 1 To .AreaRadio
            For y = 1 To .AreaRadio
                If InMapBounds(NpcList(NpcIndex).pos.Map, x + PosCasteadaX - mitadAreaRadio, PosCasteadaY + y - mitadAreaRadio) Then
                    TargetMap = MapData(NpcList(NpcIndex).pos.Map, x + PosCasteadaX - mitadAreaRadio, PosCasteadaY + y - mitadAreaRadio)
                    If afectaUsers And TargetMap.UserIndex > 0 Then
                        If Not UserList(TargetMap.UserIndex).flags.Muerto And Not EsGM(TargetMap.UserIndex) Then
                            Call NpcLanzaSpellSobreUser(NpcIndex, TargetMap.UserIndex, SpellIndex, True)
                        End If
                    End If
                    If afectaNPCs And TargetMap.NpcIndex > 0 Then
                        If NpcList(TargetMap.NpcIndex).Attackable Then
                            Call NpcLanzaSpellSobreNpc(NpcIndex, TargetMap.NpcIndex, SpellIndex)
                        End If
                    End If
                End If
            Next y
        Next x
        ' El NPC invoca otros npcs independientes
        If .Invoca = 1 Then
            For x = 1 To .cant
                If NpcList(NpcIndex).Contadores.CriaturasInvocadas >= NpcList(NpcIndex).Stats.CantidadInvocaciones Then
                    Exit Sub
                Else
                    Dim npcInvocadoIndex As Integer
                    npcInvocadoIndex = SpawnNpc(.NumNpc, NpcList(NpcIndex).pos, True, False, False)
                    Call SetNpcRef(NpcList(npcInvocadoIndex).flags.Summoner, NpcIndex)
                    NpcList(NpcIndex).Contadores.CriaturasInvocadas = NpcList(NpcIndex).Contadores.CriaturasInvocadas + 1
                    'Si es un NPC que invoca Mas NPCs
                    If NpcList(NpcIndex).Stats.CantidadInvocaciones > 0 Then
                        Dim LoopC As Long
                        'Me fijo cuantos invoca.
                        For LoopC = 1 To NpcList(NpcIndex).Stats.CantidadInvocaciones
                            'Me fijo en que posición tiene en 0 el npcInvocadoIndex
                            If Not IsValidNpcRef(NpcList(NpcIndex).Stats.NpcsInvocados(LoopC)) Then
                                'Y lo agrego
                                Call SetNpcRef(NpcList(NpcIndex).Stats.NpcsInvocados(LoopC), npcInvocadoIndex)
                                Exit For
                            End If
                        Next LoopC
                    End If
                End If
            Next x
        End If
    End With
    With NpcList(NpcIndex)
        If .Char.CastAnimation > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(.Char.charindex, .Char.CastAnimation))
        ElseIf .Char.WeaponAnim > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageArmaMov(.Char.charindex, 0))
        End If
        If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eDisplayCastMessage) Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageChatOverHead("PMAG*" & SpellIndex, .Char.charindex, vbCyan, True, .pos.x, .pos.y, _
                    RequiredSpellDisplayTime, MaxInvisibleSpellDisplayTime))
        End If
    End With
    Exit Sub
NpcLanzaSpellSobreArea_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.NpcLanzaSpellSobreArea", Erl)
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean
    On Error GoTo ErrHandler
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next
    Exit Function
ErrHandler:
End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
    On Error GoTo AgregarHechizo_Err
    Dim hIndex As Integer
    Dim j      As Integer
    hIndex = ObjData(UserList(UserIndex).invent.Object(Slot).ObjIndex).HechizoIndex
    If Not TieneHechizo(hIndex, UserIndex) Then
        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS
            If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
        Next j
        If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
            'Msg777= No tenes espacio para mas hechizos.
            Call WriteLocaleMsg(UserIndex, 777, e_FontTypeNames.FONTTYPE_INFO)
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))
            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
        End If
        UserList(UserIndex).flags.ModificoHechizos = True
    Else
        ' Msg525=Ya tenes ese hechizo.
        Call WriteLocaleMsg(UserIndex, 525, e_FontTypeNames.FONTTYPE_INFO)
    End If
    Exit Sub
AgregarHechizo_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.AgregarHechizo", Erl)
End Sub
            
Sub DecirPalabrasMagicas(ByVal Hechizo As Integer, ByVal UserIndex As Integer)
    On Error GoTo DecirPalabrasMagicas_Err
    UserList(UserIndex).Counters.timeChat = 4
    If Not IsVisible(UserList(UserIndex)) Then
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(UserIndex).Char.charindex, vbCyan, True, UserList( _
                UserIndex).pos.x, UserList(UserIndex).pos.y, RequiredSpellDisplayTime, MaxInvisibleSpellDisplayTime))
    Else
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(UserIndex).Char.charindex, vbCyan, True, UserList( _
                UserIndex).pos.x, UserList(UserIndex).pos.y, 0, 0))
    End If
    Exit Sub
DecirPalabrasMagicas_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.DecirPalabrasMagicas", Erl)
End Sub

Private Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer, Optional ByVal Slot As Integer = 0) As Boolean
    On Error GoTo PuedeLanzar_Err
    PuedeLanzar = False
    If HechizoIndex = 0 Then Exit Function
    With UserList(UserIndex)
        'Si lanza a un npc y este es solo atacable para clanes y el usuario no tiene clan, le avisa y sale de la funcion
        If IsValidNpcRef(.flags.TargetNPC) Then
            If NpcList(.flags.TargetNPC.ArrayIndex).OnlyForGuilds = 1 And .GuildIndex <= 0 Then
                'Msg2001=Debes pertenecer a un clan para atacar a este NPC
                Call WriteLocaleMsg(UserIndex, 2001, e_FontTypeNames.FONTTYPE_WARNING)
                Exit Function
            End If
            If NpcList(.flags.TargetNPC.ArrayIndex).flags.ImmuneToSpells <> 0 Then
                Call WriteLocaleMsg(UserIndex, MSG_NPC_INMUNE_TO_SPELLS, e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        End If
        If .flags.EnConsulta Then
            'Msg778= No puedes lanzar hechizos si estas en consulta.
            Call WriteLocaleMsg(UserIndex, 778, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If Hechizos(HechizoIndex).AutoLanzar And .flags.TargetUser.ArrayIndex <> UserIndex Then
            Exit Function
        End If
        If IsSet(.flags.StatusMask, eCastOnlyOnSelf) And .flags.TargetUser.ArrayIndex <> UserIndex Then
            Call WriteLocaleMsg(UserIndex, MsgCastOnlyOnSelf, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If IsSet(Hechizos(HechizoIndex).Effects, e_SpellEffects.CancelActiveEffect) And Hechizos(HechizoIndex).EotId > 0 And IsValidUserRef(.flags.TargetUser) Then
            Dim Effect As IBaseEffectOverTime
            Set Effect = FindEffectOnTarget(UserIndex, UserList(.flags.TargetUser.ArrayIndex).EffectOverTime, Hechizos(HechizoIndex).EotId)
            If Not Effect Is Nothing Then
                If Effect.EotId = Hechizos(HechizoIndex).EotId Then
                    Effect.RemoveMe = True
                    Exit Function
                End If
            End If
        End If
        If Hechizos(HechizoIndex).RequireTransform > 0 Then
            If .flags.ActiveTransform <> Hechizos(HechizoIndex).RequireTransform Then
                Call WriteLocaleMsg(UserIndex, MsgSpellRequiresTransform, e_FontTypeNames.FONTTYPE_INFO, GetNpcName(Hechizos(HechizoIndex).RequireTransform))
                Exit Function
            End If
        End If
        If IsFeatureEnabled("healers_and_tanks") And .flags.DivineBlood > 0 And IsSet(Hechizos(HechizoIndex).Effects, e_SpellEffects.eDoDamage) Then
            Call WriteLocaleMsg(UserIndex, 2095, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If .flags.Privilegios And e_PlayerType.Consejero Then
            Exit Function
        End If
        If MapInfo(.pos.Map).SinMagia And Not IsSet(Hechizos(HechizoIndex).SpellRequirementMask, eIsSkill) Then
            'Msg779= Una fuerza mística te impide lanzar hechizos en esta zona.
            Call WriteLocaleMsg(UserIndex, 779, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
        If .flags.Montado = 1 Then
            'Msg780= No puedes lanzar hechizos si estas montado.
            Call WriteLocaleMsg(UserIndex, 780, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If Hechizos(HechizoIndex).NecesitaObj > 0 Then
            If Not IsObjecIndextInInventory(UserIndex, Hechizos(HechizoIndex).NecesitaObj) And Not IsObjecIndextInInventory(UserIndex, Hechizos(HechizoIndex).NecesitaObj2) Then
                If Not IsObjecIndextInInventory(UserIndex, Hechizos(HechizoIndex).NecesitaObj) And Not IsObjecIndextInInventory(UserIndex, Hechizos(HechizoIndex).NecesitaObj2) Then
                    Call WriteLocaleMsg(UserIndex, 1634, e_FontTypeNames.FONTTYPE_INFO, ObjData(Hechizos(HechizoIndex).NecesitaObj).name) 'Msg1634=Necesitas un ¬1 para lanzar el hechizo.
                    Exit Function
                End If
            End If
        End If
        If IsValidUserRef(.flags.TargetUser) Then
            If Hechizos(HechizoIndex).TargetEffectType = e_TargetEffectType.ePositive Then
                Dim UserInteractionResult As e_InteractionResult
                UserInteractionResult = UserMod.CanHelpUser(UserIndex, .flags.TargetUser.ArrayIndex)
                If UserInteractionResult <> e_InteractionResult.eInteractionOk Then
                    Call SendHelpInteractionMessage(UserIndex, UserInteractionResult)
                    Exit Function
                End If
            End If
            If Hechizos(HechizoIndex).TargetEffectType = e_TargetEffectType.eNegative Then
                Dim UserAttackInteractionResultUser As e_AttackInteractionResult
                UserAttackInteractionResultUser = UserMod.CanAttackUser(UserIndex, .VersionId, .flags.TargetUser.ArrayIndex, .flags.TargetUser.VersionId)
                If UserAttackInteractionResultUser <> e_AttackInteractionResult.eCanAttack Then
                    Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResultUser)
                    Exit Function
                End If
            End If
        ElseIf IsValidNpcRef(.flags.TargetNPC) Then
            If Hechizos(HechizoIndex).TargetEffectType = e_TargetEffectType.eNegative Then
                Dim UserAttackInteractionResult As t_AttackInteractionResult
                UserAttackInteractionResult = UserCanAttackNpc(UserIndex, .flags.TargetNPC.ArrayIndex)
                If UserAttackInteractionResult.Result = e_AttackInteractionResult.eAttackCitizenNpc Or UserAttackInteractionResult.Result = _
                        e_AttackInteractionResult.eRemoveSafeCitizenNpc Or UserAttackInteractionResult.Result = e_AttackInteractionResult.eSameFaction Or _
                        UserAttackInteractionResult.Result = e_AttackInteractionResult.eRemoveSafe Then
                    Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
                    If UserAttackInteractionResult.CanAttack Then
                        If UserAttackInteractionResult.TurnPK Then VolverCriminal (UserIndex)
                    Else
                        Exit Function
                    End If
                End If
            End If
        End If
        If Hechizos(HechizoIndex).Cooldown > 0 And .Counters.UserHechizosInterval(Slot) > 0 Then
            Dim nowRaw            As Long
            Dim SegundosFaltantes As Long
            nowRaw = GetTickCountRaw()
            Dim Cooldown As Long
            Cooldown = Hechizos(HechizoIndex).Cooldown
            'cooldown reduction for Elven Wood items
            If .invent.EquippedWeaponObjIndex > 0 Then
                If ObjData(.invent.EquippedWeaponObjIndex).MaderaElfica > 0 Then
                    Cooldown = Cooldown / 2
                End If
            End If
            If .invent.EquippedRingAccesoryObjIndex > 0 Then
                If ObjData(.invent.EquippedRingAccesoryObjIndex).MaderaElfica > 0 Then
                    Cooldown = Cooldown / 2
                End If
            End If
            Cooldown = Cooldown * 1000
            Dim elapsedMs As Double
            elapsedMs = TicksElapsed(.Counters.UserHechizosInterval(Slot), nowRaw)
            If elapsedMs < Cooldown Then
                SegundosFaltantes = Int((Cooldown - elapsedMs) / 1000)
                Call WriteLocaleMsg(UserIndex, 1635, e_FontTypeNames.FONTTYPE_WARNING, SegundosFaltantes) 'Msg1635=Debes esperar ¬1 segundos para volver a tirar este hechizo.
                Exit Function
            End If
        End If
        If .Stats.UserSkills(e_Skill.Magia) < Hechizos(HechizoIndex).MinSkill Then
            Call WriteLocaleMsg(UserIndex, 1636, e_FontTypeNames.FONTTYPE_INFO, Hechizos(HechizoIndex).MinSkill) 'Msg1636=No tienes suficientes puntos de magia para lanzar este hechizo, necesitas ¬1 puntos.
            Exit Function
        End If
        If Hechizos(HechizoIndex).MaxLevelCasteable > 0 And .Stats.ELV > Hechizos(HechizoIndex).MaxLevelCasteable Then
            Call WriteLocaleMsg(UserIndex, 2116, e_FontTypeNames.FONTTYPE_INFO, Hechizos(HechizoIndex).MaxLevelCasteable) 'Msg2116=Para lanzar este hechizo debes ser nivel ¬1 o inferior.
            Exit Function
        End If
        If .Stats.MinHp < Hechizos(HechizoIndex).RequiredHP Then
            Call WriteLocaleMsg(UserIndex, 1637, e_FontTypeNames.FONTTYPE_INFO, Hechizos(HechizoIndex).RequiredHP) 'Msg1637=No tienes suficiente vida. Necesitas ¬1 puntos de vida.
            Exit Function
        End If
        If .Stats.MinMAN < GetSpellManaCostModifierByClass(UserIndex, Hechizos(HechizoIndex), HechizoIndex) Then
            Call WriteLocaleMsg(UserIndex, 222, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
            'Msg93=Estás muy cansado
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            'Msg2129=¡No tengo energía!
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
            Exit Function
        End If
        If .clase = e_Class.Mage And Not IsFeatureEnabled("remove-staff-requirements") Then
            If Hechizos(HechizoIndex).NeedStaff > 0 Then
                If .invent.EquippedWeaponObjIndex = 0 Then
                    'Msg781= Necesitás un báculo para lanzar este hechizo.
                    Call WriteLocaleMsg(UserIndex, 781, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
                If ObjData(.invent.EquippedWeaponObjIndex).Power < Hechizos(HechizoIndex).NeedStaff Then
                    'Msg782= Necesitás un báculo más poderoso para lanzar este hechizo.
                    Call WriteLocaleMsg(UserIndex, 782, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        End If
        If .clase = e_Class.Druid Then
            If Hechizos(HechizoIndex).RequiereInstrumento > 0 Then
                If .invent.EquippedRingAccesoryObjIndex = 0 Then
                    'Msg783= Necesitás una flauta para invocar o desinvocar a tus mascotas.
                    Call WriteLocaleMsg(UserIndex, 783, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                Else
                    If ObjData(.invent.EquippedRingAccesoryObjIndex).InstrumentoRequerido <> 1 Then
                    'Msg783= Necesitás una flauta para invocar o desinvocar a tus mascotas.
                        Call WriteLocaleMsg(UserIndex, 783, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                    End If
                End If
            End If
        End If
        If Hechizos(HechizoIndex).RequireWeaponType > 0 Then
            If .invent.EquippedWeaponObjIndex = 0 Then
                Call WriteLocaleMsg(UserIndex, GetRequiredWeaponLocaleId(Hechizos(HechizoIndex).RequireWeaponType), e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
            If ObjData(.invent.EquippedWeaponObjIndex).WeaponType <> Hechizos(HechizoIndex).RequireWeaponType Then
                Call WriteLocaleMsg(UserIndex, GetRequiredWeaponLocaleId(Hechizos(HechizoIndex).RequireWeaponType), e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        End If
        Dim RequiredItemResult As e_SpellRequirementMask
        RequiredItemResult = TestRequiredEquipedItem(.invent, Hechizos(HechizoIndex).SpellRequirementMask, 0)
        If RequiredItemResult > 0 Then
            Call SendrequiredItemMessage(UserIndex, RequiredItemResult, "para usar este hechizo.")
            Exit Function
        End If
        Dim TargetRef As t_AnyReference
        If IsValidUserRef(.flags.TargetUser) Then
            Call CastUserToAnyRef(.flags.TargetUser, TargetRef)
        ElseIf IsValidNpcRef(.flags.TargetNPC) Then
            Call CastNpcToAnyRef(.flags.TargetNPC, TargetRef)
        End If
        If IsValidRef(TargetRef) Then
            If IsDead(TargetRef) And Not IsSet(Hechizos(HechizoIndex).SpellRequirementMask, e_SpellRequirementMask.eWorkOnDead) Then
                Call WriteLocaleMsg(UserIndex, 7, e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        End If
        If IsSet(Hechizos(HechizoIndex).SpellRequirementMask, e_SpellRequirementMask.eRequireTargetOnLand) And IsValidRef(TargetRef) Then
            If TargetRef.RefType = eUser Then
                If UserList(TargetRef.ArrayIndex).flags.Nadando > 0 Or .flags.Navegando > 0 Or .flags.Montado > 0 Then
                    Call WriteLocaleMsg(UserIndex, MsgLandRequiredToUseSpell, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        End If
        If IsSet(Hechizos(HechizoIndex).SpellRequirementMask, e_SpellRequirementMask.eRequireTargetOnWater) And IsValidRef(TargetRef) Then
            If TargetRef.RefType = eUser Then
                If UserList(TargetRef.ArrayIndex).flags.Nadando = 0 And .flags.Navegando = 0 Then
                    Call WriteLocaleMsg(UserIndex, MsgWaterRequiredToUseSpell, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        End If
        PuedeLanzar = True
    End With
    Exit Function
PuedeLanzar_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.PuedeLanzar", Erl)
End Function

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.
Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)
    '***************************************************
    'Author: Uknown
    'Modification: 06/15/2008 (NicoNZ)
    'Last modification: 01/12/2020 (WyroX)
    'Sale del sub si no hay una posición valida.
    '***************************************************
    On Error GoTo HechizoInvocacion_Err
    With UserList(UserIndex)
        If .flags.EnReto Then
            'Msg784= No podés invocar criaturas durante un reto.
            Call WriteLocaleMsg(UserIndex, 784, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim h         As Integer, j As Integer, ind As Integer, Index As Integer
        Dim TargetPos As t_WorldPos
        TargetPos.Map = .flags.TargetMap
        TargetPos.x = .flags.TargetX
        TargetPos.y = .flags.TargetY
        h = .Stats.UserHechizos(.flags.Hechizo)
        If Hechizos(h).Invoca = 1 Then
            ' No puede invocar en este mapa
            If MapInfo(.pos.Map).NoMascotas Then
                'Msg785= Un gran poder te impide invocar criaturas en este mapa.
                Call WriteLocaleMsg(UserIndex, 785, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Dim MinTiempo As Integer
            Dim i         As Integer
            For i = 1 To Hechizos(h).cant
                Index = -1
                MinTiempo = IntervaloInvocacion
                For j = 1 To MAXMASCOTAS
                    If .MascotasIndex(j).ArrayIndex > 0 Then
                        If IsValidNpcRef(.MascotasIndex(j)) Then
                            If NpcList(.MascotasIndex(j).ArrayIndex).flags.NPCActive Then
                                If NpcList(.MascotasIndex(j).ArrayIndex).Contadores.TiempoExistencia > 0 And NpcList(.MascotasIndex(j).ArrayIndex).Contadores.TiempoExistencia < _
                                        MinTiempo Then
                                    Index = j
                                    MinTiempo = NpcList(.MascotasIndex(j).ArrayIndex).Contadores.TiempoExistencia
                                End If
                            Else
                                Call ClearNpcRef(.MascotasIndex(j))
                                Index = -1
                                Exit For
                            End If
                        Else
                            Index = -1
                            MinTiempo = 0
                        End If
                    ElseIf .MascotasType(j) = 0 Then
                        Index = -1
                        MinTiempo = 0
                    End If
                Next j
                If Index > -1 Then
                    If IsValidNpcRef(.MascotasIndex(Index)) Then
                        Call QuitarNPC(.MascotasIndex(Index).ArrayIndex, eSummonNew)
                    End If
                End If
                If .NroMascotas < MAXMASCOTAS Then
                    ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, False, False, False, UserIndex, Hechizos(h).wav)
                    If ind > 0 Then
                        .NroMascotas = .NroMascotas + 1
                        Index = FreeMascotaIndex(UserIndex)
                        Call SetNpcRef(.MascotasIndex(Index), ind)
                        .MascotasType(Index) = NpcList(ind).Numero
                        Call SetUserRef(NpcList(ind).MaestroUser, UserIndex)
                        NpcList(ind).Contadores.TiempoExistencia = IntervaloInvocacion
                        NpcList(ind).GiveGLD = 0
                        If IsFeatureEnabled("addjust-npc-with-caster") And IsSet(Hechizos(h).Effects, AdjustStatsWithCaster) Then
                            Call AdjustNpcStatWithCasterLevel(UserIndex, ind)
                        End If
                        Call FollowAmo(ind)
                    Else
                        Exit Sub
                    End If
                Else
                    Exit For
                End If
            Next i
            Call InfoHechizo(UserIndex)
            b = True
        ElseIf Hechizos(h).Invoca = 2 Then
            ' Si tiene mascotas
            If .NroMascotas > 0 Then
                ' Tiene que estar en zona insegura
                ' No puede invocar en este mapa
                If MapInfo(.pos.Map).NoMascotas Then
                    Call WriteLocaleMsg(UserIndex, 786, e_FontTypeNames.FONTTYPE_INFO) 'Msg786= Un gran poder te impide invocar criaturas en este mapa.
                    Exit Sub
                End If
                ' Si no están guardadas las mascotas
                If .flags.MascotasGuardadas = 0 Then
                    For i = 1 To MAXMASCOTAS
                        If IsValidNpcRef(.MascotasIndex(i)) Then
                            ' Si no es un elemental, lo "guardamos"... lo matamos
                            If NpcList(.MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia = 0 Then
                                ' Le saco el maestro, para que no me lo quite de mis mascotas
                                Call SetUserRef(NpcList(.MascotasIndex(i).ArrayIndex).MaestroUser, 0)
                                ' Lo borro
                                Call QuitarNPC(.MascotasIndex(i).ArrayIndex, eStorePets)
                                ' Saco el índice
                                Call ClearNpcRef(.MascotasIndex(i))
                                b = True
                            End If
                        Else
                            Call ClearNpcRef(.MascotasIndex(i))
                        End If
                    Next
                    .flags.MascotasGuardadas = 1
                    ' Ya están guardadas, así que las invocamos
                Else
                    For i = 1 To MAXMASCOTAS
                        ' Si está guardada y no está ya en el mapa
                        If .MascotasType(i) > 0 And .MascotasIndex(i).ArrayIndex = 0 Then
                            Call SetNpcRef(.MascotasIndex(i), SpawnNpc(.MascotasType(i), TargetPos, True, True, False, UserIndex))
                            Call SetUserRef(NpcList(.MascotasIndex(i).ArrayIndex).MaestroUser, UserIndex)
                            Call FollowAmo(.MascotasIndex(i).ArrayIndex)
                            If IsFeatureEnabled("addjust-npc-with-caster") And IsSet(Hechizos(h).Effects, AdjustStatsWithCaster) Then
                                Call AdjustNpcStatWithCasterLevel(UserIndex, .MascotasIndex(i).ArrayIndex)
                            End If
                            b = True
                        End If
                    Next
                    .flags.MascotasGuardadas = 0
                End If
            Else
                'Msg787= No tienes mascotas.
                Call WriteLocaleMsg(UserIndex, 787, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If b Then Call InfoHechizo(UserIndex)
        End If
    End With
    Exit Sub
HechizoInvocacion_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoInvocacion")
End Sub

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
    On Error GoTo HechizoTerrenoEstado_Err
    Dim PosCasteadaX As Integer
    Dim PosCasteadaY As Integer
    Dim PosCasteadaM As Integer
    Dim h            As Integer
    Dim TempX        As Integer
    Dim TempY        As Integer
    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveInvisibility) Then
        b = True
        For TempX = PosCasteadaX - 11 To PosCasteadaX + 11
            For TempY = PosCasteadaY - 11 To PosCasteadaY + 11
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, _
                                TempY).UserIndex).flags.NoDetectable = 0 Then
                            UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 0
                            Call WriteConsoleMsg(MapData(PosCasteadaM, TempX, TempY).UserIndex, PrepareMessageLocaleMsg(1869, vbNullString, e_FontTypeNames.FONTTYPE_INFOIAO)) ' Msg1869=Tu invisibilidad ya no tiene efecto.
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.charindex, _
                                    False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                        End If
                    End If
                End If
            Next TempY
        Next TempX
        Call InfoHechizo(UserIndex)
    End If
    Exit Sub
HechizoTerrenoEstado_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoTerrenoEstado", Erl)
End Sub

Private Sub HechizoSobreArea(ByVal UserIndex As Integer, ByRef b As Boolean)
    On Error GoTo HechizoSobreArea_Err
    Dim afectaUsers  As Boolean
    Dim afectaNPCs   As Boolean
    Dim TargetMap    As t_MapBlock
    Dim PosCasteadaX As Byte
    Dim PosCasteadaY As Byte
    Dim h            As Integer
    Dim x            As Long
    Dim y            As Long
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    'Envio Palabras magicas, wavs y fxs.
    If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
        Call DecirPalabrasMagicas(h, UserIndex)
    End If
    If Hechizos(h).Particle > 0 Then 'Envio Particula?
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(PosCasteadaX, PosCasteadaY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
    End If
    If Hechizos(h).ParticleViaje = 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, PosCasteadaX, PosCasteadaY))
    End If
    afectaUsers = (Hechizos(h).AreaAfecta = 1 Or Hechizos(h).AreaAfecta = 3)
    afectaNPCs = (Hechizos(h).AreaAfecta = 2 Or Hechizos(h).AreaAfecta = 3)
    For x = 1 To Hechizos(h).AreaRadio
        For y = 1 To Hechizos(h).AreaRadio
            TargetMap = MapData(UserList(UserIndex).pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + y - CInt(Hechizos(h).AreaRadio / 2))
            If afectaUsers And TargetMap.UserIndex > 0 Then
                If UserList(TargetMap.UserIndex).flags.Muerto = 0 Then
                    Call AreaHechizo(UserIndex, TargetMap.UserIndex, PosCasteadaX, PosCasteadaY, False)
                    If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
                        If Hechizos(h).ParticleViaje > 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.charindex, Hechizos(h).ParticleViaje, _
                                    Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, x, y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(TargetMap.UserIndex).pos.x, UserList( _
                                    TargetMap.UserIndex).pos.y))
                        End If
                    End If
                End If
            End If
            If afectaNPCs And TargetMap.NpcIndex > 0 Then
                If NpcList(TargetMap.NpcIndex).Attackable Then
                    Call AreaHechizo(UserIndex, TargetMap.NpcIndex, PosCasteadaX, PosCasteadaY, True)
                    If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
                        If Hechizos(h).ParticleViaje > 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.charindex, Hechizos(h).ParticleViaje, _
                                    Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, x, y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(Hechizos(h).FXgrh, NpcList(TargetMap.NpcIndex).pos.x, NpcList( _
                                    TargetMap.NpcIndex).pos.y))
                        End If
                    End If
                End If
            End If
        Next y
    Next x
    b = True
    Exit Sub
HechizoSobreArea_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoSobreArea", Erl)
End Sub

Sub HechizoPortal(ByVal UserIndex As Integer, ByRef b As Boolean)
    On Error GoTo HechizoPortal_Err
    Dim PosCasteadaX As Byte
    Dim PosCasteadaY As Byte
    Dim PosCasteadaM As Integer
    Dim uh           As Integer
    Dim TempX        As Integer
    Dim TempY        As Integer
    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    uh = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    'Envio Palabras magicas, wavs y fxs.
    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.amount > 0 Or (MapData(UserList(UserIndex).pos.Map, _
            UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES Or MapData(UserList(UserIndex).pos.Map, _
            UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).TileExit.Map > 0 Or UserList(UserIndex).flags.TargetUser.ArrayIndex <> 0 Then
        b = False
        'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", e_FontTypeNames.FONTTYPE_INFO)
        Call WriteLocaleMsg(UserIndex, 262, e_FontTypeNames.FONTTYPE_INFO)
    Else
        If Hechizos(uh).TeleportX = 1 Then
            If UserList(UserIndex).flags.Portal = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_GraphicEffects.Runa, -1, False))
                UserList(UserIndex).flags.PortalM = UserList(UserIndex).pos.Map
                UserList(UserIndex).flags.PortalX = UserList(UserIndex).flags.TargetX
                UserList(UserIndex).flags.PortalY = UserList(UserIndex).flags.TargetY
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.charindex, 600, e_AccionBarra.Intermundia))
                UserList(UserIndex).Accion.AccionPendiente = True
                UserList(UserIndex).Accion.Particula = e_GraphicEffects.Runa
                UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.Intermundia
                UserList(UserIndex).Accion.HechizoPendiente = uh
                If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
                    Call DecirPalabrasMagicas(uh, UserIndex)
                End If
                b = True
            Else
                'Msg788= No podés lanzar mas de un portal a la vez.
                Call WriteLocaleMsg(UserIndex, 788, e_FontTypeNames.FONTTYPE_INFO)
                b = False
            End If
        End If
    End If
    Exit Sub
HechizoPortal_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoPortal", Erl)
End Sub

Sub HechizoMaterializacion(ByVal UserIndex As Integer, ByRef b As Boolean)
    On Error GoTo HechizoMaterializacion_Err
    Dim h   As Integer
    Dim MAT As t_Obj
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.amount > 0 Or MapData(UserList(UserIndex).pos.Map, _
            UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Then
        b = False
        Call WriteLocaleMsg(UserIndex, 262, e_FontTypeNames.FONTTYPE_INFO)
        ' Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", e_FontTypeNames.FONTTYPE_INFO)
    Else
        MAT.amount = Hechizos(h).MaterializaCant
        MAT.ObjIndex = Hechizos(h).MaterializaObj
        Call MakeObj(MAT, UserList(UserIndex).pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, _
                Hechizos(h).TimeParticula))
        b = True
    End If
    Exit Sub
HechizoMaterializacion_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoMaterializacion", Erl)
End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)
    On Error GoTo HandleHechizoTerreno_Err
    Dim b As Boolean
    With UserList(UserIndex)
        Select Case Hechizos(uh).Tipo
            Case e_TipoHechizo.uInvocacion 'Tipo 1
                Call HechizoInvocacion(UserIndex, b)
            Case e_TipoHechizo.uEstado 'Tipo 2
                Call HechizoTerrenoEstado(UserIndex, b)
            Case e_TipoHechizo.uMaterializa 'Tipo 3
                Call HechizoMaterializacion(UserIndex, b)
            Case e_TipoHechizo.uArea 'Tipo 5
                Call HechizoSobreArea(UserIndex, b)
            Case e_TipoHechizo.uPortal 'Tipo 6
                Call HechizoPortal(UserIndex, b)
            Case e_TipoHechizo.uMultiShoot
                Dim TargetPos As t_WorldPos
                TargetPos.Map = .pos.Map
                TargetPos.x = .flags.TargetX
                TargetPos.y = .flags.TargetY
                b = MultiShot(UserIndex, TargetPos)
        End Select
        If b Then
            If Not IsSet(Hechizos(uh).SpellRequirementMask, eIsSkill) Then
                Call SubirSkill(UserIndex, Magia)
            End If
            .Stats.MinMAN = .Stats.MinMAN - GetSpellManaCostModifierByClass(UserIndex, Hechizos(uh), uh)
            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            .Stats.MinSta = .Stats.MinSta - Hechizos(uh).StaRequerido
            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            Call WriteUpdateMana(UserIndex)
            Call WriteUpdateSta(UserIndex)
        End If
    End With
    Exit Sub
HandleHechizoTerreno_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HandleHechizoTerreno", Erl)
End Sub

Function HandlePetSpell(ByVal UserIndex As Integer, ByVal uh As Integer) As Boolean
    With UserList(UserIndex)
        If .NroMascotas = 0 Then
            Exit Function
        End If
        If Hechizos(uh).EotId = 0 Then
            Exit Function
        End If
        Dim j As Integer
        For j = 1 To MAXMASCOTAS
            If IsValidNpcRef(.MascotasIndex(j)) Then
                Dim Effect As IBaseEffectOverTime
                Set Effect = FindEffectOnTarget(UserIndex, NpcList(.MascotasIndex(j).ArrayIndex).EffectOverTime, Hechizos(uh).EotId)
                If Not Effect Is Nothing Then
                    If Not EffectOverTime(Hechizos(uh).EotId).Override Then
                        Exit For
                    End If
                End If
                If Effect Is Nothing Then
                    Call CreateEffect(UserIndex, eUser, .MascotasIndex(j).ArrayIndex, eNpc, Hechizos(uh).EotId)
                Else
                    If Not Effect.Reset(UserIndex, eUser, Hechizos(uh).EotId) Then
                        Exit For
                    End If
                End If
            End If
        Next j
        If Not IsSet(Hechizos(uh).SpellRequirementMask, eIsSkill) Then
            Call SubirSkill(UserIndex, Magia)
        End If
        .Stats.MinMAN = .Stats.MinMAN - GetSpellManaCostModifierByClass(UserIndex, Hechizos(uh), uh)
        If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
        .Stats.MinSta = .Stats.MinSta - Hechizos(uh).StaRequerido
        If .Stats.MinSta < 0 Then .Stats.MinSta = 0
        Call WriteUpdateMana(UserIndex)
        Call WriteUpdateSta(UserIndex)
        HandlePetSpell = True
    End With
End Function

Function HandlePhysicalSkill(ByVal SourceIndex As Integer, _
                             ByVal SourceType As e_ReferenceType, _
                             ByVal TargetIndex As Integer, _
                             ByVal TargetType As e_ReferenceType, _
                             ByVal SpellIndex As Integer, _
                             IsAlive As Boolean) As Boolean
    Dim TargetRef As t_AnyReference
    Dim SourceRef As t_AnyReference
    Call SetRef(SourceRef, SourceIndex, SourceType)
    Call SetRef(TargetRef, TargetIndex, TargetType)
    Dim TargetPos As t_WorldPos
    Dim SourcePos As t_WorldPos
    SourcePos = GetPosition(SourceRef)
    TargetPos = GetPosition(TargetRef)
    Select Case Hechizos(SpellIndex).SkillType
        Case e_SkillType.ePushingArrow
            If Not IntervaloPermiteUsarArcos(SourceIndex, False) Then Exit Function
            Dim Damage      As Integer
            Dim objectIndex As Integer
            Dim Proyectile  As Integer
            If SourceType = eUser Then
                With UserList(SourceIndex)
                    If .invent.EquippedMunitionObjIndex = 0 Then
                        Exit Function
                    End If
                    Damage = GetUserDamageWithItem(SourceIndex, .invent.EquippedWeaponObjIndex, .invent.EquippedMunitionObjIndex) / 2
                    objectIndex = .invent.EquippedWeaponObjIndex
                    Proyectile = ObjData(.invent.EquippedMunitionObjIndex).ProjectileType
                End With
            Else
                Damage = RandomNumber(NpcList(SourceIndex).Stats.MinHIT, NpcList(SourceIndex).Stats.MaxHit)
                objectIndex = -1
                Proyectile = 1
            End If
            If RefDoDamageToTarget(SourceRef, TargetRef, Damage, e_phisical, objectIndex) = eStillAlive Then
                IsAlive = True
                If TargetRef.RefType = eUser Then
                    UserList(TargetRef.ArrayIndex).Counters.timeFx = 3
                    Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(TargetRef.ArrayIndex).Char.charindex, FXSANGRE, 0, UserList( _
                            TargetRef.ArrayIndex).pos.x, UserList(TargetRef.ArrayIndex).pos.y))
                    Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(TargetRef.ArrayIndex).pos.x, UserList( _
                            TargetRef.ArrayIndex).pos.y))
                Else
                    If NpcList(TargetRef.ArrayIndex).flags.Snd2 > 0 Then
                        Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(NpcList(TargetRef.ArrayIndex).flags.Snd2, NpcList( _
                                TargetRef.ArrayIndex).pos.x, NpcList(TargetRef.ArrayIndex).pos.y))
                    Else
                        Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(TargetRef.ArrayIndex).pos.x, NpcList( _
                                TargetRef.ArrayIndex).pos.y))
                    End If
                End If
            Else
                IsAlive = False
            End If
            If SourceRef.RefType = eUser Then
                Call SendData(SendTarget.ToPCAliveArea, SourceIndex, PrepareCreateProjectile(SourcePos.x, SourcePos.y, TargetPos.x, TargetPos.y, Proyectile))
            Else
                Call SendData(SendTarget.ToNPCAliveArea, SourceIndex, PrepareCreateProjectile(SourcePos.x, SourcePos.y, TargetPos.x, TargetPos.y, Proyectile))
            End If
            HandlePhysicalSkill = True
            Exit Function
        Case e_SkillType.eCannon
            If SourceType = eUser Then
                Debug.Assert "User cannot use this spell"
            End If
            Dim Particula        As Integer
            Dim Tiempo           As Long
            Dim CannonProyectile As Integer
            CannonProyectile = 4
            Particula = Hechizos(SpellIndex).Particle
            Tiempo = Hechizos(SpellIndex).TimeParticula
            Call SendData(SendTarget.ToNPCAliveArea, SourceIndex, PrepareMessageParticleFX(NpcList(SourceIndex).Char.charindex, Particula, Tiempo, False, , SourcePos.x, _
                    SourcePos.y))
            Call SendData(SendTarget.ToNPCAliveArea, SourceIndex, PrepareCreateProjectile(SourcePos.x, SourcePos.y, TargetPos.x, TargetPos.y, CannonProyectile))
            If Hechizos(SpellIndex).wav <> 0 Then Call SendData(SendTarget.ToNPCAliveArea, SourceIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).wav, SourcePos.x, SourcePos.y))
            Call CreateDelayedBlast(SourceIndex, SourceType, TargetPos.Map, TargetPos.x, TargetPos.y, Hechizos(SpellIndex).EotId, -1)
            HandlePhysicalSkill = False
            Exit Function
    End Select
End Function

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
    On Error GoTo HandleHechizoUsuario_Err
    Dim IsAlive As Boolean
    IsAlive = True
    Dim b      As Boolean
    Dim Effect As IBaseEffectOverTime
    With UserList(UserIndex)
        If Hechizos(uh).EotId > 0 And IsValidUserRef(.flags.TargetUser) Then
            Set Effect = FindEffectOnTarget(UserIndex, UserList(.flags.TargetUser.ArrayIndex).EffectOverTime, Hechizos(uh).EotId)
            If Not Effect Is Nothing Then
                If Not EffectOverTime(Hechizos(uh).EotId).Override Then
                    Call WriteLocaleMsg(UserIndex, MsgTargetAlreadyAffected, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
        Select Case Hechizos(uh).Tipo
            Case e_TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
                Call HechizoEstadoUsuario(UserIndex, b)
            Case e_TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
                Call HechizoPropUsuario(UserIndex, b, IsAlive)
            Case e_TipoHechizo.uCombinados
                Call HechizoCombinados(UserIndex, b, IsAlive)
            Case e_TipoHechizo.uPhysicalSkill
                b = HandlePhysicalSkill(UserIndex, eUser, .flags.TargetUser.ArrayIndex, eUser, .Stats.UserHechizos(.flags.Hechizo), IsAlive)
        End Select
        If b Then
            If Hechizos(uh).EotId > 0 And IsAlive Then
                If Effect Is Nothing Then
                    Call CreateEffect(UserIndex, eUser, .flags.TargetUser.ArrayIndex, eUser, Hechizos(uh).EotId)
                Else
                    If Not Effect.Reset(UserIndex, eUser, Hechizos(uh).EotId) Then
                        Exit Sub
                    End If
                End If
            End If
            If Not IsSet(Hechizos(uh).SpellRequirementMask, eIsSkill) Then
                Call SubirSkill(UserIndex, Magia)
            End If
            .Stats.MinMAN = .Stats.MinMAN - GetSpellManaCostModifierByClass(UserIndex, Hechizos(uh), uh)
            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            If Hechizos(uh).RequiredHP > 0 Then
                Call UserMod.ModifyHealth(UserIndex, -Hechizos(uh).RequiredHP, 1)
            End If
            .Stats.MinSta = .Stats.MinSta - Hechizos(uh).StaRequerido
            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            If IsSet(Hechizos(uh).Effects, e_SpellEffects.Resurrect) Then
                If Not PeleaSegura(UserIndex, .flags.TargetUser.ArrayIndex) Then
                    If MapInfo(.pos.Map).Seguro = 0 Then
                        Dim costoVidaResu As Long
                        costoVidaResu = UserList(.flags.TargetUser.ArrayIndex).Stats.ELV * 1.5 + .Stats.MinHp * 0.45
                        Call UserMod.ModifyHealth(UserIndex, -costoVidaResu, 1)
                    End If
                End If
            End If
            Call WriteUpdateMana(UserIndex)
            Call WriteUpdateHP(UserIndex)
            Call WriteUpdateSta(UserIndex)
        End If
    End With
    Exit Sub
HandleHechizoUsuario_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HandleHechizoUsuario", Erl)
End Sub

Public Function GetSpellManaCostModifierByClass(ByVal UserIndex As Integer, Hechizo As t_Hechizo, Optional ByVal HechizoIndex As Long) As Integer
    GetSpellManaCostModifierByClass = Hechizo.ManaRequerido
    With UserList(UserIndex)
        Select Case .clase
            Case e_Class.Bard
                If HechizoIndex = MauveFlashIndex And .invent.EquippedRingAccesoryObjIndex = MagicLuteIndex Then
                    GetSpellManaCostModifierByClass = 80
                    Exit Function
                ElseIf HechizoIndex = FireEcoIndex And .invent.EquippedRingAccesoryObjIndex = MagicLuteIndex Then
                    GetSpellManaCostModifierByClass = 70
                    Exit Function
                End If
            Case e_Class.Cleric
                If IsFeatureEnabled("healers_and_tanks") And .flags.DivineBlood > 0 Then
                    If IsSet(Hechizo.Effects, e_SpellEffects.eDoHeal) Then
                        GetSpellManaCostModifierByClass = GetSpellManaCostModifierByClass * DivineBloodManaCostMultiplier
                    End If
                End If
        End Select
    End With
End Function

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
    On Error GoTo HandleHechizoNPC_Err
    Dim b       As Boolean
    Dim Effect  As IBaseEffectOverTime
    Dim IsAlive As Boolean
    With UserList(UserIndex)
        IsAlive = True
        If Hechizos(uh).EotId > 0 And IsValidNpcRef(.flags.TargetNPC) Then
            Set Effect = FindEffectOnTarget(UserIndex, NpcList(.flags.TargetNPC.ArrayIndex).EffectOverTime, Hechizos(uh).EotId)
            If Not Effect Is Nothing Then
                If Not EffectOverTime(Hechizos(uh).EotId).Override Then
                    Call WriteLocaleMsg(UserIndex, MsgTargetAlreadyAffected, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
        Call AllMascotasAtacanNPC(.flags.TargetNPC.ArrayIndex, UserIndex)
        Select Case Hechizos(uh).Tipo
            Case e_TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
                Call HechizoEstadoNPC(.flags.TargetNPC.ArrayIndex, uh, b, UserIndex)
            Case e_TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
                Call HechizoPropNPC(uh, .flags.TargetNPC.ArrayIndex, UserIndex, b, IsAlive)
            Case e_TipoHechizo.uPhysicalSkill
                b = HandlePhysicalSkill(UserIndex, eUser, .flags.TargetNPC.ArrayIndex, eNpc, .Stats.UserHechizos(.flags.Hechizo), IsAlive)
        End Select
        If b Then
            If Hechizos(uh).EotId > 0 And IsAlive Then
                If Effect Is Nothing Then
                    Call CreateEffect(UserIndex, eUser, .flags.TargetNPC.ArrayIndex, eNpc, Hechizos(uh).EotId)
                Else
                    If Not Effect.Reset(UserIndex, eUser, Hechizos(uh).EotId) Then
                        Exit Sub
                    End If
                End If
            End If
            If Not IsSet(Hechizos(uh).SpellRequirementMask, eIsSkill) Then
                Call SubirSkill(UserIndex, Magia)
            End If
            .Stats.MinMAN = .Stats.MinMAN - GetSpellManaCostModifierByClass(UserIndex, Hechizos(uh), uh)
            If Hechizos(uh).RequiredHP > 0 Then
                If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
                Call UserMod.ModifyHealth(UserIndex, -Hechizos(uh).RequiredHP, 1)
            End If
            .Stats.MinSta = .Stats.MinSta - Hechizos(uh).StaRequerido
            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            Call WriteUpdateMana(UserIndex)
            Call WriteUpdateSta(UserIndex)
        End If
    End With
    Exit Sub
HandleHechizoNPC_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HandleHechizoNPC", Erl)
End Sub

Sub LanzarHechizo(ByVal Index As Integer, ByVal UserIndex As Integer)
    On Error GoTo LanzarHechizo_Err
    Dim uh               As Integer
    Dim SpellCastSuccess As Boolean
    uh = UserList(UserIndex).Stats.UserHechizos(Index)
    If PuedeLanzar(UserIndex, uh, Index) Then
        Select Case Hechizos(uh).Target
            Case e_TargetType.uUsuarios
                If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then
                    If Abs(UserList(UserList(UserIndex).flags.TargetUser.ArrayIndex).pos.y - UserList(UserIndex).pos.y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(UserIndex, uh)
                        SpellCastSuccess = True
                    Else
                        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    'Msg790= Este hechizo actua solo sobre usuarios.
                    Call WriteLocaleMsg(UserIndex, 790, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_TargetType.uNPC
                If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
                    If Abs(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).pos.y - UserList(UserIndex).pos.y) <= RANGO_VISION_Y Then
                        Call HandleHechizoNPC(UserIndex, uh)
                        SpellCastSuccess = True
                    Else
                        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    'Msg791= Este hechizo solo afecta a los npcs.
                    Call WriteLocaleMsg(UserIndex, 791, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_TargetType.uUsuariosYnpc
                If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then
                    If Abs(UserList(UserList(UserIndex).flags.TargetUser.ArrayIndex).pos.y - UserList(UserIndex).pos.y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(UserIndex, uh)
                        SpellCastSuccess = True
                    Else
                        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                ElseIf IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
                    If Abs(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).pos.y - UserList(UserIndex).pos.y) <= RANGO_VISION_Y Then
                        SpellCastSuccess = True
                        Call HandleHechizoNPC(UserIndex, uh)
                    Else
                        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    'Msg792= Target invalido.
                    Call WriteLocaleMsg(UserIndex, 792, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_TargetType.uTerreno
                SpellCastSuccess = True
                Call HandleHechizoTerreno(UserIndex, uh)
            Case e_TargetType.uPets
                SpellCastSuccess = HandlePetSpell(UserIndex, uh)
        End Select
    End If
    If SpellCastSuccess Then
        If Hechizos(uh).Cooldown > 0 Then
            UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCountRaw()
            If Hechizos(uh).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(uh).CdEffectId, -uh, CLng(Hechizos(uh).Cooldown) * 1000, CLng(Hechizos( _
                    uh).Cooldown) * 1000, eCD)
        End If
        If IsSet(UserList(UserIndex).flags.StatusMask, e_StatusMask.eTransformed) Then
            If UserList(UserIndex).Char.CastAnimation > 0 Then
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageDoAnimation(UserList(UserIndex).Char.charindex, UserList(UserIndex).Char.CastAnimation))
            End If
        End If
        If Hechizos(uh).TargetEffectType = e_TargetEffectType.eNegative Then
            If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then
                Call RegisterNewAttack(UserList(UserIndex).flags.TargetUser.ArrayIndex, UserIndex)
                If IsFeatureEnabled("remove-inv-on-attack") Then
                    Call RemoveUserInvisibility(UserIndex)
                End If
            End If
        ElseIf Hechizos(uh).TargetEffectType = ePositive Then
            If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then Call RegisterNewHelp(UserList(UserIndex).flags.TargetUser.ArrayIndex, UserIndex)
        End If
        Call ClearUserRef(UserList(UserIndex).flags.TargetUser)
        Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
    End If
    If UserList(UserIndex).Counters.Trabajando Then
        Call WriteMacroTrabajoToggle(UserIndex, False)
    End If
    If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    Exit Sub
LanzarHechizo_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.LanzarHechizo", Erl)
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 02/01/2008
    'Handles the Spells that afect the Stats of an User
    '24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
    '26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
    '26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
    '02/01/2008 Marcos (ByVal) - Curar Veneno no permitido en usuarios muertos.
    '***************************************************
    On Error GoTo HechizoEstadoUsuario_Err
    Dim h As Integer, targetUserIndex As Integer
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    targetUserIndex = UserList(UserIndex).flags.TargetUser.ArrayIndex
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Invisibility) Then
        If UserList(targetUserIndex).flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        If UserList(UserIndex).flags.EnReto Then
            'Msg793= No podés lanzar invisibilidad durante un reto.
            Call WriteLocaleMsg(UserIndex, 793, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Montado Then
            'Msg794= No podés lanzar invisibilidad mientras usas una montura.
            Call WriteLocaleMsg(UserIndex, 794, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(targetUserIndex).flags.Montado Then
            'Msg795= No podés lanzar invisibilidad a alguien montado.
            Call WriteLocaleMsg(UserIndex, 795, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(targetUserIndex).Counters.Saliendo Then
            If UserIndex <> targetUserIndex Then
                ' Msg666=¡El hechizo no tiene efecto!
                Call WriteLocaleMsg(UserIndex, MSG_NPC_INMUNE_TO_SPELLS, e_FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                ' Msg667=¡No podés ponerte invisible mientras te encuentres saliendo!
                Call WriteLocaleMsg(UserIndex, 667, e_FontTypeNames.FONTTYPE_WARNING)
                b = False
                Exit Sub
            End If
        End If
        If Not PeleaSegura(UserIndex, targetUserIndex) Then
            Select Case Status(UserIndex)
                Case 1, 3, 5 'Ciudadano o armada
                    If Status(targetUserIndex) <> e_Facciones.Ciudadano And Status(targetUserIndex) <> e_Facciones.Armada And Status(targetUserIndex) <> e_Facciones.consejo Then
                        If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                            ' Msg662=No puedes ayudar criminales.
                            Call WriteLocaleMsg(UserIndex, 662, e_FontTypeNames.FONTTYPE_INFO)
                            b = False
                            Exit Sub
                        ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then
                            If UserList(UserIndex).flags.Seguro = True Then
                                ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                Call WriteLocaleMsg(UserIndex, 663, e_FontTypeNames.FONTTYPE_INFO)
                                b = False
                                Exit Sub
                            Else
                                'Si tiene clan
                                If UserList(UserIndex).GuildIndex > 0 Then
                                    'Si el clan es de alineación ciudadana.
                                    If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                        'No lo dejo resucitarlo
                                        ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                        Call WriteLocaleMsg(UserIndex, 664, e_FontTypeNames.FONTTYPE_INFO)
                                        b = False
                                        Exit Sub
                                        'Si es de alineación neutral, lo dejo resucitar y lo vuelvo criminal
                                    ElseIf GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                                        Call VolverCriminal(UserIndex)
                                        Call RefreshCharStatus(UserIndex)
                                    End If
                                Else
                                    Call VolverCriminal(UserIndex)
                                    Call RefreshCharStatus(UserIndex)
                                End If
                            End If
                        End If
                    End If
                Case 2, 4 'Caos
                    If Status(targetUserIndex) <> e_Facciones.Caos And Status(targetUserIndex) <> e_Facciones.Criminal And Status(targetUserIndex) <> e_Facciones.concilio Then
                        'Msg796= No podés ayudar ciudadanos.
                        Call WriteLocaleMsg(UserIndex, 796, e_FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
            End Select
        End If
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And e_PlayerType.User Then
            If Not UserList(targetUserIndex).flags.Privilegios And e_PlayerType.User Then
                Exit Sub
            End If
        End If
        If MapInfo(UserList(targetUserIndex).pos.Map).SinInviOcul Then
            'Msg797= Una fuerza divina te impide usar invisibilidad en esta zona.
            Call WriteLocaleMsg(UserIndex, 797, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(targetUserIndex).flags.invisible = 1 Or UserList(targetUserIndex).Counters.DisabledInvisibility > 0 Then
            If targetUserIndex = UserIndex Then
                'Msg798= ¡Ya estás invisible!
                Call WriteLocaleMsg(UserIndex, 798, e_FontTypeNames.FONTTYPE_INFO)
            Else
                'Msg799= ¡El objetivo ya se encuentra invisible!
                Call WriteLocaleMsg(UserIndex, 799, e_FontTypeNames.FONTTYPE_INFO)
            End If
            b = False
            Exit Sub
        End If
        If IsSet(UserList(targetUserIndex).flags.StatusMask, eTaunting) Then
            If targetUserIndex = UserIndex Then
                'Msg800= ¡No podes ocultarte en este momento!
                Call WriteLocaleMsg(UserIndex, 800, e_FontTypeNames.FONTTYPE_INFO)
            Else
                'Msg801= ¡El objetivo no puede ocultarse!
                Call WriteLocaleMsg(UserIndex, 801, e_FontTypeNames.FONTTYPE_INFO)
            End If
            b = False
            Exit Sub
        End If
        UserList(targetUserIndex).flags.invisible = 1
        'Ladder
        'Reseteamos el contador de Invisibilidad
        'Le agrego un random al tiempo de invisibilidad de 16 a 21 segundos.
        If UserList(targetUserIndex).Counters.Invisibilidad <= 0 Then UserList(targetUserIndex).Counters.Invisibilidad = RandomNumber(Hechizos(h).Duration - 4, Hechizos( _
                h).Duration + 1)
        Call WriteContadores(targetUserIndex)
        Call SendData(SendTarget.ToPCArea, targetUserIndex, PrepareMessageSetInvisible(UserList(targetUserIndex).Char.charindex, True, UserList(targetUserIndex).pos.x, UserList( _
                targetUserIndex).pos.y))
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If Hechizos(h).EotId > 0 Then
        b = True
        Call InfoHechizo(UserIndex)
        Exit Sub
    End If
    If Hechizos(h).Mimetiza = 1 Then
        If UserList(UserIndex).flags.EnReto Then
            'Msg802= No podés mimetizarte durante un reto.
            Call WriteLocaleMsg(UserIndex, 802, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(targetUserIndex).flags.Muerto = 1 Then
            Exit Sub
        End If
        If UserList(targetUserIndex).flags.Navegando = 1 Then
            Exit Sub
        End If
        If UserList(UserIndex).flags.Navegando = 1 Then
            Exit Sub
        End If
        'Si sos user, no uses este hechizo con GMS.
        If Not EsGM(UserIndex) And EsGM(targetUserIndex) Then Exit Sub
        ' Si te mimetizaste, no importa si como bicho o User...
        If UserList(UserIndex).flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
            'Msg803= Ya te encuentras transformado. El hechizo no tuvo efecto
            Call WriteLocaleMsg(UserIndex, 803, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
        'copio el char original al mimetizado
        With UserList(UserIndex)
            .CharMimetizado.body = .Char.body
            .CharMimetizado.head = .Char.head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            .CharMimetizado.CartAnim = .Char.CartAnim
            .flags.Mimetizado = e_EstadoMimetismo.FormaUsuario
            'ahora pongo local el del enemigo
            .Char.body = UserList(targetUserIndex).Char.body
            .Char.head = UserList(targetUserIndex).Char.head
            .Char.CascoAnim = UserList(targetUserIndex).Char.CascoAnim
            .Char.ShieldAnim = UserList(targetUserIndex).Char.ShieldAnim
            .Char.WeaponAnim = UserList(targetUserIndex).Char.WeaponAnim
            .Char.CartAnim = UserList(targetUserIndex).Char.CartAnim
            .NameMimetizado = UserList(targetUserIndex).name
            If UserList(targetUserIndex).GuildIndex > 0 Then .NameMimetizado = .NameMimetizado & " <" & modGuilds.GuildName(UserList(targetUserIndex).GuildIndex) & ">"
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
            Call RefreshCharStatus(UserIndex)
        End With
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If Hechizos(h).Envenena > 0 Then
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        If UserList(targetUserIndex).flags.Envenenado = 0 Then
            If UserIndex <> targetUserIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
            End If
            UserList(targetUserIndex).flags.Envenenado = Hechizos(h).Envenena
            UserList(targetUserIndex).Counters.Veneno = Hechizos(h).Duration
            Call InfoHechizo(UserIndex)
            b = True
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1870, UserList(targetUserIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1870=¬1 ya está envenenado. El hechizo no tuvo efecto.
            b = False
        End If
    End If
    If Hechizos(h).desencantar = 1 Then
        ' Call WriteConsoleMsg(UserIndex, "Ningun efecto magico tiene efecto sobre ti ya.", e_FontTypeNames.FONTTYPE_INFOIAO)
        UserList(UserIndex).flags.Envenenado = 0
        UserList(UserIndex).Counters.Veneno = 0
        UserList(UserIndex).flags.Incinerado = 0
        UserList(UserIndex).Counters.Incineracion = 0
        If UserList(UserIndex).flags.Inmovilizado > 0 Then
            UserList(UserIndex).Counters.Inmovilizado = 0
            UserList(UserIndex).flags.Inmovilizado = 0
            If UserList(UserIndex).clase = e_Class.Warrior Or UserList(UserIndex).clase = e_Class.Hunter Or UserList(UserIndex).clase = e_Class.Thief Or UserList( _
                    UserIndex).clase = e_Class.Pirat Then
                UserList(UserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            Call WriteInmovilizaOK(UserIndex)
        End If
        If UserList(UserIndex).flags.Paralizado > 0 Then
            UserList(UserIndex).Counters.Paralisis = 0
            UserList(UserIndex).flags.Paralizado = 0
            If UserList(UserIndex).clase = e_Class.Warrior Or UserList(UserIndex).clase = e_Class.Hunter Or UserList(UserIndex).clase = e_Class.Thief Or UserList( _
                    UserIndex).clase = e_Class.Pirat Then
                UserList(UserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            Call WriteParalizeOK(UserIndex)
        End If
        If UserList(UserIndex).flags.Ceguera > 0 Then
            UserList(UserIndex).Counters.Ceguera = 0
            UserList(UserIndex).flags.Ceguera = 0
            Call WriteBlindNoMore(UserIndex)
        End If
        If UserList(UserIndex).flags.Maldicion > 0 Then
            UserList(UserIndex).flags.Maldicion = 0
            UserList(UserIndex).Counters.Maldicion = 0
        End If
        If UserList(UserIndex).flags.Estupidez > 0 Then
            UserList(UserIndex).flags.Estupidez = 0
            UserList(UserIndex).Counters.Estupidez = 0
        End If
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Incinerate) Then
        If UserIndex = targetUserIndex Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        If UserIndex <> targetUserIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        End If
        UserList(targetUserIndex).flags.Incinerado = 1
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveDebuff) Then
        Dim NegativeEffect As IBaseEffectOverTime
        Set NegativeEffect = EffectsOverTime.FindEffectOfTypeOnTarget(UserList(targetUserIndex).EffectOverTime, eDebuff)
        If Not NegativeEffect Is Nothing Then
            NegativeEffect.RemoveMe = True
            Call InfoHechizo(UserIndex)
            b = True
            Exit Sub
        End If
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.StealBuff) Then
        Dim TargetBuff As IBaseEffectOverTime
        Set TargetBuff = EffectsOverTime.FindEffectOfTypeOnTarget(UserList(targetUserIndex).EffectOverTime, eBuff)
        If Not TargetBuff Is Nothing Then
            Call EffectsOverTime.ChangeOwner(targetUserIndex, eUser, UserIndex, eUser, TargetBuff)
        End If
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.CurePoison) Then
        'Verificamos que el usuario no este muerto
        If UserList(targetUserIndex).flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        ' Si no esta envenenado, no hay nada mas que hacer
        If UserList(targetUserIndex).flags.Envenenado = 0 Then
            Call WriteLocaleMsg(UserIndex, 1871, e_FontTypeNames.FONTTYPE_INFOIAO, UserList(targetUserIndex).name) ' Msg1871=¬1 no está envenenado, el hechizo no tiene efecto.
            b = False
            Exit Sub
        End If
        'Para poder tirar curar veneno a un pk en el ring
        If Not PeleaSegura(UserIndex, targetUserIndex) Then
            If Status(targetUserIndex) = 0 And Status(UserIndex) = 1 Or Status(targetUserIndex) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        End If
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And e_PlayerType.User Then
            If Not UserList(targetUserIndex).flags.Privilegios And e_PlayerType.User Then
                Exit Sub
            End If
        End If
        UserList(targetUserIndex).flags.Envenenado = 0
        UserList(targetUserIndex).Counters.Veneno = 0
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Curse) Then
        If UserIndex = targetUserIndex Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        If UserIndex <> targetUserIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        End If
        UserList(targetUserIndex).flags.Maldicion = 1
        UserList(targetUserIndex).Counters.Maldicion = 200
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveCurse) Then
        UserList(targetUserIndex).flags.Maldicion = 0
        UserList(targetUserIndex).Counters.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.PreciseHit) Then
        UserList(targetUserIndex).flags.GolpeCertero = 1
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Paralize) Then
        If UserIndex = targetUserIndex Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If UserList(targetUserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas > 0 Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1872, UserList(targetUserIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1872=¬1 no puede volver a ser paralizado tan rápido.
            Exit Sub
        End If
        If Not UserMod.CanMove(UserList(targetUserIndex).flags, UserList(targetUserIndex).Counters) Then
            ' Msg661=No podes inmovilizar un objetivo que no puede moverse.
            Call WriteLocaleMsg(UserIndex, 661, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If IsSet(UserList(targetUserIndex).flags.StatusMask, eCCInmunity) Then
            Call WriteLocaleMsg(UserIndex, MsgCCInunity, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        If UserIndex <> targetUserIndex Then
            Call checkHechizosEfectividad(UserIndex, targetUserIndex)
            Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        End If
        Call InfoHechizo(UserIndex)
        b = True
        If UserList(targetUserIndex).clase = Warrior Or UserList(targetUserIndex).clase = Hunter Then
            UserList(targetUserIndex).Counters.Paralisis = Hechizos(h).Duration * 0.7
        Else
            UserList(targetUserIndex).Counters.Paralisis = Hechizos(h).Duration
        End If
        If UserList(targetUserIndex).flags.Paralizado = 0 Then
            UserList(targetUserIndex).flags.Paralizado = 1
            Call WriteParalizeOK(targetUserIndex)
            Call WritePosUpdate(targetUserIndex)
        End If
    End If
    If Hechizos(h).velocidad <> 0 Then
        'Verificamos que el usuario no este muerto
        If UserList(targetUserIndex).flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        If Hechizos(h).velocidad < 1 Then
            If UserIndex = targetUserIndex Then
                Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        Else
            'Para poder tirar curar veneno a un pk en el ring
            If Not PeleaSegura(UserIndex, targetUserIndex) Then
                If Status(targetUserIndex) = 0 And Status(UserIndex) = 1 Or Status(targetUserIndex) = 2 And Status(UserIndex) = 1 Then
                    If esArmada(UserIndex) Then
                        Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                    If UserList(UserIndex).flags.Seguro Then
                        Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                End If
            End If
            'Si sos user, no uses este hechizo con GMS.
            If UserList(UserIndex).flags.Privilegios And e_PlayerType.User Then
                If Not UserList(targetUserIndex).flags.Privilegios And e_PlayerType.User Then
                    Exit Sub
                End If
            End If
        End If
        Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        Call InfoHechizo(UserIndex)
        b = True
        If UserList(targetUserIndex).Counters.velocidad = 0 Then
            UserList(targetUserIndex).flags.VelocidadHechizada = Hechizos(h).velocidad
            Call ActualizarVelocidadDeUsuario(targetUserIndex)
        End If
        UserList(targetUserIndex).Counters.velocidad = Hechizos(h).Duration
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Immobilize) Then
        If UserIndex = targetUserIndex Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not UserMod.CanMove(UserList(targetUserIndex).flags, UserList(targetUserIndex).Counters) Then
            ' Msg661=No podes inmovilizar un objetivo que no puede moverse.
            Call WriteLocaleMsg(UserIndex, 661, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If UserList(targetUserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas > 0 Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1873, UserList(targetUserIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1873=¬1 no puede volver a ser inmovilizado tan rápido.
            Exit Sub
        End If
        If IsSet(UserList(targetUserIndex).flags.StatusMask, eCCInmunity) Then
            Call WriteLocaleMsg(UserIndex, MsgCCInunity, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        If UserIndex <> targetUserIndex Then
            Call checkHechizosEfectividad(UserIndex, targetUserIndex)
            Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        End If
        Call InfoHechizo(UserIndex)
        b = True
        If UserList(targetUserIndex).clase = Warrior Or UserList(targetUserIndex).clase = Hunter Then
            UserList(targetUserIndex).Counters.Inmovilizado = Hechizos(h).Duration * 0.7
        Else
            UserList(targetUserIndex).Counters.Inmovilizado = Hechizos(h).Duration
        End If
        UserList(targetUserIndex).flags.Inmovilizado = 1
        Call WriteInmovilizaOK(targetUserIndex)
        Call WritePosUpdate(targetUserIndex)
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveParalysis) Then
        'Para poder tirar remo a un pk en el ring
        If Not PeleaSegura(UserIndex, targetUserIndex) Then
            Select Case Status(UserIndex)
                Case 1, 3, 5 'Ciudadano o armada
                    If Status(targetUserIndex) <> e_Facciones.Ciudadano And Status(targetUserIndex) <> e_Facciones.Armada And Status(targetUserIndex) <> e_Facciones.consejo Then
                        If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                            ' Msg662=No puedes ayudar criminales.
                            Call WriteLocaleMsg(UserIndex, 662, e_FontTypeNames.FONTTYPE_INFO)
                            b = False
                            Exit Sub
                        ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then
                            If UserList(UserIndex).flags.Seguro = True Then
                                ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                Call WriteLocaleMsg(UserIndex, 663, e_FontTypeNames.FONTTYPE_INFO)
                                b = False
                                Exit Sub
                            Else
                                'Si tiene clan
                                If UserList(UserIndex).GuildIndex > 0 Then
                                    'Si el clan es de alineación ciudadana.
                                    If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                        'No lo dejo resucitarlo
                                        ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                        Call WriteLocaleMsg(UserIndex, 664, e_FontTypeNames.FONTTYPE_INFO)
                                        b = False
                                        Exit Sub
                                        'Si es de alineación neutral, lo dejo resucitar y lo vuelvo criminal
                                    ElseIf GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                                        Call VolverCriminal(UserIndex)
                                        Call RefreshCharStatus(UserIndex)
                                    End If
                                Else
                                    Call VolverCriminal(UserIndex)
                                    Call RefreshCharStatus(UserIndex)
                                End If
                            End If
                        End If
                    End If
                Case 2, 4 'Caos
                    If Status(targetUserIndex) <> e_Facciones.Caos And Status(targetUserIndex) <> e_Facciones.Criminal And Status(targetUserIndex) <> e_Facciones.concilio Then
                        'Msg805= No podés ayudar ciudadanos.
                        Call WriteLocaleMsg(UserIndex, 805, e_FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
            End Select
        End If
        If UserList(targetUserIndex).flags.Inmovilizado = 0 And UserList(targetUserIndex).flags.Paralizado = 0 Then
            'Msg806= El objetivo no esta paralizado.
            Call WriteLocaleMsg(UserIndex, 806, e_FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        If UserList(targetUserIndex).flags.Inmovilizado = 1 Then
            UserList(targetUserIndex).Counters.Inmovilizado = 0
            If UserList(targetUserIndex).clase = e_Class.Warrior Or UserList(targetUserIndex).clase = e_Class.Hunter Or UserList(targetUserIndex).clase = e_Class.Thief Or _
                    UserList(targetUserIndex).clase = e_Class.Pirat Then
                UserList(targetUserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            UserList(targetUserIndex).flags.Inmovilizado = 0
            Call WriteInmovilizaOK(targetUserIndex)
            Call WritePosUpdate(targetUserIndex)
        End If
        If UserList(targetUserIndex).flags.Paralizado = 1 Then
            UserList(targetUserIndex).flags.Paralizado = 0
            If UserList(targetUserIndex).clase = e_Class.Warrior Or UserList(targetUserIndex).clase = e_Class.Hunter Or UserList(targetUserIndex).clase = e_Class.Thief Or _
                    UserList(targetUserIndex).clase = e_Class.Pirat Then
                UserList(targetUserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            UserList(targetUserIndex).Counters.Paralisis = 0
            Call WriteParalizeOK(targetUserIndex)
        End If
        b = True
        Call InfoHechizo(UserIndex)
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveDumb) Then
        If UserList(targetUserIndex).flags.Estupidez = 1 Then
            'Para poder tirar remo estu a un pk en el ring
            If Not PeleaSegura(UserIndex, targetUserIndex) Then
                If Status(targetUserIndex) = 0 And Status(UserIndex) = 1 Or Status(targetUserIndex) = 2 And Status(UserIndex) = 1 Then
                    If esArmada(UserIndex) Then
                        Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                    If UserList(UserIndex).flags.Seguro Then
                        Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                End If
            End If
            UserList(targetUserIndex).flags.Estupidez = 0
            UserList(targetUserIndex).Counters.Estupidez = 0
            Call WriteDumbNoMore(targetUserIndex)
            Call InfoHechizo(UserIndex)
            b = True
        End If
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Resurrect) Then
        If UserList(targetUserIndex).flags.Muerto = 1 Then
            If UserList(UserIndex).flags.EnReto Then
                'Msg807= No podés revivir a nadie durante un reto.
                Call WriteLocaleMsg(UserIndex, 807, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).clase <> Cleric Then
                Dim PuedeRevivir As Boolean
                If UserList(UserIndex).invent.EquippedWeaponObjIndex <> 0 Then
                    If ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).Revive Then
                        PuedeRevivir = True
                    End If
                End If
                If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex <> 0 Then
                    If ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).Revive Then
                        PuedeRevivir = True
                    End If
                End If
                If UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex <> 0 Then
                    If ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).Revive Then
                        PuedeRevivir = True
                    End If
                End If
                If Not PuedeRevivir Then
                    'Msg809= Necesitás un objeto con mayor poder mágico para poder revivir.
                    Call WriteLocaleMsg(UserIndex, 809, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
            If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
                'Msg810= Deberás tener la barra de energía llena para poder resucitar.
                Call WriteLocaleMsg(UserIndex, 810, e_FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
            'Para poder tirar revivir a un pk en el ring
            If Not PeleaSegura(UserIndex, targetUserIndex) Then
                If UserList(targetUserIndex).flags.SeguroResu Then
                    ' Msg693=El usuario tiene el seguro de resurrección activado.
                    Call WriteLocaleMsg(UserIndex, 693, e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(targetUserIndex, PrepareMessageLocaleMsg(1874, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1874=¬1 está intentando revivirte. Desactiva el seguro de resurrección para permitirle hacerlo.
                    b = False
                    Exit Sub
                End If
                Select Case Status(UserIndex)
                    Case 1, 3, 5 'Ciudadano o armada
                        If Status(targetUserIndex) <> e_Facciones.Ciudadano And Status(targetUserIndex) <> e_Facciones.Armada And Status(targetUserIndex) <> e_Facciones.consejo _
                                Then
                            If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                                'Msg811= Los miembros de la armada real solo pueden revivir ciudadanos a miembros de su facción.
                                Call WriteLocaleMsg(UserIndex, 811, e_FontTypeNames.FONTTYPE_INFO)
                                b = False
                                Exit Sub
                            ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then
                                If UserList(UserIndex).flags.Seguro = True Then
                                    'Msg812= Deberás desactivar el seguro para revivir al usuario, ten en cuenta que te convertirás en criminal.
                                    Call WriteLocaleMsg(UserIndex, 812, e_FontTypeNames.FONTTYPE_INFO)
                                    b = False
                                    Exit Sub
                                Else
                                    'Si tiene clan
                                    If UserList(UserIndex).GuildIndex > 0 Then
                                        'Si el clan es de alineación ciudadana.
                                        If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                            'No lo dejo resucitarlo
                                            'Msg813= No puedes resucitar al usuario siendo fundador de un clan ciudadano.
                                            Call WriteLocaleMsg(UserIndex, 813, e_FontTypeNames.FONTTYPE_INFO)
                                            b = False
                                            Exit Sub
                                            'Si es de alineación neutral, lo dejo resucitar y lo vuelvo criminal
                                        ElseIf GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                                            Call VolverCriminal(UserIndex)
                                            Call RefreshCharStatus(UserIndex)
                                        End If
                                    Else
                                        Call VolverCriminal(UserIndex)
                                        Call RefreshCharStatus(UserIndex)
                                    End If
                                End If
                            End If
                        End If
                    Case 2, 4 'Caos
                        If Status(targetUserIndex) <> e_Facciones.Caos And Status(targetUserIndex) <> e_Facciones.Criminal And Status(targetUserIndex) <> e_Facciones.concilio Then
                            'Msg814= Los miembros del caos solo pueden revivir criminales o miembros de su facción.
                            Call WriteLocaleMsg(UserIndex, 814, e_FontTypeNames.FONTTYPE_INFO)
                            b = False
                            Exit Sub
                        End If
                End Select
            End If
            Call SendData(SendTarget.ToPCArea, targetUserIndex, PrepareMessageSetInvisible(UserList(targetUserIndex).Char.charindex, False, UserList(targetUserIndex).pos.x, _
                    UserList(targetUserIndex).pos.y))
            Call ResurrectUser(targetUserIndex, UserIndex)
            Call InfoHechizo(UserIndex)
            b = True
        Else
            b = False
        End If
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Blindness) Then
        If UserIndex = targetUserIndex Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        If UserIndex <> targetUserIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        End If
        UserList(targetUserIndex).flags.Ceguera = 1
        UserList(targetUserIndex).Counters.Ceguera = Hechizos(h).Duration
        Call WriteBlind(targetUserIndex)
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Dumb) Then
        If UserIndex = targetUserIndex Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        If UserIndex <> targetUserIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        End If
        If UserList(targetUserIndex).flags.Estupidez = 0 Then
            UserList(targetUserIndex).flags.Estupidez = 1
            UserList(targetUserIndex).Counters.Estupidez = Hechizos(h).Duration
        End If
        Call WriteDumb(targetUserIndex)
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.ToggleCleave) Then
        If UserList(UserIndex).flags.Cleave Then
            UserList(UserIndex).flags.Cleave = 0
            If Hechizos(h).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(h).CdEffectId, -1, 0, 0, eBuff)
        Else
            UserList(UserIndex).flags.Cleave = 1
            If Hechizos(h).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(h).CdEffectId, -1, -1, -1, eBuff)
        End If
        b = True
    End If
    Dim Character As t_User
    Character = UserList(UserIndex)
    If IsFeatureEnabled("healers_and_tanks") And IsSet(Hechizos(h).Effects, e_SpellEffects.ToggleDivineBlood) Then
        If UserList(UserIndex).flags.DivineBlood Then
            UserList(UserIndex).flags.DivineBlood = 0
            If Hechizos(h).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(h).CdEffectId, -1, 0, 0, eBuff)
            Character.Char.BackpackAnim = 0
            Call ChangeUserChar(UserIndex, Character.Char.body, Character.Char.head, Character.Char.Heading, Character.Char.WeaponAnim, Character.Char.ShieldAnim, _
                    Character.Char.CascoAnim, Character.Char.CartAnim, Character.Char.BackpackAnim)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(Character.Char.charindex, e_ParticleEffects.HaloGold, 45, False, 0, Character.pos.x, _
                    Character.pos.y))
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundEffects.BAOLegionHorn, Character.pos.x, Character.pos.y))
        Else
            UserList(UserIndex).flags.DivineBlood = 1
            If Hechizos(h).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(h).CdEffectId, -1, -1, -1, eBuff)
            Character.Char.BackpackAnim = 4997
            Call ChangeUserChar(UserIndex, Character.Char.body, Character.Char.head, Character.Char.Heading, Character.Char.WeaponAnim, Character.Char.ShieldAnim, _
                    Character.Char.CascoAnim, Character.Char.CartAnim, Character.Char.BackpackAnim)
            Character.Counters.timeFx = 3
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(Character.Char.charindex, e_GraphicEffects.CurarHeridasCriticasBAO, 0, Character.pos.x, _
                    Character.pos.y))
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundEffects.DivineBloodActivation, Character.pos.x, Character.pos.y))
        End If
        b = True
    End If
    Exit Sub
HechizoEstadoUsuario_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoEstadoUsuario", Erl)
End Sub

Sub checkHechizosEfectividad(ByVal UserIndex As Integer, ByVal TargetUser As Integer)
    With UserList(UserIndex)
        If UserList(TargetUser).flags.Inmovilizado + UserList(TargetUser).flags.Paralizado = 0 Then
            .Counters.controlHechizos.HechizosCasteados = .Counters.controlHechizos.HechizosCasteados + 1
            Dim efectividad As Double
            efectividad = (100 * .Counters.controlHechizos.HechizosCasteados) / .Counters.controlHechizos.HechizosTotales
            If efectividad >= 50 And .Counters.controlHechizos.HechizosTotales >= 6 Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageLocaleMsg(1638, .name & "¬" & efectividad & "¬" & .Counters.controlHechizos.HechizosCasteados & "¬" & _
                        .Counters.controlHechizos.HechizosTotales, e_FontTypeNames.FONTTYPE_TALK)) 'Msg1638=El usuario ¬1 está lanzando hechizos con una efectividad de ¬2% (Casteados: ¬3/¬4), revisar.
            End If
            Debug.Print "El usuario " & .name & " está lanzando hechizos con una efectividad de " & efectividad & "% (Casteados: " & .Counters.controlHechizos.HechizosCasteados _
                    & "/" & .Counters.controlHechizos.HechizosTotales & "), revisar."
        Else
            .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales - 1
        End If
    End With
End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
    On Error GoTo HechizoEstadoNPC_Err
    If NpcList(NpcIndex).flags.ImmuneToSpells <> 0 Then
        If UserIndex > 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_NPC_INMUNE_TO_SPELLS, e_FontTypeNames.FONTTYPE_INFO)
        End If
        b = False
        Exit Sub
    End If
    Dim UserAttackInteractionResult As t_AttackInteractionResult
    If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.Invisibility) Then
        Call InfoHechizo(UserIndex)
        NpcList(NpcIndex).flags.invisible = 1
        b = True
    End If
    If Hechizos(hIndex).Envenena > 0 Then
        UserAttackInteractionResult = UserCanAttackNpc(UserIndex, NpcIndex)
        Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
        If UserAttackInteractionResult.CanAttack Then
            If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
        Else
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        Call InfoHechizo(UserIndex)
        NpcList(NpcIndex).flags.Envenenado = Hechizos(hIndex).Envenena
        b = True
    End If
    If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.RemoveDebuff) Then
        Dim NegativeEffect As IBaseEffectOverTime
        Set NegativeEffect = EffectsOverTime.FindEffectOfTypeOnTarget(NpcList(NpcIndex).EffectOverTime, eDebuff)
        If Not NegativeEffect Is Nothing Then
            NegativeEffect.RemoveMe = True
            Call InfoHechizo(UserIndex)
            b = True
            Exit Sub
        End If
    End If
    If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.StealBuff) Then
        Dim TargetBuff As IBaseEffectOverTime
        Set TargetBuff = EffectsOverTime.FindEffectOfTypeOnTarget(NpcList(NpcIndex).EffectOverTime, eBuff)
        If Not TargetBuff Is Nothing Then
            Call EffectsOverTime.ChangeOwner(NpcIndex, eNpc, UserIndex, eUser, TargetBuff)
        End If
        Call InfoHechizo(UserIndex)
        b = True
        Exit Sub
    End If
    If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.CurePoison) Then
        If NpcList(NpcIndex).flags.Envenenado > 0 Then
            Call InfoHechizo(UserIndex)
            NpcList(NpcIndex).flags.Envenenado = 0
            b = True
        Else
            'Msg815= La criatura no esta envenenada, el hechizo no tiene efecto.
            Call WriteLocaleMsg(UserIndex, 815, e_FontTypeNames.FONTTYPE_INFOIAO)
            b = False
        End If
    End If
    If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.RemoveCurse) Then
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.Paralize) Then
        If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
            Call NPCAtacado(NpcIndex, UserIndex, True)
            Call InfoHechizo(UserIndex)
            NpcList(NpcIndex).flags.Paralizado = 1
            NpcList(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 6
            NpcList(NpcIndex).flags.Inmovilizado = 0
            NpcList(NpcIndex).Contadores.Inmovilizado = 0
            Call AnimacionIdle(NpcIndex, False)
            b = True
        Else
            Call WriteLocaleMsg(UserIndex, 381, e_FontTypeNames.FONTTYPE_INFOIAO)
            b = False
            Exit Sub
        End If
    End If
    If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.RemoveParalysis) Then
        With NpcList(NpcIndex)
            If .flags.Paralizado + .flags.Inmovilizado = 0 Then
                'Msg816= Este NPC no esta Paralizado
                Call WriteLocaleMsg(UserIndex, 816, e_FontTypeNames.FONTTYPE_INFOIAO)
                b = False
            Else
                Dim IsValidMaster As Boolean
                IsValidMaster = IsValidUserRef(.MaestroUser)
                ' Si el usuario es Armada o Caos y el NPC es de la misma faccion
                b = ((esArmada(UserIndex) Or esCaos(UserIndex)) And .flags.Faccion = UserList(UserIndex).Faccion.Status)
                'O si es mi propia mascota
                b = b Or (IsValidMaster And (.MaestroUser.ArrayIndex = UserIndex))
                'O si es mascota de otro usuario de la misma faccion
                b = b Or ((esArmada(UserIndex) And (IsValidMaster And esArmada(.MaestroUser.ArrayIndex))) Or (esCaos(UserIndex) And (IsValidMaster And esCaos( _
                        .MaestroUser.ArrayIndex))))
                If b Then
                    Call InfoHechizo(UserIndex)
                    .flags.Paralizado = 0
                    .Contadores.Paralisis = 0
                    .flags.Inmovilizado = 0
                    .Contadores.Inmovilizado = 0
                Else
                    'Msg817= Solo podés remover la Parálisis de tus mascotas o de criaturas que pertenecen a tu facción.
                    Call WriteLocaleMsg(UserIndex, 817, e_FontTypeNames.FONTTYPE_INFOIAO)
                End If
            End If
        End With
    End If
    If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.Immobilize) Then
        If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
            Call NPCAtacado(NpcIndex, UserIndex, True)
            NpcList(NpcIndex).flags.Inmovilizado = 1
            NpcList(NpcIndex).Contadores.Inmovilizado = (Hechizos(hIndex).Duration * 6.5) * 6
            NpcList(NpcIndex).flags.Paralizado = 0
            NpcList(NpcIndex).Contadores.Paralisis = 0
            Call AnimacionIdle(NpcIndex, True)
            Call InfoHechizo(UserIndex)
            b = True
        Else
            Call WriteLocaleMsg(UserIndex, 381, e_FontTypeNames.FONTTYPE_INFOIAO)
        End If
    End If
    If Hechizos(hIndex).Mimetiza = 1 Then
        If UserList(UserIndex).flags.EnReto Then
            'Msg818= No podés mimetizarte durante un reto.
            Call WriteLocaleMsg(UserIndex, 818, e_FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
            'Msg819= Ya te encuentras transformado. El hechizo no tuvo efecto
            Call WriteLocaleMsg(UserIndex, 819, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
        If UserList(UserIndex).clase = e_Class.Druid Then
            'copio el char original al mimetizado
            With UserList(UserIndex)
                .CharMimetizado.body = .Char.body
                .CharMimetizado.head = .Char.head
                .CharMimetizado.CascoAnim = .Char.CascoAnim
                .CharMimetizado.ShieldAnim = .Char.ShieldAnim
                .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                .CharMimetizado.CartAnim = .Char.CartAnim
                .flags.Mimetizado = e_EstadoMimetismo.FormaBicho
                'ahora pongo lo del NPC.
                .Char.body = NpcList(NpcIndex).Char.body
                .Char.head = NpcList(NpcIndex).Char.head
                Call ClearClothes(.Char)
                .NameMimetizado = IIf(NpcList(NpcIndex).showName = 1, NpcList(NpcIndex).name, vbNullString)
                Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                Call RefreshCharStatus(UserIndex)
            End With
        Else
            'Msg820= Solo los druidas pueden mimetizarse con criaturas.
            Call WriteLocaleMsg(UserIndex, 820, e_FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub
        End If
        Call InfoHechizo(UserIndex)
        b = True
    End If
    If Hechizos(hIndex).EotId Then
        Call InfoHechizo(UserIndex)
        b = True
    End If
    Exit Sub
HechizoEstadoNPC_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoEstadoNPC", Erl)
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean, ByRef IsAlive As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 14/08/2007
    'Handles the Spells that afect the Life NPC
    '14/08/2007 Pablo (ToxicWaste) - Orden general.
    '***************************************************
    On Error GoTo HechizoPropNPC_Err
    If NpcList(NpcIndex).flags.ImmuneToSpells <> 0 Then
        If UserIndex > 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_NPC_INMUNE_TO_SPELLS, e_FontTypeNames.FONTTYPE_INFO)
        End If
        b = False
        Exit Sub
    End If
    Dim UserAttackInteractionResult As t_AttackInteractionResult
    Dim Damage                      As Long
    Dim DamageStr                   As String
    'Salud
    If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.eDoHeal) Then
        If NpcList(NpcIndex).Stats.MinHp < NpcList(NpcIndex).Stats.MaxHp Then
            Damage = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
            Damage = Damage * UserMod.GetMagicHealingBonus(UserList(UserIndex))
            Damage = Damage * NPCs.GetSelfHealingBonus(NpcList(NpcIndex))
            If IsFeatureEnabled("elemental_tags") Then
                Call CalculateElementalTagsModifiers(UserIndex, NpcIndex, Damage)
            End If
            Call InfoHechizo(UserIndex)
            Call NPCs.DoDamageOrHeal(NpcIndex, UserIndex, eUser, Damage, e_DamageSourceType.e_magic, hIndex)
            If Damage > 0 Then
                DamageStr = PonerPuntos(Damage)
                Call WriteLocaleMsg(UserIndex, 388, e_FontTypeNames.FONTTYPE_FIGHT, "la criatura¬" & DamageStr)
            End If
            b = True
        Else
            'Msg821= La criatura no tiene heridas que curar, el hechizo no tiene efecto.
            Call WriteLocaleMsg(UserIndex, 821, e_FontTypeNames.FONTTYPE_INFOIAO)
            b = False
        End If
    ElseIf IsSet(Hechizos(hIndex).Effects, e_SpellEffects.eDoDamage) Then
        If Hechizos(hIndex).IsElementalTagsOnly Then
            If NpcList(NpcIndex).flags.ElementalTags = 0 Then
                Call WriteLocaleMsg(UserIndex, 2125, e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            End If
            If UserList(UserIndex).invent.EquippedWeaponObjIndex = 0 Then
                Call WriteLocaleMsg(UserIndex, 2126, e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            Else
                If ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).ElementalTags = 0 And UserList(UserIndex).invent.Object(UserList(UserIndex).invent.EquippedWeaponSlot).ElementalTags = 0 Then
                    Call WriteLocaleMsg(UserIndex, 2126, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
            End If
        End If
        UserAttackInteractionResult = UserCanAttackNpc(UserIndex, NpcIndex)
        Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
        If UserAttackInteractionResult.CanAttack Then
            If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
        Else
            b = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, UserIndex)
        Damage = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
        Damage = Damage + Porcentaje(Damage, 3 * UserList(UserIndex).Stats.ELV)
        Dim MagicPenetration As Integer
        If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
            MagicPenetration = ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicPenetration
            Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicAbsoluteBonus
        End If
        If UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicDamageBonus)
            Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicAbsoluteBonus
            MagicPenetration = MagicPenetration + ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicPenetration
        End If
        ' Magic Damage ring
        If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
            Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicAbsoluteBonus
            MagicPenetration = MagicPenetration + ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicPenetration
        End If
        b = True
        If NpcList(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd2, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y))
        End If
        If NpcList(NpcIndex).Stats.MagicResistance > 0 Then
            Dim DiffSkill As Integer
            DiffSkill = NpcList(NpcIndex).Stats.MagicResistance - UserList(UserIndex).Stats.UserSkills(e_Skill.Magia)
            If DiffSkill > 0 Then
                Damage = Damage - Porcentaje(Damage, max(0, (NpcList(NpcIndex).Stats.MagicDef + DiffSkill * 2) - MagicPenetration))
            Else
                Damage = Damage - Porcentaje(Damage, max(0, NpcList(NpcIndex).Stats.MagicDef - MagicPenetration))
            End If
        End If
        'Quizas tenga defenza magica el NPC.
        If Hechizos(hIndex).AntiRm = 0 Then
            Damage = Damage - NpcList(NpcIndex).Stats.defM
        End If
        Damage = Damage * UserMod.GetMagicDamageModifier(UserList(UserIndex))
        Damage = Damage * NPCs.GetMagicDamageReduction(NpcList(NpcIndex))
        If Damage < 0 Then Damage = 0
        If IsFeatureEnabled("elemental_tags") Then
            Call CalculateElementalTagsModifiers(UserIndex, NpcIndex, Damage)
        End If
        Call InfoHechizo(UserIndex)
        IsAlive = NPCs.DoDamageOrHeal(NpcIndex, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, hIndex) = eStillAlive
        If NpcList(NpcIndex).npcType = DummyTarget Then
            Call DummyTargetAttacked(NpcIndex)
        End If
    End If
    Exit Sub
HechizoPropNPC_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoPropNPC", Erl)
End Sub

Private Sub InfoHechizoDeNpcSobreUser(ByVal NpcIndex As Integer, ByVal TargetUser As Integer, ByVal Spell As Integer)
    On Error GoTo InfoHechizoDeNpcSobreUser_Err
    With UserList(TargetUser)
        If Hechizos(Spell).FXgrh > 0 Then '¿Envio FX?
            If Hechizos(Spell).ParticleViaje > 0 Then
                .Counters.timeFx = 3
                Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageParticleFXWithDestino(NpcList(NpcIndex).Char.charindex, .Char.charindex, Hechizos( _
                        Spell).ParticleViaje, Hechizos(Spell).FXgrh, Hechizos(Spell).TimeParticula, Hechizos(Spell).wav, 1, UserList(TargetUser).pos.x, UserList( _
                        TargetUser).pos.y))
            Else
                .Counters.timeFx = 3
                Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops, UserList(TargetUser).pos.x, _
                        UserList(TargetUser).pos.y))
            End If
        End If
        If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
            If Hechizos(Spell).ParticleViaje > 0 Then
                .Counters.timeFx = 3
                Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageParticleFXWithDestino(NpcList(NpcIndex).Char.charindex, .Char.charindex, Hechizos( _
                        Spell).ParticleViaje, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, Hechizos(Spell).wav, 0, UserList(TargetUser).pos.x, UserList( _
                        TargetUser).pos.y))
            Else
                .Counters.timeFx = 3
                Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageParticleFX(.Char.charindex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False, , _
                        UserList(TargetUser).pos.x, UserList(TargetUser).pos.y))
            End If
        End If
        If Hechizos(Spell).wav > 0 Then
            Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessagePlayWave(Hechizos(Spell).wav, .pos.x, .pos.y))
        End If
        If Hechizos(Spell).TimeEfect <> 0 Then
            Call WriteFlashScreen(TargetUser, Hechizos(Spell).ScreenColor, Hechizos(Spell).TimeEfect)
        End If
    End With
    Exit Sub
InfoHechizoDeNpcSobreUser_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.InfoHechizoDeNpcSobreUser", Erl)
End Sub


Private Sub InfoHechizo(ByVal UserIndex As Integer)
    On Error GoTo InfoHechizo_Err

    Dim slot As Integer
    Dim h As Integer            ' id de hechizo
    Dim skin As Integer

    ' Slot ? id (bounds-safe)
    If Not IsArrayInitialized(UserList(UserIndex).Stats.UserHechizos) Then Exit Sub
    slot = UserList(UserIndex).flags.Hechizo
    If slot < LBound(UserList(UserIndex).Stats.UserHechizos) Or _
       slot > UBound(UserList(UserIndex).Stats.UserHechizos) Then
        Call TraceError(9, "Invalid spell slot=" & slot, "modHechizos.InfoHechizo", Erl)
        Exit Sub
    End If

    h = UserList(UserIndex).Stats.UserHechizos(slot)
    If h < LBound(Hechizos) Or h > UBound(Hechizos) Then
        Call TraceError(9, "Invalid Hechizo id=" & h & " (slot=" & slot & ")", "modHechizos.InfoHechizo", Erl)
        Exit Sub
    End If

    skin = GetSkinSpellSafe(UserIndex, h)

    If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
        Call DecirPalabrasMagicas(h, UserIndex)
    End If

    If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then
        Call CastOnUser(UserIndex, h, skin, UserList(UserIndex).flags.TargetUser.ArrayIndex)
        
    ElseIf IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
    
        Call CastOnNpc(UserIndex, h, skin, UserList(UserIndex).flags.TargetNPC.ArrayIndex)
    Else
        Call CastOnTerrain(UserIndex, h, skin)
    End If

    Exit Sub
InfoHechizo_Err:
    Call TraceError(Err.Number, Err.Description & " (h=" & h & ", slot=" & slot & ")", "modHechizos.InfoHechizo", Erl)
End Sub

' --- Helper 0: skin seguro ---
Private Function GetSkinSpellSafe(ByVal UserIndex As Integer, ByVal h As Integer) As Integer
    On Error GoTo GetSkinSpellSafe_Err
    Dim res As Integer
    res = 0
    If IsArrayInitialized(UserList(UserIndex).Stats.UserSkinsHechizos) Then
        If h >= LBound(UserList(UserIndex).Stats.UserSkinsHechizos) And _
           h <= UBound(UserList(UserIndex).Stats.UserSkinsHechizos) Then
            res = UserList(UserIndex).Stats.UserSkinsHechizos(h)
        End If
    End If
    GetSkinSpellSafe = res
    Exit Function
GetSkinSpellSafe_Err:
    GetSkinSpellSafe = 0
End Function

' --- Helper 1: Target USER ---
Private Sub CastOnUser(ByVal UserIndex As Integer, ByVal h As Integer, ByVal skin As Integer, ByVal TargetIndex As Integer)
    On Error GoTo CastOnUser_Err
    Dim fxId As Integer
    ' FX
    If Hechizos(h).FXgrh > 0 Then
        UserList(TargetIndex).Counters.timeFx = 3
        If Hechizos(h).ParticleViaje > 0 Then
            Call SendData( _
               SendTarget.ToPCAliveArea, TargetIndex, _
               PrepareMessageParticleFXWithDestino( _
               UserList(UserIndex).Char.charindex, _
               UserList(TargetIndex).Char.charindex, _
               Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, _
               Hechizos(h).TimeParticula, Hechizos(h).wav, 1, _
               UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y))
        Else
            If skin > 0 Then
                fxId = skin
            Else
                fxId = Hechizos(h).FXgrh
            End If
            Call SendData( _
               SendTarget.ToPCAliveArea, TargetIndex, _
               PrepareMessageCreateFX( _
               UserList(TargetIndex).Char.charindex, fxId, Hechizos(h).loops, _
               UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y))
        End If
    End If
    ' Partículas
    If Hechizos(h).Particle > 0 Then
        UserList(TargetIndex).Counters.timeFx = 3
        If Hechizos(h).ParticleViaje > 0 Then
            Call SendData( _
               SendTarget.ToPCAliveArea, TargetIndex, _
               PrepareMessageParticleFXWithDestino( _
               UserList(TargetIndex).Char.charindex, _
               UserList(TargetIndex).Char.charindex, _
               Hechizos(h).ParticleViaje, Hechizos(h).Particle, _
               Hechizos(h).TimeParticula, Hechizos(h).wav, 0, _
               UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y))
        Else
            Call SendData( _
               SendTarget.ToPCAliveArea, TargetIndex, _
               PrepareMessageParticleFX( _
               UserList(TargetIndex).Char.charindex, _
               Hechizos(h).Particle, Hechizos(h).TimeParticula, False, , _
               UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y))
        End If
    End If
    ' Sonido (sin viaje)
    If Hechizos(h).ParticleViaje = 0 Then
        Call SendData( _
           SendTarget.ToPCAliveArea, TargetIndex, _
           PrepareMessagePlayWave(Hechizos(h).wav, _
           UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y))
    End If
    If UserIndex = TargetIndex Then
        Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, e_FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "HecMSGU*" & h & "*" & UserList(TargetIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(TargetIndex, "HecMSGA*" & h & "*" & UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)
    End If
    ' Efecto de pantalla
    If Hechizos(h).TimeEfect <> 0 Then
        Call WriteFlashScreen(UserIndex, Hechizos(h).ScreenColor, Hechizos(h).TimeEfect)
    End If

    Exit Sub
CastOnUser_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.CastOnUser", Erl)
End Sub

' --- Helper 2: Target NPC ---
Private Sub CastOnNpc(ByVal UserIndex As Integer, ByVal h As Integer, ByVal skin As Integer, ByVal TargetIndex As Integer)
    On Error GoTo CastOnNpc_Err

    Dim dead As Boolean
    Dim fxId As Integer

    dead = (NpcList(TargetIndex).Stats.MinHp < 1)

    ' FX
    If Hechizos(h).FXgrh > 0 Then
        If dead Then
            If Hechizos(h).ParticleViaje > 0 Then
                If skin > 0 Then
                    Call SendData( _
                        SendTarget.ToNPCAliveArea, TargetIndex, _
                        PrepareMessageFxPiso( _
                            skin, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
                Else
                    Call SendData( _
                        SendTarget.ToNPCAliveArea, TargetIndex, _
                        PrepareMessageParticleFXWithDestinoXY( _
                            NpcList(TargetIndex).Char.charindex, _
                            Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, _
                            Hechizos(h).TimeParticula, Hechizos(h).wav, 1, _
                            UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
                End If
            Else
                If skin > 0 Then
                    fxId = skin
                Else
                    fxId = Hechizos(h).FXgrh
                End If
                Call SendData( _
                    SendTarget.ToNPCAliveArea, TargetIndex, _
                    PrepareMessageFxPiso( _
                        fxId, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
            End If
        Else
            If Hechizos(h).ParticleViaje > 0 Then
                If skin > 0 Then
                    Call SendData( _
                        SendTarget.ToNPCAliveArea, TargetIndex, _
                        PrepareMessageCreateFX( _
                            NpcList(TargetIndex).Char.charindex, skin, Hechizos(h).loops))
                Else
                    Call SendData( _
                        SendTarget.ToNPCAliveArea, TargetIndex, _
                        PrepareMessageParticleFXWithDestino( _
                            NpcList(TargetIndex).Char.charindex, _
                            NpcList(TargetIndex).Char.charindex, _
                            Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, _
                            Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                End If
            Else
                If skin > 0 Then
                    fxId = skin
                Else
                    fxId = Hechizos(h).FXgrh
                End If
                Call SendData( _
                    SendTarget.ToNPCAliveArea, TargetIndex, _
                    PrepareMessageCreateFX( _
                        NpcList(TargetIndex).Char.charindex, fxId, Hechizos(h).loops))
            End If
        End If
    End If

    ' Partículas
    If Hechizos(h).Particle > 0 Then
        If dead Then
            Call SendData( _
                SendTarget.ToNPCAliveArea, TargetIndex, _
                PrepareMessageParticleFXWithDestinoXY( _
                    NpcList(TargetIndex).Char.charindex, _
                    Hechizos(h).ParticleViaje, Hechizos(h).Particle, _
                    Hechizos(h).TimeParticula, Hechizos(h).wav, 0, _
                    NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y))
        Else
            If Hechizos(h).ParticleViaje > 0 Then
                Call SendData( _
                    SendTarget.ToNPCAliveArea, TargetIndex, _
                    PrepareMessageParticleFXWithDestino( _
                        NpcList(TargetIndex).Char.charindex, _
                        NpcList(TargetIndex).Char.charindex, _
                        Hechizos(h).ParticleViaje, Hechizos(h).Particle, _
                        Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
            Else
                Call SendData( _
                    SendTarget.ToNPCAliveArea, TargetIndex, _
                    PrepareMessageParticleFX( _
                        NpcList(TargetIndex).Char.charindex, _
                        Hechizos(h).Particle, Hechizos(h).TimeParticula, False))
            End If
        End If
    End If

    ' Sonido (sin viaje)
    If Hechizos(h).ParticleViaje = 0 Then
        Call SendData( _
            SendTarget.ToNPCAliveArea, TargetIndex, _
            PrepareMessagePlayWave(Hechizos(h).wav, _
                NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y))
    End If

    Exit Sub
CastOnNpc_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.CastOnNpc", Erl)
End Sub

' --- Helper 3: Target Terreno ---
Private Sub CastOnTerrain(ByVal UserIndex As Integer, ByVal h As Integer, ByVal skin As Integer)
    On Error GoTo CastOnTerrain_Err

    Dim fxId As Integer

    With UserList(UserIndex)
        ' FX piso (sin duplicar)
        If Hechizos(h).FXgrh > 0 Then
            If skin > 0 Then
                fxId = skin
            Else
                fxId = Hechizos(h).FXgrh
            End If
            Call modSendData.SendToAreaByPos( _
                .pos.Map, .flags.TargetX, .flags.TargetY, _
                PrepareMessageFxPiso(fxId, .flags.TargetX, .flags.TargetY))
        End If

        ' Partícula a piso
        If Hechizos(h).Particle > 0 Then
            Call SendData( _
                SendTarget.ToPCAliveArea, UserIndex, _
                PrepareMessageParticleFXToFloor( _
                    .flags.TargetX, .flags.TargetY, _
                    Hechizos(h).Particle, Hechizos(h).TimeParticula))
        End If

        ' Sonido (jugador)
        If Hechizos(h).wav <> 0 Then
            Call SendData( _
                SendTarget.ToPCAliveArea, UserIndex, _
                PrepareMessagePlayWave(Hechizos(h).wav, .flags.TargetX, .flags.TargetY))
        End If

        ' Sonido a NPC (solo si hay NPC válido y no hay viaje)
        If Hechizos(h).ParticleViaje = 0 And IsValidNpcRef(.flags.TargetNPC) Then
            Dim n As Integer
            n = .flags.TargetNPC.ArrayIndex
            Call SendData( _
                SendTarget.ToNPCAliveArea, n, _
                PrepareMessagePlayWave(Hechizos(h).wav, _
                    NpcList(n).pos.x, NpcList(n).pos.y))
        End If
    End With

    Exit Sub
CastOnTerrain_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.CastOnTerrain", Erl)
End Sub




Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean, ByRef IsAlive As Boolean)
    On Error GoTo HechizoPropUsuario_Err
    Dim h         As Integer
    Dim Damage    As Integer
    Dim DamageStr As String
    Dim tempChr   As Integer
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    tempChr = UserList(UserIndex).flags.TargetUser.ArrayIndex
    'Hambre
    If Hechizos(h).SubeHam = 1 Then
        Call InfoHechizo(UserIndex)
        Damage = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + Damage
        If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1875, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1875=Le has restaurado ¬1 puntos de hambre a ¬2.
            Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1895, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1895=¬1 te ha restaurado ¬2 puntos de hambre.
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1896, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1896=Te has restaurado ¬1 puntos de hambre.
        End If
        Call WriteUpdateHungerAndThirst(tempChr)
        b = True
    ElseIf Hechizos(h).SubeHam = 2 Then
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        Else
            Exit Sub
        End If
        Call InfoHechizo(UserIndex)
        Damage = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - Damage
        If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1897, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1897=Le has quitado ¬1 puntos de hambre a ¬2.
            Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1898, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1898=¬1 te ha quitado ¬2 puntos de hambre.
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1899, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1899=Te has quitado ¬1 puntos de hambre.
        End If
        Call WriteUpdateHungerAndThirst(tempChr)
        b = True
        If UserList(tempChr).Stats.MinHam < 1 Then
            UserList(tempChr).Stats.MinHam = 0
        End If
    End If
    'Sed
    If Hechizos(h).SubeSed = 1 Then
        Call InfoHechizo(UserIndex)
        Damage = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + Damage
        If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1900, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1900=Le has restaurado ¬1 puntos de sed a ¬2.
            Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1901, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1901=¬1 te ha restaurado ¬2 puntos de sed.
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1902, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1902=Te has restaurado ¬1 puntos de sed.
        End If
        Call WriteUpdateHungerAndThirst(tempChr)
        b = True
    ElseIf Hechizos(h).SubeSed = 2 Then
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        End If
        Call InfoHechizo(UserIndex)
        Damage = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - Damage
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1903, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1903=Le has quitado ¬1 puntos de sed a ¬2.
            Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1904, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1904=¬1 te ha quitado ¬2 puntos de sed.
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1905, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1905=Te has quitado ¬1 puntos de sed.
        End If
        If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
        End If
        Call WriteUpdateHungerAndThirst(tempChr)
        b = True
    End If
    ' <-------- Agilidad ---------->
    If Hechizos(h).SubeAgilidad = 1 Then
        'Verifica que el usuario no este muerto
        If UserList(tempChr).flags.Muerto = 1 Then
            b = False
            Exit Sub
        End If
        If Not PeleaSegura(UserIndex, tempChr) Then
            Select Case Status(UserIndex)
                Case 1, 3, 5 'Ciudadano o armada
                    If Status(tempChr) <> e_Facciones.Ciudadano And Status(tempChr) <> e_Facciones.Armada And Status(tempChr) <> e_Facciones.consejo Then
                        If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                            ' Msg662=No puedes ayudar criminales.
                            Call WriteLocaleMsg(UserIndex, 662, e_FontTypeNames.FONTTYPE_INFO)
                            b = False
                            Exit Sub
                        ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then
                            If UserList(UserIndex).flags.Seguro = True Then
                                ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                Call WriteLocaleMsg(UserIndex, 663, e_FontTypeNames.FONTTYPE_INFO)
                                b = False
                                Exit Sub
                            Else
                                'Si tiene clan
                                If UserList(UserIndex).GuildIndex > 0 Then
                                    'Si el clan es de alineación ciudadana.
                                    If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                        'No lo dejo resucitarlo
                                        ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                        Call WriteLocaleMsg(UserIndex, 664, e_FontTypeNames.FONTTYPE_INFO)
                                        b = False
                                        Exit Sub
                                        'Si es de alineación neutral, lo dejo resucitar y lo vuelvo criminal
                                    ElseIf GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                                        Call VolverCriminal(UserIndex)
                                        Call RefreshCharStatus(UserIndex)
                                    End If
                                Else
                                    Call VolverCriminal(UserIndex)
                                    Call RefreshCharStatus(UserIndex)
                                End If
                            End If
                        End If
                    End If
                Case 2, 4 'Caos
                    If Status(tempChr) <> e_Facciones.Caos And Status(tempChr) <> e_Facciones.Criminal And Status(tempChr) <> e_Facciones.concilio Then
                        'Msg822= No podés ayudar ciudadanos.
                        Call WriteLocaleMsg(UserIndex, 822, e_FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
            End Select
        End If
        Call InfoHechizo(UserIndex)
        Damage = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) + Damage, UserList( _
                tempChr).Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)
        UserList(tempChr).flags.TomoPocion = True
        b = True
        Call WriteFYA(tempChr)
    ElseIf Hechizos(h).SubeAgilidad = 2 Then
        'Verifica que el usuario no este muerto
        If UserList(tempChr).flags.Muerto = 1 Then
            b = False
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        End If
        Call InfoHechizo(UserIndex)
        UserList(tempChr).flags.TomoPocion = True
        Damage = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        If UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) - Damage < MINATRIBUTOS Then
            UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) = MINATRIBUTOS
        Else
            UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) - Damage
        End If
        b = True
        Call WriteFYA(tempChr)
    End If
    ' <-------- Fuerza ---------->
    If Hechizos(h).SubeFuerza = 1 Then
        'Verifica que el usuario no este muerto
        If UserList(tempChr).flags.Muerto = 1 Then
            b = False
            Exit Sub
        End If
        If Not PeleaSegura(UserIndex, tempChr) Then
            Select Case Status(UserIndex)
                Case 1, 3, 5 'Ciudadano o armada
                    If Status(tempChr) <> e_Facciones.Ciudadano And Status(tempChr) <> e_Facciones.Armada And Status(tempChr) <> e_Facciones.consejo Then
                        If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                            ' Msg662=No puedes ayudar criminales.
                            Call WriteLocaleMsg(UserIndex, 662, e_FontTypeNames.FONTTYPE_INFO)
                            b = False
                            Exit Sub
                        ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then
                            If UserList(UserIndex).flags.Seguro = True Then
                                ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                Call WriteLocaleMsg(UserIndex, 663, e_FontTypeNames.FONTTYPE_INFO)
                                b = False
                                Exit Sub
                            Else
                                'Si tiene clan
                                If UserList(UserIndex).GuildIndex > 0 Then
                                    'Si el clan es de alineación ciudadana.
                                    If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                        'No lo dejo resucitarlo
                                        ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                        Call WriteLocaleMsg(UserIndex, 664, e_FontTypeNames.FONTTYPE_INFO)
                                        b = False
                                        Exit Sub
                                        'Si es de alineación neutral, lo dejo resucitar y lo vuelvo criminal
                                    ElseIf GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                                        Call VolverCriminal(UserIndex)
                                        Call RefreshCharStatus(UserIndex)
                                    End If
                                Else
                                    Call VolverCriminal(UserIndex)
                                    Call RefreshCharStatus(UserIndex)
                                End If
                            End If
                        End If
                    End If
                Case 2, 4 'Caos
                    If Status(tempChr) <> e_Facciones.Caos And Status(tempChr) <> e_Facciones.Criminal And Status(tempChr) <> e_Facciones.concilio Then
                        ' Msg665=No podés ayudar ciudadanos.
                        Call WriteLocaleMsg(UserIndex, 665, e_FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
            End Select
        End If
        Damage = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) + Damage, UserList( _
                tempChr).Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)
        UserList(tempChr).flags.TomoPocion = True
        Call WriteFYA(tempChr)
        b = True
        Call InfoHechizo(UserIndex)
        Call WriteFYA(tempChr)
    ElseIf Hechizos(h).SubeFuerza = 2 Then
        'Verifica que el usuario no este muerto
        If UserList(tempChr).flags.Muerto = 1 Then
            b = False
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        End If
        UserList(tempChr).flags.TomoPocion = True
        Damage = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        If UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) - Damage < MINATRIBUTOS Then
            UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) = MINATRIBUTOS
        Else
            UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) - Damage
        End If
        b = True
        Call InfoHechizo(UserIndex)
        Call WriteFYA(tempChr)
    End If
    'Salud
    If IsSet(Hechizos(h).Effects, e_SpellEffects.eDoHeal) Then
        'Verifica que el usuario no este muerto
        If UserList(tempChr).flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        If UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp Then
            Call WriteLocaleMsg(UserIndex, 1906, e_FontTypeNames.FONTTYPE_INFOIAO, UserList(tempChr).name) ' Msg1906=¬1 no tiene heridas para curar.
            b = False
            Exit Sub
        End If
        'Para poder tirar curar a un pk en el ring
        If Not PeleaSegura(UserIndex, tempChr) Then
            If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
            Dim trigger As e_Trigger6
            trigger = TriggerZonaPelea(UserIndex, tempChr)
            ' Están en zona segura en un ring e intenta curarse desde afuera hacia adentro o viceversa
        ElseIf trigger = TRIGGER6_PROHIBE And MapInfo(UserList(UserIndex).pos.Map).Seguro <> 0 Then
            b = False
            Exit Sub
        End If
        Damage = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
        Damage = Damage * UserMod.GetMagicHealingBonus(UserList(UserIndex))
        Damage = Damage * UserMod.GetSelfHealingBonus(UserList(tempChr))
        If IsFeatureEnabled("healers_and_tanks") And UserList(UserIndex).flags.DivineBlood > 0 Then
            Damage = Damage * DivineBloodHealingMultiplierBonus
        End If
        Call InfoHechizo(UserIndex)
        Call UserMod.DoDamageOrHeal(tempChr, UserIndex, eUser, Damage, e_DamageSourceType.e_magic, h)
        DamageStr = PonerPuntos(Damage)
        If UserIndex <> tempChr Then
            Call WriteLocaleMsg(UserIndex, 388, e_FontTypeNames.FONTTYPE_FIGHT, UserList(tempChr).name & "¬" & DamageStr)
            Call WriteLocaleMsg(tempChr, 32, e_FontTypeNames.FONTTYPE_FIGHT, UserList(UserIndex).name & "¬" & DamageStr)
        Else
            Call WriteLocaleMsg(UserIndex, 33, e_FontTypeNames.FONTTYPE_FIGHT, DamageStr)
        End If
        b = True
    ElseIf IsSet(Hechizos(h).Effects, e_SpellEffects.eDoDamage) Then
        If UserIndex = tempChr Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
        Damage = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
        Damage = Damage + Porcentaje(Damage, 3 * UserList(UserIndex).Stats.ELV)
        ' Si al hechizo le afecta el daño mágico
        Dim PorcentajeRM As Integer
        If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
            PorcentajeRM = PorcentajeRM - ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicPenetration
            Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicAbsoluteBonus
        End If
        If UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicDamageBonus)
            Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicAbsoluteBonus
            PorcentajeRM = PorcentajeRM - ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicPenetration
        End If
        If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
            Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicAbsoluteBonus
            PorcentajeRM = PorcentajeRM - ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicPenetration
        End If
        ' Si el hechizo no ignora la RM
        If Hechizos(h).AntiRm = 0 Then
            PorcentajeRM = max(0, PorcentajeRM + GetUserMR(tempChr))
            ' Resto el porcentaje total
            Damage = Damage - Porcentaje(Damage, PorcentajeRM)
        End If
        Call EffectsOverTime.TargetWillAttack(UserList(UserIndex).EffectOverTime, tempChr, eUser, e_DamageSourceType.e_magic)
        Damage = Damage * UserMod.GetMagicDamageModifier(UserList(UserIndex))
        Damage = Damage * UserMod.GetMagicDamageReduction(UserList(tempChr))
        ' Prevengo daño negativo
        If Damage < 0 Then Damage = 0
        If UserIndex <> tempChr Then
            Call checkHechizosEfectividad(UserIndex, tempChr)
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        End If
        Call InfoHechizo(UserIndex)
        IsAlive = UserMod.DoDamageOrHeal(tempChr, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, h) = eStillAlive
        Call EffectsOverTime.TargetDidHit(UserList(UserIndex).EffectOverTime, tempChr, eUser, e_DamageSourceType.e_magic)
        Call SubirSkill(tempChr, Resistencia)
        b = True
    End If
    'Mana
    If Hechizos(h).SubeMana = 1 Then
        Call InfoHechizo(UserIndex)
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + Damage
        If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1907, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1907=Le has restaurado ¬1 puntos de mana a ¬2.
            Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1908, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1908=¬1 te ha restaurado ¬2 puntos de mana.
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1909, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1909=Te has restaurado ¬1 puntos de mana.
        End If
        Call WriteUpdateMana(tempChr)
        b = True
    ElseIf Hechizos(h).SubeMana = 2 Then
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        End If
        Call InfoHechizo(UserIndex)
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1910, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1910=Le has quitado ¬1 puntos de mana a ¬2.
            Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1911, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1911=¬1 te ha quitado ¬2 puntos de mana.
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1912, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1912=Te has quitado ¬1 puntos de mana.
        End If
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - Damage
        If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
        Call WriteUpdateMana(tempChr)
        b = True
    End If
    'Stamina
    If Hechizos(h).SubeSta = 1 Then
        Call InfoHechizo(UserIndex)
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + Damage
        If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1913, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1913=Le has restaurado ¬1 puntos de vitalidad a ¬2.
            Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1914, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1914=¬1 te ha restaurado ¬2 puntos de vitalidad.
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1915, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1915=Te has restaurado ¬1 puntos de vitalidad.
        End If
        Call WriteUpdateSta(tempChr)
        b = True
    ElseIf Hechizos(h).SubeSta = 2 Then
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        End If
        Call InfoHechizo(UserIndex)
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1916, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1916=Le has quitado ¬1 puntos de vitalidad a ¬2.
            Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1917, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1917=¬1 te ha quitado ¬2 puntos de vitalidad.
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1915, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1915=Te has restaurado ¬1 puntos de vitalidad.
        End If
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - Damage
        If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
        Call WriteUpdateSta(tempChr)
        b = True
    End If
    Exit Sub
HechizoPropUsuario_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoPropUsuario", Erl)
End Sub

Sub HechizoCombinados(ByVal UserIndex As Integer, ByRef b As Boolean, ByRef IsAlive As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 02/01/2008
    '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
    '***************************************************
    On Error GoTo HechizoCombinados_Err
    Dim h                 As Integer
    Dim Damage            As Integer
    Dim targetUserIndex   As Integer
    Dim enviarInfoHechizo As Boolean
    enviarInfoHechizo = False
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    targetUserIndex = UserList(UserIndex).flags.TargetUser.ArrayIndex
    ' <-------- Agilidad ---------->
    If Hechizos(h).SubeAgilidad = 1 Then
        'Para poder tirar cl a un pk en el ring
        If Not PeleaSegura(UserIndex, targetUserIndex) Then
            If Status(targetUserIndex) = 0 And Status(UserIndex) = 1 Or Status(targetUserIndex) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                End If
            End If
        End If
        enviarInfoHechizo = True
        Damage = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(targetUserIndex).flags.DuracionEfecto = Hechizos(h).Duration
        UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) + Damage, UserList( _
                targetUserIndex).Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)
        UserList(targetUserIndex).flags.TomoPocion = True
        b = True
        Call WriteFYA(targetUserIndex)
    ElseIf Hechizos(h).SubeAgilidad = 2 Then
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        If UserIndex <> targetUserIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        End If
        enviarInfoHechizo = True
        UserList(targetUserIndex).flags.TomoPocion = True
        Damage = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(targetUserIndex).flags.DuracionEfecto = Hechizos(h).Duration
        If UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) - Damage < 6 Then
            UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) = MINATRIBUTOS
        Else
            UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) = UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) - Damage
        End If
        b = True
        Call WriteFYA(targetUserIndex)
    End If
    ' <-------- Fuerza ---------->
    If Hechizos(h).SubeFuerza = 1 Then
        'Para poder tirar fuerza a un pk en el ring
        If Not PeleaSegura(UserIndex, targetUserIndex) Then
            If Status(targetUserIndex) = 0 And Status(UserIndex) = 1 Or Status(targetUserIndex) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        End If
        Damage = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(targetUserIndex).flags.DuracionEfecto = Hechizos(h).Duration
        UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) + Damage, UserList( _
                targetUserIndex).Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)
        UserList(targetUserIndex).flags.TomoPocion = True
        b = True
        enviarInfoHechizo = True
        Call WriteFYA(targetUserIndex)
    ElseIf Hechizos(h).SubeFuerza = 2 Then
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        If UserIndex <> targetUserIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        End If
        UserList(targetUserIndex).flags.TomoPocion = True
        Damage = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(targetUserIndex).flags.DuracionEfecto = Hechizos(h).Duration
        If UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) - Damage < 6 Then
            UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) = MINATRIBUTOS
        Else
            UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) = UserList(targetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) - Damage
        End If
        b = True
        enviarInfoHechizo = True
        Call WriteFYA(targetUserIndex)
    End If
    'Salud
    If IsSet(Hechizos(h).Effects, e_SpellEffects.eDoHeal) Then
        'Verifica que el usuario no este muerto
        If UserList(targetUserIndex).flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        'Para poder tirar curar a un pk en el ring
        If Not PeleaSegura(UserIndex, targetUserIndex) Then
            If Status(targetUserIndex) = 0 And Status(UserIndex) = 1 Or Status(targetUserIndex) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        End If
        Damage = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
        Damage = Damage * UserMod.GetMagicHealingBonus(UserList(UserIndex))
        Damage = Damage * UserMod.GetSelfHealingBonus(UserList(targetUserIndex))
        If UserList(UserIndex).flags.DivineBlood > 0 And IsFeatureEnabled("healers_and_tanks") Then
            Damage = Damage * DivineBloodHealingMultiplierBonus
        End If
        enviarInfoHechizo = True
        Call UserMod.DoDamageOrHeal(targetUserIndex, UserIndex, eUser, Damage, e_DamageSourceType.e_magic, h)
        If UserIndex <> targetUserIndex Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1918, Damage & "¬" & UserList(targetUserIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1918=Le has restaurado ¬1 puntos de vida a ¬2.
            Call WriteConsoleMsg(targetUserIndex, PrepareMessageLocaleMsg(1919, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1919=¬1 te ha restaurado ¬2 puntos de vida.
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1920, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1920=Te has restaurado ¬1 puntos de vida.
        End If
        b = True
    ElseIf IsSet(Hechizos(h).Effects, e_SpellEffects.eDoDamage) Then ' Damage
        If UserIndex = targetUserIndex Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, targetUserIndex) Then Exit Sub
        Damage = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
        Damage = Damage + Porcentaje(Damage, 3 * UserList(UserIndex).Stats.ELV)
        ' mage has 30% damage reduction
        If UserList(UserIndex).clase = e_Class.Mage Then
            Damage = Damage * 0.7
        End If
        Dim MR As Integer
        ' Weapon Magic bonus
        If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
            Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicAbsoluteBonus
            MR = MR - ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicPenetration
        End If
        ' Magic ring bonus
        If UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicDamageBonus)
            Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicAbsoluteBonus
            MR = MR - ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicPenetration
        End If
        If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
            Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
            Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicAbsoluteBonus
            MR = MR - ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicPenetration
        End If
        ' Si el hechizo no ignora la RM
        If Hechizos(h).AntiRm = 0 Then
            ' Resistencia mágica armadura
            MR = max(0, MR + GetUserMR(targetUserIndex))
            If MR > 0 Then
                Damage = Damage - Porcentaje(Damage, MR)
            End If
        End If
        Call EffectsOverTime.TargetWillAttack(UserList(UserIndex).EffectOverTime, targetUserIndex, eUser, e_DamageSourceType.e_magic)
        Damage = Damage * UserMod.GetMagicDamageModifier(UserList(UserIndex))
        Damage = Damage * UserMod.GetMagicDamageReduction(UserList(targetUserIndex))
        ' Prevengo daño negativo
        If Damage < 0 Then Damage = 0
        If UserIndex <> targetUserIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, targetUserIndex)
        End If
        enviarInfoHechizo = True
        IsAlive = UserMod.DoDamageOrHeal(targetUserIndex, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, h) = eStillAlive
        Call EffectsOverTime.TargetDidHit(UserList(UserIndex).EffectOverTime, targetUserIndex, eUser, e_DamageSourceType.e_magic)
        Call SubirSkill(targetUserIndex, Resistencia)
        b = True
    End If
    Dim tU As Integer
    tU = targetUserIndex
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Invisibility) Then
        If UserList(tU).flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        If UserList(tU).Counters.Saliendo Then
            If UserIndex <> tU Then
                ' Msg666=¡El hechizo no tiene efecto!
                Call WriteLocaleMsg(UserIndex, MSG_NPC_INMUNE_TO_SPELLS, e_FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                ' Msg667=¡No podés ponerte invisible mientras te encuentres saliendo!
                Call WriteLocaleMsg(UserIndex, 667, e_FontTypeNames.FONTTYPE_WARNING)
                b = False
                Exit Sub
            End If
        End If
        If IsSet(UserList(tU).flags.StatusMask, eTaunting) Then
            ' Msg666=¡El hechizo no tiene efecto!
            Call WriteLocaleMsg(UserIndex, MSG_NPC_INMUNE_TO_SPELLS, e_FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        If Not PeleaSegura(UserIndex, tU) Then
            Select Case Status(UserIndex)
                Case 1, 3, 5 'Ciudadano o armada
                    If Status(tU) <> e_Facciones.Ciudadano And Status(tU) <> e_Facciones.Armada And Status(tU) <> e_Facciones.consejo Then
                        If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                            ' Msg662=No puedes ayudar criminales.
                            Call WriteLocaleMsg(UserIndex, 662, e_FontTypeNames.FONTTYPE_INFO)
                            b = False
                            Exit Sub
                        ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then
                            If UserList(UserIndex).flags.Seguro = True Then
                                ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                Call WriteLocaleMsg(UserIndex, 663, e_FontTypeNames.FONTTYPE_INFO)
                                b = False
                                Exit Sub
                            Else
                                'Si tiene clan
                                If UserList(UserIndex).GuildIndex > 0 Then
                                    'Si el clan es de alineación ciudadana.
                                    If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                        'No lo dejo resucitarlo
                                        ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                        Call WriteLocaleMsg(UserIndex, 664, e_FontTypeNames.FONTTYPE_INFO)
                                        b = False
                                        Exit Sub
                                        'Si es de alineación neutral, lo dejo resucitar y lo vuelvo criminal
                                    ElseIf GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                                        Call VolverCriminal(UserIndex)
                                        Call RefreshCharStatus(UserIndex)
                                    End If
                                Else
                                    Call VolverCriminal(UserIndex)
                                    Call RefreshCharStatus(UserIndex)
                                End If
                            End If
                        End If
                    End If
                Case 2, 4 'Caos
                    If Status(tU) <> e_Facciones.Caos And Status(tU) <> e_Facciones.Criminal And Status(tU) <> e_Facciones.concilio Then
                        ' Msg668=No podés ayudar ciudadanos.
                        Call WriteLocaleMsg(UserIndex, 668, e_FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
            End Select
        End If
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And e_PlayerType.User Then
            If Not UserList(tU).flags.Privilegios And e_PlayerType.User Then
                Exit Sub
            End If
        End If
        UserList(tU).flags.invisible = 1
        'Ladder
        'Reseteamos el contador de Invisibilidad
        If UserList(tU).Counters.Invisibilidad <= 0 Then UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
        Call WriteContadores(tU)
        Call SendData(SendTarget.ToPCAliveArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.charindex, True, UserList(tU).pos.x, UserList(tU).pos.y))
        enviarInfoHechizo = True
        b = True
    End If
    If Hechizos(h).Envenena > 0 Then
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).flags.Envenenado = Hechizos(h).Envenena
        enviarInfoHechizo = True
        b = True
    End If
    If Hechizos(h).desencantar = 1 Then
        ' Msg669=Has sido desencantado.
        Call WriteLocaleMsg(UserIndex, 669, e_FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.Envenenado = 0
        UserList(UserIndex).flags.Incinerado = 0
        If UserList(UserIndex).flags.Inmovilizado = 1 Then
            UserList(UserIndex).Counters.Inmovilizado = 0
            If UserList(UserIndex).clase = e_Class.Warrior Or UserList(UserIndex).clase = e_Class.Hunter Or UserList(UserIndex).clase = e_Class.Thief Or UserList( _
                    UserIndex).clase = e_Class.Pirat Then
                UserList(UserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            UserList(UserIndex).flags.Inmovilizado = 0
            Call WriteInmovilizaOK(UserIndex)
        End If
        If UserList(UserIndex).flags.Paralizado = 1 Then
            UserList(UserIndex).Counters.Paralisis = 0
            If UserList(UserIndex).clase = e_Class.Warrior Or UserList(UserIndex).clase = e_Class.Hunter Or UserList(UserIndex).clase = e_Class.Thief Or UserList( _
                    UserIndex).clase = e_Class.Pirat Then
                UserList(UserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            UserList(UserIndex).flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
        End If
        If UserList(UserIndex).flags.Ceguera = 1 Then
            UserList(UserIndex).Counters.Ceguera = 0
            UserList(UserIndex).flags.Ceguera = 0
            Call WriteBlindNoMore(UserIndex)
        End If
        If UserList(UserIndex).flags.Maldicion = 1 Then
            UserList(UserIndex).flags.Maldicion = 0
            UserList(UserIndex).Counters.Maldicion = 0
        End If
        enviarInfoHechizo = True
        b = True
    End If
    If Hechizos(h).Sanacion = 1 Then
        UserList(tU).flags.Envenenado = 0
        UserList(tU).flags.Incinerado = 0
        If UserList(tU).Counters.velocidad <> 0 Then
            UserList(tU).flags.VelocidadHechizada = 0
            UserList(tU).Counters.velocidad = 0
            Call ActualizarVelocidadDeUsuario(tU)
        End If
        enviarInfoHechizo = True
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Incinerate) Then
        If UserIndex = tU Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).Counters.Incineracion = 1
        UserList(tU).flags.Incinerado = 1
        enviarInfoHechizo = True
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.CurePoison) Then
        'Verificamos que el usuario no este muerto
        If UserList(tU).flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        'Para poder tirar curar veneno a un pk en el ring
        If Not PeleaSegura(UserIndex, tU) Then
            If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        End If
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And e_PlayerType.User Then
            If Not UserList(tU).flags.Privilegios And e_PlayerType.User Then
                Exit Sub
            End If
        End If
        UserList(tU).flags.Envenenado = 0
        UserList(tU).Counters.Veneno = 0
        enviarInfoHechizo = True
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Curse) Then
        If UserIndex = tU Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).flags.Maldicion = 1
        UserList(tU).Counters.Maldicion = 200
        enviarInfoHechizo = True
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveCurse) Then
        UserList(tU).flags.Maldicion = 0
        UserList(tU).Counters.Maldicion = 0
        enviarInfoHechizo = True
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.PreciseHit) Then
        UserList(tU).flags.GolpeCertero = 1
        enviarInfoHechizo = True
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Paralize) Then
        If UserIndex = tU Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        enviarInfoHechizo = True
        b = True
        UserList(tU).Counters.Paralisis = Hechizos(h).Duration
        If UserList(tU).flags.Paralizado = 0 Then
            UserList(tU).flags.Paralizado = 1
            Call WriteParalizeOK(tU)
            Call WritePosUpdate(tU)
        End If
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Immobilize) Then
        If UserIndex = tU Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        enviarInfoHechizo = True
        b = True
        UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration
        If UserList(tU).flags.Inmovilizado = 0 Then
            UserList(tU).flags.Inmovilizado = 1
            Call WriteInmovilizaOK(tU)
            Call WritePosUpdate(tU)
        End If
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveParalysis) Then
        'Para poder tirar remo a un pk en el ring
        If Not PeleaSegura(UserIndex, tU) Then
            If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
                If UserList(UserIndex).flags.Seguro Then
                    Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)
                End If
            End If
        End If
        If UserList(tU).flags.Inmovilizado = 1 Then
            UserList(tU).Counters.Inmovilizado = 0
            UserList(tU).flags.Inmovilizado = 0
            If UserList(tU).clase = e_Class.Warrior Or UserList(tU).clase = e_Class.Hunter Or UserList(tU).clase = e_Class.Thief Or UserList(tU).clase = e_Class.Pirat Then
                UserList(tU).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            Call WriteInmovilizaOK(tU)
            enviarInfoHechizo = True
            b = True
        End If
        If UserList(tU).flags.Paralizado = 1 Then
            UserList(tU).Counters.Paralisis = 0
            UserList(tU).flags.Paralizado = 0
            If UserList(tU).clase = e_Class.Warrior Or UserList(tU).clase = e_Class.Hunter Or UserList(tU).clase = e_Class.Thief Or UserList(tU).clase = e_Class.Pirat Then
                UserList(tU).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            Call WriteParalizeOK(tU)
            enviarInfoHechizo = True
            b = True
        End If
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Blindness) Then
        If UserIndex = tU Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        UserList(tU).flags.Ceguera = 1
        UserList(tU).Counters.Ceguera = Hechizos(h).Duration
        Call WriteBlind(tU)
        enviarInfoHechizo = True
        b = True
    End If
    If IsSet(Hechizos(h).Effects, e_SpellEffects.Dumb) Then
        If UserIndex = tU Then
            Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)
        End If
        If UserList(tU).flags.Estupidez = 0 Then
            UserList(tU).flags.Estupidez = 1
            UserList(tU).Counters.Estupidez = Hechizos(h).Duration
        End If
        Call WriteDumb(tU)
        enviarInfoHechizo = True
        b = True
    End If
    If Hechizos(h).velocidad <> 0 Then
        If Hechizos(h).velocidad < 1 Then
            If UserIndex = tU Then
                Call WriteLocaleMsg(UserIndex, 380, e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        End If
        enviarInfoHechizo = True
        b = True
        If UserList(tU).Counters.velocidad = 0 Then
            UserList(tU).flags.VelocidadHechizada = Hechizos(h).velocidad
            Call ActualizarVelocidadDeUsuario(tU)
        End If
        UserList(tU).Counters.velocidad = Hechizos(h).Duration
    End If
    If enviarInfoHechizo Then
        Call InfoHechizo(UserIndex)
    End If
    Exit Sub
HechizoCombinados_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoCombinados", Erl)
End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo UpdateUserHechizos_Err
    Dim LoopC As Byte
    'Actualiza un solo slot
    If Not UpdateAll Then
        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
            Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
        Else
            Call ChangeUserHechizo(UserIndex, Slot, 0)
        End If
    Else
        'Actualiza todos los slots
        For LoopC = 1 To MAXUSERHECHIZOS
            'Actualiza el inventario
            If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
                Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
            Else
                Call ChangeUserHechizo(UserIndex, LoopC, 0)
            End If
        Next LoopC
    End If
    Exit Sub
UpdateUserHechizos_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.UpdateUserHechizos", Erl)
End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)
    On Error GoTo ChangeUserHechizo_Err
    UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
    Call WriteChangeSpellSlot(UserIndex, Slot)
    Exit Sub
ChangeUserHechizo_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.ChangeUserHechizo", Erl)
End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)
    On Error GoTo DesplazarHechizo_Err
    If (Dire <> 1 And Dire <> -1) Then Exit Sub
    If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub
    Dim TempHechizo   As Integer
    Dim SpellInterval As Long
    With UserList(UserIndex)
        If Dire = 1 Then 'Mover arriba
            If CualHechizo = 1 Then
                ' Msg670=No podés mover el hechizo en esa direccion.
                Call WriteLocaleMsg(UserIndex, 670, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                TempHechizo = .Stats.UserHechizos(CualHechizo)
                .Stats.UserHechizos(CualHechizo) = .Stats.UserHechizos(CualHechizo - 1)
                .Stats.UserHechizos(CualHechizo - 1) = TempHechizo
                SpellInterval = .Counters.UserHechizosInterval(CualHechizo)
                .Counters.UserHechizosInterval(CualHechizo) = .Counters.UserHechizosInterval(CualHechizo - 1)
                .Counters.UserHechizosInterval(CualHechizo - 1) = SpellInterval
                'Prevent the user from casting other spells than the one he had selected when he hitted cast.
                If .flags.Hechizo = CualHechizo Then
                    .flags.Hechizo = .flags.Hechizo - 1
                ElseIf .flags.Hechizo = CualHechizo - 1 Then
                    .flags.Hechizo = .flags.Hechizo + 1
                End If
                .flags.ModificoHechizos = True
            End If
        Else 'mover abajo
            If CualHechizo = MAXUSERHECHIZOS Then
                ' Msg670=No podés mover el hechizo en esa direccion.
                Call WriteLocaleMsg(UserIndex, 670, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                TempHechizo = .Stats.UserHechizos(CualHechizo)
                .Stats.UserHechizos(CualHechizo) = .Stats.UserHechizos(CualHechizo + 1)
                .Stats.UserHechizos(CualHechizo + 1) = TempHechizo
                SpellInterval = .Counters.UserHechizosInterval(CualHechizo)
                .Counters.UserHechizosInterval(CualHechizo) = .Counters.UserHechizosInterval(CualHechizo + 1)
                .Counters.UserHechizosInterval(CualHechizo + 1) = SpellInterval
                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
                If .flags.Hechizo = CualHechizo Then
                    .flags.Hechizo = .flags.Hechizo + 1
                ElseIf .flags.Hechizo = CualHechizo + 1 Then
                    .flags.Hechizo = .flags.Hechizo - 1
                End If
                .flags.ModificoHechizos = True
            End If
        End If
    End With
    Exit Sub
DesplazarHechizo_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.DesplazarHechizo", Erl)
End Sub

Private Sub AreaHechizo(UserIndex As Integer, NpcIndex As Integer, x As Byte, y As Byte, Npc As Boolean)
    On Error GoTo AreaHechizo_Err
    Dim calculo        As Integer
    Dim TilesDifUser   As Integer
    Dim TilesDifNpc    As Integer
    Dim tilDif         As Integer
    Dim h2             As Integer
    Dim Hit            As Integer
    Dim Damage         As Integer
    Dim porcentajeDesc As Integer
    h2 = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    'Calculo de descuesto de golpe por cercania.
    TilesDifUser = x + y
    If Npc Then
        If IsSet(Hechizos(h2).Effects, e_SpellEffects.eDoDamage) Then
            TilesDifNpc = NpcList(NpcIndex).pos.x + NpcList(NpcIndex).pos.y
            tilDif = TilesDifUser - TilesDifNpc
            Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)
            Hit = Hit + Porcentaje(Hit, 3 * UserList(UserIndex).Stats.ELV)
            ' Daño mágico arma
            If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
                Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
            End If
            ' Daño mágico anillo
            If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
                Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
            End If
            ' Disminuir daño con distancia
            If tilDif <> 0 Then
                porcentajeDesc = Abs(tilDif) * 20
                Damage = Hit / 100 * porcentajeDesc
                Damage = Hit - Damage
            Else
                Damage = Hit
            End If
            ' Si el hechizo no ignora la RM
            If Hechizos(h2).AntiRm = 0 Then
                Damage = Damage - NpcList(NpcIndex).Stats.defM
            End If
            ' Prevengo daño negativo
            If Damage < 0 Then Damage = 0
            Call NPCs.DoDamageOrHeal(NpcIndex, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, h2)
            If UserList(UserIndex).ChatCombate = 1 Then
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1921, Damage & "¬" & NpcList(NpcIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1921=Le has causado ¬1 puntos de daño a ¬2.
            End If
        End If
        Exit Sub
    Else
        TilesDifNpc = UserList(NpcIndex).pos.x + UserList(NpcIndex).pos.y
        tilDif = TilesDifUser - TilesDifNpc
        If IsSet(Hechizos(h2).Effects, e_SpellEffects.eDoDamage) Then
            If UserIndex = NpcIndex Then
                Exit Sub
            End If
            If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            If UserIndex <> NpcIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
            End If
            Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)
            Hit = Hit + Porcentaje(Hit, 3 * UserList(UserIndex).Stats.ELV)
            ' Daño mágico arma
            If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
                Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
            End If
            ' Daño mágico anillo
            If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
                Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
            End If
            If tilDif <> 0 Then
                porcentajeDesc = Abs(tilDif) * 20
                Damage = Hit / 100 * porcentajeDesc
                Damage = Hit - Damage
            Else
                Damage = Hit
            End If
            ' Si el hechizo no ignora la RM
            If Hechizos(h2).AntiRm = 0 Then
                ' Resistencia mágica armadura
                If UserList(NpcIndex).invent.EquippedArmorObjIndex > 0 Then
                    Damage = Damage - Porcentaje(Damage, ObjData(UserList(NpcIndex).invent.EquippedArmorObjIndex).ResistenciaMagica)
                End If
                ' Resistencia mágica anillo
                If UserList(NpcIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
                    Damage = Damage - Porcentaje(Damage, ObjData(UserList(NpcIndex).invent.EquippedRingAccesoryObjIndex).ResistenciaMagica)
                End If
                ' Resistencia mágica escudo
                If UserList(NpcIndex).invent.EquippedShieldObjIndex > 0 Then
                    Damage = Damage - Porcentaje(Damage, ObjData(UserList(NpcIndex).invent.EquippedShieldObjIndex).ResistenciaMagica)
                End If
                ' Resistencia mágica casco
                If UserList(NpcIndex).invent.EquippedHelmetObjIndex > 0 Then
                    Damage = Damage - Porcentaje(Damage, ObjData(UserList(NpcIndex).invent.EquippedHelmetObjIndex).ResistenciaMagica)
                End If
                ' Resistencia mágica de la clase
                Damage = Damage - Damage * ModClase(UserList(NpcIndex).clase).ResistenciaMagica
            End If
            ' Prevengo daño negativo
            If Damage < 0 Then Damage = 0
            Call UserMod.DoDamageOrHeal(NpcIndex, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, h2)
            Call SubirSkill(NpcIndex, Resistencia)
            Call WriteUpdateUserStats(NpcIndex)
        End If
        If IsSet(Hechizos(h2).Effects, e_SpellEffects.eDoHeal) Then
            If Not PeleaSegura(UserIndex, NpcIndex) Then
                If Status(UserIndex) = 1 And Status(NpcIndex) <> 1 Then
                    Exit Sub
                End If
            End If
            Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)
            If tilDif <> 0 Then
                porcentajeDesc = Abs(tilDif) * 20
                Damage = Hit / 100 * porcentajeDesc
                Damage = Hit - Damage
            Else
                Damage = Hit
            End If
            Damage = Damage * UserMod.GetMagicHealingBonus(UserList(UserIndex))
            Damage = Damage * NPCs.GetSelfHealingBonus(NpcList(NpcIndex))
            Call UserMod.DoDamageOrHeal(NpcIndex, UserIndex, eUser, Damage, e_DamageSourceType.e_magic, h2)
            If UserIndex <> NpcIndex Then
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1922, Damage & "¬" & UserList(NpcIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1922=Le has restaurado ¬1 puntos de vida a ¬2.
                Call WriteConsoleMsg(NpcIndex, PrepareMessageLocaleMsg(1923, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1923=¬1 te ha restaurado ¬2 puntos de vida.
            Else
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1920, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1920=Te has restaurado ¬1 puntos de vida.
            End If
        End If
        Call WriteUpdateUserStats(NpcIndex)
    End If
    If Hechizos(h2).Envenena > 0 Then
        If UserIndex = NpcIndex Then
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
        End If
        UserList(NpcIndex).flags.Envenenado = Hechizos(h2).Envenena
        Call WriteConsoleMsg(NpcIndex, PrepareMessageLocaleMsg(1924, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1924=¬1 te ha envenenado.
    End If
    If IsSet(Hechizos(h2).Effects, e_SpellEffects.Paralize) Then
        If UserIndex = NpcIndex Then
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
        End If
        'Msg823= Has sido paralizado.
        Call WriteLocaleMsg(NpcIndex, "823", e_FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).Counters.Paralisis = Hechizos(h2).Duration
        If UserList(NpcIndex).flags.Paralizado = 0 Then
            UserList(NpcIndex).flags.Paralizado = 1
            Call WriteParalizeOK(NpcIndex)
        End If
    End If
    If IsSet(Hechizos(h2).Effects, e_SpellEffects.Immobilize) Then
        If UserIndex = NpcIndex Then
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
        End If
        'Msg824= Has sido inmovilizado.
        Call WriteLocaleMsg(NpcIndex, "824", e_FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).Counters.Inmovilizado = Hechizos(h2).Duration
        If UserList(NpcIndex).flags.Inmovilizado = 0 Then
            UserList(NpcIndex).flags.Inmovilizado = 1
            Call WriteInmovilizaOK(NpcIndex)
            Call WritePosUpdate(NpcIndex)
        End If
    End If
    If IsSet(Hechizos(h2).Effects, e_SpellEffects.Blindness) Then
        If UserIndex = NpcIndex Then
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
        End If
        UserList(NpcIndex).flags.Ceguera = 1
        UserList(NpcIndex).Counters.Ceguera = Hechizos(h2).Duration
        'Msg825= Te han cegado.
        Call WriteLocaleMsg(NpcIndex, "825", e_FontTypeNames.FONTTYPE_INFO)
        Call WriteBlind(NpcIndex)
    End If
    If Hechizos(h2).velocidad > 0 Then
        If UserIndex = NpcIndex Then
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
        End If
        If UserList(NpcIndex).Counters.velocidad = 0 Then
            UserList(NpcIndex).flags.VelocidadHechizada = Hechizos(h2).velocidad
            Call ActualizarVelocidadDeUsuario(NpcIndex)
        End If
        UserList(NpcIndex).Counters.velocidad = Hechizos(h2).Duration
    End If
    If IsSet(Hechizos(h2).Effects, e_SpellEffects.Curse) Then
        If UserIndex = NpcIndex Then
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
        End If
        'Msg826= Ahora estas maldito. No podras Atacar
        Call WriteLocaleMsg(NpcIndex, "826", e_FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.Maldicion = 1
        UserList(NpcIndex).Counters.Maldicion = Hechizos(h2).Duration
    End If
    If IsSet(Hechizos(h2).Effects, e_SpellEffects.RemoveCurse) Then
        'Msg827= Te han removido la maldicion.
        Call WriteLocaleMsg(NpcIndex, "827", e_FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.Maldicion = 0
    End If
    If IsSet(Hechizos(h2).Effects, e_SpellEffects.PreciseHit) Then
        'Msg828= Tu proximo golpe sera certero.
        Call WriteLocaleMsg(NpcIndex, "828", e_FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.GolpeCertero = 1
    End If
    If IsSet(Hechizos(h2).Effects, e_SpellEffects.Incinerate) Then
        If UserIndex = NpcIndex Then
            Exit Sub
        End If
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
        End If
        UserList(NpcIndex).flags.Incinerado = 1
        'Msg829= Has sido Incinerado.
        Call WriteLocaleMsg(NpcIndex, "829", e_FontTypeNames.FONTTYPE_INFO)
    End If
    If IsSet(Hechizos(h2).Effects, e_SpellEffects.Invisibility) Then
        'Msg830= Ahora sos invisible.
        Call WriteLocaleMsg(NpcIndex, "830", e_FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.invisible = 1
        UserList(NpcIndex).Counters.Invisibilidad = Hechizos(h2).Duration
        Call WriteContadores(NpcIndex)
        Call SendData(SendTarget.ToPCAliveArea, NpcIndex, PrepareMessageSetInvisible(UserList(NpcIndex).Char.charindex, True, UserList(NpcIndex).pos.x, UserList(NpcIndex).pos.y))
    End If
    If Hechizos(h2).Sanacion = 1 Then
        'Msg831= Has sido sanado.
        Call WriteLocaleMsg(NpcIndex, "831", e_FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.Envenenado = 0
        UserList(NpcIndex).flags.Incinerado = 0
        If UserList(NpcIndex).Counters.velocidad <> 0 Then
            UserList(NpcIndex).flags.VelocidadHechizada = 0
            UserList(NpcIndex).Counters.velocidad = 0
            Call ActualizarVelocidadDeUsuario(NpcIndex)
        End If
    End If
    If IsSet(Hechizos(h2).Effects, e_SpellEffects.RemoveParalysis) Then
        'Msg832= Has sido removido.
        Call WriteLocaleMsg(NpcIndex, "832", e_FontTypeNames.FONTTYPE_INFO)
        If UserList(NpcIndex).flags.Inmovilizado = 1 Then
            UserList(NpcIndex).Counters.Inmovilizado = 0
            UserList(NpcIndex).flags.Inmovilizado = 0
            If UserList(NpcIndex).clase = e_Class.Warrior Or UserList(NpcIndex).clase = e_Class.Hunter Or UserList(NpcIndex).clase = e_Class.Thief Or UserList(NpcIndex).clase = _
                    e_Class.Pirat Then
                UserList(NpcIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            Call WriteInmovilizaOK(NpcIndex)
        End If
        If UserList(NpcIndex).flags.Paralizado = 1 Then
            UserList(NpcIndex).flags.Paralizado = 0
            If UserList(NpcIndex).clase = e_Class.Warrior Or UserList(NpcIndex).clase = e_Class.Hunter Or UserList(NpcIndex).clase = e_Class.Thief Or UserList(NpcIndex).clase = _
                    e_Class.Pirat Then
                UserList(NpcIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            'no need to crypt this
            Call WriteParalizeOK(NpcIndex)
        End If
    End If
    If Hechizos(h2).desencantar = 1 Then
        'Msg833= Has sido desencantado.
        Call WriteLocaleMsg(NpcIndex, "833", e_FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.Envenenado = 0
        UserList(NpcIndex).Counters.Veneno = 0
        UserList(NpcIndex).flags.Incinerado = 0
        UserList(NpcIndex).Counters.Incineracion = 0
        If UserList(NpcIndex).flags.Inmovilizado = 1 Then
            UserList(NpcIndex).Counters.Inmovilizado = 0
            UserList(NpcIndex).flags.Inmovilizado = 0
            If UserList(NpcIndex).clase = e_Class.Warrior Or UserList(NpcIndex).clase = e_Class.Hunter Or UserList(NpcIndex).clase = e_Class.Thief Or UserList(NpcIndex).clase = _
                    e_Class.Pirat Then
                UserList(NpcIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            Call WriteInmovilizaOK(NpcIndex)
        End If
        If UserList(NpcIndex).flags.Paralizado = 1 Then
            UserList(NpcIndex).flags.Paralizado = 0
            If UserList(NpcIndex).clase = e_Class.Warrior Or UserList(NpcIndex).clase = e_Class.Hunter Or UserList(NpcIndex).clase = e_Class.Thief Or UserList(NpcIndex).clase = _
                    e_Class.Pirat Then
                UserList(NpcIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
            End If
            UserList(NpcIndex).Counters.Paralisis = 0
            Call WriteParalizeOK(NpcIndex)
        End If
        If UserList(NpcIndex).flags.Ceguera = 1 Then
            UserList(NpcIndex).Counters.Ceguera = 0
            UserList(NpcIndex).flags.Ceguera = 0
            Call WriteBlindNoMore(NpcIndex)
        End If
        If UserList(NpcIndex).flags.Maldicion = 1 Then
            UserList(NpcIndex).flags.Maldicion = 0
            UserList(NpcIndex).Counters.Maldicion = 0
        End If
    End If
    Exit Sub
AreaHechizo_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.AreaHechizo", Erl)
End Sub

Public Sub AdjustNpcStatWithCasterLevel(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    Dim BaseHit       As Integer
    Dim BonusDamage   As Single
    Dim BonusFromItem As Integer
    BaseHit = UserList(UserIndex).Stats.ELV
    If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
        BonusFromItem = BonusFromItem + ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus
        If ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MaderaElfica > 0 Then
            BonusFromItem = BonusFromItem * 2
        End If
    End If
    If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex Then
        BonusFromItem = BonusFromItem + ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus
        If ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MaderaElfica > 0 Then
            BonusFromItem = BonusFromItem * 2
        End If
    End If
    BonusDamage = BonusFromItem / 100
    With NpcList(NpcIndex)
        .PoderAtaque = .PoderAtaque + BaseHit
        .Stats.MinHIT = .Stats.MinHIT + (.Stats.MinHIT * BonusDamage)
        .Stats.MaxHit = .Stats.MaxHit + (.Stats.MaxHit * BonusDamage)
    End With
End Sub

Public Sub UseSpellSlot(ByVal UserIndex As Integer, ByVal spellSlot As Integer)
    On Error GoTo UseSpellSlot_Err
    With UserList(UserIndex)
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        .flags.Hechizo = spellSlot
        If UserMod.IsStun(.flags, .Counters) Then
            Call WriteLocaleMsg(UserIndex, 394, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Hechizo < 1 Or .flags.Hechizo > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
        End If
        If .flags.Hechizo <> 0 Then
            If (.flags.Privilegios And e_PlayerType.Consejero) = 0 Then
                If .Stats.UserHechizos(spellSlot) <> 0 Then
                    If Hechizos(.Stats.UserHechizos(spellSlot)).AutoLanzar = 1 Then
                        If .flags.Descansar Then Exit Sub
                        If .flags.Meditando Then
                            .flags.Meditando = False
                            .Char.FX = 0
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
                        End If
                        'If exiting, cancel
                        Call CancelExit(UserIndex)
                        Call SetUserRef(UserList(UserIndex).flags.TargetUser, UserIndex)
                        Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    Else
                        If Hechizos(.Stats.UserHechizos(spellSlot)).AreaAfecta > 0 Then
                            Call WriteWorkRequestTarget(UserIndex, e_Skill.Magia, True, Hechizos(.Stats.UserHechizos(spellSlot)).AreaRadio)
                        Else
                            Call WriteWorkRequestTarget(UserIndex, e_Skill.Magia)
                        End If
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
UseSpellSlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.UseSpellSlot", Erl)
End Sub
