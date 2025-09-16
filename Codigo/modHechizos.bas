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

Private Const FLAUTA_ELFICA             As Long = 40


Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer, Optional ByVal IgnoreVisibilityCheck As Boolean = False)
      On Error GoTo NpcLanzaSpellSobreUser_Err

      Dim Damage As Integer
      Dim DamageStr As String

100   If Spell = 0 Then Exit Sub
      Dim IsAlive As Boolean
      IsAlive = True
102   With UserList(UserIndex)
104     If .flags.Muerto Then Exit Sub
    
        '¿NPC puede ver a través de la invisibilidad?
106     If Not IgnoreVisibilityCheck Then
108       If .flags.invisible = 1 Or .flags.Oculto = 1 Or .flags.Inmunidad = 1 Then Exit Sub
        End If


110     NpcList(NpcIndex).Contadores.IntervaloLanzarHechizo = GetTickCount()
        If Hechizos(Spell).Tipo = uPhysicalSkill Then
          If Not HandlePhysicalSkill(NpcIndex, eNpc, UserIndex, eUser, Spell, IsAlive) Then
              Exit Sub
          End If
        End If
112     Call InfoHechizoDeNpcSobreUser(NpcIndex, UserIndex, Spell)
114     If IsSet(Hechizos(Spell).Effects, e_SpellEffects.eDoHeal) Then
          Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
          Damage = Damage * NPCs.GetMagicHealingBonus(NpcList(NpcIndex))
          Damage = Damage * UserMod.GetSelfHealingBonus(UserList(UserIndex))
          If Damage > 0 Then
116         Call UserMod.DoDamageOrHeal(UserIndex, NpcIndex, eNpc, Damage, e_DamageSourceType.e_magic, Spell)
120         DamageStr = PonerPuntos(Damage)
122         Call WriteLocaleMsg(UserIndex, 32, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name & "¬" & DamageStr)
          End If
          
128     ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.eDoDamage) Then
130       Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
          Damage = Damage * (1 + NpcList(NpcIndex).Stats.MagicBonus)
          ' Si el hechizo no ignora la RM
132       If Hechizos(Spell).AntiRm = 0 Then
            Dim PorcentajeRM As Integer
            PorcentajeRM = GetUserMRForNpc(UserIndex)
            ' Resto el porcentaje total
152         Damage = Damage - Porcentaje(Damage, PorcentajeRM)
            
          End If
          Damage = Damage * NPCs.GetMagicDamageModifier(NpcList(npcIndex))
          Damage = Damage * UserMod.GetMagicDamageReduction(UserList(UserIndex))
154       If Damage < 0 Then Damage = 0
156       IsAlive = UserMod.DoDamageOrHeal(UserIndex, npcIndex, eNpc, -Damage, e_DamageSourceType.e_magic, Spell) = eStillAlive
157       DamageStr = PonerPuntos(Damage)
158       Call WriteLocaleMsg(UserIndex, 1627, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name & "¬" & DamageStr) 'Msg1627=¬1 te ha quitado ¬2 puntos de vida.

162       Call SubirSkill(UserIndex, Resistencia)
          If NpcList(npcIndex).Char.CastAnimation > 0 Then Call SendData(SendTarget.ToNPCAliveArea, npcIndex, PrepareMessageCharAtaca(NpcList(npcIndex).Char.charindex, UserList(UserIndex).Char.charindex, DamageStr, NpcList(npcIndex).Char.CastAnimation))
        End If
        If IsAlive Then
            Dim Effect As IBaseEffectOverTime
            If Hechizos(Spell).EotId > 0 Then
                Set Effect = FindEffectOnTarget(npcIndex, UserList(UserIndex).EffectOverTime, Hechizos(Spell).EotId)
                If Effect Is Nothing Then
                    Call CreateEffect(npcIndex, eNpc, UserIndex, eUser, Hechizos(Spell).EotId)
                Else
                    Call Effect.Reset(npcIndex, eNpc, Hechizos(Spell).EotId)
                End If
            End If
        End If
        'Mana
170     If Hechizos(Spell).SubeMana = 1 Then
172       Damage = RandomNumber(Hechizos(Spell).MinMana, Hechizos(Spell).MaxMana)

174       .Stats.MinMAN = MinimoInt(.Stats.MinMAN + Damage, .Stats.MaxMAN)

176       Call WriteUpdateMana(UserIndex)
178       Call WriteLocaleMsg(UserIndex, 1628, e_FontTypeNames.FONTTYPE_INFO, NpcList(NpcIndex).name & "¬" & Damage) 'Msg1628=¬1 te ha restaurado ¬2 puntos de maná.

180     ElseIf Hechizos(Spell).SubeMana = 2 Then
182       Damage = RandomNumber(Hechizos(Spell).MinMana, Hechizos(Spell).MaxMana)

184       .Stats.MinMAN = MaximoInt(.Stats.MinMAN - Damage, 0)

186       Call WriteUpdateMana(UserIndex)
188       Call WriteLocaleMsg(UserIndex, 1629, e_FontTypeNames.FONTTYPE_INFO, NpcList(NpcIndex).name & "¬" & Damage) 'Msg1629=¬1 te ha quitado ¬2 puntos de maná.
        End If

190     If Hechizos(Spell).SubeAgilidad = 1 Then
192       Damage = RandomNumber(Hechizos(Spell).MinAgilidad, Hechizos(Spell).MaxAgilidad)

194       .flags.TomoPocion = True
196       .flags.DuracionEfecto = Hechizos(Spell).Duration
198       .Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(.Stats.UserAtributos(e_Atributos.Agilidad) + Damage, .Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)

200       Call WriteFYA(UserIndex)
202     ElseIf Hechizos(Spell).SubeAgilidad = 2 Then
204       Damage = RandomNumber(Hechizos(Spell).MinAgilidad, Hechizos(Spell).MaxAgilidad)

206       .flags.TomoPocion = True
208       .flags.DuracionEfecto = Hechizos(Spell).Duration
210       .Stats.UserAtributos(e_Atributos.Agilidad) = MaximoInt(MINATRIBUTOS, .Stats.UserAtributos(e_Atributos.Agilidad) - Damage)

212       Call WriteFYA(UserIndex)
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveDebuff) Then
            Dim NegativeEffect As IBaseEffectOverTime
            Set NegativeEffect = EffectsOverTime.FindEffectOfTypeOnTarget(UserList(UserIndex).EffectOverTime, eDebuff)
            If Not NegativeEffect Is Nothing Then
                NegativeEffect.RemoveMe = True
                Exit Sub
            End If
        End If
        If IsSet(Hechizos(Spell).Effects, e_SpellEffects.StealBuff) Then
            Dim TargetBuff As IBaseEffectOverTime
            Set TargetBuff = EffectsOverTime.FindEffectOfTypeOnTarget(UserList(UserIndex).EffectOverTime, eBuff)
            If Not TargetBuff Is Nothing Then
                Call EffectsOverTime.ChangeOwner(UserIndex, eUser, NpcIndex, eNpc, TargetBuff)
            End If
        End If

214     If Hechizos(Spell).SubeFuerza = 1 Then
216       Damage = RandomNumber(Hechizos(Spell).MinFuerza, Hechizos(Spell).MaxFuerza)
218       .flags.TomoPocion = True
220       .flags.DuracionEfecto = Hechizos(Spell).Duration
222       .Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(.Stats.UserAtributos(e_Atributos.Fuerza) + Damage, .Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)
224       Call WriteFYA(UserIndex)
226     ElseIf Hechizos(Spell).SubeFuerza = 2 Then
228       Damage = RandomNumber(Hechizos(Spell).MinFuerza, Hechizos(Spell).MaxFuerza)
230       .flags.TomoPocion = True
232       .flags.DuracionEfecto = Hechizos(Spell).Duration
234       .Stats.UserAtributos(e_Atributos.Fuerza) = MaximoInt(MINATRIBUTOS, .Stats.UserAtributos(e_Atributos.Fuerza) - Damage)
236       Call WriteFYA(UserIndex)
        End If
238     If IsSet(Hechizos(Spell).Effects, e_SpellEffects.Paralize) Then
240       If .flags.Paralizado = 0 Then
242         .flags.Paralizado = 1
244         .Counters.Paralisis = Hechizos(Spell).Duration / 2

246         Call WriteParalizeOK(UserIndex)
248         Call WritePosUpdate(UserIndex)
          End If
        End If
250     If IsSet(Hechizos(Spell).Effects, e_SpellEffects.Immobilize) Then
252       If .flags.Inmovilizado = 0 Then
254         .flags.Inmovilizado = 1
256         .Counters.Inmovilizado = Hechizos(Spell).Duration / 2

258         Call WriteInmovilizaOK(UserIndex)
260         Call WritePosUpdate(UserIndex)
          End If
        End If

262     If IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveParalysis) Then
264       If .flags.Paralizado > 0 Then
266         .flags.Paralizado = 0
268         .Counters.Paralisis = 0

270         Call WriteParalizeOK(UserIndex)
          End If

272       If .flags.Inmovilizado > 0 Then
274         .flags.Inmovilizado = 0
276         .Counters.Inmovilizado = 0

278         Call WriteInmovilizaOK(UserIndex)
          End If

280       Call WritePosUpdate(UserIndex)
        End If

282     If IsSet(Hechizos(Spell).Effects, e_SpellEffects.Incinerate) Then
284       If .flags.Incinerado = 0 Then
286         .flags.Incinerado = 1
288         .Counters.Incineracion = Hechizos(Spell).Duration

290         Call WriteLocaleMsg(UserIndex, 1630, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name) 'Msg1630=Has sido incinerado por ¬1.
          End If
        End If

292     If Hechizos(Spell).Envenena > 0 Then
294       If .flags.Envenenado = 0 Then
296         .flags.Envenenado = Hechizos(Spell).Envenena
298         .Counters.Veneno = Hechizos(Spell).Duration
300         Call WriteLocaleMsg(UserIndex, 1631, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name) 'Msg1631=Has sido envenenado por ¬1.
          End If
        End If

302     If IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveInvisibility) Then
304         Call UserMod.RemoveInvisibility(UserIndex)
        End If

318     If IsSet(Hechizos(Spell).Effects, e_SpellEffects.Dumb) Then
320       If .flags.Estupidez = 0 Then
322         .flags.Estupidez = IsSet(Hechizos(Spell).Effects, e_SpellEffects.Dumb)
324         .Counters.Estupidez = Hechizos(Spell).Duration
326         Call WriteLocaleMsg(UserIndex, 1632, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name) 'Msg1632=Has sido estupidizado por ¬1.
328         Call WriteDumb(UserIndex)
          End If
330     ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveDumb) Then
332       If .flags.Estupidez > 0 Then
334         .flags.Estupidez = 0
336         .Counters.Estupidez = 0
338         Call WriteLocaleMsg(UserIndex, 1633, e_FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name) 'Msg1633=¬1 te removió la estupidez.
340         Call WriteDumbNoMore(UserIndex)
          End If
        End If

342     If Hechizos(Spell).velocidad > 0 Then
344       If .Counters.velocidad = 0 Then
346         .flags.VelocidadHechizada = Hechizos(Spell).velocidad
348         .Counters.velocidad = Hechizos(Spell).Duration
350         Call ActualizarVelocidadDeUsuario(UserIndex)
          End If
        End If
        If NpcList(NpcIndex).Char.CastAnimation > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(NpcList(NpcIndex).Char.charindex, NpcList(NpcIndex).Char.CastAnimation))
        ElseIf NpcList(NpcIndex).Char.Ataque1 > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(NpcList(NpcIndex).Char.charindex, NpcList(NpcIndex).Char.Ataque1))
        ElseIf NpcList(NpcIndex).Char.WeaponAnim > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageArmaMov(NpcList(NpcIndex).Char.charindex, 0))
        End If
      End With
      
      With NpcList(NpcIndex)
        If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eDisplayCastMessage) Then
          Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, _
                        PrepareMessageChatOverHead("PMAG*" & Spell, .Char.charindex, vbCyan, True, _
                                                   .pos.x, .pos.y, RequiredSpellDisplayTime, MaxInvisibleSpellDisplayTime))
        End If
        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteConsoleMsg(UserIndex, "HecMSGA*" & Spell & "*" & .name, e_FontTypeNames.FONTTYPE_FIGHT)
        End If
      End With
      
      Exit Sub

NpcLanzaSpellSobreUser_Err:
352   Call TraceError(Err.Number, Err.Description & " Hechizo: " & Spell, "modHechizos.NpcLanzaSpellSobreUser", Erl)


End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
      On Error GoTo NpcLanzaSpellSobreNpc_Err

      Dim Damage As Integer
      Dim DamageStr As String
      Dim IsAlive As Boolean
      IsAlive = True
      If Hechizos(Spell).Tipo = uPhysicalSkill Then
        If Not HandlePhysicalSkill(NpcIndex, eNpc, TargetNPC, eNpc, Spell, IsAlive) Then
            Exit Sub
        End If
      End If
100   With NpcList(TargetNPC)
  
102     .Contadores.IntervaloLanzarHechizo = GetTickCount()
  
104     If IsSet(Hechizos(Spell).Effects, e_SpellEffects.eDoHeal) Then ' Cura
106       Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
          Damage = Damage * NPCs.GetMagicHealingBonus(NpcList(NpcIndex))
          Damage = Damage * NPCs.GetSelfHealingBonus(NpcList(TargetNPC))
108       DamageStr = PonerPuntos(Damage)
          If Hechizos(Spell).wav > 0 Then
110           Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .pos.X, .pos.y))
          End If
          If Hechizos(Spell).FXgrh > 0 Then
112           Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
          End If
          If Damage > 0 Then
114           Call SendData(SendTarget.ToPCAliveArea, TargetNPC, PrepareMessageTextCharDrop(DamageStr, .Char.charindex, vbGreen))
          End If
116       Call NPCs.DoDamageOrHeal(TargetNPC, npcIndex, eNpc, Damage, e_DamageSourceType.e_magic, Spell)
120       Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageNpcUpdateHP(TargetNPC))

122     ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.eDoDamage) Then
124       Damage = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
          Damage = Damage * NPCs.GetMagicDamageModifier(NpcList(npcIndex))
          Damage = Damage * NPCs.GetMagicDamageReduction(NpcList(TargetNPC))
126       Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.y))
128       Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
130       IsAlive = NPCs.DoDamageOrHeal(TargetNPC, npcIndex, eNpc, -Damage, e_DamageSourceType.e_magic, Spell) = eStillAlive
134       If .NPCtype = DummyTarget Then
136         Call DummyTargetAttacked(TargetNPC)
          End If
          
154     ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.Paralize) Then

156       If .flags.Paralizado = 0 Then
158         Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.y))
160         Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
162         .flags.Paralizado = 1
164         .Contadores.Paralisis = Hechizos(Spell).Duration / 2
          End If

166     ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.Immobilize) Then
168       If .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
170         Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.y))
172         Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
174         .flags.Inmovilizado = 1
176         .Contadores.Inmovilizado = Hechizos(Spell).Duration / 2
          End If
178     ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.RemoveParalysis) Then
180       If .flags.Paralizado + .flags.Inmovilizado > 0 Then
182         Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.y))
184         Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
186         .flags.Paralizado = 0
188         .Contadores.Paralisis = 0
190         .flags.Inmovilizado = 0
192         .Contadores.Inmovilizado = 0
          End If
194     ElseIf IsSet(Hechizos(Spell).Effects, e_SpellEffects.Incinerate) Then
196       If .flags.Incinerado = 0 Then
198         Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.y))

200         If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
202           Call SendData(SendTarget.ToNPCAliveArea, TargetNPC, PrepareMessageParticleFX(.Char.charindex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))

            End If

204         .flags.Incinerado = 1
          End If
        End If
        If IsAlive Then
            Dim Effect As IBaseEffectOverTime
            If Hechizos(Spell).EotId > 0 Then
                Set Effect = FindEffectOnTarget(npcIndex, NpcList(TargetNPC).EffectOverTime, Hechizos(Spell).EotId)
                If Effect Is Nothing Then
                    Call CreateEffect(npcIndex, eNpc, TargetNPC, eNpc, Hechizos(Spell).EotId)
                Else
                    Call Effect.Reset(npcIndex, eNpc, Hechizos(Spell).EotId)
                End If
            End If
        End If
      End With
      With NpcList(NpcIndex)
        If .Char.CastAnimation > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(.Char.charindex, .Char.CastAnimation))
        ElseIf .Char.Ataque1 > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(.Char.charindex, .Char.Ataque1))
        ElseIf .Char.WeaponAnim > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageArmaMov(.Char.charindex, 0))
        End If
        If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eDisplayCastMessage) Then
          Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, _
                        PrepareMessageChatOverHead("PMAG*" & Spell, .Char.charindex, vbCyan, True, _
                                                   .pos.x, .pos.y, RequiredSpellDisplayTime, MaxInvisibleSpellDisplayTime))
        End If
      End With
      Exit Sub

NpcLanzaSpellSobreNpc_Err:
206   Call TraceError(Err.Number, Err.Description, "modHechizos.NpcLanzaSpellSobreNpc", Erl)


End Sub

Public Sub NpcLanzaSpellSobreArea(ByVal NpcIndex As Integer, ByVal SpellIndex As Integer)
        On Error GoTo NpcLanzaSpellSobreArea_Err
    
        Dim afectaUsers As Boolean
        Dim afectaNPCs As Boolean
        Dim TargetMap As t_MapBlock
        Dim PosCasteadaX As Integer
        Dim PosCasteadaY As Integer
        Dim X            As Long
        Dim Y            As Long
        Dim mitadAreaRadio As Integer
      
100     NpcList(NpcIndex).Contadores.IntervaloLanzarHechizo = GetTickCount()
    
102     With Hechizos(SpellIndex)
104         afectaUsers = (.AreaAfecta = 1 Or .AreaAfecta = 3)
106         afectaNPCs = (.AreaAfecta = 2 Or .AreaAfecta = 3)
108         mitadAreaRadio = CInt(.AreaRadio / 2)
        
110         If IsValidUserRef(NpcList(npcIndex).TargetUser) Then
112             PosCasteadaX = UserList(NpcList(npcIndex).TargetUser.ArrayIndex).pos.x + RandomNumber(-2, 2)
114             PosCasteadaY = UserList(NpcList(npcIndex).TargetUser.ArrayIndex).pos.y + RandomNumber(-2, 2)
            Else
116             PosCasteadaX = NpcList(NpcIndex).Pos.X + RandomNumber(-2, 2)
118             PosCasteadaY = NpcList(NpcIndex).Pos.Y + RandomNumber(-1, 2)
            End If
       
120         For X = 1 To .AreaRadio
122             For Y = 1 To .AreaRadio
                    
                    If InMapBounds(NpcList(NpcIndex).Pos.map, X + PosCasteadaX - mitadAreaRadio, PosCasteadaY + y - mitadAreaRadio) Then
124                     TargetMap = MapData(NpcList(NpcIndex).Pos.map, X + PosCasteadaX - mitadAreaRadio, PosCasteadaY + y - mitadAreaRadio)
                    
126                     If afectaUsers And TargetMap.UserIndex > 0 Then
128                         If Not UserList(TargetMap.UserIndex).flags.Muerto And Not EsGM(TargetMap.UserIndex) Then
130                             Call NpcLanzaSpellSobreUser(NpcIndex, TargetMap.UserIndex, SpellIndex, True)
                            End If
    
                        End If
                                
132                     If afectaNPCs And TargetMap.NpcIndex > 0 Then
134                         If NpcList(TargetMap.NpcIndex).Attackable Then
136                             Call NpcLanzaSpellSobreNpc(NpcIndex, TargetMap.NpcIndex, SpellIndex)
                            End If
    
                        End If
                    End If
138             Next Y
140         Next X

            ' El NPC invoca otros npcs independientes
142         If .Invoca = 1 Then
144             For X = 1 To .cant
                    If NpcList(NpcIndex).Contadores.CriaturasInvocadas >= NpcList(NpcIndex).Stats.CantidadInvocaciones Then
                        Exit Sub
                    Else
                        Dim npcInvocadoIndex As Integer
146                      npcInvocadoIndex = SpawnNpc(.NumNpc, NpcList(NpcIndex).Pos, True, False, False)
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
148             Next X
            End If

        End With
        With NpcList(NpcIndex)
          If .Char.CastAnimation > 0 Then
              Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(.Char.charindex, .Char.CastAnimation))
          ElseIf .Char.Ataque1 > 0 Then
              Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(.Char.charindex, .Char.Ataque1))
          ElseIf .Char.WeaponAnim > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageArmaMov(.Char.charindex, 0))
          End If
          If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eDisplayCastMessage) Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, _
                          PrepareMessageChatOverHead("PMAG*" & SpellIndex, .Char.charindex, vbCyan, True, _
                                                     .pos.x, .pos.y, RequiredSpellDisplayTime, MaxInvisibleSpellDisplayTime))
          End If
        End With
        Exit Sub

NpcLanzaSpellSobreArea_Err:
150     Call TraceError(Err.Number, Err.Description, "modHechizos.NpcLanzaSpellSobreArea", Erl)

        
End Sub


Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

        On Error GoTo ErrHandler
    
        Dim j As Integer

100     For j = 1 To MAXUSERHECHIZOS

102         If UserList(UserIndex).Stats.UserHechizos(j) = i Then
104             TieneHechizo = True
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

100     hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

102     If Not TieneHechizo(hIndex, UserIndex) Then

            'Buscamos un slot vacio
104         For j = 1 To MAXUSERHECHIZOS
106             If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
108         Next j
        
110         If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
'Msg777= No tenes espacio para mas hechizos.
Call WriteLocaleMsg(UserIndex, "777", e_FontTypeNames.FONTTYPE_INFO)

            Else
114             UserList(UserIndex).Stats.UserHechizos(j) = hIndex

116             Call UpdateUserHechizos(False, UserIndex, CByte(j))

                'Quitamos del inv el item
118             Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)

            End If
            
            UserList(UserIndex).flags.ModificoHechizos = True
        Else
120         ' Msg525=Ya tenes ese hechizo.
            Call WriteLocaleMsg(UserIndex, "525", e_FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

AgregarHechizo_Err:
122     Call TraceError(Err.Number, Err.Description, "modHechizos.AgregarHechizo", Erl)

        
End Sub
            
Sub DecirPalabrasMagicas(ByVal Hechizo As Integer, ByVal UserIndex As Integer)
On Error GoTo DecirPalabrasMagicas_Err
        UserList(UserIndex).Counters.timeChat = 4
        If Not IsVisible(UserList(UserIndex)) Then
100         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(UserIndex).Char.charindex, vbCyan, True, UserList(UserIndex).pos.X, UserList(UserIndex).pos.y, RequiredSpellDisplayTime, MaxInvisibleSpellDisplayTime))
        Else
102         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(UserIndex).Char.charindex, vbCyan, True, UserList(UserIndex).pos.X, UserList(UserIndex).pos.y, 0, 0))
        End If

        Exit Sub
DecirPalabrasMagicas_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.DecirPalabrasMagicas", Erl)
End Sub

Private Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer, Optional ByVal Slot As Integer = 0) As Boolean
        On Error GoTo PuedeLanzar_Err

100     PuedeLanzar = False

        If HechizoIndex = 0 Then Exit Function
        
102     With UserList(UserIndex)

            'Si lanza a un npc y este es solo atacable para clanes y el usuario no tiene clan, le avisa y sale de la funcion
            If IsValidNpcRef(.flags.TargetNPC) Then
                If NpcList(.flags.TargetNPC.ArrayIndex).OnlyForGuilds = 1 And .GuildIndex <= 0 Then
                    'Msg2001=Debes pertenecer a un clan para atacar a este NPC
                    Call WriteLocaleMsg(UserIndex, "2001", e_FontTypeNames.FONTTYPE_WARNING)
                    Exit Function
                End If
            End If

104         If .flags.EnConsulta Then
'Msg778= No puedes lanzar hechizos si estas en consulta.
Call WriteLocaleMsg(UserIndex, "778", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
            
            If Hechizos(HechizoIndex).AutoLanzar And .flags.TargetUser.ArrayIndex <> UserIndex Then
                Exit Function
            End If
            
            If IsSet(.flags.StatusMask, eCastOnlyOnSelf) And .flags.targetUser.ArrayIndex <> UserIndex Then
                Call WriteLocaleMsg(UserIndex, MsgCastOnlyOnSelf, e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
            
108         If .flags.Muerto = 1 Then
110             Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
            
            If IsSet(Hechizos(HechizoIndex).Effects, e_SpellEffects.CancelActiveEffect) And _
                Hechizos(HechizoIndex).EotId > 0 And _
                IsValidUserRef(.flags.targetUser) Then
                Dim Effect As IBaseEffectOverTime
                Set Effect = FindEffectOnTarget(UserIndex, UserList(.flags.targetUser.ArrayIndex).EffectOverTime, Hechizos(HechizoIndex).EotId)
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
112         If .flags.Privilegios And e_PlayerType.Consejero Then
                Exit Function
            End If

114         If MapInfo(.pos.Map).SinMagia And Not IsSet(Hechizos(HechizoIndex).SpellRequirementMask, eIsSkill) Then
'Msg779= Una fuerza mística te impide lanzar hechizos en esta zona.
Call WriteLocaleMsg(UserIndex, "779", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Function
            End If
        
118         If .flags.Montado = 1 Then
'Msg780= No puedes lanzar hechizos si estas montado.
Call WriteLocaleMsg(UserIndex, "780", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

            If Hechizos(HechizoIndex).NecesitaObj > 0 Then
              If Not IsObjecIndextInInventory(UserIndex, Hechizos(HechizoIndex).NecesitaObj) And _
                 Not IsObjecIndextInInventory(UserIndex, Hechizos(HechizoIndex).NecesitaObj2) Then
                    If Not IsObjecIndextInInventory(UserIndex, Hechizos(HechizoIndex).NecesitaObj) And Not IsObjecIndextInInventory(UserIndex, Hechizos(HechizoIndex).NecesitaObj2) Then
                        Call WriteLocaleMsg(UserIndex, 1634, e_FontTypeNames.FONTTYPE_INFO, ObjData(Hechizos(HechizoIndex).NecesitaObj).name) 'Msg1634=Necesitas un ¬1 para lanzar el hechizo.
                        Exit Function
                    End If
                End If
            End If
            
            If IsValidUserRef(.flags.targetUser) Then
                If Hechizos(HechizoIndex).TargetEffectType = e_TargetEffectType.ePositive Then
                    Dim UserInteractionResult As e_InteractionResult
                    UserInteractionResult = UserMod.CanHelpUser(UserIndex, .flags.targetUser.ArrayIndex)
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

                    If UserAttackInteractionResult.result = e_AttackInteractionResult.eAttackCitizenNpc Or _
                       UserAttackInteractionResult.result = e_AttackInteractionResult.eRemoveSafeCitizenNpc Or _
                       UserAttackInteractionResult.Result = e_AttackInteractionResult.eSameFaction Or _
                       UserAttackInteractionResult.result = e_AttackInteractionResult.eRemoveSafe Then
                        Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.result)
                        If UserAttackInteractionResult.CanAttack Then
                            If UserAttackInteractionResult.TurnPK Then VolverCriminal (UserIndex)
                        Else
                            Exit Function
                        End If
                    End If
                End If
            End If

128         If Hechizos(HechizoIndex).Cooldown > 0 And .Counters.UserHechizosInterval(Slot) > 0 Then
                Dim Actual As Long
                Dim SegundosFaltantes As Long
130             Actual = GetTickCount()
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
132             If .Counters.UserHechizosInterval(Slot) + Cooldown > Actual Then
134                 SegundosFaltantes = Int((.Counters.UserHechizosInterval(Slot) + Cooldown - Actual) / 1000)
136                 Call WriteLocaleMsg(UserIndex, 1635, e_FontTypeNames.FONTTYPE_WARNING, SegundosFaltantes) 'Msg1635=Debes esperar ¬1 segundos para volver a tirar este hechizo.
                    Exit Function
                End If
            End If

138         If .Stats.UserSkills(e_Skill.Magia) < Hechizos(HechizoIndex).MinSkill Then
140             Call WriteLocaleMsg(UserIndex, 1636, e_FontTypeNames.FONTTYPE_INFO, Hechizos(HechizoIndex).MinSkill) 'Msg1636=No tienes suficientes puntos de magia para lanzar este hechizo, necesitas ¬1 puntos.
                Exit Function
            End If

142         If .Stats.MinHp < Hechizos(HechizoIndex).RequiredHP Then
144             Call WriteLocaleMsg(UserIndex, 1637, e_FontTypeNames.FONTTYPE_INFO, Hechizos(HechizoIndex).RequiredHP) 'Msg1637=No tienes suficiente vida. Necesitas ¬1 puntos de vida.
                Exit Function
            End If
            
146         If .Stats.MinMAN < ManaHechizoPorClase(UserIndex, Hechizos(HechizoIndex), HechizoIndex) Then
148             Call WriteLocaleMsg(UserIndex, "222", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

150         If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
152             Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
            
154         If .clase = e_Class.Mage And Not IsFeatureEnabled("remove-staff-requirements") Then
156             If Hechizos(HechizoIndex).NeedStaff > 0 Then
158                 If .invent.EquippedWeaponObjIndex = 0 Then
                        'Msg781= Necesitás un báculo para lanzar este hechizo.
                        Call WriteLocaleMsg(UserIndex, "781", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                
162                 If ObjData(.invent.EquippedWeaponObjIndex).Power < Hechizos(HechizoIndex).NeedStaff Then
                        'Msg782= Necesitás un báculo más poderoso para lanzar este hechizo.
                        Call WriteLocaleMsg(UserIndex, "782", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                End If
            End If
            
            If .clase = e_Class.Druid Then
                If Hechizos(HechizoIndex).RequiereInstrumento > 0 Then
                    If .invent.EquippedRingAccesoryObjIndex = 0 Or ObjData(.invent.EquippedRingAccesoryObjIndex).InstrumentoRequerido <> 1 Then
                        'Msg783= Necesitás una flauta para invocar o desinvocar a tus mascotas.
                        Call WriteLocaleMsg(UserIndex, "783", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Function
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
            If IsValidUserRef(.flags.targetUser) Then
                Call CastUserToAnyRef(.flags.targetUser, TargetRef)
            ElseIf IsValidNpcRef(.flags.TargetNPC) Then
                Call CastNpcToAnyRef(.flags.TargetNPC, TargetRef)
            End If
            
            If IsValidRef(TargetRef) Then
                If IsDead(TargetRef) And Not IsSet(Hechizos(HechizoIndex).SpellRequirementMask, e_SpellRequirementMask.eWorkOnDead) Then
                    Call WriteLocaleMsg(UserIndex, 7, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
            
            If IsSet(Hechizos(HechizoIndex).SpellRequirementMask, e_SpellRequirementMask.eRequireTargetOnLand) And _
               IsValidRef(TargetRef) Then
                If TargetRef.RefType = eUser Then
                    If UserList(TargetRef.ArrayIndex).flags.Nadando > 0 Or _
                       .flags.Navegando > 0 Or .flags.Montado > 0 Then
                        Call WriteLocaleMsg(UserIndex, MsgLandRequiredToUseSpell, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                End If
            End If
            If IsSet(Hechizos(HechizoIndex).SpellRequirementMask, e_SpellRequirementMask.eRequireTargetOnWater) And _
               IsValidRef(TargetRef) Then
               If TargetRef.RefType = eUser Then
                    If UserList(TargetRef.ArrayIndex).flags.Nadando = 0 And _
                       .flags.Navegando = 0 Then
                        Call WriteLocaleMsg(UserIndex, MsgWaterRequiredToUseSpell, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                End If
            End If
166         PuedeLanzar = True
        End With

        Exit Function

PuedeLanzar_Err:
168     Call TraceError(Err.Number, Err.Description, "modHechizos.PuedeLanzar", Erl)

        
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
    
100     With UserList(UserIndex)

102         If .flags.EnReto Then
'Msg784= No podés invocar criaturas durante un reto.
Call WriteLocaleMsg(UserIndex, "784", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
    
            Dim h As Integer, j As Integer, ind As Integer, Index As Integer
            Dim targetPos As t_WorldPos
    
106         targetPos.Map = .flags.TargetMap
108         targetPos.X = .flags.TargetX
110         targetPos.Y = .flags.TargetY
        
112         h = .Stats.UserHechizos(.flags.Hechizo)
    
114         If Hechizos(h).Invoca = 1 Then
        
                ' No puede invocar en este mapa
                If MapInfo(.Pos.Map).NoMascotas Then
'Msg785= Un gran poder te impide invocar criaturas en este mapa.
Call WriteLocaleMsg(UserIndex, "785", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                Dim MinTiempo As Integer
                Dim i As Integer
                
                For i = 1 To Hechizos(h).cant
                    Index = -1
                    MinTiempo = IntervaloInvocacion
                    For j = 1 To MAXMASCOTAS
                        If .MascotasIndex(j).ArrayIndex > 0 Then
                            If IsValidNpcRef(.MascotasIndex(j)) Then
                                If NpcList(.MascotasIndex(j).ArrayIndex).flags.NPCActive Then
                                    If NpcList(.MascotasIndex(j).ArrayIndex).Contadores.TiempoExistencia > 0 And NpcList(.MascotasIndex(j).ArrayIndex).Contadores.TiempoExistencia < MinTiempo Then
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
                        ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False, False, UserIndex)
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
            
160             Call InfoHechizo(UserIndex)
162             b = True
        
164         ElseIf Hechizos(h).Invoca = 2 Then
            
                ' Si tiene mascotas
166             If .NroMascotas > 0 Then
                    ' Tiene que estar en zona insegura
                    
                    ' No puede invocar en este mapa
                    If MapInfo(.Pos.Map).NoMascotas Then
                        Call WriteLocaleMsg(UserIndex, "786", e_FontTypeNames.FONTTYPE_INFO) 'Msg786= Un gran poder te impide invocar criaturas en este mapa.
                        Exit Sub
                    End If
                
                    ' Si no están guardadas las mascotas
170                 If .flags.MascotasGuardadas = 0 Then
172                     For i = 1 To MAXMASCOTAS
174                         If IsValidNpcRef(.MascotasIndex(i)) Then
                                ' Si no es un elemental, lo "guardamos"... lo matamos
176                             If NpcList(.MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia = 0 Then
                                    ' Le saco el maestro, para que no me lo quite de mis mascotas
178                                 Call SetUserRef(NpcList(.MascotasIndex(i).ArrayIndex).MaestroUser, 0)
                                    ' Lo borro
180                                 Call QuitarNPC(.MascotasIndex(i).ArrayIndex, eStorePets)
                                    ' Saco el índice
182                                 Call ClearNpcRef(.MascotasIndex(i))
184                                 b = True
                                End If
                            Else
                                Call ClearNpcRef(.MascotasIndex(i))
                            End If
                        Next
186                     .flags.MascotasGuardadas = 1

                    ' Ya están guardadas, así que las invocamos
                    Else
188                     For i = 1 To MAXMASCOTAS
                            ' Si está guardada y no está ya en el mapa
190                         If .MascotasType(i) > 0 And .MascotasIndex(i).ArrayIndex = 0 Then
192                             Call SetNpcRef(.MascotasIndex(i), SpawnNpc(.MascotasType(i), targetPos, True, True, False, UserIndex))
194                             Call SetUserRef(NpcList(.MascotasIndex(i).ArrayIndex).MaestroUser, UserIndex)
196                             Call FollowAmo(.MascotasIndex(i).ArrayIndex)

                                If IsFeatureEnabled("addjust-npc-with-caster") And IsSet(Hechizos(h).Effects, AdjustStatsWithCaster) Then
                                    Call AdjustNpcStatWithCasterLevel(UserIndex, .MascotasIndex(i).ArrayIndex)
                                End If

198                             b = True
                            End If
                        Next
200                     .flags.MascotasGuardadas = 0
                    End If
            
                Else
'Msg787= No tienes mascotas.
Call WriteLocaleMsg(UserIndex, "787", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

206             If b Then Call InfoHechizo(UserIndex)
            
            End If
    
        End With
    
        Exit Sub
    
HechizoInvocacion_Err:
208     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoInvocacion")


End Sub

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoTerrenoEstado_Err
        

        Dim PosCasteadaX As Integer

        Dim PosCasteadaY As Integer

        Dim PosCasteadaM As Integer

        Dim h            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     PosCasteadaX = UserList(UserIndex).flags.TargetX
102     PosCasteadaY = UserList(UserIndex).flags.TargetY
104     PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
106     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
        
108     If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveInvisibility) Then
110         b = True
112         For TempX = PosCasteadaX - 11 To PosCasteadaX + 11
114             For TempY = PosCasteadaY - 11 To PosCasteadaY + 11
116                 If InMapBounds(PosCasteadaM, TempX, TempY) Then
118                     If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                            'hay un user
120                         If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.NoDetectable = 0 Then
122                             UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 0
124                             Call WriteConsoleMsg(MapData(PosCasteadaM, TempX, TempY).UserIndex, PrepareMessageLocaleMsg(1869, vbNullString, e_FontTypeNames.FONTTYPE_INFOIAO)) ' Msg1869=Tu invisibilidad ya no tiene efecto.
126                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.charindex, False, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                            End If
                        End If
                    End If
128             Next TempY
130         Next TempX
132         Call InfoHechizo(UserIndex)
        End If
        Exit Sub
HechizoTerrenoEstado_Err:
134     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoTerrenoEstado", Erl)
End Sub

Private Sub HechizoSobreArea(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoSobreArea_Err
        
        Dim afectaUsers As Boolean
        Dim afectaNPCs As Boolean
        Dim TargetMap As t_MapBlock
        Dim PosCasteadaX As Byte
        Dim PosCasteadaY As Byte
        Dim h            As Integer
        Dim X            As Long
        Dim Y            As Long
 
100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
102     PosCasteadaX = UserList(UserIndex).flags.TargetX
104     PosCasteadaY = UserList(UserIndex).flags.TargetY
        
        'Envio Palabras magicas, wavs y fxs.
106     If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
108         Call DecirPalabrasMagicas(h, UserIndex)

        End If
    
    
118     If Hechizos(h).Particle > 0 Then 'Envio Particula?
120         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(PosCasteadaX, PosCasteadaY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
        End If

122     If Hechizos(h).ParticleViaje = 0 Then
124         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, PosCasteadaX, PosCasteadaY))
        End If
        
126     afectaUsers = (Hechizos(h).AreaAfecta = 1 Or Hechizos(h).AreaAfecta = 3)
128     afectaNPCs = (Hechizos(h).AreaAfecta = 2 Or Hechizos(h).AreaAfecta = 3)
       
130     For X = 1 To Hechizos(h).AreaRadio
132         For Y = 1 To Hechizos(h).AreaRadio
134             TargetMap = MapData(UserList(UserIndex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2))
                                
136             If afectaUsers And TargetMap.UserIndex > 0 Then
138                 If UserList(TargetMap.UserIndex).flags.Muerto = 0 Then
140                     Call AreaHechizo(UserIndex, TargetMap.UserIndex, PosCasteadaX, PosCasteadaY, False)

1110                     If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
                    
1112                         If Hechizos(h).ParticleViaje > 0 Then
1114                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, X, Y))
                                
                            Else
1116                             Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(TargetMap.userindex).Pos.X, UserList(TargetMap.userindex).Pos.y))
                
                            End If
                
                        End If
                    End If

                End If
                            
142             If afectaNPCs And TargetMap.NpcIndex > 0 Then
144                 If NpcList(TargetMap.NpcIndex).Attackable Then
11110
146                     Call AreaHechizo(UserIndex, TargetMap.NpcIndex, PosCasteadaX, PosCasteadaY, True)
                        If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
                    
11112                         If Hechizos(h).ParticleViaje > 0 Then
11114                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, X, Y))
                                
                            Else
11116                             Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageFxPiso(Hechizos(h).FXgrh, NpcList(TargetMap.NpcIndex).Pos.X, NpcList(TargetMap.NpcIndex).Pos.y))
                
                            End If
                
                        End If
                    End If

                End If
                            
148         Next Y
150     Next X

152     b = True
        
        Exit Sub

HechizoSobreArea_Err:
154     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoSobreArea", Erl)

        
End Sub

Sub HechizoPortal(ByVal UserIndex As Integer, ByRef b As Boolean)
        On Error GoTo HechizoPortal_Err
        

        Dim PosCasteadaX As Byte

        Dim PosCasteadaY As Byte

        Dim PosCasteadaM As Integer

        Dim uh           As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     PosCasteadaX = UserList(UserIndex).flags.TargetX
102     PosCasteadaY = UserList(UserIndex).flags.TargetY
104     PosCasteadaM = UserList(UserIndex).flags.TargetMap
 
106     uh = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
        'Envio Palabras magicas, wavs y fxs.
   
108     If MapData(UserList(UserIndex).pos.map, UserList(UserIndex).flags.targetX, UserList(UserIndex).flags.targetY).ObjInfo.amount > 0 Or (MapData(UserList(UserIndex).pos.map, UserList(UserIndex).flags.targetX, UserList(UserIndex).flags.targetY).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES Or MapData(UserList(UserIndex).pos.map, UserList(UserIndex).flags.targetX, UserList(UserIndex).flags.targetY).TileExit.map > 0 Or UserList(UserIndex).flags.targetUser.ArrayIndex <> 0 Then
110         b = False
            'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", e_FontTypeNames.FONTTYPE_INFO)
112         Call WriteLocaleMsg(UserIndex, "262", e_FontTypeNames.FONTTYPE_INFO)

        Else

114         If Hechizos(uh).TeleportX = 1 Then

116             If UserList(UserIndex).flags.Portal = 0 Then

118                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.Runa, -1, False))
         
120                 UserList(UserIndex).flags.PortalM = UserList(UserIndex).Pos.Map
122                 UserList(UserIndex).flags.PortalX = UserList(UserIndex).flags.TargetX
124                 UserList(UserIndex).flags.PortalY = UserList(UserIndex).flags.TargetY
            
126                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 600, e_AccionBarra.Intermundia))

128                 UserList(UserIndex).Accion.AccionPendiente = True
130                 UserList(UserIndex).Accion.Particula = e_ParticleEffects.Runa
132                 UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.Intermundia
134                 UserList(UserIndex).Accion.HechizoPendiente = uh
            
136                 If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
138                     Call DecirPalabrasMagicas(uh, UserIndex)

                    End If

140                 b = True
                Else
'Msg788= No podés lanzar mas de un portal a la vez.
Call WriteLocaleMsg(UserIndex, "788", e_FontTypeNames.FONTTYPE_INFO)
144                 b = False

                End If

            End If

        End If

        Exit Sub

HechizoPortal_Err:
146     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoPortal", Erl)

        
End Sub

Sub HechizoMaterializacion(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoMaterializacion_Err
        

        Dim h   As Integer

        Dim MAT As t_Obj

100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
 
102     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Then
104         b = False
106         Call WriteLocaleMsg(UserIndex, "262", e_FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", e_FontTypeNames.FONTTYPE_INFO)
        Else
108         MAT.amount = Hechizos(h).MaterializaCant
110         MAT.ObjIndex = Hechizos(h).MaterializaObj
112         Call MakeObj(MAT, UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY)
114         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
116         b = True

        End If

        
        Exit Sub

HechizoMaterializacion_Err:
118     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoMaterializacion", Erl)

        
End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)
        
        On Error GoTo HandleHechizoTerreno_Err
        

        '***************************************************
        'Author: Unknown
        'Last Modification: 01/10/07
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
        'usuario
        '***************************************************
        Dim b As Boolean

100     Select Case Hechizos(uh).Tipo
        
            Case e_TipoHechizo.uInvocacion 'Tipo 1
102             Call HechizoInvocacion(UserIndex, b)

104         Case e_TipoHechizo.uEstado 'Tipo 2
106             Call HechizoTerrenoEstado(UserIndex, b)

108         Case e_TipoHechizo.uMaterializa 'Tipo 3
110             Call HechizoMaterializacion(UserIndex, b)
            
112         Case e_TipoHechizo.uArea 'Tipo 5
114             Call HechizoSobreArea(UserIndex, b)
            
116         Case e_TipoHechizo.uPortal 'Tipo 6
118             Call HechizoPortal(UserIndex, b)

            Case e_TipoHechizo.uMultiShoot
                Dim targetPos As t_WorldPos
                targetPos.map = UserList(UserIndex).pos.map
                targetPos.x = UserList(UserIndex).flags.targetX
                targetPos.y = UserList(UserIndex).flags.targetY
                b = MultiShot(UserIndex, targetPos)
        End Select

124     If b Then
            If Not IsSet(Hechizos(uh).SpellRequirementMask, eIsSkill) Then
126             Call SubirSkill(UserIndex, Magia)
            End If
128         UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

130         If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
132         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

134         If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
136         Call WriteUpdateMana(UserIndex)
138         Call WriteUpdateSta(UserIndex)

        End If

        
        Exit Sub

HandleHechizoTerreno_Err:
140     Call TraceError(Err.Number, Err.Description, "modHechizos.HandleHechizoTerreno", Erl)
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
    End With
    If Not IsSet(Hechizos(uh).SpellRequirementMask, eIsSkill) Then
        Call SubirSkill(UserIndex, Magia)
    End If
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateMana(UserIndex)
    Call WriteUpdateSta(UserIndex)
    HandlePetSpell = True
End Function

Function HandlePhysicalSkill(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                        ByVal SpellIndex As Integer, IsAlive As Boolean) As Boolean
    
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
            Dim Damage As Integer
            Dim ObjectIndex As Integer
            Dim Proyectile As Integer
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
                ObjectIndex = -1
                Proyectile = 1
            End If
            If RefDoDamageToTarget(SourceRef, TargetRef, Damage, e_phisical, ObjectIndex) = eStillAlive Then
                IsAlive = True
                If TargetRef.RefType = eUser Then
                    UserList(TargetRef.ArrayIndex).Counters.timeFx = 3
                    Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(TargetRef.ArrayIndex).Char.charindex, FXSANGRE, 0, UserList(TargetRef.ArrayIndex).pos.x, UserList(TargetRef.ArrayIndex).pos.y))
                    Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(TargetRef.ArrayIndex).pos.x, UserList(TargetRef.ArrayIndex).pos.y))
                Else
                    If NpcList(TargetRef.ArrayIndex).flags.Snd2 > 0 Then
                        Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(NpcList(TargetRef.ArrayIndex).flags.Snd2, NpcList(TargetRef.ArrayIndex).pos.x, NpcList(TargetRef.ArrayIndex).pos.y))
                    Else
                        Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(TargetRef.ArrayIndex).pos.x, NpcList(TargetRef.ArrayIndex).pos.y))
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
            Dim Particula As Integer
            Dim Tiempo    As Long
            Dim CannonProyectile As Integer
            CannonProyectile = 4
            Particula = Hechizos(SpellIndex).Particle
            Tiempo = Hechizos(SpellIndex).TimeParticula
            Call SendData(SendTarget.ToNPCAliveArea, SourceIndex, PrepareMessageParticleFX(NpcList(SourceIndex).Char.charindex, Particula, Tiempo, False, , SourcePos.x, SourcePos.y))
            Call SendData(SendTarget.ToNPCAliveArea, SourceIndex, PrepareCreateProjectile(SourcePos.x, SourcePos.y, TargetPos.x, TargetPos.y, CannonProyectile))
            If Hechizos(SpellIndex).wav <> 0 Then Call SendData(SendTarget.ToNPCAliveArea, SourceIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).wav, SourcePos.x, SourcePos.y))
            Call CreateDelayedBlast(SourceIndex, SourceType, TargetPos.Map, TargetPos.x, TargetPos.y, Hechizos(SpellIndex).EotId, -1)
            HandlePhysicalSkill = False
            Exit Function
    End Select
End Function

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 01/10/07
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
        'usuario
        '***************************************************
        
        On Error GoTo HandleHechizoUsuario_Err
        
        Dim IsAlive As Boolean
        IsAlive = True
        Dim b As Boolean
        Dim Effect As IBaseEffectOverTime
        If Hechizos(uh).EotId > 0 And IsValidUserRef(UserList(UserIndex).flags.targetUser) Then
            Set Effect = FindEffectOnTarget(UserIndex, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).EffectOverTime, Hechizos(uh).EotId)
            If Not Effect Is Nothing Then
                If Not EffectOverTime(Hechizos(uh).EotId).Override Then
                    Call WriteLocaleMsg(UserIndex, MsgTargetAlreadyAffected, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
100     Select Case Hechizos(uh).Tipo

            Case e_TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
102             Call HechizoEstadoUsuario(UserIndex, b)

104         Case e_TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropUsuario(UserIndex, b, IsAlive)

108         Case e_TipoHechizo.uCombinados
110             Call HechizoCombinados(UserIndex, b, IsAlive)
            Case e_TipoHechizo.uPhysicalSkill
                b = HandlePhysicalSkill(UserIndex, eUser, UserList(UserIndex).flags.targetUser.ArrayIndex, eUser, _
                                        UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo), IsAlive)
        End Select

112     If b Then
            If Hechizos(uh).EotId > 0 And IsAlive Then
                If Effect Is Nothing Then
                    Call CreateEffect(UserIndex, eUser, UserList(UserIndex).flags.targetUser.ArrayIndex, eUser, Hechizos(uh).EotId)
                Else
                    If Not Effect.Reset(UserIndex, eUser, Hechizos(uh).EotId) Then
                        Exit Sub
                    End If
                End If
            End If
            If Not IsSet(Hechizos(uh).SpellRequirementMask, eIsSkill) Then
114             Call SubirSkill(UserIndex, Magia)
            End If
                    
116         UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - ManaHechizoPorClase(UserIndex, Hechizos(uh), uh)
           
118         If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0

120         If Hechizos(uh).RequiredHP > 0 Then
                Call UserMod.ModifyHealth(UserIndex, -Hechizos(uh).RequiredHP, 1)
            End If

128         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
130         If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0

            If IsSet(Hechizos(uh).Effects, e_SpellEffects.Resurrect) Then
            
                If Not PeleaSegura(UserIndex, UserList(UserIndex).flags.targetUser.ArrayIndex) Then
                    If MapInfo(UserList(UserIndex).Pos.map).Seguro = 0 Then
                        Dim costoVidaResu As Long
                        costoVidaResu = UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).Stats.ELV * 1.5 + UserList(UserIndex).Stats.MinHp * 0.45
                        Call UserMod.ModifyHealth(UserIndex, -costoVidaResu, 1)
                    End If
                End If
            
            End If
            
132         Call WriteUpdateMana(UserIndex)
            Call WriteUpdateHP(UserIndex)
134         Call WriteUpdateSta(UserIndex)
        End If
        Exit Sub
HandleHechizoUsuario_Err:
138     Call TraceError(Err.Number, Err.Description, "modHechizos.HandleHechizoUsuario", Erl)
End Sub

Public Function ManaHechizoPorClase(ByVal userindex As Integer, Hechizo As t_Hechizo, Optional ByVal HechizoIndex As Long) As Integer
        
    ManaHechizoPorClase = Hechizo.ManaRequerido

    Select Case UserList(UserIndex).clase
    
        Case e_Class.Bard
            If Hechizos(HechizoIndex).nombre = MauveFlashIndex And UserList(UserIndex).invent.EquippedRingAccesoryObjIndex = MagicLuteIndex Then
                ManaHechizoPorClase = 80
                Exit Function
            ElseIf Hechizos(HechizoIndex).nombre = FireEcoIndex And UserList(UserIndex).invent.EquippedRingAccesoryObjIndex = MagicLuteIndex Then
                ManaHechizoPorClase = 70
                Exit Function
            End If
           
    End Select
End Function

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
        
        On Error GoTo HandleHechizoNPC_Err
        

        '***************************************************
        'Author: Unknown
        'Last Modification: 01/10/07
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
        'usuario
        '***************************************************
        Dim b As Boolean
        Dim Effect As IBaseEffectOverTime
        Dim IsAlive As Boolean
        IsAlive = True
        If Hechizos(uh).EotId > 0 And IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
            Set Effect = FindEffectOnTarget(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).EffectOverTime, Hechizos(uh).EotId)
            If Not Effect Is Nothing Then
                If Not EffectOverTime(Hechizos(uh).EotId).Override Then
                    Call WriteLocaleMsg(UserIndex, MsgTargetAlreadyAffected, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
        Call AllMascotasAtacanNPC(UserList(UserIndex).flags.TargetNPC.ArrayIndex, UserIndex)
100     Select Case Hechizos(uh).Tipo
            Case e_TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
102             Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC.ArrayIndex, uh, b, UserIndex)

104         Case e_TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC.ArrayIndex, UserIndex, b, IsAlive)
            Case e_TipoHechizo.uPhysicalSkill
                b = HandlePhysicalSkill(UserIndex, eUser, UserList(UserIndex).flags.TargetNPC.ArrayIndex, eNpc, _
                                        UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo), IsAlive)
        End Select

108     If b Then
            If Hechizos(uh).EotId > 0 And IsAlive Then
                If Effect Is Nothing Then
                    Call CreateEffect(UserIndex, eUser, UserList(UserIndex).flags.TargetNPC.ArrayIndex, eNpc, Hechizos(uh).EotId)
                Else
                    If Not Effect.Reset(UserIndex, eUser, Hechizos(uh).EotId) Then
                        Exit Sub
                    End If
                End If
            End If
            If Not IsSet(Hechizos(uh).SpellRequirementMask, eIsSkill) Then
110             Call SubirSkill(UserIndex, Magia)
            End If
            UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - ManaHechizoPorClase(userindex, Hechizos(uh), uh)
        
116         If Hechizos(uh).RequiredHP > 0 Then
118             If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
120             Call UserMod.ModifyHealth(UserIndex, -Hechizos(uh).RequiredHP, 1)
            End If

126         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

128         If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
130         Call WriteUpdateMana(UserIndex)
132         Call WriteUpdateSta(UserIndex)

        End If

        
        Exit Sub

HandleHechizoNPC_Err:
134     Call TraceError(Err.Number, Err.Description, "modHechizos.HandleHechizoNPC", Erl)

        
End Sub

Sub LanzarHechizo(ByVal Index As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo LanzarHechizo_Err

        Dim uh As Integer
        Dim SpellCastSuccess As Boolean
        
100     uh = UserList(UserIndex).Stats.UserHechizos(Index)

102     If PuedeLanzar(UserIndex, uh, Index) Then

104         Select Case Hechizos(uh).Target

                Case e_TargetType.uUsuarios

106                 If IsValidUserRef(UserList(UserIndex).flags.targetUser) Then
108                     If Abs(UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.y - UserList(UserIndex).pos.y) <= RANGO_VISION_Y Then
110                         Call HandleHechizoUsuario(UserIndex, uh)
                            SpellCastSuccess = True
                        Else
116                         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
'Msg790= Este hechizo actua solo sobre usuarios.
Call WriteLocaleMsg(UserIndex, "790", e_FontTypeNames.FONTTYPE_INFO)
                    End If
        
120             Case e_TargetType.uNPC

122                 If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
124                     If Abs(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Pos.y - UserList(UserIndex).Pos.y) <= RANGO_VISION_Y Then
126                         Call HandleHechizoNPC(UserIndex, uh)
                            SpellCastSuccess = True
                        Else
132                         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
'Msg791= Este hechizo solo afecta a los npcs.
Call WriteLocaleMsg(UserIndex, "791", e_FontTypeNames.FONTTYPE_INFO)
                    End If
        
136             Case e_TargetType.uUsuariosYnpc

138                 If IsValidUserRef(UserList(UserIndex).flags.targetUser) Then
140                     If Abs(UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.y - UserList(UserIndex).pos.y) <= RANGO_VISION_Y Then
142                         Call HandleHechizoUsuario(UserIndex, uh)
                            SpellCastSuccess = True
                        Else
148                         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        End If

150                 ElseIf IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then

152                     If Abs(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Pos.y - UserList(UserIndex).Pos.y) <= RANGO_VISION_Y Then
                            SpellCastSuccess = True
158                         Call HandleHechizoNPC(UserIndex, uh)
                        Else
160                         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
'Msg792= Target invalido.
Call WriteLocaleMsg(UserIndex, "792", e_FontTypeNames.FONTTYPE_INFO)
                    End If
        
164             Case e_TargetType.uTerreno
                    SpellCastSuccess = True
170                 Call HandleHechizoTerreno(UserIndex, uh)
                Case e_TargetType.uPets
                    SpellCastSuccess = HandlePetSpell(UserIndex, uh)
            End Select
        End If
        If SpellCastSuccess Then
112         If Hechizos(uh).Cooldown > 0 Then
114             UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount()
                If Hechizos(uh).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(uh).CdEffectId, -uh, CLng(Hechizos(uh).Cooldown) * 1000, CLng(Hechizos(uh).Cooldown) * 1000, eCD)
            End If
            If IsSet(UserList(UserIndex).flags.StatusMask, e_StatusMask.eTransformed) Then
                If UserList(UserIndex).Char.CastAnimation > 0 Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageDoAnimation(UserList(UserIndex).Char.charindex, UserList(UserIndex).Char.CastAnimation))
                ElseIf UserList(UserIndex).Char.Ataque1 > 0 Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageDoAnimation(UserList(UserIndex).Char.charindex, UserList(UserIndex).Char.Ataque1))
                End If
            End If
            If Hechizos(uh).TargetEffectType = e_TargetEffectType.eNegative Then
                If IsValidUserRef(UserList(UserIndex).flags.targetUser) Then
                    Call RegisterNewAttack(UserList(UserIndex).flags.targetUser.ArrayIndex, UserIndex)
                    If IsFeatureEnabled("remove-inv-on-attack") Then
                        Call RemoveUserInvisibility(UserIndex)
                    End If
                End If
            ElseIf Hechizos(uh).TargetEffectType = ePositive Then
                If IsValidUserRef(UserList(UserIndex).flags.targetUser) Then Call RegisterNewHelp(UserList(UserIndex).flags.targetUser.ArrayIndex, UserIndex)
            End If
            Call ClearUserRef(UserList(UserIndex).flags.targetUser)
            Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
        End If
172     If UserList(UserIndex).Counters.Trabajando Then
174         Call WriteMacroTrabajoToggle(UserIndex, False)
        End If
176     If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
        Exit Sub
LanzarHechizo_Err:
178     Call TraceError(Err.Number, Err.Description, "modHechizos.LanzarHechizo", Erl)
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

        Dim h As Integer, tU As Integer

100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
102     tU = UserList(UserIndex).flags.targetUser.ArrayIndex

104     If IsSet(Hechizos(h).Effects, e_SpellEffects.Invisibility) Then
   
106         If UserList(tU).flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
108             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
110             b = False

                Exit Sub

            End If
            
112         If UserList(UserIndex).flags.EnReto Then
                'Msg793= No podés lanzar invisibilidad durante un reto.
                Call WriteLocaleMsg(UserIndex, "793", e_FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
124         If UserList(UserIndex).flags.Montado Then
                'Msg794= No podés lanzar invisibilidad mientras usas una montura.
                Call WriteLocaleMsg(UserIndex, "794", e_FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
                        
125         If UserList(tU).flags.Montado Then
                'Msg795= No podés lanzar invisibilidad a alguien montado.
                Call WriteLocaleMsg(UserIndex, "795", e_FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
    
128         If UserList(tU).Counters.Saliendo Then
130             If UserIndex <> tU Then
132                 ' Msg666=¡El hechizo no tiene efecto!
                    Call WriteLocaleMsg(UserIndex, "666", e_FontTypeNames.FONTTYPE_INFO)
134                 b = False

                    Exit Sub

                Else
136                 ' Msg667=¡No podés ponerte invisible mientras te encuentres saliendo!
                    Call WriteLocaleMsg(UserIndex, "667", e_FontTypeNames.FONTTYPE_WARNING)
138                 b = False

                    Exit Sub

                End If
            End If
            
            If Not PeleaSegura(UserIndex, tU) Then

                Select Case Status(UserIndex)

                    Case 1, 3, 5 'Ciudadano o armada

                        If Status(tU) <> e_Facciones.Ciudadano And Status(tU) <> e_Facciones.Armada And Status(tU) <> e_Facciones.consejo Then

                            If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                                ' Msg662=No puedes ayudar criminales.
                                Call WriteLocaleMsg(UserIndex, "662", e_FontTypeNames.FONTTYPE_INFO)
                                b = False

                                Exit Sub

                            ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then

                                If UserList(UserIndex).flags.Seguro = True Then
                                    ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                    Call WriteLocaleMsg(UserIndex, "663", e_FontTypeNames.FONTTYPE_INFO)
                                    b = False

                                    Exit Sub

                                Else

                                    'Si tiene clan
                                    If UserList(UserIndex).GuildIndex > 0 Then

                                        'Si el clan es de alineación ciudadana.
                                        If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                            'No lo dejo resucitarlo
                                            ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                            Call WriteLocaleMsg(UserIndex, "664", e_FontTypeNames.FONTTYPE_INFO)
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
                            'Msg796= No podés ayudar ciudadanos.
                            Call WriteLocaleMsg(UserIndex, "796", e_FontTypeNames.FONTTYPE_INFO)
                            b = False

                            Exit Sub

                        End If

                End Select

            End If
    
            'Si sos user, no uses este hechizo con GMS.
158         If UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then
160             If Not UserList(tU).flags.Privilegios And e_PlayerType.user Then

                    Exit Sub

                End If
            End If
            
162         If MapInfo(UserList(tU).Pos.Map).SinInviOcul Then
                'Msg797= Una fuerza divina te impide usar invisibilidad en esta zona.
                Call WriteLocaleMsg(UserIndex, "797", e_FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
166         If UserList(tU).flags.invisible = 1 Or UserList(tU).Counters.DisabledInvisibility > 0 Then

168             If tU = UserIndex Then
                    'Msg798= ¡Ya estás invisible!
                    Call WriteLocaleMsg(UserIndex, "798", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    'Msg799= ¡El objetivo ya se encuentra invisible!
                    Call WriteLocaleMsg(UserIndex, "799", e_FontTypeNames.FONTTYPE_INFO)
                End If

174             b = False

                Exit Sub

            End If

            If IsSet(UserList(tU).flags.StatusMask, eTaunting) Then
                If tU = UserIndex Then
                    'Msg800= ¡No podes ocultarte en este momento!
                    Call WriteLocaleMsg(UserIndex, "800", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    'Msg801= ¡El objetivo no puede ocultarse!
                    Call WriteLocaleMsg(UserIndex, "801", e_FontTypeNames.FONTTYPE_INFO)
                End If

                b = False

                Exit Sub

            End If
   
176         UserList(tU).flags.invisible = 1

            'Ladder
            'Reseteamos el contador de Invisibilidad
            'Le agrego un random al tiempo de invisibilidad de 16 a 21 segundos.
            If UserList(tU).Counters.Invisibilidad <= 0 Then UserList(tU).Counters.Invisibilidad = RandomNumber(Hechizos(h).Duration - 4, Hechizos(h).Duration + 1)
177         Call WriteContadores(tU)
178         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.charindex, True, UserList(tU).pos.X, UserList(tU).pos.y))
179         Call InfoHechizo(UserIndex)
180         b = True
        End If

181     If Hechizos(h).EotId > 0 Then
182         b = True
183         Call InfoHechizo(UserIndex)

184         Exit Sub

185     End If
        
188     If Hechizos(h).Mimetiza = 1 Then

190         If UserList(UserIndex).flags.EnReto Then
                'Msg802= No podés mimetizarte durante un reto.
                Call WriteLocaleMsg(UserIndex, "802", e_FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If

194         If UserList(tU).flags.Muerto = 1 Then

                Exit Sub

            End If
            
196         If UserList(tU).flags.Navegando = 1 Then

                Exit Sub

            End If

198         If UserList(UserIndex).flags.Navegando = 1 Then

                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
200         If Not EsGM(UserIndex) And EsGM(tU) Then Exit Sub
            
            ' Si te mimetizaste, no importa si como bicho o User...
202         If UserList(UserIndex).flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
                'Msg803= Ya te encuentras transformado. El hechizo no tuvo efecto
                Call WriteLocaleMsg(UserIndex, "803", e_FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
206         If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
            
            'copio el char original al mimetizado
208         With UserList(UserIndex)
210             .CharMimetizado.Body = .Char.Body
212             .CharMimetizado.Head = .Char.Head
214             .CharMimetizado.CascoAnim = .Char.CascoAnim
216             .CharMimetizado.ShieldAnim = .Char.ShieldAnim
218             .CharMimetizado.WeaponAnim = .Char.WeaponAnim
219             .CharMimetizado.CartAnim = .char.CartAnim
                
220             .flags.Mimetizado = e_EstadoMimetismo.FormaUsuario
                
                'ahora pongo local el del enemigo
222             .Char.Body = UserList(tU).Char.Body
224             .Char.Head = UserList(tU).Char.Head
226             .Char.CascoAnim = UserList(tU).Char.CascoAnim
228             .Char.ShieldAnim = UserList(tU).Char.ShieldAnim
230             .Char.WeaponAnim = UserList(tU).Char.WeaponAnim
231             .char.CartAnim = UserList(tU).char.CartAnim
232             .NameMimetizado = UserList(tU).Name

234             If UserList(tU).GuildIndex > 0 Then .NameMimetizado = .NameMimetizado & " <" & modGuilds.GuildName(UserList(tU).GuildIndex) & ">"
            
236             Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
238             Call RefreshCharStatus(UserIndex)
            End With
           
240         Call InfoHechizo(UserIndex)
242         b = True
        End If

244     If Hechizos(h).Envenena > 0 Then
246         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub

248         If UserList(tU).flags.Envenenado = 0 Then
250             If UserIndex <> tU Then
252                 Call UsuarioAtacadoPorUsuario(UserIndex, tU)
                End If

254             UserList(tU).flags.Envenenado = Hechizos(h).Envenena
256             UserList(tU).Counters.Veneno = Hechizos(h).Duration
258             Call InfoHechizo(UserIndex)
260             b = True
            Else
262             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1870, UserList(tU).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1870=¬1 ya está envenenado. El hechizo no tuvo efecto.
264             b = False
            End If
        End If

266     If Hechizos(h).desencantar = 1 Then
            ' Call WriteConsoleMsg(UserIndex, "Ningun efecto magico tiene efecto sobre ti ya.", e_FontTypeNames.FONTTYPE_INFOIAO)

268         UserList(UserIndex).flags.Envenenado = 0
270         UserList(UserIndex).Counters.Veneno = 0
272         UserList(UserIndex).flags.Incinerado = 0
274         UserList(UserIndex).Counters.Incineracion = 0
    
276         If UserList(UserIndex).flags.Inmovilizado > 0 Then
278             UserList(UserIndex).Counters.Inmovilizado = 0
280             UserList(UserIndex).flags.Inmovilizado = 0

                If UserList(UserIndex).clase = e_Class.Warrior Or UserList(UserIndex).clase = e_Class.Hunter Or UserList(UserIndex).clase = e_Class.Thief Or UserList(UserIndex).clase = e_Class.Pirat Then
                    UserList(UserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If

282             Call WriteInmovilizaOK(UserIndex)
            End If

284         If UserList(UserIndex).flags.Paralizado > 0 Then
286             UserList(UserIndex).Counters.Paralisis = 0
288             UserList(UserIndex).flags.Paralizado = 0

                If UserList(UserIndex).clase = e_Class.Warrior Or UserList(UserIndex).clase = e_Class.Hunter Or UserList(UserIndex).clase = e_Class.Thief Or UserList(UserIndex).clase = e_Class.Pirat Then
                    UserList(UserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If

290             Call WriteParalizeOK(UserIndex)
            End If
        
292         If UserList(UserIndex).flags.Ceguera > 0 Then
294             UserList(UserIndex).Counters.Ceguera = 0
296             UserList(UserIndex).flags.Ceguera = 0
298             Call WriteBlindNoMore(UserIndex)
            End If
    
300         If UserList(UserIndex).flags.Maldicion > 0 Then
302             UserList(UserIndex).flags.Maldicion = 0
304             UserList(UserIndex).Counters.Maldicion = 0
            End If

306         If UserList(UserIndex).flags.Estupidez > 0 Then
308             UserList(UserIndex).flags.Estupidez = 0
310             UserList(UserIndex).Counters.Estupidez = 0
            End If
    
312         Call InfoHechizo(UserIndex)
314         b = True
        End If

316     If IsSet(Hechizos(h).Effects, e_SpellEffects.Incinerate) Then
318         If UserIndex = tU Then
320             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
    
322         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
324         If UserIndex <> tU Then
326             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If

328         UserList(tU).flags.Incinerado = 1
330         Call InfoHechizo(UserIndex)
332         b = True
        End If

        If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveDebuff) Then

            Dim NegativeEffect As IBaseEffectOverTime

            Set NegativeEffect = EffectsOverTime.FindEffectOfTypeOnTarget(UserList(tU).EffectOverTime, eDebuff)

            If Not NegativeEffect Is Nothing Then
                NegativeEffect.RemoveMe = True
                Call InfoHechizo(UserIndex)
                b = True

                Exit Sub

            End If
        End If

        If IsSet(Hechizos(h).Effects, e_SpellEffects.StealBuff) Then

            Dim TargetBuff As IBaseEffectOverTime

            Set TargetBuff = EffectsOverTime.FindEffectOfTypeOnTarget(UserList(tU).EffectOverTime, eBuff)

            If Not TargetBuff Is Nothing Then
                Call EffectsOverTime.ChangeOwner(tU, eUser, UserIndex, eUser, TargetBuff)
            End If

            Call InfoHechizo(UserIndex)
            b = True

            Exit Sub

        End If

334     If IsSet(Hechizos(h).Effects, e_SpellEffects.CurePoison) Then

            'Verificamos que el usuario no este muerto
336         If UserList(tU).flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
338             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
340             b = False

                Exit Sub

            End If
            
            ' Si no esta envenenado, no hay nada mas que hacer
342         If UserList(tU).flags.Envenenado = 0 Then
344             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1871, UserList(tU).name, e_FontTypeNames.FONTTYPE_INFOIAO)) ' Msg1871=¬1 no está envenenado, el hechizo no tiene efecto.
346             b = False

                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
348         If Not PeleaSegura(UserIndex, tU) Then
350             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then

352                 If esArmada(UserIndex) Then
354                     Call WriteLocaleMsg(UserIndex, "379", e_FontTypeNames.FONTTYPE_INFO)
356                     b = False

                        Exit Sub

                    End If

358                 If UserList(UserIndex).flags.Seguro Then
360                     Call WriteLocaleMsg(UserIndex, "378", e_FontTypeNames.FONTTYPE_INFO)
362                     b = False

                        Exit Sub

                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
364         If UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then
366             If Not UserList(tU).flags.Privilegios And e_PlayerType.user Then

                    Exit Sub

                End If

            End If
        
368         UserList(tU).flags.Envenenado = 0
370         UserList(tU).Counters.Veneno = 0
372         Call InfoHechizo(UserIndex)
374         b = True

        End If

376     If IsSet(Hechizos(h).Effects, e_SpellEffects.Curse) Then
378         If UserIndex = tU Then
380             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
    
382         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
384         If UserIndex <> tU Then
386             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

388         UserList(tU).flags.Maldicion = 1
390         UserList(tU).Counters.Maldicion = 200
    
392         Call InfoHechizo(UserIndex)
394         b = True

        End If

396     If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveCurse) Then
398         UserList(tU).flags.Maldicion = 0
400         UserList(tU).Counters.Maldicion = 0
402         Call InfoHechizo(UserIndex)
404         b = True
        End If

406     If IsSet(Hechizos(h).Effects, e_SpellEffects.PreciseHit) Then
408         UserList(tU).flags.GolpeCertero = 1
410         Call InfoHechizo(UserIndex)
412         b = True

        End If

422     If IsSet(Hechizos(h).Effects, e_SpellEffects.Paralize) Then
424         If UserIndex = tU Then
426             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
            
            If UserList(tU).Counters.TiempoDeInmunidadParalisisNoMagicas > 0 Then
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1872, UserList(tU).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1872=¬1 no puede volver a ser paralizado tan rápido.

                Exit Sub

            End If

            If Not UserMod.CanMove(UserList(tU).flags, UserList(tU).Counters) Then
428             ' Msg661=No podes inmovilizar un objetivo que no puede moverse.
                Call WriteLocaleMsg(UserIndex, "661", e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If

            If IsSet(UserList(tU).flags.StatusMask, eCCInmunity) Then
                Call WriteLocaleMsg(UserIndex, MsgCCInunity, e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
    
432         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
434         If UserIndex <> tU Then
                Call checkHechizosEfectividad(UserIndex, tU)
436             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
438         Call InfoHechizo(UserIndex)
440         b = True
            
            If UserList(tU).clase = Warrior Or UserList(tU).clase = hunter Then
                UserList(tU).Counters.Paralisis = Hechizos(h).Duration * 0.7
            Else
                UserList(tU).Counters.Paralisis = Hechizos(h).Duration
            End If

444         If UserList(tU).flags.Paralizado = 0 Then
446             UserList(tU).flags.Paralizado = 1
448             Call WriteParalizeOK(tU)
450             Call WritePosUpdate(tU)
            End If

        End If

452     If Hechizos(h).velocidad <> 0 Then

            'Verificamos que el usuario no este muerto
454         If UserList(tU).flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
456             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
458             b = False

                Exit Sub

            End If

460         If Hechizos(h).velocidad < 1 Then
462             If UserIndex = tU Then
464                 Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)

                    Exit Sub

                End If
    
466             If Not PuedeAtacar(UserIndex, tU) Then Exit Sub

            Else

                'Para poder tirar curar veneno a un pk en el ring
468             If Not PeleaSegura(UserIndex, tU) Then
470                 If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then

472                     If esArmada(UserIndex) Then
474                         Call WriteLocaleMsg(UserIndex, "379", e_FontTypeNames.FONTTYPE_INFO)
476                         b = False

                            Exit Sub
    
                        End If
    
478                     If UserList(UserIndex).flags.Seguro Then
480                         Call WriteLocaleMsg(UserIndex, "378", e_FontTypeNames.FONTTYPE_INFO)
482                         b = False

                            Exit Sub
    
                        End If
    
                    End If
    
                End If
            
                'Si sos user, no uses este hechizo con GMS.
484             If UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then
486                 If Not UserList(tU).flags.Privilegios And e_PlayerType.user Then

                        Exit Sub

                    End If
    
                End If
            End If

488         Call UsuarioAtacadoPorUsuario(UserIndex, tU)

490         Call InfoHechizo(UserIndex)
492         b = True
                 
494         If UserList(tU).Counters.velocidad = 0 Then
496             UserList(tU).flags.VelocidadHechizada = Hechizos(h).velocidad
                
498             Call ActualizarVelocidadDeUsuario(tU)
            End If

500         UserList(tU).Counters.velocidad = Hechizos(h).Duration

        End If

502     If IsSet(Hechizos(h).Effects, e_SpellEffects.Immobilize) Then
504         If UserIndex = tU Then
506             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If

            If Not UserMod.CanMove(UserList(tU).flags, UserList(tU).Counters) Then
510             ' Msg661=No podes inmovilizar un objetivo que no puede moverse.
                Call WriteLocaleMsg(UserIndex, "661", e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
            
            If UserList(tU).Counters.TiempoDeInmunidadParalisisNoMagicas > 0 Then
515             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1873, UserList(tU).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1873=¬1 no puede volver a ser inmovilizado tan rápido.

                Exit Sub

            End If
            
            If IsSet(UserList(tU).flags.StatusMask, eCCInmunity) Then
                Call WriteLocaleMsg(UserIndex, MsgCCInunity, e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
    
516         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
518         If UserIndex <> tU Then
                Call checkHechizosEfectividad(UserIndex, tU)
520             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
            
522         Call InfoHechizo(UserIndex)
524         b = True
            
            If UserList(tU).clase = Warrior Or UserList(tU).clase = hunter Then
                UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration * 0.7
            Else
                UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration
            End If

528         UserList(tU).flags.Inmovilizado = 1
            
530         Call WriteInmovilizaOK(tU)
532         Call WritePosUpdate(tU)

        End If

534     If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveParalysis) Then
        
            'Para poder tirar remo a un pk en el ring
536         If Not PeleaSegura(UserIndex, tU) Then

                Select Case Status(UserIndex)

                    Case 1, 3, 5 'Ciudadano o armada

                        If Status(tU) <> e_Facciones.Ciudadano And Status(tU) <> e_Facciones.Armada And Status(tU) <> e_Facciones.consejo Then

                            If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                                ' Msg662=No puedes ayudar criminales.
                                Call WriteLocaleMsg(UserIndex, "662", e_FontTypeNames.FONTTYPE_INFO)
                                b = False

                                Exit Sub

                            ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then

                                If UserList(UserIndex).flags.Seguro = True Then
                                    ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                    Call WriteLocaleMsg(UserIndex, "663", e_FontTypeNames.FONTTYPE_INFO)
                                    b = False

                                    Exit Sub

                                Else

                                    'Si tiene clan
                                    If UserList(UserIndex).GuildIndex > 0 Then

                                        'Si el clan es de alineación ciudadana.
                                        If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                            'No lo dejo resucitarlo
                                            ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                            Call WriteLocaleMsg(UserIndex, "664", e_FontTypeNames.FONTTYPE_INFO)
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
                            'Msg805= No podés ayudar ciudadanos.
                            Call WriteLocaleMsg(UserIndex, "805", e_FontTypeNames.FONTTYPE_INFO)
                            b = False

                            Exit Sub

                        End If

                End Select

            End If
        
554         If UserList(tU).flags.Inmovilizado = 0 And UserList(tU).flags.Paralizado = 0 Then
                'Msg806= El objetivo no esta paralizado.
                Call WriteLocaleMsg(UserIndex, "806", e_FontTypeNames.FONTTYPE_INFO)
558             b = False

                Exit Sub

            End If
        
560         If UserList(tU).flags.Inmovilizado = 1 Then
562             UserList(tU).Counters.Inmovilizado = 0

                If UserList(tU).clase = e_Class.Warrior Or UserList(tU).clase = e_Class.Hunter Or UserList(tU).clase = e_Class.Thief Or UserList(tU).clase = e_Class.Pirat Then
                    UserList(tU).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If

564             UserList(tU).flags.Inmovilizado = 0
566             Call WriteInmovilizaOK(tU)
568             Call WritePosUpdate(tU)
            End If
            
570         If UserList(tU).flags.Paralizado = 1 Then
572             UserList(tU).flags.Paralizado = 0

                If UserList(tU).clase = e_Class.Warrior Or UserList(tU).clase = e_Class.Hunter Or UserList(tU).clase = e_Class.Thief Or UserList(tU).clase = e_Class.Pirat Then
                    UserList(tU).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If

574             UserList(tU).Counters.Paralisis = 0

576             Call WriteParalizeOK(tU)
            End If

578         b = True
580         Call InfoHechizo(UserIndex)

        End If

582     If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveDumb) Then
584         If UserList(tU).flags.Estupidez = 1 Then

                'Para poder tirar remo estu a un pk en el ring
586             If Not PeleaSegura(UserIndex, tU) Then
588                 If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then

590                     If esArmada(UserIndex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", e_FontTypeNames.FONTTYPE_INFO)
592                         Call WriteLocaleMsg(UserIndex, "379", e_FontTypeNames.FONTTYPE_INFO)
594                         b = False

                            Exit Sub

                        End If

596                     If UserList(UserIndex).flags.Seguro Then
598                         Call WriteLocaleMsg(UserIndex, "378", e_FontTypeNames.FONTTYPE_INFO)
600                         b = False

                            Exit Sub

                        Else

                            ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                        End If

                    End If

                End If
    
602             UserList(tU).flags.Estupidez = 0
604             UserList(tU).Counters.Estupidez = 0
606             Call WriteDumbNoMore(tU)

608             Call InfoHechizo(UserIndex)
610             b = True

            End If

        End If

612     If IsSet(Hechizos(h).Effects, e_SpellEffects.Resurrect) Then
614         If UserList(tU).flags.Muerto = 1 Then

616             If UserList(UserIndex).flags.EnReto Then
                    'Msg807= No podés revivir a nadie durante un reto.
                    Call WriteLocaleMsg(UserIndex, "807", e_FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
    
                'No usar resu en mapas con ResuSinEfecto
                'If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
                '   Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", e_FontTypeNames.FONTTYPE_INFO)
                '   b = False
                '   Exit Sub
                ' End If
                
620             If UserList(UserIndex).clase <> Cleric Then

                    Dim PuedeRevivir As Boolean
                    
622                 If UserList(UserIndex).invent.EquippedWeaponObjIndex <> 0 Then
624                     If ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).Revive Then
626                         PuedeRevivir = True
                        End If
                    End If
                    
628                 If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex <> 0 Then
630                     If ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).Revive Then
632                         PuedeRevivir = True
                        End If
                    End If
                    
                    If UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex <> 0 Then
                        If ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).Revive Then
                            PuedeRevivir = True
                        End If
                    End If
                        
634                 If Not PuedeRevivir Then
                        'Msg809= Necesitás un objeto con mayor poder mágico para poder revivir.
                        Call WriteLocaleMsg(UserIndex, "809", e_FontTypeNames.FONTTYPE_INFO)
638                     b = False

                        Exit Sub

                    End If
                End If
                
                If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then
                    'Msg810= Deberás tener la barra de energía llena para poder resucitar.
                    Call WriteLocaleMsg(UserIndex, "810", e_FontTypeNames.FONTTYPE_INFO)
                    b = False

                    Exit Sub

                End If

                'Para poder tirar revivir a un pk en el ring
654             If Not PeleaSegura(UserIndex, tU) Then
                    
                    If UserList(tU).flags.SeguroResu Then
                        ' Msg693=El usuario tiene el seguro de resurrección activado.
                        Call WriteLocaleMsg(UserIndex, "693", e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteConsoleMsg(tU, PrepareMessageLocaleMsg(1874, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1874=¬1 está intentando revivirte. Desactiva el seguro de resurrección para permitirle hacerlo.
                        b = False

                        Exit Sub

                    End If
                
                    Select Case Status(UserIndex)

                        Case 1, 3, 5 'Ciudadano o armada

                            If Status(tU) <> e_Facciones.Ciudadano And Status(tU) <> e_Facciones.Armada And Status(tU) <> e_Facciones.consejo Then

                                If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                                    'Msg811= Los miembros de la armada real solo pueden revivir ciudadanos a miembros de su facción.
                                    Call WriteLocaleMsg(UserIndex, "811", e_FontTypeNames.FONTTYPE_INFO)
                                    b = False

                                    Exit Sub

                                ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then

                                    If UserList(UserIndex).flags.Seguro = True Then
                                        'Msg812= Deberás desactivar el seguro para revivir al usuario, ten en cuenta que te convertirás en criminal.
                                        Call WriteLocaleMsg(UserIndex, "812", e_FontTypeNames.FONTTYPE_INFO)
                                        b = False

                                        Exit Sub

                                    Else

                                        'Si tiene clan
                                        If UserList(UserIndex).GuildIndex > 0 Then

                                            'Si el clan es de alineación ciudadana.
                                            If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                                'No lo dejo resucitarlo
                                                'Msg813= No puedes resucitar al usuario siendo fundador de un clan ciudadano.
                                                Call WriteLocaleMsg(UserIndex, "813", e_FontTypeNames.FONTTYPE_INFO)
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
                                'Msg814= Los miembros del caos solo pueden revivir criminales o miembros de su facción.
                                Call WriteLocaleMsg(UserIndex, "814", e_FontTypeNames.FONTTYPE_INFO)
                                b = False

                                Exit Sub

                            End If

                    End Select

                End If
                
                Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.charindex, False, UserList(tU).Pos.X, UserList(tU).Pos.Y))

684             Call ResurrectUser(tU)
686             Call InfoHechizo(UserIndex)

688             b = True
            Else
690             b = False

            End If

        End If

692     If IsSet(Hechizos(h).Effects, e_SpellEffects.Blindness) Then
694         If UserIndex = tU Then
696             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If
    
698         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
700         If UserIndex <> tU Then
702             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If

704         UserList(tU).flags.Ceguera = 1
706         UserList(tU).Counters.Ceguera = Hechizos(h).Duration
708         Call WriteBlind(tU)
710         Call InfoHechizo(UserIndex)
712         b = True
        End If

714     If IsSet(Hechizos(h).Effects, e_SpellEffects.Dumb) Then
716         If UserIndex = tU Then
718             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If

720         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
722         If UserIndex <> tU Then
724             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If

726         If UserList(tU).flags.Estupidez = 0 Then
728             UserList(tU).flags.Estupidez = 1
730             UserList(tU).Counters.Estupidez = Hechizos(h).Duration
            End If

732         Call WriteDumb(tU)
734         Call InfoHechizo(UserIndex)
736         b = True
        End If

738     If IsSet(Hechizos(h).Effects, e_SpellEffects.ToggleCleave) Then
            If UserList(UserIndex).flags.Cleave Then
740             UserList(UserIndex).flags.Cleave = 0

                If Hechizos(h).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(h).CdEffectId, -1, 0, 0, eBuff)
            Else
742             UserList(UserIndex).flags.Cleave = 1

                If Hechizos(h).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(h).CdEffectId, -1, -1, -1, eBuff)
            End If

744         b = True
        End If
        
        Dim Character As t_User

        Character = UserList(UserIndex)
        
        If IsSet(Hechizos(h).Effects, e_SpellEffects.ToggleDivineBlood) Then
            If UserList(UserIndex).flags.DivineBlood Then
                UserList(UserIndex).flags.DivineBlood = 0
                If Hechizos(h).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(h).CdEffectId, -1, 0, 0, eBuff)
                Character.Char.BackpackAnim = 0
                Call WriteCharacterChange(UserIndex, Character.Char.body, Character.Char.head, Character.Char.Heading, Character.Char.charindex, Character.Char.WeaponAnim, Character.Char.ShieldAnim, 0, Character.Char.BackpackAnim, 0, 0, Character.Char.CascoAnim, False, False)
            Else
                UserList(UserIndex).flags.DivineBlood = 1
                If Hechizos(h).CdEffectId > 0 Then Call WriteSendSkillCdUpdate(UserIndex, Hechizos(h).CdEffectId, -1, -1, -1, eBuff)
                Character.Char.BackpackAnim = 4997
                Call WriteCharacterChange(UserIndex, Character.Char.body, Character.Char.head, Character.Char.Heading, Character.Char.charindex, Character.Char.WeaponAnim, Character.Char.ShieldAnim, 0, Character.Char.BackpackAnim, 0, 0, Character.Char.CascoAnim, False, False)
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
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageLocaleMsg(1638, .name & "¬" & efectividad & "¬" & .Counters.controlHechizos.HechizosCasteados & "¬" & .Counters.controlHechizos.HechizosTotales, e_FontTypeNames.FONTTYPE_TALK)) 'Msg1638=El usuario ¬1 está lanzando hechizos con una efectividad de ¬2% (Casteados: ¬3/¬4), revisar.
            End If
            
            Debug.Print "El usuario " & .name & " está lanzando hechizos con una efectividad de " & efectividad & "% (Casteados: " & .Counters.controlHechizos.HechizosCasteados & "/" & .Counters.controlHechizos.HechizosTotales & "), revisar."
        Else
            .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales - 1
        End If
    End With
End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
        On Error GoTo HechizoEstadoNPC_Err
        
        Dim UserAttackInteractionResult As t_AttackInteractionResult
        
100     If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.Invisibility) Then
102         Call InfoHechizo(UserIndex)
104         NpcList(NpcIndex).flags.invisible = 1
106         b = True
        End If

108     If Hechizos(hIndex).Envenena > 0 Then

            UserAttackInteractionResult = UserCanAttackNpc(UserIndex, NpcIndex)
            Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
            If UserAttackInteractionResult.CanAttack Then
                If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
            Else
                b = False
                Exit Sub
            End If
114         Call NPCAtacado(NpcIndex, UserIndex)
116         Call InfoHechizo(UserIndex)
118         NpcList(NpcIndex).flags.Envenenado = Hechizos(hIndex).Envenena
120         b = True
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

122     If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.CurePoison) Then
124         If NpcList(NpcIndex).flags.Envenenado > 0 Then
126             Call InfoHechizo(UserIndex)
128             NpcList(NpcIndex).flags.Envenenado = 0
130             b = True
            Else
'Msg815= La criatura no esta envenenada, el hechizo no tiene efecto.
Call WriteLocaleMsg(UserIndex, "815", e_FontTypeNames.FONTTYPE_INFOIAO)
134             b = False
            End If
        End If

136     If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.RemoveCurse) Then
138         Call InfoHechizo(UserIndex)
140         b = True
        End If

150     If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.Paralize) Then
152         If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
158             Call NPCAtacado(NpcIndex, UserIndex, True)
160             Call InfoHechizo(UserIndex)
162             NpcList(NpcIndex).flags.Paralizado = 1
164             NpcList(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 6
166             NpcList(NpcIndex).flags.Inmovilizado = 0
168             NpcList(NpcIndex).Contadores.Inmovilizado = 0
170             Call AnimacionIdle(NpcIndex, False)
172             b = True
            Else
174             Call WriteLocaleMsg(UserIndex, "381", e_FontTypeNames.FONTTYPE_INFOIAO)
176             b = False
                Exit Sub
            End If
        End If

178     If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.RemoveParalysis) Then
180         With NpcList(NpcIndex)
182             If .flags.Paralizado + .flags.Inmovilizado = 0 Then
'Msg816= Este NPC no esta Paralizado
Call WriteLocaleMsg(UserIndex, "816", e_FontTypeNames.FONTTYPE_INFOIAO)
186                 b = False
                Else
                    Dim IsValidMaster As Boolean
                    IsValidMaster = IsValidUserRef(.MaestroUser)
                    ' Si el usuario es Armada o Caos y el NPC es de la misma faccion
188                 b = ((esArmada(UserIndex) Or esCaos(UserIndex)) And .flags.Faccion = UserList(UserIndex).Faccion.Status)
                    'O si es mi propia mascota
190                 b = b Or (IsValidMaster And (.MaestroUser.ArrayIndex = userIndex))
                    'O si es mascota de otro usuario de la misma faccion
192                 b = b Or ((esArmada(userIndex) And (IsValidMaster And esArmada(.MaestroUser.ArrayIndex))) Or (esCaos(userIndex) And (IsValidMaster And esCaos(.MaestroUser.ArrayIndex))))
                    
194                 If b Then
196                     Call InfoHechizo(UserIndex)
198                     .flags.Paralizado = 0
200                     .Contadores.Paralisis = 0
202                     .flags.Inmovilizado = 0
204                     .Contadores.Inmovilizado = 0
                    Else
'Msg817= Solo podés remover la Parálisis de tus mascotas o de criaturas que pertenecen a tu facción.
Call WriteLocaleMsg(UserIndex, "817", e_FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                End If
            End With
        End If
 
208     If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.Immobilize) Then
210         If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
220             Call NPCAtacado(NpcIndex, UserIndex, True)
222             NpcList(NpcIndex).flags.Inmovilizado = 1
224             NpcList(NpcIndex).Contadores.Inmovilizado = (Hechizos(hIndex).Duration * 6.5) * 6
226             NpcList(NpcIndex).flags.Paralizado = 0
228             NpcList(NpcIndex).Contadores.Paralisis = 0
230             Call AnimacionIdle(NpcIndex, True)
232             Call InfoHechizo(UserIndex)
234             b = True
            Else
236             Call WriteLocaleMsg(UserIndex, "381", e_FontTypeNames.FONTTYPE_INFOIAO)
            End If
        End If
        
238     If Hechizos(hIndex).Mimetiza = 1 Then
240         If UserList(UserIndex).flags.EnReto Then
'Msg818= No podés mimetizarte durante un reto.
Call WriteLocaleMsg(UserIndex, "818", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            End If
    
244         If UserList(UserIndex).flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
'Msg819= Ya te encuentras transformado. El hechizo no tuvo efecto
Call WriteLocaleMsg(UserIndex, "819", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
248         If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
 
250         If UserList(UserIndex).clase = e_Class.Druid Then

                'copio el char original al mimetizado
252             With UserList(UserIndex)
254                 .CharMimetizado.Body = .Char.Body
256                 .CharMimetizado.Head = .Char.Head
258                 .CharMimetizado.CascoAnim = .Char.CascoAnim
260                 .CharMimetizado.ShieldAnim = .Char.ShieldAnim
262                 .CharMimetizado.WeaponAnim = .Char.WeaponAnim
261                 .CharMimetizado.CartAnim = .char.CartAnim
264                 .flags.Mimetizado = e_EstadoMimetismo.FormaBicho
                    
                    'ahora pongo lo del NPC.
266                 .Char.Body = NpcList(NpcIndex).Char.Body
268                 .Char.Head = NpcList(NpcIndex).Char.Head
270                 Call ClearClothes(.char)
276                 .NameMimetizado = IIf(NpcList(NpcIndex).showName = 1, NpcList(NpcIndex).Name, vbNullString)

278                 Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
280                 Call RefreshCharStatus(UserIndex)
                End With
                
            Else
            
'Msg820= Solo los druidas pueden mimetizarse con criaturas.
Call WriteLocaleMsg(UserIndex, "820", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
                
            End If
        
284        Call InfoHechizo(UserIndex)
286        b = True
        End If
        If Hechizos(hIndex).EotId Then
294        Call InfoHechizo(UserIndex)
296        b = True
        End If
        Exit Sub

HechizoEstadoNPC_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoEstadoNPC", Erl)
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal npcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean, ByRef IsAlive As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 14/08/2007
        'Handles the Spells that afect the Life NPC
        '14/08/2007 Pablo (ToxicWaste) - Orden general.
        '***************************************************
        
        On Error GoTo HechizoPropNPC_Err
        
        Dim UserAttackInteractionResult As t_AttackInteractionResult
        Dim Damage As Long
        
        Dim DamageStr As String
    
        'Salud
100     If IsSet(Hechizos(hIndex).Effects, e_SpellEffects.eDoHeal) Then
102         If NpcList(NpcIndex).Stats.MinHp < NpcList(NpcIndex).Stats.MaxHp Then
104             Damage = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
105             Damage = Damage * UserMod.GetMagicHealingBonus(UserList(UserIndex))
                Damage = Damage * NPCs.GetSelfHealingBonus(NpcList(NpcIndex))

                If IsFeatureEnabled("elemental_tags") Then
                    Call CalculateElementalTagsModifiers(UserIndex, NpcIndex, Damage)
                End If
                
106             Call InfoHechizo(UserIndex)
108             Call NPCs.DoDamageOrHeal(npcIndex, UserIndex, eUser, Damage, e_DamageSourceType.e_magic, hIndex)
                
                If Damage > 0 Then
112                 DamageStr = PonerPuntos(Damage)
114                 Call WriteLocaleMsg(UserIndex, 388, e_FontTypeNames.FONTTYPE_FIGHT, "la criatura¬" & DamageStr)
                End If
120             b = True
            Else
'Msg821= La criatura no tiene heridas que curar, el hechizo no tiene efecto.
Call WriteLocaleMsg(UserIndex, "821", e_FontTypeNames.FONTTYPE_INFOIAO)
124             b = False
            End If
        
126     ElseIf IsSet(Hechizos(hIndex).Effects, e_SpellEffects.eDoDamage) Then

            UserAttackInteractionResult = UserCanAttackNpc(UserIndex, NpcIndex)
            Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
            If UserAttackInteractionResult.CanAttack Then
                If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
            Else
                b = False
                Exit Sub
            End If
                    
132         Call NPCAtacado(NpcIndex, UserIndex)
134         Damage = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
136         Damage = Damage + Porcentaje(Damage, 3 * UserList(UserIndex).Stats.ELV)
            Dim MagicPenetration As Integer
            
148         If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
150             Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
                MagicPenetration = ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicPenetration
                Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicAbsoluteBonus
            End If
151         If UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex > 0 Then
                Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicDamageBonus)
                Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicAbsoluteBonus
                MagicPenetration = MagicPenetration + ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicPenetration
            End If
            ' Magic Damage ring
152         If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
154             Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
                Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicAbsoluteBonus
                MagicPenetration = MagicPenetration + ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicPenetration
            End If
156         b = True
158         If NpcList(NpcIndex).flags.Snd2 > 0 Then
160             Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd2, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.y))
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
162         If Hechizos(hIndex).AntiRm = 0 Then
164             Damage = Damage - NpcList(npcIndex).Stats.defM
            End If
            Damage = Damage * UserMod.GetMagicDamageModifier(UserList(UserIndex))
            Damage = Damage * NPCs.GetMagicDamageReduction(NpcList(NpcIndex))
166         If Damage < 0 Then Damage = 0
            If IsFeatureEnabled("elemental_tags") Then
                Call CalculateElementalTagsModifiers(UserIndex, NpcIndex, Damage)
            End If
170         Call InfoHechizo(UserIndex)
            IsAlive = NPCs.DoDamageOrHeal(NpcIndex, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, hIndex) = eStillAlive
176         If NpcList(NpcIndex).NPCtype = DummyTarget Then
178             Call DummyTargetAttacked(NpcIndex)
            End If
        End If
        Exit Sub
HechizoPropNPC_Err:
200     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoPropNPC", Erl)
End Sub

Private Sub InfoHechizoDeNpcSobreUser(ByVal NpcIndex As Integer, ByVal TargetUser As Integer, ByVal Spell As Integer)
      On Error GoTo InfoHechizoDeNpcSobreUser_Err

100   With UserList(TargetUser)
102     If Hechizos(Spell).FXgrh > 0 Then '¿Envio FX?
104       If Hechizos(Spell).ParticleViaje > 0 Then
            .Counters.timeFx = 3
106         Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageParticleFXWithDestino(NpcList(NpcIndex).Char.charindex, .Char.charindex, Hechizos(Spell).ParticleViaje, Hechizos(Spell).FXgrh, Hechizos(Spell).TimeParticula, Hechizos(Spell).wav, 1, UserList(TargetUser).Pos.X, UserList(TargetUser).Pos.Y))
          Else
            .Counters.timeFx = 3
108         Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops, UserList(TargetUser).Pos.X, UserList(TargetUser).Pos.Y))
          End If
        End If

110     If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
112       If Hechizos(Spell).ParticleViaje > 0 Then
            .Counters.timeFx = 3
114         Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageParticleFXWithDestino(NpcList(NpcIndex).Char.charindex, .Char.charindex, Hechizos(Spell).ParticleViaje, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, Hechizos(Spell).wav, 0, UserList(TargetUser).Pos.X, UserList(TargetUser).Pos.Y))
          Else
            .Counters.timeFx = 3
116         Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageParticleFX(.Char.charindex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False, , UserList(TargetUser).Pos.X, UserList(TargetUser).Pos.Y))
          End If
        End If

118     If Hechizos(Spell).wav > 0 Then
120       Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
        End If

122     If Hechizos(Spell).TimeEfect <> 0 Then
124       Call WriteFlashScreen(TargetUser, Hechizos(Spell).ScreenColor, Hechizos(Spell).TimeEfect)
        End If

      End With
  
      Exit Sub

InfoHechizoDeNpcSobreUser_Err:
126   Call TraceError(Err.Number, Err.Description, "modHechizos.InfoHechizoDeNpcSobreUser", Erl)

End Sub

Private Sub InfoHechizo(ByVal UserIndex As Integer)
        
        On Error GoTo InfoHechizo_Err
        

        Dim h As Integer

100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
102     If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
104         Call DecirPalabrasMagicas(h, UserIndex)

        End If

106     If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then '¿El Hechizo fue tirado sobre un usuario?
108         If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
110             If Hechizos(h).ParticleViaje > 0 Then
                    UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).Counters.timeFx = 3
112                 Call SendData(SendTarget.ToPCAliveArea, UserList(UserIndex).flags.targetUser.ArrayIndex, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.charindex, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).Char.charindex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.x, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.y))
                Else
                    UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).Counters.timeFx = 3
114                 Call SendData(SendTarget.ToPCAliveArea, UserList(UserIndex).flags.targetUser.ArrayIndex, PrepareMessageCreateFX(UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).Char.charindex, Hechizos(h).FXgrh, Hechizos(h).loops, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.x, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.y))
                End If

            End If

116         If Hechizos(h).Particle > 0 Then '¿Envio Particula?
118             If Hechizos(h).ParticleViaje > 0 Then
                    UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).Counters.timeFx = 3
120                 Call SendData(SendTarget.ToPCAliveArea, UserList(UserIndex).flags.targetUser.ArrayIndex, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.charindex, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).Char.charindex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.x, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.y))
                Else
                    UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).Counters.timeFx = 3
122                 Call SendData(SendTarget.ToPCAliveArea, UserList(UserIndex).flags.targetUser.ArrayIndex, PrepareMessageParticleFX(UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).Char.charindex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False, , UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.x, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.y))
                End If

            End If
        
124         If Hechizos(h).ParticleViaje = 0 Then
126             Call SendData(SendTarget.ToPCAliveArea, UserList(UserIndex).flags.targetUser.ArrayIndex, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.x, UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).pos.y))

            End If
        
128         If Hechizos(h).TimeEfect <> 0 Then 'Envio efecto de screen
130             Call WriteFlashScreen(UserIndex, Hechizos(h).ScreenColor, Hechizos(h).TimeEfect)

            End If

132     ElseIf IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then '¿El Hechizo fue tirado sobre un npc?

134         If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
136             If NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Stats.MinHp < 1 Then

                    'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
138                 If Hechizos(h).ParticleViaje > 0 Then
140                     Call SendData(SendTarget.ToNPCAliveArea, UserList(UserIndex).flags.TargetNPC.ArrayIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.charindex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1, UserList(UserIndex).flags.targetX, UserList(UserIndex).flags.targetY))
                    Else
142                     Call SendData(SendTarget.ToNPCAliveArea, UserList(UserIndex).flags.TargetNPC.ArrayIndex, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.targetX, UserList(UserIndex).flags.targetY))
                    End If
                Else
144                 If Hechizos(h).ParticleViaje > 0 Then
146                     Call SendData(SendTarget.ToNPCAliveArea, UserList(UserIndex).flags.TargetNPC.ArrayIndex, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.charindex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                    Else
148                     Call SendData(SendTarget.ToNPCAliveArea, UserList(UserIndex).flags.TargetNPC.ArrayIndex, PrepareMessageCreateFX(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, Hechizos(h).FXgrh, Hechizos(h).loops))
                    End If
                End If

            End If
        
150         If Hechizos(h).Particle > 0 Then '¿Envio Particula?
152             If NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Stats.MinHp < 1 Then
154                 Call SendData(SendTarget.ToNPCAliveArea, UserList(UserIndex).flags.TargetNPC.ArrayIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.charindex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Pos.X, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Pos.y))
                Else
156                 If Hechizos(h).ParticleViaje > 0 Then
158                     Call SendData(SendTarget.ToNPCAliveArea, UserList(UserIndex).flags.TargetNPC.ArrayIndex, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.charindex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                    Else
160                     Call SendData(SendTarget.ToNPCAliveArea, UserList(UserIndex).flags.TargetNPC.ArrayIndex, PrepareMessageParticleFX(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))
                    End If
                End If
            End If

162         If Hechizos(h).ParticleViaje = 0 Then
164             Call SendData(SendTarget.ToNPCAliveArea, UserList(UserIndex).flags.TargetNPC.ArrayIndex, PrepareMessagePlayWave(Hechizos(h).wav, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Pos.X, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Pos.y))
            End If
        Else ' Entonces debe ser sobre el terreno
166         If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
168             Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
            End If
        
170         If Hechizos(h).Particle > 0 Then 'Envio Particula?
172             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
            End If
        
174         If Hechizos(h).wav <> 0 Then
176             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))   'Esta linea faltaba. Pablo (ToxicWaste)
            End If
    
        End If
    
178     If UserList(UserIndex).ChatCombate = 1 Then
180         If Hechizos(h).Target = e_TargetType.uTerreno Then
182             Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, e_FontTypeNames.FONTTYPE_FIGHT)
            
184         ElseIf IsValidUserRef(UserList(UserIndex).flags.targetUser) Then
                'Optimizacion de protocolo por Ladder
186             If UserIndex <> UserList(UserIndex).flags.targetUser.ArrayIndex Then
188                 Call WriteConsoleMsg(UserIndex, "HecMSGU*" & h & "*" & UserList(UserList(UserIndex).flags.targetUser.ArrayIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)
190                 Call WriteConsoleMsg(UserList(UserIndex).flags.targetUser.ArrayIndex, "HecMSGA*" & h & "*" & UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)
                Else
192                 Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, e_FontTypeNames.FONTTYPE_FIGHT)
                End If

194         ElseIf UserList(UserIndex).flags.TargetNPC.ArrayIndex > 0 Then
196             Call WriteConsoleMsg(UserIndex, "HecMSG*" & h, e_FontTypeNames.FONTTYPE_FIGHT)
            Else
198             Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, e_FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
        Exit Sub

InfoHechizo_Err:
200     Call TraceError(Err.Number, Err.Description, "modHechizos.InfoHechizo", Erl)

        
End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean, ByRef IsAlive As Boolean)
        On Error GoTo HechizoPropUsuario_Err
        

        Dim h As Integer
        Dim Damage As Integer
        Dim DamageStr As String
        Dim tempChr As Integer
    
100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
102     tempChr = UserList(UserIndex).flags.targetUser.ArrayIndex
      
        'Hambre
104     If Hechizos(h).SubeHam = 1 Then
    
106         Call InfoHechizo(UserIndex)
    
108         Damage = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
110         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + Damage
112         If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
114         If UserIndex <> tempChr Then
116             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1875, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1875=Le has restaurado ¬1 puntos de hambre a ¬2.
118             Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1895, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1895=¬1 te ha restaurado ¬2 puntos de hambre.
            Else
120             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1896, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1896=Te has restaurado ¬1 puntos de hambre.
            End If
122         Call WriteUpdateHungerAndThirst(tempChr)
124         b = True
126     ElseIf Hechizos(h).SubeHam = 2 Then

128         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
130         If UserIndex <> tempChr Then
132             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
            Else
                Exit Sub
            End If
134         Call InfoHechizo(UserIndex)
136         Damage = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
138         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - Damage
140         If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
142         If UserIndex <> tempChr Then
144             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1897, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1897=Le has quitado ¬1 puntos de hambre a ¬2.
146             Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1898, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1898=¬1 te ha quitado ¬2 puntos de hambre.
            Else
148             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1899, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1899=Te has quitado ¬1 puntos de hambre.

            End If
    
150         Call WriteUpdateHungerAndThirst(tempChr)
    
152         b = True
    
154         If UserList(tempChr).Stats.MinHam < 1 Then
156             UserList(tempChr).Stats.MinHam = 0

            End If
    
        End If

        'Sed
160     If Hechizos(h).SubeSed = 1 Then
    
162         Call InfoHechizo(UserIndex)
    
164         Damage = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
166         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + Damage

168         If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
170         If UserIndex <> tempChr Then
172             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1900, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1900=Le has restaurado ¬1 puntos de sed a ¬2.
174             Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1901, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1901=¬1 te ha restaurado ¬2 puntos de sed.
            Else
176             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1902, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1902=Te has restaurado ¬1 puntos de sed.
            End If
            
178         Call WriteUpdateHungerAndThirst(tempChr)
180         b = True
    
182     ElseIf Hechizos(h).SubeSed = 2 Then
    
184         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
186         If UserIndex <> tempChr Then
188             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
            End If
190         Call InfoHechizo(UserIndex)
192         Damage = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
194         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - Damage
    
196         If UserIndex <> tempChr Then
198             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1903, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1903=Le has quitado ¬1 puntos de sed a ¬2.
200             Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1904, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1904=¬1 te ha quitado ¬2 puntos de sed.
            Else
202             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1905, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1905=Te has quitado ¬1 puntos de sed.
            End If
    
204         If UserList(tempChr).Stats.MinAGU < 1 Then
206             UserList(tempChr).Stats.MinAGU = 0
            End If
            
210         Call WriteUpdateHungerAndThirst(tempChr)
212         b = True
        End If

        ' <-------- Agilidad ---------->
214     If Hechizos(h).SubeAgilidad = 1 Then

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
                                    Call WriteLocaleMsg(UserIndex, "662", e_FontTypeNames.FONTTYPE_INFO)
                                    b = False
                                    Exit Sub
                                ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then
                                    If UserList(UserIndex).flags.Seguro = True Then
                                        ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                        Call WriteLocaleMsg(UserIndex, "663", e_FontTypeNames.FONTTYPE_INFO)
                                        b = False
                                        Exit Sub
                                    Else
                                        'Si tiene clan
                                        If UserList(UserIndex).GuildIndex > 0 Then
                                            'Si el clan es de alineación ciudadana.
                                            If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                                'No lo dejo resucitarlo
                                                ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                                Call WriteLocaleMsg(UserIndex, "664", e_FontTypeNames.FONTTYPE_INFO)
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
Call WriteLocaleMsg(UserIndex, "822", e_FontTypeNames.FONTTYPE_INFO)
                                b = False
                                Exit Sub
                            End If
                    End Select
            End If
    
232         Call InfoHechizo(UserIndex)
234         Damage = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
236         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

238         UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) + Damage, UserList(tempChr).Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)

240         UserList(tempChr).flags.TomoPocion = True
242         b = True
244         Call WriteFYA(tempChr)
    
246     ElseIf Hechizos(h).SubeAgilidad = 2 Then
            'Verifica que el usuario no este muerto
            If UserList(tempChr).flags.Muerto = 1 Then
                b = False
                Exit Sub
            End If
    
248         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
250         If UserIndex <> tempChr Then
252             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
            End If
254         Call InfoHechizo(UserIndex)
256         UserList(tempChr).flags.TomoPocion = True
258         Damage = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
260         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

262         If UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) - Damage < MINATRIBUTOS Then
264             UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) = MINATRIBUTOS
            Else
266             UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(e_Atributos.Agilidad) - Damage

            End If
    
268         b = True
270         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
272     If Hechizos(h).SubeFuerza = 1 Then

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
                                Call WriteLocaleMsg(UserIndex, "662", e_FontTypeNames.FONTTYPE_INFO)
                                b = False
                                Exit Sub
                            ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then
                                If UserList(UserIndex).flags.Seguro = True Then
                                    ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                    Call WriteLocaleMsg(UserIndex, "663", e_FontTypeNames.FONTTYPE_INFO)
                                    b = False
                                    Exit Sub
                                Else
                                    'Si tiene clan
                                    If UserList(UserIndex).GuildIndex > 0 Then
                                        'Si el clan es de alineación ciudadana.
                                        If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                            'No lo dejo resucitarlo
                                            ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                            Call WriteLocaleMsg(UserIndex, "664", e_FontTypeNames.FONTTYPE_INFO)
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
                            Call WriteLocaleMsg(UserIndex, "665", e_FontTypeNames.FONTTYPE_INFO)
                            b = False
                            Exit Sub
                        End If
                End Select
            End If
    
290         Damage = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
292         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
294         UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) + Damage, UserList(tempChr).Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)
296         UserList(tempChr).flags.TomoPocion = True
298         Call WriteFYA(tempChr)
300         b = True
    
302         Call InfoHechizo(UserIndex)
304         Call WriteFYA(tempChr)

306     ElseIf Hechizos(h).SubeFuerza = 2 Then
            'Verifica que el usuario no este muerto
            If UserList(tempChr).flags.Muerto = 1 Then
                b = False
                Exit Sub
            End If
    
308         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
310         If UserIndex <> tempChr Then
312             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
314         UserList(tempChr).flags.TomoPocion = True
    
316         Damage = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
318         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

320         If UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) - Damage < MINATRIBUTOS Then
322             UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) = MINATRIBUTOS
            Else
324             UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(e_Atributos.Fuerza) - Damage

            End If

326         b = True
328         Call InfoHechizo(UserIndex)
330         Call WriteFYA(tempChr)

        End If

        'Salud
332     If IsSet(Hechizos(h).Effects, e_SpellEffects.eDoHeal) Then
    
            'Verifica que el usuario no este muerto
334         If UserList(tempChr).flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
336             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
338             b = False
                Exit Sub
            End If
            
340         If UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp Then
342             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1906, UserList(tempChr).name, e_FontTypeNames.FONTTYPE_INFOIAO)) ' Msg1906=¬1 no tiene heridas para curar.
344             b = False
                Exit Sub
            End If
    
            'Para poder tirar curar a un pk en el ring
348         If Not PeleaSegura(UserIndex, tempChr) Then
350             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
352                 If esArmada(UserIndex) Then
354                     Call WriteLocaleMsg(UserIndex, "379", e_FontTypeNames.FONTTYPE_INFO)
356                     b = False
                        Exit Sub

                    End If

358                 If UserList(UserIndex).flags.Seguro Then
360                     Call WriteLocaleMsg(UserIndex, "378", e_FontTypeNames.FONTTYPE_INFO)
362                     b = False
                        Exit Sub
                    End If
                End If
            Dim trigger As e_Trigger6
            trigger = TriggerZonaPelea(UserIndex, tempChr)
            ' Están en zona segura en un ring e intenta curarse desde afuera hacia adentro o viceversa
364         ElseIf trigger = TRIGGER6_PROHIBE And MapInfo(UserList(UserIndex).Pos.Map).Seguro <> 0 Then
366             b = False
                Exit Sub
            End If
368         Damage = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            Damage = Damage * UserMod.GetMagicHealingBonus(UserList(UserIndex))
            Damage = Damage * UserMod.GetSelfHealingBonus(UserList(tempChr))
            
            If UserList(UserIndex).flags.DivineBlood > 0 Then
                Damage = Damage * DivineBloodHealingMultiplierBonus
            End If
            
            
370         Call InfoHechizo(UserIndex)
            Call UserMod.DoDamageOrHeal(tempChr, UserIndex, eUser, Damage, e_DamageSourceType.e_magic, h)
376         DamageStr = PonerPuntos(Damage)
378         If UserIndex <> tempChr Then
380             Call WriteLocaleMsg(UserIndex, "388", e_FontTypeNames.FONTTYPE_FIGHT, UserList(tempChr).name & "¬" & DamageStr)
382             Call WriteLocaleMsg(tempChr, "32", e_FontTypeNames.FONTTYPE_FIGHT, UserList(UserIndex).name & "¬" & DamageStr)
            Else
384             Call WriteLocaleMsg(UserIndex, "33", e_FontTypeNames.FONTTYPE_FIGHT, DamageStr)
            End If
390         b = True

392     ElseIf IsSet(Hechizos(h).Effects, e_SpellEffects.eDoDamage) Then
    
394         If UserIndex = tempChr Then
396             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If

398         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
400         Damage = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
402         Damage = Damage + Porcentaje(Damage, 3 * UserList(UserIndex).Stats.ELV)
            ' Si al hechizo le afecta el daño mágico
            Dim PorcentajeRM As Integer
            If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
                Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
                PorcentajeRM = PorcentajeRM - ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicPenetration
                Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicAbsoluteBonus
            End If
410         If UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex > 0 Then
412             Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicDamageBonus)
                Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicAbsoluteBonus
                PorcentajeRM = PorcentajeRM - ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicPenetration
            End If
418         If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
420             Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
                Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicAbsoluteBonus
                PorcentajeRM = PorcentajeRM - ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicPenetration
            End If

            ' Si el hechizo no ignora la RM
422         If Hechizos(h).AntiRm = 0 Then
                
                PorcentajeRM = max(0, PorcentajeRM + GetUserMR(tempChr))
                ' Resto el porcentaje total
442             Damage = Damage - Porcentaje(Damage, PorcentajeRM)
            End If
            Call EffectsOverTime.TartgetWillAtack(UserList(UserIndex).EffectOverTime, tempChr, eUser, e_DamageSourceType.e_magic)
            Damage = Damage * UserMod.GetMagicDamageModifier(UserList(UserIndex))
            Damage = Damage * UserMod.GetMagicDamageReduction(UserList(tempChr))
            ' Prevengo daño negativo
444         If Damage < 0 Then Damage = 0
    
446         If UserIndex <> tempChr Then
                Call checkHechizosEfectividad(UserIndex, tempChr)
448             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
            End If
    
450         Call InfoHechizo(UserIndex)
452         IsAlive = UserMod.DoDamageOrHeal(tempChr, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, h) = eStillAlive
453         Call EffectsOverTime.TartgetDidHit(UserList(UserIndex).EffectOverTime, tempChr, eUser, e_DamageSourceType.e_magic)
460         Call SubirSkill(tempChr, Resistencia)
474         b = True

        End If

        'Mana
476     If Hechizos(h).SubeMana = 1 Then
    
478         Call InfoHechizo(UserIndex)
480         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + Damage

482         If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
484         If UserIndex <> tempChr Then
486             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1907, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1907=Le has restaurado ¬1 puntos de mana a ¬2.
488             Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1908, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1908=¬1 te ha restaurado ¬2 puntos de mana.
            Else
490             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1909, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1909=Te has restaurado ¬1 puntos de mana.

            End If
            
492         Call WriteUpdateMana(tempChr)
    
494         b = True
    
496     ElseIf Hechizos(h).SubeMana = 2 Then

498         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
500         If UserIndex <> tempChr Then
502             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
504         Call InfoHechizo(UserIndex)
    
506         If UserIndex <> tempChr Then
508             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1910, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1910=Le has quitado ¬1 puntos de mana a ¬2.
510             Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1911, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1911=¬1 te ha quitado ¬2 puntos de mana.
            Else
512             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1912, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1912=Te has quitado ¬1 puntos de mana.
            End If
    
514         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - Damage
516         If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
518         Call WriteUpdateMana(tempChr)
520         b = True
        End If

        'Stamina
522     If Hechizos(h).SubeSta = 1 Then
524         Call InfoHechizo(UserIndex)
526         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + Damage

528         If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta

530         If UserIndex <> tempChr Then
532             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1913, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1913=Le has restaurado ¬1 puntos de vitalidad a ¬2.
534             Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1914, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1914=¬1 te ha restaurado ¬2 puntos de vitalidad.
            Else
536             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1915, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1915=Te has restaurado ¬1 puntos de vitalidad.

            End If
            
538         Call WriteUpdateSta(tempChr)

540         b = True
542     ElseIf Hechizos(h).SubeSta = 2 Then

544         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
546         If UserIndex <> tempChr Then
548             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
550         Call InfoHechizo(UserIndex)
    
552         If UserIndex <> tempChr Then
554             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1916, Damage & "¬" & UserList(tempChr).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1916=Le has quitado ¬1 puntos de vitalidad a ¬2.
556             Call WriteConsoleMsg(tempChr, PrepareMessageLocaleMsg(1917, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1917=¬1 te ha quitado ¬2 puntos de vitalidad.
            Else
558             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1915, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1915=Te has restaurado ¬1 puntos de vitalidad.
            End If
    
560         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - Damage
562         If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
564         Call WriteUpdateSta(tempChr)
566         b = True
        End If
        Exit Sub
HechizoPropUsuario_Err:
568     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoPropUsuario", Erl)

        
End Sub

Sub HechizoCombinados(ByVal UserIndex As Integer, ByRef b As Boolean, ByRef IsAlive As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoCombinados_Err
        

        Dim h As Integer

        Dim Damage As Integer

        Dim TargetUserIndex           As Integer

        Dim enviarInfoHechizo As Boolean

100     enviarInfoHechizo = False
    
102     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
104     TargetUserIndex = UserList(UserIndex).flags.TargetUser.ArrayIndex
      
        ' <-------- Agilidad ---------->
106     If Hechizos(h).SubeAgilidad = 1 Then
    
            'Para poder tirar cl a un pk en el ring
108         If Not PeleaSegura(UserIndex, TargetUserIndex) Then
110             If Status(TargetUserIndex) = 0 And Status(UserIndex) = 1 Or Status(TargetUserIndex) = 2 And Status(UserIndex) = 1 Then
112                 If esArmada(UserIndex) Then
114                     Call WriteLocaleMsg(UserIndex, "379", e_FontTypeNames.FONTTYPE_INFO)
116                     b = False
                        Exit Sub

                    End If

118                 If UserList(UserIndex).flags.Seguro Then
120                     Call WriteLocaleMsg(UserIndex, "378", e_FontTypeNames.FONTTYPE_INFO)
122                     b = False
                        Exit Sub
                    Else
                    End If

                End If

            End If
    
124         enviarInfoHechizo = True
126         Damage = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
128         UserList(TargetUserIndex).flags.DuracionEfecto = Hechizos(h).Duration

130         UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) + Damage, UserList(TargetUserIndex).Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)
        
132         UserList(TargetUserIndex).flags.TomoPocion = True
134         b = True
136         Call WriteFYA(TargetUserIndex)
    
138     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
140         If Not PuedeAtacar(UserIndex, TargetUserIndex) Then Exit Sub
    
142         If UserIndex <> TargetUserIndex Then
144             Call UsuarioAtacadoPorUsuario(UserIndex, TargetUserIndex)

            End If
    
146         enviarInfoHechizo = True
    
148         UserList(TargetUserIndex).flags.TomoPocion = True
150         Damage = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
152         UserList(TargetUserIndex).flags.DuracionEfecto = Hechizos(h).Duration

154         If UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) - Damage < 6 Then
156             UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) = MINATRIBUTOS
            Else
158             UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) = UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Agilidad) - Damage
            End If
160         b = True
162         Call WriteFYA(TargetUserIndex)
        End If

        ' <-------- Fuerza ---------->
164     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
166         If Not PeleaSegura(UserIndex, TargetUserIndex) Then
168             If Status(TargetUserIndex) = 0 And Status(UserIndex) = 1 Or Status(TargetUserIndex) = 2 And Status(UserIndex) = 1 Then
170                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", e_FontTypeNames.FONTTYPE_INFO)
172                     Call WriteLocaleMsg(UserIndex, "379", e_FontTypeNames.FONTTYPE_INFO)
174                     b = False
                        Exit Sub

                    End If

176                 If UserList(UserIndex).flags.Seguro Then
178                     Call WriteLocaleMsg(UserIndex, "378", e_FontTypeNames.FONTTYPE_INFO)
180                     b = False
                        Exit Sub
                    End If
                End If
            End If
    
182         Damage = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
184         UserList(TargetUserIndex).flags.DuracionEfecto = Hechizos(h).Duration
186         UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) + Damage, UserList(TargetUserIndex).Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)
188         UserList(TargetUserIndex).flags.TomoPocion = True
190         b = True
192         enviarInfoHechizo = True
194         Call WriteFYA(TargetUserIndex)
196     ElseIf Hechizos(h).SubeFuerza = 2 Then
198         If Not PuedeAtacar(UserIndex, TargetUserIndex) Then Exit Sub
200         If UserIndex <> TargetUserIndex Then
202             Call UsuarioAtacadoPorUsuario(UserIndex, TargetUserIndex)
            End If
204         UserList(TargetUserIndex).flags.TomoPocion = True
206         Damage = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
208         UserList(TargetUserIndex).flags.DuracionEfecto = Hechizos(h).Duration
210         If UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) - Damage < 6 Then
212             UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) = MINATRIBUTOS
            Else
214             UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) = UserList(TargetUserIndex).Stats.UserAtributos(e_Atributos.Fuerza) - Damage
            End If
216         b = True
218         enviarInfoHechizo = True
220         Call WriteFYA(TargetUserIndex)
        End If

        'Salud
222     If IsSet(Hechizos(h).Effects, e_SpellEffects.eDoHeal) Then
            'Verifica que el usuario no este muerto
224         If UserList(TargetUserIndex).flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
226             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
228             b = False
                Exit Sub
            End If
            'Para poder tirar curar a un pk en el ring
230         If Not PeleaSegura(UserIndex, TargetUserIndex) Then
232             If Status(TargetUserIndex) = 0 And Status(UserIndex) = 1 Or Status(TargetUserIndex) = 2 And Status(UserIndex) = 1 Then
234                 If esArmada(UserIndex) Then
236                     Call WriteLocaleMsg(UserIndex, 379, e_FontTypeNames.FONTTYPE_INFO)
238                     b = False
                        Exit Sub
                    End If
240                 If UserList(UserIndex).flags.Seguro Then
242                     Call WriteLocaleMsg(UserIndex, 378, e_FontTypeNames.FONTTYPE_INFO)
244                     b = False
                        Exit Sub
                    End If
                End If
            End If
246         Damage = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            Damage = Damage * UserMod.GetMagicHealingBonus(UserList(UserIndex))
            Damage = Damage * UserMod.GetSelfHealingBonus(UserList(TargetUserIndex))
            
            If UserList(UserIndex).flags.DivineBlood > 0 Then
                Damage = Damage * DivineBloodHealingMultiplierBonus
            End If
            
248         enviarInfoHechizo = True
250         Call UserMod.DoDamageOrHeal(TargetUserIndex, UserIndex, eUser, Damage, e_DamageSourceType.e_magic, h)
    
254         If UserIndex <> TargetUserIndex Then
256             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1918, Damage & "¬" & UserList(targetUserIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1918=Le has restaurado ¬1 puntos de vida a ¬2.
258             Call WriteConsoleMsg(targetUserIndex, PrepareMessageLocaleMsg(1919, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1919=¬1 te ha restaurado ¬2 puntos de vida.
            Else
260             Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1920, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1920=Te has restaurado ¬1 puntos de vida.
            End If
264         b = True

266     ElseIf IsSet(Hechizos(h).Effects, e_SpellEffects.eDoDamage) Then ' Damage
268         If UserIndex = TargetUserIndex Then
270             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
271         If Not PuedeAtacar(UserIndex, TargetUserIndex) Then Exit Sub
272         Damage = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
274         Damage = Damage + Porcentaje(Damage, 3 * UserList(UserIndex).Stats.ELV)
            ' mage has 30% damage reduction
276         If UserList(UserIndex).clase = e_Class.Mage Then
278             Damage = Damage * 0.7
            End If
            Dim MR As Integer
            ' Weapon Magic bonus
280         If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
282             Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
                Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicAbsoluteBonus
                MR = MR - ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicPenetration
            End If
            
            ' Magic ring bonus
283         If UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex > 0 Then
                Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicDamageBonus)
                Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicAbsoluteBonus
                MR = MR - ObjData(UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex).MagicPenetration
            End If
284         If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
286             Damage = Damage + Porcentaje(Damage, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
                Damage = Damage + ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicAbsoluteBonus
                MR = MR - ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicPenetration
            End If
            ' Si el hechizo no ignora la RM
288         If Hechizos(h).AntiRm = 0 Then
                ' Resistencia mágica armadura
                MR = max(0, MR + GetUserMR(TargetUserIndex))
290             If MR > 0 Then
292                 Damage = Damage - Porcentaje(Damage, MR)
                End If
            End If
            Call EffectsOverTime.TartgetWillAtack(UserList(UserIndex).EffectOverTime, TargetUserIndex, eUser, e_DamageSourceType.e_magic)
            Damage = Damage * UserMod.GetMagicDamageModifier(UserList(UserIndex))
            Damage = Damage * UserMod.GetMagicDamageReduction(UserList(TargetUserIndex))
            ' Prevengo daño negativo
308         If Damage < 0 Then Damage = 0
    
312         If UserIndex <> TargetUserIndex Then
314             Call UsuarioAtacadoPorUsuario(UserIndex, TargetUserIndex)
            End If
    
316         enviarInfoHechizo = True
318         IsAlive = UserMod.DoDamageOrHeal(TargetUserIndex, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, h) = eStillAlive
321         Call EffectsOverTime.TartgetDidHit(UserList(UserIndex).EffectOverTime, TargetUserIndex, eUser, e_DamageSourceType.e_magic)
324         Call SubirSkill(TargetUserIndex, Resistencia)
336         b = True
        End If

        Dim tU As Integer
338     tU = TargetUserIndex
340     If IsSet(Hechizos(h).Effects, e_SpellEffects.Invisibility) Then
342         If UserList(tU).flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
344             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
346             b = False
                Exit Sub
            End If
    
348         If UserList(tU).Counters.Saliendo Then
350             If UserIndex <> tU Then
352                 ' Msg666=¡El hechizo no tiene efecto!
                    Call WriteLocaleMsg(UserIndex, "666", e_FontTypeNames.FONTTYPE_INFO)
354                 b = False
                    Exit Sub
                Else
356                 ' Msg667=¡No podés ponerte invisible mientras te encuentres saliendo!
                    Call WriteLocaleMsg(UserIndex, "667", e_FontTypeNames.FONTTYPE_WARNING)
358                 b = False
                    Exit Sub
                End If
            End If
           If IsSet(UserList(tU).flags.StatusMask, eTaunting) Then
               ' Msg666=¡El hechizo no tiene efecto!
                Call WriteLocaleMsg(UserIndex, "666", e_FontTypeNames.FONTTYPE_INFO)
               b = False
               Exit Sub
           End If
    
           If Not PeleaSegura(UserIndex, tU) Then
                    Select Case Status(UserIndex)
                        Case 1, 3, 5 'Ciudadano o armada
                            If Status(tU) <> e_Facciones.Ciudadano And Status(tU) <> e_Facciones.Armada And Status(tU) <> e_Facciones.consejo Then
                                If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
                                    ' Msg662=No puedes ayudar criminales.
                                    Call WriteLocaleMsg(UserIndex, "662", e_FontTypeNames.FONTTYPE_INFO)
                                    b = False
                                    Exit Sub
                                ElseIf Status(UserIndex) = e_Facciones.Ciudadano Then
                                    If UserList(UserIndex).flags.Seguro = True Then
                                        ' Msg663=Para ayudar criminales deberás desactivar el seguro.
                                        Call WriteLocaleMsg(UserIndex, "663", e_FontTypeNames.FONTTYPE_INFO)
                                        b = False
                                        Exit Sub
                                    Else
                                        'Si tiene clan
                                        If UserList(UserIndex).GuildIndex > 0 Then
                                            'Si el clan es de alineación ciudadana.
                                            If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                                'No lo dejo resucitarlo
                                                ' Msg664=No puedes ayudar a un usuario criminal perteneciendo a un clan ciudadano.
                                                Call WriteLocaleMsg(UserIndex, "664", e_FontTypeNames.FONTTYPE_INFO)
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
                                Call WriteLocaleMsg(UserIndex, "668", e_FontTypeNames.FONTTYPE_INFO)
                                b = False
                                Exit Sub
                            End If
                    End Select
                End If
    
            'Si sos user, no uses este hechizo con GMS.
378         If UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then
380             If Not UserList(tU).flags.Privilegios And e_PlayerType.user Then
                    Exit Sub
                End If
            End If
   
382         UserList(tU).flags.invisible = 1
            'Ladder
            'Reseteamos el contador de Invisibilidad
384         If UserList(tU).Counters.Invisibilidad <= 0 Then UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
386         Call WriteContadores(tU)
388         Call SendData(SendTarget.ToPCAliveArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.charindex, True, UserList(tU).Pos.X, UserList(tU).Pos.Y))
390         enviarInfoHechizo = True
392         b = True
        End If

394     If Hechizos(h).Envenena > 0 Then
396         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
398         If UserIndex <> tU Then
400             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
402         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
404         enviarInfoHechizo = True
406         b = True
        End If

408     If Hechizos(h).desencantar = 1 Then
410         ' Msg669=Has sido desencantado.
            Call WriteLocaleMsg(UserIndex, "669", e_FontTypeNames.FONTTYPE_INFO)
412         UserList(UserIndex).flags.Envenenado = 0
414         UserList(UserIndex).flags.Incinerado = 0
416         If UserList(UserIndex).flags.Inmovilizado = 1 Then
418             UserList(UserIndex).Counters.Inmovilizado = 0
                If UserList(UserIndex).clase = e_Class.Warrior Or UserList(UserIndex).clase = e_Class.Hunter Or UserList(UserIndex).clase = e_Class.Thief Or UserList(UserIndex).clase = e_Class.Pirat Then
                     UserList(UserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
420             UserList(UserIndex).flags.Inmovilizado = 0
422             Call WriteInmovilizaOK(UserIndex)
            End If
    
424         If UserList(UserIndex).flags.Paralizado = 1 Then
426             UserList(UserIndex).Counters.Paralisis = 0
                If UserList(UserIndex).clase = e_Class.Warrior Or UserList(UserIndex).clase = e_Class.Hunter Or UserList(UserIndex).clase = e_Class.Thief Or UserList(UserIndex).clase = e_Class.Pirat Then
                     UserList(UserIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
428             UserList(UserIndex).flags.Paralizado = 0
430             Call WriteParalizeOK(UserIndex)
            End If
        
432         If UserList(UserIndex).flags.Ceguera = 1 Then
434             UserList(UserIndex).Counters.Ceguera = 0
436             UserList(UserIndex).flags.Ceguera = 0
438             Call WriteBlindNoMore(UserIndex)
            End If
    
440         If UserList(UserIndex).flags.Maldicion = 1 Then
442             UserList(UserIndex).flags.Maldicion = 0
444             UserList(UserIndex).Counters.Maldicion = 0
            End If
446         enviarInfoHechizo = True
448         b = True
        End If

450     If Hechizos(h).Sanacion = 1 Then
451         UserList(tU).flags.Envenenado = 0
452         UserList(tU).flags.Incinerado = 0

            If UserList(tU).Counters.velocidad <> 0 Then
453             UserList(tU).flags.VelocidadHechizada = 0
454             UserList(tU).Counters.velocidad = 0
455             Call ActualizarVelocidadDeUsuario(tU)
            End If

456         enviarInfoHechizo = True
458         b = True
        End If

460     If IsSet(Hechizos(h).Effects, e_SpellEffects.Incinerate) Then
462         If UserIndex = tU Then
464             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
466         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
468         If UserIndex <> tU Then
470             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
472         UserList(tU).Counters.Incineracion = 1
474         UserList(tU).flags.Incinerado = 1
476         enviarInfoHechizo = True
478         b = True
        End If

480     If IsSet(Hechizos(h).Effects, e_SpellEffects.CurePoison) Then
            'Verificamos que el usuario no este muerto
482         If UserList(tU).flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
484             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
486             b = False
                Exit Sub
            End If
            'Para poder tirar curar veneno a un pk en el ring
488         If Not PeleaSegura(UserIndex, tU) Then
490             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
492                 If esArmada(UserIndex) Then
494                     Call WriteLocaleMsg(UserIndex, "379", e_FontTypeNames.FONTTYPE_INFO)
496                     b = False
                        Exit Sub
                    End If

498                 If UserList(UserIndex).flags.Seguro Then
500                     Call WriteLocaleMsg(UserIndex, "378", e_FontTypeNames.FONTTYPE_INFO)
502                     b = False
                        Exit Sub
                    End If
                End If
            End If
            'Si sos user, no uses este hechizo con GMS.
504         If UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then
506             If Not UserList(tU).flags.Privilegios And e_PlayerType.user Then
                    Exit Sub
                End If
            End If
508         UserList(tU).flags.Envenenado = 0
510         UserList(tU).Counters.Veneno = 0
512         enviarInfoHechizo = True
514         b = True
        End If

516     If IsSet(Hechizos(h).Effects, e_SpellEffects.Curse) Then
518         If UserIndex = tU Then
520             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
522         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
524         If UserIndex <> tU Then
526             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
528         UserList(tU).flags.Maldicion = 1
530         UserList(tU).Counters.Maldicion = 200
532         enviarInfoHechizo = True
534         b = True
        End If

536     If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveCurse) Then
538         UserList(tU).flags.Maldicion = 0
540         UserList(tU).Counters.Maldicion = 0
542         enviarInfoHechizo = True
544         b = True
        End If

546     If IsSet(Hechizos(h).Effects, e_SpellEffects.PreciseHit) Then
548         UserList(tU).flags.GolpeCertero = 1
550         enviarInfoHechizo = True
552         b = True
        End If

562     If IsSet(Hechizos(h).Effects, e_SpellEffects.Paralize) Then
564         If UserIndex = tU Then
566             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
568         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
570         If UserIndex <> tU Then
572             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
574         enviarInfoHechizo = True
576         b = True
578         UserList(tU).Counters.Paralisis = Hechizos(h).Duration
580         If UserList(tU).flags.Paralizado = 0 Then
582             UserList(tU).flags.Paralizado = 1
584             Call WriteParalizeOK(tU)
586             Call WritePosUpdate(tU)
            End If
        End If

588     If IsSet(Hechizos(h).Effects, e_SpellEffects.Immobilize) Then
590         If UserIndex = tU Then
592             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
594         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
596         If UserIndex <> tU Then
598             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
600         enviarInfoHechizo = True
602         b = True
604         UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration
606         If UserList(tU).flags.Inmovilizado = 0 Then
608             UserList(tU).flags.Inmovilizado = 1
610             Call WriteInmovilizaOK(tU)
612             Call WritePosUpdate(tU)
            End If
        End If

614     If IsSet(Hechizos(h).Effects, e_SpellEffects.RemoveParalysis) Then
            'Para poder tirar remo a un pk en el ring
616         If Not PeleaSegura(UserIndex, tU) Then
618             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
620                 If esArmada(UserIndex) Then
622                     Call WriteLocaleMsg(UserIndex, "379", e_FontTypeNames.FONTTYPE_INFO)
624                     b = False
                        Exit Sub
                    End If

626                 If UserList(UserIndex).flags.Seguro Then
628                     Call WriteLocaleMsg(UserIndex, "378", e_FontTypeNames.FONTTYPE_INFO)
630                     b = False
                        Exit Sub
                    Else
632                     Call VolverCriminal(UserIndex)
                    End If
                End If
            End If
            
634         If UserList(tU).flags.Inmovilizado = 1 Then
636             UserList(tU).Counters.Inmovilizado = 0
638             UserList(tU).flags.Inmovilizado = 0
                If UserList(tU).clase = e_Class.Warrior Or UserList(tU).clase = e_Class.Hunter Or UserList(tU).clase = e_Class.Thief Or UserList(tU).clase = e_Class.Pirat Then
                     UserList(tU).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
640             Call WriteInmovilizaOK(tU)
642             enviarInfoHechizo = True
644             b = True
            End If

646         If UserList(tU).flags.Paralizado = 1 Then
648             UserList(tU).Counters.Paralisis = 0
650             UserList(tU).flags.Paralizado = 0
                If UserList(tU).clase = e_Class.Warrior Or UserList(tU).clase = e_Class.Hunter Or UserList(tU).clase = e_Class.Thief Or UserList(tU).clase = e_Class.Pirat Then
                     UserList(tU).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
652             Call WriteParalizeOK(tU)
654             enviarInfoHechizo = True
656             b = True
            End If

        End If

658     If IsSet(Hechizos(h).Effects, e_SpellEffects.Blindness) Then
660         If UserIndex = tU Then
662             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
664         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
666         If UserIndex <> tU Then
668             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If

670         UserList(tU).flags.Ceguera = 1
672         UserList(tU).Counters.Ceguera = Hechizos(h).Duration
674         Call WriteBlind(tU)
676         enviarInfoHechizo = True
678         b = True
        End If

680     If IsSet(Hechizos(h).Effects, e_SpellEffects.Dumb) Then
682         If UserIndex = tU Then
684             Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If

686         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
688         If UserIndex <> tU Then
690             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If

692         If UserList(tU).flags.Estupidez = 0 Then
694             UserList(tU).flags.Estupidez = 1
696             UserList(tU).Counters.Estupidez = Hechizos(h).Duration
            End If
698         Call WriteDumb(tU)
700         enviarInfoHechizo = True
702         b = True
        End If

704     If Hechizos(h).velocidad <> 0 Then
706         If Hechizos(h).velocidad < 1 Then
708             If UserIndex = tU Then
710                 Call WriteLocaleMsg(UserIndex, "380", e_FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
712             If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            End If
            
714         enviarInfoHechizo = True
716         b = True
718         If UserList(tU).Counters.velocidad = 0 Then
720             UserList(tU).flags.VelocidadHechizada = Hechizos(h).velocidad
722             Call ActualizarVelocidadDeUsuario(tU)
            End If
724         UserList(tU).Counters.velocidad = Hechizos(h).Duration
        End If

726     If enviarInfoHechizo Then
728         Call InfoHechizo(UserIndex)
        End If
        Exit Sub

HechizoCombinados_Err:
730     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoCombinados", Erl)

        
End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
        
        On Error GoTo UpdateUserHechizos_Err

        Dim LoopC As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
104             Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
            Else
106             Call ChangeUserHechizo(UserIndex, Slot, 0)

            End If

        Else

            'Actualiza todos los slots
108         For LoopC = 1 To MAXUSERHECHIZOS

                'Actualiza el inventario
110             If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
112                 Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
                Else
114                 Call ChangeUserHechizo(UserIndex, LoopC, 0)

                End If

116         Next LoopC

        End If

        
        Exit Sub

UpdateUserHechizos_Err:
118     Call TraceError(Err.Number, Err.Description, "modHechizos.UpdateUserHechizos", Erl)

        
End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)
    On Error GoTo ChangeUserHechizo_Err
100     UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
        Call WriteChangeSpellSlot(UserIndex, Slot)
        Exit Sub
ChangeUserHechizo_Err:
    Call TraceError(Err.Number, Err.Description, "modHechizos.ChangeUserHechizo", Erl)
End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)
        
        On Error GoTo DesplazarHechizo_Err

100     If (Dire <> 1 And Dire <> -1) Then Exit Sub
102     If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

        Dim TempHechizo As Integer
        Dim SpellInterval As Long
        With UserList(UserIndex)
        
104         If Dire = 1 Then 'Mover arriba

106             If CualHechizo = 1 Then
108                 ' Msg670=No podés mover el hechizo en esa direccion.
                    Call WriteLocaleMsg(UserIndex, "670", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                
                Else
            
110                 TempHechizo = .Stats.UserHechizos(CualHechizo)
112                 .Stats.UserHechizos(CualHechizo) = .Stats.UserHechizos(CualHechizo - 1)
114                 .Stats.UserHechizos(CualHechizo - 1) = TempHechizo
                    SpellInterval = .Counters.UserHechizosInterval(CualHechizo)
115                 .Counters.UserHechizosInterval(CualHechizo) = .Counters.UserHechizosInterval(CualHechizo - 1)
116                 .Counters.UserHechizosInterval(CualHechizo - 1) = SpellInterval
                    'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
117                 If .flags.Hechizo = CualHechizo Then
118                     .flags.Hechizo = .flags.Hechizo - 1

120                 ElseIf .flags.Hechizo = CualHechizo - 1 Then
122                     .flags.Hechizo = .flags.Hechizo + 1

                    End If
                
                    .flags.ModificoHechizos = True
                End If

            Else 'mover abajo

124             If CualHechizo = MAXUSERHECHIZOS Then
126                 ' Msg670=No podés mover el hechizo en esa direccion.
                    Call WriteLocaleMsg(UserIndex, "670", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                
                Else
            
128                 TempHechizo = .Stats.UserHechizos(CualHechizo)
130                 .Stats.UserHechizos(CualHechizo) = .Stats.UserHechizos(CualHechizo + 1)
132                 .Stats.UserHechizos(CualHechizo + 1) = TempHechizo
                    SpellInterval = .Counters.UserHechizosInterval(CualHechizo)
133                 .Counters.UserHechizosInterval(CualHechizo) = .Counters.UserHechizosInterval(CualHechizo + 1)
134                 .Counters.UserHechizosInterval(CualHechizo + 1) = SpellInterval
                    'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
135                 If .flags.Hechizo = CualHechizo Then
136                     .flags.Hechizo = .flags.Hechizo + 1

138                 ElseIf .flags.Hechizo = CualHechizo + 1 Then
140                     .flags.Hechizo = .flags.Hechizo - 1

                    End If
                
                    .flags.ModificoHechizos = True
                End If

            End If
        
        End With
        
        Exit Sub

DesplazarHechizo_Err:
142     Call TraceError(Err.Number, Err.Description, "modHechizos.DesplazarHechizo", Erl)
        
End Sub

Private Sub AreaHechizo(UserIndex As Integer, NpcIndex As Integer, X As Byte, Y As Byte, npc As Boolean)
        On Error GoTo AreaHechizo_Err
        
        Dim calculo      As Integer
        Dim TilesDifUser As Integer
        Dim TilesDifNpc  As Integer
        Dim tilDif       As Integer
        Dim h2           As Integer
        Dim Hit          As Integer
        Dim Damage As Integer
        Dim porcentajeDesc As Integer

100     h2 = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

        'Calculo de descuesto de golpe por cercania.
102     TilesDifUser = X + Y

104     If npc Then
106         If IsSet(Hechizos(h2).Effects, e_SpellEffects.eDoDamage) Then
108             TilesDifNpc = NpcList(NpcIndex).Pos.X + NpcList(NpcIndex).Pos.Y
            
110             tilDif = TilesDifUser - TilesDifNpc
            
112             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

114             Hit = Hit + Porcentaje(Hit, 3 * UserList(UserIndex).Stats.ELV)
            
                ' Daño mágico arma
116             If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
118                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
                End If
                
                ' Daño mágico anillo
120             If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
122                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
                End If

                ' Disminuir daño con distancia
124             If tilDif <> 0 Then
126                 porcentajeDesc = Abs(tilDif) * 20
128                 Damage = Hit / 100 * porcentajeDesc
130                 Damage = Hit - Damage
                Else
132                 Damage = Hit
                End If
                
                ' Si el hechizo no ignora la RM
134             If Hechizos(h2).AntiRm = 0 Then
136                 Damage = Damage - NpcList(npcIndex).Stats.defM
                End If
                
                ' Prevengo daño negativo
138             If Damage < 0 Then Damage = 0
140             Call NPCs.DoDamageOrHeal(npcIndex, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, h2)
            
142             If UserList(UserIndex).ChatCombate = 1 Then
144                 Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1921, Damage & "¬" & NpcList(NpcIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1921=Le has causado ¬1 puntos de daño a ¬2.
                End If
            End If
            Exit Sub
        Else

154         TilesDifNpc = UserList(NpcIndex).Pos.X + UserList(NpcIndex).Pos.Y
156         tilDif = TilesDifUser - TilesDifNpc
158         If IsSet(Hechizos(h2).Effects, e_SpellEffects.eDoDamage) Then
160             If UserIndex = NpcIndex Then
                    Exit Sub
                End If

162             If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
164             If UserIndex <> NpcIndex Then
166                 Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
                End If
168             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)
170             Hit = Hit + Porcentaje(Hit, 3 * UserList(UserIndex).Stats.ELV)
                ' Daño mágico arma
172             If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
174                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MagicDamageBonus)
                End If
                ' Daño mágico anillo
176             If UserList(UserIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
178                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).invent.EquippedRingAccesoryObjIndex).MagicDamageBonus)
                End If

180             If tilDif <> 0 Then
182                 porcentajeDesc = Abs(tilDif) * 20
184                 Damage = Hit / 100 * porcentajeDesc
186                 Damage = Hit - Damage
                Else
188                 Damage = Hit
                End If
                
                ' Si el hechizo no ignora la RM
190             If Hechizos(h2).AntiRm = 0 Then
                    ' Resistencia mágica armadura
192                 If UserList(NpcIndex).invent.EquippedArmorObjIndex > 0 Then
194                     Damage = Damage - Porcentaje(Damage, ObjData(UserList(NpcIndex).invent.EquippedArmorObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica anillo
196                 If UserList(NpcIndex).invent.EquippedRingAccesoryObjIndex > 0 Then
198                     Damage = Damage - Porcentaje(Damage, ObjData(UserList(NpcIndex).invent.EquippedRingAccesoryObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica escudo
200                 If UserList(NpcIndex).invent.EquippedShieldObjIndex > 0 Then
202                     Damage = Damage - Porcentaje(Damage, ObjData(UserList(NpcIndex).invent.EquippedShieldObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica casco
204                 If UserList(NpcIndex).invent.EquippedHelmetObjIndex > 0 Then
206                     Damage = Damage - Porcentaje(Damage, ObjData(UserList(NpcIndex).invent.EquippedHelmetObjIndex).ResistenciaMagica)
                    End If
                   
                    ' Resistencia mágica de la clase
208                 Damage = Damage - Damage * ModClase(UserList(npcIndex).clase).ResistenciaMagica
                End If
                
                ' Prevengo daño negativo
210             If Damage < 0 Then Damage = 0
212             Call UserMod.DoDamageOrHeal(npcIndex, UserIndex, eUser, -Damage, e_DamageSourceType.e_magic, h2)
218             Call SubirSkill(NpcIndex, Resistencia)
220             Call WriteUpdateUserStats(NpcIndex)
            End If
                
230         If IsSet(Hechizos(h2).Effects, e_SpellEffects.eDoHeal) Then
232             If Not PeleaSegura(UserIndex, npcIndex) Then
234                 If Status(UserIndex) = 1 And Status(NpcIndex) <> 1 Then
                        Exit Sub
                    End If
                End If

236             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

238             If tilDif <> 0 Then
240                 porcentajeDesc = Abs(tilDif) * 20
242                 Damage = Hit / 100 * porcentajeDesc
244                 Damage = Hit - Damage
                Else
246                 Damage = Hit
                End If
                Damage = Damage * UserMod.GetMagicHealingBonus(UserList(UserIndex))
                Damage = Damage * NPCs.GetSelfHealingBonus(NpcList(NpcIndex))
248             Call UserMod.DoDamageOrHeal(npcIndex, UserIndex, eUser, Damage, e_DamageSourceType.e_magic, h2)
252             If UserIndex <> NpcIndex Then
254                 Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1922, Damage & "¬" & UserList(NpcIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1922=Le has restaurado ¬1 puntos de vida a ¬2.
256                 Call WriteConsoleMsg(NpcIndex, PrepareMessageLocaleMsg(1923, UserList(UserIndex).name & "¬" & Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1923=¬1 te ha restaurado ¬2 puntos de vida.
                Else
258                 Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1920, Damage, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1920=Te has restaurado ¬1 puntos de vida.
                End If
            End If
260         Call WriteUpdateUserStats(NpcIndex)
        End If
                
262     If Hechizos(h2).Envenena > 0 Then
264         If UserIndex = NpcIndex Then
                Exit Sub
            End If
                    
266         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
                
268         If UserIndex <> NpcIndex Then
270             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                    
272         UserList(NpcIndex).flags.Envenenado = Hechizos(h2).Envenena
274         Call WriteConsoleMsg(NpcIndex, PrepareMessageLocaleMsg(1924, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1924=¬1 te ha envenenado.

        End If
                
276     If IsSet(Hechizos(h2).Effects, e_SpellEffects.Paralize) Then
278         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
280         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
282         If UserIndex <> NpcIndex Then
284             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
            
'Msg823= Has sido paralizado.
Call WriteLocaleMsg(NpcIndex, "823", e_FontTypeNames.FONTTYPE_INFO)
288         UserList(NpcIndex).Counters.Paralisis = Hechizos(h2).Duration

290         If UserList(NpcIndex).flags.Paralizado = 0 Then
292             UserList(NpcIndex).flags.Paralizado = 1
294             Call WriteParalizeOK(NpcIndex)
            

            End If
            
        End If
                
296     If IsSet(Hechizos(h2).Effects, e_SpellEffects.Immobilize) Then
298         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
300         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
302         If UserIndex <> NpcIndex Then
304             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                    
'Msg824= Has sido inmovilizado.
Call WriteLocaleMsg(NpcIndex, "824", e_FontTypeNames.FONTTYPE_INFO)
308         UserList(NpcIndex).Counters.Inmovilizado = Hechizos(h2).Duration

310         If UserList(NpcIndex).flags.Inmovilizado = 0 Then
312             UserList(NpcIndex).flags.Inmovilizado = 1
314             Call WriteInmovilizaOK(NpcIndex)
316             Call WritePosUpdate(NpcIndex)
            
            End If

        End If
                
318     If IsSet(Hechizos(h2).Effects, e_SpellEffects.Blindness) Then
320         If UserIndex = NpcIndex Then
                Exit Sub
            End If
    
322         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
324         If UserIndex <> NpcIndex Then
326             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
            End If
                    
328         UserList(NpcIndex).flags.Ceguera = 1
330         UserList(NpcIndex).Counters.Ceguera = Hechizos(h2).Duration
'Msg825= Te han cegado.
Call WriteLocaleMsg(NpcIndex, "825", e_FontTypeNames.FONTTYPE_INFO)
334         Call WriteBlind(NpcIndex)
        End If
                
336     If Hechizos(h2).velocidad > 0 Then
338         If UserIndex = NpcIndex Then
                Exit Sub
            End If
    
340         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
342         If UserIndex <> NpcIndex Then
344             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
            End If

346         If UserList(NpcIndex).Counters.velocidad = 0 Then
348             UserList(NpcIndex).flags.VelocidadHechizada = Hechizos(h2).velocidad
350             Call ActualizarVelocidadDeUsuario(NpcIndex)
            End If

352         UserList(NpcIndex).Counters.velocidad = Hechizos(h2).Duration
        End If
                
354     If IsSet(Hechizos(h2).Effects, e_SpellEffects.Curse) Then
356         If UserIndex = NpcIndex Then
                Exit Sub
            End If
    
358         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
360         If UserIndex <> NpcIndex Then
362             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
            End If

'Msg826= Ahora estas maldito. No podras Atacar
Call WriteLocaleMsg(NpcIndex, "826", e_FontTypeNames.FONTTYPE_INFO)
366         UserList(NpcIndex).flags.Maldicion = 1
368         UserList(NpcIndex).Counters.Maldicion = Hechizos(h2).Duration
        End If
                
370     If IsSet(Hechizos(h2).Effects, e_SpellEffects.RemoveCurse) Then
'Msg827= Te han removido la maldicion.
Call WriteLocaleMsg(NpcIndex, "827", e_FontTypeNames.FONTTYPE_INFO)
374         UserList(NpcIndex).flags.Maldicion = 0
        End If
                
376     If IsSet(Hechizos(h2).Effects, e_SpellEffects.PreciseHit) Then
'Msg828= Tu proximo golpe sera certero.
Call WriteLocaleMsg(NpcIndex, "828", e_FontTypeNames.FONTTYPE_INFO)
380         UserList(NpcIndex).flags.GolpeCertero = 1
        End If
                  
388     If IsSet(Hechizos(h2).Effects, e_SpellEffects.Incinerate) Then
390         If UserIndex = NpcIndex Then
                Exit Sub
            End If
    
392         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
394         If UserIndex <> NpcIndex Then
396             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)
            End If
398         UserList(NpcIndex).flags.Incinerado = 1
'Msg829= Has sido Incinerado.
Call WriteLocaleMsg(NpcIndex, "829", e_FontTypeNames.FONTTYPE_INFO)
        End If
                
402     If IsSet(Hechizos(h2).Effects, e_SpellEffects.Invisibility) Then
'Msg830= Ahora sos invisible.
Call WriteLocaleMsg(NpcIndex, "830", e_FontTypeNames.FONTTYPE_INFO)
406         UserList(NpcIndex).flags.invisible = 1
408         UserList(NpcIndex).Counters.Invisibilidad = Hechizos(h2).Duration
410         Call WriteContadores(NpcIndex)
412         Call SendData(SendTarget.ToPCAliveArea, NpcIndex, PrepareMessageSetInvisible(UserList(NpcIndex).Char.charindex, True, UserList(NpcIndex).Pos.X, UserList(NpcIndex).Pos.Y))
        End If
                              
414     If Hechizos(h2).Sanacion = 1 Then
'Msg831= Has sido sanado.
Call WriteLocaleMsg(NpcIndex, "831", e_FontTypeNames.FONTTYPE_INFO)
418         UserList(NpcIndex).flags.Envenenado = 0
420         UserList(NpcIndex).flags.Incinerado = 0

            If UserList(NpcIndex).Counters.velocidad <> 0 Then
                UserList(NpcIndex).flags.VelocidadHechizada = 0
                UserList(NpcIndex).Counters.velocidad = 0
                Call ActualizarVelocidadDeUsuario(NpcIndex)
            End If
        End If
                
422     If IsSet(Hechizos(h2).Effects, e_SpellEffects.RemoveParalysis) Then
'Msg832= Has sido removido.
Call WriteLocaleMsg(NpcIndex, "832", e_FontTypeNames.FONTTYPE_INFO)
426         If UserList(NpcIndex).flags.Inmovilizado = 1 Then
428             UserList(NpcIndex).Counters.Inmovilizado = 0
430             UserList(NpcIndex).flags.Inmovilizado = 0
                If UserList(NpcIndex).clase = e_Class.Warrior Or UserList(NpcIndex).clase = e_Class.Hunter Or UserList(NpcIndex).clase = e_Class.Thief Or UserList(NpcIndex).clase = e_Class.Pirat Then
                     UserList(NpcIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
432             Call WriteInmovilizaOK(NpcIndex)
            End If

434         If UserList(NpcIndex).flags.Paralizado = 1 Then
436             UserList(NpcIndex).flags.Paralizado = 0
                If UserList(NpcIndex).clase = e_Class.Warrior Or UserList(NpcIndex).clase = e_Class.Hunter Or UserList(NpcIndex).clase = e_Class.Thief Or UserList(NpcIndex).clase = e_Class.Pirat Then
                     UserList(NpcIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
                'no need to crypt this
438             Call WriteParalizeOK(NpcIndex)
            End If
        End If
                
440     If Hechizos(h2).desencantar = 1 Then
'Msg833= Has sido desencantado.
Call WriteLocaleMsg(NpcIndex, "833", e_FontTypeNames.FONTTYPE_INFO)
444         UserList(NpcIndex).flags.Envenenado = 0
446         UserList(NpcIndex).Counters.Veneno = 0
448         UserList(NpcIndex).flags.Incinerado = 0
450         UserList(NpcIndex).Counters.Incineracion = 0

452         If UserList(NpcIndex).flags.Inmovilizado = 1 Then
454             UserList(NpcIndex).Counters.Inmovilizado = 0
456             UserList(NpcIndex).flags.Inmovilizado = 0
                If UserList(NpcIndex).clase = e_Class.Warrior Or UserList(NpcIndex).clase = e_Class.Hunter Or UserList(NpcIndex).clase = e_Class.Thief Or UserList(NpcIndex).clase = e_Class.Pirat Then
                     UserList(NpcIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
458             Call WriteInmovilizaOK(NpcIndex)
            End If
                    
460         If UserList(NpcIndex).flags.Paralizado = 1 Then
462             UserList(NpcIndex).flags.Paralizado = 0
                If UserList(NpcIndex).clase = e_Class.Warrior Or UserList(NpcIndex).clase = e_Class.Hunter Or UserList(NpcIndex).clase = e_Class.Thief Or UserList(NpcIndex).clase = e_Class.Pirat Then
                     UserList(NpcIndex).Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
464             UserList(NpcIndex).Counters.Paralisis = 0
466             Call WriteParalizeOK(NpcIndex)
            End If

468         If UserList(NpcIndex).flags.Ceguera = 1 Then
470             UserList(NpcIndex).Counters.Ceguera = 0
472             UserList(NpcIndex).flags.Ceguera = 0
474             Call WriteBlindNoMore(NpcIndex)
            End If

476         If UserList(NpcIndex).flags.Maldicion = 1 Then
478             UserList(NpcIndex).flags.Maldicion = 0
480             UserList(NpcIndex).Counters.Maldicion = 0
            End If
        End If
        Exit Sub
AreaHechizo_Err:
482     Call TraceError(Err.Number, Err.Description, "modHechizos.AreaHechizo", Erl)
End Sub

Private Sub AdjustNpcStatWithCasterLevel(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    Dim BaseHit As Integer
    Dim BonusDamage As Single
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

Public Sub UseSpellSlot(ByVal UserIndex As Integer, ByVal SpellSlot As Integer)
    On Error GoTo UseSpellSlot_Err
100     With UserList(UserIndex)
            
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         .flags.Hechizo = SpellSlot
            If UserMod.IsStun(.flags, .Counters) Then
                Call WriteLocaleMsg(UserIndex, 394, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
        
110         If .flags.Hechizo < 1 Or .flags.Hechizo > MAXUSERHECHIZOS Then
112             .flags.Hechizo = 0
            End If
        
114         If .flags.Hechizo <> 0 Then
116             If (.flags.Privilegios And e_PlayerType.Consejero) = 0 Then
                    If .Stats.UserHechizos(SpellSlot) <> 0 Then
120                     If Hechizos(.Stats.UserHechizos(SpellSlot)).AutoLanzar = 1 Then
122                         If .flags.Descansar Then Exit Sub
                            
124                         If .flags.Meditando Then
126                             .flags.Meditando = False
128                             .Char.FX = 0
130                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
                            End If
                        
                            'If exiting, cancel
132                         Call CancelExit(UserIndex)

134                         Call SetUserRef(UserList(UserIndex).flags.TargetUser, UserIndex)
136                         Call LanzarHechizo(.flags.Hechizo, UserIndex)
                        Else
                            If IsValidUserRef(.flags.GMMeSigue) Then
                                Call WriteNofiticarClienteCasteo(.flags.GMMeSigue.ArrayIndex, 1)
                            End If
                            If Hechizos(.Stats.UserHechizos(SpellSlot)).AreaAfecta > 0 Then
138                             Call WriteWorkRequestTarget(UserIndex, e_Skill.Magia, True, Hechizos(.Stats.UserHechizos(SpellSlot)).AreaRadio)
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
140     Call TraceError(Err.Number, Err.Description, "Protocol.UseSpellSlot", Erl)
End Sub
