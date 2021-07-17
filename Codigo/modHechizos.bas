Attribute VB_Name = "modHechizos"
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/C
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer, Optional ByVal IgnoreVisibilityCheck As Boolean = False)
      On Error GoTo NpcLanzaSpellSobreUser_Err

      Dim Daño As Integer
      Dim DañoStr As String

100   If Spell = 0 Then Exit Sub

102   With UserList(UserIndex)
104     If .flags.Muerto Then Exit Sub
    
        '¿NPC puede ver a través de la invisibilidad?
106     If Not IgnoreVisibilityCheck Then
108       If .flags.invisible = 1 Or .flags.Oculto = 1 Or .flags.Inmunidad = 1 Then Exit Sub
        End If

110     Call InfoHechizoDeNpcSobreUser(NpcIndex, UserIndex, Spell)
112     NpcList(NpcIndex).Contadores.IntervaloLanzarHechizo = GetTickCount()

114     If Hechizos(Spell).SubeHP = 1 Then
116       Daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)

118       .Stats.MinHp = MinimoInt(.Stats.MinHp + Daño, .Stats.MaxHp)

120       DañoStr = PonerPuntos(Daño)

          'Call WriteConsoleMsg(UserIndex, NpcList(NpcIndex).name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
122       Call WriteLocaleMsg(UserIndex, "32", FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).Name & "¬" & DañoStr)
124       Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageTextCharDrop(DañoStr, .Char.CharIndex, vbGreen))
126       Call WriteUpdateHP(UserIndex)

128     ElseIf Hechizos(Spell).SubeHP = 2 Then
130       Daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)

          ' Si el hechizo no ignora la RM
132       If Hechizos(Spell).AntiRm = 0 Then
            Dim PorcentajeRM As Integer

            ' Resistencia mágica armadura
134         If .Invent.ArmourEqpObjIndex > 0 Then
136           PorcentajeRM = PorcentajeRM + ObjData(.Invent.ArmourEqpObjIndex).ResistenciaMagica
            End If

            ' Resistencia mágica anillo
138         If .Invent.ResistenciaEqpObjIndex > 0 Then
140           PorcentajeRM = PorcentajeRM + ObjData(.Invent.ResistenciaEqpObjIndex).ResistenciaMagica
            End If

            ' Resistencia mágica escudo
142         If .Invent.EscudoEqpObjIndex > 0 Then
144           PorcentajeRM = PorcentajeRM + ObjData(.Invent.EscudoEqpObjIndex).ResistenciaMagica
            End If

            ' Resistencia mágica casco
146         If .Invent.CascoEqpObjIndex > 0 Then
148           PorcentajeRM = PorcentajeRM + ObjData(.Invent.CascoEqpObjIndex).ResistenciaMagica
            End If
        
150         PorcentajeRM = PorcentajeRM + 100 * ModClase(.clase).ResistenciaMagica
        
            ' Resto el porcentaje total
152         Daño = Daño - Porcentaje(Daño, PorcentajeRM)
          End If

154       If Daño < 0 Then Daño = 0

156       .Stats.MinHp = .Stats.MinHp - Daño

          'Call WriteLocaleMsg(UserIndex, "34", FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name & "¬" & DañoStr)
158       Call WriteConsoleMsg(UserIndex, NpcList(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
160       Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageTextCharDrop(DañoStr, .Char.CharIndex, vbRed))

162       Call SubirSkill(UserIndex, Resistencia)

          'Muere
164       If .Stats.MinHp < 1 Then
166         Call UserDie(UserIndex)
          Else
168         Call WriteUpdateHP(UserIndex)
          End If
        End If

        'Mana
170     If Hechizos(Spell).SubeMana = 1 Then
172       Daño = RandomNumber(Hechizos(Spell).MinMana, Hechizos(Spell).MaxMana)

174       .Stats.MinMAN = MinimoInt(.Stats.MinMAN + Daño, .Stats.MaxMAN)

176       Call WriteUpdateMana(UserIndex)
178       Call WriteConsoleMsg(UserIndex, NpcList(NpcIndex).Name & " te ha restaurado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_INFO)

180     ElseIf Hechizos(Spell).SubeMana = 2 Then
182       Daño = RandomNumber(Hechizos(Spell).MinMana, Hechizos(Spell).MaxMana)

184       .Stats.MinMAN = MaximoInt(.Stats.MinMAN - Daño, 0)

186       Call WriteUpdateMana(UserIndex)
188       Call WriteConsoleMsg(UserIndex, NpcList(NpcIndex).Name & " te ha quitado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_INFO)
        End If

190     If Hechizos(Spell).SubeAgilidad = 1 Then
192       Daño = RandomNumber(Hechizos(Spell).MinAgilidad, Hechizos(Spell).MaxAgilidad)

194       .flags.TomoPocion = True
196       .flags.DuracionEfecto = Hechizos(Spell).Duration
198       .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(.Stats.UserAtributos(eAtributos.Agilidad) + Daño, .Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2)

200       Call WriteFYA(UserIndex)
202     ElseIf Hechizos(Spell).SubeAgilidad = 2 Then
204       Daño = RandomNumber(Hechizos(Spell).MinAgilidad, Hechizos(Spell).MaxAgilidad)

206       .flags.TomoPocion = True
208       .flags.DuracionEfecto = Hechizos(Spell).Duration
210       .Stats.UserAtributos(eAtributos.Agilidad) = MaximoInt(MINATRIBUTOS, .Stats.UserAtributos(eAtributos.Agilidad) - Daño)

212       Call WriteFYA(UserIndex)
        End If

214     If Hechizos(Spell).SubeFuerza = 1 Then
216       Daño = RandomNumber(Hechizos(Spell).MinFuerza, Hechizos(Spell).MaxFuerza)

218       .flags.TomoPocion = True
220       .flags.DuracionEfecto = Hechizos(Spell).Duration
222       .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(.Stats.UserAtributos(eAtributos.Fuerza) + Daño, .Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2)

224       Call WriteFYA(UserIndex)
226     ElseIf Hechizos(Spell).SubeFuerza = 2 Then
228       Daño = RandomNumber(Hechizos(Spell).MinFuerza, Hechizos(Spell).MaxFuerza)

230       .flags.TomoPocion = True
232       .flags.DuracionEfecto = Hechizos(Spell).Duration
234       .Stats.UserAtributos(eAtributos.Fuerza) = MaximoInt(MINATRIBUTOS, .Stats.UserAtributos(eAtributos.Fuerza) - Daño)

236       Call WriteFYA(UserIndex)
        End If


238     If Hechizos(Spell).Paraliza = 1 Then
240       If .flags.Paralizado = 0 Then
242         .flags.Paralizado = 1
244         .Counters.Paralisis = Hechizos(Spell).Duration / 2

246         Call WriteParalizeOK(UserIndex)
248         Call WritePosUpdate(UserIndex)
          End If
        End If

250     If Hechizos(Spell).Inmoviliza = 1 Then
252       If .flags.Inmovilizado = 0 Then
254         .flags.Inmovilizado = 1
256         .Counters.Inmovilizado = Hechizos(Spell).Duration / 2

258         Call WriteInmovilizaOK(UserIndex)
260         Call WritePosUpdate(UserIndex)
          End If
        End If

262     If Hechizos(Spell).RemoverParalisis = 1 Then
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

282     If Hechizos(Spell).incinera > 0 Then
284       If .flags.Incinerado = 0 Then
286         .flags.Incinerado = 1
288         .Counters.Incineracion = Hechizos(Spell).Duration

290         Call WriteConsoleMsg(UserIndex, "Has sido incinerado por " & NpcList(NpcIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT)
          End If
        End If

292     If Hechizos(Spell).Envenena > 0 Then
294       If .flags.Envenenado = 0 Then
296         .flags.Envenenado = Hechizos(Spell).Envenena
298         .Counters.Veneno = Hechizos(Spell).Duration

300         Call WriteConsoleMsg(UserIndex, "Has sido incinerado por " & NpcList(NpcIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT)
          End If
        End If

302     If Hechizos(Spell).RemueveInvisibilidadParcial = 1 Then
304       If .flags.invisible + .flags.Oculto > 0 And .flags.NoDetectable = 0 Then
306         .flags.invisible = 0
308         .flags.Oculto = 0
310         .Counters.Invisibilidad = 0
312         .Counters.Ocultando = 0

314         Call WriteConsoleMsg(UserIndex, "Tu invisibilidad ya no tiene efecto.", FontTypeNames.FONTTYPE_INFOIAO)
316         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
          End If
        End If

318     If Hechizos(Spell).Estupidez > 0 Then
320       If .flags.Estupidez = 0 Then
322         .flags.Estupidez = Hechizos(Spell).Estupidez
324         .Counters.Estupidez = Hechizos(Spell).Duration

326         Call WriteConsoleMsg(UserIndex, "Has sido estupidizado por " & NpcList(NpcIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT)
328         Call WriteDumb(UserIndex)
          End If
330     ElseIf Hechizos(Spell).RemoverEstupidez > 0 Then
332       If .flags.Estupidez > 0 Then
334         .flags.Estupidez = 0
336         .Counters.Estupidez = 0

338         Call WriteConsoleMsg(UserIndex, NpcList(NpcIndex).Name & " te removio la estupidez.", FontTypeNames.FONTTYPE_FIGHT)
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

      End With

      Exit Sub

NpcLanzaSpellSobreUser_Err:
352   Call TraceError(Err.Number, Err.Description & " Hechizo: " & Spell, "modHechizos.NpcLanzaSpellSobreUser", Erl)


End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
      On Error GoTo NpcLanzaSpellSobreNpc_Err

      Dim Daño As Integer
      Dim DañoStr As String

100   With NpcList(TargetNPC)
  
102     .Contadores.IntervaloLanzarHechizo = GetTickCount()
  
104     If Hechizos(Spell).SubeHP = 1 Then ' Cura
106       Daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
108       DañoStr = PonerPuntos(Daño)
110       Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
112       Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
114       Call SendData(SendTarget.ToPCArea, TargetNPC, PrepareMessageTextCharDrop(DañoStr, .Char.CharIndex, vbGreen))

116       .Stats.MinHp = .Stats.MinHp + Daño

118       If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
120         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageNpcUpdateHP(TargetNPC))

122     ElseIf Hechizos(Spell).SubeHP = 2 Then

124       Daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
126       Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
128       Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
130       Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageTextOverChar(PonerPuntos(Daño), .Char.CharIndex, vbRed))

132       .Stats.MinHp = .Stats.MinHp - Daño

134       If .NPCtype = DummyTarget Then
136         Call DummyTargetAttacked(TargetNPC)
          End If

          ' Mascotas dan experiencia al amo
138       If .MaestroUser > 0 Then
140         Call CalcularDarExp(.MaestroUser, TargetNPC, Daño)

            ' NPC de invasión
142         If .flags.InvasionIndex Then
144           Call SumarScoreInvasion(.flags.InvasionIndex, .MaestroUser, Daño)
            End If
          End If

          'Muere
146       If .Stats.MinHp < 1 Then
148         .Stats.MinHp = 0
150         Call MuereNpc(TargetNPC, 0)
          Else
152         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageNpcUpdateHP(TargetNPC))
          End If


154     ElseIf Hechizos(Spell).Paraliza = 1 Then

156       If .flags.Paralizado = 0 Then
158         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
160         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

162         .flags.Paralizado = 1
164         .Contadores.Paralisis = Hechizos(Spell).Duration / 2

          End If

166     ElseIf Hechizos(Spell).Inmoviliza = 1 Then

168       If .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
170         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
172         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

174         .flags.Inmovilizado = 1
176         .Contadores.Inmovilizado = Hechizos(Spell).Duration / 2
          End If

178     ElseIf Hechizos(Spell).RemoverParalisis = 1 Then

180       If .flags.Paralizado + .flags.Inmovilizado > 0 Then
182         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
184         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

186         .flags.Paralizado = 0
188         .Contadores.Paralisis = 0
190         .flags.Inmovilizado = 0
192         .Contadores.Inmovilizado = 0

          End If

194     ElseIf Hechizos(Spell).incinera = 1 Then
196       If .flags.Incinerado = 0 Then
198         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))

200         If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
202           Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))

            End If

204         .flags.Incinerado = 1
          End If
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
        Dim TargetMap As MapBlock
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
        
110         If NpcList(NpcIndex).Target > 0 Then
112             PosCasteadaX = UserList(NpcList(NpcIndex).Target).Pos.X + RandomNumber(-2, 2)
114             PosCasteadaY = UserList(NpcList(NpcIndex).Target).Pos.Y + RandomNumber(-2, 2)
            Else
116             PosCasteadaX = NpcList(NpcIndex).Pos.X + RandomNumber(-2, 2)
118             PosCasteadaY = NpcList(NpcIndex).Pos.Y + RandomNumber(-1, 2)
            End If
       
120         For X = 1 To .AreaRadio
122             For Y = 1 To .AreaRadio

124                 TargetMap = MapData(NpcList(NpcIndex).Pos.Map, X + PosCasteadaX - mitadAreaRadio, PosCasteadaY + Y - mitadAreaRadio)
                
126                 If afectaUsers And TargetMap.UserIndex > 0 Then
128                     If Not UserList(TargetMap.UserIndex).flags.Muerto And Not EsGM(TargetMap.UserIndex) Then
130                         Call NpcLanzaSpellSobreUser(NpcIndex, TargetMap.UserIndex, SpellIndex, True)
                        End If

                    End If
                            
132                 If afectaNPCs And TargetMap.NpcIndex > 0 Then
134                     If NpcList(TargetMap.NpcIndex).Attackable Then
136                         Call NpcLanzaSpellSobreNpc(NpcIndex, TargetMap.NpcIndex, SpellIndex)
                        End If

                    End If
                            
138             Next Y
140         Next X

            ' El NPC invoca otros npcs independientes
142         If .Invoca = 1 Then
144             For X = 1 To .cant
146                 Call SpawnNpc(.NumNpc, NpcList(NpcIndex).Pos, True, False, False)
                    
148             Next X
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
112             Call WriteConsoleMsg(UserIndex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)

            Else
114             UserList(UserIndex).Stats.UserHechizos(j) = hIndex

116             Call UpdateUserHechizos(False, UserIndex, CByte(j))

                'Quitamos del inv el item
118             Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)

            End If
            
            UserList(UserIndex).flags.ModificoHechizos = True
        Else
120         Call WriteConsoleMsg(UserIndex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

AgregarHechizo_Err:
122     Call TraceError(Err.Number, Err.Description, "modHechizos.AgregarHechizo", Erl)

        
End Sub
            
Sub DecirPalabrasMagicas(ByVal Hechizo As Byte, ByVal UserIndex As Integer)
        
        On Error GoTo DecirPalabrasMagicas_Err

100     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(UserIndex).Char.CharIndex, vbCyan))
        Exit Sub

DecirPalabrasMagicas_Err:
102     Call TraceError(Err.Number, Err.Description, "modHechizos.DecirPalabrasMagicas", Erl)

        
End Sub

Private Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer, Optional ByVal Slot As Integer = 0) As Boolean
        On Error GoTo PuedeLanzar_Err

100     PuedeLanzar = False

        If HechizoIndex = 0 Then Exit Function
        
102     With UserList(UserIndex)

104         If UserList(UserIndex).flags.EnConsulta Then
106             Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estas en consulta.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

108         If .flags.Muerto = 1 Then
110             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        
112         If .flags.Privilegios And PlayerType.Consejero Then
                Exit Function
            End If

114         If MapInfo(.Pos.Map).SinMagia Then
116             Call WriteConsoleMsg(UserIndex, "Una fuerza mística te impide lanzar hechizos en esta zona.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Function
            End If
        
118         If .flags.Montado = 1 Then
120             Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estas montado.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

122         If Hechizos(HechizoIndex).NecesitaObj > 0 Then
124             If Not TieneObjEnInv(UserIndex, Hechizos(HechizoIndex).NecesitaObj, Hechizos(HechizoIndex).NecesitaObj2) Then
126                 Call WriteConsoleMsg(UserIndex, "Necesitas un " & ObjData(Hechizos(HechizoIndex).NecesitaObj).Name & " para lanzar el hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If

128         If Hechizos(HechizoIndex).CoolDown > 0 Then
                Dim Actual As Long
                Dim SegundosFaltantes As Long
130             Actual = GetTickCount()

132             If .Counters.UserHechizosInterval(Slot) + (Hechizos(HechizoIndex).CoolDown * 1000) < Actual Then
134                 SegundosFaltantes = Int((.Counters.UserHechizosInterval(Slot) + (Hechizos(HechizoIndex).CoolDown * 1000) - Actual) / 1000)
136                 Call WriteConsoleMsg(UserIndex, "Debes esperar " & SegundosFaltantes & " segundos para volver a tirar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Function
                End If
            End If

138         If .Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
140             Call WriteConsoleMsg(UserIndex, "No tenes suficientes puntos de magia para lanzar este hechizo, necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

142         If .Stats.MinHp < Hechizos(HechizoIndex).RequiredHP Then
144             Call WriteConsoleMsg(UserIndex, "No tenes suficiente vida. Necesitas " & Hechizos(HechizoIndex).RequiredHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

146         If .Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido Then
148             Call WriteLocaleMsg(UserIndex, "222", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

150         If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
152             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        
154         If .clase = eClass.Mage Then
156             If Hechizos(HechizoIndex).NeedStaff > 0 Then
158                 If .Invent.WeaponEqpObjIndex = 0 Then
160                     Call WriteConsoleMsg(UserIndex, "Necesitás un báculo para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function
                    End If
                
162                 If ObjData(.Invent.WeaponEqpObjIndex).Power < Hechizos(HechizoIndex).NeedStaff Then
164                     Call WriteConsoleMsg(UserIndex, "Necesitás un báculo más poderoso para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
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
104             Call WriteConsoleMsg(UserIndex, "No podés invocar criaturas durante un reto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
    
            Dim h As Integer, j As Integer, ind As Integer, Index As Integer
            Dim targetPos As WorldPos
    
106         targetPos.Map = .flags.TargetMap
108         targetPos.X = .flags.TargetX
110         targetPos.Y = .flags.TargetY
        
112         h = .Stats.UserHechizos(.flags.Hechizo)
    
114         If Hechizos(h).Invoca = 1 Then
    
116             If .NroMascotas >= MAXMASCOTAS Then Exit Sub
        
                'No deja invocar mas de 1 fatuo
118             If Hechizos(h).NumNpc = FUEGOFATUO And .NroMascotas >= 1 Then
120                 Call WriteConsoleMsg(UserIndex, "Para invocar el fuego fatuo no debes tener otras criaturas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
122             If Hechizos(h).NumNpc = ELEMENTAL_VIENTO And .NroMascotas >= 1 Then
124                 Call WriteConsoleMsg(UserIndex, "Para invocar elemental de viento no debes tener otras criaturas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
126             If Hechizos(h).NumNpc = ELEMENTAL_FUEGO And .NroMascotas >= 1 Then
128                 Call WriteConsoleMsg(UserIndex, "Para invocar el elemental de fuego no debes tener otras criaturas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
                'No permitimos se invoquen criaturas en zonas seguras
130             If MapInfo(.Pos.Map).Seguro = 1 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
132                 Call WriteConsoleMsg(UserIndex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
134             For j = 1 To Hechizos(h).cant

136                 If .NroMascotas < MAXMASCOTAS Then
138                     ind = SpawnNpc(Hechizos(h).NumNpc, targetPos, True, False, False, UserIndex)
140                     If ind > 0 Then
142                         .NroMascotas = .NroMascotas + 1
                        
144                         Index = FreeMascotaIndex(UserIndex)
                        
146                         .MascotasIndex(Index) = ind
148                         .MascotasType(Index) = NpcList(ind).Numero
                        
150                         NpcList(ind).MaestroUser = UserIndex
152                         NpcList(ind).Contadores.TiempoExistencia = IntervaloInvocacion
154                         NpcList(ind).GiveGLD = 0
                            
156                         Call FollowAmo(ind)

                        Else
                            Exit Sub
                        End If
                        
                    Else
                        Exit For
                    End If
                
158             Next j
            
160             Call InfoHechizo(UserIndex)
162             b = True
        
164         ElseIf Hechizos(h).Invoca = 2 Then
            
                ' Si tiene mascotas
166             If .NroMascotas > 0 Then
                    ' Tiene que estar en zona insegura
168                 If MapInfo(.Pos.Map).Seguro = 0 Then

                        Dim i As Integer
                    
                        ' Si no están guardadas las mascotas
170                     If .flags.MascotasGuardadas = 0 Then
172                         For i = 1 To MAXMASCOTAS
174                             If .MascotasIndex(i) > 0 Then
                                    ' Si no es un elemental, lo "guardamos"... lo matamos
176                                 If NpcList(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
                                        ' Le saco el maestro, para que no me lo quite de mis mascotas
178                                     NpcList(.MascotasIndex(i)).MaestroUser = 0
                                        ' Lo borro
180                                     Call QuitarNPC(.MascotasIndex(i))
                                        ' Saco el índice
182                                     .MascotasIndex(i) = 0
                                    
184                                     b = True
                                    End If
                                End If
                            Next
                        
186                         .flags.MascotasGuardadas = 1

                        ' Ya están guardadas, así que las invocamos
                        Else
188                         For i = 1 To MAXMASCOTAS
                                ' Si está guardada y no está ya en el mapa
190                             If .MascotasType(i) > 0 And .MascotasIndex(i) = 0 Then
192                                 .MascotasIndex(i) = SpawnNpc(.MascotasType(i), targetPos, True, True, False, UserIndex)

194                                 NpcList(.MascotasIndex(i)).MaestroUser = UserIndex
196                                 Call FollowAmo(.MascotasIndex(i))
                                
198                                 b = True
                                End If
                            Next
                        
200                         .flags.MascotasGuardadas = 0
                        End If
                
                    Else
202                     Call WriteConsoleMsg(UserIndex, "No puedes invocar tus mascotas en un mapa seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
            
                Else
204                 Call WriteConsoleMsg(UserIndex, "No tienes mascotas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

206             If b Then Call InfoHechizo(UserIndex)
            
            End If
    
        End With
    
        Exit Sub
    
HechizoInvocacion_Err:
208     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoTerrenoEstado")


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
        
108     If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
110         b = True

            'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
112         For TempX = PosCasteadaX - 11 To PosCasteadaX + 11
114             For TempY = PosCasteadaY - 11 To PosCasteadaY + 11

116                 If InMapBounds(PosCasteadaM, TempX, TempY) Then
118                     If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then

                            'hay un user
120                         If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.NoDetectable = 0 Then
122                             UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 0
124                             Call WriteConsoleMsg(MapData(PosCasteadaM, TempX, TempY).UserIndex, "Tu invisibilidad ya no tiene efecto.", FontTypeNames.FONTTYPE_INFOIAO)
126                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, False))

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
        Dim TargetMap As MapBlock
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
    
110     If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
    
112         If Hechizos(h).ParticleViaje > 0 Then
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, PosCasteadaX, PosCasteadaY))
                
            Else
116             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(Hechizos(h).FXgrh, PosCasteadaX, PosCasteadaY))

            End If

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
                    End If

                End If
                            
142             If afectaNPCs And TargetMap.NpcIndex > 0 Then
144                 If NpcList(TargetMap.NpcIndex).Attackable Then
146                     Call AreaHechizo(UserIndex, TargetMap.NpcIndex, PosCasteadaX, PosCasteadaY, True)
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
   
108     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.amount > 0 Or (MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).TileExit.Map > 0 Or UserList(UserIndex).flags.TargetUser <> 0 Then
110         b = False
            'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
112         Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)

        Else

114         If Hechizos(uh).TeleportX = 1 Then

116             If UserList(UserIndex).flags.Portal = 0 Then

118                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, -1, False))
         
120                 UserList(UserIndex).flags.PortalM = UserList(UserIndex).Pos.Map
122                 UserList(UserIndex).flags.PortalX = UserList(UserIndex).flags.TargetX
124                 UserList(UserIndex).flags.PortalY = UserList(UserIndex).flags.TargetY
            
126                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 600, Accion_Barra.Intermundia))

128                 UserList(UserIndex).Accion.AccionPendiente = True
130                 UserList(UserIndex).Accion.Particula = ParticulasIndex.Runa
132                 UserList(UserIndex).Accion.TipoAccion = Accion_Barra.Intermundia
134                 UserList(UserIndex).Accion.HechizoPendiente = uh
            
136                 If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
138                     Call DecirPalabrasMagicas(uh, UserIndex)

                    End If

140                 b = True
                Else
142                 Call WriteConsoleMsg(UserIndex, "No podés lanzar mas de un portal a la vez.", FontTypeNames.FONTTYPE_INFO)
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

        Dim MAT As obj

100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
 
102     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Then
104         b = False
106         Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
        Else
108         MAT.amount = Hechizos(h).MaterializaCant
110         MAT.ObjIndex = Hechizos(h).MaterializaObj
112         Call MakeObj(MAT, UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY)
            'Call WriteConsoleMsg(UserIndex, "Has materializado un objeto!!", FontTypeNames.FONTTYPE_INFO)
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
        
            Case TipoHechizo.uInvocacion 'Tipo 1
102             Call HechizoInvocacion(UserIndex, b)

104         Case TipoHechizo.uEstado 'Tipo 2
106             Call HechizoTerrenoEstado(UserIndex, b)

108         Case TipoHechizo.uMaterializa 'Tipo 3
110             Call HechizoMaterializacion(UserIndex, b)
            
112         Case TipoHechizo.uArea 'Tipo 5
114             Call HechizoSobreArea(UserIndex, b)
            
116         Case TipoHechizo.uPortal 'Tipo 6
118             Call HechizoPortal(UserIndex, b)

120         Case TipoHechizo.UFamiliar
122             Call InvocarFamiliar(UserIndex, b)
                
        End Select

124     If b Then
126         Call SubirSkill(UserIndex, Magia)

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

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 01/10/07
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
        'usuario
        '***************************************************
        
        On Error GoTo HandleHechizoUsuario_Err
        

        Dim b As Boolean

100     Select Case Hechizos(uh).Tipo

            Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
102             Call HechizoEstadoUsuario(UserIndex, b)

104         Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropUsuario(UserIndex, b)

108         Case TipoHechizo.uCombinados
110             Call HechizoCombinados(UserIndex, b)
    
        End Select

112     If b Then
114         Call SubirSkill(UserIndex, Magia)
            'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
116         UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
118         If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0

120         If Hechizos(uh).RequiredHP > 0 Then
122             UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - Hechizos(uh).RequiredHP
124             If UserList(UserIndex).Stats.MinHp < 0 Then UserList(UserIndex).Stats.MinHp = 1
126             Call WriteUpdateHP(UserIndex)
            End If

128         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
130         If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0

132         Call WriteUpdateMana(UserIndex)
134         Call WriteUpdateSta(UserIndex)
136         UserList(UserIndex).flags.TargetUser = 0

        End If

        
        Exit Sub

HandleHechizoUsuario_Err:
138     Call TraceError(Err.Number, Err.Description, "modHechizos.HandleHechizoUsuario", Erl)

        
End Sub

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

100     Select Case Hechizos(uh).Tipo

            Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
102             Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)

104         Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)

        End Select

108     If b Then
110         Call SubirSkill(UserIndex, Magia)
112         UserList(UserIndex).flags.TargetNPC = 0
114         UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

116         If Hechizos(uh).RequiredHP > 0 Then
118             If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
120             UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - Hechizos(uh).RequiredHP
122             Call WriteUpdateHP(UserIndex)
            End If

124         If UserList(UserIndex).Stats.MinHp < 0 Then UserList(UserIndex).Stats.MinHp = 1
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
        
100     uh = UserList(UserIndex).Stats.UserHechizos(Index)

102     If PuedeLanzar(UserIndex, uh, Index) Then

104         Select Case Hechizos(uh).Target

                Case TargetType.uUsuarios

106                 If UserList(UserIndex).flags.TargetUser > 0 Then
108                     If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
110                         Call HandleHechizoUsuario(UserIndex, uh)
                    
112                         If Hechizos(uh).CoolDown > 0 Then
114                             UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount()

                            End If

                        Else
116                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
118                     Call WriteConsoleMsg(UserIndex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
120             Case TargetType.uNPC

122                 If UserList(UserIndex).flags.TargetNPC > 0 Then
124                     If Abs(NpcList(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
126                         Call HandleHechizoNPC(UserIndex, uh)

128                         If Hechizos(uh).CoolDown > 0 Then
130                             UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount()
                    
                            End If
                    
                        Else
132                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
134                     Call WriteConsoleMsg(UserIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
136             Case TargetType.uUsuariosYnpc

138                 If UserList(UserIndex).flags.TargetUser > 0 Then
140                     If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
142                         Call HandleHechizoUsuario(UserIndex, uh)
                    
144                         If Hechizos(uh).CoolDown > 0 Then
146                             UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount()

                            End If

                        Else
148                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

150                 ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then

152                     If Abs(NpcList(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
154                         If Hechizos(uh).CoolDown > 0 Then
156                             UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount()

                            End If

158                         Call HandleHechizoNPC(UserIndex, uh)
                        Else
160                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
162                     Call WriteConsoleMsg(UserIndex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
164             Case TargetType.uTerreno

166                 If Hechizos(uh).CoolDown > 0 Then
168                     UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount()

                    End If

170                 Call HandleHechizoTerreno(UserIndex, uh)

            End Select
    
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
102     tU = UserList(UserIndex).flags.TargetUser

104     If Hechizos(h).Invisibilidad = 1 Then
   
106         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
108             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
110             b = False
                Exit Sub

            End If
            
112         If UserList(UserIndex).flags.EnReto Then
114             Call WriteConsoleMsg(UserIndex, "No podés lanzar invisibilidad durante un reto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
116         If UserList(UserIndex).flags.Navegando = 1 Then
118             Call WriteConsoleMsg(UserIndex, "No podés lanzar el hechizo mientras estás navegando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
120         If UserList(tU).flags.Navegando = 1 Then
122             Call WriteConsoleMsg(UserIndex, "No podés lanzarle el hechizo mientras está navegando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
124         If UserList(UserIndex).flags.Montado Then
126             Call WriteConsoleMsg(UserIndex, "No podés lanzar invisibilidad mientras usas una montura.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
    
128         If UserList(tU).Counters.Saliendo Then
130             If UserIndex <> tU Then
132                 Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
134                 b = False
                    Exit Sub
                Else
136                 Call WriteConsoleMsg(UserIndex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
138                 b = False
                    Exit Sub

                End If

            End If
    
            'No usar invi mapas InviSinEfecto
            ' If MapInfo(UserList(tU).Pos.map).InviSinEfecto > 0 Then
            '  Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
            '  b = False
            '   Exit Sub
            '  End If
    
            'Para poder tirar invi a un pk en el ring
140         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
142             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
144                 If esArmada(UserIndex) Then
146                     Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
148                     b = False
                        Exit Sub

                    End If

150                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
152                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
154                     b = False
                        Exit Sub
                    Else
156                     Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
158         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
160             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
            
162         If MapInfo(UserList(tU).Pos.Map).SinInviOcul Then
164             Call WriteConsoleMsg(UserIndex, "Una fuerza divina te impide usar invisibilidad en esta zona.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
166         If UserList(tU).flags.invisible = 1 Then
168             If tU = UserIndex Then
170                 Call WriteConsoleMsg(UserIndex, "¡Ya estás invisible!", FontTypeNames.FONTTYPE_INFO)
                Else
172                 Call WriteConsoleMsg(UserIndex, "¡El objetivo ya se encuentra invisible!", FontTypeNames.FONTTYPE_INFO)
                End If
174             b = False
                Exit Sub
            End If
   
176         UserList(tU).flags.invisible = 1
            'Ladder
            'Reseteamos el contador de Invisibilidad
178         UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
180         Call WriteContadores(tU)
182         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

184         Call InfoHechizo(UserIndex)
186         b = True

        End If
        
188     If Hechizos(h).Mimetiza = 1 Then

190         If UserList(UserIndex).flags.EnReto Then
192             Call WriteConsoleMsg(UserIndex, "No podés mimetizarte durante un reto.", FontTypeNames.FONTTYPE_INFO)
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
204             Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no tuvo efecto", FontTypeNames.FONTTYPE_INFO)
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
                
220             .flags.Mimetizado = e_EstadoMimetismo.FormaUsuario
                
                'ahora pongo local el del enemigo
222             .Char.Body = UserList(tU).Char.Body
224             .Char.Head = UserList(tU).Char.Head
226             .Char.CascoAnim = UserList(tU).Char.CascoAnim
228             .Char.ShieldAnim = UserList(tU).Char.ShieldAnim
230             .Char.WeaponAnim = UserList(tU).Char.WeaponAnim
232             .NameMimetizado = UserList(tU).Name

234             If UserList(tU).GuildIndex > 0 Then .NameMimetizado = .NameMimetizado & " <" & modGuilds.GuildName(UserList(tU).GuildIndex) & ">"
            
236             Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
238             Call RefreshCharStatus(UserIndex)
            End With
           
240        Call InfoHechizo(UserIndex)
242        b = True
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
262             Call WriteConsoleMsg(UserIndex, UserList(tU).Name & " ya esta envenenado. El hechizo no tuvo efecto.", FontTypeNames.FONTTYPE_INFO)
264             b = False
            End If
        End If

266     If Hechizos(h).desencantar = 1 Then
            ' Call WriteConsoleMsg(UserIndex, "Ningun efecto magico tiene efecto sobre ti ya.", FontTypeNames.FONTTYPE_INFOIAO)

268         UserList(UserIndex).flags.Envenenado = 0
270         UserList(UserIndex).Counters.Veneno = 0
272         UserList(UserIndex).flags.Incinerado = 0
274         UserList(UserIndex).Counters.Incineracion = 0
    
276         If UserList(UserIndex).flags.Inmovilizado > 0 Then
278             UserList(UserIndex).Counters.Inmovilizado = 0
280             UserList(UserIndex).flags.Inmovilizado = 0
282             Call WriteInmovilizaOK(UserIndex)
            End If

284         If UserList(UserIndex).flags.Paralizado > 0 Then
286             UserList(UserIndex).Counters.Paralisis = 0
288             UserList(UserIndex).flags.Paralizado = 0
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

316     If Hechizos(h).incinera > 0 Then
318         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
320             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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

334     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
336         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
338             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
340             b = False
                Exit Sub
            End If
            
            ' Si no esta envenenado, no hay nada mas que hacer
342         If UserList(tU).flags.Envenenado = 0 Then
344             Call WriteConsoleMsg(UserIndex, UserList(tU).Name & " no está envenenado, el hechizo no tiene efecto.", FontTypeNames.FONTTYPE_INFOIAO)
346             b = False
                Exit Sub
            End If
    
            'Para poder tirar curar veneno a un pk en el ring
348         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
350             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
352                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
354                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
356                     b = False
                        Exit Sub

                    End If

358                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
360                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
362                     b = False
                        Exit Sub

                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
364         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
366             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
368         UserList(tU).flags.Envenenado = 0
370         UserList(tU).Counters.Veneno = 0
372         Call InfoHechizo(UserIndex)
374         b = True

        End If

376     If Hechizos(h).Maldicion = 1 Then
378         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
380             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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

396     If Hechizos(h).RemoverMaldicion = 1 Then
398         UserList(tU).flags.Maldicion = 0
400         UserList(tU).Counters.Maldicion = 0
402         Call InfoHechizo(UserIndex)
404         b = True

        End If

406     If Hechizos(h).GolpeCertero = 1 Then
408         UserList(tU).flags.GolpeCertero = 1
410         Call InfoHechizo(UserIndex)
412         b = True

        End If

414     If Hechizos(h).Bendicion = 1 Then
416         UserList(tU).flags.Bendicion = 1
418         Call InfoHechizo(UserIndex)
420         b = True

        End If

422     If Hechizos(h).Paraliza = 1 Then
424         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
426             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
428         If UserList(tU).flags.Paralizado = 1 Then
430             Call WriteConsoleMsg(UserIndex, UserList(tU).Name & " ya está paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
432         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
434         If UserIndex <> tU Then
436             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
438         Call InfoHechizo(UserIndex)
440         b = True
            
442         UserList(tU).Counters.Paralisis = Hechizos(h).Duration

444         If UserList(tU).flags.Paralizado = 0 Then
446             UserList(tU).flags.Paralizado = 1
448             Call WriteParalizeOK(tU)
450             Call WritePosUpdate(tU)
            End If

        End If

452      If Hechizos(h).velocidad <> 0 Then
            'Verificamos que el usuario no este muerto
454         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
456             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
458             b = False
                Exit Sub
            End If

460         If Hechizos(h).velocidad < 1 Then
462             If UserIndex = tU Then
                    'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
464                 Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
    
466             If Not PuedeAtacar(UserIndex, tU) Then Exit Sub

            Else
                'Para poder tirar curar veneno a un pk en el ring
468             If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
470                 If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
472                     If esArmada(UserIndex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
474                         Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
476                         b = False
                            Exit Sub
    
                        End If
    
478                     If UserList(UserIndex).flags.Seguro Then
                            'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
480                         Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
482                         b = False
                            Exit Sub
    
                        End If
    
                    End If
    
                End If
            
                'Si sos user, no uses este hechizo con GMS.
484             If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
486                 If Not UserList(tU).flags.Privilegios And PlayerType.user Then
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

502     If Hechizos(h).Inmoviliza = 1 Then
504         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
506             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
508         If UserList(tU).flags.Paralizado = 1 Then
510             Call WriteConsoleMsg(UserIndex, UserList(tU).Name & " ya está paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
512         ElseIf UserList(tU).flags.Inmovilizado = 1 Then
514             Call WriteConsoleMsg(UserIndex, UserList(tU).Name & " ya está inmovilizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
516         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
518         If UserIndex <> tU Then
520             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
            
522         Call InfoHechizo(UserIndex)
524         b = True
            
526         UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration

528         UserList(tU).flags.Inmovilizado = 1
530         Call WriteInmovilizaOK(tU)
532         Call WritePosUpdate(tU)
            

        End If

534     If Hechizos(h).RemoverParalisis = 1 Then
        
            'Para poder tirar remo a un pk en el ring
536         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
538             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
540                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
542                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
544                     b = False
                        Exit Sub

                    End If

546                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
548                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
550                     b = False
                        Exit Sub
                    Else
552                     Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
        
554         If UserList(tU).flags.Inmovilizado = 0 And UserList(tU).flags.Paralizado = 0 Then
556             Call WriteConsoleMsg(UserIndex, "El objetivo no esta paralizado.", FontTypeNames.FONTTYPE_INFO)
558             b = False
                Exit Sub

            End If
        
560         If UserList(tU).flags.Inmovilizado = 1 Then
562             UserList(tU).Counters.Inmovilizado = 0
564             UserList(tU).flags.Inmovilizado = 0
566             Call WriteInmovilizaOK(tU)
568             Call WritePosUpdate(tU)
            End If
    
570         If UserList(tU).flags.Paralizado = 1 Then
572             UserList(tU).flags.Paralizado = 0
574             UserList(tU).Counters.Paralisis = 0

576             Call WriteParalizeOK(tU)
            End If

578         b = True
580         Call InfoHechizo(UserIndex)

        End If

582     If Hechizos(h).RemoverEstupidez = 1 Then
584         If UserList(tU).flags.Estupidez = 1 Then

                'Para poder tirar remo estu a un pk en el ring
586             If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
588                 If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
590                     If esArmada(UserIndex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
592                         Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
594                         b = False
                            Exit Sub

                        End If

596                     If UserList(UserIndex).flags.Seguro Then
                            'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
598                         Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
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

612     If Hechizos(h).Revivir = 1 Then
614         If UserList(tU).flags.Muerto = 1 Then

616             If UserList(UserIndex).flags.EnReto Then
618                 Call WriteConsoleMsg(UserIndex, "No podés revivir a nadie durante un reto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
    
                'No usar resu en mapas con ResuSinEfecto
                'If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
                '   Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                '   b = False
                '   Exit Sub
                ' End If
                
620             If UserList(UserIndex).clase <> Cleric Then
                    Dim PuedeRevivir As Boolean
                    
622                 If UserList(UserIndex).Invent.WeaponEqpObjIndex <> 0 Then
624                     If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Revive Then
626                         PuedeRevivir = True
                        End If
                    End If
                    
628                 If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex <> 0 Then
630                     If ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).Revive Then
632                         PuedeRevivir = True
                        End If
                    End If
                        
634                 If Not PuedeRevivir Then
636                     Call WriteConsoleMsg(UserIndex, "Necesitás un objeto con mayor poder mágico para poder revivir.", FontTypeNames.FONTTYPE_INFO)
638                     b = False
                        Exit Sub
                    End If
                End If
                
640             If UserList(tU).flags.SeguroResu Then
642                 Call WriteConsoleMsg(UserIndex, "El usuario tiene el seguro de resurrección activado.", FontTypeNames.FONTTYPE_INFO)
644                 Call WriteConsoleMsg(tU, UserList(UserIndex).Name & " está intentando revivirte. Desactiva el seguro de resurrección para permitirle hacerlo.", FontTypeNames.FONTTYPE_INFO)
646                 b = False
                    Exit Sub
                End If
        
648             If UserList(tU).Accion.TipoAccion = Accion_Barra.Resucitar Then
650                 Call WriteConsoleMsg(UserIndex, "El usuario ya esta siendo resucitado.", FontTypeNames.FONTTYPE_INFO)
652                 b = False
                    Exit Sub
                End If
        
                'Para poder tirar revivir a un pk en el ring
654             If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
656                 If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
658                     If esArmada(UserIndex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
660                         Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
662                         b = False
                            Exit Sub

                        End If

664                     If UserList(UserIndex).flags.Seguro Then
                            'call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
666                         Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
668                         b = False
                            Exit Sub
                        Else
670                         Call VolverCriminal(UserIndex)

                        End If

                    End If

                End If
672             UserList(tU).Counters.TimerBarra = 5
674             Call SendData(SendTarget.ToPCArea, tU, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, ParticulasIndex.Resucitar, UserList(tU).Counters.TimerBarra, False))
676             Call SendData(SendTarget.ToPCArea, tU, PrepareMessageBarFx(UserList(tU).Char.CharIndex, UserList(tU).Counters.TimerBarra, Accion_Barra.Resucitar))
678             UserList(tU).Accion.AccionPendiente = True
680             UserList(tU).Accion.Particula = ParticulasIndex.Resucitar
682             UserList(tU).Accion.TipoAccion = Accion_Barra.Resucitar
                
684             Call WriteUpdateHungerAndThirst(tU)
686             Call InfoHechizo(UserIndex)

688             b = True
                'Solo saco vida si es User. no quiero que exploten GMs por ahi.
        
                'Call RevivirUsuario(tU)
            Else
690             b = False

            End If

        End If

692     If Hechizos(h).Ceguera = 1 Then
694         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
696             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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

714     If Hechizos(h).Estupidez = 1 Then
716         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
718             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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

        
        Exit Sub

HechizoEstadoUsuario_Err:
738     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoEstadoUsuario", Erl)

        
End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
        On Error GoTo HechizoEstadoNPC_Err
        
100     If Hechizos(hIndex).Invisibilidad = 1 Then
102         Call InfoHechizo(UserIndex)
104         NpcList(NpcIndex).flags.invisible = 1
106         b = True

        End If

108     If Hechizos(hIndex).Envenena > 0 Then
110         If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
112             b = False
                Exit Sub

            End If

114         Call NPCAtacado(NpcIndex, UserIndex)
116         Call InfoHechizo(UserIndex)
118         NpcList(NpcIndex).flags.Envenenado = Hechizos(hIndex).Envenena
120         b = True

        End If

122     If Hechizos(hIndex).CuraVeneno = 1 Then
124         If NpcList(NpcIndex).flags.Envenenado > 0 Then
126             Call InfoHechizo(UserIndex)
128             NpcList(NpcIndex).flags.Envenenado = 0
130             b = True
            Else
132             Call WriteConsoleMsg(UserIndex, "La criatura no esta envenenada, el hechizo no tiene efecto.", FontTypeNames.FONTTYPE_INFOIAO)
134             b = False
            End If
        End If

136     If Hechizos(hIndex).RemoverMaldicion = 1 Then
138         Call InfoHechizo(UserIndex)
            'NpcList(NpcIndex).flags.Maldicion = 0
140         b = True

        End If

142     If Hechizos(hIndex).Bendicion = 1 Then
144         Call InfoHechizo(UserIndex)
146         NpcList(NpcIndex).flags.Bendicion = 1
148         b = True

        End If

150     If Hechizos(hIndex).Paraliza = 1 Then
152         If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
154             If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
156                 b = False
                    Exit Sub

                End If

158             Call NPCAtacado(NpcIndex, UserIndex)
160             Call InfoHechizo(UserIndex)
162             NpcList(NpcIndex).flags.Paralizado = 1
164             NpcList(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 6
166             NpcList(NpcIndex).flags.Inmovilizado = 0
168             NpcList(NpcIndex).Contadores.Inmovilizado = 0

170             Call AnimacionIdle(NpcIndex, False)
                
172             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
174             Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFOIAO)
176             b = False
                Exit Sub

            End If

        End If

178     If Hechizos(hIndex).RemoverParalisis = 1 Then
180         With NpcList(NpcIndex)
182             If .flags.Paralizado + .flags.Inmovilizado = 0 Then
184                 Call WriteConsoleMsg(UserIndex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFOIAO)
186                 b = False
                Else
                    ' Si el usuario es Armada o Caos y el NPC es de la misma faccion
188                 b = ((esArmada(UserIndex) Or esCaos(UserIndex)) And .flags.Faccion = UserList(UserIndex).Faccion.Status)
                    'O si es mi propia mascota
190                 b = b Or (.MaestroUser = UserIndex)
                    'O si es mascota de otro usuario de la misma faccion
192                 b = b Or ((esArmada(UserIndex) And esArmada(.MaestroUser)) Or (esCaos(UserIndex) And esCaos(.MaestroUser)))
                    
194                 If b Then
196                     Call InfoHechizo(UserIndex)
198                     .flags.Paralizado = 0
200                     .Contadores.Paralisis = 0
202                     .flags.Inmovilizado = 0
204                     .Contadores.Inmovilizado = 0
                    Else
206                     Call WriteConsoleMsg(UserIndex, "Solo podés remover la Parálisis de tus mascotas o de criaturas que pertenecen a tu facción.", FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                End If
            End With
        End If
 
208     If Hechizos(hIndex).Inmoviliza = 1 Then
210         If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
212             If NpcList(NpcIndex).flags.Paralizado <> 0 Then
214                 Call WriteConsoleMsg(UserIndex, "El NPC se encuentra paralizado.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

216             If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
218                 b = False
                    Exit Sub

                End If

220             Call NPCAtacado(NpcIndex, UserIndex)
222             NpcList(NpcIndex).flags.Inmovilizado = 1
224             NpcList(NpcIndex).Contadores.Inmovilizado = (Hechizos(hIndex).Duration * 6.5) * 6
226             NpcList(NpcIndex).flags.Paralizado = 0
228             NpcList(NpcIndex).Contadores.Paralisis = 0

230             Call AnimacionIdle(NpcIndex, True)

232             Call InfoHechizo(UserIndex)
234             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
236             Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFOIAO)

            End If

        End If

238     If Hechizos(hIndex).Mimetiza = 1 Then

240         If UserList(UserIndex).flags.EnReto Then
242             Call WriteConsoleMsg(UserIndex, "No podés mimetizarte durante un reto.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            End If
    
244         If UserList(UserIndex).flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
246             Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no tuvo efecto", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
248         If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
 
250         If UserList(UserIndex).clase = eClass.Druid Then

                'copio el char original al mimetizado
252             With UserList(UserIndex)
254                 .CharMimetizado.Body = .Char.Body
256                 .CharMimetizado.Head = .Char.Head
258                 .CharMimetizado.CascoAnim = .Char.CascoAnim
260                 .CharMimetizado.ShieldAnim = .Char.ShieldAnim
262                 .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                    
264                 .flags.Mimetizado = e_EstadoMimetismo.FormaBicho
                    
                    'ahora pongo lo del NPC.
266                 .Char.Body = NpcList(NpcIndex).Char.Body
268                 .Char.Head = NpcList(NpcIndex).Char.Head
270                 .Char.CascoAnim = NingunCasco
272                 .Char.ShieldAnim = NingunEscudo
274                 .Char.WeaponAnim = NingunArma
276                 .NameMimetizado = IIf(NpcList(NpcIndex).showName = 1, NpcList(NpcIndex).Name, vbNullString)

278                 Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
280                 Call RefreshCharStatus(UserIndex)
                End With
                
            Else
            
282             Call WriteConsoleMsg(UserIndex, "Solo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
                
            End If
        
284        Call InfoHechizo(UserIndex)
286        b = True
        End If
        
        Exit Sub

HechizoEstadoNPC_Err:
288     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoEstadoNPC", Erl)

        
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 14/08/2007
        'Handles the Spells that afect the Life NPC
        '14/08/2007 Pablo (ToxicWaste) - Orden general.
        '***************************************************
        
        On Error GoTo HechizoPropNPC_Err
        

        Dim Daño As Long
        
        Dim DañoStr As String
    
        'Salud
100     If Hechizos(hIndex).SubeHP = 1 Then
102         If NpcList(NpcIndex).Stats.MinHp < NpcList(NpcIndex).Stats.MaxHp Then
104             Daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
                'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        
106             Call InfoHechizo(UserIndex)
108             NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MinHp + Daño

110             If NpcList(NpcIndex).Stats.MinHp > NpcList(NpcIndex).Stats.MaxHp Then NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MaxHp

112             DañoStr = PonerPuntos(Daño)

                'Call WriteConsoleMsg(UserIndex, "Has curado " & Daño & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
114             Call WriteLocaleMsg(UserIndex, "388", FontTypeNames.FONTTYPE_FIGHT, "la criatura¬" & DañoStr)

116             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageTextOverChar(DañoStr, NpcList(NpcIndex).Char.CharIndex, vbGreen))
118             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageNpcUpdateHP(NpcIndex))
120             b = True
            Else
122             Call WriteConsoleMsg(UserIndex, "La criatura no tiene heridas que curar, el hechizo no tiene efecto.", FontTypeNames.FONTTYPE_INFOIAO)
124             b = False
            End If
        
126     ElseIf Hechizos(hIndex).SubeHP = 2 Then

128         If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
130             b = False
                Exit Sub
            End If
        
132         Call NPCAtacado(NpcIndex, UserIndex)
134         Daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
        
136         Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)

            ' Si al hechizo le afecta el daño mágico
138         If Hechizos(hIndex).StaffAffected Then
                ' Daño mágico arma
140             If UserList(UserIndex).clase = eClass.Mage Then
                    ' El mago tiene un 30% de daño reducido
142                 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
144                     Daño = Porcentaje(Daño, 70 + ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                    Else
146                     Daño = Daño * 0.7
                    End If
                Else
148                 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
150                     Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                    End If
                End If
                
                ' Daño mágico anillo
152             If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex > 0 Then
154                 Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).MagicDamageBonus)
                End If
            End If

156         b = True
        
158         If NpcList(NpcIndex).flags.Snd2 > 0 Then
160             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd2, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y))
            End If
        
            'Quizas tenga defenza magica el NPC.
162         If Hechizos(hIndex).AntiRm = 0 Then
164             Daño = Daño - NpcList(NpcIndex).Stats.defM
            End If
        
166         If Daño < 0 Then Daño = 0
        
168         NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MinHp - Daño
170         Call InfoHechizo(UserIndex)

            ' NPC de invasión
172         If NpcList(NpcIndex).flags.InvasionIndex Then
174             Call SumarScoreInvasion(NpcList(NpcIndex).flags.InvasionIndex, UserIndex, Daño)
            End If

176         If NpcList(NpcIndex).NPCtype = DummyTarget Then
178             Call DummyTargetAttacked(NpcIndex)
            End If
            
180         DañoStr = PonerPuntos(Daño)
        
182         If UserList(UserIndex).ChatCombate = 1 Then
                'Call WriteConsoleMsg(UserIndex, "Le has causado " & Daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
184             Call WriteLocaleMsg(UserIndex, "389", FontTypeNames.FONTTYPE_FIGHT, "la criatura¬" & DañoStr)
            End If
        
186         If NpcList(NpcIndex).MaestroUser <= 0 Then
188             Call CalcularDarExp(UserIndex, NpcIndex, Daño)
            End If
    
190         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageTextOverChar(DañoStr, NpcList(NpcIndex).Char.CharIndex, vbRed))
    
192         If NpcList(NpcIndex).Stats.MinHp < 1 Then
194             NpcList(NpcIndex).Stats.MinHp = 0
196             Call MuereNpc(NpcIndex, UserIndex)
            Else
198             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageNpcUpdateHP(NpcIndex))
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
106         Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageParticleFXWithDestino(NpcList(NpcIndex).Char.CharIndex, .Char.CharIndex, Hechizos(Spell).ParticleViaje, Hechizos(Spell).FXgrh, Hechizos(Spell).TimeParticula, Hechizos(Spell).wav, 1))
          Else
108         Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
          End If
        End If

110     If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
112       If Hechizos(Spell).ParticleViaje > 0 Then
114         Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageParticleFXWithDestino(NpcList(NpcIndex).Char.CharIndex, .Char.CharIndex, Hechizos(Spell).ParticleViaje, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, Hechizos(Spell).wav, 0))
          Else
116         Call SendData(SendTarget.ToPCArea, TargetUser, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))
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

106     If UserList(UserIndex).flags.TargetUser > 0 Then '¿El Hechizo fue tirado sobre un usuario?
108         If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
110             If Hechizos(h).ParticleViaje > 0 Then
112                 Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                Else
114                 Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageCreateFX(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                End If

            End If

116         If Hechizos(h).Particle > 0 Then '¿Envio Particula?
118             If Hechizos(h).ParticleViaje > 0 Then
120                 Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                Else
122                 Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageParticleFX(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

                End If

            End If
        
124         If Hechizos(h).ParticleViaje = 0 Then
126             Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserList(UserIndex).flags.TargetUser).Pos.X, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)

            End If
        
128         If Hechizos(h).TimeEfect <> 0 Then 'Envio efecto de screen
130             Call WriteFlashScreen(UserIndex, Hechizos(h).ScreenColor, Hechizos(h).TimeEfect)

            End If

132     ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then '¿El Hechizo fue tirado sobre un npc?

134         If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
136             If NpcList(UserList(UserIndex).flags.TargetNPC).Stats.MinHp < 1 Then

                    'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
138                 If Hechizos(h).ParticleViaje > 0 Then
140                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
                    Else
142                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))

                    End If

                Else

144                 If Hechizos(h).ParticleViaje > 0 Then
146                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                    Else
148                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageCreateFX(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                    End If

                End If

            End If
        
150         If Hechizos(h).Particle > 0 Then '¿Envio Particula?
152             If NpcList(UserList(UserIndex).flags.TargetNPC).Stats.MinHp < 1 Then
154                 Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0, NpcList(UserList(UserIndex).flags.TargetNPC).Pos.X, NpcList(UserList(UserIndex).flags.TargetNPC).Pos.Y))
                    'Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXToFloor(NpcList(UserList(UserIndex).flags.TargetNPC).Pos.X, NpcList(UserList(UserIndex).flags.TargetNPC).Pos.Y, Hechizos(H).Particle, Hechizos(H).TimeParticula))
                Else

156                 If Hechizos(h).ParticleViaje > 0 Then
158                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                    Else
160                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFX(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

                    End If

                End If

            End If

162         If Hechizos(h).ParticleViaje = 0 Then
164             Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(h).wav, NpcList(UserList(UserIndex).flags.TargetNPC).Pos.X, NpcList(UserList(UserIndex).flags.TargetNPC).Pos.Y))

            End If

        Else ' Entonces debe ser sobre el terreno

166         If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
168             Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))

            End If
        
170         If Hechizos(h).Particle > 0 Then 'Envio Particula?
172             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

            End If
        
174         If Hechizos(h).wav <> 0 Then
176             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))   'Esta linea faltaba. Pablo (ToxicWaste)

            End If
    
        End If
    
178     If UserList(UserIndex).ChatCombate = 1 Then
180         If Hechizos(h).Target = TargetType.uTerreno Then
182             Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)
            
184         ElseIf UserList(UserIndex).flags.TargetUser > 0 Then

                'Optimizacion de protocolo por Ladder
186             If UserIndex <> UserList(UserIndex).flags.TargetUser Then
188                 Call WriteConsoleMsg(UserIndex, "HecMSGU*" & h & "*" & UserList(UserList(UserIndex).flags.TargetUser).Name, FontTypeNames.FONTTYPE_FIGHT)
190                 Call WriteConsoleMsg(UserList(UserIndex).flags.TargetUser, "HecMSGA*" & h & "*" & UserList(UserIndex).Name, FontTypeNames.FONTTYPE_FIGHT)
    
                Else
192                 Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

                End If

194         ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
196             Call WriteConsoleMsg(UserIndex, "HecMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

            Else
198             Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

        
        Exit Sub

InfoHechizo_Err:
200     Call TraceError(Err.Number, Err.Description, "modHechizos.InfoHechizo", Erl)

        
End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
        On Error GoTo HechizoPropUsuario_Err
        

        Dim h As Integer
        Dim Daño As Integer
        Dim DañoStr As String
        Dim tempChr As Integer
    
100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
102     tempChr = UserList(UserIndex).flags.TargetUser
      
        'Hambre
104     If Hechizos(h).SubeHam = 1 Then
    
106         Call InfoHechizo(UserIndex)
    
108         Daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
110         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + Daño

112         If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
114         If UserIndex <> tempChr Then
116             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de hambre a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
118             Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
120             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

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
    
136         Daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
138         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - Daño
    
140         If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
142         If UserIndex <> tempChr Then
144             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de hambre a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
146             Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
148             Call WriteConsoleMsg(UserIndex, "Te has quitado " & Daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
150         Call WriteUpdateHungerAndThirst(tempChr)
    
152         b = True
    
154         If UserList(tempChr).Stats.MinHam < 1 Then
156             UserList(tempChr).Stats.MinHam = 0
158             UserList(tempChr).flags.Hambre = 1

            End If
    
        End If

        'Sed
160     If Hechizos(h).SubeSed = 1 Then
    
162         Call InfoHechizo(UserIndex)
    
164         Daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
166         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + Daño

168         If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
170         If UserIndex <> tempChr Then
172             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de sed a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
174             Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
176             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
178         Call WriteUpdateHungerAndThirst(tempChr)
    
180         b = True
    
182     ElseIf Hechizos(h).SubeSed = 2 Then
    
184         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
186         If UserIndex <> tempChr Then
188             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
190         Call InfoHechizo(UserIndex)
    
192         Daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
194         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - Daño
    
196         If UserIndex <> tempChr Then
198             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de sed a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
200             Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
202             Call WriteConsoleMsg(UserIndex, "Te has quitado " & Daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
204         If UserList(tempChr).Stats.MinAGU < 1 Then
206             UserList(tempChr).Stats.MinAGU = 0
208             UserList(tempChr).flags.Sed = 1

            End If
            
210         Call WriteUpdateHungerAndThirst(tempChr)
    
212         b = True

        End If

        ' <-------- Agilidad ---------->
214     If Hechizos(h).SubeAgilidad = 1 Then
    
            'Para poder tirar cl a un pk en el ring
216         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
218             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
220                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
222                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
224                     b = False
                        Exit Sub

                    End If

226                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
228                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
230                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
232         Call InfoHechizo(UserIndex)
234         Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
236         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

238         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + Daño, UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2)

240         UserList(tempChr).flags.TomoPocion = True
242         b = True
244         Call WriteFYA(tempChr)
    
246     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
248         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
250         If UserIndex <> tempChr Then
252             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
254         Call InfoHechizo(UserIndex)
    
256         UserList(tempChr).flags.TomoPocion = True
258         Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
260         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

262         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - Daño < MINATRIBUTOS Then
264             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
            Else
266             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - Daño

            End If
    
268         b = True
270         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
272     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
274         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
276             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
278                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
280                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
282                     b = False
                        Exit Sub

                    End If

284                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
286                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
288                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
290         Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
292         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
294         UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + Daño, UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2)

296         UserList(tempChr).flags.TomoPocion = True
            
298         Call WriteFYA(tempChr)

300         b = True
    
302         Call InfoHechizo(UserIndex)
304         Call WriteFYA(tempChr)

306     ElseIf Hechizos(h).SubeFuerza = 2 Then

308         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
310         If UserIndex <> tempChr Then
312             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
314         UserList(tempChr).flags.TomoPocion = True
    
316         Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
318         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

320         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - Daño < MINATRIBUTOS Then
322             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
            Else
324             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - Daño

            End If

326         b = True
328         Call InfoHechizo(UserIndex)
330         Call WriteFYA(tempChr)

        End If

        'Salud
332     If Hechizos(h).SubeHP = 1 Then
    
            'Verifica que el usuario no este muerto
334         If UserList(tempChr).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
336             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
338             b = False
                Exit Sub
            End If
            
340         If UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp Then
342             Call WriteConsoleMsg(UserIndex, UserList(tempChr).Name & " no tiene heridas para curar.", FontTypeNames.FONTTYPE_INFOIAO)
344             b = False
                Exit Sub
            End If
    
            'Para poder tirar curar a un pk en el ring
            Dim trigger As eTrigger6
346         trigger = TriggerZonaPelea(UserIndex, tempChr)

348         If trigger = TRIGGER6_AUSENTE Then
350             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
352                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
354                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
356                     b = False
                        Exit Sub

                    End If

358                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
360                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
362                     b = False
                        Exit Sub
                    End If

                End If
                
            ' Están en zona segura en un ring e intenta curarse desde afuera hacia adentro o viceversa
364         ElseIf trigger = TRIGGER6_PROHIBE And MapInfo(UserList(UserIndex).Pos.Map).Seguro <> 0 Then
366             b = False
                Exit Sub
            End If
       
368         Daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
370         Call InfoHechizo(UserIndex)

372         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + Daño

374         If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp

376         DañoStr = PonerPuntos(Daño)

378         If UserIndex <> tempChr Then
                'Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
380             Call WriteLocaleMsg(UserIndex, "388", FontTypeNames.FONTTYPE_FIGHT, UserList(tempChr).Name & "¬" & DañoStr)

                'Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
382             Call WriteLocaleMsg(tempChr, "32", FontTypeNames.FONTTYPE_FIGHT, UserList(UserIndex).Name & "¬" & DañoStr)
            Else
                'Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
384             Call WriteLocaleMsg(UserIndex, "33", FontTypeNames.FONTTYPE_FIGHT, DañoStr)
            End If
    
386         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageTextCharDrop(DañoStr, UserList(tempChr).Char.CharIndex, vbGreen))
388         Call WriteUpdateHP(tempChr)
    
390         b = True

392     ElseIf Hechizos(h).SubeHP = 2 Then
    
394         If UserIndex = tempChr Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
396             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

398         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
400         Daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
402         Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)

            ' Si al hechizo le afecta el daño mágico
404         If Hechizos(h).StaffAffected Then
                ' Daño mágico arma
406             If UserList(UserIndex).clase = eClass.Mage Then
                    ' El mago tiene un 30% de daño reducido
408                 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
410                     Daño = Porcentaje(Daño, 70 + ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                    Else
412                     Daño = Daño * 0.7
                    End If
                Else
414                 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
416                     Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                    End If
                End If
                
                ' Daño mágico anillo
418             If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex > 0 Then
420                 Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).MagicDamageBonus)
                End If
            End If
            
            ' Si el hechizo no ignora la RM
422         If Hechizos(h).AntiRm = 0 Then
                Dim PorcentajeRM As Integer

                ' Resistencia mágica armadura
424             If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
426                 PorcentajeRM = PorcentajeRM + ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).ResistenciaMagica
                End If
                
                ' Resistencia mágica anillo
428             If UserList(tempChr).Invent.ResistenciaEqpObjIndex > 0 Then
430                 PorcentajeRM = PorcentajeRM + ObjData(UserList(tempChr).Invent.ResistenciaEqpObjIndex).ResistenciaMagica
                End If
                
                ' Resistencia mágica escudo
432             If UserList(tempChr).Invent.EscudoEqpObjIndex > 0 Then
434                 PorcentajeRM = PorcentajeRM + ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).ResistenciaMagica
                End If
                
                ' Resistencia mágica casco
436             If UserList(tempChr).Invent.CascoEqpObjIndex > 0 Then
438                 PorcentajeRM = PorcentajeRM + ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).ResistenciaMagica
                End If
                
440             PorcentajeRM = PorcentajeRM + 100 * ModClase(UserList(tempChr).clase).ResistenciaMagica
                
                ' Resto el porcentaje total
442             Daño = Daño - Porcentaje(Daño, PorcentajeRM)
            End If

            ' Prevengo daño negativo
444         If Daño < 0 Then Daño = 0
    
446         If UserIndex <> tempChr Then
448             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
            End If
    
450         Call InfoHechizo(UserIndex)
    
452         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - Daño

454         DañoStr = PonerPuntos(Daño)
    
            'Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
456         Call WriteLocaleMsg(UserIndex, "389", FontTypeNames.FONTTYPE_FIGHT, UserList(tempChr).Name & "¬" & DañoStr)

            'Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
458         Call WriteLocaleMsg(tempChr, "34", FontTypeNames.FONTTYPE_FIGHT, UserList(UserIndex).Name & "¬" & DañoStr)
    
460         Call SubirSkill(tempChr, Resistencia)
    
462         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageTextCharDrop(Daño, UserList(tempChr).Char.CharIndex, vbRed))

            'Muere
464         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
466             Call Statistics.StoreFrag(UserIndex, tempChr)
468             Call ContarMuerte(tempChr, UserIndex)
470             Call ActStats(tempChr, UserIndex)
            Else
472             Call WriteUpdateHP(tempChr)
            End If

    
474         b = True

        End If

        'Mana
476     If Hechizos(h).SubeMana = 1 Then
    
478         Call InfoHechizo(UserIndex)
480         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + Daño

482         If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
484         If UserIndex <> tempChr Then
486             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de mana a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
488             Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
490             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

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
508             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de mana a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
510             Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
512             Call WriteConsoleMsg(UserIndex, "Te has quitado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
514         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - Daño

516         If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0

518         Call WriteUpdateMana(tempChr)

520         b = True
    
        End If

        'Stamina
522     If Hechizos(h).SubeSta = 1 Then
524         Call InfoHechizo(UserIndex)
526         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + Daño

528         If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta

530         If UserIndex <> tempChr Then
532             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
534             Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
536             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

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
554             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
556             Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
558             Call WriteConsoleMsg(UserIndex, "Te has quitado " & Daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
560         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - Daño
    
562         If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0

564         Call WriteUpdateSta(tempChr)

566         b = True

        End If

        Exit Sub

HechizoPropUsuario_Err:
568     Call TraceError(Err.Number, Err.Description, "modHechizos.HechizoPropUsuario", Erl)

        
End Sub

Sub HechizoCombinados(ByVal UserIndex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoCombinados_Err
        

        Dim h As Integer

        Dim Daño As Integer

        Dim tempChr           As Integer

        Dim enviarInfoHechizo As Boolean

100     enviarInfoHechizo = False
    
102     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
104     tempChr = UserList(UserIndex).flags.TargetUser
      
        ' <-------- Agilidad ---------->
106     If Hechizos(h).SubeAgilidad = 1 Then
    
            'Para poder tirar cl a un pk en el ring
108         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
110             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
112                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
114                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
116                     b = False
                        Exit Sub

                    End If

118                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
120                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
122                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
124         enviarInfoHechizo = True
126         Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
128         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
            'UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
            'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)

130         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + Daño, UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2)
        
132         UserList(tempChr).flags.TomoPocion = True
134         b = True
136         Call WriteFYA(tempChr)
    
138     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
140         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
142         If UserIndex <> tempChr Then
144             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
146         enviarInfoHechizo = True
    
148         UserList(tempChr).flags.TomoPocion = True
150         Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
152         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

154         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - Daño < 6 Then
156             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
            Else
158             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - Daño

            End If

            'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
160         b = True
162         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
164     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
166         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
168             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
170                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
172                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
174                     b = False
                        Exit Sub

                    End If

176                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
178                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
180                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
182         Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
184         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
186         UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + Daño, UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2)
188         UserList(tempChr).flags.TomoPocion = True
190         b = True
    
192         enviarInfoHechizo = True
194         Call WriteFYA(tempChr)
196     ElseIf Hechizos(h).SubeFuerza = 2 Then

198         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
200         If UserIndex <> tempChr Then
202             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
204         UserList(tempChr).flags.TomoPocion = True
    
206         Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
208         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        
210         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - Daño < 6 Then
212             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
            Else
214             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - Daño

            End If
   
216         b = True
218         enviarInfoHechizo = True
220         Call WriteFYA(tempChr)

        End If

        'Salud
222     If Hechizos(h).SubeHP = 1 Then
    
            'Verifica que el usuario no este muerto
224         If UserList(tempChr).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
226             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
228             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar a un pk en el ring
230         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
232             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
234                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
236                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
238                     b = False
                        Exit Sub

                    End If

240                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
242                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
244                     b = False
                        Exit Sub
                    Else

                        'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
       
246         Daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
248         enviarInfoHechizo = True

250         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + Daño

252         If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp
    
254         If UserIndex <> tempChr Then
256             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de vida a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
258             Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
260             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
262         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageTextOverChar(Daño, UserList(tempChr).Char.CharIndex, vbGreen))
    
264         b = True
266     ElseIf Hechizos(h).SubeHP = 2 Then ' Daño
    
268         If UserIndex = tempChr Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
270             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
272         Daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
274         Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)

            ' Los magos tienen 30% de daño reducido
276         If UserList(UserIndex).clase = eClass.Mage Then
278             Daño = Daño * 0.7
            End If
            
            ' Daño mágico arma
280         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
282             Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Daño mágico anillo
284         If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex > 0 Then
286             Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Si el hechizo no ignora la RM
288         If Hechizos(h).AntiRm = 0 Then
                ' Resistencia mágica armadura
290             If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
292                 Daño = Daño - Porcentaje(Daño, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica anillo
294             If UserList(tempChr).Invent.ResistenciaEqpObjIndex > 0 Then
296                 Daño = Daño - Porcentaje(Daño, ObjData(UserList(tempChr).Invent.ResistenciaEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica escudo
298             If UserList(tempChr).Invent.EscudoEqpObjIndex > 0 Then
300                 Daño = Daño - Porcentaje(Daño, ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica casco
302             If UserList(tempChr).Invent.CascoEqpObjIndex > 0 Then
304                 Daño = Daño - Porcentaje(Daño, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica de la clase
306             Daño = Daño - Daño * ModClase(UserList(tempChr).clase).ResistenciaMagica
            End If

            ' Prevengo daño negativo
308         If Daño < 0 Then Daño = 0
    
310         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
312         If UserIndex <> tempChr Then
314             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
316         enviarInfoHechizo = True
    
318         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - Daño
    
320         Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de vida a " & UserList(tempChr).Name, FontTypeNames.FONTTYPE_FIGHT)
322         Call WriteConsoleMsg(tempChr, UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
324         Call SubirSkill(tempChr, Resistencia)
326         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageTextOverChar(Daño, UserList(tempChr).Char.CharIndex, vbRed))

            'Muere
328         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
330             Call Statistics.StoreFrag(UserIndex, tempChr)
        
332             Call ContarMuerte(tempChr, UserIndex)
334             Call ActStats(tempChr, UserIndex)
            End If
    
336         b = True

        End If

        Dim tU As Integer

338     tU = tempChr

340     If Hechizos(h).Invisibilidad = 1 Then
   
342         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
344             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
346             b = False
                Exit Sub

            End If
    
348         If UserList(tU).Counters.Saliendo Then
350             If UserIndex <> tU Then
352                 Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
354                 b = False
                    Exit Sub
                Else
356                 Call WriteConsoleMsg(UserIndex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
358                 b = False
                    Exit Sub

                End If

            End If
    
            'Para poder tirar invi a un pk en el ring
360         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
362             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
364                 If esArmada(UserIndex) Then
366                     Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
368                     b = False
                        Exit Sub

                    End If

370                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
372                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
374                     b = False
                        Exit Sub
                    Else
376                     Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
378         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
380             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
   
382         UserList(tU).flags.invisible = 1
            'Ladder
            'Reseteamos el contador de Invisibilidad
384         UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
386         Call WriteContadores(tU)
388         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

390         enviarInfoHechizo = True
392         b = True

        End If

394     If Hechizos(h).Envenena > 0 Then
            ' If UserIndex = tU Then
            '    Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Exit Sub
            'End If
    
396         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
398         If UserIndex <> tU Then
400             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

402         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
404         enviarInfoHechizo = True
406         b = True

        End If

408     If Hechizos(h).desencantar = 1 Then
410         Call WriteConsoleMsg(UserIndex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)

412         UserList(UserIndex).flags.Envenenado = 0
414         UserList(UserIndex).flags.Incinerado = 0
    
416         If UserList(UserIndex).flags.Inmovilizado = 1 Then
418             UserList(UserIndex).Counters.Inmovilizado = 0
420             UserList(UserIndex).flags.Inmovilizado = 0
422             Call WriteInmovilizaOK(UserIndex)
            

            End If
    
424         If UserList(UserIndex).flags.Paralizado = 1 Then
426             UserList(UserIndex).Counters.Paralisis = 0
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

452         UserList(tU).flags.Envenenado = 0
454         UserList(tU).flags.Incinerado = 0
456         enviarInfoHechizo = True
458         b = True

        End If

460     If Hechizos(h).incinera = 1 Then
462         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
464             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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

480     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
482         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
484             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
486             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
488         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
490             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
492                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
494                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
496                     b = False
                        Exit Sub

                    End If

498                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
500                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
502                     b = False
                        Exit Sub
                    Else

                        '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
504         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
506             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
508         UserList(tU).flags.Envenenado = 0
510         UserList(tU).Counters.Veneno = 0
512         enviarInfoHechizo = True
514         b = True

        End If

516     If Hechizos(h).Maldicion = 1 Then
518         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
520             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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

536     If Hechizos(h).RemoverMaldicion = 1 Then
538         UserList(tU).flags.Maldicion = 0
540         UserList(tU).Counters.Maldicion = 0
542         enviarInfoHechizo = True
544         b = True

        End If

546     If Hechizos(h).GolpeCertero = 1 Then
548         UserList(tU).flags.GolpeCertero = 1
550         enviarInfoHechizo = True
552         b = True

        End If

554     If Hechizos(h).Bendicion = 1 Then
556         UserList(tU).flags.Bendicion = 1
558         enviarInfoHechizo = True
560         b = True

        End If

562     If Hechizos(h).Paraliza = 1 Then
564         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
566             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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

588     If Hechizos(h).Inmoviliza = 1 Then
590         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
592             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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

614     If Hechizos(h).RemoverParalisis = 1 Then
        
            'Para poder tirar remo a un pk en el ring
616         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
618             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
620                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
622                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
624                     b = False
                        Exit Sub

                    End If

626                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
628                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
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
640             Call WriteInmovilizaOK(tU)
642             enviarInfoHechizo = True
            
644             b = True

            End If

646         If UserList(tU).flags.Paralizado = 1 Then
648             UserList(tU).Counters.Paralisis = 0
650             UserList(tU).flags.Paralizado = 0

652             Call WriteParalizeOK(tU)
654             enviarInfoHechizo = True
            
656             b = True

            End If

        End If

658     If Hechizos(h).Ceguera = 1 Then
660         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
662             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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

680     If Hechizos(h).Estupidez = 1 Then
682         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
684             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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
                    'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
710                 Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
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
        

        'Call LogTarea("Sub UpdateUserHechizos")

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
        

        'Call LogTarea("ChangeUserHechizo")
    
100     UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
    
102     If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
104         Call WriteChangeSpellSlot(UserIndex, Slot)
        Else
106         Call WriteChangeSpellSlot(UserIndex, Slot)

        End If

        
        Exit Sub

ChangeUserHechizo_Err:
108     Call TraceError(Err.Number, Err.Description, "modHechizos.ChangeUserHechizo", Erl)

        
End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)
        
        On Error GoTo DesplazarHechizo_Err

100     If (Dire <> 1 And Dire <> -1) Then Exit Sub
102     If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

        Dim TempHechizo As Integer
        
        With UserList(UserIndex)
        
104         If Dire = 1 Then 'Mover arriba

106             If CualHechizo = 1 Then
108                 Call WriteConsoleMsg(UserIndex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                
                Else
            
110                 TempHechizo = .Stats.UserHechizos(CualHechizo)
112                 .Stats.UserHechizos(CualHechizo) = .Stats.UserHechizos(CualHechizo - 1)
114                 .Stats.UserHechizos(CualHechizo - 1) = TempHechizo

                    'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
116                 If .flags.Hechizo = CualHechizo Then
118                     .flags.Hechizo = .flags.Hechizo - 1

120                 ElseIf .flags.Hechizo = CualHechizo - 1 Then
122                     .flags.Hechizo = .flags.Hechizo + 1

                    End If
                
                    .flags.ModificoHechizos = True
                End If

            Else 'mover abajo

124             If CualHechizo = MAXUSERHECHIZOS Then
126                 Call WriteConsoleMsg(UserIndex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                
                Else
            
128                 TempHechizo = .Stats.UserHechizos(CualHechizo)
130                 .Stats.UserHechizos(CualHechizo) = .Stats.UserHechizos(CualHechizo + 1)
132                 .Stats.UserHechizos(CualHechizo + 1) = TempHechizo

                    'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
134                 If .flags.Hechizo = CualHechizo Then
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
        Dim Daño As Integer
        Dim porcentajeDesc As Integer

100     h2 = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

        'Calculo de descuesto de golpe por cercania.
102     TilesDifUser = X + Y

104     If npc Then
106         If Hechizos(h2).SubeHP = 2 Then
108             TilesDifNpc = NpcList(NpcIndex).Pos.X + NpcList(NpcIndex).Pos.Y
            
110             tilDif = TilesDifUser - TilesDifNpc
            
112             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

114             Hit = Hit + Porcentaje(Hit, 3 * UserList(UserIndex).Stats.ELV)
            
                ' Daño mágico arma
116             If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
118                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                End If
                
                ' Daño mágico anillo
120             If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex > 0 Then
122                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).MagicDamageBonus)
                End If

                ' Disminuir daño con distancia
124             If tilDif <> 0 Then
126                 porcentajeDesc = Abs(tilDif) * 20
128                 Daño = Hit / 100 * porcentajeDesc
130                 Daño = Hit - Daño
                Else
132                 Daño = Hit
                End If
                
                ' Si el hechizo no ignora la RM
134             If Hechizos(h2).AntiRm = 0 Then
136                 Daño = Daño - NpcList(NpcIndex).Stats.defM
                End If
                
                ' Prevengo daño negativo
138             If Daño < 0 Then Daño = 0
            
140             NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MinHp - Daño
            
142             If UserList(UserIndex).ChatCombate = 1 Then
144                 Call WriteConsoleMsg(UserIndex, "Le has causado " & Daño & " puntos de daño a " & NpcList(NpcIndex).Name, FontTypeNames.FONTTYPE_FIGHT)

                End If
            
146             Call CalcularDarExp(UserIndex, NpcIndex, Daño)
                
148             If NpcList(NpcIndex).Stats.MinHp <= 0 Then
                    'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + NpcList(NpcIndex).GiveEXP
                    'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + NpcList(NpcIndex).GiveGLD
150                 Call MuereNpc(NpcIndex, UserIndex)
                Else
152                 Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageNpcUpdateHP(NpcIndex))
                End If

            End If

            Exit Sub
        Else

154         TilesDifNpc = UserList(NpcIndex).Pos.X + UserList(NpcIndex).Pos.Y
156         tilDif = TilesDifUser - TilesDifNpc

158         If Hechizos(h2).SubeHP = 2 Then
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
172             If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
174                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                End If
                
                ' Daño mágico anillo
176             If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex > 0 Then
178                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).MagicDamageBonus)
                End If

180             If tilDif <> 0 Then
182                 porcentajeDesc = Abs(tilDif) * 20
184                 Daño = Hit / 100 * porcentajeDesc
186                 Daño = Hit - Daño
                Else
188                 Daño = Hit
                End If
                
                ' Si el hechizo no ignora la RM
190             If Hechizos(h2).AntiRm = 0 Then
                    ' Resistencia mágica armadura
192                 If UserList(NpcIndex).Invent.ArmourEqpObjIndex > 0 Then
194                     Daño = Daño - Porcentaje(Daño, ObjData(UserList(NpcIndex).Invent.ArmourEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica anillo
196                 If UserList(NpcIndex).Invent.ResistenciaEqpObjIndex > 0 Then
198                     Daño = Daño - Porcentaje(Daño, ObjData(UserList(NpcIndex).Invent.ResistenciaEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica escudo
200                 If UserList(NpcIndex).Invent.EscudoEqpObjIndex > 0 Then
202                     Daño = Daño - Porcentaje(Daño, ObjData(UserList(NpcIndex).Invent.EscudoEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica casco
204                 If UserList(NpcIndex).Invent.CascoEqpObjIndex > 0 Then
206                     Daño = Daño - Porcentaje(Daño, ObjData(UserList(NpcIndex).Invent.CascoEqpObjIndex).ResistenciaMagica)
                    End If
                   
                    ' Resistencia mágica de la clase
208                 Daño = Daño - Daño * ModClase(UserList(NpcIndex).clase).ResistenciaMagica
                End If
                
                ' Prevengo daño negativo
210             If Daño < 0 Then Daño = 0

212             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp - Daño
                    
214             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de vida a " & UserList(NpcIndex).Name, FontTypeNames.FONTTYPE_FIGHT)
216             Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
218             Call SubirSkill(NpcIndex, Resistencia)
220             Call WriteUpdateUserStats(NpcIndex)
                
                'Muere
222             If UserList(NpcIndex).Stats.MinHp < 1 Then
                    'Store it!
224                 Call Statistics.StoreFrag(UserIndex, NpcIndex)
                        
226                 Call ContarMuerte(NpcIndex, UserIndex)
228                 Call ActStats(NpcIndex, UserIndex)
                End If

            End If
                
230         If Hechizos(h2).SubeHP = 1 Then
232             If (TriggerZonaPelea(UserIndex, NpcIndex) <> TRIGGER6_PERMITE) Then
234                 If Status(UserIndex) = 1 And Status(NpcIndex) <> 1 Then
                        Exit Sub

                    End If

                End If

236             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

238             If tilDif <> 0 Then
240                 porcentajeDesc = Abs(tilDif) * 20
242                 Daño = Hit / 100 * porcentajeDesc
244                 Daño = Hit - Daño
                Else
246                 Daño = Hit

                End If
 
248             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp + Daño

250             If UserList(NpcIndex).Stats.MinHp > UserList(NpcIndex).Stats.MaxHp Then UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MaxHp
 
252             If UserIndex <> NpcIndex Then
254                 Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de vida a " & UserList(NpcIndex).Name, FontTypeNames.FONTTYPE_FIGHT)
256                 Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                Else
258                 Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

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
274         Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).Name & " te ha envenenado.", FontTypeNames.FONTTYPE_FIGHT)

        End If
                
276     If Hechizos(h2).Paraliza = 1 Then
278         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
280         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
282         If UserIndex <> NpcIndex Then
284             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
            
286         Call WriteConsoleMsg(NpcIndex, "Has sido paralizado.", FontTypeNames.FONTTYPE_INFO)
288         UserList(NpcIndex).Counters.Paralisis = Hechizos(h2).Duration

290         If UserList(NpcIndex).flags.Paralizado = 0 Then
292             UserList(NpcIndex).flags.Paralizado = 1
294             Call WriteParalizeOK(NpcIndex)
            

            End If
            
        End If
                
296     If Hechizos(h2).Inmoviliza = 1 Then
298         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
300         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
302         If UserIndex <> NpcIndex Then
304             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                    
306         Call WriteConsoleMsg(NpcIndex, "Has sido inmovilizado.", FontTypeNames.FONTTYPE_INFO)
308         UserList(NpcIndex).Counters.Inmovilizado = Hechizos(h2).Duration

310         If UserList(NpcIndex).flags.Inmovilizado = 0 Then
312             UserList(NpcIndex).flags.Inmovilizado = 1
314             Call WriteInmovilizaOK(NpcIndex)
316             Call WritePosUpdate(NpcIndex)
            
            End If

        End If
                
318     If Hechizos(h2).Ceguera = 1 Then
320         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
322         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
324         If UserIndex <> NpcIndex Then
326             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                    
328         UserList(NpcIndex).flags.Ceguera = 1
330         UserList(NpcIndex).Counters.Ceguera = Hechizos(h2).Duration
332         Call WriteConsoleMsg(NpcIndex, "Te han cegado.", FontTypeNames.FONTTYPE_INFO)
            
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
                
354     If Hechizos(h2).Maldicion = 1 Then
356         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
358         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
360         If UserIndex <> NpcIndex Then
362             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If

364         Call WriteConsoleMsg(NpcIndex, "Ahora estas maldito. No podras Atacar", FontTypeNames.FONTTYPE_INFO)
366         UserList(NpcIndex).flags.Maldicion = 1
368         UserList(NpcIndex).Counters.Maldicion = Hechizos(h2).Duration

        End If
                
370     If Hechizos(h2).RemoverMaldicion = 1 Then
372         Call WriteConsoleMsg(NpcIndex, "Te han removido la maldicion.", FontTypeNames.FONTTYPE_INFO)
374         UserList(NpcIndex).flags.Maldicion = 0

        End If
                
376     If Hechizos(h2).GolpeCertero = 1 Then
378         Call WriteConsoleMsg(NpcIndex, "Tu proximo golpe sera certero.", FontTypeNames.FONTTYPE_INFO)
380         UserList(NpcIndex).flags.GolpeCertero = 1

        End If
                
382     If Hechizos(h2).Bendicion = 1 Then
384         Call WriteConsoleMsg(NpcIndex, "Has sido bendecido.", FontTypeNames.FONTTYPE_INFO)
386         UserList(NpcIndex).flags.Bendicion = 1

        End If
                  
388     If Hechizos(h2).incinera = 1 Then
390         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
392         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
394         If UserIndex <> NpcIndex Then
396             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If

398         UserList(NpcIndex).flags.Incinerado = 1
400         Call WriteConsoleMsg(NpcIndex, "Has sido Incinerado.", FontTypeNames.FONTTYPE_INFO)

        End If
                
402     If Hechizos(h2).Invisibilidad = 1 Then
404         Call WriteConsoleMsg(NpcIndex, "Ahora sos invisible.", FontTypeNames.FONTTYPE_INFO)
406         UserList(NpcIndex).flags.invisible = 1
408         UserList(NpcIndex).Counters.Invisibilidad = Hechizos(h2).Duration
410         Call WriteContadores(NpcIndex)
412         Call SendData(SendTarget.ToPCArea, NpcIndex, PrepareMessageSetInvisible(UserList(NpcIndex).Char.CharIndex, True))

        End If
                              
414     If Hechizos(h2).Sanacion = 1 Then
416         Call WriteConsoleMsg(NpcIndex, "Has sido sanado.", FontTypeNames.FONTTYPE_INFO)
418         UserList(NpcIndex).flags.Envenenado = 0
420         UserList(NpcIndex).flags.Incinerado = 0

        End If
                
422     If Hechizos(h2).RemoverParalisis = 1 Then
424         Call WriteConsoleMsg(NpcIndex, "Has sido removido.", FontTypeNames.FONTTYPE_INFO)

426         If UserList(NpcIndex).flags.Inmovilizado = 1 Then
428             UserList(NpcIndex).Counters.Inmovilizado = 0
430             UserList(NpcIndex).flags.Inmovilizado = 0
432             Call WriteInmovilizaOK(NpcIndex)
            

            End If

434         If UserList(NpcIndex).flags.Paralizado = 1 Then
436             UserList(NpcIndex).flags.Paralizado = 0
                'no need to crypt this
438             Call WriteParalizeOK(NpcIndex)
            

            End If

        End If
                
440     If Hechizos(h2).desencantar = 1 Then
442         Call WriteConsoleMsg(NpcIndex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)
                    
444         UserList(NpcIndex).flags.Envenenado = 0
446         UserList(NpcIndex).Counters.Veneno = 0
448         UserList(NpcIndex).flags.Incinerado = 0
450         UserList(NpcIndex).Counters.Incineracion = 0

452         If UserList(NpcIndex).flags.Inmovilizado = 1 Then
454             UserList(NpcIndex).Counters.Inmovilizado = 0
456             UserList(NpcIndex).flags.Inmovilizado = 0
458             Call WriteInmovilizaOK(NpcIndex)

            End If
                    
460         If UserList(NpcIndex).flags.Paralizado = 1 Then
462             UserList(NpcIndex).flags.Paralizado = 0
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
