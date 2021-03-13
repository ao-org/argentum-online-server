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

Public Const SUPERANILLO       As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer, Optional ByVal IgnoreVisibilityCheck As Boolean = False)
        'Guardia caos
        
        On Error GoTo NpcLanzaSpellSobreUser_Err
        
100     With UserList(UserIndex)

102         If Spell = 0 Then Exit Sub
        
            '¿NPC puede ver a través de la invisibilidad?
104         If Not IgnoreVisibilityCheck Then
106             If .flags.invisible = 1 Or .flags.Oculto = 1 Or .flags.Inmunidad = 1 Then Exit Sub
            End If

            'NpcList(NpcIndex).CanAttack = 0
            Dim Daño As Integer
            
            Dim DañoStr As String

108         If Hechizos(Spell).SubeHP = 1 Then

110             Daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

116             .Stats.MinHp = .Stats.MinHp + Daño

118             If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp

                DañoStr = PonerPuntos(Daño)
    
                'Call WriteConsoleMsg(UserIndex, NpcList(NpcIndex).name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
120             Call WriteLocaleMsg(UserIndex, "32", FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name & "¬" & DañoStr)

121             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageTextCharDrop(DañoStr, .Char.CharIndex, vbGreen))

122             Call WriteUpdateHP(UserIndex)

126         ElseIf Hechizos(Spell).SubeHP = 2 Then

128             Daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)

                ' Si el hechizo no ignora la RM
404             If Hechizos(Spell).AntiRm = 0 Then
                    Dim PorcentajeRM As Integer

                    ' Resistencia mágica armadura
406                 If .Invent.ArmourEqpObjIndex > 0 Then
408                     PorcentajeRM = PorcentajeRM + ObjData(.Invent.ArmourEqpObjIndex).ResistenciaMagica
                    End If

                    ' Resistencia mágica anillo
410                 If .Invent.ResistenciaEqpObjIndex > 0 Then
412                     PorcentajeRM = PorcentajeRM + ObjData(.Invent.ResistenciaEqpObjIndex).ResistenciaMagica
                    End If

                    ' Resistencia mágica escudo
414                 If .Invent.EscudoEqpObjIndex > 0 Then
416                     PorcentajeRM = PorcentajeRM + ObjData(.Invent.EscudoEqpObjIndex).ResistenciaMagica
                    End If

                    ' Resistencia mágica casco
418                 If .Invent.CascoEqpObjIndex > 0 Then
420                     PorcentajeRM = PorcentajeRM + ObjData(.Invent.CascoEqpObjIndex).ResistenciaMagica
                    End If

                    ' Resto el porcentaje total
                    Daño = Daño - Porcentaje(Daño, PorcentajeRM)
                End If
        
138             If Daño < 0 Then Daño = 0
        
140             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
142             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
                
146             If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?

148                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))
          
                End If

150             .Stats.MinHp = .Stats.MinHp - Daño
        
                Call WriteConsoleMsg(UserIndex, NpcList(NpcIndex).name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                
152             'Call WriteLocaleMsg(UserIndex, "34", FontTypeNames.FONTTYPE_FIGHT, NpcList(NpcIndex).name & "¬" & DañoStr)

153             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageTextCharDrop(DañoStr, .Char.CharIndex, vbRed))

154             Call SubirSkill(UserIndex, Resistencia)
        
                'Muere
156             If .Stats.MinHp < 1 Then
158                 Call UserDie(UserIndex)
                Else
160                 Call WriteUpdateHP(UserIndex)
                End If
    
162         ElseIf Hechizos(Spell).Paraliza = 1 Then

164             If .flags.Paralizado = 0 Then
166                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
168                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

170                 .flags.Paralizado = 1
172                 .Counters.Paralisis = Hechizos(Spell).Duration / 2
          
174                 Call WriteParalizeOK(UserIndex)
176                 Call WritePosUpdate(UserIndex)

                End If

178         ElseIf Hechizos(Spell).incinera = 1 Then
180             Debug.Print "incinerar"

182             If .flags.Incinerado = 0 Then
184                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))

186                 If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
188                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))

                    End If

190                 .flags.Incinerado = 1
192                 Call WriteConsoleMsg(UserIndex, "Has sido incinerado por " & NpcList(NpcIndex).name & ".", FontTypeNames.FONTTYPE_FIGHT)

                End If

            End If
    
        End With

        Exit Sub

NpcLanzaSpellSobreUser_Err:
194     Call RegistrarError(Err.Number, Err.Description & " Hechizo: " & Spell, "modHechizos.NpcLanzaSpellSobreUser", Erl)

196     Resume Next
        
End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
        'solo hechizos ofensivos!
        
        On Error GoTo NpcLanzaSpellSobreNpc_Err
        

100     If NpcList(NpcIndex).CanAttack = 0 Then Exit Sub

102     NpcList(NpcIndex).CanAttack = 0

        Dim Daño As Integer

104     If Hechizos(Spell).SubeHP = 2 Then
    
106         Daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
108         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, NpcList(TargetNPC).Pos.X, NpcList(TargetNPC).Pos.Y))
110         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(NpcList(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
112         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageTextOverChar(PonerPuntos(Daño), NpcList(TargetNPC).Char.CharIndex, vbRed))
        
114         NpcList(TargetNPC).Stats.MinHp = NpcList(TargetNPC).Stats.MinHp - Daño

            If NpcList(NpcIndex).NPCtype = DummyTarget Then
                NpcList(NpcIndex).Contadores.UltimoAtaque = 30
            End If

            ' Mascotas dan experiencia al amo
116         If NpcList(NpcIndex).MaestroUser > 0 Then
118             Call CalcularDarExp(NpcList(NpcIndex).MaestroUser, TargetNPC, Daño)
            End If
        
            'Muere
120         If NpcList(TargetNPC).Stats.MinHp < 1 Then
122             NpcList(TargetNPC).Stats.MinHp = 0
                ' If NpcList(NpcIndex).MaestroUser > 0 Then
                '  Call MuereNpc(TargetNPC, NpcList(NpcIndex).MaestroUser)
                '  Else
124             Call MuereNpc(TargetNPC, 0)

                '  End If
            End If
    
        End If
    
        
        Exit Sub

NpcLanzaSpellSobreNpc_Err:
126     Call RegistrarError(Err.Number, Err.Description, "modHechizos.NpcLanzaSpellSobreNpc", Erl)
128     Resume Next
        
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

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal slot As Integer)
        
        On Error GoTo AgregarHechizo_Err
        

        Dim hIndex As Integer

        Dim j      As Integer

100     hIndex = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).HechizoIndex

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
118             Call QuitarUserInvItem(UserIndex, CByte(slot), 1)

            End If

        Else
120         Call WriteConsoleMsg(UserIndex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

AgregarHechizo_Err:
122     Call RegistrarError(Err.Number, Err.Description, "modHechizos.AgregarHechizo", Erl)
124     Resume Next
        
End Sub
            
Sub DecirPalabrasMagicas(ByVal Hechizo As Byte, ByVal UserIndex As Integer)
        
        On Error GoTo DecirPalabrasMagicas_Err

100     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(UserIndex).Char.CharIndex, vbCyan))
        Exit Sub

DecirPalabrasMagicas_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modHechizos.DecirPalabrasMagicas", Erl)

        
End Sub

Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer, Optional ByVal slot As Integer = 0) As Boolean
        
    On Error GoTo PuedeLanzar_Err
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If MapInfo(.Pos.Map).SinMagia Then
            Call WriteConsoleMsg(UserIndex, "Una fuerza mística te impide lanzar hechizos en esta zona.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If

        If Hechizos(HechizoIndex).NecesitaObj > 0 Then
            If Not TieneObjEnInv(UserIndex, Hechizos(HechizoIndex).NecesitaObj, Hechizos(HechizoIndex).NecesitaObj2) Then
                Call WriteConsoleMsg(UserIndex, "Necesitas un " & ObjData(Hechizos(HechizoIndex).NecesitaObj).name & " para lanzar el hechizo.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        End If

        If Hechizos(HechizoIndex).CoolDown > 0 Then
            Dim Actual As Long
            Dim SegundosFaltantes As Long
            Actual = GetTickCount()

            If .Counters.UserHechizosInterval(slot) + (Hechizos(HechizoIndex).CoolDown * 1000) < Actual Then
                SegundosFaltantes = Int((.Counters.UserHechizosInterval(slot) + (Hechizos(HechizoIndex).CoolDown * 1000) - Actual) / 1000)
                Call WriteConsoleMsg(UserIndex, "Debes esperar " & SegundosFaltantes & " segundos para volver a tirar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                Exit Function
            End If
        End If

        If .Stats.UserSkills(eSkill.magia) < Hechizos(HechizoIndex).MinSkill Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes puntos de magia para lanzar este hechizo, necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .Stats.MinHp < Hechizos(HechizoIndex).RequiredHP Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficiente vida. Necesitas " & Hechizos(HechizoIndex).RequiredHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido Then
            Call WriteLocaleMsg(UserIndex, "222", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
            Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        If .clase = eClass.Mage Then
            If Hechizos(HechizoIndex).NeedStaff > 0 Then
                If .Invent.WeaponEqpObjIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Necesitás un báculo para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
                
                If ObjData(.Invent.WeaponEqpObjIndex).Power < Hechizos(HechizoIndex).NeedStaff Then
                    Call WriteConsoleMsg(UserIndex, "Necesitás un báculo más poderoso para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        End If

        PuedeLanzar = True

    End With

    Exit Function

PuedeLanzar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modHechizos.PuedeLanzar", Erl)
    Resume Next
        
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

            If .flags.EnReto Then
                Call WriteConsoleMsg(UserIndex, "No podés invocar criaturas durante un reto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
    
            Dim h As Integer, j As Integer, ind As Integer, index As Integer
            Dim TargetPos As WorldPos
    
102         TargetPos.Map = .flags.TargetMap
104         TargetPos.X = .flags.TargetX
106         TargetPos.Y = .flags.TargetY
        
108         h = .Stats.UserHechizos(.flags.Hechizo)
    
110         If Hechizos(h).Invoca = 1 Then
    
112             If .NroMascotas >= MAXMASCOTAS Then Exit Sub
        
                'No deja invocar mas de 1 fatuo
114             If Hechizos(h).NumNpc = FUEGOFATUO And .NroMascotas >= 1 Then
116                 Call WriteConsoleMsg(UserIndex, "Para invocar el fuego fatuo no debes tener otras criaturas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
                'No permitimos se invoquen criaturas en zonas seguras
118             If MapInfo(.Pos.Map).Seguro = 1 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
120                 Call WriteConsoleMsg(UserIndex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
122             For j = 1 To Hechizos(h).cant
                
124                 If .NroMascotas < MAXMASCOTAS Then
126                     ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False, False, UserIndex)
128                     If ind > 0 Then
130                         .NroMascotas = .NroMascotas + 1
                        
132                         index = FreeMascotaIndex(UserIndex)
                        
134                         .MascotasIndex(index) = ind
136                         .MascotasType(index) = NpcList(ind).Numero
                        
138                         NpcList(ind).MaestroUser = UserIndex
140                         NpcList(ind).Contadores.TiempoExistencia = IntervaloInvocacion
142                         NpcList(ind).GiveGLD = 0
                        
144                         Call FollowAmo(ind)
                        Else
                            Exit Sub
                        End If
                        
                    Else
                        Exit For
                    End If
                
146             Next j
            
148             Call InfoHechizo(UserIndex)
150             b = True
        
152         ElseIf Hechizos(h).Invoca = 2 Then
            
                ' Si tiene mascotas
154             If .NroMascotas > 0 Then
                    ' Tiene que estar en zona insegura
156                 If MapInfo(.Pos.Map).Seguro = 0 Then

                        Dim i As Integer
                    
                        ' Si no están guardadas las mascotas
158                     If .flags.MascotasGuardadas = 0 Then
160                         For i = 1 To MAXMASCOTAS
162                             If .MascotasIndex(i) > 0 Then
                                    ' Si no es un elemental, lo "guardamos"... lo matamos
164                                 If NpcList(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
                                        ' Le saco el maestro, para que no me lo quite de mis mascotas
166                                     NpcList(.MascotasIndex(i)).MaestroUser = 0
                                        ' Lo borro
168                                     Call QuitarNPC(.MascotasIndex(i))
                                        ' Saco el índice
170                                     .MascotasIndex(i) = 0
                                    
172                                     b = True
                                    End If
                                End If
                            Next
                        
174                         .flags.MascotasGuardadas = 1

                        ' Ya están guardadas, así que las invocamos
                        Else
176                         For i = 1 To MAXMASCOTAS
                                ' Si está guardada y no está ya en el mapa
178                             If .MascotasType(i) > 0 And .MascotasIndex(i) = 0 Then
180                                 .MascotasIndex(i) = SpawnNpc(.MascotasType(i), TargetPos, True, True, False, UserIndex)

182                                 NpcList(.MascotasIndex(i)).MaestroUser = UserIndex
184                                 Call FollowAmo(.MascotasIndex(i))
                                
186                                 b = True
                                End If
                            Next
                        
188                         .flags.MascotasGuardadas = 0
                        End If
                
                    Else
190                     Call WriteConsoleMsg(UserIndex, "No puedes invocar tus mascotas en un mapa seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
            
                Else
192                 Call WriteConsoleMsg(UserIndex, "No tienes mascotas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

194             If b Then Call InfoHechizo(UserIndex)
            
            End If
    
        End With
    
        Exit Sub
    
HechizoInvocacion_Err:
196     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoTerrenoEstado")
198     Resume Next

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
134     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoTerrenoEstado", Erl)
136     Resume Next
        
End Sub

Sub HechizoSobreArea(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoSobreArea_Err
        

        Dim PosCasteadaX As Byte

        Dim PosCasteadaY As Byte

        Dim PosCasteadaM As Integer

        Dim h            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     PosCasteadaX = UserList(UserIndex).flags.TargetX
102     PosCasteadaY = UserList(UserIndex).flags.TargetY
104     PosCasteadaM = UserList(UserIndex).flags.TargetMap
 
106     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

        Dim X         As Long

        Dim Y         As Long
    
        Dim NPCIndex2 As Integer

        Dim Cuantos   As Long
    
        'Envio Palabras magicas, wavs y fxs.
108     If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
110         Call DecirPalabrasMagicas(h, UserIndex)

        End If
    
112     If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
    
114         If Hechizos(h).ParticleViaje > 0 Then
116             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
                
            Else
118             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))

            End If

        End If
    
120     If Hechizos(h).Particle > 0 Then 'Envio Particula?
122         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

        End If
    
124     If Hechizos(h).ParticleViaje = 0 Then
126         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, PosCasteadaX, PosCasteadaY))  'Esta linea faltaba. Pablo (ToxicWaste)

        End If

        Dim cuantosuser As Byte

        Dim nameuser    As String
       
128     Select Case Hechizos(h).AreaAfecta

            Case 1

130             For X = 1 To Hechizos(h).AreaRadio
132                 For Y = 1 To Hechizos(h).AreaRadio

134                     If MapData(UserList(UserIndex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex > 0 Then
136                         NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex

                            'If NPCIndex2 <> UserIndex Then
138                         If UserList(NPCIndex2).flags.Muerto = 0 Then
                                        
140                             AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
142                             cuantosuser = cuantosuser + 1
                                ' nameuser = nameuser & "," & NpcList(NPCIndex2).Name
                                            
                            End If

                            ' End If
                        End If

                    Next
                Next
                    
                ' If cuantosuser > 0 Then
                '     Call WriteConsoleMsg(UserIndex, "Has alcanzado a " & cuantosuser & " usuarios.", FontTypeNames.FONTTYPE_FIGHT)
                ' End If
144         Case 2

146             For X = 1 To Hechizos(h).AreaRadio
148                 For Y = 1 To Hechizos(h).AreaRadio

150                     If MapData(UserList(UserIndex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
152                         NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

154                         If NpcList(NPCIndex2).Attackable Then
156                             AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
158                             Cuantos = Cuantos + 1

                            End If

                        End If

                    Next
                Next
                
                ' If Cuantos > 0 Then
                '  Call WriteConsoleMsg(UserIndex, "Has alcanzado a " & Cuantos & " criaturas.", FontTypeNames.FONTTYPE_FIGHT)
                '  End If
160         Case 3

162             For X = 1 To Hechizos(h).AreaRadio
164                 For Y = 1 To Hechizos(h).AreaRadio

166                     If MapData(UserList(UserIndex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex > 0 Then
168                         NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex

                            'If NPCIndex2 <> UserIndex Then
170                         If UserList(NPCIndex2).flags.Muerto = 0 Then
172                             AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
174                             cuantosuser = cuantosuser + 1

                            End If

                            ' End If
                        End If
                            
176                     If MapData(UserList(UserIndex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
178                         NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

180                         If NpcList(NPCIndex2).Attackable Then
182                             AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
184                             Cuantos = Cuantos + 1
            
                            End If

                        End If
                            
                    Next
                Next
                
                ' If Cuantos > 0 Then
                '   Call WriteConsoleMsg(UserIndex, "Has alcanzado a " & Cuantos & " criaturas", FontTypeNames.FONTTYPE_FIGHT)
                ' End If
                'If cuantosuser > 0 Then
                ' Call WriteConsoleMsg(UserIndex, "Has alcanzado a a " & cuantosuser & " usuarios.", FontTypeNames.FONTTYPE_FIGHT)
                ' End If
        End Select

186     b = True

        
        Exit Sub

HechizoSobreArea_Err:
188     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoSobreArea", Erl)
190     Resume Next
        
End Sub

Sub HechizoPortal(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoPortal_Err
        

100     If UserList(UserIndex).flags.BattleModo = 1 Then
102         b = False
            'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
104         Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
        Else

            Dim PosCasteadaX As Byte

            Dim PosCasteadaY As Byte

            Dim PosCasteadaM As Integer

            Dim uh           As Integer

            Dim TempX        As Integer

            Dim TempY        As Integer

106         PosCasteadaX = UserList(UserIndex).flags.TargetX
108         PosCasteadaY = UserList(UserIndex).flags.TargetY
110         PosCasteadaM = UserList(UserIndex).flags.TargetMap
 
112         uh = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
            'Envio Palabras magicas, wavs y fxs.
   
114         If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or (MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).TileExit.Map > 0 Or UserList(UserIndex).flags.TargetUser <> 0 Then
116             b = False
                'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
118             Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)

            Else

120             If Hechizos(uh).TeleportX = 1 Then

122                 If UserList(UserIndex).flags.Portal = 0 Then

124                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, -1, False))
            
126                     UserList(UserIndex).flags.PortalM = UserList(UserIndex).Pos.Map
128                     UserList(UserIndex).flags.PortalX = UserList(UserIndex).flags.TargetX
130                     UserList(UserIndex).flags.PortalY = UserList(UserIndex).flags.TargetY
            
132                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 600, Accion_Barra.Intermundia))

134                     UserList(UserIndex).Accion.AccionPendiente = True
136                     UserList(UserIndex).Accion.Particula = ParticulasIndex.Runa
138                     UserList(UserIndex).Accion.TipoAccion = Accion_Barra.Intermundia
140                     UserList(UserIndex).Accion.HechizoPendiente = uh
            
142                     If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
144                         Call DecirPalabrasMagicas(uh, UserIndex)

                        End If

146                     b = True
                    Else
148                     Call WriteConsoleMsg(UserIndex, "No podés lanzar mas de un portal a la vez.", FontTypeNames.FONTTYPE_INFO)
150                     b = False

                    End If

                End If

            End If

        End If

        
        Exit Sub

HechizoPortal_Err:
152     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoPortal", Erl)
154     Resume Next
        
End Sub

Sub HechizoMaterializacion(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoMaterializacion_Err
        

        Dim h   As Integer

        Dim MAT As obj

100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
 
102     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Then
104         b = False
106         Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
        Else
108         MAT.Amount = Hechizos(h).MaterializaCant
110         MAT.ObjIndex = Hechizos(h).MaterializaObj
112         Call MakeObj(MAT, UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY)
            'Call WriteConsoleMsg(UserIndex, "Has materializado un objeto!!", FontTypeNames.FONTTYPE_INFO)
114         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
116         b = True

        End If

        
        Exit Sub

HechizoMaterializacion_Err:
118     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoMaterializacion", Erl)
120     Resume Next
        
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

                ' Call InvocarFamiliar(UserIndex, b)
        End Select

        'If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Or UserList(UserIndex).flags.TargetUser <> 0 Then
        '  b = False
        '  Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)

        'Else

122     If b Then
124         Call SubirSkill(UserIndex, magia)

130         UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

132         If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
134         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

136         If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
138         Call WriteUpdateMana(UserIndex)
140         Call WriteUpdateSta(UserIndex)

        End If

        
        Exit Sub

HandleHechizoTerreno_Err:
142     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HandleHechizoTerreno", Erl)
144     Resume Next
        
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
114         Call SubirSkill(UserIndex, magia)
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
138     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HandleHechizoUsuario", Erl)
140     Resume Next
        
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
110         Call SubirSkill(UserIndex, magia)
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
134     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HandleHechizoNPC", Erl)
136     Resume Next
        
End Sub

Sub LanzarHechizo(index As Integer, UserIndex As Integer)
        
        On Error GoTo LanzarHechizo_Err
        
100     If UserList(UserIndex).flags.EnConsulta Then
102         Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Dim uh As Integer
104         uh = UserList(UserIndex).Stats.UserHechizos(index)

106     If PuedeLanzar(UserIndex, uh, index) Then

108         Select Case Hechizos(uh).Target

                Case TargetType.uUsuarios

110                 If UserList(UserIndex).flags.TargetUser > 0 Then
112                     If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
114                         Call HandleHechizoUsuario(UserIndex, uh)
                    
116                         If Hechizos(uh).CoolDown > 0 Then
118                             UserList(UserIndex).Counters.UserHechizosInterval(index) = GetTickCount()

                            End If

                        Else
120                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
122                     Call WriteConsoleMsg(UserIndex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
124             Case TargetType.uNPC

126                 If UserList(UserIndex).flags.TargetNPC > 0 Then
128                     If Abs(NpcList(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
130                         Call HandleHechizoNPC(UserIndex, uh)

132                         If Hechizos(uh).CoolDown > 0 Then
134                             UserList(UserIndex).Counters.UserHechizosInterval(index) = GetTickCount()
                    
                            End If
                    
                        Else
136                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
138                     Call WriteConsoleMsg(UserIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
140             Case TargetType.uUsuariosYnpc

142                 If UserList(UserIndex).flags.TargetUser > 0 Then
144                     If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
146                         Call HandleHechizoUsuario(UserIndex, uh)
                    
148                         If Hechizos(uh).CoolDown > 0 Then
150                             UserList(UserIndex).Counters.UserHechizosInterval(index) = GetTickCount()

                            End If

                        Else
152                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

154                 ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then

156                     If Abs(NpcList(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
158                         If Hechizos(uh).CoolDown > 0 Then
160                             UserList(UserIndex).Counters.UserHechizosInterval(index) = GetTickCount()

                            End If

162                         Call HandleHechizoNPC(UserIndex, uh)
                        Else
164                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
166                     Call WriteConsoleMsg(UserIndex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
168             Case TargetType.uTerreno

170                 If Hechizos(uh).CoolDown > 0 Then
172                     UserList(UserIndex).Counters.UserHechizosInterval(index) = GetTickCount()

                    End If

174                 Call HandleHechizoTerreno(UserIndex, uh)

            End Select
    
        End If

176     If UserList(UserIndex).Counters.Trabajando Then
178         Call WriteMacroTrabajoToggle(UserIndex, False)

        End If

180     If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    
        
        Exit Sub

LanzarHechizo_Err:
182     Call RegistrarError(Err.Number, Err.Description, "modHechizos.LanzarHechizo", Erl)
184     Resume Next
        
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

            If UserList(UserIndex).flags.EnReto Then
                Call WriteConsoleMsg(UserIndex, "No podés lanzar invisibilidad durante un reto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
   
106         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
108             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
110             b = False
                Exit Sub

            End If
    
112         If UserList(tU).Counters.Saliendo Then
114             If UserIndex <> tU Then
116                 Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
118                 b = False
                    Exit Sub
                Else
120                 Call WriteConsoleMsg(UserIndex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
122                 b = False
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
124         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
126             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
128                 If esArmada(UserIndex) Then
130                     Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
132                     b = False
                        Exit Sub

                    End If

134                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
136                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
138                     b = False
                        Exit Sub
                    Else
140                     Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
142         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
144             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
            
146         If MapInfo(UserList(tU).Pos.Map).SinInviOcul Then
148             Call WriteConsoleMsg(UserIndex, "Una fuerza divina te impide usar invisibilidad en esta zona.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
150         If UserList(tU).flags.invisible = 1 Then
152             If tU = UserIndex Then
154                 Call WriteConsoleMsg(UserIndex, "¡Ya estás invisible!", FontTypeNames.FONTTYPE_INFO)
                Else
156                 Call WriteConsoleMsg(UserIndex, "¡El objetivo ya se encuentra invisible!", FontTypeNames.FONTTYPE_INFO)
                End If
158             b = False
                Exit Sub
            End If
   
160         UserList(tU).flags.invisible = 1
            'Ladder
            'Reseteamos el contador de Invisibilidad
162         UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
164         Call WriteContadores(tU)
166         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

168         Call InfoHechizo(UserIndex)
170         b = True

        End If
        
172     If Hechizos(h).Mimetiza = 1 Then

            If UserList(UserIndex).flags.EnReto Then
                Call WriteConsoleMsg(UserIndex, "No podés mimetizarte durante un reto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

174         If UserList(tU).flags.Muerto = 1 Then
                Exit Sub
            End If
            
176         If UserList(tU).flags.Navegando = 1 Then
                Exit Sub
            End If
178         If UserList(UserIndex).flags.Navegando = 1 Then
                Exit Sub
            End If
            
            'Si sos user, no uses este hechizo con GMS.
180         If Not EsGM(UserIndex) And EsGM(tU) Then Exit Sub
            
182         If UserList(UserIndex).flags.Mimetizado = 1 Then
184             Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no tuvo efecto", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
186         If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
            
            'copio el char original al mimetizado
            
188         With UserList(UserIndex)
190             .CharMimetizado.Body = .Char.Body
192             .CharMimetizado.Head = .Char.Head
194             .CharMimetizado.CascoAnim = .Char.CascoAnim
196             .CharMimetizado.ShieldAnim = .Char.ShieldAnim
198             .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                
200             .flags.Mimetizado = 1
                
                'ahora pongo local el del enemigo
202             .Char.Body = UserList(tU).Char.Body
204             .Char.Head = UserList(tU).Char.Head
206             .Char.CascoAnim = UserList(tU).Char.CascoAnim
208             .Char.ShieldAnim = UserList(tU).Char.ShieldAnim
210             .Char.WeaponAnim = UserList(tU).Char.WeaponAnim
212             .NameMimetizado = UserList(tU).name
214             If UserList(tU).GuildIndex > 0 Then .NameMimetizado = .NameMimetizado & " <" & modGuilds.GuildName(UserList(tU).GuildIndex) & ">"
            
216             Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
218             Call RefreshCharStatus(UserIndex)
            End With
           
220        Call InfoHechizo(UserIndex)
222        b = True
        End If

224     If Hechizos(h).Envenena > 0 Then
            ' If UserIndex = tU Then
            '    Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Exit Sub
            'End If
    
226         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
228         If UserIndex <> tU Then
230             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

232         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
234         Call InfoHechizo(UserIndex)
236         b = True

        End If

238     If Hechizos(h).desencantar = 1 Then
            ' Call WriteConsoleMsg(UserIndex, "Ningun efecto magico tiene efecto sobre ti ya.", FontTypeNames.FONTTYPE_INFOIAO)

240         UserList(UserIndex).flags.Envenenado = 0
242         UserList(UserIndex).flags.Incinerado = 0
    
244         If UserList(UserIndex).flags.Inmovilizado = 1 Then
246             UserList(UserIndex).Counters.Inmovilizado = 0
248             UserList(UserIndex).flags.Inmovilizado = 0
250             Call WriteInmovilizaOK(UserIndex)
            

            End If
    
252         If UserList(UserIndex).flags.Paralizado = 1 Then
254             UserList(UserIndex).flags.Paralizado = 0
256             Call WriteParalizeOK(UserIndex)
            
           
            End If
        
258         If UserList(UserIndex).flags.Ceguera = 1 Then
260             UserList(UserIndex).flags.Ceguera = 0
262             Call WriteBlindNoMore(UserIndex)
            

            End If
    
264         If UserList(UserIndex).flags.Maldicion = 1 Then
266             UserList(UserIndex).flags.Maldicion = 0
268             UserList(UserIndex).Counters.Maldicion = 0

            End If
    
270         Call InfoHechizo(UserIndex)
272         b = True

        End If

274     If Hechizos(h).incinera = 1 Then
276         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
278             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
280         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
282         If UserIndex <> tU Then
284             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

286         UserList(tU).flags.Incinerado = 1
288         Call InfoHechizo(UserIndex)
290         b = True

        End If

292     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
294         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
296             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
298             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
300         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
302             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
304                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
306                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
308                     b = False
                        Exit Sub

                    End If

310                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
312                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
314                     b = False
                        Exit Sub
                    Else

                        '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
316         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
318             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
320         UserList(tU).flags.Envenenado = 0
322         Call InfoHechizo(UserIndex)
324         b = True

        End If

326     If Hechizos(h).Maldicion = 1 Then
328         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
330             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
332         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
334         If UserIndex <> tU Then
336             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

338         UserList(tU).flags.Maldicion = 1
340         UserList(tU).Counters.Maldicion = 200
    
342         Call InfoHechizo(UserIndex)
344         b = True

        End If

346     If Hechizos(h).RemoverMaldicion = 1 Then
348         UserList(tU).flags.Maldicion = 0
350         Call InfoHechizo(UserIndex)
352         b = True

        End If

354     If Hechizos(h).GolpeCertero = 1 Then
356         UserList(tU).flags.GolpeCertero = 1
358         Call InfoHechizo(UserIndex)
360         b = True

        End If

362     If Hechizos(h).Bendicion = 1 Then
364         UserList(tU).flags.Bendicion = 1
366         Call InfoHechizo(UserIndex)
368         b = True

        End If

370     If Hechizos(h).Paraliza = 1 Then
372         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
374             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
376         If UserList(tU).flags.Paralizado = 1 Then
378             Call WriteConsoleMsg(UserIndex, UserList(tU).name & " ya está paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
380         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
382         If UserIndex <> tU Then
384             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
386         Call InfoHechizo(UserIndex)
388         b = True

390         If UserList(tU).Invent.ResistenciaEqpObjIndex = SUPERANILLO Then
392             Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
394             Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            
                Exit Sub

            End If
            
396         UserList(tU).Counters.Paralisis = Hechizos(h).Duration

398         If UserList(tU).flags.Paralizado = 0 Then
400             UserList(tU).flags.Paralizado = 1
402             Call WriteParalizeOK(tU)
404             Call WritePosUpdate(tU)
            End If

        End If

406     If Hechizos(h).Velocidad > 0 Then
408         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
410             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
412         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
414         If UserIndex <> tU Then
416             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
418         Call InfoHechizo(UserIndex)
420         b = True
                 
422         If UserList(tU).Counters.Velocidad = 0 Then
424             UserList(tU).flags.VelocidadBackup = UserList(tU).Char.speeding
            End If

426         UserList(tU).Char.speeding = Hechizos(h).Velocidad
428         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSpeedingACT(UserList(tU).Char.CharIndex, UserList(tU).Char.speeding))
            'End If
430         UserList(tU).Counters.Velocidad = Hechizos(h).Duration

        End If

432     If Hechizos(h).Inmoviliza = 1 Then
434         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
436             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
438         If UserList(tU).flags.Paralizado = 1 Then
440             Call WriteConsoleMsg(UserIndex, UserList(tU).name & " ya está paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
442         ElseIf UserList(tU).flags.Inmovilizado = 1 Then
444             Call WriteConsoleMsg(UserIndex, UserList(tU).name & " ya está inmovilizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
446         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
448         If UserIndex <> tU Then
450             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
            
452         Call InfoHechizo(UserIndex)
454         b = True
            '  If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
            '   Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            '
            '    Exit Sub
            ' End If
            
456         UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration

458         UserList(tU).flags.Inmovilizado = 1
460         Call WriteInmovilizaOK(tU)
462         Call WritePosUpdate(tU)
            

        End If

464     If Hechizos(h).RemoverParalisis = 1 Then
        
            'Para poder tirar remo a un pk en el ring
466         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
468             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
470                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
472                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
474                     b = False
                        Exit Sub

                    End If

476                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
478                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
480                     b = False
                        Exit Sub
                    Else
482                     Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
        
484         If UserList(tU).flags.Inmovilizado = 0 And UserList(tU).flags.Paralizado = 0 Then
486             Call WriteConsoleMsg(UserIndex, "El objetivo no esta paralizado.", FontTypeNames.FONTTYPE_INFO)
488             b = False
                Exit Sub

            End If
        
490         If UserList(tU).flags.Inmovilizado = 1 Then
492             UserList(tU).Counters.Inmovilizado = 0
494             UserList(tU).flags.Inmovilizado = 0
496             Call WriteInmovilizaOK(tU)
498             Call WritePosUpdate(tU)
                ' Call InfoHechizo(UserIndex)
            

                'b = True
            End If
    
500         If UserList(tU).flags.Paralizado = 1 Then
502             UserList(tU).flags.Paralizado = 0
                'no need to crypt this
504             Call WriteParalizeOK(tU)
            

                '  b = True
            End If

506         b = True
508         Call InfoHechizo(UserIndex)

        End If

510     If Hechizos(h).RemoverEstupidez = 1 Then
512         If UserList(tU).flags.Estupidez = 1 Then

                'Para poder tirar remo estu a un pk en el ring
514             If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
516                 If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
518                     If esArmada(UserIndex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
520                         Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
522                         b = False
                            Exit Sub

                        End If

524                     If UserList(UserIndex).flags.Seguro Then
                            'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
526                         Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
528                         b = False
                            Exit Sub
                        Else

                            ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                        End If

                    End If

                End If
    
530             UserList(tU).flags.Estupidez = 0
                'no need to crypt this
532             Call WriteDumbNoMore(tU)
            
534             Call InfoHechizo(UserIndex)
536             b = True

            End If

        End If

538     If Hechizos(h).Revivir = 1 Then
540         If UserList(tU).flags.Muerto = 1 Then

                If UserList(UserIndex).flags.EnReto Then
                    Call WriteConsoleMsg(UserIndex, "No podés revivir a nadie durante un reto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
    
                'No usar resu en mapas con ResuSinEfecto
                'If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
                '   Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                '   b = False
                '   Exit Sub
                ' End If
                
                If UserList(UserIndex).clase <> Cleric Then
                    Dim PuedeRevivir As Boolean
                    
                    If UserList(UserIndex).Invent.WeaponEqpObjIndex <> 0 Then
                        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Revive Then
                            PuedeRevivir = True
                        End If
                    End If
                    
                    If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex <> 0 Then
                        If ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).Revive Then
                            PuedeRevivir = True
                        End If
                    End If
                        
                    If Not PuedeRevivir Then
                        Call WriteConsoleMsg(UserIndex, "Necesitás un objeto con mayor poder mágico para poder revivir.", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    End If
                End If
                
                If UserList(tU).flags.SeguroResu Then
                    Call WriteConsoleMsg(UserIndex, "El usuario tiene el seguro de resurrección activado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(tU, UserList(UserIndex).name & " está intentando revivirte. Desactiva el seguro de resurrección para permitirle hacerlo.", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
        
542             If UserList(tU).Accion.TipoAccion = Accion_Barra.Resucitar Then
544                 Call WriteConsoleMsg(UserIndex, "El usuario ya esta siendo resucitado.", FontTypeNames.FONTTYPE_INFO)
546                 b = False
                    Exit Sub
                End If
        
                'Para poder tirar revivir a un pk en el ring
548             If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
550                 If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
552                     If esArmada(UserIndex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
554                         Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
556                         b = False
                            Exit Sub

                        End If

558                     If UserList(UserIndex).flags.Seguro Then
                            'call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
560                         Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
562                         b = False
                            Exit Sub
                        Else
564                         Call VolverCriminal(UserIndex)

                        End If

                    End If

                End If
                        
566             Call SendData(SendTarget.ToPCArea, tU, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, ParticulasIndex.Resucitar, 600, False))
568             Call SendData(SendTarget.ToPCArea, tU, PrepareMessageBarFx(UserList(tU).Char.CharIndex, 600, Accion_Barra.Resucitar))
570             UserList(tU).Accion.AccionPendiente = True
572             UserList(tU).Accion.Particula = ParticulasIndex.Resucitar
574             UserList(tU).Accion.TipoAccion = Accion_Barra.Resucitar
                
576             Call WriteUpdateHungerAndThirst(tU)
578             Call InfoHechizo(UserIndex)

580             b = True
                'Solo saco vida si es User. no quiero que exploten GMs por ahi.
        
                'Call RevivirUsuario(tU)
            Else
582             b = False

            End If

        End If

584     If Hechizos(h).Ceguera = 1 Then
586         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
588             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
590         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
592         If UserIndex <> tU Then
594             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

596         UserList(tU).flags.Ceguera = 1
598         UserList(tU).Counters.Ceguera = Hechizos(h).Duration

600         Call WriteBlind(tU)
        
602         Call InfoHechizo(UserIndex)
604         b = True

        End If

606     If Hechizos(h).Estupidez = 1 Then
608         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
610             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

612         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
614         If UserIndex <> tU Then
616             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

618         If UserList(tU).flags.Estupidez = 0 Then
620             UserList(tU).flags.Estupidez = 1
622             UserList(tU).Counters.Estupidez = Hechizos(h).Duration

            End If

624         Call WriteDumb(tU)
        

626         Call InfoHechizo(UserIndex)
628         b = True

        End If

        
        Exit Sub

HechizoEstadoUsuario_Err:
630     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoEstadoUsuario", Erl)
632     Resume Next
        
End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
        
        On Error GoTo HechizoEstadoNPC_Err
        

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 04/13/2008
        'Handles the Spells that afect the Stats of an NPC
        '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
        'removidos por users de su misma faccion.
        '***************************************************
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
124         Call InfoHechizo(UserIndex)
126         NpcList(NpcIndex).flags.Envenenado = 0
128         b = True

        End If

130     If Hechizos(hIndex).RemoverMaldicion = 1 Then
132         Call InfoHechizo(UserIndex)
            'NpcList(NpcIndex).flags.Maldicion = 0
134         b = True

        End If

136     If Hechizos(hIndex).Bendicion = 1 Then
138         Call InfoHechizo(UserIndex)
140         NpcList(NpcIndex).flags.Bendicion = 1
142         b = True

        End If

144     If Hechizos(hIndex).Paraliza = 1 Then
146         If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
148             If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
150                 b = False
                    Exit Sub

                End If

152             Call NPCAtacado(NpcIndex, UserIndex)
154             Call InfoHechizo(UserIndex)
156             NpcList(NpcIndex).flags.Paralizado = 1
158             NpcList(NpcIndex).flags.Inmovilizado = 0
160             NpcList(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 6

162             Call AnimacionIdle(NpcIndex, False)
                
164             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
166             Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)
168             b = False
                Exit Sub

            End If

        End If

170     If Hechizos(hIndex).RemoverParalisis = 1 Then
172         If NpcList(NpcIndex).flags.Paralizado = 1 Or NpcList(NpcIndex).flags.Inmovilizado = 1 Then
174             If NpcList(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
176                 If esArmada(UserIndex) Then
178                     Call InfoHechizo(UserIndex)
180                     NpcList(NpcIndex).flags.Paralizado = 0
182                     NpcList(NpcIndex).Contadores.Paralisis = 0
184                     b = True
                        Exit Sub
                    Else
186                     Call WriteConsoleMsg(UserIndex, "Solo podés Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
188                     b = False
                        Exit Sub

                    End If
                
190                 Call WriteConsoleMsg(UserIndex, "Solo podés Remover la Parálisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
192                 b = False
                    Exit Sub
                Else

194                 If NpcList(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
196                     If esCaos(UserIndex) Then
198                         Call InfoHechizo(UserIndex)
200                         NpcList(NpcIndex).flags.Paralizado = 0
202                         NpcList(NpcIndex).Contadores.Paralisis = 0
204                         b = True
                            Exit Sub
                        Else
206                         Call WriteConsoleMsg(UserIndex, "Solo podés Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
208                         b = False
                            Exit Sub

                        End If

                    End If

                End If

            Else
210             Call WriteConsoleMsg(UserIndex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFO)
212             b = False
                Exit Sub

            End If

        End If
 
214     If Hechizos(hIndex).Inmoviliza = 1 And NpcList(NpcIndex).flags.Paralizado = 0 Then
216         If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
218             If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
220                 b = False
                    Exit Sub

                End If

222             Call NPCAtacado(NpcIndex, UserIndex)
224             NpcList(NpcIndex).flags.Inmovilizado = 1
226             NpcList(NpcIndex).flags.Paralizado = 0
228             NpcList(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 6

230             Call AnimacionIdle(NpcIndex, True)

232             Call InfoHechizo(UserIndex)
234             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
236             Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

238     If Hechizos(hIndex).Mimetiza = 1 Then

            If UserList(UserIndex).flags.EnReto Then
                Call WriteConsoleMsg(UserIndex, "No podés mimetizarte durante un reto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
    
240         If UserList(UserIndex).flags.Mimetizado = 1 Then
242             Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no tuvo efecto", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
244         If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
            
                
246         If UserList(UserIndex).clase = eClass.Druid Then
                'copio el char original al mimetizado
248             With UserList(UserIndex)
250                 .CharMimetizado.Body = .Char.Body
252                 .CharMimetizado.Head = .Char.Head
254                 .CharMimetizado.CascoAnim = .Char.CascoAnim
256                 .CharMimetizado.ShieldAnim = .Char.ShieldAnim
258                 .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                    
260                 .flags.Mimetizado = 1
                    
                    'ahora pongo lo del NPC.
262                 .Char.Body = NpcList(NpcIndex).Char.Body
264                 .Char.Head = NpcList(NpcIndex).Char.Head
266                 .Char.CascoAnim = NingunCasco
268                 .Char.ShieldAnim = NingunEscudo
270                 .Char.WeaponAnim = NingunArma
272                 .NameMimetizado = IIf(NpcList(NpcIndex).showName = 1, NpcList(NpcIndex).name, vbNullString)

274                 Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
276                 Call RefreshCharStatus(UserIndex)
                End With
            Else
278             Call WriteConsoleMsg(UserIndex, "Solo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
280        Call InfoHechizo(UserIndex)
282        b = True
        End If
        
        Exit Sub

HechizoEstadoNPC_Err:
284     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoEstadoNPC", Erl)
286     Resume Next
        
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
102         Daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
            'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        
104         Call InfoHechizo(UserIndex)
106         NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MinHp + Daño

108         If NpcList(NpcIndex).Stats.MinHp > NpcList(NpcIndex).Stats.MaxHp Then NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MaxHp

109         DañoStr = PonerPuntos(Daño)

            'Call WriteConsoleMsg(UserIndex, "Has curado " & Daño & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
110         Call WriteLocaleMsg(UserIndex, "388", FontTypeNames.FONTTYPE_FIGHT, "la criatura¬" & DañoStr)

112         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageTextOverChar(DañoStr, NpcList(NpcIndex).Char.CharIndex, vbGreen))
114         b = True
        
116     ElseIf Hechizos(hIndex).SubeHP = 2 Then

118         If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
120             b = False
                Exit Sub
            End If
        
122         Call NPCAtacado(NpcIndex, UserIndex)
124         Daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
        
126         Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)

            ' Si al hechizo le afecta el daño mágico
127         If Hechizos(hIndex).StaffAffected Then
                ' Daño mágico arma
128             If UserList(UserIndex).clase = eClass.Mage Then
                    ' El mago tiene un 30% de daño reducido
129                 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
130                     Daño = Porcentaje(Daño, 70 + ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                    Else
131                     Daño = Daño * 0.7
                    End If
                Else
132                 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
133                     Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                    End If
                End If
                
                ' Daño mágico anillo
134             If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex > 0 Then
136                 Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).MagicDamageBonus)
                End If
            End If

140         b = True
        
142         If NpcList(NpcIndex).flags.Snd2 > 0 Then
144             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd2, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y))
            End If
        
            'Quizas tenga defenza magica el NPC.
146         If Hechizos(hIndex).AntiRm = 0 Then
148             Daño = Daño - NpcList(NpcIndex).Stats.defM
            End If
        
150         If Daño < 0 Then Daño = 0
        
152         NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MinHp - Daño
154         Call InfoHechizo(UserIndex)

            If NpcList(NpcIndex).NPCtype = DummyTarget Then
                NpcList(NpcIndex).Contadores.UltimoAtaque = 30
            End If
            
            DañoStr = PonerPuntos(Daño)
        
156         If UserList(UserIndex).ChatCombate = 1 Then
                'Call WriteConsoleMsg(UserIndex, "Le has causado " & Daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
158             Call WriteLocaleMsg(UserIndex, "389", FontTypeNames.FONTTYPE_FIGHT, "la criatura¬" & DañoStr)
            End If
        
            If NpcList(NpcIndex).MaestroUser <= 0 Then
160             Call CalcularDarExp(UserIndex, NpcIndex, Daño)
            End If
    
162         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageTextOverChar(DañoStr, NpcList(NpcIndex).Char.CharIndex, vbRed))
    
164         If NpcList(NpcIndex).Stats.MinHp < 1 Then
166             NpcList(NpcIndex).Stats.MinHp = 0
168             Call MuereNpc(NpcIndex, UserIndex)
            End If

        End If

        
        Exit Sub

HechizoPropNPC_Err:
170     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoPropNPC", Erl)
172     Resume Next
        
End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)
        
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
188                 Call WriteConsoleMsg(UserIndex, "HecMSGU*" & h & "*" & UserList(UserList(UserIndex).flags.TargetUser).name, FontTypeNames.FONTTYPE_FIGHT)
190                 Call WriteConsoleMsg(UserList(UserIndex).flags.TargetUser, "HecMSGA*" & h & "*" & UserList(UserIndex).name, FontTypeNames.FONTTYPE_FIGHT)
    
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
200     Call RegistrarError(Err.Number, Err.Description, "modHechizos.InfoHechizo", Erl)
202     Resume Next
        
End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoPropUsuario_Err
        

        Dim h As Integer

        Dim Daño As Integer
        
        Dim DañoStr As String

        Dim tempChr           As Integer
    
102     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
104     tempChr = UserList(UserIndex).flags.TargetUser
      
        'Hambre
106     If Hechizos(h).SubeHam = 1 Then
    
108         Call InfoHechizo(UserIndex)
    
110         Daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
112         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + Daño

114         If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
116         If UserIndex <> tempChr Then
118             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
120             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
122             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
124         Call WriteUpdateHungerAndThirst(tempChr)
126         b = True
    
128     ElseIf Hechizos(h).SubeHam = 2 Then

130         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
132         If UserIndex <> tempChr Then
134             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
            Else
                Exit Sub

            End If
    
136         Call InfoHechizo(UserIndex)
    
138         Daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
140         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - Daño
    
142         If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
144         If UserIndex <> tempChr Then
146             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
148             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
150             Call WriteConsoleMsg(UserIndex, "Te has quitado " & Daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
152         Call WriteUpdateHungerAndThirst(tempChr)
    
154         b = True
    
156         If UserList(tempChr).Stats.MinHam < 1 Then
158             UserList(tempChr).Stats.MinHam = 0
160             UserList(tempChr).flags.Hambre = 1

            End If
    
        End If

        'Sed
162     If Hechizos(h).SubeSed = 1 Then
    
164         Call InfoHechizo(UserIndex)
    
166         Daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
168         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + Daño

170         If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
172         If UserIndex <> tempChr Then
174             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
176             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
178             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
180         Call WriteUpdateHungerAndThirst(tempChr)
    
182         b = True
    
184     ElseIf Hechizos(h).SubeSed = 2 Then
    
186         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
188         If UserIndex <> tempChr Then
190             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
192         Call InfoHechizo(UserIndex)
    
194         Daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
196         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - Daño
    
198         If UserIndex <> tempChr Then
200             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
202             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
204             Call WriteConsoleMsg(UserIndex, "Te has quitado " & Daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
206         If UserList(tempChr).Stats.MinAGU < 1 Then
208             UserList(tempChr).Stats.MinAGU = 0
210             UserList(tempChr).flags.Sed = 1

            End If
            
212         Call WriteUpdateHungerAndThirst(tempChr)
    
214         b = True

        End If

        ' <-------- Agilidad ---------->
216     If Hechizos(h).SubeAgilidad = 1 Then
    
            'Para poder tirar cl a un pk en el ring
218         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
220             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
222                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
224                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
226                     b = False
                        Exit Sub

                    End If

228                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
230                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
232                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
234         Call InfoHechizo(UserIndex)
236         Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
238         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

240         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + Daño, UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2)

242         UserList(tempChr).flags.TomoPocion = True
244         b = True
246         Call WriteFYA(tempChr)
    
248     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
250         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
252         If UserIndex <> tempChr Then
254             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
256         Call InfoHechizo(UserIndex)
    
258         UserList(tempChr).flags.TomoPocion = True
260         Daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
262         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

264         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - Daño < MINATRIBUTOS Then
266             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
            Else
268             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - Daño

            End If
    
270         b = True
272         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
274     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
276         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
278             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
280                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
282                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
284                     b = False
                        Exit Sub

                    End If

286                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
288                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
290                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
292         Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
294         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
296         UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + Daño, UserList(tempChr).Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2)

298         UserList(tempChr).flags.TomoPocion = True
            
300         Call WriteFYA(tempChr)

302         b = True
    
304         Call InfoHechizo(UserIndex)
306         Call WriteFYA(tempChr)

308     ElseIf Hechizos(h).SubeFuerza = 2 Then

310         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
312         If UserIndex <> tempChr Then
314             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
316         UserList(tempChr).flags.TomoPocion = True
    
318         Daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
320         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

322         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - Daño < MINATRIBUTOS Then
324             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
            Else
326             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - Daño

            End If

328         b = True
330         Call InfoHechizo(UserIndex)
332         Call WriteFYA(tempChr)

        End If

        'Salud
334     If Hechizos(h).SubeHP = 1 Then
    
            'Verifica que el usuario no este muerto
336         If UserList(tempChr).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
338             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
340             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar a un pk en el ring
342         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
344             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
346                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
348                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
350                     b = False
                        Exit Sub

                    End If

352                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
354                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
356                     b = False
                        Exit Sub
                    Else

                        'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
       
358         Daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            ' daño = daño + Porcentaje(daño, 2 * UserList(UserIndex).Stats.ELV)
    
360         Call InfoHechizo(UserIndex)

362         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + Daño

364         If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp

            DañoStr = PonerPuntos(Daño)

366         If UserIndex <> tempChr Then
                'Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
368             Call WriteLocaleMsg(UserIndex, "388", FontTypeNames.FONTTYPE_FIGHT, UserList(tempChr).name & "¬" & DañoStr)

                'Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
370             Call WriteLocaleMsg(tempChr, "32", FontTypeNames.FONTTYPE_FIGHT, UserList(UserIndex).name & "¬" & DañoStr)
            Else
                'Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
372             Call WriteLocaleMsg(UserIndex, "33", FontTypeNames.FONTTYPE_FIGHT, DañoStr)
            End If
    
374         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageTextCharDrop(DañoStr, UserList(tempChr).Char.CharIndex, vbGreen))
376         Call WriteUpdateHP(tempChr)
    
378         b = True

380     ElseIf Hechizos(h).SubeHP = 2 Then
    
382         If UserIndex = tempChr Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
384             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

386         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
388         Daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
390         Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)

            ' Si al hechizo le afecta el daño mágico
392         If Hechizos(h).StaffAffected Then
                ' Daño mágico arma
394             If UserList(UserIndex).clase = eClass.Mage Then
                    ' El mago tiene un 30% de daño reducido
396                 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
397                     Daño = Porcentaje(Daño, 70 + ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                    Else
398                     Daño = Daño * 0.7
                    End If
                Else
399                 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
400                     Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                    End If
                End If
                
                ' Daño mágico anillo
402             If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex > 0 Then
403                 Daño = Daño + Porcentaje(Daño, ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).MagicDamageBonus)
                End If
            End If
            
            ' Si el hechizo no ignora la RM
404         If Hechizos(h).AntiRm = 0 Then
                Dim PorcentajeRM As Integer

                ' Resistencia mágica armadura
406             If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
408                 PorcentajeRM = PorcentajeRM + ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).ResistenciaMagica
                End If
                
                ' Resistencia mágica anillo
410             If UserList(tempChr).Invent.ResistenciaEqpObjIndex > 0 Then
412                 PorcentajeRM = PorcentajeRM + ObjData(UserList(tempChr).Invent.ResistenciaEqpObjIndex).ResistenciaMagica
                End If
                
                ' Resistencia mágica escudo
414             If UserList(tempChr).Invent.EscudoEqpObjIndex > 0 Then
416                 PorcentajeRM = PorcentajeRM + ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).ResistenciaMagica
                End If
                
                ' Resistencia mágica casco
418             If UserList(tempChr).Invent.CascoEqpObjIndex > 0 Then
420                 PorcentajeRM = PorcentajeRM + ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).ResistenciaMagica
                End If

                ' Resto el porcentaje total
                Daño = Daño - Porcentaje(Daño, PorcentajeRM)
            End If

            ' Prevengo daño negativo
422         If Daño < 0 Then Daño = 0
    
424         If UserIndex <> tempChr Then
426             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
            End If
    
428         Call InfoHechizo(UserIndex)
    
430         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - Daño

431         DañoStr = PonerPuntos(Daño)
    
            'Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
432         Call WriteLocaleMsg(UserIndex, "389", FontTypeNames.FONTTYPE_FIGHT, UserList(tempChr).name & "¬" & DañoStr)

            'Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
434         Call WriteLocaleMsg(tempChr, "34", FontTypeNames.FONTTYPE_FIGHT, UserList(UserIndex).name & "¬" & DañoStr)
    
436         Call SubirSkill(tempChr, Resistencia)
    
438         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageTextCharDrop(Daño, UserList(tempChr).Char.CharIndex, vbRed))

            'Muere
440         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
442             Call Statistics.StoreFrag(UserIndex, tempChr)
444             Call ContarMuerte(tempChr, UserIndex)
446             Call ActStats(tempChr, UserIndex)
            Else
448             Call WriteUpdateHP(tempChr)
            End If

    
450         b = True

        End If

        'Mana
452     If Hechizos(h).SubeMana = 1 Then
    
454         Call InfoHechizo(UserIndex)
456         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + Daño

458         If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
460         If UserIndex <> tempChr Then
462             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
464             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
466             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
468         Call WriteUpdateMana(tempChr)
    
470         b = True
    
472     ElseIf Hechizos(h).SubeMana = 2 Then

474         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
476         If UserIndex <> tempChr Then
478             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
480         Call InfoHechizo(UserIndex)
    
482         If UserIndex <> tempChr Then
484             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
486             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
488             Call WriteConsoleMsg(UserIndex, "Te has quitado " & Daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
490         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - Daño

492         If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0

494         Call WriteUpdateMana(tempChr)

496         b = True
    
        End If

        'Stamina
498     If Hechizos(h).SubeSta = 1 Then
500         Call InfoHechizo(UserIndex)
502         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + Daño

504         If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta

506         If UserIndex <> tempChr Then
508             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
510             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
512             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
514         Call WriteUpdateSta(tempChr)

516         b = True
518     ElseIf Hechizos(h).SubeSta = 2 Then

520         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
522         If UserIndex <> tempChr Then
524             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
526         Call InfoHechizo(UserIndex)
    
528         If UserIndex <> tempChr Then
530             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
532             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
534             Call WriteConsoleMsg(UserIndex, "Te has quitado " & Daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
536         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - Daño
    
538         If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0

540         Call WriteUpdateSta(tempChr)

542         b = True

        End If

        Exit Sub

HechizoPropUsuario_Err:
548     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoPropUsuario", Erl)
550     Resume Next
        
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
256             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
258             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
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
            End If

            ' Prevengo daño negativo
306         If Daño < 0 Then Daño = 0
    
308         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
310         If UserIndex <> tempChr Then
312             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
314         enviarInfoHechizo = True
    
316         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - Daño
    
318         Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
320         Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
322         Call SubirSkill(tempChr, Resistencia)
324         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageTextOverChar(Daño, UserList(tempChr).Char.CharIndex, vbRed))

            'Muere
326         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
328             Call Statistics.StoreFrag(UserIndex, tempChr)
        
330             Call ContarMuerte(tempChr, UserIndex)
332             Call ActStats(tempChr, UserIndex)

                'Call UserDie(tempChr)
            End If
    
334         b = True

        End If

        Dim tU As Integer

336     tU = tempChr

338     If Hechizos(h).Invisibilidad = 1 Then
   
340         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
342             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
344             b = False
                Exit Sub

            End If
    
346         If UserList(tU).Counters.Saliendo Then
348             If UserIndex <> tU Then
350                 Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
352                 b = False
                    Exit Sub
                Else
354                 Call WriteConsoleMsg(UserIndex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
356                 b = False
                    Exit Sub

                End If

            End If
    
            'Para poder tirar invi a un pk en el ring
358         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
360             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
362                 If esArmada(UserIndex) Then
364                     Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
366                     b = False
                        Exit Sub

                    End If

368                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
370                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
372                     b = False
                        Exit Sub
                    Else
374                     Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
376         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
378             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
   
380         UserList(tU).flags.invisible = 1
            'Ladder
            'Reseteamos el contador de Invisibilidad
382         UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
384         Call WriteContadores(tU)
386         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

388         enviarInfoHechizo = True
390         b = True

        End If

392     If Hechizos(h).Envenena > 0 Then
            ' If UserIndex = tU Then
            '    Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Exit Sub
            'End If
    
394         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
396         If UserIndex <> tU Then
398             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

400         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
402         enviarInfoHechizo = True
404         b = True

        End If

406     If Hechizos(h).desencantar = 1 Then
408         Call WriteConsoleMsg(UserIndex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)

410         UserList(UserIndex).flags.Envenenado = 0
412         UserList(UserIndex).flags.Incinerado = 0
    
414         If UserList(UserIndex).flags.Inmovilizado = 1 Then
416             UserList(UserIndex).Counters.Inmovilizado = 0
418             UserList(UserIndex).flags.Inmovilizado = 0
420             Call WriteInmovilizaOK(UserIndex)
            

            End If
    
422         If UserList(UserIndex).flags.Paralizado = 1 Then
424             UserList(UserIndex).flags.Paralizado = 0
426             Call WriteParalizeOK(UserIndex)
            
           
            End If
        
428         If UserList(UserIndex).flags.Ceguera = 1 Then
430             UserList(UserIndex).flags.Ceguera = 0
432             Call WriteBlindNoMore(UserIndex)
            

            End If
    
434         If UserList(UserIndex).flags.Maldicion = 1 Then
436             UserList(UserIndex).flags.Maldicion = 0
438             UserList(UserIndex).Counters.Maldicion = 0

            End If
    
440         enviarInfoHechizo = True
442         b = True

        End If

444     If Hechizos(h).Sanacion = 1 Then

446         UserList(tU).flags.Envenenado = 0
448         UserList(tU).flags.Incinerado = 0
450         enviarInfoHechizo = True
452         b = True

        End If

454     If Hechizos(h).incinera = 1 Then
456         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
458             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
460         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
462         If UserIndex <> tU Then
464             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

466         UserList(tU).flags.Incinerado = 1
468         enviarInfoHechizo = True
470         b = True

        End If

472     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
474         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
476             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
478             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
480         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
482             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
484                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
486                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
488                     b = False
                        Exit Sub

                    End If

490                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
492                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
494                     b = False
                        Exit Sub
                    Else

                        '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
496         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
498             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
500         UserList(tU).flags.Envenenado = 0
502         enviarInfoHechizo = True
504         b = True

        End If

506     If Hechizos(h).Maldicion = 1 Then
508         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
510             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
512         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
514         If UserIndex <> tU Then
516             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

518         UserList(tU).flags.Maldicion = 1
520         UserList(tU).Counters.Maldicion = 200
    
522         enviarInfoHechizo = True
524         b = True

        End If

526     If Hechizos(h).RemoverMaldicion = 1 Then
528         UserList(tU).flags.Maldicion = 0
530         enviarInfoHechizo = True
532         b = True

        End If

534     If Hechizos(h).GolpeCertero = 1 Then
536         UserList(tU).flags.GolpeCertero = 1
538         enviarInfoHechizo = True
540         b = True

        End If

542     If Hechizos(h).Bendicion = 1 Then
544         UserList(tU).flags.Bendicion = 1
546         enviarInfoHechizo = True
548         b = True

        End If

550     If Hechizos(h).Paraliza = 1 Then
552         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
554             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
556         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
558         If UserIndex <> tU Then
560             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
562         enviarInfoHechizo = True
564         b = True

566         If UserList(tU).Invent.ResistenciaEqpObjIndex = SUPERANILLO Then
568             Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
570             Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            
                Exit Sub

            End If
            
572         UserList(tU).Counters.Paralisis = Hechizos(h).Duration

574         If UserList(tU).flags.Paralizado = 0 Then
576             UserList(tU).flags.Paralizado = 1
578             Call WriteParalizeOK(tU)
580             Call WritePosUpdate(tU)
            End If

        End If

582     If Hechizos(h).Inmoviliza = 1 Then
584         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
586             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
588         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
590         If UserIndex <> tU Then
592             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
594         enviarInfoHechizo = True
596         b = True
            
598         UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration

600         If UserList(tU).flags.Inmovilizado = 0 Then
602             UserList(tU).flags.Inmovilizado = 1
604             Call WriteInmovilizaOK(tU)
606             Call WritePosUpdate(tU)
            

            End If

        End If

608     If Hechizos(h).RemoverParalisis = 1 Then
        
            'Para poder tirar remo a un pk en el ring
610         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
612             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
614                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
616                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
618                     b = False
                        Exit Sub

                    End If

620                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
622                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
624                     b = False
                        Exit Sub
                    Else
626                     Call VolverCriminal(UserIndex)

                    End If

                End If
            
            End If

628         If UserList(tU).flags.Inmovilizado = 1 Then
630             UserList(tU).Counters.Inmovilizado = 0
632             UserList(tU).flags.Inmovilizado = 0
634             Call WriteInmovilizaOK(tU)
636             enviarInfoHechizo = True
            
638             b = True

            End If

640         If UserList(tU).flags.Paralizado = 1 Then
642             UserList(tU).flags.Paralizado = 0
                'no need to crypt this
644             Call WriteParalizeOK(tU)
646             enviarInfoHechizo = True
            
648             b = True

            End If

        End If

650     If Hechizos(h).Ceguera = 1 Then
652         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
654             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
656         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
658         If UserIndex <> tU Then
660             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

662         UserList(tU).flags.Ceguera = 1
664         UserList(tU).Counters.Ceguera = Hechizos(h).Duration

666         Call WriteBlind(tU)
        
668         enviarInfoHechizo = True
670         b = True

        End If

672     If Hechizos(h).Estupidez = 1 Then
674         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
676             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

678         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
680         If UserIndex <> tU Then
682             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

684         If UserList(tU).flags.Estupidez = 0 Then
686             UserList(tU).flags.Estupidez = 1
688             UserList(tU).Counters.Estupidez = Hechizos(h).Duration

            End If

690         Call WriteDumb(tU)
        

692         enviarInfoHechizo = True
694         b = True

        End If

696     If Hechizos(h).Velocidad > 0 Then

698         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
700         If UserIndex <> tU Then
702             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
704         enviarInfoHechizo = True
706         b = True
            
708         If UserList(tU).Counters.Velocidad = 0 Then
710             UserList(tU).flags.VelocidadBackup = UserList(tU).Char.speeding

            End If

712         UserList(tU).Char.speeding = Hechizos(h).Velocidad
714         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSpeedingACT(UserList(tU).Char.CharIndex, UserList(tU).Char.speeding))
            
716         UserList(tU).Counters.Velocidad = Hechizos(h).Duration

        End If

718     If enviarInfoHechizo Then
720         Call InfoHechizo(UserIndex)

        End If

    

        
        Exit Sub

HechizoCombinados_Err:
722     Call RegistrarError(Err.Number, Err.Description, "modHechizos.HechizoCombinados", Erl)
724     Resume Next
        
End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal slot As Byte)
        
        On Error GoTo UpdateUserHechizos_Err
        

        'Call LogTarea("Sub UpdateUserHechizos")

        Dim LoopC As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(UserIndex).Stats.UserHechizos(slot) > 0 Then
104             Call ChangeUserHechizo(UserIndex, slot, UserList(UserIndex).Stats.UserHechizos(slot))
            Else
106             Call ChangeUserHechizo(UserIndex, slot, 0)

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
118     Call RegistrarError(Err.Number, Err.Description, "modHechizos.UpdateUserHechizos", Erl)
120     Resume Next
        
End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Hechizo As Integer)
        
        On Error GoTo ChangeUserHechizo_Err
        

        'Call LogTarea("ChangeUserHechizo")
    
100     UserList(UserIndex).Stats.UserHechizos(slot) = Hechizo
    
102     If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
104         Call WriteChangeSpellSlot(UserIndex, slot)
        Else
106         Call WriteChangeSpellSlot(UserIndex, slot)

        End If

        
        Exit Sub

ChangeUserHechizo_Err:
108     Call RegistrarError(Err.Number, Err.Description, "modHechizos.ChangeUserHechizo", Erl)
110     Resume Next
        
End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)
        
        On Error GoTo DesplazarHechizo_Err
        

100     If (Dire <> 1 And Dire <> -1) Then Exit Sub
102     If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

        Dim TempHechizo As Integer

104     If Dire = 1 Then 'Mover arriba
106         If CualHechizo = 1 Then
108             Call WriteConsoleMsg(UserIndex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
110             TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
112             UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
114             UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
116             If UserList(UserIndex).flags.Hechizo > 0 Then
118                 UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo - 1

                End If

            End If

        Else 'mover abajo

120         If CualHechizo = MAXUSERHECHIZOS Then
122             Call WriteConsoleMsg(UserIndex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
124             TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
126             UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
128             UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
130             If UserList(UserIndex).flags.Hechizo > 0 Then
132                 UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo + 1

                End If

            End If

        End If

        
        Exit Sub

DesplazarHechizo_Err:
134     Call RegistrarError(Err.Number, Err.Description, "modHechizos.DesplazarHechizo", Erl)
136     Resume Next
        
End Sub

Sub AreaHechizo(UserIndex As Integer, NpcIndex As Integer, X As Byte, Y As Byte, npc As Boolean)
        
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
144                 Call WriteConsoleMsg(UserIndex, "Le has causado " & Daño & " puntos de daño a " & NpcList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)

                End If
            
146             Call CalcularDarExp(UserIndex, NpcIndex, Daño)
                
148             If NpcList(NpcIndex).Stats.MinHp <= 0 Then
                    'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + NpcList(NpcIndex).GiveEXP
                    'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + NpcList(NpcIndex).GiveGLD
150                 Call MuereNpc(NpcIndex, UserIndex)
                End If

                Exit Sub

            End If

        Else

152         TilesDifNpc = UserList(NpcIndex).Pos.X + UserList(NpcIndex).Pos.Y
154         tilDif = TilesDifUser - TilesDifNpc

156         If Hechizos(h2).SubeHP = 2 Then
158             If UserIndex = NpcIndex Then
                    Exit Sub
                End If

160             If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
                
162             If UserIndex <> NpcIndex Then
164                 Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

                End If
                
166             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

168             Hit = Hit + Porcentaje(Hit, 3 * UserList(UserIndex).Stats.ELV)
            
                ' Daño mágico arma
170             If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
172                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                End If
                
                ' Daño mágico anillo
174             If UserList(UserIndex).Invent.DañoMagicoEqpObjIndex > 0 Then
176                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.DañoMagicoEqpObjIndex).MagicDamageBonus)
                End If

178             If tilDif <> 0 Then
180                 porcentajeDesc = Abs(tilDif) * 20
182                 Daño = Hit / 100 * porcentajeDesc
184                 Daño = Hit - Daño
                Else
186                 Daño = Hit
                End If
                
                ' Si el hechizo no ignora la RM
188             If Hechizos(h2).AntiRm = 0 Then
                    ' Resistencia mágica armadura
190                 If UserList(NpcIndex).Invent.ArmourEqpObjIndex > 0 Then
192                     Daño = Daño - Porcentaje(Daño, ObjData(UserList(NpcIndex).Invent.ArmourEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica anillo
194                 If UserList(NpcIndex).Invent.ResistenciaEqpObjIndex > 0 Then
196                     Daño = Daño - Porcentaje(Daño, ObjData(UserList(NpcIndex).Invent.ResistenciaEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica escudo
198                 If UserList(NpcIndex).Invent.EscudoEqpObjIndex > 0 Then
200                     Daño = Daño - Porcentaje(Daño, ObjData(UserList(NpcIndex).Invent.EscudoEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica casco
202                 If UserList(NpcIndex).Invent.CascoEqpObjIndex > 0 Then
204                     Daño = Daño - Porcentaje(Daño, ObjData(UserList(NpcIndex).Invent.CascoEqpObjIndex).ResistenciaMagica)
                    End If
                End If
                
                ' Prevengo daño negativo
206             If Daño < 0 Then Daño = 0

208             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp - Daño
                    
210             Call WriteConsoleMsg(UserIndex, "Le has quitado " & Daño & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
212             Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).name & " te ha quitado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
214             Call SubirSkill(NpcIndex, Resistencia)
216             Call WriteUpdateUserStats(NpcIndex)
                
                'Muere
218             If UserList(NpcIndex).Stats.MinHp < 1 Then
                    'Store it!
220                 Call Statistics.StoreFrag(UserIndex, NpcIndex)
                        
222                 Call ContarMuerte(NpcIndex, UserIndex)
224                 Call ActStats(NpcIndex, UserIndex)

                    'Call UserDie(NpcIndex)
                End If

            End If
                
226         If Hechizos(h2).SubeHP = 1 Then
228             If (TriggerZonaPelea(UserIndex, NpcIndex) <> TRIGGER6_PERMITE) Then
230                 If Status(UserIndex) = 1 And Status(NpcIndex) <> 1 Then
                        Exit Sub

                    End If

                End If

232             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

234             If tilDif <> 0 Then
236                 porcentajeDesc = Abs(tilDif) * 20
238                 Daño = Hit / 100 * porcentajeDesc
240                 Daño = Hit - Daño
                Else
242                 Daño = Hit

                End If
 
244             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp + Daño

246             If UserList(NpcIndex).Stats.MinHp > UserList(NpcIndex).Stats.MaxHp Then UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MaxHp

            End If
 
248         If UserIndex <> NpcIndex Then
250             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & Daño & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
252             Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).name & " te ha restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
254             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & Daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
                    
256         Call WriteUpdateUserStats(NpcIndex)

        End If
                
258     If Hechizos(h2).Envenena > 0 Then
260         If UserIndex = NpcIndex Then
                Exit Sub

            End If
                    
262         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
                
264         If UserIndex <> NpcIndex Then
266             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                    
268         UserList(NpcIndex).flags.Envenenado = Hechizos(h2).Envenena
270         Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).name & " te ha envenenado.", FontTypeNames.FONTTYPE_FIGHT)

        End If
                
272     If Hechizos(h2).Paraliza = 1 Then
274         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
276         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
278         If UserIndex <> NpcIndex Then
280             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
            
282         Call WriteConsoleMsg(NpcIndex, "Has sido paralizado.", FontTypeNames.FONTTYPE_INFO)
284         UserList(NpcIndex).Counters.Paralisis = Hechizos(h2).Duration

286         If UserList(NpcIndex).flags.Paralizado = 0 Then
288             UserList(NpcIndex).flags.Paralizado = 1
290             Call WriteParalizeOK(NpcIndex)
            

            End If
            
        End If
                
292     If Hechizos(h2).Inmoviliza = 1 Then
294         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
296         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
298         If UserIndex <> NpcIndex Then
300             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                    
302         Call WriteConsoleMsg(NpcIndex, "Has sido inmovilizado.", FontTypeNames.FONTTYPE_INFO)
304         UserList(NpcIndex).Counters.Inmovilizado = Hechizos(h2).Duration

306         If UserList(NpcIndex).flags.Inmovilizado = 0 Then
308             UserList(NpcIndex).flags.Inmovilizado = 1
310             Call WriteInmovilizaOK(NpcIndex)
312             Call WritePosUpdate(NpcIndex)
            
            End If

        End If
                
314     If Hechizos(h2).Ceguera = 1 Then
316         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
318         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
320         If UserIndex <> NpcIndex Then
322             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                    
324         UserList(NpcIndex).flags.Ceguera = 1
326         UserList(NpcIndex).Counters.Ceguera = Hechizos(h2).Duration
328         Call WriteConsoleMsg(NpcIndex, "Te han cegado.", FontTypeNames.FONTTYPE_INFO)
            
330         Call WriteBlind(NpcIndex)
        

        End If
                
332     If Hechizos(h2).Velocidad > 0 Then
    
334         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
336         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
338         If UserIndex <> NpcIndex Then
340             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If

342         If UserList(NpcIndex).Counters.Velocidad = 0 Then
344             UserList(NpcIndex).flags.VelocidadBackup = UserList(NpcIndex).Char.speeding

            End If

346         UserList(NpcIndex).Char.speeding = Hechizos(h2).Velocidad
348         Call SendData(SendTarget.ToPCArea, NpcIndex, PrepareMessageSpeedingACT(UserList(NpcIndex).Char.CharIndex, UserList(NpcIndex).Char.speeding))
350         UserList(NpcIndex).Counters.Velocidad = Hechizos(h2).Duration

        End If
                
352     If Hechizos(h2).Maldicion = 1 Then
354         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
356         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
358         If UserIndex <> NpcIndex Then
360             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If

362         Call WriteConsoleMsg(NpcIndex, "Ahora estas maldito. No podras Atacar", FontTypeNames.FONTTYPE_INFO)
364         UserList(NpcIndex).flags.Maldicion = 1
366         UserList(NpcIndex).Counters.Maldicion = Hechizos(h2).Duration

        End If
                
368     If Hechizos(h2).RemoverMaldicion = 1 Then
370         Call WriteConsoleMsg(NpcIndex, "Te han removido la maldicion.", FontTypeNames.FONTTYPE_INFO)
372         UserList(NpcIndex).flags.Maldicion = 0

        End If
                
374     If Hechizos(h2).GolpeCertero = 1 Then
376         Call WriteConsoleMsg(NpcIndex, "Tu proximo golpe sera certero.", FontTypeNames.FONTTYPE_INFO)
378         UserList(NpcIndex).flags.GolpeCertero = 1

        End If
                
380     If Hechizos(h2).Bendicion = 1 Then
382         Call WriteConsoleMsg(NpcIndex, "Has sido bendecido.", FontTypeNames.FONTTYPE_INFO)
384         UserList(NpcIndex).flags.Bendicion = 1

        End If
                  
386     If Hechizos(h2).incinera = 1 Then
388         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
390         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
392         If UserIndex <> NpcIndex Then
394             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If

396         UserList(NpcIndex).flags.Incinerado = 1
398         Call WriteConsoleMsg(NpcIndex, "Has sido Incinerado.", FontTypeNames.FONTTYPE_INFO)

        End If
                
400     If Hechizos(h2).Invisibilidad = 1 Then
402         Call WriteConsoleMsg(NpcIndex, "Ahora sos invisible.", FontTypeNames.FONTTYPE_INFO)
404         UserList(NpcIndex).flags.invisible = 1
406         UserList(NpcIndex).Counters.Invisibilidad = Hechizos(h2).Duration
408         Call WriteContadores(NpcIndex)
410         Call SendData(SendTarget.ToPCArea, NpcIndex, PrepareMessageSetInvisible(UserList(NpcIndex).Char.CharIndex, True))

        End If
                              
412     If Hechizos(h2).Sanacion = 1 Then
414         Call WriteConsoleMsg(NpcIndex, "Has sido sanado.", FontTypeNames.FONTTYPE_INFO)
416         UserList(NpcIndex).flags.Envenenado = 0
418         UserList(NpcIndex).flags.Incinerado = 0

        End If
                
420     If Hechizos(h2).RemoverParalisis = 1 Then
422         Call WriteConsoleMsg(NpcIndex, "Has sido removido.", FontTypeNames.FONTTYPE_INFO)

424         If UserList(NpcIndex).flags.Inmovilizado = 1 Then
426             UserList(NpcIndex).Counters.Inmovilizado = 0
428             UserList(NpcIndex).flags.Inmovilizado = 0
430             Call WriteInmovilizaOK(NpcIndex)
            

            End If

432         If UserList(NpcIndex).flags.Paralizado = 1 Then
434             UserList(NpcIndex).flags.Paralizado = 0
                'no need to crypt this
436             Call WriteParalizeOK(NpcIndex)
            

            End If

        End If
                
438     If Hechizos(h2).desencantar = 1 Then
440         Call WriteConsoleMsg(NpcIndex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)
                    
442         UserList(NpcIndex).flags.Envenenado = 0
444         UserList(NpcIndex).flags.Incinerado = 0
                    
446         If UserList(NpcIndex).flags.Inmovilizado = 1 Then
448             UserList(NpcIndex).Counters.Inmovilizado = 0
450             UserList(NpcIndex).flags.Inmovilizado = 0
452             Call WriteInmovilizaOK(NpcIndex)
            

            End If
                    
454         If UserList(NpcIndex).flags.Paralizado = 1 Then
456             UserList(NpcIndex).flags.Paralizado = 0
458             Call WriteParalizeOK(NpcIndex)
            
                       
            End If
                    
460         If UserList(NpcIndex).flags.Ceguera = 1 Then
462             UserList(NpcIndex).flags.Ceguera = 0
464             Call WriteBlindNoMore(NpcIndex)
            

            End If
                    
466         If UserList(NpcIndex).flags.Maldicion = 1 Then
468             UserList(NpcIndex).flags.Maldicion = 0
470             UserList(NpcIndex).Counters.Maldicion = 0

            End If

        End If
        
        
        Exit Sub

AreaHechizo_Err:
472     Call RegistrarError(Err.Number, Err.Description, "modHechizos.AreaHechizo", Erl)
474     Resume Next
        
End Sub
