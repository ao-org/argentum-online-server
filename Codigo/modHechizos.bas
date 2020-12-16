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

Public Const HELEMENTAL_FUEGO  As Integer = 26

Public Const HELEMENTAL_TIERRA As Integer = 28

Public Const SUPERANILLO       As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByVal Spell As Integer, Optional ByVal IgnoreVisibilityCheck As Boolean = False)
        'Guardia caos
        
        On Error GoTo NpcLanzaSpellSobreUser_Err
        
100     With UserList(Userindex)

            If Spell = 0 Then Exit Sub
        
            '¿NPC puede ver a través de la invisibilidad?
102         If Not IgnoreVisibilityCheck Then
104             If .flags.invisible = 1 Or .flags.Oculto = 1 Or .flags.Inmunidad = 1 Then Exit Sub
            End If

            'Npclist(NpcIndex).CanAttack = 0
            Dim daño As Integer

106         If Hechizos(Spell).SubeHP = 1 Then

108             daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
110             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
112             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

114             .Stats.MinHp = .Stats.MinHp + daño

116             If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
    
118             Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
120             Call WriteUpdateHP(Userindex)
122             Call SubirSkill(Userindex, Resistencia)

124         ElseIf Hechizos(Spell).SubeHP = 2 Then
        
126             daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
        
128             If .Invent.CascoEqpObjIndex > 0 Then
130                 daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

                End If
        
132             If .Invent.AnilloEqpObjIndex > 0 Then
134                 daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

                End If
        
136             If daño < 0 Then daño = 0
        
138             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
140             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
142             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageEfectOverHead(daño, .Char.CharIndex))
                
144             If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?

146                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))
          
                End If

148             .Stats.MinHp = .Stats.MinHp - daño
        
150             Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
152             Call SubirSkill(Userindex, Resistencia)
        
                'Muere
154             If .Stats.MinHp < 1 Then
156                 .Stats.MinHp = 0
                    'If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                    'restarCriminalidad (UserIndex)
                    ' End If
158                 Call UserDie(Userindex)
                    '[Barrin 1-12-03]
                    ' If Npclist(NpcIndex).MaestroUser > 0 Then
                    'Store it!
                    'Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, UserIndex)
                
                    'Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
                    ' Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
                    '  End If
                    '[/Barrin]
                End If
            
160             Call WriteUpdateHP(Userindex)
    
162         ElseIf Hechizos(Spell).Paraliza = 1 Then

164             If .flags.Paralizado = 0 Then
166                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
168                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

170                 .flags.Paralizado = 1
172                 .Counters.Paralisis = Hechizos(Spell).Duration / 2
          
174                 Call WriteParalizeOK(Userindex)
176                 Call WritePosUpdate(Userindex)

                End If

178         ElseIf Hechizos(Spell).incinera = 1 Then
180             Debug.Print "incinerar"

182             If .flags.Incinerado = 0 Then
184                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))

186                 If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
188                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))

                    End If

190                 .flags.Incinerado = 1
192                 Call WriteConsoleMsg(Userindex, "Has sido incinerado por " & Npclist(NpcIndex).name & ".", FontTypeNames.FONTTYPE_FIGHT)

                End If

            End If
    
        End With

        Exit Sub

NpcLanzaSpellSobreUser_Err:
194     Call RegistrarError(Err.Number, Err.description & " Hechizo: " & Spell, "modHechizos.NpcLanzaSpellSobreUser", Erl)

196     Resume Next
        
End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
        'solo hechizos ofensivos!
        
        On Error GoTo NpcLanzaSpellSobreNpc_Err
        

100     If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub

102     Npclist(NpcIndex).CanAttack = 0

        Dim daño As Integer

104     If Hechizos(Spell).SubeHP = 2 Then
    
106         daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
108         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, Npclist(TargetNPC).Pos.X, Npclist(TargetNPC).Pos.Y))
110         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
112         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageEfectOverHead(daño, Npclist(TargetNPC).Char.CharIndex))
        
114         Npclist(TargetNPC).Stats.MinHp = Npclist(TargetNPC).Stats.MinHp - daño

            ' Mascotas dan experiencia al amo
116         If Npclist(NpcIndex).MaestroUser > 0 Then
118             Call CalcularDarExp(Npclist(NpcIndex).MaestroUser, TargetNPC, daño)
            End If
        
            'Muere
120         If Npclist(TargetNPC).Stats.MinHp < 1 Then
122             Npclist(TargetNPC).Stats.MinHp = 0
                ' If Npclist(NpcIndex).MaestroUser > 0 Then
                '  Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
                '  Else
124             Call MuereNpc(TargetNPC, 0)

                '  End If
            End If
    
        End If
    
        
        Exit Sub

NpcLanzaSpellSobreNpc_Err:
126     Call RegistrarError(Err.Number, Err.description, "modHechizos.NpcLanzaSpellSobreNpc", Erl)
128     Resume Next
        
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal Userindex As Integer) As Boolean

        On Error GoTo ErrHandler
    
        Dim j As Integer

100     For j = 1 To MAXUSERHECHIZOS

102         If UserList(Userindex).Stats.UserHechizos(j) = i Then
104             TieneHechizo = True
                Exit Function

            End If

        Next

        Exit Function
ErrHandler:

End Function

Sub AgregarHechizo(ByVal Userindex As Integer, ByVal slot As Integer)
        
        On Error GoTo AgregarHechizo_Err
        

        Dim hIndex As Integer

        Dim j      As Integer

100     hIndex = ObjData(UserList(Userindex).Invent.Object(slot).ObjIndex).HechizoIndex

102     If Not TieneHechizo(hIndex, Userindex) Then

            'Buscamos un slot vacio
104         For j = 1 To MAXUSERHECHIZOS

106             If UserList(Userindex).Stats.UserHechizos(j) = 0 Then Exit For
108         Next j
        
110         If UserList(Userindex).Stats.UserHechizos(j) <> 0 Then
112             Call WriteConsoleMsg(Userindex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
            Else
114             UserList(Userindex).Stats.UserHechizos(j) = hIndex
116             Call UpdateUserHechizos(False, Userindex, CByte(j))
                'Quitamos del inv el item
118             Call QuitarUserInvItem(Userindex, CByte(slot), 1)

            End If

        Else
120         Call WriteConsoleMsg(Userindex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

AgregarHechizo_Err:
122     Call RegistrarError(Err.Number, Err.description, "modHechizos.AgregarHechizo", Erl)
124     Resume Next
        
End Sub
            
Sub DecirPalabrasMagicas(ByVal Hechizo As Byte, ByVal Userindex As Integer)

        On Error Resume Next

100     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(Userindex).Char.CharIndex, vbCyan))
        Exit Sub

End Sub

Function PuedeLanzar(ByVal Userindex As Integer, ByVal HechizoIndex As Integer, Optional ByVal slot As Integer = 0) As Boolean
        
        On Error GoTo PuedeLanzar_Err
        

100     If UserList(Userindex).flags.Muerto = 0 Then

            Dim wp2 As WorldPos

102         wp2.Map = UserList(Userindex).flags.TargetMap
104         wp2.X = UserList(Userindex).flags.TargetX
106         wp2.Y = UserList(Userindex).flags.TargetY
    
108         If Hechizos(HechizoIndex).NecesitaObj > 0 Then
110             If TieneObjEnInv(Userindex, Hechizos(HechizoIndex).NecesitaObj, Hechizos(HechizoIndex).NecesitaObj2) Then
112                 PuedeLanzar = True
                    'Exit Function
               
                Else
114                 Call WriteConsoleMsg(Userindex, "Necesitas un " & ObjData(Hechizos(HechizoIndex).NecesitaObj).name & " para lanzar el hechizo.", FontTypeNames.FONTTYPE_INFO)
116                 PuedeLanzar = False
                    Exit Function

                End If

            End If
    
118         If Hechizos(HechizoIndex).CoolDown > 0 Then

                Dim actual            As Long

                Dim segundosFaltantes As Long

120             actual = GetTickCount()

122             If UserList(Userindex).Counters.UserHechizosInterval(slot) + (Hechizos(HechizoIndex).CoolDown * 1000) > actual Then
124                 segundosFaltantes = Int((UserList(Userindex).Counters.UserHechizosInterval(slot) + (Hechizos(HechizoIndex).CoolDown * 1000) - actual) / 1000)
126                 Call WriteConsoleMsg(Userindex, "Debes esperar " & segundosFaltantes & " segundos para volver a tirar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
128                 PuedeLanzar = False
                    Exit Function
                Else
130                 PuedeLanzar = True

                End If

            End If
    
132         If UserList(Userindex).Stats.MinHp > Hechizos(HechizoIndex).RequiredHP Then
134             If UserList(Userindex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
136                 If UserList(Userindex).Stats.UserSkills(eSkill.magia) >= Hechizos(HechizoIndex).MinSkill Then
138                     If UserList(Userindex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
140                         PuedeLanzar = True
                        Else
142                         Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
144                         PuedeLanzar = False

                        End If
                    
                    Else
146                     Call WriteConsoleMsg(Userindex, "No tenes suficientes puntos de magia para lanzar este hechizo, necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos.", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteLocaleMsg(UserIndex, "221", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
148                     PuedeLanzar = False

                    End If

                Else
150                 Call WriteLocaleMsg(Userindex, "222", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "No tenes suficiente mana. Necesitas " & Hechizos(HechizoIndex).ManaRequerido & " puntos de mana.", FontTypeNames.FONTTYPE_INFO)
152                 PuedeLanzar = False

                End If

            Else
154             Call WriteConsoleMsg(Userindex, "No tenes suficiente vida. Necesitas " & Hechizos(HechizoIndex).RequiredHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
156             PuedeLanzar = False

            End If

        Else
            'Call WriteConsoleMsg(UserIndex, "No podes lanzar hechizos porque estas muerto.", FontTypeNames.FONTTYPE_INFO)
158         Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
160         PuedeLanzar = False

        End If

        
        Exit Function

PuedeLanzar_Err:
162     Call RegistrarError(Err.Number, Err.description, "modHechizos.PuedeLanzar", Erl)
164     Resume Next
        
End Function

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HechizoInvocacion(ByVal Userindex As Integer, ByRef b As Boolean)
        '***************************************************
        'Author: Uknown
        'Modification: 06/15/2008 (NicoNZ)
        'Last modification: 01/12/2020 (WyroX)
        'Sale del sub si no hay una posición valida.
        '***************************************************
    
        On Error GoTo HechizoInvocacion_Err
    
100     With UserList(Userindex)
    
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
116                 Call WriteConsoleMsg(Userindex, "Para invocar el fuego fatuo no debes tener otras criaturas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
                'No permitimos se invoquen criaturas en zonas seguras
118             If MapInfo(.Pos.Map).Seguro = 1 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
120                 Call WriteConsoleMsg(Userindex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
122             For j = 1 To Hechizos(h).cant
                
124                 If .NroMascotas < MAXMASCOTAS Then
126                     ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False)
128                     If ind > 0 Then
130                         .NroMascotas = .NroMascotas + 1
                        
132                         index = FreeMascotaIndex(Userindex)
                        
134                         .MascotasIndex(index) = ind
136                         .MascotasType(index) = Npclist(ind).Numero
                        
138                         Npclist(ind).MaestroUser = Userindex
140                         Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
142                         Npclist(ind).GiveGLD = 0
                        
144                         Call FollowAmo(ind)
                        Else
                            Exit Sub
                        End If
                        
                    Else
                        Exit For
                    End If
                
146             Next j
            
148             Call InfoHechizo(Userindex)
150             b = True
        
152         ElseIf Hechizos(h).Invoca = 2 Then
            
                ' Si tiene mascotas
154             If .NroMascotas > 0 Then
                    ' Tiene que estar en zona insegura
156                 If MapInfo(.Pos.Map).Seguro = 1 Then

                        Dim i As Integer
                    
                        ' Si no están guardadas las mascotas
158                     If .flags.MascotasGuardadas = 0 Then
160                         For i = 1 To MAXMASCOTAS
162                             If .MascotasIndex(i) > 0 Then
                                    ' Si no es un elemental, lo "guardamos"... lo matamos
164                                 If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
                                        ' Le saco el maestro, para que no me lo quite de mis mascotas
166                                     Npclist(.MascotasIndex(i)).MaestroUser = 0
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
180                                 .MascotasIndex(i) = SpawnNpc(.MascotasType(i), TargetPos, True, True)

182                                 Npclist(.MascotasIndex(i)).MaestroUser = Userindex
184                                 Call FollowAmo(.MascotasIndex(i))
                                
186                                 b = True
                                End If
                            Next
                        
188                         .flags.MascotasGuardadas = 0
                        End If
                
                    Else
190                     Call WriteConsoleMsg(Userindex, "No puedes invocar tus mascotas en un mapa seguro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
            
                Else
192                 Call WriteConsoleMsg(Userindex, "No tienes mascotas.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

194             If b Then Call InfoHechizo(Userindex)
            
            End If
    
        End With
    
        Exit Sub
    
HechizoInvocacion_Err:
196     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoTerrenoEstado")
198     Resume Next

End Sub

Sub HechizoTerrenoEstado(ByVal Userindex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoTerrenoEstado_Err
        

        Dim PosCasteadaX As Integer

        Dim PosCasteadaY As Integer

        Dim PosCasteadaM As Integer

        Dim h            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     PosCasteadaX = UserList(Userindex).flags.TargetX
102     PosCasteadaY = UserList(Userindex).flags.TargetY
104     PosCasteadaM = UserList(Userindex).flags.TargetMap
    
106     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
        
108     If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
110         b = True

            'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
112         For TempX = PosCasteadaX - 11 To PosCasteadaX + 11
114             For TempY = PosCasteadaY - 11 To PosCasteadaY + 11

116                 If InMapBounds(PosCasteadaM, TempX, TempY) Then
118                     If MapData(PosCasteadaM, TempX, TempY).Userindex > 0 Then

                            'hay un user
120                         If UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.NoDetectable = 0 Then
122                             UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.invisible = 0
124                             Call WriteConsoleMsg(MapData(PosCasteadaM, TempX, TempY).Userindex, "Tu invisibilidad ya no tiene efecto.", FontTypeNames.FONTTYPE_INFOIAO)
126                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).Char.CharIndex, False))

                            End If

                        End If

                    End If

128             Next TempY
130         Next TempX
    
132         Call InfoHechizo(Userindex)

        End If

        
        Exit Sub

HechizoTerrenoEstado_Err:
134     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoTerrenoEstado", Erl)
136     Resume Next
        
End Sub

Sub HechizoSobreArea(ByVal Userindex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoSobreArea_Err
        

        Dim PosCasteadaX As Byte

        Dim PosCasteadaY As Byte

        Dim PosCasteadaM As Integer

        Dim h            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     PosCasteadaX = UserList(Userindex).flags.TargetX
102     PosCasteadaY = UserList(Userindex).flags.TargetY
104     PosCasteadaM = UserList(Userindex).flags.TargetMap
 
106     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)

        Dim X         As Long

        Dim Y         As Long
    
        Dim NPCIndex2 As Integer

        Dim Cuantos   As Long
    
        'Envio Palabras magicas, wavs y fxs.
108     If UserList(Userindex).flags.NoPalabrasMagicas = 0 Then
110         Call DecirPalabrasMagicas(h, Userindex)

        End If
    
112     If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
    
114         If Hechizos(h).ParticleViaje > 0 Then
116             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXWithDestinoXY(UserList(Userindex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))
                
            Else
118             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))

            End If

        End If
    
120     If Hechizos(h).Particle > 0 Then 'Envio Particula?
122         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXToFloor(UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

        End If
    
124     If Hechizos(h).ParticleViaje = 0 Then
126         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(h).wav, PosCasteadaX, PosCasteadaY))  'Esta linea faltaba. Pablo (ToxicWaste)

        End If

        Dim cuantosuser As Byte

        Dim nameuser    As String
       
128     Select Case Hechizos(h).AreaAfecta

            Case 1

130             For X = 1 To Hechizos(h).AreaRadio
132                 For Y = 1 To Hechizos(h).AreaRadio

134                     If MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).Userindex > 0 Then
136                         NPCIndex2 = MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).Userindex

                            'If NPCIndex2 <> UserIndex Then
138                         If UserList(NPCIndex2).flags.Muerto = 0 Then
                                        
140                             AreaHechizo Userindex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
142                             cuantosuser = cuantosuser + 1
                                ' nameuser = nameuser & "," & Npclist(NPCIndex2).Name
                                            
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

150                     If MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
152                         NPCIndex2 = MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

154                         If Npclist(NPCIndex2).Attackable Then
156                             AreaHechizo Userindex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
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

166                     If MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).Userindex > 0 Then
168                         NPCIndex2 = MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).Userindex

                            'If NPCIndex2 <> UserIndex Then
170                         If UserList(NPCIndex2).flags.Muerto = 0 Then
172                             AreaHechizo Userindex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
174                             cuantosuser = cuantosuser + 1

                            End If

                            ' End If
                        End If
                            
176                     If MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
178                         NPCIndex2 = MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

180                         If Npclist(NPCIndex2).Attackable Then
182                             AreaHechizo Userindex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
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
188     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoSobreArea", Erl)
190     Resume Next
        
End Sub

Sub HechizoPortal(ByVal Userindex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoPortal_Err
        

100     If UserList(Userindex).flags.BattleModo = 1 Then
102         b = False
            'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
104         Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFO)
        Else

            Dim PosCasteadaX As Byte

            Dim PosCasteadaY As Byte

            Dim PosCasteadaM As Integer

            Dim uh           As Integer

            Dim TempX        As Integer

            Dim TempY        As Integer

106         PosCasteadaX = UserList(Userindex).flags.TargetX
108         PosCasteadaY = UserList(Userindex).flags.TargetY
110         PosCasteadaM = UserList(Userindex).flags.TargetMap
 
112         uh = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
    
            'Envio Palabras magicas, wavs y fxs.
   
114         If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).ObjInfo.Amount > 0 Or (MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Or MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).TileExit.Map > 0 Or UserList(Userindex).flags.TargetUser <> 0 Then
116             b = False
                'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
118             Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFO)

            Else

120             If Hechizos(uh).TeleportX = 1 Then

122                 If UserList(Userindex).flags.Portal = 0 Then

124                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, -1, False))
            
126                     UserList(Userindex).flags.PortalM = UserList(Userindex).Pos.Map
128                     UserList(Userindex).flags.PortalX = UserList(Userindex).flags.TargetX
130                     UserList(Userindex).flags.PortalY = UserList(Userindex).flags.TargetY
            
132                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 600, Accion_Barra.Intermundia))

134                     UserList(Userindex).Accion.AccionPendiente = True
136                     UserList(Userindex).Accion.Particula = ParticulasIndex.Runa
138                     UserList(Userindex).Accion.TipoAccion = Accion_Barra.Intermundia
140                     UserList(Userindex).Accion.HechizoPendiente = uh
            
142                     If UserList(Userindex).flags.NoPalabrasMagicas = 0 Then
144                         Call DecirPalabrasMagicas(uh, Userindex)

                        End If

146                     b = True
                    Else
148                     Call WriteConsoleMsg(Userindex, "No podés lanzar mas de un portal a la vez.", FontTypeNames.FONTTYPE_INFO)
150                     b = False

                    End If

                End If

            End If

        End If

        
        Exit Sub

HechizoPortal_Err:
152     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoPortal", Erl)
154     Resume Next
        
End Sub

Sub HechizoMaterializacion(ByVal Userindex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoMaterializacion_Err
        

        Dim h   As Integer

        Dim MAT As obj

100     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
 
102     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).Blocked Then
104         b = False
106         Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
        Else
108         MAT.Amount = Hechizos(h).MaterializaCant
110         MAT.ObjIndex = Hechizos(h).MaterializaObj
112         Call MakeObj(MAT, UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY)
            'Call WriteConsoleMsg(UserIndex, "Has materializado un objeto!!", FontTypeNames.FONTTYPE_INFO)
114         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXToFloor(UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
116         b = True

        End If

        
        Exit Sub

HechizoMaterializacion_Err:
118     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoMaterializacion", Erl)
120     Resume Next
        
End Sub

Sub HandleHechizoTerreno(ByVal Userindex As Integer, ByVal uh As Integer)
        
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
102             Call HechizoInvocacion(Userindex, b)

104         Case TipoHechizo.uEstado 'Tipo 2
106             Call HechizoTerrenoEstado(Userindex, b)

108         Case TipoHechizo.uMaterializa 'Tipo 3
110             Call HechizoMaterializacion(Userindex, b)
            
112         Case TipoHechizo.uArea 'Tipo 5
114             Call HechizoSobreArea(Userindex, b)
            
116         Case TipoHechizo.uPortal 'Tipo 6
118             Call HechizoPortal(Userindex, b)

120         Case TipoHechizo.UFamiliar

                ' Call InvocarFamiliar(UserIndex, b)
        End Select

        'If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Or UserList(UserIndex).flags.TargetUser <> 0 Then
        '  b = False
        '  Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)

        'Else

122     If b Then
124         Call SubirSkill(Userindex, magia)

            'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
126         If UserList(Userindex).clase = eClass.Druid And UserList(Userindex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
128             UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido * 0.7
            Else
130             UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido

            End If

132         If UserList(Userindex).Stats.MinMAN < 0 Then UserList(Userindex).Stats.MinMAN = 0
134         UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Hechizos(uh).StaRequerido

136         If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
138         Call WriteUpdateMana(Userindex)
140         Call WriteUpdateSta(Userindex)

        End If

        
        Exit Sub

HandleHechizoTerreno_Err:
142     Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoTerreno", Erl)
144     Resume Next
        
End Sub

Sub HandleHechizoUsuario(ByVal Userindex As Integer, ByVal uh As Integer)
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
102             Call HechizoEstadoUsuario(Userindex, b)

104         Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropUsuario(Userindex, b)

108         Case TipoHechizo.uCombinados
110             Call HechizoCombinados(Userindex, b)
    
        End Select

112     If b Then
114         Call SubirSkill(Userindex, magia)
            'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
116         UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
118         If UserList(Userindex).Stats.MinMAN < 0 Then UserList(Userindex).Stats.MinMAN = 0

120         If Hechizos(uh).RequiredHP > 0 Then
122             UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp - Hechizos(uh).RequiredHP
124             If UserList(Userindex).Stats.MinHp < 0 Then UserList(Userindex).Stats.MinHp = 1
126             Call WriteUpdateHP(Userindex)
            End If

128         UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Hechizos(uh).StaRequerido
130         If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0

132         Call WriteUpdateMana(Userindex)
134         Call WriteUpdateSta(Userindex)
136         UserList(Userindex).flags.TargetUser = 0

        End If

        
        Exit Sub

HandleHechizoUsuario_Err:
138     Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoUsuario", Erl)
140     Resume Next
        
End Sub

Sub HandleHechizoNPC(ByVal Userindex As Integer, ByVal uh As Integer)
        
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
102             Call HechizoEstadoNPC(UserList(Userindex).flags.TargetNPC, uh, b, Userindex)

104         Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropNPC(uh, UserList(Userindex).flags.TargetNPC, Userindex, b)

        End Select

108     If b Then
110         Call SubirSkill(Userindex, magia)
112         UserList(Userindex).flags.TargetNPC = 0
114         UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido

116         If Hechizos(uh).RequiredHP > 0 Then
118             If UserList(Userindex).Stats.MinMAN < 0 Then UserList(Userindex).Stats.MinMAN = 0
120             UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp - Hechizos(uh).RequiredHP
122             Call WriteUpdateHP(Userindex)
            End If

124         If UserList(Userindex).Stats.MinHp < 0 Then UserList(Userindex).Stats.MinHp = 1
126         UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Hechizos(uh).StaRequerido

128         If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
130         Call WriteUpdateMana(Userindex)
132         Call WriteUpdateSta(Userindex)

        End If

        
        Exit Sub

HandleHechizoNPC_Err:
134     Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoNPC", Erl)
136     Resume Next
        
End Sub

Sub LanzarHechizo(index As Integer, Userindex As Integer)
        
        On Error GoTo LanzarHechizo_Err
        

        Dim uh As Integer

100     uh = UserList(Userindex).Stats.UserHechizos(index)

102     If PuedeLanzar(Userindex, uh, index) Then

104         Select Case Hechizos(uh).Target

                Case TargetType.uUsuarios

106                 If UserList(Userindex).flags.TargetUser > 0 Then
108                     If Abs(UserList(UserList(Userindex).flags.TargetUser).Pos.Y - UserList(Userindex).Pos.Y) <= RANGO_VISION_Y Then
110                         Call HandleHechizoUsuario(Userindex, uh)
                    
112                         If Hechizos(uh).CoolDown > 0 Then
114                             UserList(Userindex).Counters.UserHechizosInterval(index) = GetTickCount()

                            End If

                        Else
116                         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
118                     Call WriteConsoleMsg(Userindex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
120             Case TargetType.uNPC

122                 If UserList(Userindex).flags.TargetNPC > 0 Then
124                     If Abs(Npclist(UserList(Userindex).flags.TargetNPC).Pos.Y - UserList(Userindex).Pos.Y) <= RANGO_VISION_Y Then
126                         Call HandleHechizoNPC(Userindex, uh)

128                         If Hechizos(uh).CoolDown > 0 Then
130                             UserList(Userindex).Counters.UserHechizosInterval(index) = GetTickCount()
                    
                            End If
                    
                        Else
132                         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
134                     Call WriteConsoleMsg(Userindex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
136             Case TargetType.uUsuariosYnpc

138                 If UserList(Userindex).flags.TargetUser > 0 Then
140                     If Abs(UserList(UserList(Userindex).flags.TargetUser).Pos.Y - UserList(Userindex).Pos.Y) <= RANGO_VISION_Y Then
142                         Call HandleHechizoUsuario(Userindex, uh)
                    
144                         If Hechizos(uh).CoolDown > 0 Then
146                             UserList(Userindex).Counters.UserHechizosInterval(index) = GetTickCount()

                            End If

                        Else
148                         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

150                 ElseIf UserList(Userindex).flags.TargetNPC > 0 Then

152                     If Abs(Npclist(UserList(Userindex).flags.TargetNPC).Pos.Y - UserList(Userindex).Pos.Y) <= RANGO_VISION_Y Then
154                         If Hechizos(uh).CoolDown > 0 Then
156                             UserList(Userindex).Counters.UserHechizosInterval(index) = GetTickCount()

                            End If

158                         Call HandleHechizoNPC(Userindex, uh)
                        Else
160                         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
162                     Call WriteConsoleMsg(Userindex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
164             Case TargetType.uTerreno

166                 If Hechizos(uh).CoolDown > 0 Then
168                     UserList(Userindex).Counters.UserHechizosInterval(index) = GetTickCount()

                    End If

170                 Call HandleHechizoTerreno(Userindex, uh)

            End Select
    
        End If

172     If UserList(Userindex).Counters.Trabajando Then
174         Call WriteMacroTrabajoToggle(Userindex, False)

        End If

176     If UserList(Userindex).Counters.Ocultando Then UserList(Userindex).Counters.Ocultando = UserList(Userindex).Counters.Ocultando - 1
    
        
        Exit Sub

LanzarHechizo_Err:
178     Call RegistrarError(Err.Number, Err.description, "modHechizos.LanzarHechizo", Erl)
180     Resume Next
        
End Sub

Sub HechizoEstadoUsuario(ByVal Userindex As Integer, ByRef b As Boolean)
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

100     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
102     tU = UserList(Userindex).flags.TargetUser

104     If Hechizos(h).Invisibilidad = 1 Then
   
106         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
108             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
110             b = False
                Exit Sub

            End If
    
112         If UserList(tU).Counters.Saliendo Then
114             If Userindex <> tU Then
116                 Call WriteConsoleMsg(Userindex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
118                 b = False
                    Exit Sub
                Else
120                 Call WriteConsoleMsg(Userindex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
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
124         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
126             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
128                 If esArmada(Userindex) Then
130                     Call WriteConsoleMsg(Userindex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
132                     b = False
                        Exit Sub

                    End If

134                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
136                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
138                     b = False
                        Exit Sub
                    Else
140                     Call VolverCriminal(Userindex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
142         If UserList(Userindex).flags.Privilegios And PlayerType.user Then
144             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
            
146         If UserList(tU).flags.invisible = 1 Then
148             If tU = Userindex Then
150                 Call WriteConsoleMsg(Userindex, "¡Ya estás invisible!", FontTypeNames.FONTTYPE_INFO)
                Else
152                 Call WriteConsoleMsg(Userindex, "¡El objetivo ya se encuentra invisible!", FontTypeNames.FONTTYPE_INFO)
                End If
154             b = False
                Exit Sub
            End If
   
156         UserList(tU).flags.invisible = 1
            'Ladder
            'Reseteamos el contador de Invisibilidad
158         UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
160         Call WriteContadores(tU)
162         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

164         Call InfoHechizo(Userindex)
166         b = True

        End If
        
168     If Hechizos(h).Mimetiza = 1 Then
170         If UserList(tU).flags.Muerto = 1 Then
                Exit Sub
            End If
            
172         If UserList(tU).flags.Navegando = 1 Then
                Exit Sub
            End If
174         If UserList(Userindex).flags.Navegando = 1 Then
                Exit Sub
            End If
            
            'Si sos user, no uses este hechizo con GMS.
176         If UserList(Userindex).flags.Privilegios And PlayerType.user Then
178             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub
                End If
            End If
            
180         If UserList(Userindex).flags.Mimetizado = 1 Then
182             Call WriteConsoleMsg(Userindex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
184         If UserList(Userindex).flags.AdminInvisible = 1 Then Exit Sub
            
            'copio el char original al mimetizado
            
186         With UserList(Userindex)
188             .CharMimetizado.Body = .Char.Body
190             .CharMimetizado.Head = .Char.Head
192             .CharMimetizado.CascoAnim = .Char.CascoAnim
194             .CharMimetizado.ShieldAnim = .Char.ShieldAnim
196             .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                
198             .flags.Mimetizado = 1
                
                'ahora pongo local el del enemigo
200             .Char.Body = UserList(tU).Char.Body
202             .Char.Head = UserList(tU).Char.Head
204             .Char.CascoAnim = UserList(tU).Char.CascoAnim
206             .Char.ShieldAnim = UserList(tU).Char.ShieldAnim
208             .Char.WeaponAnim = UserList(tU).Char.WeaponAnim
            
210             Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            End With
           
212        Call InfoHechizo(Userindex)
214        b = True
        End If

216     If Hechizos(h).Envenena > 0 Then
            ' If UserIndex = tU Then
            '    Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Exit Sub
            'End If
    
218         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
220         If Userindex <> tU Then
222             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

224         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
226         Call InfoHechizo(Userindex)
228         b = True

        End If

230     If Hechizos(h).desencantar = 1 Then
            ' Call WriteConsoleMsg(UserIndex, "Ningun efecto magico tiene efecto sobre ti ya.", FontTypeNames.FONTTYPE_INFOIAO)

232         UserList(Userindex).flags.Envenenado = 0
234         UserList(Userindex).flags.Incinerado = 0
    
236         If UserList(Userindex).flags.Inmovilizado = 1 Then
238             UserList(Userindex).Counters.Inmovilizado = 0
240             UserList(Userindex).flags.Inmovilizado = 0
242             Call WriteInmovilizaOK(Userindex)
            

            End If
    
244         If UserList(Userindex).flags.Paralizado = 1 Then
246             UserList(Userindex).flags.Paralizado = 0
248             Call WriteParalizeOK(Userindex)
            
           
            End If
        
250         If UserList(Userindex).flags.Ceguera = 1 Then
252             UserList(Userindex).flags.Ceguera = 0
254             Call WriteBlindNoMore(Userindex)
            

            End If
    
256         If UserList(Userindex).flags.Maldicion = 1 Then
258             UserList(Userindex).flags.Maldicion = 0
260             UserList(Userindex).Counters.Maldicion = 0

            End If
    
262         Call InfoHechizo(Userindex)
264         b = True

        End If

266     If Hechizos(h).incinera = 1 Then
268         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
270             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
272         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
274         If Userindex <> tU Then
276             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

278         UserList(tU).flags.Incinerado = 1
280         Call InfoHechizo(Userindex)
282         b = True

        End If

284     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
286         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
288             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
290             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
292         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
294             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
296                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
298                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
300                     b = False
                        Exit Sub

                    End If

302                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
304                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
306                     b = False
                        Exit Sub
                    Else

                        '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
308         If UserList(Userindex).flags.Privilegios And PlayerType.user Then
310             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
312         UserList(tU).flags.Envenenado = 0
314         Call InfoHechizo(Userindex)
316         b = True

        End If

318     If Hechizos(h).Maldicion = 1 Then
320         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
322             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
324         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
326         If Userindex <> tU Then
328             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

330         UserList(tU).flags.Maldicion = 1
332         UserList(tU).Counters.Maldicion = 200
    
334         Call InfoHechizo(Userindex)
336         b = True

        End If

338     If Hechizos(h).RemoverMaldicion = 1 Then
340         UserList(tU).flags.Maldicion = 0
342         Call InfoHechizo(Userindex)
344         b = True

        End If

346     If Hechizos(h).GolpeCertero = 1 Then
348         UserList(tU).flags.GolpeCertero = 1
350         Call InfoHechizo(Userindex)
352         b = True

        End If

354     If Hechizos(h).Bendicion = 1 Then
356         UserList(tU).flags.Bendicion = 1
358         Call InfoHechizo(Userindex)
360         b = True

        End If

362     If Hechizos(h).Paraliza = 1 Then
364         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
366             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
368         If UserList(tU).flags.Paralizado = 1 Then
370             Call WriteConsoleMsg(Userindex, UserList(tU).name & " ya está paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
372         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
374         If Userindex <> tU Then
376             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
378         Call InfoHechizo(Userindex)
380         b = True

382         If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
384             Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
386             Call WriteConsoleMsg(Userindex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            
                Exit Sub

            End If
            
388         UserList(tU).Counters.Paralisis = Hechizos(h).Duration

390         If UserList(tU).flags.Paralizado = 0 Then
392             UserList(tU).flags.Paralizado = 1
394             Call WriteParalizeOK(tU)
396             Call WritePosUpdate(tU)
            End If

        End If

398     If Hechizos(h).Velocidad > 0 Then
400         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
402             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
404         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
406         If Userindex <> tU Then
408             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
410         Call InfoHechizo(Userindex)
412         b = True
                 
414         If UserList(tU).Counters.Velocidad = 0 Then
416             UserList(tU).flags.VelocidadBackup = UserList(tU).Char.speeding
            End If

418         UserList(tU).Char.speeding = Hechizos(h).Velocidad
420         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSpeedingACT(UserList(tU).Char.CharIndex, UserList(tU).Char.speeding))
            'End If
422         UserList(tU).Counters.Velocidad = Hechizos(h).Duration

        End If

424     If Hechizos(h).Inmoviliza = 1 Then
426         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
428             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
430         If UserList(tU).flags.Paralizado = 1 Then
432             Call WriteConsoleMsg(Userindex, UserList(tU).name & " ya está paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
434         ElseIf UserList(tU).flags.Inmovilizado = 1 Then
436             Call WriteConsoleMsg(Userindex, UserList(tU).name & " ya está inmovilizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
438         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
440         If Userindex <> tU Then
442             Call UsuarioAtacadoPorUsuario(Userindex, tU)
            End If
            
444         Call InfoHechizo(Userindex)
446         b = True
            '  If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
            '   Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            '
            '    Exit Sub
            ' End If
            
448         UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration

450         UserList(tU).flags.Inmovilizado = 1
452         Call WriteInmovilizaOK(tU)
454         Call WritePosUpdate(tU)
            

        End If

456     If Hechizos(h).RemoverParalisis = 1 Then
        
            'Para poder tirar remo a un pk en el ring
458         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
460             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
462                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
464                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
466                     b = False
                        Exit Sub

                    End If

468                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
470                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
472                     b = False
                        Exit Sub
                    Else
474                     Call VolverCriminal(Userindex)

                    End If

                End If

            End If
        
476         If UserList(tU).flags.Inmovilizado = 0 And UserList(tU).flags.Paralizado = 0 Then
478             Call WriteConsoleMsg(Userindex, "El objetivo no esta paralizado.", FontTypeNames.FONTTYPE_INFO)
480             b = False
                Exit Sub

            End If
        
482         If UserList(tU).flags.Inmovilizado = 1 Then
484             UserList(tU).Counters.Inmovilizado = 0
486             UserList(tU).flags.Inmovilizado = 0
488             Call WriteInmovilizaOK(tU)
490             Call WritePosUpdate(tU)
                ' Call InfoHechizo(UserIndex)
            

                'b = True
            End If
    
492         If UserList(tU).flags.Paralizado = 1 Then
494             UserList(tU).flags.Paralizado = 0
                'no need to crypt this
496             Call WriteParalizeOK(tU)
            

                '  b = True
            End If

498         b = True
500         Call InfoHechizo(Userindex)

        End If

502     If Hechizos(h).RemoverEstupidez = 1 Then
504         If UserList(tU).flags.Estupidez = 1 Then

                'Para poder tirar remo estu a un pk en el ring
506             If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
508                 If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
510                     If esArmada(Userindex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
512                         Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
514                         b = False
                            Exit Sub

                        End If

516                     If UserList(Userindex).flags.Seguro Then
                            'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
518                         Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
520                         b = False
                            Exit Sub
                        Else

                            ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                        End If

                    End If

                End If
    
522             UserList(tU).flags.Estupidez = 0
                'no need to crypt this
524             Call WriteDumbNoMore(tU)
            
526             Call InfoHechizo(Userindex)
528             b = True

            End If

        End If

530     If Hechizos(h).Revivir = 1 Then
532         If UserList(tU).flags.Muerto = 1 Then
    
                'No usar resu en mapas con ResuSinEfecto
                'If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
                '   Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                '   b = False
                '   Exit Sub
                ' End If
        
534             If UserList(tU).Accion.TipoAccion = Accion_Barra.Resucitar Then
536                 Call WriteConsoleMsg(Userindex, "El usuario ya esta siendo resucitado.", FontTypeNames.FONTTYPE_INFO)
538                 b = False
                    Exit Sub

                End If
        
                'Para poder tirar revivir a un pk en el ring
540             If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
542                 If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
544                     If esArmada(Userindex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
546                         Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
548                         b = False
                            Exit Sub

                        End If

550                     If UserList(Userindex).flags.Seguro Then
                            'call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
552                         Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
554                         b = False
                            Exit Sub
                        Else
556                         Call VolverCriminal(Userindex)

                        End If

                    End If

                End If
                        
558             Call SendData(SendTarget.ToPCArea, tU, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, ParticulasIndex.Resucitar, 600, False))
560             Call SendData(SendTarget.ToPCArea, tU, PrepareMessageBarFx(UserList(tU).Char.CharIndex, 600, Accion_Barra.Resucitar))
562             UserList(tU).Accion.AccionPendiente = True
564             UserList(tU).Accion.Particula = ParticulasIndex.Resucitar
566             UserList(tU).Accion.TipoAccion = Accion_Barra.Resucitar
                
                'Pablo Toxic Waste (GD: 29/04/07)
                'UserList(tU).Stats.MinAGU = 0
                'UserList(tU).flags.Sed = 1
                'UserList(tU).Stats.MinHam = 0
                'UserList(tU).flags.Hambre = 1
568             Call WriteUpdateHungerAndThirst(tU)
570             Call InfoHechizo(Userindex)
                'UserList(tU).Stats.MinMAN = 0
                'UserList(tU).Stats.MinSta = 0
572             b = True
                'Solo saco vida si es User. no quiero que exploten GMs por ahi.
        
                'Call RevivirUsuario(tU)
            Else
574             b = False

            End If

        End If

576     If Hechizos(h).Ceguera = 1 Then
578         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
580             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
582         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
584         If Userindex <> tU Then
586             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

588         UserList(tU).flags.Ceguera = 1
590         UserList(tU).Counters.Ceguera = Hechizos(h).Duration

592         Call WriteBlind(tU)
        
594         Call InfoHechizo(Userindex)
596         b = True

        End If

598     If Hechizos(h).Estupidez = 1 Then
600         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
602             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

604         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
606         If Userindex <> tU Then
608             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

610         If UserList(tU).flags.Estupidez = 0 Then
612             UserList(tU).flags.Estupidez = 1
614             UserList(tU).Counters.Estupidez = Hechizos(h).Duration

            End If

616         Call WriteDumb(tU)
        

618         Call InfoHechizo(Userindex)
620         b = True

        End If

        
        Exit Sub

HechizoEstadoUsuario_Err:
622     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoEstadoUsuario", Erl)
624     Resume Next
        
End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal Userindex As Integer)
        
        On Error GoTo HechizoEstadoNPC_Err
        

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 04/13/2008
        'Handles the Spells that afect the Stats of an NPC
        '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
        'removidos por users de su misma faccion.
        '***************************************************
100     If Hechizos(hIndex).Invisibilidad = 1 Then
102         Call InfoHechizo(Userindex)
104         Npclist(NpcIndex).flags.invisible = 1
106         b = True

        End If

108     If Hechizos(hIndex).Envenena > 0 Then
110         If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
112             b = False
                Exit Sub

            End If

114         Call NPCAtacado(NpcIndex, Userindex)
116         Call InfoHechizo(Userindex)
118         Npclist(NpcIndex).flags.Envenenado = Hechizos(hIndex).Envenena
120         b = True

        End If

122     If Hechizos(hIndex).CuraVeneno = 1 Then
124         Call InfoHechizo(Userindex)
126         Npclist(NpcIndex).flags.Envenenado = 0
128         b = True

        End If

130     If Hechizos(hIndex).RemoverMaldicion = 1 Then
132         Call InfoHechizo(Userindex)
            'Npclist(NpcIndex).flags.Maldicion = 0
134         b = True

        End If

136     If Hechizos(hIndex).Bendicion = 1 Then
138         Call InfoHechizo(Userindex)
140         Npclist(NpcIndex).flags.Bendicion = 1
142         b = True

        End If

144     If Hechizos(hIndex).Paraliza = 1 Then
146         If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
148             If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
150                 b = False
                    Exit Sub

                End If

152             Call NPCAtacado(NpcIndex, Userindex)
154             Call InfoHechizo(Userindex)
156             Npclist(NpcIndex).flags.Paralizado = 1
158             Npclist(NpcIndex).flags.Inmovilizado = 0
160             Npclist(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 6

162             Call AnimacionIdle(NpcIndex, False)
                
164             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
166             Call WriteLocaleMsg(Userindex, "381", FontTypeNames.FONTTYPE_INFO)
168             b = False
                Exit Sub

            End If

        End If

170     If Hechizos(hIndex).RemoverParalisis = 1 Then
172         If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1 Then
174             If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
176                 If esArmada(Userindex) Then
178                     Call InfoHechizo(Userindex)
180                     Npclist(NpcIndex).flags.Paralizado = 0
182                     Npclist(NpcIndex).Contadores.Paralisis = 0
184                     b = True
                        Exit Sub
                    Else
186                     Call WriteConsoleMsg(Userindex, "Solo podés Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
188                     b = False
                        Exit Sub

                    End If
                
190                 Call WriteConsoleMsg(Userindex, "Solo podés Remover la Parálisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
192                 b = False
                    Exit Sub
                Else

194                 If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
196                     If esCaos(Userindex) Then
198                         Call InfoHechizo(Userindex)
200                         Npclist(NpcIndex).flags.Paralizado = 0
202                         Npclist(NpcIndex).Contadores.Paralisis = 0
204                         b = True
                            Exit Sub
                        Else
206                         Call WriteConsoleMsg(Userindex, "Solo podés Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
208                         b = False
                            Exit Sub

                        End If

                    End If

                End If

            Else
210             Call WriteConsoleMsg(Userindex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFO)
212             b = False
                Exit Sub

            End If

        End If
 
214     If Hechizos(hIndex).Inmoviliza = 1 And Npclist(NpcIndex).flags.Paralizado = 0 Then
216         If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
218             If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
220                 b = False
                    Exit Sub

                End If

222             Call NPCAtacado(NpcIndex, Userindex)
224             Npclist(NpcIndex).flags.Inmovilizado = 1
226             Npclist(NpcIndex).flags.Paralizado = 0
228             Npclist(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 6

230             Call AnimacionIdle(NpcIndex, True)

232             Call InfoHechizo(Userindex)
234             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
236             Call WriteLocaleMsg(Userindex, "381", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

238     If Hechizos(hIndex).Mimetiza = 1 Then
    
240         If UserList(Userindex).flags.Mimetizado = 1 Then
242             Call WriteConsoleMsg(Userindex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
244         If UserList(Userindex).flags.AdminInvisible = 1 Then Exit Sub
            
                
246         If UserList(Userindex).clase = eClass.Druid Then
                'copio el char original al mimetizado
248             With UserList(Userindex)
250                 .CharMimetizado.Body = .Char.Body
252                 .CharMimetizado.Head = .Char.Head
254                 .CharMimetizado.CascoAnim = .Char.CascoAnim
256                 .CharMimetizado.ShieldAnim = .Char.ShieldAnim
258                 .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                    
260                 .flags.Mimetizado = 1
                    
                    'ahora pongo lo del NPC.
262                 .Char.Body = Npclist(NpcIndex).Char.Body
264                 .Char.Head = Npclist(NpcIndex).Char.Head
266                 .Char.CascoAnim = NingunCasco
268                 .Char.ShieldAnim = NingunEscudo
270                 .Char.WeaponAnim = NingunArma
                
272                 Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                End With
            Else
274             Call WriteConsoleMsg(Userindex, "Solo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
276        Call InfoHechizo(Userindex)
278        b = True
        End If
        
        Exit Sub

HechizoEstadoNPC_Err:
280     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoEstadoNPC", Erl)
282     Resume Next
        
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 14/08/2007
        'Handles the Spells that afect the Life NPC
        '14/08/2007 Pablo (ToxicWaste) - Orden general.
        '***************************************************
        
        On Error GoTo HechizoPropNPC_Err
        

        Dim daño As Long
    
        'Salud
100     If Hechizos(hIndex).SubeHP = 1 Then
102         daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
            'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        
104         Call InfoHechizo(Userindex)
106         Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp + daño

108         If Npclist(NpcIndex).Stats.MinHp > Npclist(NpcIndex).Stats.MaxHp Then Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MaxHp
110         Call WriteConsoleMsg(Userindex, "Has curado " & daño & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
112         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(daño, Npclist(NpcIndex).Char.CharIndex, vbGreen))
114         b = True
        
116     ElseIf Hechizos(hIndex).SubeHP = 2 Then

118         If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
120             b = False
                Exit Sub

            End If
        
122         Call NPCAtacado(NpcIndex, Userindex)
124         daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
        
126         daño = daño + Porcentaje(daño, 3 * UserList(Userindex).Stats.ELV)
    
            ' Daño mágico arma
128         If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
130             daño = daño + Porcentaje(daño, ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Daño mágico anillo
132         If UserList(Userindex).Invent.AnilloEqpObjIndex > 0 Then
134             daño = daño + Porcentaje(daño, ObjData(UserList(Userindex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
            End If

136         b = True
        
138         If Npclist(NpcIndex).flags.Snd2 > 0 Then
140             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
            End If
        
            'Quizas tenga defenza magica el NPC.
142         If Hechizos(hIndex).AntiRm = 0 Then
144             daño = daño - Npclist(NpcIndex).Stats.defM
            End If
        
146         If daño < 0 Then daño = 0
        
148         Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - daño
150         Call InfoHechizo(Userindex)
        
152         If UserList(Userindex).ChatCombate = 1 Then
154             Call WriteConsoleMsg(Userindex, "Le has causado " & daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
156         Call CalcularDarExp(Userindex, NpcIndex, daño)
    
158         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(daño, Npclist(NpcIndex).Char.CharIndex))
    
160         If Npclist(NpcIndex).Stats.MinHp < 1 Then
162             Npclist(NpcIndex).Stats.MinHp = 0
164             Call MuereNpc(NpcIndex, Userindex)

            End If

        End If

        
        Exit Sub

HechizoPropNPC_Err:
166     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoPropNPC", Erl)
168     Resume Next
        
End Sub

Sub InfoHechizo(ByVal Userindex As Integer)
        
        On Error GoTo InfoHechizo_Err
        

        Dim h As Integer

100     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
    
102     If UserList(Userindex).flags.NoPalabrasMagicas = 0 Then
104         Call DecirPalabrasMagicas(h, Userindex)

        End If

106     If UserList(Userindex).flags.TargetUser > 0 Then '¿El Hechizo fue tirado sobre un usuario?
108         If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
110             If Hechizos(h).ParticleViaje > 0 Then
112                 Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(Userindex).Char.CharIndex, UserList(UserList(Userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                Else
114                 Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessageCreateFX(UserList(UserList(Userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                End If

            End If

116         If Hechizos(h).Particle > 0 Then '¿Envio Particula?
118             If Hechizos(h).ParticleViaje > 0 Then
120                 Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(Userindex).Char.CharIndex, UserList(UserList(Userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                Else
122                 Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessageParticleFX(UserList(UserList(Userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

                End If

            End If
        
124         If Hechizos(h).ParticleViaje = 0 Then
126             Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserList(Userindex).flags.TargetUser).Pos.X, UserList(UserList(Userindex).flags.TargetUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)

            End If
        
128         If Hechizos(h).TimeEfect <> 0 Then 'Envio efecto de screen
130             Call WriteEfectToScreen(Userindex, Hechizos(h).ScreenColor, Hechizos(h).TimeEfect)

            End If

132     ElseIf UserList(Userindex).flags.TargetNPC > 0 Then '¿El Hechizo fue tirado sobre un npc?

134         If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
136             If Npclist(UserList(Userindex).flags.TargetNPC).Stats.MinHp < 1 Then

                    'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
138                 If Hechizos(h).ParticleViaje > 0 Then
140                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(Userindex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))
                    Else
142                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))

                    End If

                Else

144                 If Hechizos(h).ParticleViaje > 0 Then
146                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(Userindex).Char.CharIndex, Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                    Else
148                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageCreateFX(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                    End If

                End If

            End If
        
150         If Hechizos(h).Particle > 0 Then '¿Envio Particula?
152             If Npclist(UserList(Userindex).flags.TargetNPC).Stats.MinHp < 1 Then
154                 Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(Userindex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0, Npclist(UserList(Userindex).flags.TargetNPC).Pos.X, Npclist(UserList(Userindex).flags.TargetNPC).Pos.Y))
                    'Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXToFloor(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.X, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y, Hechizos(H).Particle, Hechizos(H).TimeParticula))
                Else

156                 If Hechizos(h).ParticleViaje > 0 Then
158                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(Userindex).Char.CharIndex, Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                    Else
160                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFX(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

                    End If

                End If

            End If

162         If Hechizos(h).ParticleViaje = 0 Then
164             Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(h).wav, Npclist(UserList(Userindex).flags.TargetNPC).Pos.X, Npclist(UserList(Userindex).flags.TargetNPC).Pos.Y))

            End If

        Else ' Entonces debe ser sobre el terreno

166         If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
168             Call modSendData.SendToAreaByPos(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))

            End If
        
170         If Hechizos(h).Particle > 0 Then 'Envio Particula?
172             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXToFloor(UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

            End If
        
174         If Hechizos(h).wav <> 0 Then
176             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(h).wav, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))   'Esta linea faltaba. Pablo (ToxicWaste)

            End If
    
        End If
    
178     If UserList(Userindex).ChatCombate = 1 Then
180         If Hechizos(h).Target = TargetType.uTerreno Then
182             Call WriteConsoleMsg(Userindex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)
            
184         ElseIf UserList(Userindex).flags.TargetUser > 0 Then

                'Optimizacion de protocolo por Ladder
186             If Userindex <> UserList(Userindex).flags.TargetUser Then
188                 Call WriteConsoleMsg(Userindex, "HecMSGU*" & h & "*" & UserList(UserList(Userindex).flags.TargetUser).name, FontTypeNames.FONTTYPE_FIGHT)
190                 Call WriteConsoleMsg(UserList(Userindex).flags.TargetUser, "HecMSGA*" & h & "*" & UserList(Userindex).name, FontTypeNames.FONTTYPE_FIGHT)
    
                Else
192                 Call WriteConsoleMsg(Userindex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

                End If

194         ElseIf UserList(Userindex).flags.TargetNPC > 0 Then
196             Call WriteConsoleMsg(Userindex, "HecMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

            Else
198             Call WriteConsoleMsg(Userindex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

        
        Exit Sub

InfoHechizo_Err:
200     Call RegistrarError(Err.Number, Err.description, "modHechizos.InfoHechizo", Erl)
202     Resume Next
        
End Sub

Sub HechizoPropUsuario(ByVal Userindex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoPropUsuario_Err
        

        Dim h As Integer

        Dim daño As Integer

        Dim tempChr           As Integer

        Dim enviarInfoHechizo As Boolean

100     enviarInfoHechizo = False
    
102     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
104     tempChr = UserList(Userindex).flags.TargetUser
      
        'Hambre
106     If Hechizos(h).SubeHam = 1 Then
    
108         enviarInfoHechizo = True
    
110         daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
112         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + daño

114         If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
116         If Userindex <> tempChr Then
118             Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
120             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
122             Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
124         Call WriteUpdateHungerAndThirst(tempChr)
126         b = True
    
128     ElseIf Hechizos(h).SubeHam = 2 Then

130         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
132         If Userindex <> tempChr Then
134             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
            Else
                Exit Sub

            End If
    
136         enviarInfoHechizo = True
    
138         daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
140         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
142         If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
144         If Userindex <> tempChr Then
146             Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
148             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
150             Call WriteConsoleMsg(Userindex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

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
    
164         enviarInfoHechizo = True
    
166         daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
168         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + daño

170         If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
172         If Userindex <> tempChr Then
174             Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
176             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
178             Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
180         Call WriteUpdateHungerAndThirst(tempChr)
    
182         b = True
    
184     ElseIf Hechizos(h).SubeSed = 2 Then
    
186         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
188         If Userindex <> tempChr Then
190             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
192         enviarInfoHechizo = True
    
194         daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
196         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
198         If Userindex <> tempChr Then
200             Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
202             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
204             Call WriteConsoleMsg(Userindex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

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
218         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
220             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
222                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
224                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
226                     b = False
                        Exit Sub

                    End If

228                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
230                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
232                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
234         enviarInfoHechizo = True
236         daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
238         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

240         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
         
242         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)

244         UserList(tempChr).flags.TomoPocion = True
246         b = True
248         Call WriteFYA(tempChr)
    
250     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
252         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
254         If Userindex <> tempChr Then
256             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
258         enviarInfoHechizo = True
    
260         UserList(tempChr).flags.TomoPocion = True
262         daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
264         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

266         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño < MINATRIBUTOS Then
268             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
            Else
270             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño

            End If
    
272         b = True
274         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
276     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
278         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
280             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
282                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
284                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
286                     b = False
                        Exit Sub

                    End If

288                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
290                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
292                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
294         daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
296         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
298         UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño

300         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    
302         UserList(tempChr).flags.TomoPocion = True
            
304         Call WriteFYA(tempChr)

306         b = True
    
308         enviarInfoHechizo = True
310         Call WriteFYA(tempChr)

312     ElseIf Hechizos(h).SubeFuerza = 2 Then

314         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
316         If Userindex <> tempChr Then
318             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
320         UserList(tempChr).flags.TomoPocion = True
    
322         daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
324         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

326         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño < MINATRIBUTOS Then
328             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
            Else
330             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño

            End If

332         b = True
334         enviarInfoHechizo = True
336         Call WriteFYA(tempChr)

        End If

        'Salud
338     If Hechizos(h).SubeHP = 1 Then
    
            'Verifica que el usuario no este muerto
340         If UserList(tempChr).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
342             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
344             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar a un pk en el ring
346         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
348             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
350                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
352                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
354                     b = False
                        Exit Sub

                    End If

356                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
358                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
360                     b = False
                        Exit Sub
                    Else

                        'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
       
362         daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            ' daño = daño + Porcentaje(daño, 2 * UserList(UserIndex).Stats.ELV)
    
364         enviarInfoHechizo = True

366         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + daño

368         If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp
    
370         If Userindex <> tempChr Then
372             Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
374             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
376             Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
378         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex, vbGreen))
380         Call WriteUpdateHP(tempChr)
    
382         b = True

384     ElseIf Hechizos(h).SubeHP = 2 Then
    
386         If Userindex = tempChr Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
388             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

390         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
392         daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
394         daño = daño + Porcentaje(daño, 3 * UserList(Userindex).Stats.ELV)

            ' Daño mágico arma
396         If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
398             daño = daño + Porcentaje(daño, ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Daño mágico anillo
400         If UserList(Userindex).Invent.AnilloEqpObjIndex > 0 Then
402             daño = daño + Porcentaje(daño, ObjData(UserList(Userindex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Si el hechizo no ignora la RM
404         If Hechizos(h).AntiRm = 0 Then
                ' Resistencia mágica armadura
406             If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
408                 daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica anillo
410             If UserList(tempChr).Invent.AnilloEqpObjIndex > 0 Then
412                 daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica escudo
414             If UserList(tempChr).Invent.EscudoEqpObjIndex > 0 Then
416                 daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica casco
418             If UserList(tempChr).Invent.CascoEqpObjIndex > 0 Then
420                 daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).ResistenciaMagica)
                End If
            End If

            ' Prevengo daño negativo
422         If daño < 0 Then daño = 0
    
424         If Userindex <> tempChr Then
426             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
            End If
    
428         enviarInfoHechizo = True
    
430         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - daño
    
432         Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
434         Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    
436         Call SubirSkill(tempChr, Resistencia)
    
438         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex))

            'Muere
440         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
442             Call Statistics.StoreFrag(Userindex, tempChr)
444             Call ContarMuerte(tempChr, Userindex)
446             UserList(tempChr).Stats.MinHp = 0
448             Call ActStats(tempChr, Userindex)

                '  Call UserDie(tempChr)
            End If
            
450         Call WriteUpdateHP(tempChr)
    
452         b = True

        End If

        'Mana
454     If Hechizos(h).SubeMana = 1 Then
    
456         enviarInfoHechizo = True
458         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + daño

460         If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
462         If Userindex <> tempChr Then
464             Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
466             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
468             Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
470         Call WriteUpdateMana(tempChr)
    
472         b = True
    
474     ElseIf Hechizos(h).SubeMana = 2 Then

476         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
478         If Userindex <> tempChr Then
480             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
482         enviarInfoHechizo = True
    
484         If Userindex <> tempChr Then
486             Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
488             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
490             Call WriteConsoleMsg(Userindex, "Te has quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
492         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño

494         If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0

496         Call WriteUpdateMana(tempChr)

498         b = True
    
        End If

        'Stamina
500     If Hechizos(h).SubeSta = 1 Then
502         Call InfoHechizo(Userindex)
504         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + daño

506         If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta

508         If Userindex <> tempChr Then
510             Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
512             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
514             Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
516         Call WriteUpdateSta(tempChr)

518         b = True
520     ElseIf Hechizos(h).SubeSta = 2 Then

522         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
524         If Userindex <> tempChr Then
526             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
528         enviarInfoHechizo = True
    
530         If Userindex <> tempChr Then
532             Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
534             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
536             Call WriteConsoleMsg(Userindex, "Te has quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
538         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
540         If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0

542         Call WriteUpdateSta(tempChr)

544         b = True

        End If

546     If enviarInfoHechizo Then
548         Call InfoHechizo(Userindex)

        End If

    

        
        Exit Sub

HechizoPropUsuario_Err:
550     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoPropUsuario", Erl)
552     Resume Next
        
End Sub

Sub HechizoCombinados(ByVal Userindex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoCombinados_Err
        

        Dim h As Integer

        Dim daño As Integer

        Dim tempChr           As Integer

        Dim enviarInfoHechizo As Boolean

100     enviarInfoHechizo = False
    
102     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
104     tempChr = UserList(Userindex).flags.TargetUser
      
        ' <-------- Agilidad ---------->
106     If Hechizos(h).SubeAgilidad = 1 Then
    
            'Para poder tirar cl a un pk en el ring
108         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
110             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
112                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
114                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
116                     b = False
                        Exit Sub

                    End If

118                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
120                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
122                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
124         enviarInfoHechizo = True
126         daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
128         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
            'UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
            'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)

130         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño

132         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
        
134         UserList(tempChr).flags.TomoPocion = True
136         b = True
138         Call WriteFYA(tempChr)
    
140     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
142         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
144         If Userindex <> tempChr Then
146             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
148         enviarInfoHechizo = True
    
150         UserList(tempChr).flags.TomoPocion = True
152         daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
154         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

156         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño < 6 Then
158             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
            Else
160             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño

            End If

            'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
162         b = True
164         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
166     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
168         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
170             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
172                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
174                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
176                     b = False
                        Exit Sub

                    End If

178                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
180                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
182                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
184         daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
186         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
188         UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño

190         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
    
192         UserList(tempChr).flags.TomoPocion = True
194         b = True
    
196         enviarInfoHechizo = True
198         Call WriteFYA(tempChr)
200     ElseIf Hechizos(h).SubeFuerza = 2 Then

202         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
204         If Userindex <> tempChr Then
206             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
208         UserList(tempChr).flags.TomoPocion = True
    
210         daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
212         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        
214         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño < 6 Then
216             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
            Else
218             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño

            End If
   
220         b = True
222         enviarInfoHechizo = True
224         Call WriteFYA(tempChr)

        End If

        'Salud
226     If Hechizos(h).SubeHP = 1 Then
    
            'Verifica que el usuario no este muerto
228         If UserList(tempChr).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
230             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
232             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar a un pk en el ring
234         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
236             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
238                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
240                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
242                     b = False
                        Exit Sub

                    End If

244                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
246                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
248                     b = False
                        Exit Sub
                    Else

                        'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
       
250         daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
252         enviarInfoHechizo = True

254         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + daño

256         If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp
    
258         If Userindex <> tempChr Then
260             Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
262             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
264             Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
266         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex, vbGreen))
    
268         b = True
270     ElseIf Hechizos(h).SubeHP = 2 Then
    
272         If Userindex = tempChr Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
274             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
276         daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
278         daño = daño + Porcentaje(daño, 3 * UserList(Userindex).Stats.ELV)
            
            ' Daño mágico arma
280         If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
282             daño = daño + Porcentaje(daño, ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Daño mágico anillo
284         If UserList(Userindex).Invent.AnilloEqpObjIndex > 0 Then
286             daño = daño + Porcentaje(daño, ObjData(UserList(Userindex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Si el hechizo no ignora la RM
288         If Hechizos(h).AntiRm = 0 Then
                ' Resistencia mágica armadura
290             If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
292                 daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica anillo
294             If UserList(tempChr).Invent.AnilloEqpObjIndex > 0 Then
296                 daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica escudo
298             If UserList(tempChr).Invent.EscudoEqpObjIndex > 0 Then
300                 daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica casco
302             If UserList(tempChr).Invent.CascoEqpObjIndex > 0 Then
304                 daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).ResistenciaMagica)
                End If
            End If

            ' Prevengo daño negativo
306         If daño < 0 Then daño = 0
    
308         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
310         If Userindex <> tempChr Then
312             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
314         enviarInfoHechizo = True
    
316         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - daño
    
318         Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
320         Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
322         Call SubirSkill(tempChr, Resistencia)
324         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex))

            'Muere
326         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
328             Call Statistics.StoreFrag(Userindex, tempChr)
        
330             Call ContarMuerte(tempChr, Userindex)
332             UserList(tempChr).Stats.MinHp = 0
334             Call ActStats(tempChr, Userindex)

                'Call UserDie(tempChr)
            End If
    
336         b = True

        End If

        Dim tU As Integer

338     tU = tempChr

340     If Hechizos(h).Invisibilidad = 1 Then
   
342         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
344             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
346             b = False
                Exit Sub

            End If
    
348         If UserList(tU).Counters.Saliendo Then
350             If Userindex <> tU Then
352                 Call WriteConsoleMsg(Userindex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
354                 b = False
                    Exit Sub
                Else
356                 Call WriteConsoleMsg(Userindex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
358                 b = False
                    Exit Sub

                End If

            End If
    
            'Para poder tirar invi a un pk en el ring
360         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
362             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
364                 If esArmada(Userindex) Then
366                     Call WriteConsoleMsg(Userindex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
368                     b = False
                        Exit Sub

                    End If

370                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
372                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
374                     b = False
                        Exit Sub
                    Else
376                     Call VolverCriminal(Userindex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
378         If UserList(Userindex).flags.Privilegios And PlayerType.user Then
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
    
396         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
398         If Userindex <> tU Then
400             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

402         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
404         enviarInfoHechizo = True
406         b = True

        End If

408     If Hechizos(h).desencantar = 1 Then
410         Call WriteConsoleMsg(Userindex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)

412         UserList(Userindex).flags.Envenenado = 0
414         UserList(Userindex).flags.Incinerado = 0
    
416         If UserList(Userindex).flags.Inmovilizado = 1 Then
418             UserList(Userindex).Counters.Inmovilizado = 0
420             UserList(Userindex).flags.Inmovilizado = 0
422             Call WriteInmovilizaOK(Userindex)
            

            End If
    
424         If UserList(Userindex).flags.Paralizado = 1 Then
426             UserList(Userindex).flags.Paralizado = 0
428             Call WriteParalizeOK(Userindex)
            
           
            End If
        
430         If UserList(Userindex).flags.Ceguera = 1 Then
432             UserList(Userindex).flags.Ceguera = 0
434             Call WriteBlindNoMore(Userindex)
            

            End If
    
436         If UserList(Userindex).flags.Maldicion = 1 Then
438             UserList(Userindex).flags.Maldicion = 0
440             UserList(Userindex).Counters.Maldicion = 0

            End If
    
442         enviarInfoHechizo = True
444         b = True

        End If

446     If Hechizos(h).Sanacion = 1 Then

448         UserList(tU).flags.Envenenado = 0
450         UserList(tU).flags.Incinerado = 0
452         enviarInfoHechizo = True
454         b = True

        End If

456     If Hechizos(h).incinera = 1 Then
458         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
460             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
462         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
464         If Userindex <> tU Then
466             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

468         UserList(tU).flags.Incinerado = 1
470         enviarInfoHechizo = True
472         b = True

        End If

474     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
476         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
478             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
480             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
482         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
484             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
486                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
488                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
490                     b = False
                        Exit Sub

                    End If

492                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
494                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
496                     b = False
                        Exit Sub
                    Else

                        '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
498         If UserList(Userindex).flags.Privilegios And PlayerType.user Then
500             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
502         UserList(tU).flags.Envenenado = 0
504         enviarInfoHechizo = True
506         b = True

        End If

508     If Hechizos(h).Maldicion = 1 Then
510         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
512             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
514         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
516         If Userindex <> tU Then
518             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

520         UserList(tU).flags.Maldicion = 1
522         UserList(tU).Counters.Maldicion = 200
    
524         enviarInfoHechizo = True
526         b = True

        End If

528     If Hechizos(h).RemoverMaldicion = 1 Then
530         UserList(tU).flags.Maldicion = 0
532         enviarInfoHechizo = True
534         b = True

        End If

536     If Hechizos(h).GolpeCertero = 1 Then
538         UserList(tU).flags.GolpeCertero = 1
540         enviarInfoHechizo = True
542         b = True

        End If

544     If Hechizos(h).Bendicion = 1 Then
546         UserList(tU).flags.Bendicion = 1
548         enviarInfoHechizo = True
550         b = True

        End If

552     If Hechizos(h).Paraliza = 1 Then
554         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
556             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
558         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
560         If Userindex <> tU Then
562             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
564         enviarInfoHechizo = True
566         b = True

568         If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
570             Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
572             Call WriteConsoleMsg(Userindex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            
                Exit Sub

            End If
            
574         UserList(tU).Counters.Paralisis = Hechizos(h).Duration

576         If UserList(tU).flags.Paralizado = 0 Then
578             UserList(tU).flags.Paralizado = 1
580             Call WriteParalizeOK(tU)
582             Call WritePosUpdate(tU)
            End If

        End If

584     If Hechizos(h).Inmoviliza = 1 Then
586         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
588             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
590         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
592         If Userindex <> tU Then
594             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
596         enviarInfoHechizo = True
598         b = True
            
600         UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration

602         If UserList(tU).flags.Inmovilizado = 0 Then
604             UserList(tU).flags.Inmovilizado = 1
606             Call WriteInmovilizaOK(tU)
608             Call WritePosUpdate(tU)
            

            End If

        End If

610     If Hechizos(h).RemoverParalisis = 1 Then
        
            'Para poder tirar remo a un pk en el ring
612         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
614             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
616                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
618                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
620                     b = False
                        Exit Sub

                    End If

622                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
624                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
626                     b = False
                        Exit Sub
                    Else
628                     Call VolverCriminal(Userindex)

                    End If

                End If
            
            End If

630         If UserList(tU).flags.Inmovilizado = 1 Then
632             UserList(tU).Counters.Inmovilizado = 0
634             UserList(tU).flags.Inmovilizado = 0
636             Call WriteInmovilizaOK(tU)
638             enviarInfoHechizo = True
            
640             b = True

            End If

642         If UserList(tU).flags.Paralizado = 1 Then
644             UserList(tU).flags.Paralizado = 0
                'no need to crypt this
646             Call WriteParalizeOK(tU)
648             enviarInfoHechizo = True
            
650             b = True

            End If

        End If

652     If Hechizos(h).Ceguera = 1 Then
654         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
656             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
658         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
660         If Userindex <> tU Then
662             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

664         UserList(tU).flags.Ceguera = 1
666         UserList(tU).Counters.Ceguera = Hechizos(h).Duration

668         Call WriteBlind(tU)
        
670         enviarInfoHechizo = True
672         b = True

        End If

674     If Hechizos(h).Estupidez = 1 Then
676         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
678             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

680         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
682         If Userindex <> tU Then
684             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

686         If UserList(tU).flags.Estupidez = 0 Then
688             UserList(tU).flags.Estupidez = 1
690             UserList(tU).Counters.Estupidez = Hechizos(h).Duration

            End If

692         Call WriteDumb(tU)
        

694         enviarInfoHechizo = True
696         b = True

        End If

698     If Hechizos(h).Velocidad > 0 Then

700         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
702         If Userindex <> tU Then
704             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
706         enviarInfoHechizo = True
708         b = True
            
710         If UserList(tU).Counters.Velocidad = 0 Then
712             UserList(tU).flags.VelocidadBackup = UserList(tU).Char.speeding

            End If

714         UserList(tU).Char.speeding = Hechizos(h).Velocidad
716         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSpeedingACT(UserList(tU).Char.CharIndex, UserList(tU).Char.speeding))
            
718         UserList(tU).Counters.Velocidad = Hechizos(h).Duration

        End If

720     If enviarInfoHechizo Then
722         Call InfoHechizo(Userindex)

        End If

    

        
        Exit Sub

HechizoCombinados_Err:
724     Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoCombinados", Erl)
726     Resume Next
        
End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal slot As Byte)
        
        On Error GoTo UpdateUserHechizos_Err
        

        'Call LogTarea("Sub UpdateUserHechizos")

        Dim LoopC As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(Userindex).Stats.UserHechizos(slot) > 0 Then
104             Call ChangeUserHechizo(Userindex, slot, UserList(Userindex).Stats.UserHechizos(slot))
            Else
106             Call ChangeUserHechizo(Userindex, slot, 0)

            End If

        Else

            'Actualiza todos los slots
108         For LoopC = 1 To MAXUSERHECHIZOS

                'Actualiza el inventario
110             If UserList(Userindex).Stats.UserHechizos(LoopC) > 0 Then
112                 Call ChangeUserHechizo(Userindex, LoopC, UserList(Userindex).Stats.UserHechizos(LoopC))
                Else
114                 Call ChangeUserHechizo(Userindex, LoopC, 0)

                End If

116         Next LoopC

        End If

        
        Exit Sub

UpdateUserHechizos_Err:
118     Call RegistrarError(Err.Number, Err.description, "modHechizos.UpdateUserHechizos", Erl)
120     Resume Next
        
End Sub

Sub ChangeUserHechizo(ByVal Userindex As Integer, ByVal slot As Byte, ByVal Hechizo As Integer)
        
        On Error GoTo ChangeUserHechizo_Err
        

        'Call LogTarea("ChangeUserHechizo")
    
100     UserList(Userindex).Stats.UserHechizos(slot) = Hechizo
    
102     If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
104         Call WriteChangeSpellSlot(Userindex, slot)
        Else
106         Call WriteChangeSpellSlot(Userindex, slot)

        End If

        
        Exit Sub

ChangeUserHechizo_Err:
108     Call RegistrarError(Err.Number, Err.description, "modHechizos.ChangeUserHechizo", Erl)
110     Resume Next
        
End Sub

Public Sub DesplazarHechizo(ByVal Userindex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)
        
        On Error GoTo DesplazarHechizo_Err
        

100     If (Dire <> 1 And Dire <> -1) Then Exit Sub
102     If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

        Dim TempHechizo As Integer

104     If Dire = 1 Then 'Mover arriba
106         If CualHechizo = 1 Then
108             Call WriteConsoleMsg(Userindex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
110             TempHechizo = UserList(Userindex).Stats.UserHechizos(CualHechizo)
112             UserList(Userindex).Stats.UserHechizos(CualHechizo) = UserList(Userindex).Stats.UserHechizos(CualHechizo - 1)
114             UserList(Userindex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
116             If UserList(Userindex).flags.Hechizo > 0 Then
118                 UserList(Userindex).flags.Hechizo = UserList(Userindex).flags.Hechizo - 1

                End If

            End If

        Else 'mover abajo

120         If CualHechizo = MAXUSERHECHIZOS Then
122             Call WriteConsoleMsg(Userindex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
124             TempHechizo = UserList(Userindex).Stats.UserHechizos(CualHechizo)
126             UserList(Userindex).Stats.UserHechizos(CualHechizo) = UserList(Userindex).Stats.UserHechizos(CualHechizo + 1)
128             UserList(Userindex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
130             If UserList(Userindex).flags.Hechizo > 0 Then
132                 UserList(Userindex).flags.Hechizo = UserList(Userindex).flags.Hechizo + 1

                End If

            End If

        End If

        
        Exit Sub

DesplazarHechizo_Err:
134     Call RegistrarError(Err.Number, Err.description, "modHechizos.DesplazarHechizo", Erl)
136     Resume Next
        
End Sub

Sub AreaHechizo(Userindex As Integer, NpcIndex As Integer, X As Byte, Y As Byte, npc As Boolean)
        
        On Error GoTo AreaHechizo_Err
        

        Dim calculo      As Integer

        Dim TilesDifUser As Integer

        Dim TilesDifNpc  As Integer

        Dim tilDif       As Integer

        Dim h2           As Integer

        Dim Hit          As Integer

        Dim daño As Integer

        Dim porcentajeDesc As Integer

100     h2 = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)

        'Calculo de descuesto de golpe por cercania.
102     TilesDifUser = X + Y

104     If npc Then
106         If Hechizos(h2).SubeHP = 2 Then
108             TilesDifNpc = Npclist(NpcIndex).Pos.X + Npclist(NpcIndex).Pos.Y
            
110             tilDif = TilesDifUser - TilesDifNpc
            
112             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

114             Hit = Hit + Porcentaje(Hit, 3 * UserList(Userindex).Stats.ELV)
            
                ' Daño mágico arma
116             If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
118                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                End If
                
                ' Daño mágico anillo
120             If UserList(Userindex).Invent.AnilloEqpObjIndex > 0 Then
122                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(Userindex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
                End If

                ' Disminuir daño con distancia
124             If tilDif <> 0 Then
126                 porcentajeDesc = Abs(tilDif) * 20
128                 daño = Hit / 100 * porcentajeDesc
130                 daño = Hit - daño
                Else
132                 daño = Hit
                End If
                
                ' Si el hechizo no ignora la RM
134             If Hechizos(h2).AntiRm = 0 Then
136                 daño = daño - Npclist(NpcIndex).Stats.defM
                End If
                
                ' Prevengo daño negativo
138             If daño < 0 Then daño = 0
            
140             Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - daño
            
142             If UserList(Userindex).ChatCombate = 1 Then
144                 Call WriteConsoleMsg(Userindex, "Le has causado " & daño & " puntos de daño a " & Npclist(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)

                End If
            
146             Call CalcularDarExp(Userindex, NpcIndex, daño)
                
148             If Npclist(NpcIndex).Stats.MinHp <= 0 Then
                    'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + Npclist(NpcIndex).GiveEXP
                    'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Npclist(NpcIndex).GiveGLD
150                 Call MuereNpc(NpcIndex, Userindex)
                End If

                Exit Sub

            End If

        Else

152         TilesDifNpc = UserList(NpcIndex).Pos.X + UserList(NpcIndex).Pos.Y
154         tilDif = TilesDifUser - TilesDifNpc

156         If Hechizos(h2).SubeHP = 2 Then
158             If Userindex = NpcIndex Then
                    Exit Sub
                End If

160             If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
                
162             If Userindex <> NpcIndex Then
164                 Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

                End If
                
166             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

168             Hit = Hit + Porcentaje(Hit, 3 * UserList(Userindex).Stats.ELV)
            
                ' Daño mágico arma
170             If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
172                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                End If
                
                ' Daño mágico anillo
174             If UserList(Userindex).Invent.AnilloEqpObjIndex > 0 Then
176                 Hit = Hit + Porcentaje(Hit, ObjData(UserList(Userindex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
                End If

178             If tilDif <> 0 Then
180                 porcentajeDesc = Abs(tilDif) * 20
182                 daño = Hit / 100 * porcentajeDesc
184                 daño = Hit - daño
                Else
186                 daño = Hit
                End If
                
                ' Si el hechizo no ignora la RM
188             If Hechizos(h2).AntiRm = 0 Then
                    ' Resistencia mágica armadura
190                 If UserList(NpcIndex).Invent.ArmourEqpObjIndex > 0 Then
192                     daño = daño - Porcentaje(daño, ObjData(UserList(NpcIndex).Invent.ArmourEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica anillo
194                 If UserList(NpcIndex).Invent.AnilloEqpObjIndex > 0 Then
196                     daño = daño - Porcentaje(daño, ObjData(UserList(NpcIndex).Invent.AnilloEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica escudo
198                 If UserList(NpcIndex).Invent.EscudoEqpObjIndex > 0 Then
200                     daño = daño - Porcentaje(daño, ObjData(UserList(NpcIndex).Invent.EscudoEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica casco
202                 If UserList(NpcIndex).Invent.CascoEqpObjIndex > 0 Then
204                     daño = daño - Porcentaje(daño, ObjData(UserList(NpcIndex).Invent.CascoEqpObjIndex).ResistenciaMagica)
                    End If
                End If
                
                ' Prevengo daño negativo
206             If daño < 0 Then daño = 0

208             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp - daño
                    
210             Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
212             Call WriteConsoleMsg(NpcIndex, UserList(Userindex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
214             Call SubirSkill(NpcIndex, Resistencia)
216             Call WriteUpdateUserStats(NpcIndex)
                
                'Muere
218             If UserList(NpcIndex).Stats.MinHp < 1 Then
                    'Store it!
220                 Call Statistics.StoreFrag(Userindex, NpcIndex)
                        
222                 Call ContarMuerte(NpcIndex, Userindex)
224                 UserList(NpcIndex).Stats.MinHp = 0
226                 Call ActStats(NpcIndex, Userindex)

                    'Call UserDie(NpcIndex)
                End If

            End If
                
228         If Hechizos(h2).SubeHP = 1 Then
230             If (TriggerZonaPelea(Userindex, NpcIndex) <> TRIGGER6_PERMITE) Then
232                 If Status(Userindex) = 1 And Status(NpcIndex) <> 1 Then
                        Exit Sub

                    End If

                End If

234             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

236             If tilDif <> 0 Then
238                 porcentajeDesc = Abs(tilDif) * 20
240                 daño = Hit / 100 * porcentajeDesc
242                 daño = Hit - daño
                Else
244                 daño = Hit

                End If
 
246             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp + daño

248             If UserList(NpcIndex).Stats.MinHp > UserList(NpcIndex).Stats.MaxHp Then UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MaxHp

            End If
 
250         If Userindex <> NpcIndex Then
252             Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
254             Call WriteConsoleMsg(NpcIndex, UserList(Userindex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
256             Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
                    
258         Call WriteUpdateUserStats(NpcIndex)

        End If
                
260     If Hechizos(h2).Envenena > 0 Then
262         If Userindex = NpcIndex Then
                Exit Sub

            End If
                    
264         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
                
266         If Userindex <> NpcIndex Then
268             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If
                    
270         UserList(NpcIndex).flags.Envenenado = Hechizos(h2).Envenena
272         Call WriteConsoleMsg(NpcIndex, UserList(Userindex).name & " te ha envenenado.", FontTypeNames.FONTTYPE_FIGHT)

        End If
                
274     If Hechizos(h2).Paraliza = 1 Then
276         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
278         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
280         If Userindex <> NpcIndex Then
282             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If
            
284         Call WriteConsoleMsg(NpcIndex, "Has sido paralizado.", FontTypeNames.FONTTYPE_INFO)
286         UserList(NpcIndex).Counters.Paralisis = Hechizos(h2).Duration

288         If UserList(NpcIndex).flags.Paralizado = 0 Then
290             UserList(NpcIndex).flags.Paralizado = 1
292             Call WriteParalizeOK(NpcIndex)
            

            End If
            
        End If
                
294     If Hechizos(h2).Inmoviliza = 1 Then
296         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
298         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
300         If Userindex <> NpcIndex Then
302             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If
                    
304         Call WriteConsoleMsg(NpcIndex, "Has sido inmovilizado.", FontTypeNames.FONTTYPE_INFO)
306         UserList(NpcIndex).Counters.Inmovilizado = Hechizos(h2).Duration

308         If UserList(NpcIndex).flags.Inmovilizado = 0 Then
310             UserList(NpcIndex).flags.Inmovilizado = 1
312             Call WriteInmovilizaOK(NpcIndex)
314             Call WritePosUpdate(NpcIndex)
            
            End If

        End If
                
316     If Hechizos(h2).Ceguera = 1 Then
318         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
320         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
322         If Userindex <> NpcIndex Then
324             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If
                    
326         UserList(NpcIndex).flags.Ceguera = 1
328         UserList(NpcIndex).Counters.Ceguera = Hechizos(h2).Duration
330         Call WriteConsoleMsg(NpcIndex, "Te han cegado.", FontTypeNames.FONTTYPE_INFO)
            
332         Call WriteBlind(NpcIndex)
        

        End If
                
334     If Hechizos(h2).Velocidad > 0 Then
    
336         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
338         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
340         If Userindex <> NpcIndex Then
342             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If

344         If UserList(NpcIndex).Counters.Velocidad = 0 Then
346             UserList(NpcIndex).flags.VelocidadBackup = UserList(NpcIndex).Char.speeding

            End If

348         UserList(NpcIndex).Char.speeding = Hechizos(h2).Velocidad
350         Call SendData(SendTarget.ToPCArea, NpcIndex, PrepareMessageSpeedingACT(UserList(NpcIndex).Char.CharIndex, UserList(NpcIndex).Char.speeding))
352         UserList(NpcIndex).Counters.Velocidad = Hechizos(h2).Duration

        End If
                
354     If Hechizos(h2).Maldicion = 1 Then
356         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
358         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
360         If Userindex <> NpcIndex Then
362             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

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
390         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
392         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
394         If Userindex <> NpcIndex Then
396             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

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
446         UserList(NpcIndex).flags.Incinerado = 0
                    
448         If UserList(NpcIndex).flags.Inmovilizado = 1 Then
450             UserList(NpcIndex).Counters.Inmovilizado = 0
452             UserList(NpcIndex).flags.Inmovilizado = 0
454             Call WriteInmovilizaOK(NpcIndex)
            

            End If
                    
456         If UserList(NpcIndex).flags.Paralizado = 1 Then
458             UserList(NpcIndex).flags.Paralizado = 0
460             Call WriteParalizeOK(NpcIndex)
            
                       
            End If
                    
462         If UserList(NpcIndex).flags.Ceguera = 1 Then
464             UserList(NpcIndex).flags.Ceguera = 0
466             Call WriteBlindNoMore(NpcIndex)
            

            End If
                    
468         If UserList(NpcIndex).flags.Maldicion = 1 Then
470             UserList(NpcIndex).flags.Maldicion = 0
472             UserList(NpcIndex).Counters.Maldicion = 0

            End If

        End If
        
        
        Exit Sub

AreaHechizo_Err:
474     Call RegistrarError(Err.Number, Err.description, "modHechizos.AreaHechizo", Erl)
476     Resume Next
        
End Sub
