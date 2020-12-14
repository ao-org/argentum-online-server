Attribute VB_Name = "Acciones"

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
'for more information about ORE please visit http://www.baronsoft.com/
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

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

        On Error Resume Next

        '¿Rango Visión? (ToxicWaste)
100     If (Abs(UserList(Userindex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(Userindex).Pos.X - X) > RANGO_VISION_X) Then
            Exit Sub

        End If

        '¿Posicion valida?
102     If InMapBounds(Map, X, Y) Then
   
            Dim FoundChar      As Byte

            Dim FoundSomething As Byte

            Dim TempCharIndex  As Integer
       
104         If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                'Set the target NPC
106             UserList(Userindex).flags.TargetNPC = MapData(Map, X, Y).NpcIndex
108             UserList(Userindex).flags.TargetNpcTipo = Npclist(MapData(Map, X, Y).NpcIndex).NPCtype
        
110             If Npclist(MapData(Map, X, Y).NpcIndex).Comercia = 1 Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
112                 If UserList(Userindex).flags.Muerto = 1 Then
114                     Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
116                 If UserList(Userindex).flags.Comerciando Then
                        Exit Sub

                    End If
            
118                 If Distancia(Npclist(UserList(Userindex).flags.TargetNPC).Pos, UserList(Userindex).Pos) > 6 Then
120                     Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Iniciamos la rutina pa' comerciar.
122                 Call IniciarComercioNPC(Userindex)
        
124             ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Banquero Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
126                 If UserList(Userindex).flags.Muerto = 1 Then
128                     Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
130                 If UserList(Userindex).flags.Comerciando Then
                        Exit Sub

                    End If
            
132                 If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).Pos, UserList(Userindex).Pos) > 6 Then
134                     Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'A depositar de una
136                 Call IniciarBanco(Userindex)
            
138             ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Pirata Then  'VIAJES

                    '¿Esta el user muerto? Si es asi no puede comerciar
140                 If UserList(Userindex).flags.Muerto = 1 Then
142                     Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
144                 If UserList(Userindex).flags.Comerciando Then
                        Exit Sub

                    End If
            
146                 If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).Pos, UserList(Userindex).Pos) > 5 Then
148                     Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
150                     Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
152                 If Npclist(MapData(Map, X, Y).NpcIndex).SoundOpen <> 0 Then
154                     Call WritePlayWave(Userindex, Npclist(MapData(Map, X, Y).NpcIndex).SoundOpen, NO_3D_SOUND, NO_3D_SOUND)

                    End If

                    'A depositar de unaIniciarTransporte
156                 Call WriteViajarForm(Userindex, MapData(Map, X, Y).NpcIndex)
                    Exit Sub
            
158             ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.ResucitadorNewbie Then

160                 If Distancia(UserList(Userindex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 5 Then
                        'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
162                     Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
164                 UserList(Userindex).flags.Envenenado = 0
166                 UserList(Userindex).flags.Incinerado = 0
      
                    'Revivimos si es necesario
168                 If UserList(Userindex).flags.Muerto = 1 And (Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or EsNewbie(Userindex)) Then
170                     Call WriteConsoleMsg(Userindex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)
172                     Call RevivirUsuario(Userindex)
174                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Resucitar, 30, False))
176                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("204", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                
                    Else

                        'curamos totalmente
178                     If UserList(Userindex).Stats.MinHp <> UserList(Userindex).Stats.MaxHp Then
180                         UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MaxHp
182                         Call WritePlayWave(Userindex, "101", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
                            'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...¡Has sido curado!", FontTypeNames.FONTTYPE_INFO)
184                         Call WriteLocaleMsg(Userindex, "83", FontTypeNames.FONTTYPE_INFOIAO)
                    
186                         Call WriteUpdateUserStats(Userindex)

188                         If Status(Userindex) = 2 Or Status(Userindex) = 0 Then
190                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.CurarCrimi, 100, False))
                            Else
           
192                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Curar, 100, False))

                            End If

                        End If

                    End If
            
                    'Sistema Battle
            
194             ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.BattleModo Then

196                 If Distancia(UserList(Userindex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 5 Then
198                     Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
200                 If BattleActivado = 0 Then
202                     Call WriteChatOverHead(Userindex, "Actualmente el battle se encuentra desactivado.", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                        Exit Sub

                    End If
                        
204                 If UserList(Userindex).clase = eClass.Trabajador Then
206                     Call WriteConsoleMsg(Userindex, "Los trabajadores no pueden ingresar al battle.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
208                 If UserList(Userindex).Stats.ELV < BattleMinNivel Then
210                     Call WriteConsoleMsg(Userindex, "Exclusivo para personajes superiores a nivel " & BattleMinNivel, FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
212                 If UserList(Userindex).flags.Comerciando Then
214                     Call WriteConsoleMsg(Userindex, "No podes ingresar al battle si estas comerciando.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
216                 If UserList(Userindex).flags.EnTorneo = True Then
218                     Call WriteConsoleMsg(Userindex, "No podes ingresar al battle estando anotado en un evento.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
220                 If UserList(Userindex).Accion.TipoAccion = Accion_Barra.BattleModo Then Exit Sub
222                 If UserList(Userindex).donador.activo = 0 Then
224                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, 400, False))
226                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 400, Accion_Barra.BattleModo))
                    Else
228                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, 50, False))
230                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 50, Accion_Barra.BattleModo))

                    End If

232                 UserList(Userindex).Accion.AccionPendiente = True
234                 UserList(Userindex).Accion.Particula = ParticulasIndex.Runa
236                 UserList(Userindex).Accion.TipoAccion = Accion_Barra.BattleModo
            
                    'Sistema Battle
         
238             ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Subastador Then

240                 If UserList(Userindex).flags.Muerto = 1 Then
242                     Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
244                 If Distancia(Npclist(UserList(Userindex).flags.TargetNPC).Pos, UserList(Userindex).Pos) > 1 Then
                        ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del subastador.", FontTypeNames.FONTTYPE_INFO)
246                     Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

248                 Call IniciarSubasta(Userindex)
            
250             ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Quest Then

252                 If UserList(Userindex).flags.Muerto = 1 Then
254                     Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
256                 Call EnviarQuest(Userindex)
            
258             ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Enlistador Then

260                 If UserList(Userindex).flags.Muerto = 1 Then
262                     Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
264                 If Distancia(Npclist(UserList(Userindex).flags.TargetNPC).Pos, UserList(Userindex).Pos) > 4 Then
266                     Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
        
268                 If Npclist(UserList(Userindex).flags.TargetNPC).flags.Faccion = 0 Then
270                     If UserList(Userindex).Faccion.ArmadaReal = 0 Then
272                         Call EnlistarArmadaReal(Userindex)
                        Else
274                         Call RecompensaArmadaReal(Userindex)

                        End If

                    Else

276                     If UserList(Userindex).Faccion.FuerzasCaos = 0 Then
278                         Call EnlistarCaos(Userindex)
                        Else
280                         Call RecompensaCaos(Userindex)

                        End If

                    End If

282             ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Gobernador Then

284                 If UserList(Userindex).flags.Muerto = 1 Then
286                     Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
288                 If Distancia(Npclist(UserList(Userindex).flags.TargetNPC).Pos, UserList(Userindex).Pos) > 3 Then
290                     Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del gobernador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Dim DeDonde As String
            
292                 If UserList(Userindex).Hogar = Npclist(UserList(Userindex).flags.TargetNPC).GobernadorDe Then
294                     Call WriteChatOverHead(Userindex, "Ya perteneces a esta ciudad. Gracias por ser uno más de nosotros.", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                        Exit Sub

                    End If
            
296                 If UserList(Userindex).Faccion.Status = 0 Or UserList(Userindex).Faccion.Status = 2 Then
298                     If Npclist(UserList(Userindex).flags.TargetNPC).GobernadorDe = eCiudad.cBanderbill Then
300                         Call WriteChatOverHead(Userindex, "Aquí no aceptamos criminales.", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                            Exit Sub

                        End If

                    End If
            
302                 If UserList(Userindex).Faccion.Status = 3 Or UserList(Userindex).Faccion.Status = 1 Then
304                     If Npclist(UserList(Userindex).flags.TargetNPC).GobernadorDe = eCiudad.cArghal Then
306                         Call WriteChatOverHead(Userindex, "¡¡Sal de aquí ciudadano asqueroso!!", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                            Exit Sub

                        End If

                    End If
            
308                 If UserList(Userindex).Hogar <> Npclist(UserList(Userindex).flags.TargetNPC).GobernadorDe Then
            
310                     UserList(Userindex).PosibleHogar = Npclist(UserList(Userindex).flags.TargetNPC).GobernadorDe
                
312                     Select Case UserList(Userindex).PosibleHogar

                            Case eCiudad.cUllathorpe
314                             DeDonde = "Ullathorpe"
                            
316                         Case eCiudad.cNix
318                             DeDonde = "Nix"
                
320                         Case eCiudad.cBanderbill
322                             DeDonde = "Banderbill"
                        
324                         Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
326                             DeDonde = "Lindos"
                            
328                         Case eCiudad.cArghal
330                             DeDonde = " Arghal"
                            
332                         Case eCiudad.CHillidan
334                             DeDonde = " Hillidan"
                            
336                         Case Else
338                             DeDonde = "Ullathorpe"

                        End Select
                    
340                     UserList(Userindex).flags.pregunta = 3
342                     Call WritePreguntaBox(Userindex, "¿Te gustaria ser ciudadano de " & DeDonde & "?")
                
                    End If

                End If
        
                '¿Es un obj?
344         ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
346             UserList(Userindex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        
348             Select Case ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
350                     Call AccionParaPuerta(Map, X, Y, Userindex)

352                 Case eOBJType.otCarteles 'Es un cartel
354                     Call AccionParaCartel(Map, X, Y, Userindex)

356                 Case eOBJType.OtCorreo 'Es un cartel
                        'Call AccionParaCorreo(Map, x, Y, UserIndex)

358                 Case eOBJType.otForos 'Foro
                        'Call AccionParaForo(Map, X, Y, UserIndex)
360                     Call WriteConsoleMsg(Userindex, "El foro está temporalmente deshabilitado.", FontTypeNames.FONTTYPE_EJECUCION)

362                 Case eOBJType.OtPozos 'Pozos
                        'Call AccionParaPozos(Map, x, Y, UserIndex)

364                 Case eOBJType.otArboles 'Pozos
                        'Call AccionParaArboles(Map, x, Y, UserIndex)

366                 Case eOBJType.otYunque 'Pozos
368                     Call AccionParaYunque(Map, X, Y, Userindex)

370                 Case eOBJType.otLeña    'Leña

372                     If MapData(Map, X, Y).ObjInfo.ObjIndex = FOGATA_APAG And UserList(Userindex).flags.Muerto = 0 Then
374                         Call AccionParaRamita(Map, X, Y, Userindex)

                        End If

                End Select

                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
376         ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
378             UserList(Userindex).flags.TargetObj = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
        
380             Select Case ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
382                     Call AccionParaPuerta(Map, X + 1, Y, Userindex)
            
                End Select

384         ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
386             UserList(Userindex).flags.TargetObj = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex

388             Select Case ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
390                     Call AccionParaPuerta(Map, X + 1, Y + 1, Userindex)
            
                End Select

392         ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
394             UserList(Userindex).flags.TargetObj = MapData(Map, X, Y + 1).ObjInfo.ObjIndex

396             Select Case ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
398                     Call AccionParaPuerta(Map, X, Y + 1, Userindex)

                End Select

            'ElseIf HayAgua(Map, x, Y) Then
                'Call AccionParaAgua(Map, x, Y, UserIndex)

            End If

        End If

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
108         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        '¿Hay mensajes?
        Dim f As String, tit As String, men As String, BASE As String, auxcad As String

110     f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).ForoID) & ".for"

112     If FileExist(f, vbNormal) Then

            Dim num As Integer

114         num = val(GetVar(f, "INFO", "CantMSG"))
116         BASE = Left$(f, Len(f) - 4)

            Dim i As Integer

            Dim n As Integer

118         For i = 1 To num
120             n = FreeFile
122             f = BASE & i & ".for"
124             Open f For Input Shared As #n
126             Input #n, tit
128             men = vbNullString
130             auxcad = vbNullString

132             Do While Not EOF(n)
134                 Input #n, auxcad
136                 men = men & vbCrLf & auxcad
                Loop
138             Close #n
140             Call WriteAddForumMsg(Userindex, tit, men)
        
            Next

        End If

142     Call WriteShowForumForm(Userindex)

End Sub

Sub AccionParaPozos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
108         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

110     If MapData(Map, X, Y).ObjInfo.Amount <= 1 Then
112         Call WriteConsoleMsg(Userindex, "El pozo esta drenado, regresa mas tarde...", FontTypeNames.FONTTYPE_EJECUCION)
            Exit Sub

        End If

114     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 1 Then
116         If UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN Then
118             Call WriteConsoleMsg(Userindex, "No tenes necesidad del pozo...", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

120         UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
122         MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
124         Call WriteConsoleMsg(Userindex, "Sientes la frescura del pozo. ¡Tu maná a sido restaurada!", FontTypeNames.FONTTYPE_EJECUCION)
126         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
128         Call WriteUpdateUserStats(Userindex)
            Exit Sub

        End If

130     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 2 Then
132         If UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU Then
134             Call WriteConsoleMsg(Userindex, "No tenes necesidad del pozo...", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

136         UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
138         UserList(Userindex).flags.Sed = 0 'Bug reparado 27/01/13
140         MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
142         Call WriteConsoleMsg(Userindex, "Sientes la frescura del pozo. ¡Ya no sientes sed!", FontTypeNames.FONTTYPE_EJECUCION)
144         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
146         Call WriteUpdateHungerAndThirst(Userindex)
            Exit Sub

        End If

End Sub

Sub AccionParaArboles(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
108         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

110     If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
112         Call WriteConsoleMsg(Userindex, "Esta prohibido manipular árboles en las ciudades.", FontTypeNames.FONTTYPE_INFOIAO)
114         Call WriteWorkRequestTarget(Userindex, 0)
            Exit Sub

        End If

116     If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) < 40 Then
118         Call WriteConsoleMsg(Userindex, "No tenes suficientes conocimientos para comer del arbol. Necesitas al menos 40 skill en supervivencia.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

120     If MapData(Map, X, Y).ObjInfo.Amount <= 1 Then
122         Call WriteConsoleMsg(Userindex, "El árbol no tiene más frutos para dar.", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

124     If UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MaxHam Then
126         Call WriteConsoleMsg(Userindex, "No tenes hambre.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

128     UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MinHam + 5
130     UserList(Userindex).Stats.MaxHam = 100
132     UserList(Userindex).flags.Hambre = 0 'Bug reparado 27/01/13
    
134     MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
    
136     If Not UserList(Userindex).flags.UltimoMensaje = 40 Then
138         Call WriteConsoleMsg(Userindex, "Logras conseguir algunos frutos del árbol, ya no sientes tanta hambre.", FontTypeNames.FONTTYPE_INFOIAO)
140         UserList(Userindex).flags.UltimoMensaje = 40

        End If
    
142     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(e_SoundIndex.MORFAR_MANZANA, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
144     Call WriteUpdateHungerAndThirst(Userindex)

End Sub

Sub AccionParaAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
108         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

110     If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
112         Call WriteConsoleMsg(Userindex, "Esta prohibido beber agua en las orillas de las ciudades.", FontTypeNames.FONTTYPE_INFO)
114         Call WriteWorkRequestTarget(Userindex, 0)
            Exit Sub

        End If

116     If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) < 30 Then
118         Call WriteConsoleMsg(Userindex, "No tenes suficientes conocimientos para beber del agua. Necesitas al menos 30 skill en supervivencia.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

120     If UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU Then
122         Call WriteConsoleMsg(Userindex, "No tenes sed.", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

124     UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU + 5
126     UserList(Userindex).flags.Sed = 0 'Bug reparado 27/01/13
    
128     If Not UserList(Userindex).flags.UltimoMensaje = 41 Then
130         Call WriteConsoleMsg(Userindex, "Has bebido, ya no sientes tanta sed.", FontTypeNames.FONTTYPE_INFOIAO)
132         UserList(Userindex).flags.UltimoMensaje = 41

        End If
    
134     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
136     Call WriteUpdateHungerAndThirst(Userindex)

End Sub

Sub AccionParaYunque(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
108         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        ' Herramientas: SubTipo 7 - Martillo de Herrero
110     If ObjData(UserList(Userindex).Invent.HerramientaEqpObjIndex).Subtipo <> 7 Then
            'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
112         Call WriteConsoleMsg(Userindex, "Antes debes tener equipado un martillo de herrero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

114     Call EnivarArmasConstruibles(Userindex)
116     Call EnivarArmadurasConstruibles(Userindex)
118     Call WriteShowBlacksmithForm(Userindex)

        'UserList(UserIndex).Invent.HerramientaEqpObjIndex = objindex
        'UserList(UserIndex).Invent.HerramientaEqpSlot = slot

End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Userindex As Integer)

        On Error Resume Next

        Dim MiObj As obj

        Dim wp    As WorldPos

100     If Not (Distance(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, X, Y) > 2) Then
102         If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
104             If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then

                    'Abre la puerta
106                 If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
108                     MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
110                     Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(MapData(Map, X, Y).ObjInfo.ObjIndex, X, Y))
                    
112                     Call BloquearPuerta(Map, X, Y, False)
                      
                        'Sonido
114                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                    Else
116                     Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    'Cierra puerta
118                 MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
120                 Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(MapData(Map, X, Y).ObjInfo.ObjIndex, X, Y))
                                
122                 Call BloquearPuerta(Map, X, Y, True)

124                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, X, Y))

                End If
        
126             UserList(Userindex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
            Else
128             Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)

            End If

        Else
130         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)

            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim MiObj As obj

100     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
102         If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
104             Call WriteShowSignal(Userindex, MapData(Map, X, Y).ObjInfo.ObjIndex)

            End If
  
        End If

End Sub

Sub AccionParaCorreo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
        
        On Error GoTo AccionParaCorreo_Err
        

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If UserList(Userindex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto.", FontTypeNames.FONTTYPE_INFO)
108         Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

110     If Distancia(Pos, UserList(Userindex).Pos) > 4 Then
112         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

114     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 47 Then
116         Call WriteListaCorreo(Userindex, False)

        End If

        
        Exit Sub

AccionParaCorreo_Err:
118     Call RegistrarError(Err.Number, Err.description, "Argentum20Server.Acciones.AccionParaCorreo", Erl)
120     Resume Next
        
End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim Suerte As Byte
        Dim exito  As Byte
        Dim raise  As Integer
    
        Dim Pos    As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y
    
106     With UserList(Userindex)
    
108         If Distancia(Pos, .Pos) > 2 Then
110             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

112         If MapInfo(Map).lluvia And Lloviendo Then
114             Call WriteConsoleMsg(Userindex, "Esta lloviendo, no podés encender una fogata aquí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

116         If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Seguro = 1 Then
118             Call WriteConsoleMsg(Userindex, "En zona segura no podés hacer fogatas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

120         If MapData(Map, X - 1, Y).ObjInfo.ObjIndex = FOGATA Or _
               MapData(Map, X + 1, Y).ObjInfo.ObjIndex = FOGATA Or _
               MapData(Map, X, Y - 1).ObjInfo.ObjIndex = FOGATA Or _
               MapData(Map, X, Y + 1).ObjInfo.ObjIndex = FOGATA Then
           
122             Call WriteConsoleMsg(Userindex, "Debes alejarte un poco de la otra fogata.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

124         If .Stats.UserSkills(Supervivencia) > 1 And .Stats.UserSkills(Supervivencia) < 6 Then
126             Suerte = 3
        
128         ElseIf .Stats.UserSkills(Supervivencia) >= 6 And .Stats.UserSkills(Supervivencia) <= 10 Then
130             Suerte = 2
        
132         ElseIf .Stats.UserSkills(Supervivencia) >= 10 And .Stats.UserSkills(Supervivencia) Then
134             Suerte = 1

            End If

136         exito = RandomNumber(1, Suerte)

138         If exito = 1 Then
    
140             If MapInfo(.Pos.Map).zone <> Ciudad Then
                
                    Dim obj As obj
142                 obj.ObjIndex = FOGATA
144                 obj.Amount = 1
        
146                 Call WriteConsoleMsg(Userindex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
        
148                 Call MakeObj(obj, Map, X, Y)

                Else
        
150                 Call WriteConsoleMsg(Userindex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
            
                    Exit Sub

                End If

            Else
        
152             Call WriteConsoleMsg(Userindex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)

            End If
    
        End With

154     Call SubirSkill(Userindex, Supervivencia)

End Sub
