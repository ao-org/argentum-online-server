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

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo Accion_Err
    
        

        

        '¿Rango Visión? (ToxicWaste)
100     If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
            Exit Sub

        End If

        '¿Posicion valida?
102     If InMapBounds(Map, X, Y) Then
   
            Dim FoundChar      As Byte

            Dim FoundSomething As Byte

            Dim TempCharIndex  As Integer
       
104         If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                'Set the target NPC
106             UserList(UserIndex).flags.TargetNPC = MapData(Map, X, Y).NpcIndex
108             UserList(UserIndex).flags.TargetNpcTipo = NpcList(MapData(Map, X, Y).NpcIndex).NPCtype
        
110             If NpcList(MapData(Map, X, Y).NpcIndex).Comercia = 1 Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
112                 If UserList(UserIndex).flags.Muerto = 1 Then
114                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
116                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub

                    End If
            
118                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 6 Then
120                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
                    If NpcList(MapData(Map, X, Y).NpcIndex).Movement = Caminata Then
                        NpcList(MapData(Map, X, Y).NpcIndex).Contadores.IntervaloMovimiento = GetTickCount + 15000 - NpcList(MapData(Map, X, Y).NpcIndex).IntervaloMovimiento ' 15 segundos
                    End If
            
                    'Iniciamos la rutina pa' comerciar.
122                 Call IniciarComercioNPC(UserIndex)
        
124             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Banquero Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
126                 If UserList(UserIndex).flags.Muerto = 1 Then
128                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
130                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub

                    End If
            
132                 If Distancia(NpcList(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 6 Then
134                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'A depositar de una
136                 Call IniciarBanco(UserIndex)
            
138             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Pirata Then  'VIAJES

                    '¿Esta el user muerto? Si es asi no puede comerciar
140                 If UserList(UserIndex).flags.Muerto = 1 Then
142                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
144                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub

                    End If
            
146                 If Distancia(NpcList(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 5 Then
148                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
150                     Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
152                 If NpcList(MapData(Map, X, Y).NpcIndex).SoundOpen <> 0 Then
154                     Call WritePlayWave(UserIndex, NpcList(MapData(Map, X, Y).NpcIndex).SoundOpen, NO_3D_SOUND, NO_3D_SOUND)

                    End If

                    'A depositar de unaIniciarTransporte
156                 Call WriteViajarForm(UserIndex, MapData(Map, X, Y).NpcIndex)
                    Exit Sub
            
158             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.ResucitadorNewbie Then

160                 If Distancia(UserList(UserIndex).Pos, NpcList(MapData(Map, X, Y).NpcIndex).Pos) > 5 Then
                        'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
162                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
                    If NpcList(MapData(Map, X, Y).NpcIndex).Movement = Caminata Then
                        NpcList(MapData(Map, X, Y).NpcIndex).Contadores.IntervaloMovimiento = GetTickCount + 5000 - NpcList(MapData(Map, X, Y).NpcIndex).IntervaloMovimiento ' 5 segundos
                    End If
            
164                 UserList(UserIndex).flags.Envenenado = 0
166                 UserList(UserIndex).flags.Incinerado = 0
      
                    'Revivimos si es necesario
168                 If UserList(UserIndex).flags.Muerto = 1 And (NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
170                     Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)
172                     Call RevivirUsuario(UserIndex)
174                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 30, False))
176                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                
                    Else

                        'curamos totalmente
178                     If UserList(UserIndex).Stats.MinHp <> UserList(UserIndex).Stats.MaxHp Then
180                         UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
182                         Call WritePlayWave(UserIndex, "117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                            'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...¡Has sido curado!", FontTypeNames.FONTTYPE_INFO)
184                         Call WriteLocaleMsg(UserIndex, "83", FontTypeNames.FONTTYPE_INFOIAO)
                    
186                         Call WriteUpdateUserStats(UserIndex)

188                         If Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
190                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.CurarCrimi, 100, False))
                            Else
           
192                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Curar, 100, False))

                            End If

                        End If

                    End If
            
                    'Sistema Battle
            
194             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.BattleModo Then

196                 If Distancia(UserList(UserIndex).Pos, NpcList(MapData(Map, X, Y).NpcIndex).Pos) > 5 Then
198                     Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
200                 If BattleActivado = 0 Then
202                     Call WriteChatOverHead(UserIndex, "Actualmente el battle se encuentra desactivado.", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                        Exit Sub

                    End If
                        
204                 If UserList(UserIndex).clase = eClass.Trabajador Then
206                     Call WriteConsoleMsg(UserIndex, "Los trabajadores no pueden ingresar al battle.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
208                 If UserList(UserIndex).Stats.ELV < BattleMinNivel Then
210                     Call WriteConsoleMsg(UserIndex, "Exclusivo para personajes superiores a nivel " & BattleMinNivel, FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
212                 If UserList(UserIndex).flags.Comerciando Then
214                     Call WriteConsoleMsg(UserIndex, "No podes ingresar al battle si estas comerciando.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
216                 If UserList(UserIndex).flags.EnTorneo = True Then
218                     Call WriteConsoleMsg(UserIndex, "No podes ingresar al battle estando anotado en un evento.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
220                 If UserList(UserIndex).Accion.TipoAccion = Accion_Barra.BattleModo Then Exit Sub
222                 If UserList(UserIndex).donador.activo = 0 Then
224                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 400, False))
226                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 400, Accion_Barra.BattleModo))
                    Else
228                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 50, False))
230                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 50, Accion_Barra.BattleModo))

                    End If

232                 UserList(UserIndex).Accion.AccionPendiente = True
234                 UserList(UserIndex).Accion.Particula = ParticulasIndex.Runa
236                 UserList(UserIndex).Accion.TipoAccion = Accion_Barra.BattleModo
            
                    'Sistema Battle
         
238             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Subastador Then

240                 If UserList(UserIndex).flags.Muerto = 1 Then
242                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
244                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 1 Then
                        ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del subastador.", FontTypeNames.FONTTYPE_INFO)
246                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
                    If NpcList(MapData(Map, X, Y).NpcIndex).Movement = Caminata Then
                        NpcList(MapData(Map, X, Y).NpcIndex).Contadores.IntervaloMovimiento = GetTickCount + 20000 - NpcList(MapData(Map, X, Y).NpcIndex).IntervaloMovimiento ' 20 segundos
                    End If

248                 Call IniciarSubasta(UserIndex)
            
250             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Quest Then

252                 If UserList(UserIndex).flags.Muerto = 1 Then
254                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
                    If NpcList(MapData(Map, X, Y).NpcIndex).Movement = Caminata Then
                        NpcList(MapData(Map, X, Y).NpcIndex).Contadores.IntervaloMovimiento = GetTickCount + 15000 - NpcList(MapData(Map, X, Y).NpcIndex).IntervaloMovimiento ' 15 segundos
                    End If
            
256                 Call EnviarQuest(UserIndex)
            
258             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Enlistador Then

260                 If UserList(UserIndex).flags.Muerto = 1 Then
262                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
264                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 4 Then
266                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
        
268                 If NpcList(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
270                     If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
272                         Call EnlistarArmadaReal(UserIndex)
                        Else
274                         Call RecompensaArmadaReal(UserIndex)

                        End If

                    Else

276                     If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
278                         Call EnlistarCaos(UserIndex)
                        Else
280                         Call RecompensaCaos(UserIndex)

                        End If

                    End If

282             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Gobernador Then

284                 If UserList(UserIndex).flags.Muerto = 1 Then
286                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
288                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
290                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del gobernador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Dim DeDonde As String
                    Dim Gobernador As npc
                        Gobernador = NpcList(UserList(UserIndex).flags.TargetNPC)
            
292                 If UserList(UserIndex).Hogar = Gobernador.GobernadorDe Then
294                     Call WriteChatOverHead(UserIndex, "Ya perteneces a esta ciudad. Gracias por ser uno más de nosotros.", Gobernador.Char.CharIndex, vbWhite)
                        Exit Sub

                    End If
            
296                 If UserList(UserIndex).Faccion.Status = 0 Or UserList(UserIndex).Faccion.Status = 2 Then
298                     If Gobernador.GobernadorDe = eCiudad.cBanderbill Then
300                         Call WriteChatOverHead(UserIndex, "Aquí no aceptamos criminales.", Gobernador.Char.CharIndex, vbWhite)
                            Exit Sub

                        End If

                    End If
            
302                 If UserList(UserIndex).Faccion.Status = 3 Or UserList(UserIndex).Faccion.Status = 1 Then
304                     If Gobernador.GobernadorDe = eCiudad.cArkhein Then
306                         Call WriteChatOverHead(UserIndex, "¡¡Sal de aquí ciudadano asqueroso!!", Gobernador.Char.CharIndex, vbWhite)
                            Exit Sub

                        End If

                    End If
            
308                 If UserList(UserIndex).Hogar <> Gobernador.GobernadorDe Then
            
310                     UserList(UserIndex).PosibleHogar = Gobernador.GobernadorDe
                
312                     Select Case UserList(UserIndex).PosibleHogar

                            Case eCiudad.cUllathorpe
314                             DeDonde = "Ullathorpe"
                            
316                         Case eCiudad.cNix
318                             DeDonde = "Nix"
                
320                         Case eCiudad.cBanderbill
322                             DeDonde = "Banderbill"
                        
324                         Case eCiudad.cLindos
326                             DeDonde = "Lindos"
                            
328                         Case eCiudad.cArghal
330                             DeDonde = " Arghal"
                            
332                         Case eCiudad.cArkhein
334                             DeDonde = " Arkhein"

336                         Case Else
338                             DeDonde = "Ullathorpe"

                        End Select
                    
340                     UserList(UserIndex).flags.pregunta = 3
342                     Call WritePreguntaBox(UserIndex, "¿Te gustaria ser ciudadano de " & DeDonde & "?")
                
                    End If

                End If
        
                '¿Es un obj?
344         ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
346             UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        
348             Select Case ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
350                     Call AccionParaPuerta(Map, X, Y, UserIndex)

352                 Case eOBJType.otCarteles 'Es un cartel
354                     Call AccionParaCartel(Map, X, Y, UserIndex)

356                 Case eOBJType.OtCorreo 'Es un cartel
                        'Call AccionParaCorreo(Map, x, Y, UserIndex)

358                 Case eOBJType.otForos 'Foro
                        'Call AccionParaForo(Map, X, Y, UserIndex)
360                     Call WriteConsoleMsg(UserIndex, "El foro está temporalmente deshabilitado.", FontTypeNames.FONTTYPE_EJECUCION)

362                 Case eOBJType.OtPozos 'Pozos
                        'Call AccionParaPozos(Map, x, Y, UserIndex)

364                 Case eOBJType.otArboles 'Pozos
                        'Call AccionParaArboles(Map, x, Y, UserIndex)

366                 Case eOBJType.otYunque 'Pozos
368                     Call AccionParaYunque(Map, X, Y, UserIndex)

370                 Case eOBJType.otLeña    'Leña

372                     If MapData(Map, X, Y).ObjInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
374                         Call AccionParaRamita(Map, X, Y, UserIndex)

                        End If

                End Select

                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
376         ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
378             UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
        
380             Select Case ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
382                     Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
                End Select

384         ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
386             UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex

388             Select Case ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
390                     Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
                End Select

392         ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
394             UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).ObjInfo.ObjIndex

396             Select Case ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
398                     Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

                End Select

            'ElseIf HayAgua(Map, x, Y) Then
                'Call AccionParaAgua(Map, x, Y, UserIndex)

            End If

        End If

        
        Exit Sub

Accion_Err:
400     Call RegistrarError(Err.Number, Err.Description, "Acciones.Accion", Erl)

        
End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo AccionParaForo_Err
    
        

        

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
108         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
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
140             Call WriteAddForumMsg(UserIndex, tit, men)
        
            Next

        End If

142     Call WriteShowForumForm(UserIndex)

        
        Exit Sub

AccionParaForo_Err:
144     Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaForo", Erl)

        
End Sub

Sub AccionParaPozos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo AccionParaPozos_Err
    
        

        

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
108         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

110     If MapData(Map, X, Y).ObjInfo.Amount <= 1 Then
112         Call WriteConsoleMsg(UserIndex, "El pozo esta drenado, regresa mas tarde...", FontTypeNames.FONTTYPE_EJECUCION)
            Exit Sub

        End If

114     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 1 Then
116         If UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then
118             Call WriteConsoleMsg(UserIndex, "No tenes necesidad del pozo...", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

120         UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
122         MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
124         Call WriteConsoleMsg(UserIndex, "Sientes la frescura del pozo. ¡Tu maná a sido restaurada!", FontTypeNames.FONTTYPE_EJECUCION)
126         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
128         Call WriteUpdateUserStats(UserIndex)
            Exit Sub

        End If

130     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 2 Then
132         If UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU Then
134             Call WriteConsoleMsg(UserIndex, "No tenes necesidad del pozo...", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

136         UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
138         UserList(UserIndex).flags.Sed = 0 'Bug reparado 27/01/13
140         MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
142         Call WriteConsoleMsg(UserIndex, "Sientes la frescura del pozo. ¡Ya no sientes sed!", FontTypeNames.FONTTYPE_EJECUCION)
144         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
146         Call WriteUpdateHungerAndThirst(UserIndex)
            Exit Sub

        End If

        
        Exit Sub

AccionParaPozos_Err:
148     Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaPozos", Erl)

        
End Sub

Sub AccionParaArboles(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo AccionParaArboles_Err
    
        

        

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
108         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

110     If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
112         Call WriteConsoleMsg(UserIndex, "Esta prohibido manipular árboles en las ciudades.", FontTypeNames.FONTTYPE_INFOIAO)
114         Call WriteWorkRequestTarget(UserIndex, 0)
            Exit Sub

        End If

116     If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 40 Then
118         Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para comer del arbol. Necesitas al menos 40 skill en supervivencia.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

120     If MapData(Map, X, Y).ObjInfo.Amount <= 1 Then
122         Call WriteConsoleMsg(UserIndex, "El árbol no tiene más frutos para dar.", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

124     If UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam Then
126         Call WriteConsoleMsg(UserIndex, "No tenes hambre.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

128     UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam + 5
130     UserList(UserIndex).Stats.MaxHam = 100
132     UserList(UserIndex).flags.Hambre = 0 'Bug reparado 27/01/13
    
134     MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - 1
    
136     If Not UserList(UserIndex).flags.UltimoMensaje = 40 Then
138         Call WriteConsoleMsg(UserIndex, "Logras conseguir algunos frutos del árbol, ya no sientes tanta hambre.", FontTypeNames.FONTTYPE_INFOIAO)
140         UserList(UserIndex).flags.UltimoMensaje = 40

        End If
    
142     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
144     Call WriteUpdateHungerAndThirst(UserIndex)

        
        Exit Sub

AccionParaArboles_Err:
146     Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaArboles", Erl)

        
End Sub

Sub AccionParaAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo AccionParaAgua_Err
    
        

        

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
108         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

110     If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
112         Call WriteConsoleMsg(UserIndex, "Esta prohibido beber agua en las orillas de las ciudades.", FontTypeNames.FONTTYPE_INFO)
114         Call WriteWorkRequestTarget(UserIndex, 0)
            Exit Sub

        End If

116     If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 30 Then
118         Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para beber del agua. Necesitas al menos 30 skill en supervivencia.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

120     If UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU Then
122         Call WriteConsoleMsg(UserIndex, "No tenes sed.", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

124     UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + 5
126     UserList(UserIndex).flags.Sed = 0 'Bug reparado 27/01/13
    
128     If Not UserList(UserIndex).flags.UltimoMensaje = 41 Then
130         Call WriteConsoleMsg(UserIndex, "Has bebido, ya no sientes tanta sed.", FontTypeNames.FONTTYPE_INFOIAO)
132         UserList(UserIndex).flags.UltimoMensaje = 41

        End If
    
134     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
136     Call WriteUpdateHungerAndThirst(UserIndex)

        
        Exit Sub

AccionParaAgua_Err:
138     Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaAgua", Erl)

        
End Sub

Sub AccionParaYunque(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo AccionParaYunque_Err
    
        

        

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
108         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        ' Herramientas: SubTipo 7 - Martillo de Herrero
110     If ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo <> 7 Then
            'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
112         Call WriteConsoleMsg(UserIndex, "Antes debes tener equipado un martillo de herrero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

114     Call EnivarArmasConstruibles(UserIndex)
116     Call EnivarArmadurasConstruibles(UserIndex)
118     Call WriteShowBlacksmithForm(UserIndex)

        'UserList(UserIndex).Invent.HerramientaEqpObjIndex = objindex
        'UserList(UserIndex).Invent.HerramientaEqpSlot = slot

        
        Exit Sub

AccionParaYunque_Err:
120     Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaYunque", Erl)

        
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal UserIndex As Integer, Optional ByVal SinDistancia As Boolean)
        On Error GoTo Handler

        Dim puerta As ObjData 'ver ReyarB
        
        

        
        If Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2 And Not SinDistancia Then
            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If


        puerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex)

        If puerta.Llave = 1 And Not SinDistancia Then
            Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If puerta.Cerrada = 1 Then 'Abre la puerta
            MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexAbierta
            Call BloquearPuerta(Map, X, Y, False)

        Else 'Cierra puerta
            MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexCerrada
            Call BloquearPuerta(Map, X, Y, True)

        End If

        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 1 Then
             Call AccionParaPuerta(Map, X - 3, Y + 1, UserIndex, True)
        End If

        Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(MapData(Map, X, Y).ObjInfo.ObjIndex, X, Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex

        Exit Sub

Handler:
132 Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaPuerta", Erl)
134 Resume Next

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

        On Error GoTo Handler

        Dim MiObj As obj

100     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
102         If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
104             Call WriteShowSignal(UserIndex, MapData(Map, X, Y).ObjInfo.ObjIndex)

            End If
  
        End If
        
        Exit Sub
        
Handler:
106 Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaCartel", Erl)
108 Resume Next

End Sub

Sub AccionParaCorreo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo AccionParaCorreo_Err
        

        Dim Pos As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If UserList(UserIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto.", FontTypeNames.FONTTYPE_INFO)
108         Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

110     If Distancia(Pos, UserList(UserIndex).Pos) > 4 Then
112         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

114     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 47 Then
116         Call WriteListaCorreo(UserIndex, False)

        End If

        
        Exit Sub

AccionParaCorreo_Err:
118     Call RegistrarError(Err.Number, Err.Description, "Argentum20Server.Acciones.AccionParaCorreo", Erl)
120     Resume Next
        
End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error GoTo Handler

        Dim Suerte As Byte
        Dim exito  As Byte
        Dim raise  As Integer
    
        Dim Pos    As WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y
    
106     With UserList(UserIndex)
    
108         If Distancia(Pos, .Pos) > 2 Then
110             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

112         If MapInfo(Map).lluvia And Lloviendo Then
114             Call WriteConsoleMsg(UserIndex, "Esta lloviendo, no podés encender una fogata aquí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

116         If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Seguro = 1 Then
118             Call WriteConsoleMsg(UserIndex, "En zona segura no podés hacer fogatas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

120         If MapData(Map, X - 1, Y).ObjInfo.ObjIndex = FOGATA Or _
               MapData(Map, X + 1, Y).ObjInfo.ObjIndex = FOGATA Or _
               MapData(Map, X, Y - 1).ObjInfo.ObjIndex = FOGATA Or _
               MapData(Map, X, Y + 1).ObjInfo.ObjIndex = FOGATA Then
           
122             Call WriteConsoleMsg(UserIndex, "Debes alejarte un poco de la otra fogata.", FontTypeNames.FONTTYPE_INFO)
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
        
146                 Call WriteConsoleMsg(UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
        
148                 Call MakeObj(obj, Map, X, Y)

                Else
        
150                 Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
            
                    Exit Sub

                End If

            Else
        
152             Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)

            End If
    
        End With

154     Call SubirSkill(UserIndex, Supervivencia)

        Exit Sub
        
Handler:
156 Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaRamita", Erl)
158 Resume Next

End Sub
