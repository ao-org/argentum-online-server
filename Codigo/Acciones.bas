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
122                 If NpcList(MapData(Map, X, Y).NpcIndex).Movement = TipoAI.Caminata Then
124                     NpcList(MapData(Map, X, Y).NpcIndex).Contadores.IntervaloMovimiento = GetTickCount + 15000 - NpcList(MapData(Map, X, Y).NpcIndex).IntervaloMovimiento ' 15 segundos
                    End If
            
                    'Iniciamos la rutina pa' comerciar.
126                 Call IniciarComercioNPC(UserIndex)
        
128             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Banquero Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
130                 If UserList(UserIndex).flags.Muerto = 1 Then
132                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
134                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub

                    End If
            
136                 If Distancia(NpcList(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 6 Then
138                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'A depositar de una
140                 Call IniciarBanco(UserIndex)
            
142             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Pirata Then  'VIAJES

                    '¿Esta el user muerto? Si es asi no puede comerciar
144                 If UserList(UserIndex).flags.Muerto = 1 Then
146                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
148                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub

                    End If
            
150                 If Distancia(NpcList(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 5 Then
152                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
154                     Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
156                 If NpcList(MapData(Map, X, Y).NpcIndex).SoundOpen <> 0 Then
158                     Call WritePlayWave(UserIndex, NpcList(MapData(Map, X, Y).NpcIndex).SoundOpen, NO_3D_SOUND, NO_3D_SOUND)

                    End If

                    'A depositar de unaIniciarTransporte
160                 Call WriteViajarForm(UserIndex, MapData(Map, X, Y).NpcIndex)
                    Exit Sub
            
162             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.ResucitadorNewbie Then

164                 If Distancia(UserList(UserIndex).Pos, NpcList(MapData(Map, X, Y).NpcIndex).Pos) > 5 Then
                        'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
166                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
168                 If NpcList(MapData(Map, X, Y).NpcIndex).Movement = Caminata Then
170                     NpcList(MapData(Map, X, Y).NpcIndex).Contadores.IntervaloMovimiento = GetTickCount + 5000 - NpcList(MapData(Map, X, Y).NpcIndex).IntervaloMovimiento ' 5 segundos
                    End If
            
172                 UserList(UserIndex).flags.Envenenado = 0
174                 UserList(UserIndex).flags.Incinerado = 0
      
                    'Revivimos si es necesario
176                 If UserList(UserIndex).flags.Muerto = 1 And (NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
178                     Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)
180                     Call RevivirUsuario(UserIndex)
182                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 30, False))
184                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                
                    Else

                        'curamos totalmente
186                     If UserList(UserIndex).Stats.MinHp <> UserList(UserIndex).Stats.MaxHp Then
188                         UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
190                         Call WritePlayWave(UserIndex, "117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                            'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...¡Has sido curado!", FontTypeNames.FONTTYPE_INFO)
192                         Call WriteLocaleMsg(UserIndex, "83", FontTypeNames.FONTTYPE_INFOIAO)
                    
194                         Call WriteUpdateUserStats(UserIndex)

196                         If Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
198                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.CurarCrimi, 100, False))
                            Else
           
200                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Curar, 100, False))

                            End If

                        End If

                    End If
         
202             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Subastador Then

204                 If UserList(UserIndex).flags.Muerto = 1 Then
206                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
208                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 1 Then
                        ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del subastador.", FontTypeNames.FONTTYPE_INFO)
210                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
212                 If NpcList(MapData(Map, X, Y).NpcIndex).Movement = Caminata Then
214                     NpcList(MapData(Map, X, Y).NpcIndex).Contadores.IntervaloMovimiento = GetTickCount + 20000 - NpcList(MapData(Map, X, Y).NpcIndex).IntervaloMovimiento ' 20 segundos
                    End If

216                 Call IniciarSubasta(UserIndex)
            
218             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Quest Then

220                 If UserList(UserIndex).flags.Muerto = 1 Then
222                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
224                 If NpcList(MapData(Map, X, Y).NpcIndex).Movement = Caminata Then
226                     NpcList(MapData(Map, X, Y).NpcIndex).Contadores.IntervaloMovimiento = GetTickCount + 15000 - NpcList(MapData(Map, X, Y).NpcIndex).IntervaloMovimiento ' 15 segundos
                    End If
            
228                 Call EnviarQuest(UserIndex)
            
230             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Enlistador Then

232                 If UserList(UserIndex).flags.Muerto = 1 Then
234                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
236                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 4 Then
238                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
        
240                 If NpcList(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
242                     If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
244                         Call EnlistarArmadaReal(UserIndex)
                        Else
246                         Call RecompensaArmadaReal(UserIndex)

                        End If
                    Else
248                     If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
250                         Call EnlistarCaos(UserIndex)
                        Else
252                         Call RecompensaCaos(UserIndex)
                        End If
                    End If

254             ElseIf NpcList(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Gobernador Then

256                 If UserList(UserIndex).flags.Muerto = 1 Then
258                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
260                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
262                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del gobernador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Dim DeDonde As String
                    Dim Gobernador As npc
264                     Gobernador = NpcList(UserList(UserIndex).flags.TargetNPC)
            
266                 If UserList(UserIndex).Hogar = Gobernador.GobernadorDe Then
268                     Call WriteChatOverHead(UserIndex, "Ya perteneces a esta ciudad. Gracias por ser uno más de nosotros.", Gobernador.Char.CharIndex, vbWhite)
                        Exit Sub

                    End If
            
270                 If UserList(UserIndex).Faccion.Status = 0 Or UserList(UserIndex).Faccion.Status = 2 Then
272                     If Gobernador.GobernadorDe = eCiudad.cBanderbill Then
274                         Call WriteChatOverHead(UserIndex, "Aquí no aceptamos criminales.", Gobernador.Char.CharIndex, vbWhite)
                            Exit Sub

                        End If

                    End If
            
276                 If UserList(UserIndex).Faccion.Status = 3 Or UserList(UserIndex).Faccion.Status = 1 Then
278                     If Gobernador.GobernadorDe = eCiudad.cArkhein Then
280                         Call WriteChatOverHead(UserIndex, "¡¡Sal de aquí ciudadano asqueroso!!", Gobernador.Char.CharIndex, vbWhite)
                            Exit Sub

                        End If

                    End If
            
282                 If UserList(UserIndex).Hogar <> Gobernador.GobernadorDe Then
            
284                     UserList(UserIndex).PosibleHogar = Gobernador.GobernadorDe
                
286                     Select Case UserList(UserIndex).PosibleHogar

                            Case eCiudad.cUllathorpe
288                             DeDonde = "Ullathorpe"
                            
290                         Case eCiudad.cNix
292                             DeDonde = "Nix"
                
294                         Case eCiudad.cBanderbill
296                             DeDonde = "Banderbill"
                        
298                         Case eCiudad.cLindos
300                             DeDonde = "Lindos"
                            
302                         Case eCiudad.cArghal
304                             DeDonde = " Arghal"
                            
306                         Case eCiudad.cArkhein
308                             DeDonde = " Arkhein"

310                         Case Else
312                             DeDonde = "Ullathorpe"

                        End Select
                    
314                     UserList(UserIndex).flags.pregunta = 3
316                     Call WritePreguntaBox(UserIndex, "¿Te gustaria ser ciudadano de " & DeDonde & "?")
                
                    End If

                End If
        
                '¿Es un obj?
318         ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
320             UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        
322             Select Case ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
324                     Call AccionParaPuerta(Map, X, Y, UserIndex)

326                 Case eOBJType.otCarteles 'Es un cartel
328                     Call AccionParaCartel(Map, X, Y, UserIndex)

330                 Case eOBJType.OtCorreo 'Es un cartel
                        'Call AccionParaCorreo(Map, x, Y, UserIndex)

332                 Case eOBJType.otForos 'Foro
                        'Call AccionParaForo(Map, X, Y, UserIndex)
334                     Call WriteConsoleMsg(UserIndex, "El foro está temporalmente deshabilitado.", FontTypeNames.FONTTYPE_EJECUCION)

336                 Case eOBJType.OtPozos 'Pozos
                        'Call AccionParaPozos(Map, x, Y, UserIndex)

338                 Case eOBJType.otArboles 'Pozos
                        'Call AccionParaArboles(Map, x, Y, UserIndex)

340                 Case eOBJType.otYunque 'Pozos
342                     Call AccionParaYunque(Map, X, Y, UserIndex)

344                 Case eOBJType.otLeña    'Leña

346                     If MapData(Map, X, Y).ObjInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
348                         Call AccionParaRamita(Map, X, Y, UserIndex)

                        End If

                End Select

                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
350         ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
352             UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
        
354             Select Case ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
356                     Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
                End Select

358         ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
360             UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex

362             Select Case ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
364                     Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
                End Select

366         ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
368             UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).ObjInfo.ObjIndex

370             Select Case ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
372                     Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

                End Select

            'ElseIf HayAgua(Map, x, Y) Then
                'Call AccionParaAgua(Map, x, Y, UserIndex)

            End If

        End If

        
        Exit Sub

Accion_Err:
374     Call RegistrarError(Err.Number, Err.Description, "Acciones.Accion", Erl)

        
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

110     If MapData(Map, X, Y).ObjInfo.amount <= 1 Then
112         Call WriteConsoleMsg(UserIndex, "El pozo esta drenado, regresa mas tarde...", FontTypeNames.FONTTYPE_EJECUCION)
            Exit Sub

        End If

114     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 1 Then
116         If UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN Then
118             Call WriteConsoleMsg(UserIndex, "No tenes necesidad del pozo...", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

120         UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
122         MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount - 1
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
140         MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount - 1
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

120     If MapData(Map, X, Y).ObjInfo.amount <= 1 Then
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
    
134     MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount - 1
    
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
        
110     If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
112         Call WriteConsoleMsg(UserIndex, "Debes tener equipado un martillo de herrero para trabajar con el yunque.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
114     If ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo <> 7 Then
            'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
116         Call WriteConsoleMsg(UserIndex, "La herramienta que tienes no es la correcta, necesitas un martillo de herrero para poder trabajar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

118     Call EnivarArmasConstruibles(UserIndex)
120     Call EnivarArmadurasConstruibles(UserIndex)
122     Call WriteShowBlacksmithForm(UserIndex)

        'UserList(UserIndex).Invent.HerramientaEqpObjIndex = objindex
        'UserList(UserIndex).Invent.HerramientaEqpSlot = slot

        
        Exit Sub

AccionParaYunque_Err:
124     Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaYunque", Erl)

        
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal UserIndex As Integer, Optional ByVal SinDistancia As Boolean)
        On Error GoTo Handler

        Dim puerta As ObjData 'ver ReyarB
        
        

        
100     If Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2 And Not SinDistancia Then
            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
102         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If


104     puerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex)
106     If puerta.Llave = 1 And Not SinDistancia Then
108         Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

110     If puerta.Cerrada = 1 Then 'Abre la puerta
112         MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexAbierta
114         Call BloquearPuerta(Map, X, Y, False)

        Else 'Cierra puerta
116         MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexCerrada
118         Call BloquearPuerta(Map, X, Y, True)

        End If

120     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 1 Then
122          Call AccionParaPuerta(Map, X - 3, Y + 1, UserIndex, True)
        End If

124     Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(MapData(Map, X, Y).ObjInfo.ObjIndex, MapData(Map, X, Y).ObjInfo.amount, X, Y))
126     If puerta.GrhIndex = 11445 Or puerta.GrhIndex = 11444 Then
128         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA_DUCTO, X, Y))
130     ElseIf puerta.GrhIndex = 11447 Or puerta.GrhIndex = 11446 Then
        Else
132         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        End If
134     UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex

        Exit Sub

Handler:
136 Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaPuerta", Erl)
138 Resume Next

End Sub


Sub AccionParaPuertaNpc(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal NpcIndex As Integer)
        On Error GoTo Handler

        Dim puerta As ObjData 'ver ReyarB


100     puerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex)

102     If puerta.Cerrada = 1 Then 'Abre la puerta
104         MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexAbierta
106         Call BloquearPuerta(Map, X, Y, False)

        Else 'Cierra puerta
108         MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexCerrada
110         Call BloquearPuerta(Map, X, Y, True)

        End If

112     Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(MapData(Map, X, Y).ObjInfo.ObjIndex, MapData(Map, X, Y).ObjInfo.amount, X, Y))

114     Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))

        Exit Sub

Handler:
116 Call RegistrarError(Err.Number, Err.Description, "Acciones.AccionParaPuertaNpc", Erl)
118 Resume Next

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
144                 obj.amount = 1
        
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
