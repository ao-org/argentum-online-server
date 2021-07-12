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
106             TempCharIndex = MapData(Map, X, Y).NpcIndex

                'Set the target NPC
108             UserList(UserIndex).flags.TargetNPC = TempCharIndex
110             UserList(UserIndex).flags.TargetNpcTipo = NpcList(TempCharIndex).NPCtype
        
112             If NpcList(TempCharIndex).Comercia = 1 Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
114                 If UserList(UserIndex).flags.Muerto = 1 Then
116                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
118                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub

                    End If
            
120                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 6 Then
122                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
124                 If NpcList(TempCharIndex).Movement = TipoAI.Caminata Then
126                     NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 15000 - NpcList(TempCharIndex).IntervaloMovimiento ' 15 segundos
                    End If
            
                    'Iniciamos la rutina pa' comerciar.
128                 Call IniciarComercioNPC(UserIndex)
        
130             ElseIf NpcList(TempCharIndex).NPCtype = eNPCType.Banquero Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
132                 If UserList(UserIndex).flags.Muerto = 1 Then
134                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
136                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub

                    End If
            
138                 If Distancia(NpcList(TempCharIndex).Pos, UserList(UserIndex).Pos) > 6 Then
140                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'A depositar de una
142                 Call IniciarBanco(UserIndex)
            
144             ElseIf NpcList(TempCharIndex).NPCtype = eNPCType.Pirata Then  'VIAJES

                    '¿Esta el user muerto? Si es asi no puede comerciar
146                 If UserList(UserIndex).flags.Muerto = 1 Then
148                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
150                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub

                    End If
            
152                 If Distancia(NpcList(TempCharIndex).Pos, UserList(UserIndex).Pos) > 5 Then
154                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
156                     Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del banquero.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
158                 If NpcList(TempCharIndex).SoundOpen <> 0 Then
160                     Call WritePlayWave(UserIndex, NpcList(TempCharIndex).SoundOpen, NO_3D_SOUND, NO_3D_SOUND)

                    End If

                    'A depositar de unaIniciarTransporte
162                 Call WriteViajarForm(UserIndex, TempCharIndex)
                    Exit Sub
            
164             ElseIf NpcList(TempCharIndex).NPCtype = eNPCType.Revividor Or NpcList(TempCharIndex).NPCtype = eNPCType.ResucitadorNewbie Then

166                 If Distancia(UserList(UserIndex).Pos, NpcList(TempCharIndex).Pos) > 5 Then
                        'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
168                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
170                 If NpcList(TempCharIndex).Movement = Caminata Then
172                     NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 5000 - NpcList(TempCharIndex).IntervaloMovimiento ' 5 segundos
                    End If
            
174                 UserList(UserIndex).flags.Envenenado = 0
176                 UserList(UserIndex).flags.Incinerado = 0
      
                    'Revivimos si es necesario
178                 If UserList(UserIndex).flags.Muerto = 1 And (NpcList(TempCharIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
180                     Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)
182                     Call RevivirUsuario(UserIndex)
184                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 30, False))
186                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                
                    Else

                        'curamos totalmente
188                     If UserList(UserIndex).Stats.MinHp <> UserList(UserIndex).Stats.MaxHp Then
190                         UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
192                         Call WritePlayWave(UserIndex, "117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                            'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...¡Has sido curado!", FontTypeNames.FONTTYPE_INFO)
194                         Call WriteLocaleMsg(UserIndex, "83", FontTypeNames.FONTTYPE_INFOIAO)
                    
196                         Call WriteUpdateUserStats(UserIndex)

198                         If Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
200                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.CurarCrimi, 100, False))
                            Else
           
202                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Curar, 100, False))

                            End If

                        End If

                    End If
         
204             ElseIf NpcList(TempCharIndex).NPCtype = eNPCType.Subastador Then

206                 If UserList(UserIndex).flags.Muerto = 1 Then
208                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
210                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 1 Then
                        ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del subastador.", FontTypeNames.FONTTYPE_INFO)
212                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
214                 If NpcList(TempCharIndex).Movement = Caminata Then
216                     NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 20000 - NpcList(TempCharIndex).IntervaloMovimiento ' 20 segundos
                    End If

218                 Call IniciarSubasta(UserIndex)
            
220             ElseIf NpcList(TempCharIndex).NPCtype = eNPCType.Quest Then

222                 If UserList(UserIndex).flags.Muerto = 1 Then
224                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
226                 If NpcList(TempCharIndex).Movement = Caminata Then
228                     NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 15000 - NpcList(TempCharIndex).IntervaloMovimiento ' 15 segundos
                    End If
            
230                 Call EnviarQuest(UserIndex)
            
232             ElseIf NpcList(TempCharIndex).NPCtype = eNPCType.Enlistador Then

234                 If UserList(UserIndex).flags.Muerto = 1 Then
236                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
238                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 4 Then
240                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
        
242                 If NpcList(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
244                     If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
246                         Call EnlistarArmadaReal(UserIndex)
                        Else
248                         Call RecompensaArmadaReal(UserIndex)

                        End If
                    Else
250                     If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
252                         Call EnlistarCaos(UserIndex)
                        Else
254                         Call RecompensaCaos(UserIndex)
                        End If
                    End If

256             ElseIf NpcList(TempCharIndex).NPCtype = eNPCType.Gobernador Then

258                 If UserList(UserIndex).flags.Muerto = 1 Then
260                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
262                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
264                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del gobernador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Dim DeDonde As String
                    Dim Gobernador As npc
266                     Gobernador = NpcList(UserList(UserIndex).flags.TargetNPC)
            
268                 If UserList(UserIndex).Hogar = Gobernador.GobernadorDe Then
270                     Call WriteChatOverHead(UserIndex, "Ya perteneces a esta ciudad. Gracias por ser uno más de nosotros.", Gobernador.Char.CharIndex, vbWhite)
                        Exit Sub

                    End If
            
272                 If UserList(UserIndex).Faccion.Status = 0 Or UserList(UserIndex).Faccion.Status = 2 Then
274                     If Gobernador.GobernadorDe = eCiudad.cBanderbill Then
276                         Call WriteChatOverHead(UserIndex, "Aquí no aceptamos criminales.", Gobernador.Char.CharIndex, vbWhite)
                            Exit Sub

                        End If

                    End If
            
278                 If UserList(UserIndex).Faccion.Status = 3 Or UserList(UserIndex).Faccion.Status = 1 Then
280                     If Gobernador.GobernadorDe = eCiudad.cArkhein Then
282                         Call WriteChatOverHead(UserIndex, "¡¡Sal de aquí ciudadano asqueroso!!", Gobernador.Char.CharIndex, vbWhite)
                            Exit Sub

                        End If

                    End If
            
284                 If UserList(UserIndex).Hogar <> Gobernador.GobernadorDe Then
            
286                     UserList(UserIndex).PosibleHogar = Gobernador.GobernadorDe
                
288                     Select Case UserList(UserIndex).PosibleHogar

                            Case eCiudad.cUllathorpe
290                             DeDonde = "Ullathorpe"
                            
292                         Case eCiudad.cNix
294                             DeDonde = "Nix"
                
296                         Case eCiudad.cBanderbill
298                             DeDonde = "Banderbill"
                        
300                         Case eCiudad.cLindos
302                             DeDonde = "Lindos"
                            
304                         Case eCiudad.cArghal
306                             DeDonde = " Arghal"
                            
308                         Case eCiudad.cArkhein
310                             DeDonde = " Arkhein"

312                         Case Else
314                             DeDonde = "Ullathorpe"

                        End Select
                    
316                     UserList(UserIndex).flags.pregunta = 3
318                     Call WritePreguntaBox(UserIndex, "¿Te gustaria ser ciudadano de " & DeDonde & "?")
                
                    End If
                    
320             ElseIf NpcList(TempCharIndex).Craftea > 0 Then
322                 If UserList(UserIndex).flags.Muerto = 1 Then
324                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
            
326                 If Distancia(NpcList(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
328                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
330                 UserList(UserIndex).flags.Crafteando = NpcList(TempCharIndex).Craftea
332                 Call WriteOpenCrafting(UserIndex, NpcList(TempCharIndex).Craftea)
                End If
        
                '¿Es un obj?
334         ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
336             UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        
338             Select Case ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
340                     Call AccionParaPuerta(Map, X, Y, UserIndex)

342                 Case eOBJType.otCarteles 'Es un cartel
344                     Call AccionParaCartel(Map, X, Y, UserIndex)

346                 Case eOBJType.OtCorreo 'Es un cartel
                        'Call AccionParaCorreo(Map, x, Y, UserIndex)

348                 Case eOBJType.otForos 'Foro
                        'Call AccionParaForo(Map, X, Y, UserIndex)
350                     Call WriteConsoleMsg(UserIndex, "El foro está temporalmente deshabilitado.", FontTypeNames.FONTTYPE_EJECUCION)

352                 Case eOBJType.OtPozos 'Pozos
                        'Call AccionParaPozos(Map, x, Y, UserIndex)

354                 Case eOBJType.otArboles 'Pozos
                        'Call AccionParaArboles(Map, x, Y, UserIndex)

356                 Case eOBJType.otYunque 'Pozos
358                     Call AccionParaYunque(Map, X, Y, UserIndex)

360                 Case eOBJType.otLeña    'Leña

362                     If MapData(Map, X, Y).ObjInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
364                         Call AccionParaRamita(Map, X, Y, UserIndex)

                        End If

                End Select

                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
366         ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
368             UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
        
370             Select Case ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
372                     Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
                End Select

374         ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
376             UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex

378             Select Case ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
380                     Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
                End Select

382         ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
384             UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).ObjInfo.ObjIndex

386             Select Case ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case eOBJType.otPuertas 'Es una puerta
388                     Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

                End Select

            'ElseIf HayAgua(Map, x, Y) Then
                'Call AccionParaAgua(Map, x, Y, UserIndex)

            End If

        End If

        
        Exit Sub

Accion_Err:
390     Call TraceError(Err.Number, Err.Description, "Acciones.Accion", Erl)

        
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
144     Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaForo", Erl)

        
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
148     Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaPozos", Erl)

        
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
146     Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaArboles", Erl)

        
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
138     Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaAgua", Erl)

        
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
124     Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaYunque", Erl)

        
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
126     If puerta.GrhIndex = 11445 Or puerta.GrhIndex = 11444 Or puerta.GrhIndex = 59878 Or puerta.GrhIndex = 59877 Then
128         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA_DUCTO, X, Y))
        Else
132         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        End If
134     UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex

        Exit Sub

Handler:
136 Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaPuerta", Erl)


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
116 Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaPuertaNpc", Erl)


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
106 Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaCartel", Erl)


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
118     Call TraceError(Err.Number, Err.Description, "Argentum20Server.Acciones.AccionParaCorreo", Erl)

        
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
156 Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaRamita", Erl)


End Sub
