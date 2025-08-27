Attribute VB_Name = "Acciones"
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


Public Function get_map_name(ByVal map As Long) As String
On Error GoTo get_map_name_Err
        get_map_name = MapInfo(map).map_name
        Exit Function
get_map_name_Err:
     Call TraceError(Err.Number, Err.Description, "ModLadder.get_map_name", Erl)
End Function


Function PuedeUsarObjeto(UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByVal writeInConsole As Boolean = False) As Byte
        On Error GoTo PuedeUsarObjeto_Err

        Dim Objeto As t_ObjData
        Dim Msg As String, i As Long
     Objeto = ObjData(ObjIndex)
                
     With UserList(UserIndex)
     
         If EsGM(UserIndex) Then
             PuedeUsarObjeto = 0
             Msg = ""
    
         ElseIf Objeto.Newbie = 1 And Not EsNewbie(UserIndex) Then
             PuedeUsarObjeto = 7
             Msg = "Solo los newbies pueden usar este objeto."
                
         ElseIf .Stats.ELV < Objeto.MinELV Then
             PuedeUsarObjeto = 6
             Msg = "Necesitas ser nivel " & Objeto.MinELV & " para usar este objeto."
             
         ElseIf .Stats.ELV > Objeto.MaxLEV And Objeto.MaxLEV > 0 Then
             PuedeUsarObjeto = 6
             Msg = "Este objeto no puede ser utilizado por personajes de nivel " & Objeto.MaxLEV & " o superior."
     
         ElseIf Not FaccionPuedeUsarItem(UserIndex, ObjIndex) And JerarquiaPuedeUsarItem(UserIndex, ObjIndex) Then
             PuedeUsarObjeto = 3
             Msg = "Tu facción no te permite utilizarlo."
    
         ElseIf Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
             PuedeUsarObjeto = 2
             Msg = "Tu clase no puede utilizar este objeto."
    
         ElseIf Not SexoPuedeUsarItem(UserIndex, ObjIndex) Then
             PuedeUsarObjeto = 1
             Msg = "Tu sexo no puede utilizar este objeto."
    
         ElseIf Not RazaPuedeUsarItem(UserIndex, ObjIndex) Then
             PuedeUsarObjeto = 5
             Msg = "Tu raza no puede utilizar este objeto."
         ElseIf (Objeto.SkillIndex > 0) Then
             If (.Stats.UserSkills(Objeto.SkillIndex) < Objeto.SkillRequerido) Then
                 PuedeUsarObjeto = 4
                 Msg = "Necesitas " & Objeto.SkillRequerido & " puntos en " & SkillsNames(Objeto.SkillIndex) & " para usar este item."
                Else
                 PuedeUsarObjeto = 0
                 Msg = ""
                End If
            Else
             PuedeUsarObjeto = 0
             Msg = ""
            End If
    End With
     If writeInConsole And Msg <> "" Then Call WriteConsoleMsg(UserIndex, Msg, e_FontTypeNames.FONTTYPE_INFO)

        Exit Function

PuedeUsarObjeto_Err:
     Call TraceError(Err.Number, Err.Description, "ModLadder.PuedeUsarObjeto", Erl)

End Function


Public Sub CompletarAccionFin(ByVal UserIndex As Integer)
        
        On Error GoTo CompletarAccionFin_Err
        

        Dim obj  As t_ObjData

        Dim Slot As Byte

     Select Case UserList(UserIndex).Accion.TipoAccion

            Case e_AccionBarra.Runa
             obj = ObjData(UserList(UserIndex).Accion.RunaObj)
             Slot = UserList(UserIndex).Accion.ObjSlot

             Select Case obj.TipoRuna

                    Case e_RuneType.ReturnHome 'lleva a la ciudad de origen vivo o muerto

                        Dim DeDonde As t_CityWorldPos

                        Dim map     As Integer

                        Dim X       As Byte

                        Dim y       As Byte
        
                     If UserList(UserIndex).flags.Muerto = 0 Then

                         Select Case UserList(UserIndex).Hogar

                                Case e_Ciudad.cUllathorpe
                                 DeDonde = CityUllathorpe
                        
                             Case e_Ciudad.cNix
                                 DeDonde = CityNix
            
                             Case e_Ciudad.cBanderbill
                                 DeDonde = CityBanderbill
                    
                             Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                                 DeDonde = CityLindos
                        
                             Case e_Ciudad.cArghal
                                 DeDonde = CityArghal
                                 
                             Case e_Ciudad.cForgat
                                 DeDonde = CityForgat

                             Case e_Ciudad.cEldoria
                                 DeDonde = CityEldoria
                        
                             Case e_Ciudad.cArkhein
                                 DeDonde = CityArkhein
                        
                             Case Else
                                 DeDonde = CityUllathorpe

                            End Select

                         map = DeDonde.map
                         X = DeDonde.X
                         y = DeDonde.y
                        Else

                         If MapInfo(UserList(UserIndex).Pos.map).ResuCiudad <> 0 Then

                             Select Case MapInfo(UserList(UserIndex).Pos.map).ResuCiudad

                                    Case e_Ciudad.cUllathorpe
                                     DeDonde = CityUllathorpe
                        
                                 Case e_Ciudad.cNix
                                     DeDonde = CityNix
            
                                 Case e_Ciudad.cBanderbill
                                     DeDonde = CityBanderbill
                    
                                 Case e_Ciudad.cLindos
                                     DeDonde = CityLindos
                        
                                 Case e_Ciudad.cArghal
                                     DeDonde = CityArghal
                                     
                                 Case e_Ciudad.cForgat
                                     DeDonde = CityForgat
                        
                                 Case e_Ciudad.cArkhein
                                     DeDonde = CityArkhein
                                 
                                 Case e_Ciudad.cEldoria
                                     DeDonde = CityEldoria

                        
                                 Case Else
                                     DeDonde = CityUllathorpe

                                End Select

                            Else

                             Select Case UserList(UserIndex).Hogar

                                    Case e_Ciudad.cUllathorpe
                                     DeDonde = CityUllathorpe
                        
                                 Case e_Ciudad.cNix
                                     DeDonde = CityNix
            
                                 Case e_Ciudad.cBanderbill
                                     DeDonde = CityBanderbill
                    
                                 Case e_Ciudad.cLindos
                                     DeDonde = CityLindos
                        
                                 Case e_Ciudad.cArghal
                                     DeDonde = CityArghal
                                     
                                 Case e_Ciudad.cForgat
                                     DeDonde = CityForgat
                        
                                 Case e_Ciudad.cArkhein
                                     DeDonde = CityArkhein

                                 Case e_Ciudad.cEldoria
                                     DeDonde = CityEldoria
                        
                                 Case Else
                                     DeDonde = CityUllathorpe

                                End Select

                            End If
                
                         map = DeDonde.MapaResu
                         X = DeDonde.ResuX
                         y = DeDonde.ResuY
                
                            Dim Resu As Boolean
                
                         Resu = True
            
                        End If
                
                     Call FindLegalPos(UserIndex, map, X, y)
                     Call WarpUserChar(UserIndex, map, X, y, True)
                        'Msg1065= Has regresado a tu ciudad de origen.
                        Call WriteLocaleMsg(UserIndex, "1065", e_FontTypeNames.FONTTYPE_WARNING)

                        'Call WriteFlashScreen(UserIndex, &HA4FFFF, 150, True)
                     If UserList(UserIndex).flags.Navegando = 1 Then

                            Dim barca As t_ObjData

                         barca = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
                         Call DoNavega(UserIndex, barca, UserList(UserIndex).Invent.BarcoSlot)

                        End If
                
                     If Resu Then
                
                         UserList(UserIndex).Counters.TimerBarra = 5
                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, e_ParticulasIndex.Resucitar, UserList(UserIndex).Counters.TimerBarra, False))
                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Counters.TimerBarra, e_AccionBarra.Resucitar))

                
                         UserList(UserIndex).Accion.AccionPendiente = True
                         UserList(UserIndex).Accion.Particula = e_ParticulasIndex.Resucitar
                         UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.Resucitar

                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))
                            'Msg82=El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...
                         Call WriteLocaleMsg(UserIndex, "82", e_FontTypeNames.FONTTYPE_INFOIAO)

                        End If
                
                     If Not Resu Then
                         UserList(UserIndex).Accion.AccionPendiente = False
                         UserList(UserIndex).Accion.Particula = 0
                         UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion

                        End If

                     UserList(UserIndex).Accion.HechizoPendiente = 0
                     UserList(UserIndex).Accion.RunaObj = 0
                     UserList(UserIndex).Accion.ObjSlot = 0
              
                 Case e_RuneType.Escape
                     map = obj.HastaMap
                     X = obj.HastaX
                     y = obj.HastaY
            
                     If obj.DesdeMap = 0 Then
                         Call FindLegalPos(UserIndex, map, X, y)
                         Call WarpUserChar(UserIndex, map, X, y, True)
                            'Msg1066= Te has teletransportado por el mundo.
                            Call WriteLocaleMsg(UserIndex, "1066", e_FontTypeNames.FONTTYPE_WARNING)
                         Call QuitarUserInvItem(UserIndex, Slot, 1)
                         Call UpdateUserInv(False, UserIndex, Slot)
                        Else

                         If UserList(UserIndex).Pos.map <> obj.DesdeMap Then
                            'Msg1067= Esta runa no puede ser usada desde aquí.
                            Call WriteLocaleMsg(UserIndex, "1067", e_FontTypeNames.FONTTYPE_INFO)
                            Else
                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                             Call UpdateUserInv(False, UserIndex, Slot)
                             Call FindLegalPos(UserIndex, map, X, y)
                             Call WarpUserChar(UserIndex, map, X, y, True)
                            'Msg1068= Te has teletransportado por el mundo.
                            Call WriteLocaleMsg(UserIndex, "1068", e_FontTypeNames.FONTTYPE_WARNING)

                            End If

                        End If
        
                     UserList(UserIndex).Accion.Particula = 0
                     UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
                     UserList(UserIndex).Accion.HechizoPendiente = 0
                     UserList(UserIndex).Accion.RunaObj = 0
                     UserList(UserIndex).Accion.ObjSlot = 0
                     UserList(UserIndex).Accion.AccionPendiente = False


                    Case e_RuneType.MesonSafePassage

                        If UserList(UserIndex).Pos.Map = MAP_MESON_HOSTIGADO or UserList(UserIndex).Pos.Map = MAP_MESON_HOSTIGADO_TRADING_ZONE Then
                            'mensaje de error de "no puedes usar la runa estando en el meson"
                            Call WriteLocaleMsg(UserIndex, "2081", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If

                        If obj.HastaMap <> MAP_MESON_HOSTIGADO Then
                            'mensaje de error de runa invalida, hay algo mal dateado llamar a un gm o avisar a soporte
                            Call WriteLocaleMsg(UserIndex, "2080", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        UserList(UserIndex).flags.ReturnPos = UserList(UserIndex).Pos
                        
                        Map = obj.HastaMap
                        x = obj.HastaX
                        y = obj.HastaY
                        
                        Call WarpUserChar(UserIndex, Map, x, y, True)
                        'Msg1066= Te has teletransportado por el mundo.
                        Call WriteLocaleMsg(UserIndex, "1066", e_FontTypeNames.FONTTYPE_WARNING)
                        
                        UserList(UserIndex).Accion.Particula = 0
                        UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
                        UserList(UserIndex).Accion.HechizoPendiente = 0
                        UserList(UserIndex).Accion.RunaObj = 0
                        UserList(UserIndex).Accion.ObjSlot = 0
                        UserList(UserIndex).Accion.AccionPendiente = False
                        
                End Select
                
         Case e_AccionBarra.Hogar
             Call HomeArrival(UserIndex)
             UserList(UserIndex).Accion.AccionPendiente = False
             UserList(UserIndex).Accion.Particula = 0
             UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
            

         Case e_AccionBarra.Intermundia
        
             If UserList(UserIndex).flags.Muerto = 0 Then

                    Dim uh As Integer

                    Dim Mapaf, Xf, Yf As Integer

                 uh = UserList(UserIndex).Accion.HechizoPendiente
    
                 Mapaf = Hechizos(uh).TeleportXMap
                 Xf = Hechizos(uh).TeleportXX
                 Yf = Hechizos(uh).TeleportXY
    
                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(uh).wav, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))  'Esta linea faltaba. Pablo (ToxicWaste)
                 'Msg1069= ¡Has abierto la puerta a intermundia!
                 Call WriteLocaleMsg(UserIndex, "1069", e_FontTypeNames.FONTTYPE_INFO)
                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, e_ParticulasIndex.Runa, -1, True))
                 UserList(UserIndex).flags.Portal = 10
                 UserList(UserIndex).flags.PortalMDestino = Mapaf
                 UserList(UserIndex).flags.PortalYDestino = Xf
                 UserList(UserIndex).flags.PortalXDestino = Yf
                
                    Dim Mapa As Integer

                 Mapa = UserList(UserIndex).flags.PortalM
                 X = UserList(UserIndex).flags.PortalX
                 y = UserList(UserIndex).flags.PortalY
                 MapData(Mapa, X, y).Particula = e_ParticulasIndex.TpVerde
                 MapData(Mapa, X, y).TimeParticula = -1
                 MapData(Mapa, X, y).TileExit.map = UserList(UserIndex).flags.PortalMDestino
                 MapData(Mapa, X, y).TileExit.X = UserList(UserIndex).flags.PortalXDestino
                 MapData(Mapa, X, y).TileExit.y = UserList(UserIndex).flags.PortalYDestino
                
                 Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(X, y, e_ParticulasIndex.TpVerde, -1))
                
                 Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageLightFXToFloor(X, y, &HFF80C0, 105))

                End If
                    
             UserList(UserIndex).Accion.Particula = 0
             UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
             UserList(UserIndex).Accion.HechizoPendiente = 0
             UserList(UserIndex).Accion.RunaObj = 0
             UserList(UserIndex).Accion.ObjSlot = 0
             UserList(UserIndex).Accion.AccionPendiente = False
            
                '
         Case e_AccionBarra.Resucitar
             ' Msg585=¡Has sido resucitado!
             Call WriteLocaleMsg(UserIndex, "585", e_FontTypeNames.FONTTYPE_INFO)
             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, e_ParticulasIndex.Resucitar, 250, True))
             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))
             Call RevivirUsuario(UserIndex, True)
                
             UserList(UserIndex).Accion.Particula = 0
             UserList(UserIndex).Accion.TipoAccion = e_AccionBarra.CancelarAccion
             UserList(UserIndex).Accion.HechizoPendiente = 0
             UserList(UserIndex).Accion.RunaObj = 0
             UserList(UserIndex).Accion.ObjSlot = 0
             UserList(UserIndex).Accion.AccionPendiente = False
                      
        End Select
               
        Exit Sub

CompletarAccionFin_Err:
     Call TraceError(Err.Number, Err.Description, "ModLadder.CompletarAccionFin", Erl)

        
End Sub

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
        
        If UserIndex <= 0 Then Exit Sub

        '¿Posicion valida?
102     If InMapBounds(Map, X, Y) Then
   
            Dim FoundChar      As Byte

            Dim FoundSomething As Byte

            Dim TempCharIndex  As Integer
       
104         If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
106             TempCharIndex = MapData(Map, X, Y).NpcIndex

                'Set the target NPC
108             Call SetNpcRef(UserList(UserIndex).flags.TargetNPC, TempCharIndex)
110             UserList(UserIndex).flags.TargetNpcTipo = NpcList(TempCharIndex).NPCtype
        
112             If NpcList(TempCharIndex).Comercia = 1 Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
114                 If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
116                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    'Is it already in commerce mode??
118                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub

                    End If
            
120                 If Distancia(NpcList(TempCharIndex).Pos, UserList(UserIndex).Pos) > 4 Then
122                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
124                 If NpcList(TempCharIndex).Movement = e_TipoAI.Caminata Then
126                     NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 15000 - NpcList(TempCharIndex).IntervaloMovimiento
                    End If
            
                    'Iniciamos la rutina pa' comerciar.
128                 Call IniciarComercioNPC(UserIndex)
        
130             ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.Banquero Then

                    '¿Esta el user muerto? Si es asi no puede comerciar
132                 If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
134                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'Is it already in commerce mode??
136                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub
                    End If
138                 If Distancia(NpcList(TempCharIndex).Pos, UserList(userindex).Pos) > 4 Then
140                     Call WriteLocaleMsg(userindex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'A depositar de una
142                 Call IniciarBanco(UserIndex)
            
144             ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.Pirata Then  'VIAJES

                    '¿Esta el user muerto? Si es asi no puede comerciar
146                 If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
148                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'Is it already in commerce mode??
150                 If UserList(UserIndex).flags.Comerciando Then
                        Exit Sub
                    End If
            
152                 If Distancia(NpcList(TempCharIndex).Pos, UserList(userindex).Pos) > 4 Then
154                     Call WriteLocaleMsg(userindex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg1070= Estas demasiado lejos del vendedor de pasajes.
                        Call WriteLocaleMsg(UserIndex, "1070", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
158                 If NpcList(TempCharIndex).SoundOpen <> 0 Then
160                     Call WritePlayWave(UserIndex, NpcList(TempCharIndex).SoundOpen, NO_3D_SOUND, NO_3D_SOUND, , 1)
                    End If

                    'A depositar de unaIniciarTransporte
162                 Call WriteViajarForm(UserIndex, TempCharIndex)
                    Exit Sub
            
164             ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.Revividor Or NpcList(TempCharIndex).NPCtype = e_NPCType.ResucitadorNewbie Then

166                 If Distancia(UserList(UserIndex).Pos, NpcList(TempCharIndex).Pos) > 5 Then
                        'Msg8=El sacerdote no puede curarte debido a que estas demasiado lejos.
168                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    '  Hacemos que se detenga a hablar un momento :P
170                 If NpcList(TempCharIndex).Movement = Caminata Then
172                     NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 5000 - NpcList(TempCharIndex).IntervaloMovimiento ' 5 segundos
                    End If
            
174                 UserList(UserIndex).flags.Envenenado = 0
176                 UserList(UserIndex).flags.Incinerado = 0
      
                    'Revivimos si es necesario
178                 If UserList(UserIndex).flags.Muerto = 1 And (NpcList(TempCharIndex).NPCtype = e_NPCType.Revividor Or EsNewbie(UserIndex)) Then
180                     ' Msg585=¡Has sido resucitado!
                        Call WriteLocaleMsg(UserIndex, "585", e_FontTypeNames.FONTTYPE_INFO)
182                     Call RevivirUsuario(UserIndex)
184                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, e_ParticulasIndex.Resucitar, 30, False))
186                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                
                    Else

                        'curamos totalmente
188                     If UserList(UserIndex).Stats.MinHp <> UserList(UserIndex).Stats.MaxHp Then
190                         UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
192                         Call WritePlayWave(UserIndex, "117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                            'Msg83=El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...¡Has sido curado!"
194                         Call WriteLocaleMsg(UserIndex, "83", e_FontTypeNames.FONTTYPE_INFOIAO)
                    
196                         Call WriteUpdateUserStats(UserIndex)

198                         If Status(UserIndex) = 4 Or Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
200                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, e_ParticulasIndex.CurarCrimi, 100, False))
                            Else
           
202                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, e_ParticulasIndex.Curar, 100, False))

                            End If

                        End If

                    End If
         
204             ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.Subastador Then

206                 If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
208                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
210                 If Distancia(NpcList(TempCharIndex).Pos, UserList(UserIndex).Pos) > 1 Then
212                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
214                 If NpcList(TempCharIndex).Movement = Caminata Then
216                     NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 20000 - NpcList(TempCharIndex).IntervaloMovimiento
                    End If

218                 Call IniciarSubasta(UserIndex)
            
220             ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.Quest Then

222                 If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
224                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call WritePlayWave(UserIndex, NpcList(TempCharIndex).SoundOpen, NpcList(TempCharIndex).pos.x, NpcList(TempCharIndex).pos.y, 2, 1)
230                 Call EnviarQuest(UserIndex)
            
232             ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.Enlistador Then

234                 If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
236                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
238                 If Distancia(NpcList(TempCharIndex).Pos, UserList(UserIndex).Pos) > 3 Then
240                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
        
242                 If NpcList(TempCharIndex).flags.Faccion = 0 Then
244                     If UserList(UserIndex).Faccion.Status <> e_Facciones.Armada And UserList(UserIndex).Faccion.Status <> e_Facciones.consejo Then
246                         Call EnlistarArmadaReal(UserIndex)
                        Else
248                         Call RecompensaArmadaReal(UserIndex)

                        End If
                    Else
250                     If UserList(UserIndex).Faccion.Status <> e_Facciones.Caos And UserList(UserIndex).Faccion.Status <> e_Facciones.concilio Then
252                         Call EnlistarCaos(UserIndex)
                        Else
254                         Call RecompensaCaos(UserIndex)
                        End If
                    End If

256             ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.Gobernador Then

258                 If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
260                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
            
262                 If Distancia(NpcList(TempCharIndex).Pos, UserList(UserIndex).Pos) > 3 Then
264                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg8=Estas demasiado lejos del gobernador.
                        Exit Sub

                    End If

                    Dim DeDonde As String
                    Dim Gobernador As t_Npc
266                     Gobernador = NpcList(TempCharIndex)
            
268                 If UserList(UserIndex).Hogar = Gobernador.GobernadorDe Then
270                     Call WriteLocaleChatOverHead(UserIndex, "1349", "", Gobernador.Char.charindex, vbWhite) ' Msg1349=Ya perteneces a esta ciudad. Gracias por ser uno más de nosotros.
                        Exit Sub

                    End If
            
272                 If UserList(UserIndex).Faccion.Status = 0 Or UserList(UserIndex).Faccion.Status = 2 Then
274                     If Gobernador.GobernadorDe = e_Ciudad.cBanderbill Then
276                         Call WriteLocaleChatOverHead(UserIndex, "1350", "", Gobernador.Char.charindex, vbWhite) ' Msg1350=Aquí no aceptamos criminales.
                            Exit Sub

                        End If

                    End If
            
278                 If UserList(UserIndex).Faccion.Status = 3 Or UserList(UserIndex).Faccion.Status = 1 Then
280                     If Gobernador.GobernadorDe = e_Ciudad.cArkhein Then
282                         Call WriteLocaleChatOverHead(UserIndex, "1351", "", Gobernador.Char.charindex, vbWhite) ' Msg1351=¡¡Sal de aquí ciudadano asqueroso!!

                            Exit Sub

                        End If

                    End If
            
284                 If UserList(UserIndex).Hogar <> Gobernador.GobernadorDe Then
            
286                     UserList(UserIndex).PosibleHogar = Gobernador.GobernadorDe
                
288                     Select Case UserList(UserIndex).PosibleHogar

                            Case e_Ciudad.cUllathorpe
290                             DeDonde = "Ullathorpe"
                            
292                         Case e_Ciudad.cNix
294                             DeDonde = "Nix"
                
296                         Case e_Ciudad.cBanderbill
298                             DeDonde = "Banderbill"
                        
300                         Case e_Ciudad.cLindos
302                             DeDonde = "Lindos"
                            
304                         Case e_Ciudad.cArghal
306                             DeDonde = "Arghal"
                        
                            Case e_Ciudad.cForgat
                                DeDonde = "Forgat"

                            Case e_Ciudad.cEldoria
                                DeDonde = "Eldoria"
                            
308                         Case e_Ciudad.cArkhein
310                             DeDonde = "Arkhein"

312                         Case Else
314                             DeDonde = "Ullathorpe"

                        End Select
                    
316                     UserList(UserIndex).flags.pregunta = 3
318                     Call WritePreguntaBox(UserIndex, 1592, DeDonde)
                
                    End If
                ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.EntregaPesca Then
                    Dim i As Integer, j As Integer
                    Dim PuntosTotales As Long
                                                            
                    Dim CantPecesEspeciales As Long
                    Dim OroTotal As Long
                    CantPecesEspeciales = UBound(PecesEspeciales)
                                                
                    If CantPecesEspeciales > 0 Then
                        For i = 1 To MAX_INVENTORY_SLOTS
                            For j = 1 To CantPecesEspeciales
                                If UserList(UserIndex).Invent.Object(i).ObjIndex = PecesEspeciales(j).ObjIndex Then
                                    PuntosTotales = PuntosTotales + (ObjData(UserList(UserIndex).Invent.Object(i).ObjIndex).PuntosPesca * UserList(UserIndex).Invent.Object(i).amount)
                                    OroTotal = OroTotal + (ObjData(UserList(userindex).Invent.Object(i).ObjIndex).Valor * UserList(userindex).Invent.Object(i).amount)
                                End If
                            Next j
                        Next i
                    End If
                    
                    If PuntosTotales > 0 Then
319                     UserList(UserIndex).flags.pregunta = 5
                        Call WritePreguntaBox(UserIndex, 1593, PuntosTotales & "¬" & PonerPuntos(OroTotal * 1.2)) 'Msg1593= Tienes un total de ¬1 puntos y ¬2 monedas de oro para reclamar, ¿Deseas aceptar?
                    Else
                        Dim charIndexstr As Integer
                        charIndexStr = str(NpcList(TempCharIndex).Char.charindex)
                        Call WriteLocaleChatOverHead(UserIndex, "1352", "", charindexstr, &HFFFF00) ' Msg1352=No tienes ningún trofeo de pesca para entregar.
                    End If
                ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.AO20Shop Then
322                 If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
324                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
                    
                    Call WriteShopInit(UserIndex)
                
                ElseIf NpcList(TempCharIndex).NPCtype = e_NPCType.AO20ShopPjs Then
323                 If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
325                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
               
                    Call WriteShopPjsInit(UserIndex)
                ElseIf NpcList(TempCharIndex).npcType = e_NPCType.EventMaster Then
                    If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
                        Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    If Distancia(NpcList(TempCharIndex).pos, UserList(UserIndex).pos) > 4 Then
                        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call WriteUpdateLobbyList(UserIndex)
320             ElseIf NpcList(TempCharIndex).Craftea > 0 Then
                    If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
                        Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
            
326                 If Distancia(NpcList(TempCharIndex).Pos, UserList(UserIndex).Pos) > 3 Then
328                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
330                 UserList(UserIndex).flags.Crafteando = NpcList(TempCharIndex).Craftea
332                 Call WriteOpenCrafting(UserIndex, NpcList(TempCharIndex).Craftea)
                End If
        
                '¿Es un obj?
334         ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
336             UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
        
338             Select Case ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType
            
                    Case e_OBJType.otPuertas 'Es una puerta
340                     Call AccionParaPuerta(Map, X, Y, UserIndex)

342                 Case e_OBJType.otCarteles 'Es un cartel
344                     Call AccionParaCartel(Map, X, Y, UserIndex)

346                 Case e_OBJType.OtCorreo 'Es un cartel
                        'Call AccionParaCorreo(Map, x, Y, UserIndex)
                        ' Msg586=El correo está temporalmente deshabilitado.
                        Call WriteLocaleMsg(UserIndex, "586", e_FontTypeNames.FONTTYPE_EJECUCION)

356                 Case e_OBJType.otYunque 'Pozos
358                     Call AccionParaYunque(Map, X, Y, UserIndex)

360                 Case e_OBJType.otLeña    'Leña

362                     If MapData(Map, X, Y).ObjInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
364                         Call AccionParaRamita(Map, X, Y, UserIndex)

                        End If
                    Case Else
                        Exit Sub

                End Select

                '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
366         ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
368             UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
        
370             Select Case ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType
            
                    Case e_OBJType.otPuertas 'Es una puerta
372                     Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
                End Select

374         ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
376             UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex

378             Select Case ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case e_OBJType.otPuertas 'Es una puerta
380                     Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
                End Select

382         ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
384             UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).ObjInfo.ObjIndex

386             Select Case ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType
            
                    Case e_OBJType.otPuertas 'Es una puerta
388                     Call AccionParaPuerta(Map, X, Y + 1, UserIndex)

                End Select
            End If

        End If

        
        Exit Sub

Accion_Err:
390     Call TraceError(Err.Number, Err.Description, "Acciones.Accion", Erl)

        
End Sub


Sub AccionParaYunque(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo AccionParaYunque_Err

        Dim Pos As t_WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y

106     If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
            ' Msg8=Estas demasiado lejos.
            Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
110     If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
            'Msg1071= Debes tener equipado un martillo de herrero para trabajar con el yunque.
            Call WriteLocaleMsg(UserIndex, "1071", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
114     If ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo <> 7 Then
            'Msg1072= La herramienta que tienes no es la correcta, necesitas un martillo de herrero para poder trabajar.
            Call WriteLocaleMsg(UserIndex, "1072", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

118     Call EnivarArmasConstruibles(UserIndex)
120     Call EnivarArmadurasConstruibles(UserIndex)
        Call SendCraftableElementRunes(UserIndex)
122     Call WriteShowBlacksmithForm(UserIndex)

        Exit Sub

AccionParaYunque_Err:
124     Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaYunque", Erl)

        
End Sub

Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal UserIndex As Integer, Optional ByVal SinDistancia As Boolean)
        On Error GoTo Handler

        Dim puerta As t_ObjData 'ver ReyarB
        
100     If Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2 And Not SinDistancia Then
102         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If


104     puerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex)
106     If puerta.Llave = 1 And Not SinDistancia Then
            If puerta.GrhIndex = 11445 Or puerta.GrhIndex = 11444 Or puerta.GrhIndex = 59878 Or puerta.GrhIndex = 59877 Then
            'Msg1073= Al parecer, alguien cerró esta puerta. Debe haber algún interruptor por algún lado...
            Call WriteLocaleMsg(UserIndex, "1073", e_FontTypeNames.FONTTYPE_INFO)
            Else
            'Msg1074= La puerta esta cerrada con llave.
            Call WriteLocaleMsg(UserIndex, "1074", e_FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If

110     If puerta.Cerrada = 1 Then 'Abre la puerta
112         MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexAbierta
114         Call BloquearPuerta(Map, X, Y, False)
            If puerta.GrhIndex = 11445 Or puerta.GrhIndex = 11444 Or puerta.GrhIndex = 59878 Or puerta.GrhIndex = 59877 Then
            'Msg1075= Has abierto la compuerta del ducto.
            Call WriteLocaleMsg(UserIndex, "1075", e_FontTypeNames.FONTTYPE_INFO)
            End If

        Else 'Cierra puerta
116         MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexCerrada
118         Call BloquearPuerta(Map, X, Y, True)

        End If

120     If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo = 1 Then
122          Call AccionParaPuerta(Map, X - 3, Y + 1, UserIndex, True)
        End If

124     Call modSendData.SendToAreaByPos(Map, x, y, PrepareMessageObjectCreate(MapData(Map, x, y).ObjInfo.ObjIndex, MapData(Map, x, y).ObjInfo.amount, x, y))
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

        Dim puerta As t_ObjData 'ver ReyarB


100     puerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex)

102     If puerta.Cerrada = 1 Then 'Abre la puerta
104         MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexAbierta
106         Call BloquearPuerta(Map, X, Y, False)

        Else 'Cierra puerta
108         MapData(Map, X, Y).ObjInfo.ObjIndex = puerta.IndexCerrada
110         Call BloquearPuerta(Map, X, Y, True)

        End If

112     Call modSendData.SendToAreaByPos(Map, x, y, PrepareMessageObjectCreate(MapData(Map, x, y).ObjInfo.ObjIndex, MapData(Map, x, y).ObjInfo.amount, x, y))

114     Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))

        Exit Sub

Handler:
116 Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaPuertaNpc", Erl)


End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

        On Error GoTo Handler

        Dim MiObj As t_Obj

100     'If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
102     If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
104         Call WriteShowSignal(UserIndex, MapData(Map, X, Y).ObjInfo.ObjIndex)
        Else
            Call WriteShowPapiro(UserIndex)
        End If
  
        'End If
        
        Exit Sub
        
Handler:
106 Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaCartel", Erl)

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    On Error GoTo Handler

        Dim Suerte As Byte
        Dim exito  As Byte
        Dim raise  As Integer
    
        Dim Pos    As t_WorldPos

100     Pos.Map = Map
102     Pos.X = X
104     Pos.Y = Y
    
106     With UserList(UserIndex)
    
108         If Distancia(Pos, .Pos) > 2 Then
110             ' Msg8=Estas demasiado lejos.
                Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

112         If MapInfo(Map).lluvia And Lloviendo Then
                'Msg1076= Esta lloviendo, no podés encender una fogata aquí.
                Call WriteLocaleMsg(UserIndex, "1076", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

116         If MapData(Map, X, Y).trigger = e_Trigger.ZONASEGURA Or MapInfo(Map).Seguro = 1 Then
                'Msg1077= En zona segura no podés hacer fogatas.
                Call WriteLocaleMsg(UserIndex, "1077", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

120         If MapData(Map, X - 1, Y).ObjInfo.ObjIndex = FOGATA Or _
               MapData(Map, X + 1, Y).ObjInfo.ObjIndex = FOGATA Or _
               MapData(Map, X, Y - 1).ObjInfo.ObjIndex = FOGATA Or _
               MapData(Map, X, Y + 1).ObjInfo.ObjIndex = FOGATA Then
           
                'Msg1078= Debes alejarte un poco de la otra fogata.
                Call WriteLocaleMsg(UserIndex, "1078", e_FontTypeNames.FONTTYPE_INFO)
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
                
                    Dim obj As t_Obj
142                 obj.ObjIndex = FOGATA
144                 obj.amount = 1
        
                    'Msg1079= Has prendido la fogata.
                    Call WriteLocaleMsg(UserIndex, "1079", e_FontTypeNames.FONTTYPE_INFO)
        
148                 Call MakeObj(obj, Map, X, Y)

                Else
        
                    'Msg1080= La ley impide realizar fogatas en las ciudades.
                    Call WriteLocaleMsg(UserIndex, "1080", e_FontTypeNames.FONTTYPE_INFO)
            
                    Exit Sub

                End If

            Else
        
                'Msg1081= No has podido hacer fuego.
                Call WriteLocaleMsg(UserIndex, "1081", e_FontTypeNames.FONTTYPE_INFO)

            End If
    
        End With

154     Call SubirSkill(UserIndex, Supervivencia)

        Exit Sub
        
Handler:
156 Call TraceError(Err.Number, Err.Description, "Acciones.AccionParaRamita", Erl)


End Sub
