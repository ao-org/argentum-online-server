Attribute VB_Name = "ModLadder"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)

Private theTime      As SYSTEMTIME

Public ClaveApertura As String

Public PaquetesCount As Long

Private Type SYSTEMTIME

    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer

End Type

Public Enum Accion_Barra

    Runa = 1
    Resucitar = 2
    Intermundia = 3
    GoToPareja = 5
    Hogar = 6
    CancelarAccion = 99

End Enum

Public Function GetTickCount() As Long
        
        On Error GoTo GetTickCount_Err
    
        
    
100     GetTickCount = timeGetTime And &H7FFFFFFF
    
        
        Exit Function

GetTickCount_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLadder.GetTickCount", Erl)

        
End Function

Function GetTimeFormated() As String
        
        On Error GoTo GetTimeFormated_Err
        
        Dim Elapsed As Single
100     Elapsed = (GetTickCount() - HoraMundo) / DuracionDia
        
        Dim Mins As Long
102     Mins = (Elapsed - Fix(Elapsed)) * 1440

        Dim Horita    As Byte

        Dim Minutitos As Byte

104     Horita = Fix(Mins / 60)
106     Minutitos = Mins Mod 60

108     GetTimeFormated = Right$("00" & Horita, 2) & ":" & Right$("00" & Minutitos, 2)

        
        Exit Function

GetTimeFormated_Err:
110     Call TraceError(Err.Number, Err.Description, "ModLadder.GetTimeFormated", Erl)
112
        
End Function

Public Sub GetHoraActual()
        
        On Error GoTo GetHoraActual_Err
        
100     GetSystemTime theTime

102     HoraActual = (theTime.wHour - 3)

104     If HoraActual = -3 Then HoraActual = 21
106     If HoraActual = -2 Then HoraActual = 22
108     If HoraActual = -1 Then HoraActual = 23
110     frmMain.lblhora.Caption = HoraActual & ":" & Format(theTime.wMinute, "00") & ":" & Format(theTime.wSecond, "00")
112     HoraEvento = HoraActual

        
        Exit Sub

GetHoraActual_Err:
114     Call TraceError(Err.Number, Err.Description, "ModLadder.GetHoraActual", Erl)
116
        
End Sub

Public Function DarNameMapa(ByVal Map As Long) As String
        
        On Error GoTo DarNameMapa_Err
        
100     DarNameMapa = MapInfo(Map).map_name

        
        Exit Function

DarNameMapa_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLadder.DarNameMapa", Erl)
104
        
End Function

Public Sub CompletarAccionFin(ByVal UserIndex As Integer)
        
        On Error GoTo CompletarAccionFin_Err
        

        Dim obj  As ObjData

        Dim Slot As Byte

100     Select Case UserList(UserIndex).Accion.TipoAccion

            Case Accion_Barra.Runa
102             obj = ObjData(UserList(UserIndex).Accion.RunaObj)
104             Slot = UserList(UserIndex).Accion.ObjSlot

106             Select Case obj.TipoRuna

                    Case 1 'Cuando esta muerto lleva al lugar de Origen

                        Dim DeDonde As CityWorldPos

                        Dim Map     As Integer

                        Dim X       As Byte

                        Dim Y       As Byte
        
108                     If UserList(UserIndex).flags.Muerto = 0 Then

110                         Select Case UserList(UserIndex).Hogar

                                Case eCiudad.cUllathorpe
112                                 DeDonde = CityUllathorpe
                        
114                             Case eCiudad.cNix
116                                 DeDonde = CityNix
            
118                             Case eCiudad.cBanderbill
120                                 DeDonde = CityBanderbill
                    
122                             Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
124                                 DeDonde = CityLindos
                        
126                             Case eCiudad.cArghal
128                                 DeDonde = CityArghal
                        
130                             Case eCiudad.cArkhein
132                                 DeDonde = CityArkhein
                        
134                             Case Else
136                                 DeDonde = CityUllathorpe

                            End Select

138                         Map = DeDonde.Map
140                         X = DeDonde.X
142                         Y = DeDonde.Y
                        Else

144                         If MapInfo(UserList(UserIndex).Pos.Map).ResuCiudad <> 0 Then

146                             Select Case MapInfo(UserList(UserIndex).Pos.Map).ResuCiudad

                                    Case eCiudad.cUllathorpe
148                                     DeDonde = CityUllathorpe
                        
150                                 Case eCiudad.cNix
152                                     DeDonde = CityNix
            
154                                 Case eCiudad.cBanderbill
156                                     DeDonde = CityBanderbill
                    
158                                 Case eCiudad.cLindos
160                                     DeDonde = CityLindos
                        
162                                 Case eCiudad.cArghal
164                                     DeDonde = CityArghal
                        
166                                 Case eCiudad.cArkhein
168                                     DeDonde = CityArkhein
                        
170                                 Case Else
172                                     DeDonde = CityUllathorpe

                                End Select

                            Else

174                             Select Case UserList(UserIndex).Hogar

                                    Case eCiudad.cUllathorpe
176                                     DeDonde = CityUllathorpe
                        
178                                 Case eCiudad.cNix
180                                     DeDonde = CityNix
            
182                                 Case eCiudad.cBanderbill
184                                     DeDonde = CityBanderbill
                    
186                                 Case eCiudad.cLindos
188                                     DeDonde = CityLindos
                        
190                                 Case eCiudad.cArghal
192                                     DeDonde = CityArghal
                        
194                                 Case eCiudad.cArkhein
196                                     DeDonde = CityArkhein
                        
198                                 Case Else
200                                     DeDonde = CityUllathorpe

                                End Select

                            End If
                
202                         Map = DeDonde.MapaResu
204                         X = DeDonde.ResuX
206                         Y = DeDonde.ResuY
                
                            Dim Resu As Boolean
                
208                         Resu = True
            
                        End If
                
210                     Call FindLegalPos(UserIndex, Map, X, Y)
212                     Call WarpUserChar(UserIndex, Map, X, Y, True)
214                     Call WriteConsoleMsg(UserIndex, "Has regresado a tu ciudad de origen.", FontTypeNames.FONTTYPE_WARNING)

                        'Call WriteFlashScreen(UserIndex, &HA4FFFF, 150, True)
216                     If UserList(UserIndex).flags.Navegando = 1 Then

                            Dim barca As ObjData

218                         barca = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
220                         Call DoNavega(UserIndex, barca, UserList(UserIndex).Invent.BarcoSlot)

                        End If
                
222                     If Resu Then
                
                            UserList(UserIndex).Counters.TimerBarra = 5
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, UserList(UserIndex).Counters.TimerBarra, False))
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Counters.TimerBarra, Accion_Barra.Resucitar))

                
234                         UserList(UserIndex).Accion.AccionPendiente = True
236                         UserList(UserIndex).Accion.Particula = ParticulasIndex.Resucitar
238                         UserList(UserIndex).Accion.TipoAccion = Accion_Barra.Resucitar

240                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                            'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...", FontTypeNames.FONTTYPE_INFO)
242                         Call WriteLocaleMsg(UserIndex, "82", FontTypeNames.FONTTYPE_INFOIAO)

                        End If
                
244                     If Not Resu Then
246                         UserList(UserIndex).Accion.AccionPendiente = False
248                         UserList(UserIndex).Accion.Particula = 0
250                         UserList(UserIndex).Accion.TipoAccion = Accion_Barra.CancelarAccion

                        End If

252                     UserList(UserIndex).Accion.HechizoPendiente = 0
254                     UserList(UserIndex).Accion.RunaObj = 0
256                     UserList(UserIndex).Accion.ObjSlot = 0
              
258                 Case 2
260                     Map = obj.HastaMap
262                     X = obj.HastaX
264                     Y = obj.HastaY
            
266                     If obj.DesdeMap = 0 Then
268                         Call FindLegalPos(UserIndex, Map, X, Y)
270                         Call WarpUserChar(UserIndex, Map, X, Y, True)
272                         Call WriteConsoleMsg(UserIndex, "Te has teletransportado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
274                         Call QuitarUserInvItem(UserIndex, Slot, 1)
276                         Call UpdateUserInv(False, UserIndex, Slot)
                        Else

278                         If UserList(UserIndex).Pos.Map <> obj.DesdeMap Then
280                             Call WriteConsoleMsg(UserIndex, "Esta runa no puede ser usada desde aquí.", FontTypeNames.FONTTYPE_INFO)
                            Else
282                             Call QuitarUserInvItem(UserIndex, Slot, 1)
284                             Call UpdateUserInv(False, UserIndex, Slot)
286                             Call FindLegalPos(UserIndex, Map, X, Y)
288                             Call WarpUserChar(UserIndex, Map, X, Y, True)
290                             Call WriteConsoleMsg(UserIndex, "Te has teletransportado por el mundo.", FontTypeNames.FONTTYPE_WARNING)

                            End If

                        End If
        
292                     UserList(UserIndex).Accion.Particula = 0
294                     UserList(UserIndex).Accion.TipoAccion = Accion_Barra.CancelarAccion
296                     UserList(UserIndex).Accion.HechizoPendiente = 0
298                     UserList(UserIndex).Accion.RunaObj = 0
300                     UserList(UserIndex).Accion.ObjSlot = 0
302                     UserList(UserIndex).Accion.AccionPendiente = False

            
                End Select
                
304         Case Accion_Barra.Hogar
306             Call HomeArrival(UserIndex)
308             UserList(UserIndex).Accion.AccionPendiente = False
310             UserList(UserIndex).Accion.Particula = 0
312             UserList(UserIndex).Accion.TipoAccion = Accion_Barra.CancelarAccion
            

314         Case Accion_Barra.Intermundia
        
316             If UserList(UserIndex).flags.Muerto = 0 Then

                    Dim uh As Integer

                    Dim Mapaf, Xf, Yf As Integer

318                 uh = UserList(UserIndex).Accion.HechizoPendiente
    
320                 Mapaf = Hechizos(uh).TeleportXMap
322                 Xf = Hechizos(uh).TeleportXX
324                 Yf = Hechizos(uh).TeleportXY
    
326                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(uh).wav, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))  'Esta linea faltaba. Pablo (ToxicWaste)
328                 Call WriteConsoleMsg(UserIndex, "¡Has abierto la puerta a intermundia!", FontTypeNames.FONTTYPE_INFO)
330                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, -1, True))
332                 UserList(UserIndex).flags.Portal = 10
334                 UserList(UserIndex).flags.PortalMDestino = Mapaf
336                 UserList(UserIndex).flags.PortalYDestino = Xf
338                 UserList(UserIndex).flags.PortalXDestino = Yf
                
                    Dim Mapa As Integer

340                 Mapa = UserList(UserIndex).flags.PortalM
342                 X = UserList(UserIndex).flags.PortalX
344                 Y = UserList(UserIndex).flags.PortalY
346                 MapData(Mapa, X, Y).Particula = ParticulasIndex.TpVerde
348                 MapData(Mapa, X, Y).TimeParticula = -1
350                 MapData(Mapa, X, Y).TileExit.Map = UserList(UserIndex).flags.PortalMDestino
352                 MapData(Mapa, X, Y).TileExit.X = UserList(UserIndex).flags.PortalXDestino
354                 MapData(Mapa, X, Y).TileExit.Y = UserList(UserIndex).flags.PortalYDestino
                
                    'Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.Intermundia, -1))
356                 Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.TpVerde, -1))
                
358                 Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageLightFXToFloor(X, Y, &HFF80C0, 105))

                End If
                    
360             UserList(UserIndex).Accion.Particula = 0
362             UserList(UserIndex).Accion.TipoAccion = Accion_Barra.CancelarAccion
364             UserList(UserIndex).Accion.HechizoPendiente = 0
366             UserList(UserIndex).Accion.RunaObj = 0
368             UserList(UserIndex).Accion.ObjSlot = 0
370             UserList(UserIndex).Accion.AccionPendiente = False
            
                '
372         Case Accion_Barra.Resucitar
374             Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)
376             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 250, True))
378             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("117", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
380             Call RevivirUsuario(UserIndex, True)
                
382             UserList(UserIndex).Accion.Particula = 0
384             UserList(UserIndex).Accion.TipoAccion = Accion_Barra.CancelarAccion
386             UserList(UserIndex).Accion.HechizoPendiente = 0
388             UserList(UserIndex).Accion.RunaObj = 0
390             UserList(UserIndex).Accion.ObjSlot = 0
392             UserList(UserIndex).Accion.AccionPendiente = False
                      
        End Select
               
        Exit Sub

CompletarAccionFin_Err:
394     Call TraceError(Err.Number, Err.Description, "ModLadder.CompletarAccionFin", Erl)
396
        
End Sub

Public Function TieneObjEnInv(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ObjIndex2 As Integer = 0) As Boolean
        On Error GoTo TieneObjEnInv_Err
        
        'Devuelve el slot del inventario donde se encuentra el obj
        'Creaado por Ladder 25/09/2014
        Dim i As Byte

100     For i = 1 To 36

102         If UserList(UserIndex).Invent.Object(i).ObjIndex = ObjIndex Then
104             TieneObjEnInv = True
                Exit Function

            End If

106         If ObjIndex2 > 0 Then
108             If UserList(UserIndex).Invent.Object(i).ObjIndex = ObjIndex2 Then
110                 TieneObjEnInv = True
                    Exit Function

                End If

            End If

112     Next i

114     TieneObjEnInv = False

        
        Exit Function

TieneObjEnInv_Err:
116     Call TraceError(Err.Number, Err.Description, "ModLadder.TieneObjEnInv", Erl)
118
        
End Function


Public Function CantidadObjEnInv(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
        
        On Error GoTo CantidadObjEnInv_Err
        
        'Devuelve el amount si tiene el ObjIndex en el inventario, sino devuelve 0
        'Creaado por Ladder 25/09/2014
        Dim i As Byte

100     For i = 1 To 36

102         If UserList(UserIndex).Invent.Object(i).ObjIndex = ObjIndex Then
104             CantidadObjEnInv = UserList(UserIndex).Invent.Object(i).amount
                Exit Function
            End If


106     Next i

108     CantidadObjEnInv = 0

        
        Exit Function

CantidadObjEnInv_Err:
110     Call TraceError(Err.Number, Err.Description, "ModLadder.CantidadObjEnInv", Erl)
112
        
End Function

Public Function SumarTiempo(segundos As Integer) As String
        
        On Error GoTo SumarTiempo_Err
        

        Dim a As Variant, b As Variant

        Dim X As Integer

        Dim T As String

100     T = "00:00:00" 'Lo inicializamos en 0 horas, 0 minutos, 0 segundos
102     a = Format("00:00:01", "hh:mm:ss") 'guardamos en una variable el formato de 1 segundos

104     For X = 1 To segundos 'hacemos segundo a segundo
106         b = Format(T, "hh:mm:ss") 'En B guardamos un formato de hora:minuto:segundo segun lo que tenia T
108         T = Format(TimeValue(a) + TimeValue(b), "hh:mm:ss") 'asignamos a T la suma de A + B (osea, sumamos logicamente 1 segundo)
110     Next X

112     SumarTiempo = T 'a la funcion le damos el valor que hallamos en T

        
        Exit Function

SumarTiempo_Err:
114     Call TraceError(Err.Number, Err.Description, "ModLadder.SumarTiempo", Erl)
116
        
End Function

Public Sub AgregarAConsola(ByVal Text As String)
        
        On Error GoTo AgregarAConsola_Err
        

100     frmMain.List1.AddItem (Text)

        
        Exit Sub

AgregarAConsola_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLadder.AgregarAConsola", Erl)
104
        
End Sub

' TODO: Crear enum para la respuesta
Function PuedeUsarObjeto(UserIndex As Integer, ByVal ObjIndex As Integer, Optional ByVal writeInConsole As Boolean = False) As Byte
        On Error GoTo PuedeUsarObjeto_Err

        Dim Objeto As ObjData
        Dim Msg As String, i As Long
100     Objeto = ObjData(ObjIndex)
                
102     If EsGM(UserIndex) Then
104         PuedeUsarObjeto = 0
106         Msg = ""

108     ElseIf Objeto.Newbie = 1 And Not EsNewbie(UserIndex) Then
110         PuedeUsarObjeto = 7
112         Msg = "Solo los newbies pueden usar este objeto."
            
114     ElseIf UserList(UserIndex).Stats.ELV < Objeto.MinELV Then
116         PuedeUsarObjeto = 6
118         Msg = "Necesitas ser nivel " & Objeto.MinELV & " para usar este objeto."

120     ElseIf Not FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
122         PuedeUsarObjeto = 3
124         Msg = "Tu facción no te permite utilizarlo."

126     ElseIf Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
128         PuedeUsarObjeto = 2
130         Msg = "Tu clase no puede utilizar este objeto."

132     ElseIf Not SexoPuedeUsarItem(UserIndex, ObjIndex) Then
134         PuedeUsarObjeto = 1
136         Msg = "Tu sexo no puede utilizar este objeto."

138     ElseIf Not RazaPuedeUsarItem(UserIndex, ObjIndex) Then
140         PuedeUsarObjeto = 5
142         Msg = "Tu raza no puede utilizar este objeto."
144     ElseIf (Objeto.SkillIndex > 0) Then
146         If (UserList(UserIndex).Stats.UserSkills(Objeto.SkillIndex) < Objeto.SkillRequerido) Then
148             PuedeUsarObjeto = 4
150             Msg = "Necesitas " & Objeto.SkillRequerido & " puntos en " & SkillsNames(Objeto.SkillIndex) & " para usar este item."
            Else
152             PuedeUsarObjeto = 0
154             Msg = ""
            End If
        Else
156         PuedeUsarObjeto = 0
158         Msg = ""
        End If

160     If writeInConsole And Msg <> "" Then Call WriteConsoleMsg(UserIndex, Msg, FontTypeNames.FONTTYPE_INFO)

        Exit Function

PuedeUsarObjeto_Err:
162     Call TraceError(Err.Number, Err.Description, "ModLadder.PuedeUsarObjeto", Erl)
164     'Resume Next ' WyroX: Si hay error que salga directamente

End Function

Public Function RequiereOxigeno(ByVal UserMap) As Boolean
        On Error GoTo RequiereOxigeno_Err
        
100     RequiereOxigeno = (UserMap = 265) Or _
                          (UserMap = 266) Or _
                          (UserMap = 267)
        
        Exit Function

RequiereOxigeno_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLadder.RequiereOxigeno", Erl)
104
        
End Function
