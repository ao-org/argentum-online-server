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
    BattleModo = 4
    GoToPareja = 5
    CancelarAccion = 99

End Enum

Public Function GetTickCount() As Long
    
100     GetTickCount = timeGetTime And &H7FFFFFFF
    
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
110     Call RegistrarError(Err.Number, Err.description, "ModLadder.GetTimeFormated", Erl)
112     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "ModLadder.GetHoraActual", Erl)
116     Resume Next
        
End Sub

Public Function DarNameMapa(ByVal Map As Long) As String
        
        On Error GoTo DarNameMapa_Err
        
100     DarNameMapa = MapInfo(Map).map_name

        
        Exit Function

DarNameMapa_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModLadder.DarNameMapa", Erl)
104     Resume Next
        
End Function

Public Sub CompletarAccionFin(ByVal Userindex As Integer)
        
        On Error GoTo CompletarAccionFin_Err
        

        Dim obj  As ObjData

        Dim slot As Byte

100     Select Case UserList(Userindex).Accion.TipoAccion

            Case Accion_Barra.Runa
102             obj = ObjData(UserList(Userindex).Accion.RunaObj)
104             slot = UserList(Userindex).Accion.ObjSlot

106             Select Case obj.TipoRuna

                    Case 1 'Cuando esta muerto lleva al lugar de Origen

                        Dim DeDonde As CityWorldPos

                        Dim Map     As Integer

                        Dim X       As Byte

                        Dim Y       As Byte
        
108                     If UserList(Userindex).flags.Muerto = 0 Then

110                         Select Case UserList(Userindex).Hogar

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
                        
130                             Case eCiudad.CHillidan
132                                 DeDonde = CityHillidan
                        
134                             Case Else
136                                 DeDonde = CityUllathorpe

                            End Select

138                         Map = DeDonde.Map
140                         X = DeDonde.X
142                         Y = DeDonde.Y
                        Else

144                         If MapInfo(UserList(Userindex).Pos.Map).extra2 <> 0 Then

146                             Select Case MapInfo(UserList(Userindex).Pos.Map).extra2

                                    Case eCiudad.cUllathorpe
148                                     DeDonde = CityUllathorpe
                        
150                                 Case eCiudad.cNix
152                                     DeDonde = CityNix
            
154                                 Case eCiudad.cBanderbill
156                                     DeDonde = CityBanderbill
                    
158                                 Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
160                                     DeDonde = CityLindos
                        
162                                 Case eCiudad.cArghal
164                                     DeDonde = CityArghal
                        
166                                 Case eCiudad.CHillidan
168                                     DeDonde = CityHillidan
                        
170                                 Case Else
172                                     DeDonde = CityUllathorpe

                                End Select

                            Else

174                             Select Case UserList(Userindex).Hogar

                                    Case eCiudad.cUllathorpe
176                                     DeDonde = CityUllathorpe
                        
178                                 Case eCiudad.cNix
180                                     DeDonde = CityNix
            
182                                 Case eCiudad.cBanderbill
184                                     DeDonde = CityBanderbill
                    
186                                 Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
188                                     DeDonde = CityLindos
                        
190                                 Case eCiudad.cArghal
192                                     DeDonde = CityArghal
                        
194                                 Case eCiudad.CHillidan
196                                     DeDonde = CityHillidan
                        
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
                
210                     Call FindLegalPos(Userindex, Map, X, Y)
212                     Call WarpUserChar(Userindex, Map, X, Y, True)
214                     Call WriteConsoleMsg(Userindex, "Has regresado a tu ciudad de origen.", FontTypeNames.FONTTYPE_WARNING)

                        'Call WriteEfectToScreen(UserIndex, &HA4FFFF, 150, True)
216                     If UserList(Userindex).flags.Navegando = 1 Then

                            Dim barca As ObjData

218                         barca = ObjData(UserList(Userindex).Invent.BarcoObjIndex)
220                         Call DoNavega(Userindex, barca, UserList(Userindex).Invent.BarcoSlot)

                        End If
                
222                     If Resu Then
                
224                         If UserList(Userindex).donador.activo = 0 Then ' Donador no espera tiempo
226                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Resucitar, 400, False))
228                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 400, Accion_Barra.Resucitar))
                            Else
230                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Resucitar, 10, False))
232                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 10, Accion_Barra.Resucitar))

                            End If
                
234                         UserList(Userindex).Accion.AccionPendiente = True
236                         UserList(Userindex).Accion.Particula = ParticulasIndex.Resucitar
238                         UserList(Userindex).Accion.TipoAccion = Accion_Barra.Resucitar

240                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("104", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...", FontTypeNames.FONTTYPE_INFO)
242                         Call WriteLocaleMsg(Userindex, "82", FontTypeNames.FONTTYPE_INFOIAO)

                        End If
                
244                     If Not Resu Then
246                         UserList(Userindex).Accion.AccionPendiente = False
248                         UserList(Userindex).Accion.Particula = 0
250                         UserList(Userindex).Accion.TipoAccion = Accion_Barra.CancelarAccion

                        End If

252                     UserList(Userindex).Accion.HechizoPendiente = 0
254                     UserList(Userindex).Accion.RunaObj = 0
256                     UserList(Userindex).Accion.ObjSlot = 0
              
258                 Case 2
260                     Map = obj.HastaMap
262                     X = obj.HastaX
264                     Y = obj.HastaY
            
266                     If obj.DesdeMap = 0 Then
268                         Call FindLegalPos(Userindex, Map, X, Y)
270                         Call WarpUserChar(Userindex, Map, X, Y, True)
272                         Call WriteConsoleMsg(Userindex, "Te has teletransportado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
274                         Call QuitarUserInvItem(Userindex, slot, 1)
276                         Call UpdateUserInv(False, Userindex, slot)
                        Else

278                         If UserList(Userindex).Pos.Map <> obj.DesdeMap Then
280                             Call WriteConsoleMsg(Userindex, "Esta runa no puede ser usada desde aquí.", FontTypeNames.FONTTYPE_INFO)
                            Else
282                             Call QuitarUserInvItem(Userindex, slot, 1)
284                             Call UpdateUserInv(False, Userindex, slot)
286                             Call FindLegalPos(Userindex, Map, X, Y)
288                             Call WarpUserChar(Userindex, Map, X, Y, True)
290                             Call WriteConsoleMsg(Userindex, "Te has teletransportado por el mundo.", FontTypeNames.FONTTYPE_WARNING)

                            End If

                        End If
        
292                     UserList(Userindex).Accion.Particula = 0
294                     UserList(Userindex).Accion.TipoAccion = Accion_Barra.CancelarAccion
296                     UserList(Userindex).Accion.HechizoPendiente = 0
298                     UserList(Userindex).Accion.RunaObj = 0
300                     UserList(Userindex).Accion.ObjSlot = 0
302                     UserList(Userindex).Accion.AccionPendiente = False

304                 Case 3

                        Dim parejaindex As Integer
    
306                     If Not UserList(Userindex).flags.BattleModo Then
                    
                            ' If UserList(UserIndex).donador.activo = 1 Then
308                         If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
310                             If UserList(Userindex).flags.Casado = 1 Then
312                                 parejaindex = NameIndex(UserList(Userindex).flags.Pareja)
                            
314                                 If parejaindex > 0 Then
316                                     If Not UserList(parejaindex).flags.BattleModo Then
318                                         Call WarpToLegalPos(Userindex, UserList(parejaindex).Pos.Map, UserList(parejaindex).Pos.X, UserList(parejaindex).Pos.Y, True)
320                                         Call WriteConsoleMsg(Userindex, "Te has teletransportado hacia tu pareja.", FontTypeNames.FONTTYPE_INFOIAO)
322                                         Call WriteConsoleMsg(parejaindex, "Tu pareja se ha teletransportado hacia vos.", FontTypeNames.FONTTYPE_INFOIAO)
                                        Else
324                                         Call WriteConsoleMsg(Userindex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)

                                        End If
                                    
                                    Else
326                                     Call WriteConsoleMsg(Userindex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)

                                    End If

                                Else
328                                 Call WriteConsoleMsg(Userindex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)

                                End If

                            Else
330                             Call WriteConsoleMsg(Userindex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If
                    
                            'Else
                            '   Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
                            '  End If
                        Else
332                         Call WriteConsoleMsg(Userindex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
            
                        End If
            
                End Select

334         Case Accion_Barra.Intermundia
        
336             If UserList(Userindex).flags.Muerto = 0 Then

                    Dim uh As Integer

                    Dim Mapaf, Xf, Yf As Integer

338                 uh = UserList(Userindex).Accion.HechizoPendiente
    
340                 Mapaf = Hechizos(uh).TeleportXMap
342                 Xf = Hechizos(uh).TeleportXX
344                 Yf = Hechizos(uh).TeleportXY
    
346                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(uh).wav, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))  'Esta linea faltaba. Pablo (ToxicWaste)
348                 Call WriteConsoleMsg(Userindex, "¡Has abierto la puerta a intermundia!", FontTypeNames.FONTTYPE_INFO)
350                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, -1, True))
352                 UserList(Userindex).flags.Portal = 10
354                 UserList(Userindex).flags.PortalMDestino = Mapaf
356                 UserList(Userindex).flags.PortalYDestino = Xf
358                 UserList(Userindex).flags.PortalXDestino = Yf
                
                    Dim Mapa As Integer

360                 Mapa = UserList(Userindex).flags.PortalM
362                 X = UserList(Userindex).flags.PortalX
364                 Y = UserList(Userindex).flags.PortalY
366                 MapData(Mapa, X, Y).Particula = ParticulasIndex.TpVerde
368                 MapData(Mapa, X, Y).TimeParticula = -1
370                 MapData(Mapa, X, Y).TileExit.Map = UserList(Userindex).flags.PortalMDestino
372                 MapData(Mapa, X, Y).TileExit.X = UserList(Userindex).flags.PortalXDestino
374                 MapData(Mapa, X, Y).TileExit.Y = UserList(Userindex).flags.PortalYDestino
                
                    'Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.Intermundia, -1))
376                 Call SendData(SendTarget.toMap, UserList(Userindex).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.TpVerde, -1))
                
378                 Call SendData(SendTarget.toMap, UserList(Userindex).flags.PortalM, PrepareMessageLightFXToFloor(X, Y, &HFF80C0, 105))

                End If
                    
380             UserList(Userindex).Accion.Particula = 0
382             UserList(Userindex).Accion.TipoAccion = Accion_Barra.CancelarAccion
384             UserList(Userindex).Accion.HechizoPendiente = 0
386             UserList(Userindex).Accion.RunaObj = 0
388             UserList(Userindex).Accion.ObjSlot = 0
390             UserList(Userindex).Accion.AccionPendiente = False
            
                '
392         Case Accion_Barra.Resucitar
394             Call WriteConsoleMsg(Userindex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)
396             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Resucitar, 250, True))
398             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("204", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
400             Call RevivirUsuario(Userindex)
                
402             UserList(Userindex).Accion.Particula = 0
404             UserList(Userindex).Accion.TipoAccion = Accion_Barra.CancelarAccion
406             UserList(Userindex).Accion.HechizoPendiente = 0
408             UserList(Userindex).Accion.RunaObj = 0
410             UserList(Userindex).Accion.ObjSlot = 0
412             UserList(Userindex).Accion.AccionPendiente = False
        
414         Case Accion_Barra.BattleModo
        
416             If UserList(Userindex).flags.BattleModo = 1 Then
418                 Call Cerrar_Usuario(Userindex)
                
                    ' Dim mapaa As Integer
                    '  Dim xa As Integer
                    ' Dim ya As Integer
                    ' mapaa = CInt(ReadField(1, GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Position"), 45))
                    ' xa = CInt(ReadField(2, GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Position"), 45))
                    ' ya = CInt(ReadField(3, GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Position"), 45))

                    ' Call WarpUserChar(UserIndex, mapaa, xa, ya, False)
                
                    ' Call RelogearUser(UserIndex, UserList(UserIndex).name, UserList(UserIndex).cuenta)
                Else
                
420                 If UserList(Userindex).flags.invisible = 1 Or UserList(Userindex).flags.Oculto = 1 Then

422                     UserList(Userindex).flags.Oculto = 0
424                     UserList(Userindex).flags.invisible = 0
426                     UserList(Userindex).Counters.TiempoOculto = 0
428                     UserList(Userindex).Counters.Invisibilidad = 0
                
430                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, False))

                    End If
                
432                 Call SaveUser(Userindex)  'Guardo el PJ

434                 X = 50
436                 Y = 54
438                 Call FindLegalPos(Userindex, 336, X, Y)
                
440                 Call WarpUserChar(Userindex, 336, X, Y, True)
                
442                 UserList(Userindex).flags.BattleModo = 1

444                 If UserList(Userindex).flags.Subastando Then
446                     Call CancelarSubasta

                    End If
                
448                 Call AumentarPJ(Userindex)
450                 Call WriteConsoleMsg(Userindex, "Battle> Ahora tu personaje se encuentra en modo batalla. Recuerda que todos los cambios que se realicen sobre éste no tendran efecto mientras te encuentres aquí. Cuando desees salir, solamente toca ESC o escribe /SALIR y relogea con tu personaje.", FontTypeNames.FONTTYPE_CITIZEN)
                
                End If

452             UserList(Userindex).Accion.AccionPendiente = False
454             UserList(Userindex).Accion.Particula = 0
456             UserList(Userindex).Accion.TipoAccion = Accion_Barra.CancelarAccion
458             UserList(Userindex).Accion.HechizoPendiente = 0
460             UserList(Userindex).Accion.RunaObj = 0
462             UserList(Userindex).Accion.ObjSlot = 0
                
464         Case Accion_Barra.GoToPareja
    
466             If Not UserList(Userindex).flags.BattleModo Then
                    
                    ' If UserList(UserIndex).donador.activo = 1 Then
468                 If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
470                     If UserList(Userindex).flags.Casado = 1 Then
472                         parejaindex = NameIndex(UserList(Userindex).flags.Pareja)
                            
474                         If parejaindex > 0 Then
476                             If Not UserList(parejaindex).flags.BattleModo Then
478                                 Call WarpToLegalPos(Userindex, UserList(parejaindex).Pos.Map, UserList(parejaindex).Pos.X, UserList(parejaindex).Pos.Y, True)
480                                 Call WriteConsoleMsg(Userindex, "Te has teletransportado hacia tu pareja.", FontTypeNames.FONTTYPE_INFOIAO)
482                                 Call WriteConsoleMsg(parejaindex, "Tu pareja se ha teletransportado hacia vos.", FontTypeNames.FONTTYPE_INFOIAO)
                                Else
484                                 Call WriteConsoleMsg(Userindex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)

                                End If
                                    
                            Else
486                             Call WriteConsoleMsg(Userindex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If

                        Else
488                         Call WriteConsoleMsg(Userindex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

                    Else
490                     Call WriteConsoleMsg(Userindex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If
                    
                    ' Else
                    ' Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
                    'End If
                Else
492                 Call WriteConsoleMsg(Userindex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
            
                End If
            
        End Select
       
        
        Exit Sub

CompletarAccionFin_Err:
494     Call RegistrarError(Err.Number, Err.description, "ModLadder.CompletarAccionFin", Erl)
496     Resume Next
        
End Sub

Public Function General_Get_Line_Count(ByVal FileName As String) As Long

        '**************************************************************
        'Author: Augusto José Rando
        'Last Modify Date: 6/11/2005
        '
        '**************************************************************
        On Error GoTo ErrorHandler

        Dim n As Integer, tmpStr As String

100     If LenB(FileName) Then
102         n = FreeFile()
    
104         Open FileName For Input As #n
    
106         Do While Not EOF(n)
108             General_Get_Line_Count = General_Get_Line_Count + 1
110             Line Input #n, tmpStr
            Loop
    
112         Close n

        End If

        Exit Function

ErrorHandler:

End Function

Public Function Integer_To_String(ByVal Var As Integer) As String
        
        On Error GoTo Integer_To_String_Err
        

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
        Dim temp As String
        
        'Convertimos a hexa
100     temp = hex$(Var)
    
        'Nos aseguramos tenga 4 bytes de largo
102     While Len(temp) < 4

104         temp = "0" & temp
        Wend
    
        'Convertimos a string
106     Integer_To_String = Chr$(val("&H" & Left$(temp, 2))) & Chr$(val("&H" & Right$(temp, 2)))
        Exit Function

ErrorHandler:

        
        Exit Function

Integer_To_String_Err:
108     Call RegistrarError(Err.Number, Err.description, "ModLadder.Integer_To_String", Erl)
110     Resume Next
        
End Function

Public Function String_To_Integer(ByRef str As String, ByVal Start As Integer) As Integer

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
        On Error GoTo Error_Handler
    
        Dim temp_str As String
    
        'Asergurarse sea válido
100     If Len(str) < Start - 1 Or Len(str) = 0 Then Exit Function
    
        'Convertimos a hexa el valor ascii del segundo byte
102     temp_str = hex$(Asc(mid$(str, Start + 1, 1)))
    
        'Nos aseguramos tenga 2 bytes (los ceros a la izquierda cuentan por ser el segundo byte)
104     While Len(temp_str) < 2

106         temp_str = "0" & temp_str
        Wend
    
        'Convertimos a integer
108     String_To_Integer = val("&H" & hex$(Asc(mid$(str, Start, 1))) & temp_str)
            
        Exit Function
        
Error_Handler:
        
End Function

Public Function Byte_To_String(ByVal Var As Byte) As String
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        'Convierte un byte a string
        '**************************************************************
        
        On Error GoTo Byte_To_String_Err
        
100     Byte_To_String = Chr$(val("&H" & hex$(Var)))
        Exit Function

ErrorHandler:

        
        Exit Function

Byte_To_String_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModLadder.Byte_To_String", Erl)
104     Resume Next
        
End Function

Public Function String_To_Byte(ByRef str As String, ByVal Start As Integer) As Byte

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
        On Error GoTo Error_Handler
    
100     If Len(str) < Start Then Exit Function
    
102     String_To_Byte = Asc(mid$(str, Start, 1))
    
        Exit Function
        
Error_Handler:

End Function

Public Function Long_To_String(ByVal Var As Long) As String
        
        On Error GoTo Long_To_String_Err
        

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
        'No aceptamos valores que usen los 4 últimos its
100     If Var > &HFFFFFFF Then GoTo ErrorHandler
    
        Dim temp As String
    
        'Vemos si el cuarto byte es cero
102     If (Var And &HFF&) = 0 Then Var = Var Or &H80000001
    
        'Vemos si el tercer byte es cero
104     If (Var And &HFF00&) = 0 Then Var = Var Or &H40000100
    
        'Vemos si el segundo byte es cero
106     If (Var And &HFF0000) = 0 Then Var = Var Or &H20010000
    
        'Vemos si el primer byte es cero
108     If Var < &H1000000 Then Var = Var Or &H10000000
    
        'Convertimos a hexa
110     temp = hex$(Var)
    
        'Nos aseguramos tenga 8 bytes de largo
112     While Len(temp) < 8

114         temp = "0" & temp
        Wend
    
        'Convertimos a string
116     Long_To_String = Chr$(val("&H" & Left$(temp, 2))) & Chr$(val("&H" & mid$(temp, 3, 2))) & Chr$(val("&H" & mid$(temp, 5, 2))) & Chr$(val("&H" & mid$(temp, 7, 2)))
        Exit Function

ErrorHandler:

        
        Exit Function

Long_To_String_Err:
118     Call RegistrarError(Err.Number, Err.description, "ModLadder.Long_To_String", Erl)
120     Resume Next
        
End Function

Public Function String_To_Long(ByRef str As String, ByVal Start As Integer) As Long
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 3/12/2005
        '
        '**************************************************************
    
        On Error GoTo ErrorHandler
        
100     If Len(str) < Start - 3 Then Exit Function
    
        Dim temp_str  As String

        Dim temp_str2 As String

        Dim temp_str3 As String
    
        'Tomamos los últimos 3 bytes y convertimos sus valroes ASCII a hexa
102     temp_str = hex$(Asc(mid$(str, Start + 1, 1)))
104     temp_str2 = hex$(Asc(mid$(str, Start + 2, 1)))
106     temp_str3 = hex$(Asc(mid$(str, Start + 3, 1)))
    
        'Nos aseguramos todos midan 2 bytes (los ceros a la izquierda cuentan por ser bytes 2, 3 y 4)
108     While Len(temp_str) < 2

110         temp_str = "0" & temp_str
        Wend
    
112     While Len(temp_str2) < 2

114         temp_str2 = "0" & temp_str2
        Wend
    
116     While Len(temp_str3) < 2

118         temp_str3 = "0" & temp_str3
        Wend
    
        'Convertimos a una única cadena hexa
120     String_To_Long = val("&H" & hex$(Asc(mid$(str, Start, 1))) & temp_str & temp_str2 & temp_str3)
    
        'Si el cuarto byte era cero
122     If String_To_Long And &H80000000 Then String_To_Long = String_To_Long Xor &H80000001
    
        'Si el tercer byte era cero
124     If String_To_Long And &H40000000 Then String_To_Long = String_To_Long Xor &H40000100
    
        'Si el segundo byte era cero
126     If String_To_Long And &H20000000 Then String_To_Long = String_To_Long Xor &H20010000
    
        'Si el primer byte era cero
128     If String_To_Long And &H10000000 Then String_To_Long = String_To_Long Xor &H10000000
        
        Exit Function
        
ErrorHandler:

End Function

Public Function TieneObjEnInv(ByVal Userindex As Integer, ByVal ObjIndex As Integer, Optional ObjIndex2 As Integer = 0) As Boolean
        
        On Error GoTo TieneObjEnInv_Err
        

        'Devuelve el slot del inventario donde se encuentra el obj
        'Creaado por Ladder 25/09/2014
        Dim i As Byte

100     For i = 1 To 36

102         If UserList(Userindex).Invent.Object(i).ObjIndex = ObjIndex Then
104             TieneObjEnInv = True
                Exit Function

            End If

106         If ObjIndex2 > 0 Then
108             If UserList(Userindex).Invent.Object(i).ObjIndex = ObjIndex2 Then
110                 TieneObjEnInv = True
                    Exit Function

                End If

            End If

112     Next i

114     TieneObjEnInv = False

        
        Exit Function

TieneObjEnInv_Err:
116     Call RegistrarError(Err.Number, Err.description, "ModLadder.TieneObjEnInv", Erl)
118     Resume Next
        
End Function
Public Function CantidadObjEnInv(ByVal Userindex As Integer, ByVal ObjIndex As Integer) As Integer
        
        On Error GoTo CantidadObjEnInv_Err
        
        'Devuelve el amount si tiene el ObjIndex en el inventario, sino devuelve 0
        'Creaado por Ladder 25/09/2014
        Dim i As Byte

100     For i = 1 To 36

102         If UserList(Userindex).Invent.Object(i).ObjIndex = ObjIndex Then
104             CantidadObjEnInv = UserList(Userindex).Invent.Object(i).Amount
                Exit Function
            End If


106     Next i

108     CantidadObjEnInv = 0

        
        Exit Function

CantidadObjEnInv_Err:
110     Call RegistrarError(Err.Number, Err.description, "ModLadder.CantidadObjEnInv", Erl)
112     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "ModLadder.SumarTiempo", Erl)
116     Resume Next
        
End Function

Public Sub AgregarAConsola(ByVal Text As String)
        
        On Error GoTo AgregarAConsola_Err
        

100     frmMain.List1.AddItem (Text)

        
        Exit Sub

AgregarAConsola_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModLadder.AgregarAConsola", Erl)
104     Resume Next
        
End Sub

Function PuedeUsarObjeto(Userindex As Integer, ByVal ObjIndex As Integer) As Byte
        
        On Error GoTo PuedeUsarObjeto_Err
        

100     If UserList(Userindex).Stats.ELV < ObjData(ObjIndex).MinELV Then
102         PuedeUsarObjeto = 6
            Exit Function

        End If

104     Select Case ObjData(ObjIndex).OBJType

            Case otWeapon

106             If Not ClasePuedeUsarItem(Userindex, ObjIndex) Then
108                 PuedeUsarObjeto = 2
                    Exit Function

                End If
       
110         Case otNUDILLOS

112             If Not ClasePuedeUsarItem(Userindex, ObjIndex) Then
114                 PuedeUsarObjeto = 2
                    Exit Function

                End If
        
116         Case otArmadura
            
118             If Not CheckRazaUsaRopa(Userindex, ObjIndex) Then
120                 PuedeUsarObjeto = 5
                    Exit Function

                End If
                
122             If Not SexoPuedeUsarItem(Userindex, ObjIndex) Then
124                 PuedeUsarObjeto = 1
                    Exit Function

                End If

126             If Not ClasePuedeUsarItem(Userindex, ObjIndex) Then
128                 PuedeUsarObjeto = 2
                    Exit Function

                End If

130         Case otCASCO
            
132             If Not ClasePuedeUsarItem(Userindex, ObjIndex) Then
134                 PuedeUsarObjeto = 2
                    Exit Function

                End If
                
136         Case otESCUDO
            
138             If Not ClasePuedeUsarItem(Userindex, ObjIndex) Then
140                 PuedeUsarObjeto = 2
                    Exit Function

                End If
            
142         Case otPergaminos
            
144             If Not ClasePuedeUsarItem(Userindex, ObjIndex) Then
146                 PuedeUsarObjeto = 2
                    Exit Function

                End If

148         Case otMonturas

150             If Not CheckClaseTipo(Userindex, ObjIndex) Then
152                 PuedeUsarObjeto = 2
                    Exit Function

                End If
                
154             If Not CheckRazaTipo(Userindex, ObjIndex) Then
156                 PuedeUsarObjeto = 5
                    Exit Function

                End If
                
158         Case otHerramientas
160             If Not ClasePuedeUsarItem(Userindex, ObjIndex) Then
162                 PuedeUsarObjeto = 2
                    Exit Function

                End If
            
        End Select

164     PuedeUsarObjeto = 0

        
        Exit Function

PuedeUsarObjeto_Err:
166     Call RegistrarError(Err.Number, Err.description, "ModLadder.PuedeUsarObjeto", Erl)
168     Resume Next
        
End Function

Public Function RequiereOxigeno(ByVal UserMap) As Boolean
        
        On Error GoTo RequiereOxigeno_Err
        

100     Select Case UserMap

                'Case 55
                ' RequiereOxigeno = True
            Case 331
102             RequiereOxigeno = True

104         Case 332
106             RequiereOxigeno = True

108         Case 333
110             RequiereOxigeno = True

112         Case Else
114             RequiereOxigeno = False

        End Select

        
        Exit Function

RequiereOxigeno_Err:
116     Call RegistrarError(Err.Number, Err.description, "ModLadder.RequiereOxigeno", Erl)
118     Resume Next
        
End Function
