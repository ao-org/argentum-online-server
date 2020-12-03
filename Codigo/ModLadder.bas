Attribute VB_Name = "ModLadder"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

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

Function GetTimeFormated() As String
        
        On Error GoTo GetTimeFormated_Err
        
        Dim Elapsed As Single
        Elapsed = ((timeGetTime And &H7FFFFFFF) - HoraMundo) / DuracionDia
        
        Dim Mins As Long
        Mins = (Elapsed - Fix(Elapsed)) * 1440

        Dim Horita    As Byte

        Dim Minutitos As Byte

100     Horita = Fix(Mins / 60)
102     Minutitos = Mins Mod 60

104     GetTimeFormated = Right$("00" & Horita, 2) & ":" & Right$("00" & Minutitos, 2)

        
        Exit Function

GetTimeFormated_Err:
        Call RegistrarError(Err.Number, Err.description, "ModLadder.GetTimeFormated", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "ModLadder.GetHoraActual", Erl)
        Resume Next
        
End Sub

Public Function DarNameMapa(ByVal Map As Long) As String
        
        On Error GoTo DarNameMapa_Err
        
100     DarNameMapa = MapInfo(Map).map_name

        
        Exit Function

DarNameMapa_Err:
        Call RegistrarError(Err.Number, Err.description, "ModLadder.DarNameMapa", Erl)
        Resume Next
        
End Function

Public Sub CompletarAccionFin(ByVal UserIndex As Integer)
        
        On Error GoTo CompletarAccionFin_Err
        

        Dim obj  As ObjData

        Dim slot As Byte

100     Select Case UserList(UserIndex).Accion.TipoAccion

            Case Accion_Barra.Runa
102             obj = ObjData(UserList(UserIndex).Accion.RunaObj)
104             slot = UserList(UserIndex).Accion.ObjSlot

106             Select Case obj.TipoRuna

                    Case 1 'Cuando esta muerto lleva al lugar de Origen

                        Dim DeDonde As CityWorldPos

                        Dim Map     As Integer

                        Dim X       As Byte

                        Dim Y       As Byte
        
108                     If UserList(UserIndex).flags.Muerto = 0 Then

110                         Select Case UserList(Userindex).Hogar

                                Case eCiudad.cUllathorpe
112                                 DeDonde = Ullathorpe
                        
114                             Case eCiudad.cNix
116                                 DeDonde = Nix
            
118                             Case eCiudad.cBanderbill
120                                 DeDonde = Banderbill
                    
122                             Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
124                                 DeDonde = Lindos
                        
126                             Case eCiudad.cArghal
128                                 DeDonde = Arghal
                        
130                             Case eCiudad.CHillidan
132                                 DeDonde = Hillidan
                        
134                             Case Else
136                                 DeDonde = Ullathorpe

                            End Select

138                         Map = DeDonde.Map
140                         X = DeDonde.X
142                         Y = DeDonde.Y

                        Else

144                         If MapInfo(UserList(UserIndex).Pos.Map).extra2 <> 0 Then

146                             Select Case MapInfo(UserList(UserIndex).Pos.Map).extra2

                                    Case eCiudad.cUllathorpe
148                                     DeDonde = Ullathorpe
                        
150                                 Case eCiudad.cNix
152                                     DeDonde = Nix
            
154                                 Case eCiudad.cBanderbill
156                                     DeDonde = Banderbill
                    
158                                 Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
160                                     DeDonde = Lindos
                        
162                                 Case eCiudad.cArghal
164                                     DeDonde = Arghal
                        
166                                 Case eCiudad.CHillidan
168                                     DeDonde = Hillidan
                        
170                                 Case Else
172                                     DeDonde = Ullathorpe

                                End Select

                            Else

174                             Select Case UserList(UserIndex).Hogar

                                    Case eCiudad.cUllathorpe
176                                     DeDonde = Ullathorpe
                        
178                                 Case eCiudad.cNix
180                                     DeDonde = Nix
            
182                                 Case eCiudad.cBanderbill
184                                     DeDonde = Banderbill
                    
186                                 Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
188                                     DeDonde = Lindos
                        
190                                 Case eCiudad.cArghal
192                                     DeDonde = Arghal
                        
194                                 Case eCiudad.CHillidan
196                                     DeDonde = Hillidan
                        
198                                 Case Else
200                                     DeDonde = Ullathorpe

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

                        'Call WriteEfectToScreen(UserIndex, &HA4FFFF, 150, True)
216                     If UserList(UserIndex).flags.Navegando = 1 Then

                            Dim barca As ObjData

218                         barca = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
220                         Call DoNavega(UserIndex, barca, UserList(UserIndex).Invent.BarcoSlot)

                        End If
                
222                     If Resu Then
                
224                         If UserList(UserIndex).donador.activo = 0 Then ' Donador no espera tiempo
226                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 400, False))
228                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 400, Accion_Barra.Resucitar))
                            Else
230                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 10, False))
232                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 10, Accion_Barra.Resucitar))

                            End If
                
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
274                         Call QuitarUserInvItem(UserIndex, slot, 1)
276                         Call UpdateUserInv(False, UserIndex, slot)
                        Else

278                         If UserList(UserIndex).Pos.Map <> obj.DesdeMap Then
280                             Call WriteConsoleMsg(UserIndex, "Esta runa no puede ser usada desde aquí.", FontTypeNames.FONTTYPE_INFO)
                            Else
282                             Call QuitarUserInvItem(UserIndex, slot, 1)
284                             Call UpdateUserInv(False, UserIndex, slot)
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

304                 Case 3

                        Dim parejaindex As Integer
    
306                     If Not UserList(UserIndex).flags.BattleModo Then
                    
                            ' If UserList(UserIndex).donador.activo = 1 Then
308                         If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
310                             If UserList(UserIndex).flags.Casado = 1 Then
312                                 parejaindex = NameIndex(UserList(UserIndex).flags.Pareja)
                            
314                                 If parejaindex > 0 Then
316                                     If Not UserList(parejaindex).flags.BattleModo Then
318                                         Call WarpToLegalPos(UserIndex, UserList(parejaindex).Pos.Map, UserList(parejaindex).Pos.X, UserList(parejaindex).Pos.Y, True)
320                                         Call WriteConsoleMsg(UserIndex, "Te has teletransportado hacia tu pareja.", FontTypeNames.FONTTYPE_INFOIAO)
322                                         Call WriteConsoleMsg(parejaindex, "Tu pareja se ha teletransportado hacia vos.", FontTypeNames.FONTTYPE_INFOIAO)
                                        Else
324                                         Call WriteConsoleMsg(UserIndex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)

                                        End If
                                    
                                    Else
326                                     Call WriteConsoleMsg(UserIndex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)

                                    End If

                                Else
328                                 Call WriteConsoleMsg(UserIndex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)

                                End If

                            Else
330                             Call WriteConsoleMsg(UserIndex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If
                    
                            'Else
                            '   Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
                            '  End If
                        Else
332                         Call WriteConsoleMsg(UserIndex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
            
                        End If
            
                End Select

334         Case Accion_Barra.Intermundia
        
336             If UserList(UserIndex).flags.Muerto = 0 Then

                    Dim uh As Integer

                    Dim Mapaf, Xf, Yf As Integer

338                 uh = UserList(UserIndex).Accion.HechizoPendiente
    
340                 Mapaf = Hechizos(uh).TeleportXMap
342                 Xf = Hechizos(uh).TeleportXX
344                 Yf = Hechizos(uh).TeleportXY
    
346                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(uh).wav, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))  'Esta linea faltaba. Pablo (ToxicWaste)
348                 Call WriteConsoleMsg(UserIndex, "¡Has abierto la puerta a intermundia!", FontTypeNames.FONTTYPE_INFO)
350                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, -1, True))
352                 UserList(UserIndex).flags.Portal = 10
354                 UserList(UserIndex).flags.PortalMDestino = Mapaf
356                 UserList(UserIndex).flags.PortalYDestino = Xf
358                 UserList(UserIndex).flags.PortalXDestino = Yf
                
                    Dim Mapa As Integer

360                 Mapa = UserList(UserIndex).flags.PortalM
362                 X = UserList(UserIndex).flags.PortalX
364                 Y = UserList(UserIndex).flags.PortalY
366                 MapData(Mapa, X, Y).Particula = ParticulasIndex.TpVerde
368                 MapData(Mapa, X, Y).TimeParticula = -1
370                 MapData(Mapa, X, Y).TileExit.Map = UserList(UserIndex).flags.PortalMDestino
372                 MapData(Mapa, X, Y).TileExit.X = UserList(UserIndex).flags.PortalXDestino
374                 MapData(Mapa, X, Y).TileExit.Y = UserList(UserIndex).flags.PortalYDestino
                
                    'Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.Intermundia, -1))
376                 Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.TpVerde, -1))
                
378                 Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageLightFXToFloor(X, Y, &HFF80C0, 105))

                End If
                    
380             UserList(UserIndex).Accion.Particula = 0
382             UserList(UserIndex).Accion.TipoAccion = Accion_Barra.CancelarAccion
384             UserList(UserIndex).Accion.HechizoPendiente = 0
386             UserList(UserIndex).Accion.RunaObj = 0
388             UserList(UserIndex).Accion.ObjSlot = 0
390             UserList(UserIndex).Accion.AccionPendiente = False
            
                '
392         Case Accion_Barra.Resucitar
394             Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)
396             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 250, True))
398             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("204", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
400             Call RevivirUsuario(UserIndex)
                
402             UserList(UserIndex).Accion.Particula = 0
404             UserList(UserIndex).Accion.TipoAccion = Accion_Barra.CancelarAccion
406             UserList(UserIndex).Accion.HechizoPendiente = 0
408             UserList(UserIndex).Accion.RunaObj = 0
410             UserList(UserIndex).Accion.ObjSlot = 0
412             UserList(UserIndex).Accion.AccionPendiente = False
        
414         Case Accion_Barra.BattleModo
        
416             If UserList(UserIndex).flags.BattleModo = 1 Then
418                 Call Cerrar_Usuario(UserIndex)
                
                    ' Dim mapaa As Integer
                    '  Dim xa As Integer
                    ' Dim ya As Integer
                    ' mapaa = CInt(ReadField(1, GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Position"), 45))
                    ' xa = CInt(ReadField(2, GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Position"), 45))
                    ' ya = CInt(ReadField(3, GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Position"), 45))

                    ' Call WarpUserChar(UserIndex, mapaa, xa, ya, False)
                
                    ' Call RelogearUser(UserIndex, UserList(UserIndex).name, UserList(UserIndex).cuenta)
                Else
                
420                 If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then

422                     UserList(UserIndex).flags.Oculto = 0
424                     UserList(UserIndex).flags.invisible = 0
426                     UserList(UserIndex).Counters.TiempoOculto = 0
428                     UserList(UserIndex).Counters.Invisibilidad = 0
                
430                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))

                    End If
                
432                 Call SaveUser(UserIndex)  'Guardo el PJ

434                 X = 50
436                 Y = 54
438                 Call FindLegalPos(UserIndex, 336, X, Y)
                
440                 Call WarpUserChar(UserIndex, 336, X, Y, True)
                
442                 UserList(UserIndex).flags.BattleModo = 1

444                 If UserList(UserIndex).flags.Subastando Then
446                     Call CancelarSubasta

                    End If
                
448                 Call AumentarPJ(UserIndex)
450                 Call WriteConsoleMsg(UserIndex, "Battle> Ahora tu personaje se encuentra en modo batalla. Recuerda que todos los cambios que se realicen sobre éste no tendran efecto mientras te encuentres aquí. Cuando desees salir, solamente toca ESC o escribe /SALIR y relogea con tu personaje.", FontTypeNames.FONTTYPE_CITIZEN)
                
                End If

452             UserList(UserIndex).Accion.AccionPendiente = False
454             UserList(UserIndex).Accion.Particula = 0
456             UserList(UserIndex).Accion.TipoAccion = Accion_Barra.CancelarAccion
458             UserList(UserIndex).Accion.HechizoPendiente = 0
460             UserList(UserIndex).Accion.RunaObj = 0
462             UserList(UserIndex).Accion.ObjSlot = 0
                
464         Case Accion_Barra.GoToPareja
    
466             If Not UserList(UserIndex).flags.BattleModo Then
                    
                    ' If UserList(UserIndex).donador.activo = 1 Then
468                 If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
470                     If UserList(UserIndex).flags.Casado = 1 Then
472                         parejaindex = NameIndex(UserList(UserIndex).flags.Pareja)
                            
474                         If parejaindex > 0 Then
476                             If Not UserList(parejaindex).flags.BattleModo Then
478                                 Call WarpToLegalPos(UserIndex, UserList(parejaindex).Pos.Map, UserList(parejaindex).Pos.X, UserList(parejaindex).Pos.Y, True)
480                                 Call WriteConsoleMsg(UserIndex, "Te has teletransportado hacia tu pareja.", FontTypeNames.FONTTYPE_INFOIAO)
482                                 Call WriteConsoleMsg(parejaindex, "Tu pareja se ha teletransportado hacia vos.", FontTypeNames.FONTTYPE_INFOIAO)
                                Else
484                                 Call WriteConsoleMsg(UserIndex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)

                                End If
                                    
                            Else
486                             Call WriteConsoleMsg(UserIndex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If

                        Else
488                         Call WriteConsoleMsg(UserIndex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

                    Else
490                     Call WriteConsoleMsg(UserIndex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If
                    
                    ' Else
                    ' Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
                    'End If
                Else
492                 Call WriteConsoleMsg(UserIndex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
            
                End If
            
        End Select
       
        
        Exit Sub

CompletarAccionFin_Err:
        Call RegistrarError(Err.Number, Err.description, "ModLadder.CompletarAccionFin", Erl)
        Resume Next
        
End Sub

Public Function General_Get_Line_Count(ByVal FileName As String) As Long

    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2005
    '
    '**************************************************************
    On Error GoTo ErrorHandler

    Dim n As Integer, tmpStr As String

    If LenB(FileName) Then
        n = FreeFile()
    
        Open FileName For Input As #n
    
        Do While Not EOF(n)
            General_Get_Line_Count = General_Get_Line_Count + 1
            Line Input #n, tmpStr
        Loop
    
        Close n

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
        Call RegistrarError(Err.Number, Err.description, "ModLadder.Integer_To_String", Erl)
        Resume Next
        
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
    If Len(str) < Start - 1 Or Len(str) = 0 Then Exit Function
    
    'Convertimos a hexa el valor ascii del segundo byte
    temp_str = hex$(Asc(mid$(str, Start + 1, 1)))
    
    'Nos aseguramos tenga 2 bytes (los ceros a la izquierda cuentan por ser el segundo byte)
    While Len(temp_str) < 2

        temp_str = "0" & temp_str
    Wend
    
    'Convertimos a integer
    String_To_Integer = val("&H" & hex$(Asc(mid$(str, Start, 1))) & temp_str)
            
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
        Call RegistrarError(Err.Number, Err.description, "ModLadder.Byte_To_String", Erl)
        Resume Next
        
End Function

Public Function String_To_Byte(ByRef str As String, ByVal Start As Integer) As Byte

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 3/12/2005
    '
    '**************************************************************
    On Error GoTo Error_Handler
    
    If Len(str) < Start Then Exit Function
    
    String_To_Byte = Asc(mid$(str, Start, 1))
    
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
        Call RegistrarError(Err.Number, Err.description, "ModLadder.Long_To_String", Erl)
        Resume Next
        
End Function

Public Function String_To_Long(ByRef str As String, ByVal Start As Integer) As Long
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 3/12/2005
    '
    '**************************************************************
    
    On Error GoTo ErrorHandler
        
    If Len(str) < Start - 3 Then Exit Function
    
    Dim temp_str  As String

    Dim temp_str2 As String

    Dim temp_str3 As String
    
    'Tomamos los últimos 3 bytes y convertimos sus valroes ASCII a hexa
    temp_str = hex$(Asc(mid$(str, Start + 1, 1)))
    temp_str2 = hex$(Asc(mid$(str, Start + 2, 1)))
    temp_str3 = hex$(Asc(mid$(str, Start + 3, 1)))
    
    'Nos aseguramos todos midan 2 bytes (los ceros a la izquierda cuentan por ser bytes 2, 3 y 4)
    While Len(temp_str) < 2

        temp_str = "0" & temp_str
    Wend
    
    While Len(temp_str2) < 2

        temp_str2 = "0" & temp_str2
    Wend
    
    While Len(temp_str3) < 2

        temp_str3 = "0" & temp_str3
    Wend
    
    'Convertimos a una única cadena hexa
    String_To_Long = val("&H" & hex$(Asc(mid$(str, Start, 1))) & temp_str & temp_str2 & temp_str3)
    
    'Si el cuarto byte era cero
    If String_To_Long And &H80000000 Then String_To_Long = String_To_Long Xor &H80000001
    
    'Si el tercer byte era cero
    If String_To_Long And &H40000000 Then String_To_Long = String_To_Long Xor &H40000100
    
    'Si el segundo byte era cero
    If String_To_Long And &H20000000 Then String_To_Long = String_To_Long Xor &H20010000
    
    'Si el primer byte era cero
    If String_To_Long And &H10000000 Then String_To_Long = String_To_Long Xor &H10000000
        
    Exit Function
        
ErrorHandler:

End Function

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
        Call RegistrarError(Err.Number, Err.description, "ModLadder.TieneObjEnInv", Erl)
        Resume Next
        
End Function
Public Function CantidadObjEnInv(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
        
        On Error GoTo CantidadObjEnInv_Err
        
        'Devuelve el amount si tiene el ObjIndex en el inventario, sino devuelve 0
        'Creaado por Ladder 25/09/2014
        Dim i As Byte

100     For i = 1 To 36

102         If UserList(UserIndex).Invent.Object(i).ObjIndex = ObjIndex Then
104             CantidadObjEnInv = UserList(UserIndex).Invent.Object(i).Amount
                Exit Function
            End If


112     Next i

114     CantidadObjEnInv = 0

        
        Exit Function

CantidadObjEnInv_Err:
        Call RegistrarError(Err.Number, Err.description, "ModLadder.CantidadObjEnInv", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "ModLadder.SumarTiempo", Erl)
        Resume Next
        
End Function

Public Sub AgregarAConsola(ByVal Text As String)
        
        On Error GoTo AgregarAConsola_Err
        

100     frmMain.List1.AddItem (Text)

        
        Exit Sub

AgregarAConsola_Err:
        Call RegistrarError(Err.Number, Err.description, "ModLadder.AgregarAConsola", Erl)
        Resume Next
        
End Sub

Function PuedeUsarObjeto(UserIndex As Integer, ByVal ObjIndex As Integer) As Byte
        
        On Error GoTo PuedeUsarObjeto_Err
        

100     If UserList(UserIndex).Stats.ELV < ObjData(ObjIndex).MinELV Then
102         PuedeUsarObjeto = 6
            Exit Function

        End If

104     Select Case ObjData(ObjIndex).OBJType

            Case otWeapon

106             If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
108                 PuedeUsarObjeto = 2
                    Exit Function

                End If
       
110         Case otNUDILLOS

112             If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
114                 PuedeUsarObjeto = 2
                    Exit Function

                End If
        
116         Case otArmadura
            
118             If Not CheckRazaUsaRopa(UserIndex, ObjIndex) Then
120                 PuedeUsarObjeto = 5
                    Exit Function

                End If
                
122             If Not SexoPuedeUsarItem(UserIndex, ObjIndex) Then
124                 PuedeUsarObjeto = 1
                    Exit Function

                End If

126             If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
128                 PuedeUsarObjeto = 2
                    Exit Function

                End If

130         Case otCASCO
            
132             If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
134                 PuedeUsarObjeto = 2
                    Exit Function

                End If
                
136         Case otESCUDO
            
138             If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
140                 PuedeUsarObjeto = 2
                    Exit Function

                End If
            
142         Case otPergaminos
            
144             If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
146                 PuedeUsarObjeto = 2
                    Exit Function

                End If

148         Case otMonturas

150             If Not CheckClaseTipo(UserIndex, ObjIndex) Then
152                 PuedeUsarObjeto = 2
                    Exit Function

                End If
                
154             If Not CheckRazaTipo(UserIndex, ObjIndex) Then
156                 PuedeUsarObjeto = 5
                    Exit Function

                End If
                
            Case otHerramientas
                If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                    PuedeUsarObjeto = 2
                    Exit Function

                End If
            
        End Select

158     PuedeUsarObjeto = 0

        
        Exit Function

PuedeUsarObjeto_Err:
        Call RegistrarError(Err.Number, Err.description, "ModLadder.PuedeUsarObjeto", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "ModLadder.RequiereOxigeno", Erl)
        Resume Next
        
End Function
