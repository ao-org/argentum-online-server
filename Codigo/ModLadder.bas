Attribute VB_Name = "ModLadder"
Option Explicit
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private theTime As SYSTEMTIME
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


Function GetTimeFormated(Mins As Integer) As String
Dim Horita As Byte
Dim Minutitos As Byte
Dim a As String

Horita = Fix(Mins / 60)
Minutitos = Mins - 60 * Horita
If Minutitos < 10 Then
    GetTimeFormated = Horita & ":0" & Minutitos
Else
    GetTimeFormated = Horita & ":" & Minutitos
End If
End Function
Public Sub GetHoraActual()
GetSystemTime theTime

HoraActual = (theTime.wHour - 3)
If HoraActual = -3 Then HoraActual = 21
If HoraActual = -2 Then HoraActual = 22
If HoraActual = -1 Then HoraActual = 23
frmMain.lblhora.Caption = HoraActual & ":" & Format(theTime.wMinute, "00") & ":" & Format(theTime.wSecond, "00")
HoraEvento = HoraActual
End Sub

#If Lac Then
Public Sub LoadAntiCheat()
Dim i As Integer

Lac_Camina = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Caminar")))
Lac_Lanzar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Lanzar")))
Lac_Usar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Usar")))
Lac_Tirar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Tirar")))
Lac_Pociones = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Pociones")))
Lac_Pegar = CLng(val(GetVar$(App.Path & "\AntiCheats.ini", "INTERVALOS", "Pegar")))
For i = 1 To MaxUsers
ResetearLac i
Next


End Sub
Public Sub ResetearLac(UserIndex As Integer)
With UserList(UserIndex).Lac
.LCaminar.Init Lac_Camina
.LPociones.Init Lac_Pociones
.LUsar.Init Lac_Usar
.LPegar.Init Lac_Pegar
.LLanzar.Init Lac_Lanzar
.LTirar.Init Lac_Tirar
End With

End Sub
Public Sub CargaLac(UserIndex As Integer)
With UserList(UserIndex).Lac
Set .LCaminar = New Cls_InterGTC
Set .LLanzar = New Cls_InterGTC
Set .LPegar = New Cls_InterGTC
Set .LPociones = New Cls_InterGTC
Set .LTirar = New Cls_InterGTC
Set .LUsar = New Cls_InterGTC

.LCaminar.Init Lac_Camina
.LPociones.Init Lac_Pociones
.LUsar.Init Lac_Usar
.LPegar.Init Lac_Pegar
.LLanzar.Init Lac_Lanzar
.LTirar.Init Lac_Tirar
End With

End Sub
Public Sub DescargaLac(UserIndex As Integer)
'Exit Sub
With UserList(UserIndex).Lac
Set .LCaminar = Nothing
Set .LLanzar = Nothing
Set .LPegar = Nothing
Set .LPociones = Nothing
Set .LTirar = Nothing
Set .LUsar = Nothing
End With
End Sub
#End If


Public Function DarNameMapa(ByVal Map As Long) As String
DarNameMapa = MapInfo(Map).map_name
End Function

Public Sub CompletarAccionFin(ByVal UserIndex As Integer)
Dim obj As ObjData
Dim slot As Byte

Select Case UserList(UserIndex).accion.TipoAccion

    Case Accion_Barra.Runa
        obj = ObjData(UserList(UserIndex).accion.RunaObj)
        slot = UserList(UserIndex).accion.ObjSlot
     Select Case obj.TipoRuna
            Case 1 'Cuando esta muerto lleva al lugar de Origen
                Dim DeDonde As CityWorldPos
                Dim Map As Integer
                Dim x As Byte
                Dim Y As Byte
                
        
        If UserList(UserIndex).flags.Muerto = 0 Then
        

            Select Case UserList(UserIndex).Hogar
                    Case eCiudad.cUllathorpe
                        DeDonde = CityUllathorpe
                        
                    Case eCiudad.cNix
                        DeDonde = CityNix
            
                    Case eCiudad.cBanderbill
                        DeDonde = CityBanderbill
                    
                    Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                        DeDonde = CityLindos
                        
                    Case eCiudad.cArghal
                        DeDonde = CityArghal
                        
                    Case eCiudad.CHillidan
                        DeDonde = CityHillidan
                        
                    Case Else
                        DeDonde = CityUllathorpe
                End Select
                    Map = DeDonde.Map
                    x = DeDonde.x
                    Y = DeDonde.Y
        Else
            If MapInfo(UserList(UserIndex).Pos.Map).extra2 <> 0 Then
                Select Case MapInfo(UserList(UserIndex).Pos.Map).extra2
                    Case eCiudad.cUllathorpe
                        DeDonde = CityUllathorpe
                        
                    Case eCiudad.cNix
                        DeDonde = CityNix
            
                    Case eCiudad.cBanderbill
                        DeDonde = CityBanderbill
                    
                    Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                        DeDonde = CityLindos
                        
                    Case eCiudad.cArghal
                        DeDonde = CityArghal
                        
                    Case eCiudad.CHillidan
                        DeDonde = CityHillidan
                        
                    Case Else
                        DeDonde = CityUllathorpe
                End Select
            Else
                Select Case UserList(UserIndex).Hogar
                    Case eCiudad.cUllathorpe
                        DeDonde = CityUllathorpe
                        
                    Case eCiudad.cNix
                        DeDonde = CityNix
            
                    Case eCiudad.cBanderbill
                        DeDonde = CityBanderbill
                    
                    Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                        DeDonde = CityLindos
                        
                    Case eCiudad.cArghal
                        DeDonde = CityArghal
                        
                    Case eCiudad.CHillidan
                        DeDonde = CityHillidan
                        
                    Case Else
                        DeDonde = CityUllathorpe
                End Select
            End If
                 
                
                Map = DeDonde.MapaResu
                x = DeDonde.ResuX
                Y = DeDonde.ResuY
                
                
                
                Dim Resu As Boolean
                
                Resu = True
            
        End If
        
                
                Call FindLegalPos(UserIndex, Map, x, Y)
                Call WarpUserChar(UserIndex, Map, x, Y, True)
                Call WriteConsoleMsg(UserIndex, "Has regresado a tu ciudad de origen.", FontTypeNames.FONTTYPE_WARNING)
                'Call WriteEfectToScreen(UserIndex, &HA4FFFF, 150, True)
                If UserList(UserIndex).flags.Navegando = 1 Then
                Dim barca As ObjData
                barca = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
                Call DoNavega(UserIndex, barca, UserList(UserIndex).Invent.BarcoSlot)
                End If
                
                If Resu Then
                
                If UserList(UserIndex).donador.activo = 0 Then ' Donador no espera tiempo
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 400, False))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 400, Accion_Barra.Resucitar))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 10, False))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 10, Accion_Barra.Resucitar))
                End If
                
                UserList(UserIndex).accion.AccionPendiente = True
                UserList(UserIndex).accion.Particula = ParticulasIndex.Resucitar
                UserList(UserIndex).accion.TipoAccion = Accion_Barra.Resucitar

                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, "82", FontTypeNames.FONTTYPE_INFOIAO)
                End If
                
                If Not Resu Then
                UserList(UserIndex).accion.AccionPendiente = False
                UserList(UserIndex).accion.Particula = 0
                UserList(UserIndex).accion.TipoAccion = Accion_Barra.CancelarAccion
                End If
                UserList(UserIndex).accion.HechizoPendiente = 0
                UserList(UserIndex).accion.RunaObj = 0
                UserList(UserIndex).accion.ObjSlot = 0
        
              
            Case 2
                Map = obj.HastaMap
                x = obj.HastaX
                Y = obj.HastaY
            
                If obj.DesdeMap = 0 Then
                    Call FindLegalPos(UserIndex, Map, x, Y)
                    Call WarpUserChar(UserIndex, Map, x, Y, True)
                    Call WriteConsoleMsg(UserIndex, "Te has teletransportado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    Call UpdateUserInv(False, UserIndex, slot)
                Else
                    If UserList(UserIndex).Pos.Map <> obj.DesdeMap Then
                        Call WriteConsoleMsg(UserIndex, "Esta runa no puede ser usada desde aquí.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                        Call UpdateUserInv(False, UserIndex, slot)
                        Call FindLegalPos(UserIndex, Map, x, Y)
                        Call WarpUserChar(UserIndex, Map, x, Y, True)
                        Call WriteConsoleMsg(UserIndex, "Te has teletransportado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
                    End If
                End If
                
                
        
                UserList(UserIndex).accion.Particula = 0
                UserList(UserIndex).accion.TipoAccion = Accion_Barra.CancelarAccion
                UserList(UserIndex).accion.HechizoPendiente = 0
                UserList(UserIndex).accion.RunaObj = 0
                UserList(UserIndex).accion.ObjSlot = 0
                UserList(UserIndex).accion.AccionPendiente = False
            Case 3
                            Dim parejaindex As Integer
    
    
            If Not UserList(UserIndex).flags.BattleModo Then
                    
              ' If UserList(UserIndex).donador.activo = 1 Then
                    If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                        If UserList(UserIndex).flags.Casado = 1 Then
                            parejaindex = NameIndex(UserList(UserIndex).flags.Pareja)
                            
                                If parejaindex > 0 Then
                                    If Not UserList(parejaindex).flags.BattleModo Then
                                        Call WarpToLegalPos(UserIndex, UserList(parejaindex).Pos.Map, UserList(parejaindex).Pos.x, UserList(parejaindex).Pos.Y, True)
                                        Call WriteConsoleMsg(UserIndex, "Te has teletransportado hacia tu pareja.", FontTypeNames.FONTTYPE_INFOIAO)
                                        Call WriteConsoleMsg(parejaindex, "Tu pareja se ha teletransportado hacia vos.", FontTypeNames.FONTTYPE_INFOIAO)
                                    Else
                                        Call WriteConsoleMsg(UserIndex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)
                                    End If
                                    
                                Else
                                    Call WriteConsoleMsg(UserIndex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)
                                End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                    
                'Else
                 '   Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
              '  End If
            Else
                Call WriteConsoleMsg(UserIndex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
            
            End If
            
            
            
            
            End Select
        Case Accion_Barra.Intermundia
        
            If UserList(UserIndex).flags.Muerto = 0 Then
                Dim uh As Integer
                 Dim Mapaf, Xf, Yf As Integer
                uh = UserList(UserIndex).accion.HechizoPendiente
    
                Mapaf = Hechizos(uh).TeleportXMap
                Xf = Hechizos(uh).TeleportXX
                Yf = Hechizos(uh).TeleportXY
    
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(uh).wav, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))  'Esta linea faltaba. Pablo (ToxicWaste)
                Call WriteConsoleMsg(UserIndex, "¡Has abierto la puerta a intermundia!", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, -1, True))
                UserList(UserIndex).flags.Portal = 10
                UserList(UserIndex).flags.PortalMDestino = Mapaf
                UserList(UserIndex).flags.PortalYDestino = Xf
                UserList(UserIndex).flags.PortalXDestino = Yf
                
                
                Dim Mapa As Integer
                Mapa = UserList(UserIndex).flags.PortalM
                x = UserList(UserIndex).flags.PortalX
                Y = UserList(UserIndex).flags.PortalY
                MapData(Mapa, x, Y).Particula = ParticulasIndex.TpVerde
                MapData(Mapa, x, Y).TimeParticula = -1
                MapData(Mapa, x, Y).TileExit.Map = UserList(UserIndex).flags.PortalMDestino
                MapData(Mapa, x, Y).TileExit.x = UserList(UserIndex).flags.PortalXDestino
                MapData(Mapa, x, Y).TileExit.Y = UserList(UserIndex).flags.PortalYDestino
                
                'Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.Intermundia, -1))
                Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(x, Y, ParticulasIndex.TpVerde, -1))
                
                Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageLightFXToFloor(x, Y, &HFF80C0, 105))
               End If
                    
                UserList(UserIndex).accion.Particula = 0
                UserList(UserIndex).accion.TipoAccion = Accion_Barra.CancelarAccion
                UserList(UserIndex).accion.HechizoPendiente = 0
                UserList(UserIndex).accion.RunaObj = 0
                UserList(UserIndex).accion.ObjSlot = 0
                UserList(UserIndex).accion.AccionPendiente = False
            
            
            
            
            '
        Case Accion_Barra.Resucitar
                Call WriteConsoleMsg(UserIndex, "¡Has sido resucitado!", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 250, True))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("204", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                Call RevivirUsuario(UserIndex)
                
                
                UserList(UserIndex).accion.Particula = 0
                UserList(UserIndex).accion.TipoAccion = Accion_Barra.CancelarAccion
                UserList(UserIndex).accion.HechizoPendiente = 0
                UserList(UserIndex).accion.RunaObj = 0
                UserList(UserIndex).accion.ObjSlot = 0
                UserList(UserIndex).accion.AccionPendiente = False
        
        Case Accion_Barra.BattleModo
        
                If UserList(UserIndex).flags.BattleModo = 1 Then
                    Call Cerrar_Usuario(UserIndex)
                
               ' Dim mapaa As Integer
              '  Dim xa As Integer
               ' Dim ya As Integer
               ' mapaa = CInt(ReadField(1, GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Position"), 45))
               ' xa = CInt(ReadField(2, GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Position"), 45))
               ' ya = CInt(ReadField(3, GetVar(CharPath & UserList(UserIndex).name & ".chr", "INIT", "Position"), 45))

               ' Call WarpUserChar(UserIndex, mapaa, xa, ya, False)
                
                
               ' Call RelogearUser(UserIndex, UserList(UserIndex).name, UserList(UserIndex).cuenta)
                Else
                
                If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then

                        UserList(UserIndex).flags.Oculto = 0
                        UserList(UserIndex).flags.invisible = 0
                        UserList(UserIndex).Counters.TiempoOculto = 0
                        UserList(UserIndex).Counters.Invisibilidad = 0
                
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
                End If
                
                Call SaveUser(UserIndex)  'Guardo el PJ
                

                    x = 50
                    Y = 54
                    Call FindLegalPos(UserIndex, 336, x, Y)
                
                    Call WarpUserChar(UserIndex, 336, x, Y, True)
                
                UserList(UserIndex).flags.BattleModo = 1
                If UserList(UserIndex).flags.Subastando Then
                    Call CancelarSubasta
                End If
                
                
                
                Call AumentarPJ(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Battle> Ahora tu personaje se encuentra en modo batalla. Recuerda que todos los cambios que se realicen sobre éste no tendran efecto mientras te encuentres aquí. Cuando desees salir, solamente toca ESC o escribe /SALIR y relogea con tu personaje.", FontTypeNames.FONTTYPE_CITIZEN)
                
                End If
                UserList(UserIndex).accion.AccionPendiente = False
                UserList(UserIndex).accion.Particula = 0
                UserList(UserIndex).accion.TipoAccion = Accion_Barra.CancelarAccion
                UserList(UserIndex).accion.HechizoPendiente = 0
                UserList(UserIndex).accion.RunaObj = 0
                UserList(UserIndex).accion.ObjSlot = 0
                
        Case Accion_Barra.GoToPareja
            
    
    
            If Not UserList(UserIndex).flags.BattleModo Then
                    
               ' If UserList(UserIndex).donador.activo = 1 Then
                    If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                        If UserList(UserIndex).flags.Casado = 1 Then
                            parejaindex = NameIndex(UserList(UserIndex).flags.Pareja)
                            
                                If parejaindex > 0 Then
                                    If Not UserList(parejaindex).flags.BattleModo Then
                                        Call WarpToLegalPos(UserIndex, UserList(parejaindex).Pos.Map, UserList(parejaindex).Pos.x, UserList(parejaindex).Pos.Y, True)
                                        Call WriteConsoleMsg(UserIndex, "Te has teletransportado hacia tu pareja.", FontTypeNames.FONTTYPE_INFOIAO)
                                        Call WriteConsoleMsg(parejaindex, "Tu pareja se ha teletransportado hacia vos.", FontTypeNames.FONTTYPE_INFOIAO)
                                    Else
                                        Call WriteConsoleMsg(UserIndex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)
                                    End If
                                    
                                Else
                                    Call WriteConsoleMsg(UserIndex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)
                                End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                    
               ' Else
                   ' Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
                'End If
            Else
                Call WriteConsoleMsg(UserIndex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
            
            End If
            
            
End Select

                


       
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
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    Dim temp As String
        
    'Convertimos a hexa
    temp = hex$(Var)
    
    'Nos aseguramos tenga 4 bytes de largo
    While Len(temp) < 4
        temp = "0" & temp
    Wend
    
    'Convertimos a string
    Integer_To_String = Chr$(val("&H" & Left$(temp, 2))) & Chr$(val("&H" & Right$(temp, 2)))
Exit Function

ErrorHandler:
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
    Byte_To_String = Chr$(val("&H" & hex$(Var)))
Exit Function

ErrorHandler:
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
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    'No aceptamos valores que usen los 4 últimos its
    If Var > &HFFFFFFF Then GoTo ErrorHandler
    
    Dim temp As String
    
    'Vemos si el cuarto byte es cero
    If (Var And &HFF&) = 0 Then _
        Var = Var Or &H80000001
    
    'Vemos si el tercer byte es cero
    If (Var And &HFF00&) = 0 Then _
        Var = Var Or &H40000100
    
    'Vemos si el segundo byte es cero
    If (Var And &HFF0000) = 0 Then _
        Var = Var Or &H20010000
    
    'Vemos si el primer byte es cero
    If Var < &H1000000 Then _
        Var = Var Or &H10000000
    
    'Convertimos a hexa
    temp = hex$(Var)
    
    'Nos aseguramos tenga 8 bytes de largo
    While Len(temp) < 8
        temp = "0" & temp
    Wend
    
    'Convertimos a string
    Long_To_String = Chr$(val("&H" & Left$(temp, 2))) & Chr$(val("&H" & mid$(temp, 3, 2))) & Chr$(val("&H" & mid$(temp, 5, 2))) & Chr$(val("&H" & mid$(temp, 7, 2)))
Exit Function

ErrorHandler:
End Function

Public Function String_To_Long(ByRef str As String, ByVal Start As Integer) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/12/2005
'
'**************************************************************
    
    On Error GoTo ErrorHandler
        
    If Len(str) < Start - 3 Then Exit Function
    
    Dim temp_str As String
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
    If String_To_Long And &H80000000 Then _
        String_To_Long = String_To_Long Xor &H80000001
    
    'Si el tercer byte era cero
    If String_To_Long And &H40000000 Then _
        String_To_Long = String_To_Long Xor &H40000100
    
    'Si el segundo byte era cero
    If String_To_Long And &H20000000 Then _
        String_To_Long = String_To_Long Xor &H20010000
    
    'Si el primer byte era cero
    If String_To_Long And &H10000000 Then _
        String_To_Long = String_To_Long Xor &H10000000
        
    Exit Function
        
ErrorHandler:

End Function


Public Function TieneObjEnInv(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ObjIndex2 As Integer = 0) As Boolean
'Devuelve el slot del inventario donde se encuentra el obj
'Creaado por Ladder 25/09/2014
Dim i As Byte

For i = 1 To 36
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ObjIndex Then
        TieneObjEnInv = True
        Exit Function
    End If
    If ObjIndex2 > 0 Then
        If UserList(UserIndex).Invent.Object(i).ObjIndex = ObjIndex2 Then
            TieneObjEnInv = True
            Exit Function
    End If
    End If
Next i


TieneObjEnInv = False

End Function
Public Function SumarTiempo(segundos As Integer) As String
Dim a As Variant, b As Variant
Dim x As Integer
Dim T As String

T = "00:00:00" 'Lo inicializamos en 0 horas, 0 minutos, 0 segundos
a = Format("00:00:01", "hh:mm:ss") 'guardamos en una variable el formato de 1 segundos

For x = 1 To segundos 'hacemos segundo a segundo
    b = Format(T, "hh:mm:ss") 'En B guardamos un formato de hora:minuto:segundo segun lo que tenia T
    T = Format(TimeValue(a) + TimeValue(b), "hh:mm:ss") 'asignamos a T la suma de A + B (osea, sumamos logicamente 1 segundo)
Next x

SumarTiempo = T 'a la funcion le damos el valor que hallamos en T
End Function
Public Sub AgregarAConsola(ByVal Text As String)


frmMain.List1.AddItem (Text)
End Sub

Function PuedeUsarObjeto(UserIndex As Integer, ByVal ObjIndex As Integer) As Byte


If UserList(UserIndex).Stats.ELV < ObjData(ObjIndex).MinELV Then
    PuedeUsarObjeto = 6
    Exit Function
End If


Select Case ObjData(ObjIndex).OBJType
    Case otWeapon

            If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                 PuedeUsarObjeto = 2
                 Exit Function
            End If
       
    Case otNUDILLOS

        If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
             PuedeUsarObjeto = 2
             Exit Function
        End If
        
    Case otArmadura
         

            
                If Not CheckRazaUsaRopa(UserIndex, ObjIndex) Then
                    PuedeUsarObjeto = 5
                    Exit Function
                End If
                
                If Not SexoPuedeUsarItem(UserIndex, ObjIndex) Then
                    PuedeUsarObjeto = 1
                    Exit Function
                End If
                 

                    If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                         PuedeUsarObjeto = 2
                         Exit Function
                    End If
                


            Case otCASCO
            
                 If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                      PuedeUsarObjeto = 2
                      Exit Function
                 End If
                

                
            Case otESCUDO
            
                If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                    PuedeUsarObjeto = 2
                    Exit Function
                End If

            
            Case otPergaminos
            
            
             If Not ClasePuedeUsarItem(UserIndex, ObjIndex) Then
                      PuedeUsarObjeto = 2
                      Exit Function
                 End If
            

        Case otMonturas
                If Not CheckClaseTipo(UserIndex, ObjIndex) Then
                      PuedeUsarObjeto = 2
                      Exit Function
                 End If
                
                If Not CheckRazaTipo(UserIndex, ObjIndex) Then
                      PuedeUsarObjeto = 5
                      Exit Function
                 End If
            
            
            
End Select

PuedeUsarObjeto = 0

End Function

Public Function RequiereOxigeno(ByVal UserMap) As Boolean


Select Case UserMap
    'Case 55
       ' RequiereOxigeno = True
    Case 331
        RequiereOxigeno = True
    Case 332
        RequiereOxigeno = True
    Case 333
        RequiereOxigeno = True
    Case Else
        RequiereOxigeno = False
End Select


End Function
