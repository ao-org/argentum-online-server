Attribute VB_Name = "Hogar"
Option Explicit

'Cantidad de Ciudades
Public Const NUMCIUDADES    As Byte = 6

Type HomeDistance
    distanceToCity(1 To NUMCIUDADES) As Integer
End Type

Public Ciudades(1 To NUMCIUDADES)         As WorldPos
Public distanceToCities()                 As HomeDistance

Public Const MATRIX_INITIAL_MAP           As Integer = 1
Public Const GOHOME_PENALTY               As Integer = 5

Public Function getLimit(ByVal Mapa As Integer, ByVal side As Byte) As Integer

        '***************************************************
        'Author: Budi
        'Last Modification: 31/01/2010
        'Retrieves the limit in the given side in the given map.
        'TODO: This should be set in the .inf map file.
        '***************************************************
        Dim X As Long

        Dim Y As Long

100     If Mapa <= 0 Then Exit Function

102     For X = 15 To 87
104         For Y = 0 To 3

106             Select Case side

                    Case eHeading.NORTH
108                     getLimit = MapData(Mapa, X, 7 + Y).TileExit.Map

110                 Case eHeading.EAST
112                     getLimit = MapData(Mapa, 92 - Y, X).TileExit.Map

114                 Case eHeading.SOUTH
116                     getLimit = MapData(Mapa, X, 94 - Y).TileExit.Map

118                 Case eHeading.WEST
120                     getLimit = MapData(Mapa, 9 + Y, X).TileExit.Map

                End Select

122             If getLimit > 0 Then Exit Function
124         Next Y
126     Next X

End Function

Public Sub generateMatrix(ByVal Mapa As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim i As Integer

        Dim j As Integer
    
100     ReDim distanceToCities(1 To NumMaps) As HomeDistance
    
102     For j = 1 To NUMCIUDADES
104         For i = 1 To NumMaps
106             distanceToCities(i).distanceToCity(j) = -1
108         Next i
110     Next j
    
112     For j = 1 To NUMCIUDADES
114         For i = 1 To 4

116             Select Case i

                    Case eHeading.NORTH
118                     Call setDistance(getLimit(Ciudades(j).Map, eHeading.NORTH), j, i, 0, -1)

120                 Case eHeading.EAST
122                     Call setDistance(getLimit(Ciudades(j).Map, eHeading.EAST), j, i, 1, 0)

124                 Case eHeading.SOUTH
126                     Call setDistance(getLimit(Ciudades(j).Map, eHeading.SOUTH), j, i, 0, 1)

128                 Case eHeading.WEST
130                     Call setDistance(getLimit(Ciudades(j).Map, eHeading.WEST), j, i, -1, 0)

                End Select

132         Next i
134     Next j

End Sub

Public Sub setDistance(ByVal Mapa As Integer, _
                       ByVal city As Byte, _
                       ByVal side As Integer, _
                       Optional ByVal X As Integer = 0, _
                       Optional ByVal Y As Integer = 0)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim i   As Integer

        Dim lim As Integer

100     If Mapa <= 0 Or Mapa > NumMaps Then Exit Sub

102     If distanceToCities(Mapa).distanceToCity(city) >= 0 Then Exit Sub

104     If Mapa = Ciudades(city).Map Then
106         distanceToCities(Mapa).distanceToCity(city) = 0
        Else
108         distanceToCities(Mapa).distanceToCity(city) = Abs(X) + Abs(Y)

        End If

110     For i = 1 To 4
112         lim = getLimit(Mapa, i)

114         If lim > 0 Then

116             Select Case i

                    Case eHeading.NORTH
118                     Call setDistance(lim, city, i, X, Y - 1)

120                 Case eHeading.EAST
122                     Call setDistance(lim, city, i, X + 1, Y)

124                 Case eHeading.SOUTH
126                     Call setDistance(lim, city, i, X, Y + 1)

128                 Case eHeading.WEST
130                     Call setDistance(lim, city, i, X - 1, Y)

                End Select

            End If

132     Next i

End Sub

Public Sub goHome(ByVal Userindex As Integer)
        '***************************************************
        'Author: Budi
        'Last Modification: 01/06/2010
        '01/06/2010: ZaMa - Ahora usa otro tipo de intervalo
        '***************************************************

        Dim Distance As Long
        Dim Tiempo   As Long
    
100     With UserList(Userindex)

102         If .flags.Muerto = 1 Then
        
104             If .flags.lastMap = 0 Then
106                 Distance = distanceToCities(.Pos.Map).distanceToCity(.Hogar)
                
                Else
108                 Distance = distanceToCities(.flags.lastMap).distanceToCity(.Hogar) + GOHOME_PENALTY

                End If
            
110             Tiempo = (Distance + 1) * 13 'seg
            
                'If Tiempo > 60 Then Tiempo = 60
            
112             Call IntervaloGoHome(Userindex, Tiempo * 1000, True)
            
114             If .flags.lastMap = 0 Then
116                 Call WriteConsoleMsg(Userindex, "Te encuentras a " & CStr(Distance) & " mapas de " & MapInfo(Ciudades(.Hogar).Map).map_name & ", este viaje durara " & CStr(Tiempo) & " segundos.", FontTypeNames.FONTTYPE_FIGHT)
                Else
118                 Call WriteConsoleMsg(Userindex, "Te encuentras en un dungeon o en las catacumbas, viajarás a " & MapInfo(Ciudades(.Hogar).Map).map_name & " en " & CStr(Tiempo) & " segundos.", FontTypeNames.FONTTYPE_FIGHT)
                End If
            
            Else
        
120             Call WriteConsoleMsg(Userindex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
        End With
    
End Sub

''
' Maneja el tiempo de arrivo al hogar
'
' @param UserIndex  El index del usuario a ser afectado por el /hogar
'

Public Sub TravelingEffect(ByVal Userindex As Integer)
        '******************************************************
        'Author: ZaMa
        'Last Update: 01/06/2010 (ZaMa)
        '******************************************************

        ' Si ya paso el tiempo de penalizacion
100     If IntervaloGoHome(Userindex) Then
102         Call HomeArrival(Userindex)
        End If

End Sub


Public Function GetHomeArrivalTime(ByVal Userindex As Integer) As Integer

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        'Calculates the time left to arrive home.
        '**************************************************************
        Dim TActual As Long
    
100     TActual = GetTickCount()
    
102     With UserList(Userindex)
104         GetHomeArrivalTime = (.Counters.goHome - TActual) * 0.001
        End With

End Function

Public Sub HomeArrival(ByVal Userindex As Integer)
        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        'Teleports user to its home.
        '**************************************************************
    
        Dim tX   As Integer
        Dim tY   As Integer
        Dim tMap As Integer

100     With UserList(Userindex)

            'Antes de que el pj llegue a la ciudad, lo hacemos dejar de navegar para que no se buguee.
102         If .flags.Navegando = 1 Then
104             .Char.Body = iCuerpoMuerto
106             .Char.Head = iCabezaMuerto
108             .Char.ShieldAnim = NingunEscudo
110             .Char.WeaponAnim = NingunArma
112             .Char.CascoAnim = NingunCasco
            
114             .flags.Navegando = 0
            
116             Call WriteNavigateToggle(Userindex)

                'Le sacamos el navegando, pero no le mostramos a los demas porque va a ser sumoneado hasta ulla.
            End If
        
118         tX = Ciudades(.Hogar).X
120         tY = Ciudades(.Hogar).Y
122         tMap = Ciudades(.Hogar).Map
        
124         Call FindLegalPos(Userindex, tMap, CByte(tX), CByte(tY))
126         Call WarpUserChar(Userindex, tMap, tX, tY, True)
        
128         Call WriteConsoleMsg(Userindex, "El viaje ha terminado.", FontTypeNames.FONTTYPE_INFOBOLD)
        
130         .flags.Traveling = 0
132         .Counters.goHome = 0
        
        End With
    
End Sub

Public Function IntervaloGoHome(ByVal Userindex As Integer, _
                                Optional ByVal TimeInterval As Long, _
                                Optional ByVal Actualizar As Boolean = False) As Boolean

        '**************************************************************
        'Author: ZaMa
        'Last Modify by: ZaMa
        'Last Modify Date: 01/06/2010
        '01/06/2010: ZaMa - Add the Timer which determines wether the user can be teleported to its home or not
        '**************************************************************
    
        Dim TActual As Long
100         TActual = GetTickCount()
    
102     With UserList(Userindex)

            ' Inicializa el timer
104         If Actualizar Then
        
106             .flags.Traveling = 1
108             .Counters.goHome = TActual + TimeInterval
            
            Else

110             If TActual >= .Counters.goHome Then
112                 IntervaloGoHome = True
                End If

            End If

        End With

End Function

Public Sub HandleHome(ByVal Userindex As Integer)

        '***************************************************
        'Author: Budi
        'Creation Date: 06/01/2010
        'Last Modification: 05/06/10
        'Pato - 05/06/10: Add the UCase$ to prevent problems.
        '***************************************************
    
100     With UserList(Userindex)
        
102         Call .incomingData.ReadInteger

104         If .flags.Muerto = 0 Then
106             Call WriteConsoleMsg(Userindex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
        
            '¿Zona Segura?
108         If MapInfo(.Pos.Map).Seguro = 1 Then
110             Call WriteConsoleMsg(Userindex, "No puedes usar este comando en zona segura.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            
            End If
        
            'Si el mapa tiene alguna restriccion (newbie, dungeon, etc...), no lo dejamos viajar.
112         If MapInfo(.Pos.Map).zone = "NEWBIE" Then
114             Call WriteConsoleMsg(Userindex, "No pueder viajar a tu hogar desde este mapa.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            
            End If
        
            'Si es un mapa comun y no esta en cana
116         If .Counters.Pena <> 0 Then
118             Call WriteConsoleMsg(Userindex, "No puedes usar este comando en prisión.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
        
120         If .flags.Muerto = 0 Then
122             Call WriteConsoleMsg(Userindex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

124         If .flags.Traveling = 0 Then
            
126             If .Pos.Map <> Ciudades(.Hogar).Map Then
128                 Call goHome(Userindex)
                
                Else
130                 Call WriteConsoleMsg(Userindex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

132             .flags.Traveling = 0
134             .Counters.goHome = 0
            
136             Call WriteConsoleMsg(Userindex, "Ya hay un viaje en curso.", FontTypeNames.FONTTYPE_INFO)
            
            End If
        
        End With

End Sub
