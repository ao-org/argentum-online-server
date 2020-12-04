Attribute VB_Name = "Hogar"
Option Explicit

Type HomeDistance
    distanceToCity(1 To NUMCIUDADES) As Integer
End Type

Public Ciudades(1 To NUMCIUDADES)         As CityWorldPos
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

    If Mapa <= 0 Then Exit Function

    For X = 15 To 87
        For Y = 0 To 3

            Select Case side

                Case eHeading.NORTH
                    getLimit = MapData(Mapa, X, 7 + Y).TileExit.Map

                Case eHeading.EAST
                    getLimit = MapData(Mapa, 92 - Y, X).TileExit.Map

                Case eHeading.SOUTH
                    getLimit = MapData(Mapa, X, 94 - Y).TileExit.Map

                Case eHeading.WEST
                    getLimit = MapData(Mapa, 9 + Y, X).TileExit.Map

            End Select

            If getLimit > 0 Then Exit Function
        Next Y
    Next X

End Function

Public Sub generateMatrix(ByVal Mapa As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    Dim j As Integer
    
    ReDim distanceToCities(1 To NumMaps) As HomeDistance
    
    For j = 1 To NUMCIUDADES
        For i = 1 To NumMaps
            distanceToCities(i).distanceToCity(j) = -1
        Next i
    Next j
    
    For j = 1 To NUMCIUDADES
        For i = 1 To 4

            Select Case i

                Case eHeading.NORTH
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.NORTH), j, i, 0, -1)

                Case eHeading.EAST
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.EAST), j, i, 1, 0)

                Case eHeading.SOUTH
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.SOUTH), j, i, 0, 1)

                Case eHeading.WEST
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.WEST), j, i, -1, 0)

            End Select

        Next i
    Next j

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

    If Mapa <= 0 Or Mapa > NumMaps Then Exit Sub

    If distanceToCities(Mapa).distanceToCity(city) >= 0 Then Exit Sub

    If Mapa = Ciudades(city).Map Then
        distanceToCities(Mapa).distanceToCity(city) = 0
    Else
        distanceToCities(Mapa).distanceToCity(city) = Abs(X) + Abs(Y)

    End If

    For i = 1 To 4
        lim = getLimit(Mapa, i)

        If lim > 0 Then

            Select Case i

                Case eHeading.NORTH
                    Call setDistance(lim, city, i, X, Y - 1)

                Case eHeading.EAST
                    Call setDistance(lim, city, i, X + 1, Y)

                Case eHeading.SOUTH
                    Call setDistance(lim, city, i, X, Y + 1)

                Case eHeading.WEST
                    Call setDistance(lim, city, i, X - 1, Y)

            End Select

        End If

    Next i

End Sub

Public Sub goHome(ByVal Userindex As Integer)
    '***************************************************
    'Author: Budi
    'Last Modification: 01/06/2010
    '01/06/2010: ZaMa - Ahora usa otro tipo de intervalo
    '***************************************************

    Dim Distance As Long
    Dim Tiempo   As Long
    
    With UserList(Userindex)

        If .flags.Muerto = 1 Then
        
            If .flags.lastMap = 0 Then
                Distance = distanceToCities(.Pos.Map).distanceToCity(.Hogar)
                
            Else
                Distance = distanceToCities(.flags.lastMap).distanceToCity(.Hogar) + GOHOME_PENALTY

            End If
            
            Tiempo = (Distance + 1) * 13 'seg
            
            'If Tiempo > 60 Then Tiempo = 60
            
            Call IntervaloGoHome(Userindex, Tiempo * 1000, True)
                
            Call WriteConsoleMsg(Userindex, "Te encuentras a " & CStr(Distance) & " mapas de " & MapInfo(Ciudades(.Hogar).Map).map_name & ", este viaje durara " & CStr(Tiempo) & " segundos.", FontTypeNames.FONTTYPE_FIGHT)
            
        Else
        
            Call WriteConsoleMsg(Userindex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)

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
    If IntervaloGoHome(Userindex) Then
        Call HomeArrival(Userindex)
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
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(Userindex)
        GetHomeArrivalTime = (.Counters.goHome - TActual) * 0.001
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

    With UserList(Userindex)

        'Antes de que el pj llegue a la ciudad, lo hacemos dejar de navegar para que no se buguee.
        If .flags.Navegando = 1 Then
            .Char.Body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            
            .flags.Navegando = 0
            
            Call WriteNavigateToggle(Userindex)

            'Le sacamos el navegando, pero no le mostramos a los demas porque va a ser sumoneado hasta ulla.
        End If
        
        tX = Ciudades(.Hogar).X
        tY = Ciudades(.Hogar).Y
        tMap = Ciudades(.Hogar).Map
        
        Call FindLegalPos(Userindex, tMap, CByte(tX), CByte(tY))
        Call WarpUserChar(Userindex, tMap, tX, tY, True)
        
        Call WriteConsoleMsg(Userindex, "El viaje ha terminado.", FontTypeNames.FONTTYPE_INFOBOLD)
        
        .flags.Traveling = 0
        .Counters.goHome = 0
        
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
        TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(Userindex)

        ' Inicializa el timer
        If Actualizar Then
        
            .flags.Traveling = 1
            .Counters.goHome = TActual + TimeInterval
            
        Else

            If TActual >= .Counters.goHome Then
                IntervaloGoHome = True
            End If

        End If

    End With

End Function

''
' Handles the "Home" message.
'
' @param    userIndex The index of the user sending the message.
Public Sub HandleHome(ByVal Userindex As Integer)

    '***************************************************
    'Author: Budi
    'Creation Date: 06/01/2010
    'Last Modification: 05/06/10
    'Pato - 05/06/10: Add the UCase$ to prevent problems.
    '***************************************************
    
    With UserList(Userindex)
        
        Call .incomingData.ReadInteger

        If .flags.Muerto = 0 Then
            Call WriteConsoleMsg(Userindex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

        'Si es un mapa comUn y no esta en cana
        If .Counters.Pena <> 0 Then
            Call WriteConsoleMsg(Userindex, "No puedes usar este comando aqui.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
        
        If .flags.Traveling = 0 Then
            
            If .Pos.Map <> Ciudades(.Hogar).Map Then
                Call goHome(Userindex)
                
            Else
                Call WriteConsoleMsg(Userindex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)

            End If

        Else

            .flags.Traveling = 0
            .Counters.goHome = 0
            
            Call WriteConsoleMsg(Userindex, "Ya hay un viaje en curso.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        
    End With

End Sub

