Attribute VB_Name = "Extra"
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

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByRef x As Byte, ByRef Y As Byte)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 26/03/2009
    'Search for a Legal pos for the user who is being teleported.
    '***************************************************

    If MapData(Map, x, Y).UserIndex <> 0 Or MapData(Map, x, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(Map, x, Y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace     As Boolean

        Dim tX             As Long

        Dim tY             As Long

        Dim Rango          As Long

        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = x - Rango To x + Rango

                    'Reviso que no haya User ni NPC
                    If MapData(Map, tX, tY).UserIndex = 0 And MapData(Map, tX, tY).NpcIndex = 0 Then
                        
                        If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        
                        Exit For

                    End If

                Next tX
        
                If FoundPlace Then Exit For
            Next tY
            
            If FoundPlace Then Exit For
        Next Rango
    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            x = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(Map, x, Y).UserIndex

            If OtherUserIndex <> 0 Then

                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then

                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        

                    End If

                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        

                    End If

                End If
            
                Call CloseSocket(OtherUserIndex)

            End If

        End If

    End If

End Sub

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
    EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie

End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************
    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)

End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************
    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)

End Function

Public Function EsGM(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 23/01/2007
    '***************************************************
    EsGM = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))

End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)

    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 23/01/2007
    'Handles the Map passage of Users. Allows the existance
    'of exclusive maps for Newbies, Royal Army and Caos Legion members
    'and enables GMs to enter every map without restriction.
    'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
    '***************************************************
    On Error GoTo Errhandler

    Dim nPos   As WorldPos

    Dim FxFlag As Boolean

    'Controla las salidas
    If InMapBounds(Map, x, Y) Then
    
        If MapData(Map, x, Y).ObjInfo.ObjIndex > 0 Then
            FxFlag = ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport

        End If

        If (MapData(Map, x, Y).TileExit.Map > 0) And (MapData(Map, x, Y).TileExit.Map <= NumMaps) Then
    
            '¿Es mapa de newbies?
            If UCase$(MapInfo(MapData(Map, x, Y).TileExit.Map).restrict_mode) = "NEWBIE" Then

                '¿El usuario es un newbie?
                If EsNewbie(UserIndex) Or EsGM(UserIndex) Then
                    If LegalPos(MapData(Map, x, Y).TileExit.Map, MapData(Map, x, Y).TileExit.x, MapData(Map, x, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex), , , False) Then
                        Call WarpUserChar(UserIndex, MapData(Map, x, Y).TileExit.Map, MapData(Map, x, Y).TileExit.x, MapData(Map, x, Y).TileExit.Y, FxFlag)
                
                    Else
                        Call ClosestLegalPos(MapData(Map, x, Y).TileExit, nPos)

                        If nPos.x <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.Y, FxFlag)

                        End If

                    End If

                Else 'No es newbie
                    Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                    Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                    If nPos.x <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.Y, FxFlag)

                    End If

                End If

            Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.

                If LegalPos(MapData(Map, x, Y).TileExit.Map, MapData(Map, x, Y).TileExit.x, MapData(Map, x, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex), , , False) Then
                    Call WarpUserChar(UserIndex, MapData(Map, x, Y).TileExit.Map, MapData(Map, x, Y).TileExit.x, MapData(Map, x, Y).TileExit.Y, FxFlag)
                Else
                    Call ClosestLegalPos(MapData(Map, x, Y).TileExit, nPos)

                    If nPos.x <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.Y, FxFlag)

                    End If

                End If

            End If

            'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
            Dim aN As Integer
    
            aN = UserList(UserIndex).flags.AtacadoPorNpc

            If aN > 0 Then
                Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                Npclist(aN).flags.AttackedBy = vbNullString

            End If
    
            aN = UserList(UserIndex).flags.NPCAtacado

            If aN > 0 Then
                If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
                    Npclist(aN).flags.AttackedFirstBy = vbNullString

                End If

            End If

            UserList(UserIndex).flags.AtacadoPorNpc = 0
            UserList(UserIndex).flags.NPCAtacado = 0

        End If
    
    End If

    Exit Sub

Errhandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.description)

End Sub

Function InRangoVision(ByVal UserIndex As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean

    If x > UserList(UserIndex).Pos.x - MinXBorder And x < UserList(UserIndex).Pos.x + MinXBorder Then
        If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
            InRangoVision = True
            Exit Function

        End If

    End If

    InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, x As Integer, Y As Integer) As Boolean

    If x > Npclist(NpcIndex).Pos.x - MinXBorder And x < Npclist(NpcIndex).Pos.x + MinXBorder Then
        If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
            InRangoVisionNPC = True
            Exit Function

        End If

    End If

    InRangoVisionNPC = False

End Function

Function InMapBounds(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean
            
    If (Map <= 0 Or Map > NumMaps) Or x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True

    End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
    '*****************************************************************
    'Author: Unknown (original version)
    'Last Modification: 24/01/2007 (ToxicWaste)
    'Encuentra la posicion legal mas cercana y la guarda en nPos
    '*****************************************************************

    Dim Notfound As Boolean

    Dim LoopC    As Integer

    Dim tX       As Integer

    Dim tY       As Integer

    nPos.Map = Pos.Map

    Do While Not LegalPos(Pos.Map, nPos.x, nPos.Y, PuedeAgua, PuedeTierra, , False)

        If LoopC > 12 Then
            Notfound = True
            Exit Do

        End If
    
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.x - LoopC To Pos.x + LoopC
            
                If LegalPos(nPos.Map, tX, tY, PuedeAgua, PuedeTierra, , False) Then
                    nPos.x = tX
                    nPos.Y = tY
                    '¿Hay objeto?
                
                    tX = Pos.x + LoopC
                    tY = Pos.Y + LoopC
  
                End If
        
            Next tX
        Next tY
    
        LoopC = LoopC + 1
    
    Loop

    If Notfound = True Then
        nPos.x = 0
        nPos.Y = 0

    End If

End Sub

Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
    '*****************************************************************
    'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
    '*****************************************************************

    Dim Notfound As Boolean

    Dim LoopC    As Integer

    Dim tX       As Integer

    Dim tY       As Integer

    nPos.Map = Pos.Map

    Do While Not LegalPos(Pos.Map, nPos.x, nPos.Y)

        If LoopC > 12 Then
            Notfound = True
            Exit Do

        End If
    
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.x - LoopC To Pos.x + LoopC
            
                If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                    nPos.x = tX
                    nPos.Y = tY
                    '¿Hay objeto?
                
                    tX = Pos.x + LoopC
                    tY = Pos.Y + LoopC
  
                End If
        
            Next tX
        Next tY
    
        LoopC = LoopC + 1
    
    Loop

    If Notfound = True Then
        nPos.x = 0
        nPos.Y = 0

    End If

End Sub

Function NameIndex(ByVal name As String) As Integer

    Dim UserIndex As Integer

    '¿Nombre valido?
    If LenB(name) = 0 Then
        NameIndex = 0
        Exit Function

    End If

    If InStrB(name, "+") <> 0 Then
        name = UCase$(Replace(name, "+", " "))

    End If

    UserIndex = 1

    Do Until UCase$(UserList(UserIndex).name) = UCase$(name)
    
        UserIndex = UserIndex + 1
    
        If UserIndex > MaxUsers Then
            NameIndex = 0
            Exit Function

        End If
    
    Loop
 
    NameIndex = UserIndex
 
End Function

Function IP_Index(ByVal inIP As String) As Integer
 
    Dim UserIndex As Integer

    '¿Nombre valido?
    If LenB(inIP) = 0 Then
        IP_Index = 0
        Exit Function

    End If
  
    UserIndex = 1

    Do Until UserList(UserIndex).ip = inIP
    
        UserIndex = UserIndex + 1
    
        If UserIndex > MaxUsers Then
            IP_Index = 0
            Exit Function

        End If
    
    Loop
 
    IP_Index = UserIndex

    Exit Function

End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean

    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers

        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
                CheckForSameIP = True
                Exit Function

            End If

        End If

    Next LoopC

    CheckForSameIP = False

End Function

Function CheckForSameName(ByVal name As String) As Boolean

    'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long

    For LoopC = 1 To LastUser

        If UserList(LoopC).flags.UserLogged Then
        
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
        
            If UCase$(UserList(LoopC).name) = UCase$(name) Then
                CheckForSameName = True
                Exit Function

            End If

        End If

    Next LoopC

    CheckForSameName = False

End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)

    '*****************************************************************
    'Toma una posicion y se mueve hacia donde esta perfilado
    '*****************************************************************
    Dim x  As Integer

    Dim Y  As Integer

    Dim nX As Integer

    Dim nY As Integer

    x = Pos.x
    Y = Pos.Y

    If Head = eHeading.NORTH Then
        nX = x
        nY = Y - 1

    End If

    If Head = eHeading.SOUTH Then
        nX = x
        nY = Y + 1

    End If

    If Head = eHeading.EAST Then
        nX = x + 1
        nY = Y

    End If

    If Head = eHeading.WEST Then
        nX = x - 1
        nY = Y

    End If

    'Devuelve valores
    Pos.x = nX
    Pos.Y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal Montado As Boolean = False, Optional ByVal PuedeTraslado As Boolean = True) As Boolean
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 23/01/2007
    'Checks if the position is Legal.
    '***************************************************
    '¿Es un mapa valido?

    If (Map <= 0 Or Map > NumMaps) Or (x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPos = False
    Else

        If PuedeAgua And PuedeTierra Then
            LegalPos = (MapData(Map, x, Y).Blocked = 0) And (MapData(Map, x, Y).UserIndex = 0) And (MapData(Map, x, Y).NpcIndex = 0) And (PuedeTraslado Or MapData(Map, x, Y).TileExit.Map = 0)

        ElseIf PuedeTierra And Not PuedeAgua Then
            LegalPos = (MapData(Map, x, Y).Blocked = 0) And (MapData(Map, x, Y).UserIndex = 0) And (MapData(Map, x, Y).NpcIndex = 0) And (Not HayAgua(Map, x, Y)) And (PuedeTraslado Or MapData(Map, x, Y).TileExit.Map = 0)

        ElseIf PuedeAgua And Not PuedeTierra Then
            LegalPos = (MapData(Map, x, Y).Blocked = 0) And (MapData(Map, x, Y).UserIndex = 0) And (MapData(Map, x, Y).NpcIndex = 0) And (HayAgua(Map, x, Y)) And (PuedeTraslado Or MapData(Map, x, Y).TileExit.Map = 0)
        Else
            LegalPos = False

        End If
   
    End If

End Function

Function LegalPosNPC(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean

    If (Map <= 0 Or Map > NumMaps) Or (x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
    Else

        If AguaValida = 0 Then
            LegalPosNPC = (MapData(Map, x, Y).Blocked = 0) And (MapData(Map, x, Y).UserIndex = 0) And (MapData(Map, x, Y).NpcIndex = 0) And (MapData(Map, x, Y).trigger <> eTrigger.POSINVALIDA) And Not HayAgua(Map, x, Y)
        Else
            LegalPosNPC = (MapData(Map, x, Y).Blocked = 0) And (MapData(Map, x, Y).UserIndex = 0) And (MapData(Map, x, Y).NpcIndex = 0) And (MapData(Map, x, Y).trigger <> eTrigger.POSINVALIDA)

        End If
 
    End If

End Function

Sub SendHelp(ByVal Index As Integer)

    Dim NumHelpLines As Integer

    Dim LoopC        As Integer

    NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

    For LoopC = 1 To NumHelpLines
        Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
    Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    If Npclist(NpcIndex).NroExpresiones > 0 Then

        Dim randomi

        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))

    End If

End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)

    'Responde al click del usuario sobre el mapa
    Dim FoundChar      As Byte

    Dim FoundSomething As Byte

    Dim TempCharIndex  As Integer

    Dim Stat           As String

    Dim ft             As FontTypeNames

    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.x - x) > RANGO_VISION_X) Then
        Exit Sub

    End If

    '¿Posicion valida?
    If InMapBounds(Map, x, Y) Then
        UserList(UserIndex).flags.TargetMap = Map
        UserList(UserIndex).flags.TargetX = x
        UserList(UserIndex).flags.TargetY = Y

        '¿Es un obj?
        If MapData(Map, x, Y).ObjInfo.ObjIndex > 0 Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = x
            UserList(UserIndex).flags.TargetObjY = Y
            FoundSomething = 1
        ElseIf MapData(Map, x + 1, Y).ObjInfo.ObjIndex > 0 Then

            'Informa el nombre
            If ObjData(MapData(Map, x + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                UserList(UserIndex).flags.TargetObjMap = Map
                UserList(UserIndex).flags.TargetObjX = x + 1
                UserList(UserIndex).flags.TargetObjY = Y
                FoundSomething = 1

            End If

        ElseIf MapData(Map, x + 1, Y + 1).ObjInfo.ObjIndex > 0 Then

            If ObjData(MapData(Map, x + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                'Informa el nombre
                UserList(UserIndex).flags.TargetObjMap = Map
                UserList(UserIndex).flags.TargetObjX = x + 1
                UserList(UserIndex).flags.TargetObjY = Y + 1
                FoundSomething = 1

            End If

        ElseIf MapData(Map, x, Y + 1).ObjInfo.ObjIndex > 0 Then

            If ObjData(MapData(Map, x, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                'Informa el nombre
                UserList(UserIndex).flags.TargetObjMap = Map
                UserList(UserIndex).flags.TargetObjX = x
                UserList(UserIndex).flags.TargetObjY = Y + 1
                FoundSomething = 1

            End If

        End If
    
        If FoundSomething = 1 Then
            UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex

            If UserList(UserIndex).Counters.Trabajando = 0 Then
                If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then

                    Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "* - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.Amount & "", FontTypeNames.FONTTYPE_INFO)
            
                Else

                    If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = eOBJType.otYacimiento Then
                        Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
                        Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).name & " - (Minerales disponibles: " & MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.Amount & ")", FontTypeNames.FONTTYPE_INFO)

                    ElseIf ObjData(UserList(UserIndex).flags.TargetObj).OBJType = eOBJType.otArboles Then
                        Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
                        Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).name & " - (Recursos disponibles: " & MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.Amount & ")", FontTypeNames.FONTTYPE_INFO)
                    
                    Else
                        Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "*", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If
    
        End If

        '¿Es un personaje?
        If Y + 1 <= YMaxMapSize Then
            If MapData(Map, x, Y + 1).UserIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y + 1).UserIndex
                FoundChar = 1

            End If

            If MapData(Map, x, Y + 1).NpcIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y + 1).NpcIndex
                FoundChar = 2

            End If

        End If

        '¿Es un personaje?
        If FoundChar = 0 Then
            If MapData(Map, x, Y).UserIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y).UserIndex
                FoundChar = 1

            End If

            If MapData(Map, x, Y).NpcIndex > 0 Then
                TempCharIndex = MapData(Map, x, Y).NpcIndex
                FoundChar = 2

            End If

        End If
    
        'Reaccion al personaje
        If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
            If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios And PlayerType.Dios Then
            
                'If LenB(UserList(TempCharIndex).DescRM) = 0 Then 'No tiene descRM y quiere que se vea su nombre.
                
                If UserList(TempCharIndex).flags.Privilegios = user Then
                
                    Dim Fragsnick As String

                    'If Abs(CInt(UserList(TempCharIndex).Stats.ELV) - CInt(UserList(UserIndex).Stats.ELV)) < 10 Then
                    Stat = Stat & " <" & ListaClases(UserList(TempCharIndex).clase) & " " & ListaRazas(UserList(TempCharIndex).raza) & " Nivel: " & UserList(TempCharIndex).Stats.ELV & ">"

                    'End If
                    If EsNewbie(TempCharIndex) Then
                        Stat = Stat & " <Newbie>"

                    End If

                    If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 49 Then
                        If UserList(TempCharIndex).flags.Envenenado > 0 Then
                            Fragsnick = " | Envenenado "

                        End If

                        If UserList(TempCharIndex).flags.Ceguera = 1 Then
                            Fragsnick = Fragsnick & " | Ciego "

                        End If

                        If UserList(TempCharIndex).flags.Incinerado = 1 Then
                            Fragsnick = Fragsnick & " | Incinerado "

                        End If

                        If UserList(TempCharIndex).flags.Paralizado = 1 Then
                            Fragsnick = Fragsnick & " | Paralizado "

                        End If

                        If UserList(TempCharIndex).flags.Inmovilizado = 1 Then
                            Fragsnick = Fragsnick & " | Inmovilizado "

                        End If

                        If UserList(TempCharIndex).Counters.Trabajando > 0 Then
                            Fragsnick = Fragsnick & " | Trabajando "

                        End If

                        If UserList(TempCharIndex).flags.invisible = 1 Then
                            Fragsnick = Fragsnick & " | Invisible "

                        End If

                        If UserList(TempCharIndex).flags.Oculto = 1 Then
                            Fragsnick = Fragsnick & " | Oculto "

                        End If

                        If UserList(TempCharIndex).flags.Estupidez = 1 Then
                            Fragsnick = Fragsnick & " | Estupido "

                        End If

                        If UserList(TempCharIndex).flags.Maldicion = 1 Then
                            Fragsnick = Fragsnick & " | Maldito "

                        End If

                        If UserList(TempCharIndex).flags.Silenciado = 1 Then
                            Fragsnick = Fragsnick & " | Silenciado "

                        End If

                        If UserList(TempCharIndex).flags.Comerciando = True Then
                            Fragsnick = Fragsnick & " | Comerciando "

                        End If

                        If UserList(TempCharIndex).flags.Descansar = 1 Then
                            Fragsnick = Fragsnick & " | Descansando "

                        End If

                        If UserList(TempCharIndex).flags.Meditando = True Then
                            Fragsnick = Fragsnick & " | Concentrado "
            
                        End If

                        If UserList(TempCharIndex).flags.BattleModo = 1 Then
                            Fragsnick = Fragsnick & " | Modo Battle "
            
                        End If
                        
                        If UserList(TempCharIndex).Stats.MinHp = 0 Then
                            Stat = Stat & " <Muerto>"
                        ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.1) Then
                            Stat = Stat & " <Casi muerto" & Fragsnick & ">"
                        ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.5) Then
                            Stat = Stat & " <Malherido" & Fragsnick & ">"
                        ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.75) Then
                            Stat = Stat & " <Herido" & Fragsnick & ">"
                        ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.99) Then
                            Stat = Stat & " <Levemente herido" & Fragsnick & ">"
                        Else
                            Stat = Stat & " <Intacto" & Fragsnick & ">"

                        End If

                    End If
                
                    If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                        Stat = Stat & " <" & TituloReal(TempCharIndex) & ">"
                        ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                    ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                        Stat = Stat & " <" & TituloCaos(TempCharIndex) & ">"
                        ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA

                    End If
                
                    If UserList(TempCharIndex).GuildIndex > 0 Then
                        Stat = Stat & " <" & modGuilds.GuildName(UserList(TempCharIndex).GuildIndex) & ">"

                    End If

                End If ' If user > 0 then
                            
                If Not UserList(TempCharIndex).flags.Privilegios And PlayerType.user Then
                    If UserList(TempCharIndex).flags.Privilegios = Consejero Then
                        Stat = Stat & " <Game Desing>"
                        ft = FontTypeNames.FONTTYPE_GM

                    End If

                    If UserList(TempCharIndex).flags.Privilegios = SemiDios Then
                        Stat = Stat & " <Game Master>"
                        ft = FontTypeNames.FONTTYPE_GM

                    End If

                    If UserList(TempCharIndex).flags.Privilegios = Dios Then
                        Stat = Stat & " <Administrador>"
                        ft = FontTypeNames.FONTTYPE_DIOS

                    End If
                        
                    If UserList(TempCharIndex).flags.Privilegios = PlayerType.Admin Then
                        Stat = Stat & " <Administrador>"
                        ft = FontTypeNames.FONTTYPE_DIOS

                    End If
                    
                ElseIf UserList(TempCharIndex).Faccion.Status = 0 Then
                    ft = FontTypeNames.FONTTYPE_New_Gris
                ElseIf UserList(TempCharIndex).Faccion.Status = 1 Then
                    ft = FontTypeNames.FONTTYPE_CITIZEN

                End If
                    
                If UserList(TempCharIndex).flags.Casado = 1 Then
                    Stat = Stat & " <Pareja de " & UserList(TempCharIndex).flags.Pareja & ">"

                End If
                    
                If Len(UserList(TempCharIndex).Desc) > 0 Then
                    Stat = "Ves a [" & UserList(TempCharIndex).name & "]" & Stat & " - " & UserList(TempCharIndex).Desc
                Else
                    Stat = "Ves a [" & UserList(TempCharIndex).name & "]" & Stat

                End If
                 
                ' Else  'Si tiene descRM la muestro siempre.
                '   Stat = UserList(TempCharIndex).DescRM
                '   ft = FontTypeNames.FONTTYPE_INFOBOLD
                ' End If
            
                If LenB(Stat) > 0 Then
                    Call WriteConsoleMsg(UserIndex, Stat, ft)

                End If
            
                FoundSomething = 1
                UserList(UserIndex).flags.TargetUser = TempCharIndex
                UserList(UserIndex).flags.TargetNPC = 0
                UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun

            End If

        End If

        If FoundChar = 2 Then '¿Encontro un NPC?

            Dim estatus As String
            
            'If UserList(UserIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
            '  estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
            '  Else
                        
            If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 75 Then
                If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.1) Then
                    estatus = "<Agonizando"
                ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.2) Then
                    estatus = "<Casi muerto"
                ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.5) Then
                    estatus = "<Malherido"
                ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.7) Then
                    estatus = "<Herido"
                ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.95) Then
                    estatus = "<Levemente herido"
                Else
                    estatus = "<Intacto"

                End If

            Else
                estatus = "<" & Npclist(TempCharIndex).Stats.MinHp & "/" & Npclist(TempCharIndex).Stats.MaxHp
                        
            End If
                        
            If Npclist(TempCharIndex).flags.Envenenado > 0 Then
                estatus = estatus & " | Envenenado"

            End If
                        
            If Npclist(TempCharIndex).flags.Paralizado = 1 Then
                If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 49 Then
                    estatus = estatus & " | Paralizado (" & CInt(Npclist(TempCharIndex).Contadores.Paralisis / 6.5) & "s)"
                Else
                    estatus = estatus & " | Paralizado"

                End If

            End If
                        
            If Npclist(TempCharIndex).flags.Inmovilizado = 1 Then
                If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 49 Then
                    estatus = estatus & " | Inmovilizado (" & CInt(Npclist(TempCharIndex).Contadores.Paralisis / 6.5) & "s)"
                Else
                    estatus = estatus & " | Inmovilizado"

                End If

            End If
                        
            estatus = estatus & ">"
    
            'End If
            
            If Len(Npclist(TempCharIndex).Desc) > 1 Then
                'Optimizacion de protocolo por Ladder
                Call WriteChatOverHead(UserIndex, "NPCDESC*" & Npclist(TempCharIndex).Numero, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
            ElseIf TempCharIndex = CentinelaNPCIndex Then
                'Enviamos nuevamente el texto del centinela según quien pregunta
                Call modCentinela.CentinelaSendClave(UserIndex)
            Else
                
                Call WriteConsoleMsg(UserIndex, "NPCNAME*" & Npclist(TempCharIndex).Numero & "*" & " " & estatus, FontTypeNames.FONTTYPE_INFO)
                ' If UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                ' Call WriteConsoleMsg(UserIndex, "Le pegó primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, Npclist(TempCharIndex).Char.CharIndex, FontTypeNames.FONTTYPE_INFO)
                'End If
                
            End If

            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNPC = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
        
        End If
    
        If FoundChar = 0 Then
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
            UserList(UserIndex).flags.TargetUser = 0

        End If
    
        '*** NO ENCOTRO NADA ***
        If FoundSomething = 0 Then
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
            UserList(UserIndex).flags.TargetObjMap = 0
            UserList(UserIndex).flags.TargetObjX = 0
            UserList(UserIndex).flags.TargetObjY = 0

            ' Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", FontTypeNames.FONTTYPE_INFO)
        End If

    Else

        If FoundSomething = 0 Then
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
            UserList(UserIndex).flags.TargetObjMap = 0
            UserList(UserIndex).flags.TargetObjX = 0
            UserList(UserIndex).flags.TargetObjY = 0

            '  Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", FontTypeNames.FONTTYPE_INFO)
        End If

    End If

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading

    '*****************************************************************
    'Devuelve la direccion en la cual el target se encuentra
    'desde pos, 0 si la direc es igual
    '*****************************************************************
    Dim x As Integer

    Dim Y As Integer

    x = Pos.x - Target.x
    Y = Pos.Y - Target.Y

    'NE
    If Sgn(x) = -1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
        Exit Function

    End If

    'NW
    If Sgn(x) = 1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
        Exit Function

    End If

    'SW
    If Sgn(x) = 1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
        Exit Function

    End If

    'SE
    If Sgn(x) = -1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
        Exit Function

    End If

    'Sur
    If Sgn(x) = 0 And Sgn(Y) = -1 Then
        FindDirection = eHeading.SOUTH
        Exit Function

    End If

    'norte
    If Sgn(x) = 0 And Sgn(Y) = 1 Then
        FindDirection = eHeading.NORTH
        Exit Function

    End If

    'oeste
    If Sgn(x) = 1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.WEST
        Exit Function

    End If

    'este
    If Sgn(x) = -1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.EAST
        Exit Function

    End If

    'misma
    If Sgn(x) = 0 And Sgn(Y) = 0 Then
        FindDirection = 0
        Exit Function

    End If

End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean

    ItemNoEsDeMapa = ObjData(Index).OBJType <> eOBJType.otPuertas And ObjData(Index).OBJType <> eOBJType.otForos And ObjData(Index).OBJType <> eOBJType.otCarteles And ObjData(Index).OBJType <> eOBJType.otArboles And ObjData(Index).OBJType <> eOBJType.otYacimiento And ObjData(Index).OBJType <> eOBJType.otTeleport And ObjData(Index).OBJType <> eOBJType.OtCorreo And ObjData(Index).OBJType <> eOBJType.OtDecoraciones

End Function

'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
    MostrarCantidad = ObjData(Index).OBJType <> eOBJType.otPuertas And ObjData(Index).OBJType <> eOBJType.otForos And ObjData(Index).OBJType <> eOBJType.otCarteles And ObjData(Index).OBJType <> eOBJType.otYacimiento And ObjData(Index).OBJType <> eOBJType.otArboles And ObjData(Index).OBJType <> eOBJType.OtCorreo And ObjData(Index).OBJType <> eOBJType.otTeleport

End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

    EsObjetoFijo = OBJType = eOBJType.otForos Or OBJType = eOBJType.otCarteles Or OBJType = eOBJType.otArboles Or OBJType = eOBJType.otYacimiento
    OBJType = eOBJType.OtDecoraciones

End Function
