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

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByRef X As Byte, ByRef Y As Byte)
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 26/03/2009
        'Search for a Legal pos for the user who is being teleported.
        '***************************************************
        
        On Error GoTo FindLegalPos_Err
        

100     If MapData(Map, X, Y).UserIndex <> 0 Or MapData(Map, X, Y).NpcIndex <> 0 Then
                    
            ' Se teletransporta a la misma pos a la que estaba
102         If MapData(Map, X, Y).UserIndex = UserIndex Then Exit Sub
                            
            Dim FoundPlace     As Boolean

            Dim tX             As Long

            Dim tY             As Long

            Dim Rango          As Long

            Dim OtherUserIndex As Integer
    
104         For Rango = 1 To 5
106             For tY = Y - Rango To Y + Rango
108                 For tX = X - Rango To X + Rango

                        'Reviso que no haya User ni NPC
110                     If MapData(Map, tX, tY).UserIndex = 0 And MapData(Map, tX, tY).NpcIndex = 0 Then
                        
112                         If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        
                            Exit For

                        End If

114                 Next tX
        
116                 If FoundPlace Then Exit For
118             Next tY
            
120             If FoundPlace Then Exit For
122         Next Rango
    
124         If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
126             X = tX
128             Y = tY
            Else
                'Muy poco probable, pero..
                'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
130             OtherUserIndex = MapData(Map, X, Y).UserIndex

132             If OtherUserIndex <> 0 Then

                    'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
134                 If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then

                        'Le avisamos al que estaba comerciando que se tuvo que ir.
136                     If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
138                         Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
140                         Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        

                        End If

                        'Lo sacamos.
142                     If UserList(OtherUserIndex).flags.UserLogged Then
144                         Call FinComerciarUsu(OtherUserIndex)
146                         Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        

                        End If

                    End If
            
148                 Call CloseSocket(OtherUserIndex)

                End If

            End If

        End If

        
        Exit Sub

FindLegalPos_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.FindLegalPos", Erl)
        Resume Next
        
End Sub

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo EsNewbie_Err
        
100     EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie

        
        Exit Function

EsNewbie_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.EsNewbie", Erl)
        Resume Next
        
End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 23/01/2007
        '***************************************************
        
        On Error GoTo esArmada_Err
        
100     esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)

        
        Exit Function

esArmada_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.esArmada", Erl)
        Resume Next
        
End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 23/01/2007
        '***************************************************
        
        On Error GoTo esCaos_Err
        
100     esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)

        
        Exit Function

esCaos_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.esCaos", Erl)
        Resume Next
        
End Function

Public Function EsGM(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 23/01/2007
        '***************************************************
        
        On Error GoTo EsGM_Err
        
100     EsGM = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))

        
        Exit Function

EsGM_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.EsGM", Erl)
        Resume Next
        
End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 23/01/2007
    'Handles the Map passage of Users. Allows the existance
    'of exclusive maps for Newbies, Royal Army and Caos Legion members
    'and enables GMs to enter every map without restriction.
    'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
    '***************************************************
    On Error GoTo ErrHandler

    Dim nPos   As WorldPos

    Dim FxFlag As Boolean

    'Controla las salidas
    If InMapBounds(Map, X, Y) Then
    
        If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
            FxFlag = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport

        End If

        If (MapData(Map, X, Y).TileExit.Map > 0) And (MapData(Map, X, Y).TileExit.Map <= NumMaps) Then
    
            '¿Es mapa de newbies?
            If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).restrict_mode) = "NEWBIE" Then

                '¿El usuario es un newbie?
                If EsNewbie(UserIndex) Or EsGM(UserIndex) Then
                    If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex), , , False) Then
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, FxFlag)
                
                    Else
                        Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)

                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                        End If

                    End If

                Else 'No es newbie
                    Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                    Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

                    End If

                End If

            Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.

                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex), , , False) Then
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, FxFlag)
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)

                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)

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

ErrHandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.description)

End Sub

Function InRangoVision(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo InRangoVision_Err
        

100     If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
102         If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
104             InRangoVision = True
                Exit Function

            End If

        End If

106     InRangoVision = False

        
        Exit Function

InRangoVision_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.InRangoVision", Erl)
        Resume Next
        
End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean
        
        On Error GoTo InRangoVisionNPC_Err
        

100     If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
102         If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
104             InRangoVisionNPC = True
                Exit Function

            End If

        End If

106     InRangoVisionNPC = False

        
        Exit Function

InRangoVisionNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.InRangoVisionNPC", Erl)
        Resume Next
        
End Function

Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo InMapBounds_Err
        
            
100     If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
102         InMapBounds = False
        Else
104         InMapBounds = True

        End If

        
        Exit Function

InMapBounds_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.InMapBounds", Erl)
        Resume Next
        
End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
        '*****************************************************************
        'Author: Unknown (original version)
        'Last Modification: 24/01/2007 (ToxicWaste)
        'Encuentra la posicion legal mas cercana y la guarda en nPos
        '*****************************************************************
        
        On Error GoTo ClosestLegalPos_Err
        

        Dim Notfound As Boolean

        Dim LoopC    As Integer

        Dim tX       As Integer

        Dim tY       As Integer

100     nPos.Map = Pos.Map

102     Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, , False)

104         If LoopC > 12 Then
106             Notfound = True
                Exit Do

            End If
    
108         For tY = Pos.Y - LoopC To Pos.Y + LoopC
110             For tX = Pos.X - LoopC To Pos.X + LoopC
            
112                 If LegalPos(nPos.Map, tX, tY, PuedeAgua, PuedeTierra, , False) Then
114                     nPos.X = tX
116                     nPos.Y = tY
                        '¿Hay objeto?
                
118                     tX = Pos.X + LoopC
120                     tY = Pos.Y + LoopC
  
                    End If
        
122             Next tX
124         Next tY
    
126         LoopC = LoopC + 1
    
        Loop

128     If Notfound = True Then
130         nPos.X = 0
132         nPos.Y = 0

        End If

        
        Exit Sub

ClosestLegalPos_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.ClosestLegalPos", Erl)
        Resume Next
        
End Sub

Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
        '*****************************************************************
        'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
        '*****************************************************************
        
        On Error GoTo ClosestStablePos_Err
        

        Dim Notfound As Boolean

        Dim LoopC    As Integer

        Dim tX       As Integer

        Dim tY       As Integer

100     nPos.Map = Pos.Map

102     Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)

104         If LoopC > 12 Then
106             Notfound = True
                Exit Do

            End If
    
108         For tY = Pos.Y - LoopC To Pos.Y + LoopC
110             For tX = Pos.X - LoopC To Pos.X + LoopC
            
112                 If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
114                     nPos.X = tX
116                     nPos.Y = tY
                        '¿Hay objeto?
                
118                     tX = Pos.X + LoopC
120                     tY = Pos.Y + LoopC
  
                    End If
        
122             Next tX
124         Next tY
    
126         LoopC = LoopC + 1
    
        Loop

128     If Notfound = True Then
130         nPos.X = 0
132         nPos.Y = 0

        End If

        
        Exit Sub

ClosestStablePos_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.ClosestStablePos", Erl)
        Resume Next
        
End Sub

Function NameIndex(ByVal name As String) As Integer
        
        On Error GoTo NameIndex_Err
        

        Dim UserIndex As Integer

        '¿Nombre valido?
100     If LenB(name) = 0 Then
102         NameIndex = 0
            Exit Function

        End If

104     If InStrB(name, "+") <> 0 Then
106         name = UCase$(Replace(name, "+", " "))

        End If

108     UserIndex = 1

110     Do Until UCase$(UserList(UserIndex).name) = UCase$(name)
    
112         UserIndex = UserIndex + 1
    
114         If UserIndex > MaxUsers Then
116             NameIndex = 0
                Exit Function

            End If
    
        Loop
 
118     NameIndex = UserIndex
 
        
        Exit Function

NameIndex_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.NameIndex", Erl)
        Resume Next
        
End Function

Function IP_Index(ByVal inIP As String) As Integer
        
        On Error GoTo IP_Index_Err
        
 
        Dim UserIndex As Integer

        '¿Nombre valido?
100     If LenB(inIP) = 0 Then
102         IP_Index = 0
            Exit Function

        End If
  
104     UserIndex = 1

106     Do Until UserList(UserIndex).ip = inIP
    
108         UserIndex = UserIndex + 1
    
110         If UserIndex > MaxUsers Then
112             IP_Index = 0
                Exit Function

            End If
    
        Loop
 
114     IP_Index = UserIndex

        Exit Function

        
        Exit Function

IP_Index_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.IP_Index", Erl)
        Resume Next
        
End Function

Function ContarMismaIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Integer
        
        On Error GoTo CheckForSameIP_Err
        

        Dim LoopC As Integer

100     For LoopC = 1 To MaxUsers

102         If UserList(LoopC).flags.UserLogged = True Then
104             If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
106                 ContarMismaIP = ContarMismaIP + 1
                End If

            End If

108     Next LoopC

        
        Exit Function

CheckForSameIP_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.MaxConexionesIP", Erl)
        Resume Next
        
End Function

Function CheckForSameName(ByVal name As String) As Boolean
        
        On Error GoTo CheckForSameName_Err
        

        'Controlo que no existan usuarios con el mismo nombre
        Dim LoopC As Long

100     For LoopC = 1 To LastUser

102         If UserList(LoopC).flags.UserLogged Then
        
                'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
                'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
                'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
                'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
                'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
        
104             If UCase$(UserList(LoopC).name) = UCase$(name) Then
106                 CheckForSameName = True
                    Exit Function

                End If

            End If

108     Next LoopC

110     CheckForSameName = False

        
        Exit Function

CheckForSameName_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.CheckForSameName", Erl)
        Resume Next
        
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
        
        On Error GoTo HeadtoPos_Err
        

        '*****************************************************************
        'Toma una posicion y se mueve hacia donde esta perfilado
        '*****************************************************************
        Dim X  As Integer

        Dim Y  As Integer

        Dim nX As Integer

        Dim nY As Integer

100     X = Pos.X
102     Y = Pos.Y

104     If Head = eHeading.NORTH Then
106         nX = X
108         nY = Y - 1

        End If

110     If Head = eHeading.SOUTH Then
112         nX = X
114         nY = Y + 1

        End If

116     If Head = eHeading.EAST Then
118         nX = X + 1
120         nY = Y

        End If

122     If Head = eHeading.WEST Then
124         nX = X - 1
126         nY = Y

        End If

        'Devuelve valores
128     Pos.X = nX
130     Pos.Y = nY

        
        Exit Sub

HeadtoPos_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.HeadtoPos", Erl)
        Resume Next
        
End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal Montado As Boolean = False, Optional ByVal PuedeTraslado As Boolean = True) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Checks if the position is Legal.
        '***************************************************
        '¿Es un mapa valido?
        
        On Error GoTo LegalPos_Err
        

100     If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
102         LegalPos = False
        Else

104         If PuedeAgua And PuedeTierra Then
106             LegalPos = (MapData(Map, X, Y).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (PuedeTraslado Or MapData(Map, X, Y).TileExit.Map = 0)

108         ElseIf PuedeTierra And Not PuedeAgua Then
110             LegalPos = (MapData(Map, X, Y).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And ((MapData(Map, X, Y).Blocked And FLAG_AGUA) = 0) And (PuedeTraslado Or MapData(Map, X, Y).TileExit.Map = 0)

112         ElseIf PuedeAgua And Not PuedeTierra Then
114             LegalPos = (MapData(Map, X, Y).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And ((MapData(Map, X, Y).Blocked And FLAG_AGUA) <> 0) And (PuedeTraslado Or MapData(Map, X, Y).TileExit.Map = 0)
            Else
116             LegalPos = False

            End If
   
        End If

        
        Exit Function

LegalPos_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.LegalPos", Erl)
        Resume Next
        
End Function

Function LegalWalk(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As eHeading, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal Montado As Boolean = False, Optional ByVal PuedeTraslado As Boolean = True) As Boolean
        On Error GoTo LegalPos_Err
        

100     If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
102         Exit Function
        End If

104     If PuedeAgua And PuedeTierra Then
106         LegalWalk = (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (PuedeTraslado Or MapData(Map, X, Y).TileExit.Map = 0)

108     ElseIf PuedeTierra And Not PuedeAgua Then
110         LegalWalk = (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And ((MapData(Map, X, Y).Blocked And FLAG_AGUA) = 0) And (PuedeTraslado Or MapData(Map, X, Y).TileExit.Map = 0)

112     ElseIf PuedeAgua And Not PuedeTierra Then
114         LegalWalk = (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And ((MapData(Map, X, Y).Blocked And FLAG_AGUA) <> 0) And (PuedeTraslado Or MapData(Map, X, Y).TileExit.Map = 0)
        
        Else
116         LegalWalk = False
        End If
        
        LegalWalk = LegalWalk And ((MapData(Map, X, Y).Blocked And 2 ^ (Heading - 1)) = 0)
        
        Exit Function

LegalPos_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.LegalWalk", Erl)
        Resume Next
        
End Function

Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean
        
        On Error GoTo LegalPosNPC_Err
        

100     If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
102         LegalPosNPC = False
        
        ElseIf MapData(Map, X, Y).TileExit.Map > 0 Then
            LegalPosNPC = False
        
        Else

104         If AguaValida = 0 Then
106             LegalPosNPC = (MapData(Map, X, Y).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (MapData(Map, X, Y).trigger <> eTrigger.POSINVALIDA Or IsPet) And (MapData(Map, X, Y).Blocked And FLAG_AGUA) = 0
            Else
108             LegalPosNPC = (MapData(Map, X, Y).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (MapData(Map, X, Y).trigger <> eTrigger.POSINVALIDA Or IsPet)
            End If
 
        End If

        
        Exit Function

LegalPosNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.LegalPosNPC", Erl)
        Resume Next
        
End Function

Sub SendHelp(ByVal Index As Integer)
        
        On Error GoTo SendHelp_Err
        

        Dim NumHelpLines As Integer

        Dim LoopC        As Integer

100     NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

102     For LoopC = 1 To NumHelpLines
104         Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
106     Next LoopC

        
        Exit Sub

SendHelp_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.SendHelp", Erl)
        Resume Next
        
End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo Expresar_Err
        

100     If Npclist(NpcIndex).NroExpresiones > 0 Then

            Dim randomi

102         randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
104         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))

        End If

        
        Exit Sub

Expresar_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.Expresar", Erl)
        Resume Next
        
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo LookatTile_Err
        

        'Responde al click del usuario sobre el mapa
        Dim FoundChar      As Byte

        Dim FoundSomething As Byte

        Dim TempCharIndex  As Integer

        Dim Stat           As String

        Dim ft             As FontTypeNames

        '¿Rango Visión? (ToxicWaste)
100     If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
            Exit Sub

        End If

        '¿Posicion valida?
102     If InMapBounds(Map, X, Y) Then
104         UserList(UserIndex).flags.TargetMap = Map
106         UserList(UserIndex).flags.TargetX = X
108         UserList(UserIndex).flags.TargetY = Y

            '¿Es un obj?
110         If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                'Informa el nombre
112             UserList(UserIndex).flags.TargetObjMap = Map
114             UserList(UserIndex).flags.TargetObjX = X
116             UserList(UserIndex).flags.TargetObjY = Y
118             FoundSomething = 1
120         ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then

                'Informa el nombre
122             If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
124                 UserList(UserIndex).flags.TargetObjMap = Map
126                 UserList(UserIndex).flags.TargetObjX = X + 1
128                 UserList(UserIndex).flags.TargetObjY = Y
130                 FoundSomething = 1

                End If

132         ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then

134             If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                    'Informa el nombre
136                 UserList(UserIndex).flags.TargetObjMap = Map
138                 UserList(UserIndex).flags.TargetObjX = X + 1
140                 UserList(UserIndex).flags.TargetObjY = Y + 1
142                 FoundSomething = 1

                End If

144         ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then

146             If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                    'Informa el nombre
148                 UserList(UserIndex).flags.TargetObjMap = Map
150                 UserList(UserIndex).flags.TargetObjX = X
152                 UserList(UserIndex).flags.TargetObjY = Y + 1
154                 FoundSomething = 1

                End If

            End If
    
156         If FoundSomething = 1 Then
158             UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex

160             If UserList(UserIndex).Counters.Trabajando = 0 Then
162                 If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then

164                     Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "* - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.Amount & "", FontTypeNames.FONTTYPE_INFO)
            
                    Else

166                     If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = eOBJType.otYacimiento Then
168                         Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
170                         Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).name & " - (Minerales disponibles: " & MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.Amount & ")", FontTypeNames.FONTTYPE_INFO)

172                     ElseIf ObjData(UserList(UserIndex).flags.TargetObj).OBJType = eOBJType.otArboles Then
174                         Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
176                         Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).name & " - (Recursos disponibles: " & MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.Amount & ")", FontTypeNames.FONTTYPE_INFO)
                    
                        Else
178                         Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "*", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                End If
    
            End If

            '¿Es un personaje?
180         If Y + 1 <= YMaxMapSize Then
182             If MapData(Map, X, Y + 1).UserIndex > 0 Then
184                 TempCharIndex = MapData(Map, X, Y + 1).UserIndex
186                 FoundChar = 1

                End If

188             If MapData(Map, X, Y + 1).NpcIndex > 0 Then
190                 TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
192                 FoundChar = 2

                End If

            End If

            '¿Es un personaje?
194         If FoundChar = 0 Then
196             If MapData(Map, X, Y).UserIndex > 0 Then
198                 TempCharIndex = MapData(Map, X, Y).UserIndex
200                 FoundChar = 1

                End If

202             If MapData(Map, X, Y).NpcIndex > 0 Then
204                 TempCharIndex = MapData(Map, X, Y).NpcIndex
206                 FoundChar = 2

                End If

            End If
    
            'Reaccion al personaje
208         If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
210             If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios And PlayerType.Dios Then
            
                    'If LenB(UserList(TempCharIndex).DescRM) = 0 Then 'No tiene descRM y quiere que se vea su nombre.
                
212                 If UserList(TempCharIndex).flags.Privilegios = user Then
                
                        Dim Fragsnick As String

                        'If Abs(CInt(UserList(TempCharIndex).Stats.ELV) - CInt(UserList(UserIndex).Stats.ELV)) < 10 Then
214                     Stat = Stat & " <" & ListaClases(UserList(TempCharIndex).clase) & " " & ListaRazas(UserList(TempCharIndex).raza) & " Nivel: " & UserList(TempCharIndex).Stats.ELV & ">"

                        'End If
216                     If EsNewbie(TempCharIndex) Then
218                         Stat = Stat & " <Newbie>"

                        End If

220                     If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 49 Then
222                         If UserList(TempCharIndex).flags.Envenenado > 0 Then
224                             Fragsnick = " | Envenenado "

                            End If

226                         If UserList(TempCharIndex).flags.Ceguera = 1 Then
228                             Fragsnick = Fragsnick & " | Ciego "

                            End If

230                         If UserList(TempCharIndex).flags.Incinerado = 1 Then
232                             Fragsnick = Fragsnick & " | Incinerado "

                            End If

234                         If UserList(TempCharIndex).flags.Paralizado = 1 Then
236                             Fragsnick = Fragsnick & " | Paralizado "

                            End If

238                         If UserList(TempCharIndex).flags.Inmovilizado = 1 Then
240                             Fragsnick = Fragsnick & " | Inmovilizado "

                            End If

242                         If UserList(TempCharIndex).Counters.Trabajando > 0 Then
244                             Fragsnick = Fragsnick & " | Trabajando "

                            End If

246                         If UserList(TempCharIndex).flags.invisible = 1 Then
248                             Fragsnick = Fragsnick & " | Invisible "

                            End If

250                         If UserList(TempCharIndex).flags.Oculto = 1 Then
252                             Fragsnick = Fragsnick & " | Oculto "

                            End If

254                         If UserList(TempCharIndex).flags.Estupidez = 1 Then
256                             Fragsnick = Fragsnick & " | Estupido "

                            End If

258                         If UserList(TempCharIndex).flags.Maldicion = 1 Then
260                             Fragsnick = Fragsnick & " | Maldito "

                            End If

262                         If UserList(TempCharIndex).flags.Silenciado = 1 Then
264                             Fragsnick = Fragsnick & " | Silenciado "

                            End If

266                         If UserList(TempCharIndex).flags.Comerciando = True Then
268                             Fragsnick = Fragsnick & " | Comerciando "

                            End If

270                         If UserList(TempCharIndex).flags.Descansar = 1 Then
272                             Fragsnick = Fragsnick & " | Descansando "

                            End If

274                         If UserList(TempCharIndex).flags.Meditando = True Then
276                             Fragsnick = Fragsnick & " | Concentrado "
            
                            End If

278                         If UserList(TempCharIndex).flags.BattleModo = 1 Then
280                             Fragsnick = Fragsnick & " | Modo Battle "
            
                            End If
                        
282                         If UserList(TempCharIndex).Stats.MinHp = 0 Then
284                             Stat = Stat & " <Muerto>"
286                         ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.1) Then
288                             Stat = Stat & " <Casi muerto" & Fragsnick & ">"
290                         ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.5) Then
292                             Stat = Stat & " <Malherido" & Fragsnick & ">"
294                         ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.75) Then
296                             Stat = Stat & " <Herido" & Fragsnick & ">"
298                         ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.99) Then
300                             Stat = Stat & " <Levemente herido" & Fragsnick & ">"
                            Else
302                             Stat = Stat & " <Intacto" & Fragsnick & ">"

                            End If

                        End If
                
304                     If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
306                         Stat = Stat & " <" & TituloReal(TempCharIndex) & ">"
308                         ft = FontTypeNames.FONTTYPE_CONSEJOVesA
310                     ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
312                         Stat = Stat & " <" & TituloCaos(TempCharIndex) & ">"
314                         ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA

                        End If
                
316                     If UserList(TempCharIndex).GuildIndex > 0 Then
318                         Stat = Stat & " <" & modGuilds.GuildName(UserList(TempCharIndex).GuildIndex) & ">"

                        End If

                    End If ' If user > 0 then
                            
320                 If Not UserList(TempCharIndex).flags.Privilegios And PlayerType.user Then
322                     If UserList(TempCharIndex).flags.Privilegios = Consejero Then
324                         Stat = Stat & " <Game Desing>"
326                         ft = FontTypeNames.FONTTYPE_GM

                        End If

328                     If UserList(TempCharIndex).flags.Privilegios = SemiDios Then
330                         Stat = Stat & " <Game Master>"
332                         ft = FontTypeNames.FONTTYPE_GM

                        End If

334                     If UserList(TempCharIndex).flags.Privilegios = Dios Then
336                         Stat = Stat & " <Administrador>"
338                         ft = FontTypeNames.FONTTYPE_DIOS

                        End If
                        
340                     If UserList(TempCharIndex).flags.Privilegios = PlayerType.Admin Then
342                         Stat = Stat & " <Administrador>"
344                         ft = FontTypeNames.FONTTYPE_DIOS

                        End If
                    
346                 ElseIf UserList(TempCharIndex).Faccion.Status = 0 Then
348                     ft = FontTypeNames.FONTTYPE_New_Gris
350                 ElseIf UserList(TempCharIndex).Faccion.Status = 1 Then
352                     ft = FontTypeNames.FONTTYPE_CITIZEN

                    End If
                    
354                 If UserList(TempCharIndex).flags.Casado = 1 Then
356                     Stat = Stat & " <Pareja de " & UserList(TempCharIndex).flags.Pareja & ">"

                    End If
                    
358                 If Len(UserList(TempCharIndex).Desc) > 0 Then
360                     Stat = "Ves a [" & UserList(TempCharIndex).name & "]" & Stat & " - " & UserList(TempCharIndex).Desc
                    Else
362                     Stat = "Ves a [" & UserList(TempCharIndex).name & "]" & Stat

                    End If
                 
                    ' Else  'Si tiene descRM la muestro siempre.
                    '   Stat = UserList(TempCharIndex).DescRM
                    '   ft = FontTypeNames.FONTTYPE_INFOBOLD
                    ' End If
            
364                 If LenB(Stat) > 0 Then
366                     Call WriteConsoleMsg(UserIndex, Stat, ft)

                    End If
            
368                 FoundSomething = 1
370                 UserList(UserIndex).flags.TargetUser = TempCharIndex
372                 UserList(UserIndex).flags.TargetNPC = 0
374                 UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun

                End If

            End If

376         If FoundChar = 2 Then '¿Encontro un NPC?

                Dim estatus As String
            
                'If UserList(UserIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                '  estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
                '  Else
                        
378             If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 75 Then
380                 If Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.1) Then
382                     estatus = "<Agonizando"
384                 ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.2) Then
386                     estatus = "<Casi muerto"
388                 ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.5) Then
390                     estatus = "<Malherido"
392                 ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.7) Then
394                     estatus = "<Herido"
396                 ElseIf Npclist(TempCharIndex).Stats.MinHp < (Npclist(TempCharIndex).Stats.MaxHp * 0.95) Then
398                     estatus = "<Levemente herido"
                    Else
400                     estatus = "<Intacto"

                    End If

                Else
402                 estatus = "<" & Npclist(TempCharIndex).Stats.MinHp & "/" & Npclist(TempCharIndex).Stats.MaxHp
                        
                End If
                        
404             If Npclist(TempCharIndex).flags.Envenenado > 0 Then
406                 estatus = estatus & " | Envenenado"

                End If
                        
408             If Npclist(TempCharIndex).flags.Paralizado = 1 Then
410                 If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 49 Then
412                     estatus = estatus & " | Paralizado (" & CInt(Npclist(TempCharIndex).Contadores.Paralisis / 6.5) & "s)"
                    Else
414                     estatus = estatus & " | Paralizado"

                    End If

                End If
                        
416             If Npclist(TempCharIndex).flags.Inmovilizado = 1 Then
418                 If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 49 Then
420                     estatus = estatus & " | Inmovilizado (" & CInt(Npclist(TempCharIndex).Contadores.Paralisis / 6.5) & "s)"
                    Else
422                     estatus = estatus & " | Inmovilizado"

                    End If

                End If
                        
424             estatus = estatus & ">"
    
                'End If
            
426             If Len(Npclist(TempCharIndex).Desc) > 1 Then
                    'Optimizacion de protocolo por Ladder
428                 Call WriteChatOverHead(UserIndex, "NPCDESC*" & Npclist(TempCharIndex).Numero, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
430             ElseIf TempCharIndex = CentinelaNPCIndex Then
                    'Enviamos nuevamente el texto del centinela según quien pregunta
432                 Call modCentinela.CentinelaSendClave(UserIndex)
                
                ElseIf Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call WriteConsoleMsg(UserIndex, "NPCNAME*" & Npclist(TempCharIndex).Numero & "* es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).name & " " & estatus, FontTypeNames.FONTTYPE_INFO)
                
                Else
                
434                 Call WriteConsoleMsg(UserIndex, "NPCNAME*" & Npclist(TempCharIndex).Numero & "*" & " " & estatus, FontTypeNames.FONTTYPE_INFO)
                    ' If UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                    ' Call WriteConsoleMsg(UserIndex, "Le pegó primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, Npclist(TempCharIndex).Char.CharIndex, FontTypeNames.FONTTYPE_INFO)
                    'End If
                
                End If

436             FoundSomething = 1
438             UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
440             UserList(UserIndex).flags.TargetNPC = TempCharIndex
442             UserList(UserIndex).flags.TargetUser = 0
444             UserList(UserIndex).flags.TargetObj = 0
        
            End If
    
446         If FoundChar = 0 Then
448             UserList(UserIndex).flags.TargetNPC = 0
450             UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
452             UserList(UserIndex).flags.TargetUser = 0

            End If
    
            '*** NO ENCOTRO NADA ***
454         If FoundSomething = 0 Then
456             UserList(UserIndex).flags.TargetNPC = 0
458             UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
460             UserList(UserIndex).flags.TargetUser = 0
462             UserList(UserIndex).flags.TargetObj = 0
464             UserList(UserIndex).flags.TargetObjMap = 0
466             UserList(UserIndex).flags.TargetObjX = 0
468             UserList(UserIndex).flags.TargetObjY = 0

                ' Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", FontTypeNames.FONTTYPE_INFO)
            End If

        Else

470         If FoundSomething = 0 Then
472             UserList(UserIndex).flags.TargetNPC = 0
474             UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
476             UserList(UserIndex).flags.TargetUser = 0
478             UserList(UserIndex).flags.TargetObj = 0
480             UserList(UserIndex).flags.TargetObjMap = 0
482             UserList(UserIndex).flags.TargetObjX = 0
484             UserList(UserIndex).flags.TargetObjY = 0

                '  Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", FontTypeNames.FONTTYPE_INFO)
            End If

        End If

        
        Exit Sub

LookatTile_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.LookatTile", Erl)
        Resume Next
        
End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
        
        On Error GoTo FindDirection_Err
        

        '*****************************************************************
        'Devuelve la direccion en la cual el target se encuentra
        'desde pos, 0 si la direc es igual
        '*****************************************************************
        Dim X As Integer

        Dim Y As Integer

100     X = Pos.X - Target.X
102     Y = Pos.Y - Target.Y

        'NE
104     If Sgn(X) = -1 And Sgn(Y) = 1 Then
106         FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
            Exit Function

        End If

        'NW
108     If Sgn(X) = 1 And Sgn(Y) = 1 Then
110         FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
            Exit Function

        End If

        'SW
112     If Sgn(X) = 1 And Sgn(Y) = -1 Then
114         FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
            Exit Function

        End If

        'SE
116     If Sgn(X) = -1 And Sgn(Y) = -1 Then
118         FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
            Exit Function

        End If

        'Sur
120     If Sgn(X) = 0 And Sgn(Y) = -1 Then
122         FindDirection = eHeading.SOUTH
            Exit Function

        End If

        'norte
124     If Sgn(X) = 0 And Sgn(Y) = 1 Then
126         FindDirection = eHeading.NORTH
            Exit Function

        End If

        'oeste
128     If Sgn(X) = 1 And Sgn(Y) = 0 Then
130         FindDirection = eHeading.WEST
            Exit Function

        End If

        'este
132     If Sgn(X) = -1 And Sgn(Y) = 0 Then
134         FindDirection = eHeading.EAST
            Exit Function

        End If

        'misma
136     If Sgn(X) = 0 And Sgn(Y) = 0 Then
138         FindDirection = 0
            Exit Function

        End If

        
        Exit Function

FindDirection_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.FindDirection", Erl)
        Resume Next
        
End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean
        
        On Error GoTo ItemNoEsDeMapa_Err
        

100     ItemNoEsDeMapa = ObjData(Index).OBJType <> eOBJType.otPuertas And ObjData(Index).OBJType <> eOBJType.otForos And ObjData(Index).OBJType <> eOBJType.otCarteles And ObjData(Index).OBJType <> eOBJType.otArboles And ObjData(Index).OBJType <> eOBJType.otYacimiento And ObjData(Index).OBJType <> eOBJType.otTeleport And ObjData(Index).OBJType <> eOBJType.OtCorreo And ObjData(Index).OBJType <> eOBJType.OtDecoraciones

        
        Exit Function

ItemNoEsDeMapa_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.ItemNoEsDeMapa", Erl)
        Resume Next
        
End Function

'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
        
        On Error GoTo MostrarCantidad_Err
        
100     MostrarCantidad = ObjData(Index).OBJType <> eOBJType.otPuertas And ObjData(Index).OBJType <> eOBJType.otForos And ObjData(Index).OBJType <> eOBJType.otCarteles And ObjData(Index).OBJType <> eOBJType.otYacimiento And ObjData(Index).OBJType <> eOBJType.otArboles And ObjData(Index).OBJType <> eOBJType.OtCorreo And ObjData(Index).OBJType <> eOBJType.otTeleport

        
        Exit Function

MostrarCantidad_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.MostrarCantidad", Erl)
        Resume Next
        
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean
        
        On Error GoTo EsObjetoFijo_Err
        

100     EsObjetoFijo = OBJType = eOBJType.otForos Or OBJType = eOBJType.otCarteles Or OBJType = eOBJType.otArboles Or OBJType = eOBJType.otYacimiento Or OBJType = eOBJType.OtDecoraciones

        
        Exit Function

EsObjetoFijo_Err:
        Call RegistrarError(Err.Number, Err.description, "Extra.EsObjetoFijo", Erl)
        Resume Next
        
End Function
