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

Public m_NameIndex As New Dictionary

Public Function NameIndex(ByRef UserName As String) As Integer

100     NameIndex = m_NameIndex(UCase$(Replace(UserName, "+", " ")))
    
End Function


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
    
104         For Rango = 0 To 5
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
140                         Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", e_FontTypeNames.FONTTYPE_TALK)
                        

                        End If

                        'Lo sacamos.
142                     If UserList(OtherUserIndex).flags.UserLogged Then
144                         Call FinComerciarUsu(OtherUserIndex)
146                         Call WriteShowMessageBox(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        

                        End If

                    End If
            
148                 Call CloseSocket(OtherUserIndex)

                End If

            End If

        End If

        
        Exit Sub

FindLegalPos_Err:
150     Call TraceError(Err.Number, Err.Description, "Extra.FindLegalPos", Erl)

        
End Sub

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
        On Error GoTo EsNewbie_Err
        
100     If UserIndex > 0 Then EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie

        Exit Function

EsNewbie_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.EsNewbie", Erl)

        
End Function

Public Function esCiudadano(ByVal UserIndex As Integer) As Boolean
        On Error GoTo esCiudadano_Err
        
100     If UserIndex > 0 Then esCiudadano = Status(UserIndex) = Ciudadano
        
        Exit Function

esCiudadano_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.esCiudadano", Erl)

End Function

Public Function esCriminal(ByVal UserIndex As Integer) As Boolean
        On Error GoTo esCriminal_Err
        
100     If UserIndex > 0 Then esCriminal = Status(UserIndex) = Criminal
        
        Exit Function

esCriminal_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.esCriminal", Erl)

End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
        On Error GoTo esArmada_Err
        
100     If UserIndex > 0 Then esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)
        
        Exit Function

esArmada_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.esArmada", Erl)

        
End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean
        On Error GoTo esCaos_Err
        
100     If UserIndex > 0 Then esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
     
        Exit Function

esCaos_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.esCaos", Erl)

        
End Function

Public Function EsGM(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 23/01/2007
        '***************************************************
        
        On Error GoTo EsGM_Err
        
100     If UserIndex = 0 Then
102         EsGM = False
            Exit Function
        End If
        
104     EsGM = (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) <> 0
        
        Exit Function

EsGM_Err:
106     Call TraceError(Err.Number, Err.Description, "Extra.EsGM", Erl)

        
End Function

Private Function CheckMapRestrictions(ByVal UserIndex As Integer, ByVal Map As Integer) As Boolean

100     With UserList(UserIndex)

102         If EsGM(UserIndex) Then
104             CheckMapRestrictions = True
                Exit Function
            End If

106         If MapInfo(Map).Newbie And Not EsNewbie(UserIndex) Then
108             If .flags.UltimoMensaje <> 101 Then
110                 Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
112                 .flags.UltimoMensaje = 101
                End If
                Exit Function
            End If

114         If MapInfo(Map).NoPKs And (Status(UserIndex) = 0 Or Status(UserIndex) = 2) Then
116             If .flags.UltimoMensaje <> 102 Then
118                 Call WriteConsoleMsg(UserIndex, "Sólo los ciudadanos pueden entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
120                 .flags.UltimoMensaje = 102
                End If
                Exit Function
            End If

122         If MapInfo(Map).NoCiudadanos And (Status(UserIndex) = 1 Or Status(UserIndex) = 3) Then
124             If .flags.UltimoMensaje <> 103 Then
126                 Call WriteConsoleMsg(UserIndex, "Sólo los criminales pueden entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
128                 .flags.UltimoMensaje = 103
                End If
                Exit Function
            End If

130         If MapInfo(Map).SoloClanes And .GuildIndex <= 0 Then
132             If .flags.UltimoMensaje <> 104 Then
134                 Call WriteConsoleMsg(UserIndex, "Necesitas pertenecer a un clan para entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
136                 .flags.UltimoMensaje = 104
                End If
                Exit Function
            End If

138         If MapInfo(Map).MinLevel <> 0 And .Stats.ELV < MapInfo(Map).MinLevel Then
140             If .flags.UltimoMensaje <> 105 Then
142                 Call WriteConsoleMsg(UserIndex, "Necesitas ser al menos nivel " & MapInfo(Map).MinLevel & " para entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
144                 .flags.UltimoMensaje = 105
                End If
                Exit Function
            End If

146         If MapInfo(Map).MaxLevel <> 0 And .Stats.ELV >= MapInfo(Map).MaxLevel Then
148             If .flags.UltimoMensaje <> 106 Then
150                 Call WriteConsoleMsg(UserIndex, "Sólo los personajes inferiores a nivel " & MapInfo(Map).MaxLevel & " pueden entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
152                 .flags.UltimoMensaje = 106
                End If
                Exit Function
            End If

154         CheckMapRestrictions = True

        End With

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

        Dim nPos   As t_WorldPos

        Dim EsTeleport As Boolean
        
        Dim aN As Integer
        
100     With UserList(UserIndex)

            'Controla las salidas
102         If InMapBounds(Map, X, Y) Then
        
104             If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
106                 EsTeleport = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = e_OBJType.otTeleport
                End If
    
108             If (MapData(Map, X, Y).TileExit.Map > 0) And (MapData(Map, X, Y).TileExit.Map <= NumMaps) Then
    
                    ' WyroX: Restricciones de mapas
110                 If CheckMapRestrictions(UserIndex, MapData(Map, X, Y).TileExit.Map) Then
112                     If EsMapaInterdimensional(MapData(Map, X, Y).TileExit.Map) And Not EsMapaInterdimensional(.Pos.Map) Then
114                         .flags.ReturnPos = .Pos
                        End If

116                     If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, .flags.Navegando = 1, , , False) Then
118                         Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, EsTeleport)
                
                        Else
120                         Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
    
122                         If nPos.X <> 0 And nPos.Y <> 0 Then
124                             Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, EsTeleport)
                            End If
    
                        End If
                
                    ' Si hay un teleport: movemos al usuario para que no se quede bloqueándolo
126                 ElseIf EsTeleport Then
128                     Call ClosestLegalPos(.Pos, nPos)
    
130                      If nPos.X <> 0 And nPos.Y <> 0 Then
132                          Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, EsTeleport)
                        End If
                    End If
    
                    'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
134                 aN = .flags.AtacadoPorNpc
    
136                 If aN > 0 Then
138                     NpcList(aN).Movement = NpcList(aN).flags.OldMovement
140                     NpcList(aN).Hostile = NpcList(aN).flags.OldHostil
142                     NpcList(aN).flags.AttackedBy = vbNullString
144                     NpcList(aN).Target = 0
                    End If
        
146                 aN = .flags.NPCAtacado
    
148                 If aN > 0 Then
150                     If NpcList(aN).flags.AttackedFirstBy = .Name Then
152                         NpcList(aN).flags.AttackedFirstBy = vbNullString
    
                        End If
    
                    End If
    
154                 .flags.AtacadoPorNpc = 0
156                 .flags.NPCAtacado = 0
    
158             ElseIf MapData(Map, X, Y).TileExit.Map < 0 Then
160                 If .flags.ReturnPos.Map <> 0 Then
162                     If LegalPos(.flags.ReturnPos.Map, .flags.ReturnPos.X, .flags.ReturnPos.Y, .flags.Navegando = 1, , , False) Then
164                         Call WarpUserChar(UserIndex, .flags.ReturnPos.Map, .flags.ReturnPos.X, .flags.ReturnPos.Y, False)
                        
                        Else
166                         Call ClosestLegalPos(.flags.ReturnPos, nPos)
                        
168                         If nPos.X <> 0 And nPos.Y <> 0 Then
170                             Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, EsTeleport)
                            End If
                        End If
                        
                        'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
172                     aN = .flags.AtacadoPorNpc
        
174                     If aN > 0 Then
176                         NpcList(aN).Movement = NpcList(aN).flags.OldMovement
178                         NpcList(aN).Hostile = NpcList(aN).flags.OldHostil
180                         NpcList(aN).flags.AttackedBy = vbNullString
182                         NpcList(aN).Target = 0
                        End If
            
184                     aN = .flags.NPCAtacado
        
186                     If aN > 0 Then
188                         If NpcList(aN).flags.AttackedFirstBy = .Name Then
190                             NpcList(aN).flags.AttackedFirstBy = vbNullString
        
                            End If
        
                        End If
        
192                     .flags.AtacadoPorNpc = 0
194                     .flags.NPCAtacado = 0
                    End If
                End If
            End If

        End With

        Exit Sub

ErrHandler:
196     Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.Description)

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
108     Call TraceError(Err.Number, Err.Description, "Extra.InRangoVision", Erl)

        
End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean
        
        On Error GoTo InRangoVisionNPC_Err
        

100     If X > NpcList(NpcIndex).Pos.X - MinXBorder And X < NpcList(NpcIndex).Pos.X + MinXBorder Then
102         If Y > NpcList(NpcIndex).Pos.Y - MinYBorder And Y < NpcList(NpcIndex).Pos.Y + MinYBorder Then
104             InRangoVisionNPC = True
                Exit Function

            End If

        End If

106     InRangoVisionNPC = False

        
        Exit Function

InRangoVisionNPC_Err:
108     Call TraceError(Err.Number, Err.Description, "Extra.InRangoVisionNPC", Erl)

        
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
106     Call TraceError(Err.Number, Err.Description, "Extra.InMapBounds", Erl)

        
End Function

Function ClosestLegalPosNPC(ByVal NpcIndex As Integer, ByVal MaxRange As Integer, Optional ByVal IgnoreUsers As Boolean) As t_WorldPos

        On Error GoTo ErrHandler

        Dim LoopC    As Integer
        Dim tX       As Integer
        Dim tY       As Integer
    
100     With NpcList(NpcIndex)

            Do
102             tY = .Pos.Y - LoopC
104             For tX = .Pos.X - LoopC To .Pos.X + LoopC
106                 If ValidNPCSpawnPos(ClosestLegalPosNPC, .Pos.Map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers) Then
                        Exit Function
                    End If
                Next

108             tX = .Pos.X - LoopC
110             For tY = .Pos.Y - LoopC + 1 To .Pos.Y + LoopC - 1
112                 If ValidNPCSpawnPos(ClosestLegalPosNPC, .Pos.Map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers) Then
                        Exit Function
                    End If
                Next

114             tX = .Pos.X + LoopC
116             For tY = .Pos.Y - LoopC + 1 To .Pos.Y + LoopC - 1
118                 If ValidNPCSpawnPos(ClosestLegalPosNPC, .Pos.Map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers) Then
                        Exit Function
                    End If
                Next

120             tY = .Pos.Y + LoopC
122             For tX = .Pos.X - LoopC To .Pos.X + LoopC
124                 If ValidNPCSpawnPos(ClosestLegalPosNPC, .Pos.Map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers) Then
                        Exit Function
                    End If
                Next

126             LoopC = LoopC + 1

128         Loop While LoopC <= MaxRange

        End With

        Exit Function

ErrHandler:
130     Call TraceError(Err.Number, Err.Description, "Extra.ClosestLegalPosNPC")

        
End Function

Private Function ValidNPCSpawnPos(OutPos As t_WorldPos, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Boolean, ByVal TierraValida As Boolean, ByVal IgnoreUsers As Boolean) As Boolean

100     If LegalPos(Map, X, Y, AguaValida, TierraValida, , False) Then
102         If TestSpawnTrigger(Map, X, Y) Then
104             If Not HayPCarea(Map, X, Y) Or IgnoreUsers Then
106                 ValidNPCSpawnPos = True
108                 OutPos.Map = Map
110                 OutPos.X = X
112                 OutPos.Y = Y
                    Exit Function
                End If
            End If
        End If

End Function

Sub ClosestLegalPos(Pos As t_WorldPos, ByRef nPos As t_WorldPos, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True)
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
                        Exit Sub
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
134     Call TraceError(Err.Number, Err.Description, "Extra.ClosestLegalPos", Erl)

        
End Sub

Sub ClosestStablePos(Pos As t_WorldPos, ByRef nPos As t_WorldPos)
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
134     Call TraceError(Err.Number, Err.Description, "Extra.ClosestStablePos", Erl)

        
End Sub

Function IP_Index(ByVal inIP As String) As Integer
        
        On Error GoTo IP_Index_Err
        
 
        Dim UserIndex As Integer

        '¿Nombre valido?
100     If LenB(inIP) = 0 Then
102         IP_Index = 0
            Exit Function

        End If
  
104     UserIndex = 1

106     Do Until UserList(UserIndex).IP = inIP
    
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
116     Call TraceError(Err.Number, Err.Description, "Extra.IP_Index", Erl)

        
End Function

Function ContarMismaIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Integer
        
        On Error GoTo CheckForSameIP_Err
        

        Dim LoopC As Integer

100     For LoopC = 1 To MaxUsers

102         If UserList(LoopC).flags.UserLogged = True Then
104             If UserList(LoopC).IP = UserIP And UserIndex <> LoopC Then
106                 ContarMismaIP = ContarMismaIP + 1
                End If

            End If

108     Next LoopC

        
        Exit Function

CheckForSameIP_Err:
110     Call TraceError(Err.Number, Err.Description, "Extra.MaxConexionesIP", Erl)

        
End Function

Sub HeadtoPos(ByVal Head As e_Heading, ByRef Pos As t_WorldPos)
        
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

104     If Head = e_Heading.NORTH Then
106         nX = X
108         nY = Y - 1

        End If

110     If Head = e_Heading.SOUTH Then
112         nX = X
114         nY = Y + 1

        End If

116     If Head = e_Heading.EAST Then
118         nX = X + 1
120         nY = Y

        End If

122     If Head = e_Heading.WEST Then
124         nX = X - 1
126         nY = Y

        End If

        'Devuelve valores
128     Pos.X = nX
130     Pos.Y = nY

        
        Exit Sub

HeadtoPos_Err:
132     Call TraceError(Err.Number, Err.Description, "Extra.HeadtoPos", Erl)

        
End Sub

' Autor: WyroX - 20/01/2021
' Retorna el heading recibo como parámetro pero rotado, según el valor R.
' Si R es 1, rota en sentido horario. Si R es -1, en sentido antihorario.
Function Rotate_Heading(ByVal Heading As e_Heading, ByVal R As Integer) As e_Heading
    
100     Rotate_Heading = (Heading + R + 3) Mod 4 + 1
    
End Function

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal Montado As Boolean = False, Optional ByVal PuedeTraslado As Boolean = True, Optional ByVal PuedeBloqueoParcial As Boolean = True) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Checks if the position is Legal.
        '***************************************************
        '¿Es un mapa valido?
        
        On Error GoTo LegalPos_Err
        
100     If Map <= 0 Or Map > NumMaps Then Exit Function
        
102     If X < MinXBorder Or X > MaxXBorder Then Exit Function
        
104     If Y < MinYBorder Or Y > MaxYBorder Then Exit Function
        
106     With MapData(Map, X, Y)
        
108         If .NpcIndex <> 0 Then Exit Function

110         If .UserIndex <> 0 Then Exit Function
            
112         If Not PuedeTraslado Then
114             If .TileExit.Map > 0 Then Exit Function
            End If
            
116         If Not PuedeAgua Then
118             If (.Blocked And FLAG_AGUA) <> 0 Then Exit Function
            End If
            
120         If Not PuedeTierra Then
122             If (.Blocked And FLAG_AGUA) = 0 Then Exit Function
            End If
            
            If PuedeBloqueoParcial Then
124             If (.Blocked And e_Block.ALL_SIDES) = e_Block.ALL_SIDES Then Exit Function
            Else
                If (.Blocked And e_Block.ALL_SIDES) > 0 Then Exit Function
            End If
            
        End With
        
        
126     LegalPos = True

        Exit Function

LegalPos_Err:
128     Call TraceError(Err.Number, Err.Description, "Extra.LegalPos", Erl)
        
End Function

Function LegalWalk(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As e_Heading, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal Montado As Boolean = False, Optional ByVal PuedeTraslado As Boolean = True) As Boolean
        On Error GoTo LegalWalk_Err
        

100     If Map <= 0 Or Map > NumMaps Then Exit Function
        
102     If X < MinXBorder Or X > MaxXBorder Then Exit Function
        
104     If Y < MinYBorder Or Y > MaxYBorder Then Exit Function
        
106     With MapData(Map, X, Y)
        
108         If .NpcIndex <> 0 Then Exit Function

110         If .UserIndex <> 0 Then
112             If UserList(.UserIndex).flags.AdminInvisible = 0 And UserList(.UserIndex).flags.Muerto = 0 Then
                    Exit Function
                End If
            End If
            
114         If Not PuedeTraslado Then
116             If .TileExit.Map > 0 Then Exit Function
            End If
            
118         If Not PuedeAgua Then
120             If (.Blocked And FLAG_AGUA) <> 0 Then Exit Function
            End If
            
122         If Not PuedeTierra Then
124             If (.Blocked And FLAG_AGUA) = 0 Then Exit Function
            End If
            
126         If (.Blocked And 2 ^ (Heading - 1)) <> 0 Then Exit Function
        
        End With
        
128     LegalWalk = True

        
        Exit Function

LegalWalk_Err:
130     Call TraceError(Err.Number, Err.Description, "Extra.LegalWalk", Erl)

        
End Function

Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean
        
        On Error GoTo LegalPosNPC_Err
        

100     If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
102         LegalPosNPC = False
        
104     ElseIf MapData(Map, X, Y).TileExit.Map > 0 Then
106         LegalPosNPC = False
        
        Else

108         If AguaValida = 0 Then
110             LegalPosNPC = (MapData(Map, X, Y).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (MapData(Map, X, Y).trigger <> e_Trigger.POSINVALIDA Or IsPet) And (MapData(Map, X, Y).Blocked And FLAG_AGUA) = 0
            Else
112             LegalPosNPC = (MapData(Map, X, Y).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (MapData(Map, X, Y).trigger <> e_Trigger.POSINVALIDA Or IsPet)
            End If
 
        End If

        
        Exit Function

LegalPosNPC_Err:
114     Call TraceError(Err.Number, Err.Description, "Extra.LegalPosNPC", Erl)

        
End Function

Function LegalWalkNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As e_Heading, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal IgnoraInvalida As Boolean = False, Optional ByVal PuedePisar As Boolean, Optional ByVal esGuardia As Boolean = False) As Boolean
    ' Reescrito por WyroX

    On Error GoTo LegalWalkNPC_Err

100 If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        Exit Function
    End If

102 With MapData(Map, X, Y)

104     If .TileExit.Map Then Exit Function

106     If Not PuedeAgua Then
108         If .Blocked And FLAG_AGUA Then
                Exit Function
            End If
        End If

110     If Not PuedeTierra Then
112         If (.Blocked And FLAG_AGUA) = 0 Then
                Exit Function
            End If
        End If

114     If Not PuedePisar Then
116         If .UserIndex Then
118             If UserList(.UserIndex).flags.AdminInvisible = 0 And UserList(.UserIndex).flags.Muerto = 0 Then
                    Exit Function
                End If
120         ElseIf .NpcIndex Then
                Exit Function
            End If
        End If
        
122     If Not IgnoraInvalida Then
124         If .trigger = e_Trigger.POSINVALIDA Then
                Exit Function
            End If
        End If
        
126     If Not esGuardia Then
128         If .Blocked And 2 ^ (Heading - 1) Then
                Exit Function
            End If
        Else
130         If (.Blocked And 2 ^ (Heading - 1)) And Not HayPuerta(Map, X + 1, Y) And Not HayPuerta(Map, X, Y) And Not HayPuerta(Map, X + 1, Y - 1) And Not HayPuerta(Map, X, Y - 1) Then Exit Function
        End If


    End With
    
132 LegalWalkNPC = True
    
    Exit Function

LegalWalkNPC_Err:
134     Call TraceError(Err.Number, Err.Description, "Extra.LegalWalkNPC", Erl)

        
End Function

Sub SendHelp(ByVal Index As Integer)
        
        On Error GoTo SendHelp_Err
        

        Dim NumHelpLines As Integer

        Dim LoopC        As Integer

100     NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

102     For LoopC = 1 To NumHelpLines
104         Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), e_FontTypeNames.FONTTYPE_INFO)
106     Next LoopC

        
        Exit Sub

SendHelp_Err:
108     Call TraceError(Err.Number, Err.Description, "Extra.SendHelp", Erl)

        
End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo Expresar_Err
        

100     If NpcList(NpcIndex).NroExpresiones > 0 Then

            Dim randomi

102         randomi = RandomNumber(1, NpcList(NpcIndex).NroExpresiones)
104         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(NpcList(NpcIndex).Expresiones(randomi), NpcList(NpcIndex).Char.CharIndex, vbWhite))

        End If

        
        Exit Sub

Expresar_Err:
106     Call TraceError(Err.Number, Err.Description, "Extra.Expresar", Erl)

        
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo LookatTile_Err
        

        'Responde al click del usuario sobre el mapa
        Dim FoundChar      As Byte

        Dim FoundSomething As Byte

        Dim TempCharIndex  As Integer

        Dim Stat           As String

        Dim ft             As e_FontTypeNames

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
122             If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = e_OBJType.otPuertas Then
124                 UserList(UserIndex).flags.TargetObjMap = Map
126                 UserList(UserIndex).flags.TargetObjX = X + 1
128                 UserList(UserIndex).flags.TargetObjY = Y
130                 FoundSomething = 1

                End If

132         ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then

134             If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = e_OBJType.otPuertas Then
                    'Informa el nombre
136                 UserList(UserIndex).flags.TargetObjMap = Map
138                 UserList(UserIndex).flags.TargetObjX = X + 1
140                 UserList(UserIndex).flags.TargetObjY = Y + 1
142                 FoundSomething = 1

                End If

144         ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then

146             If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = e_OBJType.otPuertas Then
                    'Informa el nombre
148                 UserList(UserIndex).flags.TargetObjMap = Map
150                 UserList(UserIndex).flags.TargetObjX = X
152                 UserList(UserIndex).flags.TargetObjY = Y + 1
154                 FoundSomething = 1

                End If

            End If
    
156         If FoundSomething = 1 Then
158             UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex

160             If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then

162                 Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "* - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount & "", e_FontTypeNames.FONTTYPE_INFO)
            
                Else

164                 If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = e_OBJType.otYacimiento Then
166                     Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
168                     Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).Name & " - (Minerales disponibles: " & MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount & ")", e_FontTypeNames.FONTTYPE_INFO)

170                 ElseIf ObjData(UserList(UserIndex).flags.TargetObj).OBJType = e_OBJType.otArboles Then
172                     Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
174                     Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).Name & " - (Recursos disponibles: " & MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount & ")", e_FontTypeNames.FONTTYPE_INFO)
                    
176                 ElseIf ObjData(UserList(UserIndex).flags.TargetObj).OBJType = e_OBJType.otTeleport Then
178                     If MapData(Map, X, Y).TileExit.Map > 0 Then
180                         If LenB(MapInfo(MapData(Map, X, Y).TileExit.Map).map_name) <> 0 Then
182                             Call WriteConsoleMsg(UserIndex, "Portal a " & MapInfo(MapData(Map, X, Y).TileExit.Map).map_name, e_FontTypeNames.FONTTYPE_INFO)
                            Else
184                             Call WriteConsoleMsg(UserIndex, "Portal a un mapa desconocido...", e_FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
186                         Call WriteConsoleMsg(UserIndex, "Portal hacia la quinta dimensión", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
188                     Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "*", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
    
            End If

            '¿Es un personaje?
190         If Y + 1 <= YMaxMapSize Then
192             If MapData(Map, X, Y + 1).UserIndex > 0 Then
194                 TempCharIndex = MapData(Map, X, Y + 1).UserIndex
196                 FoundChar = 1

                End If

198             If MapData(Map, X, Y + 1).NpcIndex > 0 Then
200                 TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
202                 FoundChar = 2

                End If

            End If

            '¿Es un personaje?
204         If FoundChar = 0 Then
206             If MapData(Map, X, Y).UserIndex > 0 Then
208                 TempCharIndex = MapData(Map, X, Y).UserIndex
210                 FoundChar = 1

                End If

212             If MapData(Map, X, Y).NpcIndex > 0 Then
214                 TempCharIndex = MapData(Map, X, Y).NpcIndex
216                 FoundChar = 2

                End If

            End If
    
            'Reaccion al personaje
218         If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
220             If UserList(TempCharIndex).flags.AdminInvisible = 0 Or CompararPrivilegiosUser(UserIndex, TempCharIndex) >= 0 Then
                    
222                 If UserList(TempCharIndex).showName Or CompararPrivilegiosUser(UserIndex, TempCharIndex) >= 0 Then
                
224                     If UserList(TempCharIndex).flags.Privilegios = user Then
                    
                            Dim Fragsnick As String
    
226                         If EsGM(UserIndex) Then
228                             Stat = Stat & " <" & ListaClases(UserList(TempCharIndex).clase) & " " & ListaRazas(UserList(TempCharIndex).raza) & " Nivel: " & UserList(TempCharIndex).Stats.ELV & ">"
                                Stat = Stat & " (ELO " & UserList(TempCharIndex).Stats.ELO & ")"
                            End If

232                         If EsNewbie(TempCharIndex) Then
234                             Stat = Stat & " <Newbie>"
                            End If
    
236                         If UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) > 49 Then
238                             If UserList(TempCharIndex).flags.Envenenado > 0 Then
240                                 Fragsnick = " | Envenenado"
                                End If
    
242                             If UserList(TempCharIndex).flags.Ceguera = 1 Then
244                                 Fragsnick = Fragsnick & " | Ciego"
                                End If
    
246                             If UserList(TempCharIndex).flags.Incinerado = 1 Then
248                                 Fragsnick = Fragsnick & " | Incinerado"
                                End If
    
250                             If UserList(TempCharIndex).flags.Paralizado = 1 Then
252                                 Fragsnick = Fragsnick & " | Paralizado"
                                End If
    
254                             If UserList(TempCharIndex).flags.Inmovilizado = 1 Then
256                                 Fragsnick = Fragsnick & " | Inmovilizado"
                                End If
    
258                             If UserList(TempCharIndex).Counters.Trabajando > 0 Then
260                                 Fragsnick = Fragsnick & " | Trabajando"
                                End If
    
262                             If UserList(TempCharIndex).flags.invisible = 1 Then
264                                 Fragsnick = Fragsnick & " | Invisible"
                                End If
    
266                             If UserList(TempCharIndex).flags.Oculto = 1 Then
268                                 Fragsnick = Fragsnick & " | Oculto"
                                End If
    
270                             If UserList(TempCharIndex).flags.Estupidez = 1 Then
272                                 Fragsnick = Fragsnick & " | Estupido"
                                End If
    
274                             If UserList(TempCharIndex).flags.Maldicion = 1 Then
276                                 Fragsnick = Fragsnick & " | Maldito"
                                End If
    
278                             If UserList(TempCharIndex).flags.Silenciado = 1 Then
280                                 Fragsnick = Fragsnick & " | Silenciado"
                                End If
    
282                             If UserList(TempCharIndex).flags.Comerciando = True Then
284                                 Fragsnick = Fragsnick & " | Comerciando"
                                End If
    
286                             If UserList(TempCharIndex).flags.Descansar = 1 Then
288                                 Fragsnick = Fragsnick & " | Descansando"
                                End If
    
290                             If UserList(TempCharIndex).flags.Meditando Then
292                                 Fragsnick = Fragsnick & " | Concentrado"
                                End If
                                
294                             If UserList(TempCharIndex).Stats.MinHp = 0 Then
296                                 Stat = Stat & " <Muerto>"
298                             ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.1) Then
300                                 Stat = Stat & " <Casi muerto" & Fragsnick & ">"
302                             ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.5) Then
304                                 Stat = Stat & " <Malherido" & Fragsnick & ">"
306                             ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.75) Then
308                                 Stat = Stat & " <Herido" & Fragsnick & ">"
310                             ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.99) Then
312                                 Stat = Stat & " <Levemente herido" & Fragsnick & ">"
                                Else
314                                 Stat = Stat & " <Intacto" & Fragsnick & ">"
                                End If
    
                            End If
                    
316                         If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
318                             Stat = Stat & " <" & TituloReal(TempCharIndex) & ">"

320                             ft = e_FontTypeNames.FONTTYPE_CONSEJOVesA
322                         ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
324                             Stat = Stat & " <" & TituloCaos(TempCharIndex) & ">"
326                             ft = e_FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                            End If
                    
328                         If UserList(TempCharIndex).GuildIndex > 0 Then
330                             Stat = Stat & " <" & modGuilds.GuildName(UserList(TempCharIndex).GuildIndex) & ">"
                            End If
    
                        End If ' If user > 0 then
                                
332                     If EsGM(TempCharIndex) Then
334                         ft = e_FontTypeNames.FONTTYPE_GM
                        
336                         If UserList(TempCharIndex).flags.Privilegios = e_PlayerType.Consejero Then
338                             Stat = Stat & " <Consejero>"
                            End If
    
340                         If UserList(TempCharIndex).flags.Privilegios = e_PlayerType.SemiDios Then
342                             Stat = Stat & " <Semi-Dios>"
    
                            End If
    
344                         If UserList(TempCharIndex).flags.Privilegios = e_PlayerType.Dios Then
346                             Stat = Stat & " <Dios>"
    
                            End If
                            
348                         If UserList(TempCharIndex).flags.Privilegios = e_PlayerType.Admin Then
350                             Stat = Stat & " <Administrador>"
    
                            End If
                        
352                     ElseIf UserList(TempCharIndex).Faccion.Status = 0 Then
354                         ft = e_FontTypeNames.FONTTYPE_CRIMINAL
                        
356                     ElseIf UserList(TempCharIndex).Faccion.Status = 1 Then
358                         ft = e_FontTypeNames.FONTTYPE_CITIZEN
                        
                        End If
                        
360                     If UserList(TempCharIndex).flags.Casado = 1 Then
362                         Stat = Stat & " <Pareja de " & UserList(TempCharIndex).flags.Pareja & ">"
    
                        End If
                        
364                     If Len(UserList(TempCharIndex).Desc) > 0 Then
366                         Stat = "Ves a " & UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).Desc
                        Else
368                         Stat = "Ves a " & UserList(TempCharIndex).Name & Stat
    
                        End If
                
370                     If LenB(Stat) > 0 Then
372                         If UserList(TempCharIndex).flags.Muerto Then
374                             Call WriteConsoleMsg(UserIndex, Stat, e_FontTypeNames.FONTTYPE_New_Gris)
                            Else
376                             Call WriteConsoleMsg(UserIndex, Stat, ft)
                            End If
                        
                        
                        End If
                
378                     FoundSomething = 1
380                     UserList(UserIndex).flags.TargetUser = TempCharIndex
382                     UserList(UserIndex).flags.TargetNPC = 0
384                     UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun

                    Else
386                     Call WriteConsoleMsg(UserIndex, "Ves a ??? <Game Master>", e_FontTypeNames.FONTTYPE_GM)
    
388                     UserList(UserIndex).flags.TargetUser = TempCharIndex
390                     UserList(UserIndex).flags.TargetNPC = 0
392                     UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
    
                    End If
                End If
            End If

394         If FoundChar = 2 Then '¿Encontro un NPC?

                Dim estatus As String

396             If EsGM(UserIndex) Or UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 75 Then
398                 estatus = "<" & NpcList(TempCharIndex).Stats.MinHp & "/" & NpcList(TempCharIndex).Stats.MaxHp
                    
400             ElseIf UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 50 Then
402                 If NpcList(TempCharIndex).Stats.MaxHp <> 0 Then
404                     estatus = "<" & Round((NpcList(TempCharIndex).Stats.MinHp / NpcList(TempCharIndex).Stats.MaxHp) * 100#, 0) & "%"
                    End If
 
406             ElseIf UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 25 Then
                
408                 If NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.1) Then
410                     estatus = "<Agonizando"
412                 ElseIf NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.2) Then
414                     estatus = "<Casi muerto"
416                 ElseIf NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.5) Then
418                     estatus = "<Malherido"
420                 ElseIf NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.7) Then
422                     estatus = "<Herido"
424                 ElseIf NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.95) Then
426                     estatus = "<Levemente herido"
                    Else
428                     estatus = "<Intacto"
                    End If
                    
                Else
430                 If NpcList(TempCharIndex).Stats.MinHp < NpcList(TempCharIndex).Stats.MaxHp Then
432                     estatus = "<Herido"
                    Else
434                     estatus = "<Intacto"
                    End If
                        
                End If
                        
436             If NpcList(TempCharIndex).flags.Envenenado > 0 Then
438                 estatus = estatus & " | Envenenado"
                End If
                        
440             If NpcList(TempCharIndex).flags.Paralizado = 1 Then
442                 If UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 100 Then
444                     estatus = estatus & " | Paralizado (" & CInt(NpcList(TempCharIndex).Contadores.Paralisis / 6.5) & "s)"
                    Else
446                     estatus = estatus & " | Paralizado"
                    End If
                End If
                        
448             If NpcList(TempCharIndex).flags.Inmovilizado = 1 Then
450                 If UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 100 Then
452                     estatus = estatus & " | Inmovilizado (" & CInt(NpcList(TempCharIndex).Contadores.Inmovilizado / 6.5) & "s)"
                    Else
454                     estatus = estatus & " | Inmovilizado"

                    End If

                End If
                        
456             estatus = estatus & ">"
    
                'End If
            
458             If Len(NpcList(TempCharIndex).Desc) > 1 Then
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
460                 If NpcList(TempCharIndex).Movement = Caminata Then
462                     NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 5000 + Len(NpcList(TempCharIndex).Desc) * 50 - NpcList(TempCharIndex).IntervaloMovimiento ' 5 segundos + 1 segundo cada 20 caracteres
                    End If

                    'Optimizacion de protocolo por Ladder
464                 Call WriteChatOverHead(UserIndex, "NPCDESC*" & NpcList(TempCharIndex).Numero, NpcList(TempCharIndex).Char.CharIndex, vbWhite)
466             ElseIf TempCharIndex = CentinelaNPCIndex Then
                    'Enviamos nuevamente el texto del centinela según quien pregunta
468                 Call modCentinela.CentinelaSendClave(UserIndex)
                
470             ElseIf NpcList(TempCharIndex).MaestroUser > 0 Then
472                 Call WriteConsoleMsg(UserIndex, "NPCNAME*" & NpcList(TempCharIndex).Numero & "* es mascota de " & UserList(NpcList(TempCharIndex).MaestroUser).Name & " " & estatus, e_FontTypeNames.FONTTYPE_INFO)
                
                Else
                
474                 Call WriteConsoleMsg(UserIndex, "NPCNAME*" & NpcList(TempCharIndex).Numero & "*" & " " & estatus, e_FontTypeNames.FONTTYPE_INFO)
                    ' If UserList(UserIndex).flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
                    ' Call WriteConsoleMsg(UserIndex, "Le pegó primero: " & NpcList(TempCharIndex).flags.AttackedFirstBy & ".", e_FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, NpcList(TempCharIndex).Char.CharIndex, e_FontTypeNames.FONTTYPE_INFO)
                    'End If
                
                End If

476             FoundSomething = 1
478             UserList(UserIndex).flags.TargetNpcTipo = NpcList(TempCharIndex).NPCtype
480             UserList(UserIndex).flags.TargetNPC = TempCharIndex
482             UserList(UserIndex).flags.TargetUser = 0
484             UserList(UserIndex).flags.TargetObj = 0


            
            Dim i As Long, j As Long
                    
486                 For i = 1 To MAXUSERQUESTS
                    
488                     With UserList(UserIndex).QuestStats.Quests(i)
                    
490                         If .QuestIndex Then
492                             If QuestList(.QuestIndex).RequiredTargetNPCs Then
                    
494                                 For j = 1 To QuestList(.QuestIndex).RequiredTargetNPCs
                    
496                                     If QuestList(.QuestIndex).RequiredTargetNPC(j).NpcIndex = NpcList(TempCharIndex).Numero Then
498                                         If QuestList(.QuestIndex).RequiredTargetNPC(j).amount > .NPCsTarget(j) Then
500                                             .NPCsTarget(j) = .NPCsTarget(j) + 1
                    
                                            End If
                                                
502                                         If QuestList(.QuestIndex).RequiredTargetNPC(j).amount = .NPCsTarget(j) Then
504                                             Call FinishQuest(UserIndex, .QuestIndex, i)
506                                             Call WriteUpdateNPCSimbolo(UserIndex, TempCharIndex, 1)
508                                             Call WriteChatOverHead(UserIndex, "¡Quest Finalizada!", NpcList(TempCharIndex).Char.CharIndex, vbYellow)
510                                             Call WriteConsoleMsg(UserIndex, "Quest Finalizada!", e_FontTypeNames.FONTTYPE_INFOIAO)
                                            End If
                    
                                        End If
                    
512                                 Next j
                    
                                End If
                    
                            End If
                    
                        End With
                    
514                 Next i
        
            End If
    
516         If FoundChar = 0 Then
518             UserList(UserIndex).flags.TargetNPC = 0
520             UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
522             UserList(UserIndex).flags.TargetUser = 0

            End If
    
            '*** NO ENCOTRO NADA ***
524         If FoundSomething = 0 Then
526             UserList(UserIndex).flags.TargetNPC = 0
528             UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
530             UserList(UserIndex).flags.TargetUser = 0
532             UserList(UserIndex).flags.TargetObj = 0
534             UserList(UserIndex).flags.TargetObjMap = 0
536             UserList(UserIndex).flags.TargetObjX = 0
538             UserList(UserIndex).flags.TargetObjY = 0

                ' Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        Else

540         If FoundSomething = 0 Then
542             UserList(UserIndex).flags.TargetNPC = 0
544             UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
546             UserList(UserIndex).flags.TargetUser = 0
548             UserList(UserIndex).flags.TargetObj = 0
550             UserList(UserIndex).flags.TargetObjMap = 0
552             UserList(UserIndex).flags.TargetObjX = 0
554             UserList(UserIndex).flags.TargetObjY = 0

                '  Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End If

        
        Exit Sub

LookatTile_Err:
556         Call TraceError(Err.Number, Err.Description & " Pos: " & Map & "-" & X & "-" & Y & " Usuario: " & UserList(UserIndex).Name & " Extra.LookatTile", Erl)
558     If FoundChar = 2 Then
560         Call TraceError(Err.Number, Err.Description & " Pos: " & Map & "-" & X & "-" & Y & "Npc: " & NpcList(TempCharIndex).Numero & " Usuario: " & UserList(UserIndex).Name & "Extra.LookatTile X LADDER", Erl)
        End If

        
End Sub

Function FindDirection(Pos As t_WorldPos, Target As t_WorldPos) As e_Heading
        
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
106         FindDirection = IIf(RandomNumber(0, 1), e_Heading.NORTH, e_Heading.EAST)
            Exit Function

        End If

        'NW
108     If Sgn(X) = 1 And Sgn(Y) = 1 Then
110         FindDirection = IIf(RandomNumber(0, 1), e_Heading.WEST, e_Heading.NORTH)
            Exit Function

        End If

        'SW
112     If Sgn(X) = 1 And Sgn(Y) = -1 Then
114         FindDirection = IIf(RandomNumber(0, 1), e_Heading.WEST, e_Heading.SOUTH)
            Exit Function

        End If

        'SE
116     If Sgn(X) = -1 And Sgn(Y) = -1 Then
118         FindDirection = IIf(RandomNumber(0, 1), e_Heading.SOUTH, e_Heading.EAST)
            Exit Function

        End If

        'Sur
120     If Sgn(X) = 0 And Sgn(Y) = -1 Then
122         FindDirection = e_Heading.SOUTH
            Exit Function

        End If

        'norte
124     If Sgn(X) = 0 And Sgn(Y) = 1 Then
126         FindDirection = e_Heading.NORTH
            Exit Function

        End If

        'oeste
128     If Sgn(X) = 1 And Sgn(Y) = 0 Then
130         FindDirection = e_Heading.WEST
            Exit Function

        End If

        'este
132     If Sgn(X) = -1 And Sgn(Y) = 0 Then
134         FindDirection = e_Heading.EAST
            Exit Function

        End If

        'misma
136     If Sgn(X) = 0 And Sgn(Y) = 0 Then
138         FindDirection = 0
            Exit Function

        End If

        
        Exit Function

FindDirection_Err:
140     Call TraceError(Err.Number, Err.Description, "Extra.FindDirection", Erl)

        
End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean
        
        On Error GoTo ItemNoEsDeMapa_Err
        

100     ItemNoEsDeMapa = ObjData(Index).OBJType <> e_OBJType.otPuertas And ObjData(Index).OBJType <> e_OBJType.otForos And ObjData(Index).OBJType <> e_OBJType.otCarteles And ObjData(Index).OBJType <> e_OBJType.otArboles And ObjData(Index).OBJType <> e_OBJType.otYacimiento And ObjData(Index).OBJType <> e_OBJType.otTeleport And ObjData(Index).OBJType <> e_OBJType.OtCorreo And ObjData(Index).OBJType <> e_OBJType.OtDecoraciones

        
        Exit Function

ItemNoEsDeMapa_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.ItemNoEsDeMapa", Erl)

        
End Function

'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
        
        On Error GoTo MostrarCantidad_Err
        
100     MostrarCantidad = ObjData(Index).OBJType <> e_OBJType.otPuertas And ObjData(Index).OBJType <> e_OBJType.otForos And ObjData(Index).OBJType <> e_OBJType.otCarteles And ObjData(Index).OBJType <> e_OBJType.otYacimiento And ObjData(Index).OBJType <> e_OBJType.otArboles And ObjData(Index).OBJType <> e_OBJType.OtCorreo And ObjData(Index).OBJType <> e_OBJType.otTeleport

        
        Exit Function

MostrarCantidad_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.MostrarCantidad", Erl)

        
End Function

Public Function EsObjetoFijo(ByVal OBJType As e_OBJType) As Boolean
        
        On Error GoTo EsObjetoFijo_Err
        

100     EsObjetoFijo = OBJType = e_OBJType.otForos Or OBJType = e_OBJType.otCarteles Or OBJType = e_OBJType.otArboles Or OBJType = e_OBJType.otYacimiento Or OBJType = e_OBJType.OtDecoraciones

        
        Exit Function

EsObjetoFijo_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.EsObjetoFijo", Erl)

        
End Function

Public Function HayPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
100     If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
102         HayPuerta = (ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = e_OBJType.otPuertas) And ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 And (ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0)
        End If
End Function

Public Sub CargarMapasEspeciales()

        Dim File As clsIniManager
100     Set File = New clsIniManager
    
102     Call File.Initialize(DatPath & "MapasEspeciales.dat")
    
        Dim Cantidad As Integer
104     Cantidad = val(File.GetValue("MapasInterdimensionales", "Cantidad"))
    
        Dim i As Integer
    
106     If Cantidad > 0 Then
108         ReDim MapasInterdimensionales(1 To Cantidad)
        
110         For i = 1 To Cantidad
112             MapasInterdimensionales(i) = val(File.GetValue("MapasInterdimensionales", "Mapa" & i))
            Next
        Else
114         ReDim MapasInterdimensionales(0)
        End If
    
116     Cantidad = val(File.GetValue("MapasIgnoranLimpieza", "Cantidad"))
    
118     If Cantidad > 0 Then
120         ReDim MapasIgnoranLimpieza(1 To Cantidad)
        
122         For i = 1 To Cantidad
124             MapasIgnoranLimpieza(i) = val(File.GetValue("MapasIgnoranLimpieza", "Mapa" & i))
            Next
        Else
126         ReDim MapasIgnoranLimpieza(0)
        End If
    
128     Set File = Nothing

End Sub

Public Sub CargarTipoUsuario(ByVal UserIndex As Integer)
                
        Dim cantidadDonadores As Integer
100     If Not FileExist(DatPath & "donadores.dat", vbArchive) Then
            Exit Sub
        End If
        
        Dim IniFile As clsIniManager
106     Set IniFile = New clsIniManager
108     Call IniFile.Initialize(DatPath & "donadores.dat")
        
        cantidadDonadores = val(IniFile.GetValue("INIT", "Cantidad"))
        
        If cantidadDonadores > 0 Then
        
            Dim emailDonador As String
            Dim rango As eTipoUsuario
            
            Dim i As Integer
            
            For i = 1 To cantidadDonadores
            
                Dim str() As String
                str = Split(IniFile.GetValue("DONADOR", "Donador" & i), "#")
                emailDonador = str(0)
                rango = val(str(1))
                
                If LCase(UserList(UserIndex).Email) = LCase(emailDonador) Then
                    UserList(UserIndex).Stats.tipoUsuario = rango
                    Set IniFile = Nothing
                    Exit Sub
                End If
                
            Next i
            
        End If
        
        UserList(UserIndex).Stats.tipoUsuario = eTipoUsuario.tNormal
152     Set IniFile = Nothing
End Sub



