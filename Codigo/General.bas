Attribute VB_Name = "General"

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

Public Type TDonador

    activo As Byte
    CreditoDonador As Integer
    FechaExpiracion As Date

End Type

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Global LeerNPCs As New clsIniReader

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer)
        
        On Error GoTo DarCuerpoDesnudo_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/14/07
        'Da cuerpo desnudo a un usuario
        '***************************************************
        Dim CuerpoDesnudo As Integer

100     Select Case UserList(UserIndex).genero

            Case eGenero.Hombre

102             Select Case UserList(UserIndex).raza

                    Case eRaza.Humano
104                     CuerpoDesnudo = 21 'ok

106                 Case eRaza.Drow
108                     CuerpoDesnudo = 32 ' ok

110                 Case eRaza.Elfo
112                     CuerpoDesnudo = 510 'Revisar

114                 Case eRaza.Gnomo
116                     CuerpoDesnudo = 508 'Revisar

118                 Case eRaza.Enano
120                     CuerpoDesnudo = 53 'ok

122                 Case eRaza.Orco
124                     CuerpoDesnudo = 248 ' ok

                End Select

126         Case eGenero.Mujer

128             Select Case UserList(UserIndex).raza

                    Case eRaza.Humano
130                     CuerpoDesnudo = 39 'ok

132                 Case eRaza.Drow
134                     CuerpoDesnudo = 40 'ok

136                 Case eRaza.Elfo
138                     CuerpoDesnudo = 511 'Revisar

140                 Case eRaza.Gnomo
142                     CuerpoDesnudo = 509 'Revisar

144                 Case eRaza.Enano
146                     CuerpoDesnudo = 60 ' ok

148                 Case eRaza.Orco
150                     CuerpoDesnudo = 249 'ok

                End Select

        End Select

152     UserList(UserIndex).Char.Body = CuerpoDesnudo

154     UserList(UserIndex).flags.Desnudo = 1

        
        Exit Sub

DarCuerpoDesnudo_Err:
156     Call RegistrarError(Err.Number, Err.Description, "General.DarCuerpoDesnudo", Erl)
158     Resume Next
        
End Sub

Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Byte)
        'b ahora es boolean,
        'b=true bloquea el tile en (x,y)
        'b=false desbloquea el tile en (x,y)
        'toMap = true -> Envia los datos a todo el mapa
        'toMap = false -> Envia los datos al user
        'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
        'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
        ' WyroX: Uso bloqueo parcial
        
        On Error GoTo Bloquear_Err
        
        ' Envío sólo los flags de bloq
100     b = b And eBlock.ALL_SIDES

102     If toMap Then
104         Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
        Else
106         Call WriteBlockPosition(sndIndex, X, Y, b)
        End If

        
        Exit Sub

Bloquear_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.Bloquear", Erl)
110     Resume Next
        
End Sub

Sub MostrarBloqueosPuerta(ByVal toMap As Boolean, _
                          ByVal sndIndex As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer)
        
        On Error GoTo MostrarBloqueosPuerta_Err
        
        Dim Map       As Integer
        Dim ModPuerta As Integer
        
100     If toMap Then
102         Map = sndIndex
        Else
104         Map = UserList(sndIndex).Pos.Map
        End If
        
        ModPuerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo

        Select Case ModPuerta
        
            Case 0
106             ' Bloqueos superiores
                Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
108             Call Bloquear(toMap, sndIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)

                ' Bloqueos inferiores
110             Call Bloquear(toMap, sndIndex, X, Y + 1, MapData(Map, X, Y + 1).Blocked)
112             Call Bloquear(toMap, sndIndex, X - 1, Y + 1, MapData(Map, X - 1, Y + 1).Blocked)

            Case 1
                ' para palancas o teclas sin modicar bloqueos en X,Y
                
            Case 2
                ' Bloqueos superiores
                Call Bloquear(toMap, sndIndex, X, Y - 1, MapData(Map, X, Y - 1).Blocked)
                Call Bloquear(toMap, sndIndex, X - 1, Y - 1, MapData(Map, X - 1, Y - 1).Blocked)
                Call Bloquear(toMap, sndIndex, X + 1, Y - 1, MapData(Map, X + 1, Y - 1).Blocked)
                ' Bloqueos inferiores
                Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
                Call Bloquear(toMap, sndIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                Call Bloquear(toMap, sndIndex, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                
            Case 3
                ' Bloqueos superiores
                Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
                Call Bloquear(toMap, sndIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                Call Bloquear(toMap, sndIndex, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                ' Bloqueos inferiores
                Call Bloquear(toMap, sndIndex, X, Y + 1, MapData(Map, X, Y + 1).Blocked)
                Call Bloquear(toMap, sndIndex, X - 1, Y + 1, MapData(Map, X - 1, Y + 1).Blocked)
                Call Bloquear(toMap, sndIndex, X + 1, Y + 1, MapData(Map, X + 1, Y + 1).Blocked)

            Case 4
                ' Bloqueos superiores
                Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
                ' Bloqueos inferiores
                Call Bloquear(toMap, sndIndex, X, Y + 1, MapData(Map, X, Y + 1).Blocked)

            Case 5 'Ver WyroX
                ' Bloqueos vertical ver ReyarB
                Call Bloquear(toMap, sndIndex, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                Call Bloquear(toMap, sndIndex, X + 1, Y - 1, MapData(Map, X + 1, Y - 1).Blocked)

                ' Bloqueos horizontal
                Call Bloquear(toMap, sndIndex, X, Y - 2, MapData(Map, X, Y - 2).Blocked)
                Call Bloquear(toMap, sndIndex, X - 1, Y - 2, MapData(Map, X - 1, Y - 2).Blocked)


            Case 6 ' Ver WyroX
                ' Bloqueos superiores ver ReyarB
                Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
                Call Bloquear(toMap, sndIndex, X, Y - 1, MapData(Map, X, Y - 1).Blocked)

                ' Bloqueos inferiores
                Call Bloquear(toMap, sndIndex, X, Y - 2, MapData(Map, X, Y - 2).Blocked)
                Call Bloquear(toMap, sndIndex, X + 1, Y - 2, MapData(Map, X + 1, Y - 2).Blocked)

        End Select

        Exit Sub

MostrarBloqueosPuerta_Err:
114     Call RegistrarError(Err.Number, Err.Description, "General.MostrarBloqueosPuerta", Erl)
        
End Sub

Sub BloquearPuerta(ByVal Map As Integer, _
                   ByVal X As Integer, _
                   ByVal Y As Integer, _
                   ByVal Bloquear As Boolean)
        
        On Error GoTo BloquearPuerta_Err
        Dim ModPuerta As Integer
        
        'ver reyarb
        ModPuerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo

        Select Case ModPuerta
        
            Case 0 'puerta 2 tiles
100
                ' Bloqueos superiores
                MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or eBlock.NORTH, MapData(Map, X, Y).Blocked And Not eBlock.NORTH)
102             MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or eBlock.NORTH, MapData(Map, X - 1, Y).Blocked And Not eBlock.NORTH)

                ' Cambio bloqueos inferiores
104             MapData(Map, X, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X, Y + 1).Blocked Or eBlock.SOUTH, MapData(Map, X, Y + 1).Blocked And Not eBlock.SOUTH)
106             MapData(Map, X - 1, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y + 1).Blocked Or eBlock.SOUTH, MapData(Map, X - 1, Y + 1).Blocked And Not eBlock.SOUTH)

            Case 1
                ' para palancas o teclas sin modicar bloqueos en X,Y

            Case 2 ' puerta 3 tiles 1 arriba
                ' Bloqueos superiores
                MapData(Map, X, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X, Y - 1).Blocked Or eBlock.NORTH, MapData(Map, X, Y - 1).Blocked And Not eBlock.NORTH)
                MapData(Map, X - 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y - 1).Blocked Or eBlock.NORTH, MapData(Map, X - 1, Y - 1).Blocked And Not eBlock.NORTH)
                MapData(Map, X + 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y - 1).Blocked Or eBlock.NORTH, MapData(Map, X + 1, Y - 1).Blocked And Not eBlock.NORTH)
                ' Cambio bloqueos inferiores
                MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or eBlock.SOUTH, MapData(Map, X, Y).Blocked And Not eBlock.SOUTH)
                MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or eBlock.SOUTH, MapData(Map, X - 1, Y).Blocked And Not eBlock.SOUTH)
                MapData(Map, X + 1, Y).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y).Blocked Or eBlock.SOUTH, MapData(Map, X + 1, Y).Blocked And Not eBlock.SOUTH)
                
            Case 3 ' puerta 3 tiles
                ' Bloqueos superiores
                MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or eBlock.NORTH, MapData(Map, X, Y).Blocked And Not eBlock.NORTH)
                MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or eBlock.NORTH, MapData(Map, X - 1, Y).Blocked And Not eBlock.NORTH)
                MapData(Map, X + 1, Y).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y).Blocked Or eBlock.NORTH, MapData(Map, X + 1, Y).Blocked And Not eBlock.NORTH)
                ' Cambio bloqueos inferiores
                MapData(Map, X, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X, Y + 1).Blocked Or eBlock.SOUTH, MapData(Map, X, Y + 1).Blocked And Not eBlock.SOUTH)
                MapData(Map, X - 1, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y + 1).Blocked Or eBlock.SOUTH, MapData(Map, X - 1, Y + 1).Blocked And Not eBlock.SOUTH)
                MapData(Map, X + 1, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y + 1).Blocked Or eBlock.SOUTH, MapData(Map, X + 1, Y + 1).Blocked And Not eBlock.SOUTH)
        
            Case 4 'puerta 1 tiles
                ' Bloqueos superiores
                MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or eBlock.NORTH, MapData(Map, X, Y).Blocked And Not eBlock.NORTH)
                ' Cambio bloqueos inferiores
                MapData(Map, X, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X, Y + 1).Blocked Or eBlock.SOUTH, MapData(Map, X, Y + 1).Blocked And Not eBlock.SOUTH)
                
            Case 5 'Ver WyroX
                ' Bloqueos  vertical ver ReyarB
                MapData(Map, X + 1, Y).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y).Blocked Or eBlock.ALL_SIDES, MapData(Map, X + 1, Y).Blocked And Not eBlock.ALL_SIDES)
                MapData(Map, X + 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y - 1).Blocked Or eBlock.ALL_SIDES, MapData(Map, X + 1, Y - 1).Blocked And Not eBlock.ALL_SIDES)
                
                ' Cambio horizontal
                MapData(Map, X, Y - 2).Blocked = IIf(Bloquear, MapData(Map, X, Y - 2).Blocked Or eBlock.ALL_SIDES, MapData(Map, X, Y - 2).Blocked And Not eBlock.ALL_SIDES)
                MapData(Map, X - 1, Y - 2).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y - 2).Blocked Or eBlock.ALL_SIDES, MapData(Map, X - 1, Y - 2).Blocked And Not eBlock.ALL_SIDES)


            Case 6 ' Ver Wyrox
                ' Bloqueos vertical ver ReyarB
                MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or eBlock.ALL_SIDES, MapData(Map, X - 1, Y).Blocked And Not eBlock.ALL_SIDES)
                MapData(Map, X - 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y - 1).Blocked Or eBlock.ALL_SIDES, MapData(Map, X - 1, Y - 1).Blocked And Not eBlock.ALL_SIDES)
                
                ' Cambio bloqueos Puerta abierta
                MapData(Map, X, Y - 2).Blocked = IIf(Bloquear, MapData(Map, X, Y - 2).Blocked Or eBlock.ALL_SIDES, MapData(Map, X, Y - 2).Blocked And Not eBlock.ALL_SIDES)
                MapData(Map, X + 1, Y + 2).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y - 2).Blocked Or eBlock.ALL_SIDES, MapData(Map, X + 1, Y - 2).Blocked And Not eBlock.ALL_SIDES)

                
        End Select

        ' Mostramos a todos
108     Call MostrarBloqueosPuerta(True, Map, X, Y)
        
        Exit Sub

BloquearPuerta_Err:
110     Call RegistrarError(Err.Number, Err.Description, "General.BloquearPuerta", Erl)
        
End Sub

Function HayCosta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayCosta_Err
        

        'Ladder 10 - 2 - 2010
        'Chequea si hay costa en los tiles proximos al usuario
100     If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
102         If ((MapData(Map, X, Y).Graphic(1) >= 22552 And MapData(Map, X, Y).Graphic(1) <= 22599) Or (MapData(Map, X, Y).Graphic(1) >= 7283 And MapData(Map, X, Y).Graphic(1) <= 7378) Or (MapData(Map, X, Y).Graphic(1) >= 13387 And MapData(Map, X, Y).Graphic(1) <= 13482)) And MapData(Map, X, Y).Graphic(2) = 0 Then
104             HayCosta = True
            Else
106             HayCosta = False

            End If

        Else
108         HayCosta = False

        End If

        
        Exit Function

HayCosta_Err:
110     Call RegistrarError(Err.Number, Err.Description, "General.HayCosta", Erl)
112     Resume Next
        
End Function

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayAgua_Err
        

100     With MapData(Map, X, Y)
102         If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
104             HayAgua = ((.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
                    (.Graphic(1) >= 124 And .Graphic(1) <= 139) Or _
                    (.Graphic(1) >= 24223 And .Graphic(1) <= 24238) Or _
                    (.Graphic(1) >= 24303 And .Graphic(1) <= 24318) Or _
                    (.Graphic(1) >= 468 And .Graphic(1) <= 483) Or _
                    (.Graphic(1) >= 44668 And .Graphic(1) <= 44939) Or _
                    (.Graphic(1) >= 24143 And .Graphic(1) <= 24158)) Or _
                    (.Graphic(1) >= 12628 And .Graphic(1) <= 12643) Or _
                    (.Graphic(1) >= 2948 And .Graphic(1) <= 2963)
            Else
106             HayAgua = False
    
            End If
        End With

        
        Exit Function

HayAgua_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.HayAgua", Erl)
110     Resume Next
        
End Function

Function EsArbol(ByVal GrhIndex As Long) As Boolean
        
        On Error GoTo EsArbol_Err
    
        
100     EsArbol = GrhIndex = 7000 Or GrhIndex = 7001 Or GrhIndex = 7002 Or GrhIndex = 641 Or GrhIndex = 26075 Or GrhIndex = 643 Or GrhIndex = 644 Or _
           GrhIndex = 647 Or GrhIndex = 26076 Or GrhIndex = 7222 Or GrhIndex = 7223 Or GrhIndex = 7224 Or GrhIndex = 7225 Or GrhIndex = 7226 Or _
           GrhIndex = 26077 Or GrhIndex = 26079 Or GrhIndex = 735 Or GrhIndex = 32343 Or GrhIndex = 32344 Or GrhIndex = 26080 Or GrhIndex = 26081 Or _
           GrhIndex = 32345 Or GrhIndex = 32346 Or GrhIndex = 32347 Or GrhIndex = 32348 Or GrhIndex = 32349 Or GrhIndex = 32350 Or GrhIndex = 32351 Or _
           GrhIndex = 32352 Or GrhIndex = 14961 Or GrhIndex = 14950 Or GrhIndex = 14951 Or GrhIndex = 14952 Or GrhIndex = 14953 Or GrhIndex = 14954 Or _
           GrhIndex = 14955 Or GrhIndex = 14956 Or GrhIndex = 14957 Or GrhIndex = 14958 Or GrhIndex = 14959 Or GrhIndex = 14962 Or GrhIndex = 14963 Or _
           GrhIndex = 14964 Or GrhIndex = 14967 Or GrhIndex = 14968 Or GrhIndex = 14969 Or GrhIndex = 14970 Or GrhIndex = 14971 Or GrhIndex = 14972 Or _
           GrhIndex = 14973 Or GrhIndex = 14974 Or GrhIndex = 14975 Or GrhIndex = 14976 Or GrhIndex = 14978 Or GrhIndex = 14980 Or GrhIndex = 14982 Or _
           GrhIndex = 14983 Or GrhIndex = 14984 Or GrhIndex = 14985 Or GrhIndex = 14987 Or GrhIndex = 14988 Or GrhIndex = 26078 Or GrhIndex = 26192

        
        Exit Function

EsArbol_Err:
102     Call RegistrarError(Err.Number, Err.Description, "General.EsArbol", Erl)

        
End Function

Private Function HayLava(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayLava_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/12/07
        '***************************************************
100     If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
102         If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
104             HayLava = True
            Else
106             HayLava = False

            End If

        Else
108         HayLava = False

        End If

        
        Exit Function

HayLava_Err:
110     Call RegistrarError(Err.Number, Err.Description, "General.HayLava", Erl)
112     Resume Next
        
End Function

Sub ApagarFogatas()

        'Ladder /ApagarFogatas
        On Error GoTo ErrHandler

        Dim obj As obj
100         obj.ObjIndex = FOGATA_APAG
102         obj.Amount = 1

        Dim MapaActual As Long
        Dim Y          As Long
        Dim X          As Long

104     For MapaActual = 1 To NumMaps
106         For Y = YMinMapSize To YMaxMapSize
108             For X = XMinMapSize To XMaxMapSize

110                 If MapInfo(MapaActual).lluvia Then
                
112                     If MapData(MapaActual, X, Y).ObjInfo.ObjIndex = FOGATA Then
                    
114                         Call EraseObj(MAX_INVENTORY_OBJS, MapaActual, X, Y)
116                         Call MakeObj(obj, MapaActual, X, Y)

                        End If

                    End If

118             Next X
120         Next Y
122     Next MapaActual

        Exit Sub
    
ErrHandler:
124     Call LogError("Error producido al apagar las fogatas de " & X & "-" & Y & " del mapa: " & MapaActual & "    -" & Err.Description)

End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)
        
        On Error GoTo EnviarSpawnList_Err
        

        Dim K          As Long
        Dim npcNames() As String

100     'Debug.Print UBound(SpawnList)
102     ReDim npcNames(1 To UBound(SpawnList)) As String

104     For K = 1 To UBound(SpawnList)
106         npcNames(K) = SpawnList(K).NpcName
108     Next K

110     Call WriteSpawnList(UserIndex, npcNames())

        
        Exit Sub

EnviarSpawnList_Err:
112     Call RegistrarError(Err.Number, Err.Description, "General.EnviarSpawnList", Erl)
114     Resume Next
        
End Sub

Public Sub LeerLineaComandos()
        
        On Error GoTo LeerLineaComandos_Err
        

        Dim rdata As String

100     rdata = Command
102     rdata = Right$(rdata, Len(rdata))
104     ClaveApertura = ReadField(1, rdata, Asc("*")) ' NICK

        
        Exit Sub

LeerLineaComandos_Err:
106     Call RegistrarError(Err.Number, Err.Description, "General.LeerLineaComandos", Erl)
108     Resume Next
        
End Sub

Private Sub InicializarConstantes()
        
        On Error GoTo InicializarConstantes_Err
    
        
    
100     LastBackup = Format(Now, "Short Time")
102     minutos = Format(Now, "Short Time")
    
104     IniPath = App.Path & "\"

206     ListaRazas(eRaza.Humano) = "Humano"
208     ListaRazas(eRaza.Elfo) = "Elfo"
210     ListaRazas(eRaza.Drow) = "Elfo Oscuro"
212     ListaRazas(eRaza.Gnomo) = "Gnomo"
214     ListaRazas(eRaza.Enano) = "Enano"
        'ListaRazas(eRaza.Orco) = "Orco"
    
216     ListaClases(eClass.Mage) = "Mago"
218     ListaClases(eClass.Cleric) = "Clérigo"
220     ListaClases(eClass.Warrior) = "Guerrero"
222     ListaClases(eClass.Assasin) = "Asesino"
224     ListaClases(eClass.Bard) = "Bardo"
226     ListaClases(eClass.Druid) = "Druida"
228     ListaClases(eClass.Paladin) = "Paladín"
230     ListaClases(eClass.Hunter) = "Cazador"
232     ListaClases(eClass.Trabajador) = "Trabajador"
234     ListaClases(eClass.Pirat) = "Pirata"
236     ListaClases(eClass.Thief) = "Ladrón"
238     ListaClases(eClass.Bandit) = "Bandido"
    
240     SkillsNames(eSkill.magia) = "Magia"
242     SkillsNames(eSkill.Robar) = "Robar"
244     SkillsNames(eSkill.Tacticas) = "Destreza en combate"
246     SkillsNames(eSkill.Armas) = "Combate con armas"
248     SkillsNames(eSkill.Meditar) = "Meditar"
250     SkillsNames(eSkill.Apuñalar) = "Apuñalar"
252     SkillsNames(eSkill.Ocultarse) = "Ocultarse"
254     SkillsNames(eSkill.Supervivencia) = "Supervivencia"
256     SkillsNames(eSkill.Comerciar) = "Comercio"
258     SkillsNames(eSkill.Defensa) = "Defensa con escudo"
260     SkillsNames(eSkill.liderazgo) = "Liderazgo"
262     SkillsNames(eSkill.Proyectiles) = "Armas a distancia"
264     SkillsNames(eSkill.Wrestling) = "Combate sin armas"
266     SkillsNames(eSkill.Navegacion) = "Navegación"
268     SkillsNames(eSkill.equitacion) = "Equitación"
270     SkillsNames(eSkill.Resistencia) = "Resistencia mágica"
272     SkillsNames(eSkill.Talar) = "Tala"
274     SkillsNames(eSkill.Pescar) = "Pesca"
276     SkillsNames(eSkill.Mineria) = "Minería"
278     SkillsNames(eSkill.Herreria) = "Herrería"
280     SkillsNames(eSkill.Carpinteria) = "Carpintería"
282     SkillsNames(eSkill.Alquimia) = "Alquimia"
284     SkillsNames(eSkill.Sastreria) = "Sastrería"
286     SkillsNames(eSkill.Domar) = "Domar"
   
288     ListaAtributos(eAtributos.Fuerza) = "Fuerza"
290     ListaAtributos(eAtributos.Agilidad) = "Agilidad"
292     ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
294     ListaAtributos(eAtributos.Constitucion) = "Constitución"
296     ListaAtributos(eAtributos.Carisma) = "Carisma"
    
298     centinelaActivado = False
    
300     IniPath = App.Path & "\"
    
        'Bordes del mapa
302     MinXBorder = XMinMapSize + (XWindow \ 2)
304     MaxXBorder = XMaxMapSize - (XWindow \ 2)
306     MinYBorder = YMinMapSize + (YWindow \ 2)
308     MaxYBorder = YMaxMapSize - (YWindow \ 2)
    
        
        Exit Sub

InicializarConstantes_Err:
310     Call RegistrarError(Err.Number, Err.Description, "General.InicializarConstantes", Erl)

        
End Sub

Sub Main()

    On Error GoTo Handler

    ' Me fijo si ya hay un proceso llamado server.exe abierto
    If GetProcess(App.EXEName & ".exe") > 1 Then
    
        ' Si lo hay, pregunto si lo queremos cerrar.
        If MsgBox("Se ha encontrado mas de 1 instancia abierta de esta aplicación, ¿ Desea continuar ?", vbYesNo) = vbNo Then
            
            ' Cerramos esta instancia de la aplicacion.
            End

        End If
        
    End If
    
    Call LeerLineaComandos
    
    Call CargarRanking
    
    Dim f As Date
    
    Call ChDir(App.Path)
    Call ChDrive(App.Path)

    Call InicializarConstantes
    
    frmCargando.Show
    
    'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")
    
    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    
    Call LoadGuildsDB
    
    Call LoadConfiguraciones
    Call CargarCodigosDonador
    Call loadAdministrativeUsers

    '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    
    MaxUsers = 0
    Call LoadSini
    Call LoadIntervalos
    Call CargarForbidenWords
    Call CargaApuestas
    Call CargarSpawnList
    Call LoadMotd
    Call BanIpCargar
    
    frmCargando.Label1(2).Caption = "Conectando base de datos y limpiando usuarios logueados"

    'Conecto base de datos
    Call Database_Connect
        
    'Reinicio los users online
    Call SetUsersLoggedDatabase(0)
        
    'Leo el record de usuarios
    RecordUsuarios = LeerRecordUsuariosDatabase()
        
    'Tarea pesada
    Call LogoutAllUsersAndAccounts

    '*************************************************
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat
    '*************************************************
    
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    'Call LoadOBJData
    Call LoadOBJData

    Call InitTesoro
    Call InitRegalo
        
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos
        
    frmCargando.Label1(2).Caption = "Cargando Objetos de Herrería"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    
    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpintería"
    Call LoadObjCarpintero
    
    frmCargando.Label1(2).Caption = "Cargando Objetos de Alquimista"
    Call LoadObjAlquimista
    
    frmCargando.Label1(2).Caption = "Cargando Objetos de Sastre"
    Call LoadObjSastre
    
    frmCargando.Label1(2).Caption = "Cargando Pesca"
    Call LoadPesca
    
    frmCargando.Label1(2).Caption = "Cargando Recursos Especiales"
    Call LoadRecursosEspeciales

    frmCargando.Label1(2).Caption = "Cargando Rangos de Faccion"
    Call LoadRangosFaccion

    frmCargando.Label1(2).Caption = "Cargando Recompensas de Faccion"
    Call LoadRecompensasFaccion
    
    frmCargando.Label1(2).Caption = "Cargando Balance.dat"
    Call LoadBalance    '4/01/08 Pablo ToxicWaste
    
    frmCargando.Label1(2).Caption = "Cargando Ciudades.dat"
    Call CargarCiudades
    
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando WorldBackup"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData
    End If
        
    Call CargarInfoRetos
    Call CargarInfoEventos
    
    ' Pretorianos
    frmCargando.Label1(2).Caption = "Cargando Pretorianos.dat"
    Call LoadPretorianData
    
    frmCargando.Label1(2).Caption = "Cargando Logros.ini"
    Call CargarLogros ' Ladder 22/04/2015
    
    frmCargando.Label1(2).Caption = "Cargando Baneos Temporales"
    Call LoadBans
    
    frmCargando.Label1(2).Caption = "Cargando Usuarios Donadores"
    Call LoadDonadores
    Call LoadObjDonador
    Call LoadQuests

    EstadoGlobal = True
    
    Call InicializarLimpieza

    'Comentado porque hay worldsave en ese mapa!
    'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Dim LoopC As Integer
    
    'Resetea las conexiones de los usuarios
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    With frmMain
        .Minuto.Enabled = True
        .TimerGuardarUsuarios.Enabled = True
        .TimerGuardarUsuarios.Interval = IntervaloTimerGuardarUsuarios
        .tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .Auditoria.Enabled = True
        .KillLog.Enabled = True
        .TIMER_AI.Enabled = True
        .Invasion.Enabled = True
    End With
    
    Subasta.SubastaHabilitada = True
    Subasta.HaySubastaActiva = False
    Call ResetMeteo
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    'Configuracion de los sockets
    
    Call SecurityIp.InitIpTables(1000)
        
    'Cierra el socket de escucha
    If LastSockListen >= 0 Then Call apiclosesocket(LastSockListen)

    Call IniciaWsApi(frmMain.hwnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

    If SockListen <> -1 Then
        Call WriteVar(IniPath & "Server.ini", "INIT", "LastSockListen", SockListen) _
                ' Guarda el socket escuchando
    Else
        Call MsgBox("Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly)

    End If
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Call GetHoraActual
    
    HoraMundo = GetTickCount() - DuracionDia \ 2

    frmCargando.Visible = False
    Unload frmCargando

    'Log
    Dim n As Integer
    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & Time & " server iniciado " & App.Major & "." & App.Minor & "." & App.Revision
    Close #n
    
    'Ocultar
    Call frmMain.InitMain(HideMe)
    
    tInicioServer = GetTickCount()

    Exit Sub
        
Handler:
    Call RegistrarError(Err.Number, Err.Description, "General.Main", Erl)

    Resume Next

End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
        '*****************************************************************
        'Se fija si existe el archivo
        '*****************************************************************
        
        On Error GoTo FileExist_Err
        
100     FileExist = LenB(dir$(File, FileType)) <> 0

        
        Exit Function

FileExist_Err:
102     Call RegistrarError(Err.Number, Err.Description, "General.FileExist", Erl)
104     Resume Next
        
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
        
        On Error GoTo ReadField_Err
        

        '*****************************************************************
        'Gets a field from a string
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/15/2004
        'Gets a field from a delimited string
        '*****************************************************************
        Dim i          As Long

        Dim LastPos    As Long

        Dim CurrentPos As Long

        Dim delimiter  As String * 1
    
100     delimiter = Chr$(SepASCII)
    
102     For i = 1 To Pos
104         LastPos = CurrentPos
106         CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
108     Next i
    
110     If CurrentPos = 0 Then
112         ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
        Else
114         ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

        End If

        
        Exit Function

ReadField_Err:
116     Call RegistrarError(Err.Number, Err.Description, "General.ReadField", Erl)
118     Resume Next
        
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
        
        On Error GoTo MapaValido_Err
        
100     MapaValido = Map >= 1 And Map <= NumMaps

        
        Exit Function

MapaValido_Err:
102     Call RegistrarError(Err.Number, Err.Description, "General.MapaValido", Erl)
104     Resume Next
        
End Function

Sub MostrarNumUsers()
        
        On Error GoTo MostrarNumUsers_Err
        

100     Call SendData(SendTarget.ToAll, 0, PrepareMessageOnlineUser(NumUsers))
102     frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers
    
104     Call SetUsersLoggedDatabase(NumUsers)

        
        Exit Sub

MostrarNumUsers_Err:
106     Call RegistrarError(Err.Number, Err.Description, "General.MostrarNumUsers", Erl)
108     Resume Next
        
End Sub

Public Sub LogCriticEvent(Desc As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
104     Print #nfile, Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
104     Print #nfile, Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogIndex(ByVal index As Integer, ByVal Desc As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\" & index & ".log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogError(Desc As String)

    Dim nfile As Integer: nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

End Sub

Public Sub LogPerformance(Desc As String)

    Dim nfile As Integer: nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\Performance.log" For Append Shared As #nfile
        Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

End Sub

Public Sub LogConsulta(Desc As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\ConsultasGM.log" For Append Shared As #nfile
104     Print #nfile, Date & " - " & Time & " - " & Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogStatic(Desc As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogTarea(Desc As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile(1) ' obtenemos un canal
102     Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogClanes(ByVal str As String)
        
        On Error GoTo LogClanes_Err
        

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        
        Exit Sub

LogClanes_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.LogClanes", Erl)
110     Resume Next
        
End Sub

Public Sub LogIP(ByVal str As String)
        
        On Error GoTo LogIP_Err
        

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\IP.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        
        Exit Sub

LogIP_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.LogIP", Erl)
110     Resume Next
        
End Sub

Public Sub LogDesarrollo(ByVal str As String)
        
        On Error GoTo LogDesarrollo_Err
        

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        
        Exit Sub

LogDesarrollo_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.LogDesarrollo", Erl)
110     Resume Next
        
End Sub

Public Sub LogGM(nombre As String, texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
        'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
102     Open App.Path & "\logs\" & nombre & ".log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & texto
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogPremios(GM As String, UserName As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, Motivo As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal

102     Open App.Path & "\logs\PremiosOtorgados.log" For Append Shared As #nfile
104     Print #nfile, "[" & GM & "]" & vbNewLine
106     Print #nfile, Date & " " & Time & vbNewLine
108     Print #nfile, "Item: " & ObjData(ObjIndex).name & " (" & ObjIndex & ") Cantidad: " & Cantidad & vbNewLine
110     Print #nfile, "Motivo: " & Motivo & vbNewLine & vbNewLine
112     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogDatabaseError(Desc As String)
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 09/10/2018
        '***************************************************

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
    
102     Open App.Path & "\logs\Database.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " - " & Desc
106     Close #nfile
     
108     Debug.Print "Error en la BD: " & Desc & vbNewLine & _
            "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
            
        Exit Sub
    
ErrHandler:

End Sub

Public Sub SaveDayStats()
        ''On Error GoTo errhandler
        ''
        ''Dim nfile As Integer
        ''nfile = FreeFile ' obtenemos un canal
        ''Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile
        ''
        ''Print #nfile, "<stats>"
        ''Print #nfile, "<ao>"
        ''Print #nfile, "<dia>" & Date & "</dia>"
        ''Print #nfile, "<hora>" & Time & "</hora>"
        ''Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
        ''Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
        ''Print #nfile, "</ao>"
        ''Print #nfile, "</stats>"
        ''
        ''
        ''Close #nfile
    
        On Error GoTo SaveDayStats_Err
    
        Exit Sub

ErrHandler:

    
        Exit Sub

SaveDayStats_Err:
100     Call RegistrarError(Err.Number, Err.Description, "General.SaveDayStats", Erl)
102     Resume Next
    
End Sub

Public Sub LogAsesinato(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal

102     Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & texto
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub logVentaCasa(ByVal texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal

102     Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
104     Print #nfile, "----------------------------------------------------------"
106     Print #nfile, Date & " " & Time & " " & texto
108     Print #nfile, "----------------------------------------------------------"
110     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogHackAttemp(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
104     Print #nfile, "----------------------------------------------------------"
106     Print #nfile, Date & " " & Time & " " & texto
108     Print #nfile, "----------------------------------------------------------"
110     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogCheating(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\CH.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & texto
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
104     Print #nfile, "----------------------------------------------------------"
106     Print #nfile, Date & " " & Time & " " & texto
108     Print #nfile, "----------------------------------------------------------"
110     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogAntiCheat(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & texto
106     Print #nfile, ""
108     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
        
        On Error GoTo ValidInputNP_Err
        

        Dim Arg As String

        Dim i   As Integer

100     For i = 1 To 33

102         Arg = ReadField(i, cad, 44)

104         If LenB(Arg) = 0 Then Exit Function

106     Next i

108     ValidInputNP = True

        
        Exit Function

ValidInputNP_Err:
110     Call RegistrarError(Err.Number, Err.Description, "General.ValidInputNP", Erl)
112     Resume Next
        
End Function

Sub Restart()
        
        On Error GoTo Restart_Err
        
        'Se asegura de que los sockets estan cerrados e ignora cualquier err

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

        Dim LoopC As Long

        'Cierra el socket de escucha
110     If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
        'Inicia el socket de escucha
112     SockListen = ListenForConnect(Puerto, hWndMsg, "")

114     For LoopC = 1 To MaxUsers
116         Call CloseSocket(LoopC)
        Next

        'Initialize statistics!!
        'Call Statistics.Initialize

118     For LoopC = 1 To UBound(UserList())
120         Set UserList(LoopC).incomingData = Nothing
122         Set UserList(LoopC).outgoingData = Nothing
124     Next LoopC

126     ReDim UserList(1 To MaxUsers) As user

128     For LoopC = 1 To MaxUsers
130         UserList(LoopC).ConnID = -1
132         UserList(LoopC).ConnIDValida = False
134         Set UserList(LoopC).incomingData = New clsByteQueue
136         Set UserList(LoopC).outgoingData = New clsByteQueue
138     Next LoopC

140     LastUser = 0
142     NumUsers = 0

144     Call FreeNPCs
146     Call FreeCharIndexes

148     Call LoadSini
150     Call LoadIntervalos
152     Call LoadOBJData
154     Call LoadPesca
156     Call LoadRecursosEspeciales

158     Call LoadMapData

160     Call CargarHechizos

188     If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

        'Log it
        Dim n As Integer

190     n = FreeFile
192     Open App.Path & "\logs\Main.log" For Append Shared As #n
194     Print #n, Date & " " & Time & " servidor reiniciado."
196     Close #n

        'Ocultar
        Call frmMain.InitMain(HideMe)
    
        Exit Sub

Restart_Err:
204     Call RegistrarError(Err.Number, Err.Description, "General.Restart", Erl)

        
End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo Intemperie_Err
        
    
100     If MapInfo(UserList(UserIndex).Pos.Map).zone <> "DUNGEON" Then
102         If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 1 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 2 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger < 10 Then Intemperie = True
        Else
104         Intemperie = False

        End If
    
        
        Exit Function

Intemperie_Err:
106     Call RegistrarError(Err.Number, Err.Description, "General.Intemperie", Erl)
108     Resume Next
        
End Function

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
        
        On Error GoTo TiempoInvocacion_Err
    
        
        Dim i As Integer
100     For i = 1 To MAXMASCOTAS
102         If UserList(UserIndex).MascotasIndex(i) > 0 Then
104             If NpcList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
106                NpcList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
                   NpcList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
108                If NpcList(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
                End If
            End If
110     Next i
        
        Exit Sub

TiempoInvocacion_Err:
112     Call RegistrarError(Err.Number, Err.Description, "General.TiempoInvocacion", Erl)

        
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoFrio_Err
        
100     If Not Intemperie(UserIndex) Then Exit Sub
        
102     With UserList(UserIndex)
            
104         If .Invent.ArmourEqpObjIndex > 0 Then
                ' WyroX: Ropa invernal
                If ObjData(.Invent.ArmourEqpObjIndex).Invernal Then Exit Sub
            End If
            
106         If .Counters.Frio < IntervaloFrio Then
108             .Counters.Frio = .Counters.Frio + 1

            Else

110             If MapInfo(.Pos.Map).terrain = Nieve Then
112                 Call WriteConsoleMsg(UserIndex, "¡Estás muriendo de frío, abrígate o morirás!", FontTypeNames.FONTTYPE_INFO)

                    ' WyroX: Sin ropa perdés vida más rápido que con una ropa no-invernal
                    Dim MinDaño As Integer, MaxDaño As Integer
                    If .flags.Desnudo = 0 Then
                        MinDaño = 17
                        MaxDaño = 23
                    Else
                        MinDaño = 27
                        MaxDaño = 33
                    End If

                    ' WyroX: Agrego aleatoriedad
                    Dim Daño As Integer
114                 Daño = Porcentaje(.Stats.MaxHp, RandomNumber(MinDaño, MaxDaño))

116                 .Stats.MinHp = .Stats.MinHp - Daño
            
118                 If .Stats.MinHp < 1 Then

120                     Call WriteConsoleMsg(UserIndex, "¡Has muerto de frío!", FontTypeNames.FONTTYPE_INFO)

122                     Call UserDie(UserIndex)

                    Else
124                     Call WriteUpdateHP(UserIndex)
                    End If
                End If
        
126             .Counters.Frio = 0

            End If
        
        End With
        
        Exit Sub

EfectoFrio_Err:
128     Call RegistrarError(Err.Number, Err.Description, "General.EfectoFrio", Erl)

130     Resume Next
        
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoLava_Err

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/12/07
        'If user is standing on lava, take health points from him
        '***************************************************
        
100     With UserList(UserIndex)
        
102         If .Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
104             .Counters.Lava = .Counters.Lava + 1
        
            Else

106             If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
108                 Call WriteConsoleMsg(UserIndex, "¡Quítate de la lava, te estás quemando!", FontTypeNames.FONTTYPE_INFO)
110                 .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHp, 5)
            
112                 If .Stats.MinHp < 1 Then
114                     Call WriteConsoleMsg(UserIndex, "¡Has muerto quemado!", FontTypeNames.FONTTYPE_INFO)
116                     Call UserDie(UserIndex)
                    Else
118                     Call WriteUpdateHP(UserIndex)
                    End If
                End If
        
120             .Counters.Lava = 0

            End If
        
        End With
        

        
        Exit Sub

EfectoLava_Err:
122     Call RegistrarError(Err.Number, Err.Description, "General.EfectoLava", Erl)

124     Resume Next
        
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
    '******************************************************
    'Author: Unknown
    'Last Update: 04/11/2008 (NicoNZ)
    '
    '******************************************************
        
        On Error GoTo EfectoMimetismo_Err
    
        
        Dim Barco As ObjData
    
100     With UserList(UserIndex)
102         If .Counters.Mimetismo < IntervaloInvisible Then
104             .Counters.Mimetismo = .Counters.Mimetismo + 1

            Else
                'restore old char
106             Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
            
108             If .flags.Navegando Then
                    Call EquiparBarco(UserIndex)
                Else
124                 .Char.Body = .CharMimetizado.Body
126                 .Char.Head = .CharMimetizado.Head
128                 .Char.CascoAnim = .CharMimetizado.CascoAnim
130                 .Char.ShieldAnim = .CharMimetizado.ShieldAnim
132                 .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                End If
                
134             .Counters.Mimetismo = 0
136             .flags.Mimetizado = e_EstadoMimetismo.Desactivado
            
138             With .Char
140                 Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
142                 Call RefreshCharStatus(UserIndex)
                End With
                
            End If
            
        End With
        
        Exit Sub

EfectoMimetismo_Err:
144     Call RegistrarError(Err.Number, Err.Description, "General.EfectoMimetismo", Erl)

        
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoInvisibilidad_Err
        

100     If UserList(UserIndex).Counters.Invisibilidad > 0 Then
102         UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad - 1
        Else
104         UserList(UserIndex).Counters.Invisibilidad = 0
106         UserList(UserIndex).flags.invisible = 0

108         If UserList(UserIndex).flags.Oculto = 0 Then
                ' Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
110             Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
114             Call WriteContadores(UserIndex)

            End If

        End If

        
        Exit Sub

EfectoInvisibilidad_Err:
116     Call RegistrarError(Err.Number, Err.Description, "General.EfectoInvisibilidad", Erl)
118     Resume Next
        
End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
        On Error GoTo EfectoParalisisNpc_Err
        
100     If NpcList(NpcIndex).Contadores.Paralisis > 0 Then
102         NpcList(NpcIndex).Contadores.Paralisis = NpcList(NpcIndex).Contadores.Paralisis - 1
        Else
104         NpcList(NpcIndex).flags.Paralizado = 0

        End If

        
        Exit Sub

EfectoParalisisNpc_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.EfectoParalisisNpc", Erl)
110     Resume Next
        
End Sub

Public Sub EfectoInmovilizadoNpc(ByVal NpcIndex As Integer)
        On Error GoTo EfectoInmovilizadoNpc_Err

100     If NpcList(NpcIndex).Contadores.Inmovilizado > 0 Then
102         NpcList(NpcIndex).Contadores.Inmovilizado = NpcList(NpcIndex).Contadores.Inmovilizado - 1
        Else
104         NpcList(NpcIndex).flags.Inmovilizado = 0

        End If

        Exit Sub

EfectoInmovilizadoNpc_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.EfectoInmovilizadoNpc", Erl)
110     Resume Next
        
End Sub


Public Sub EfectoCeguera(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoCeguera_Err
        

100     If UserList(UserIndex).Counters.Ceguera > 0 Then
102         UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
        Else

104         If UserList(UserIndex).flags.Ceguera = 1 Then
106             UserList(UserIndex).flags.Ceguera = 0
108             Call WriteBlindNoMore(UserIndex)

            End If

        End If

        
        Exit Sub

EfectoCeguera_Err:
110     Call RegistrarError(Err.Number, Err.Description, "General.EfectoCeguera", Erl)
112     Resume Next
        
End Sub

Public Sub EfectoEstupidez(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoEstupidez_Err
        

100     If UserList(UserIndex).Counters.Estupidez > 0 Then
102         UserList(UserIndex).Counters.Estupidez = UserList(UserIndex).Counters.Estupidez - 1

        Else

104         If UserList(UserIndex).flags.Estupidez = 1 Then
106             UserList(UserIndex).flags.Estupidez = 0
108             Call WriteDumbNoMore(UserIndex)

            End If

        End If

        
        Exit Sub

EfectoEstupidez_Err:
110     Call RegistrarError(Err.Number, Err.Description, "General.EfectoEstupidez", Erl)
112     Resume Next
        
End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoParalisisUser_Err
        

100     If UserList(UserIndex).Counters.Paralisis > 0 Then
102         UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
        Else
104         UserList(UserIndex).flags.Paralizado = 0
            'UserList(UserIndex).Flags.AdministrativeParalisis = 0
106         Call WriteParalizeOK(UserIndex)

        End If

        
        Exit Sub

EfectoParalisisUser_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.EfectoParalisisUser", Erl)
110     Resume Next
        
End Sub

Public Sub EfectoVelocidadUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoVelocidadUser_Err
        

100     If UserList(UserIndex).Counters.velocidad > 0 Then
102         UserList(UserIndex).Counters.velocidad = UserList(UserIndex).Counters.velocidad - 1
        Else
            UserList(UserIndex).flags.VelocidadHechizada = 0
            Call ActualizarVelocidadDeUsuario(UserIndex)
        End If

        Exit Sub

EfectoVelocidadUser_Err:
110     Call RegistrarError(Err.Number, Err.Description, "General.EfectoVelocidadUser", Erl)
112     Resume Next
        
End Sub

Public Sub EfectoMaldicionUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoMaldicionUser_Err
        

100     If UserList(UserIndex).Counters.Maldicion > 0 Then
102         UserList(UserIndex).Counters.Maldicion = UserList(UserIndex).Counters.Maldicion - 1
    
        Else
104         UserList(UserIndex).flags.Maldicion = 0
106         Call WriteConsoleMsg(UserIndex, "¡La magia perdió su efecto! Ya puedes atacar.", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
        End If

        
        Exit Sub

EfectoMaldicionUser_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.EfectoMaldicionUser", Erl)
110     Resume Next
        
End Sub

Public Sub EfectoInmoUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoInmoUser_Err
        

100     If UserList(UserIndex).Counters.Inmovilizado > 0 Then
102         UserList(UserIndex).Counters.Inmovilizado = UserList(UserIndex).Counters.Inmovilizado - 1
        Else
104         UserList(UserIndex).flags.Inmovilizado = 0
            'UserList(UserIndex).Flags.AdministrativeParalisis = 0
106         Call WriteInmovilizaOK(UserIndex)

        End If

        
        Exit Sub

EfectoInmoUser_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.EfectoInmoUser", Erl)
110     Resume Next
        
End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
        On Error GoTo RecStamina_Err

        Dim trigger As Byte
        Dim Suerte As Integer

        With UserList(UserIndex)
            trigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger

            If trigger = 1 And trigger = 2 And trigger = 4 Then Exit Sub

            If .Stats.MinSta < .Stats.MaxSta Then

                If .Counters.STACounter < Intervalo Then
                    .Counters.STACounter = .Counters.STACounter + 1
                    Exit Sub

                End If

                .Counters.STACounter = 0

                If .Counters.Trabajando > 0 Then Exit Sub  'Trabajando no sube energía. (ToxicWaste)

                EnviarStats = True

                Select Case .Stats.UserSkills(eSkill.Supervivencia)
                    Case 0 To 10
                        Suerte = 5
                    Case 11 To 20
                        Suerte = 7
                    Case 21 To 30
                        Suerte = 9
                    Case 31 To 40
                        Suerte = 11
                    Case 41 To 50
                        Suerte = 13
                    Case 51 To 60
                        Suerte = 15
                    Case 61 To 70
                        Suerte = 17
                    Case 71 To 80
                        Suerte = 19
                    Case 81 To 90
                        Suerte = 21
                    Case 91 To 99
                        Suerte = 23
                    Case 100
                        Suerte = 25
                End Select

                If .flags.RegeneracionSta = 1 Then Suerte = 45

                .Stats.MinSta = .Stats.MinSta + RandomNumber(1, Porcentaje(.Stats.MaxSta, Suerte))

                If .Stats.MinSta > .Stats.MaxSta Then
                    .Stats.MinSta = .Stats.MaxSta

                End If

            End If
        End With

        Exit Sub

RecStamina_Err:
        Call RegistrarError(Err.Number, Err.Description, "General.RecStamina", Erl)
        Resume Next

End Sub

Public Sub PierdeEnergia(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

        On Error GoTo RecStamina_Err

100     With UserList(UserIndex)

102         If .Stats.MinSta > 0 And .flags.RegeneracionSta = 0 Then
    
104             If .Counters.STACounter < Intervalo Then
106                 .Counters.STACounter = .Counters.STACounter + 1
                Else
            
108                 .Counters.STACounter = 0
    
110                 EnviarStats = True
            
                    Dim Cantidad As Integer
    
112                 Cantidad = RandomNumber(1, Porcentaje(.Stats.MaxSta, (MAXSKILLPOINTS * 1.5 - .Stats.UserSkills(eSkill.Supervivencia)) * 0.25))
114                 .Stats.MinSta = .Stats.MinSta - Cantidad
    
116                 If .Stats.MinSta < 0 Then
118                     .Stats.MinSta = 0
                    End If
    
                End If
    
            End If

        End With
        
        Exit Sub

RecStamina_Err:
120     Call RegistrarError(Err.Number, Err.Description, "General.PierdeEnergia", Erl)
122     Resume Next
        
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)

        On Error GoTo EfectoVeneno_Err

        Dim damage As Long

100     If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
102         UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
        Else
104         Call CancelExit(UserIndex)

            With UserList(UserIndex)
              'Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
              Call WriteLocaleMsg(UserIndex, "47", FontTypeNames.FONTTYPE_VENENO)
              Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, ParticulasIndex.Envenena, 30, False))
              .Counters.Veneno = 0

              ' El veneno saca un porcentaje de vida random.
              damage = RandomNumber(3, 5)
              damage = (1 + damage * .Stats.MaxHp \ 100) ' Redondea para arriba
              .Stats.MinHp = UserList(UserIndex).Stats.MinHp - damage

              If .ChatCombate = 1 Then
                  ' "El veneno te ha causado ¬1 puntos de daño."
                  Call WriteLocaleMsg(UserIndex, "390", FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(damage))
              End If

              If UserList(UserIndex).Stats.MinHp < 1 Then
                  Call UserDie(UserIndex)
              Else
                  Call WriteUpdateHP(UserIndex)
              End If
            End With

        End If

        Exit Sub

EfectoVeneno_Err:
124     Call RegistrarError(Err.Number, Err.Description, "General.EfectoVeneno", Erl)
126     Resume Next

End Sub

Public Sub EfectoAhogo(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoAhogo_Err
        

        Dim n As Integer

100     If RequiereOxigeno(UserList(UserIndex).Pos.Map) Then
102         If UserList(UserIndex).Counters.Ahogo < 70 Then
104             UserList(UserIndex).Counters.Ahogo = UserList(UserIndex).Counters.Ahogo + 1
            Else
106             Call WriteConsoleMsg(UserIndex, "Te estás ahogando, si no consigues oxígeno morirás.", FontTypeNames.FONTTYPE_EJECUCION)
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 205, 30, False))
108             UserList(UserIndex).Counters.Ahogo = 0
110             n = RandomNumber(150, 200)
112             UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - n

114             If UserList(UserIndex).Stats.MinHp < 1 Then
116                 Call UserDie(UserIndex)
                Else
120                 Call WriteUpdateHP(UserIndex)
                End If

            End If

        Else
122         UserList(UserIndex).flags.Ahogandose = 0

        End If

        
        Exit Sub

EfectoAhogo_Err:
124     Call RegistrarError(Err.Number, Err.Description, "General.EfectoAhogo", Erl)
126     Resume Next
        
End Sub

' El incineramiento tiene una logica particular, que es hacer daño sostenido en el tiempo.
Public Sub EfectoIncineramiento(ByVal UserIndex As Integer)
        On Error GoTo EfectoIncineramiento_Err

        Dim damage As Integer

        With UserList(UserIndex)

            ' 5 Mini intervalitos, dentro del intervalo total de incineracion
            If .Counters.Incineracion Mod (IntervaloIncineracion \ 5) = 0 Then
                ' "Te estás incinerando, si no te curas morirás.
                Call WriteLocaleMsg(UserIndex, "392", FontTypeNames.FONTTYPE_FIGHT)
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, ParticulasIndex.Incinerar, 30, False))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 73, 0))

                damage = RandomNumber(35, 45)
                .Stats.MinHp = .Stats.MinHp - damage

                If .ChatCombate = 1 Then
                    ' "El fuego te ha causado ¬1 puntos de daño."
                    Call WriteLocaleMsg(UserIndex, "391", FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(damage))
                End If

                If UserList(UserIndex).Stats.MinHp < 1 Then
                    Call UserDie(UserIndex)
                Else
                    Call WriteUpdateHP(UserIndex)
                End If
            End If

            .Counters.Incineracion = .Counters.Incineracion + 1

            If .Counters.Incineracion > IntervaloIncineracion Then
                ' Se termino la incineracion
                .flags.Incinerado = 0
                .Counters.Incineracion = 0
                Exit Sub

            End If
        End With

        Exit Sub

EfectoIncineramiento_Err:
        Call RegistrarError(Err.Number, Err.Description, "General.EfectoIncineramiento", Erl)
        Resume Next

End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)
        
        On Error GoTo DuracionPociones_Err
        

        'Controla la duracion de las pociones
100     If UserList(UserIndex).flags.DuracionEfecto > 0 Then
102         UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1

104         If UserList(UserIndex).flags.DuracionEfecto = 0 Then
106             UserList(UserIndex).flags.TomoPocion = False
108             UserList(UserIndex).flags.TipoPocion = 0

                'volvemos los atributos al estado normal
                Dim LoopX As Integer

110             For LoopX = 1 To NUMATRIBUTOS
112                 UserList(UserIndex).Stats.UserAtributos(LoopX) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopX)
                Next
114             Call WriteFYA(UserIndex)

            End If

        End If

        
        Exit Sub

DuracionPociones_Err:
116     Call RegistrarError(Err.Number, Err.Description, "General.DuracionPociones", Erl)
118     Resume Next
        
End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)
        
        On Error GoTo HambreYSed_Err
        

100     If Not UserList(UserIndex).flags.Privilegios And PlayerType.user Then Exit Sub

        'Sed
104     If UserList(UserIndex).Stats.MinAGU > 0 Then
106         If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
108             UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
            Else
110             UserList(UserIndex).Counters.AGUACounter = 0
112             UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
        
114             If UserList(UserIndex).Stats.MinAGU <= 0 Then
116                 UserList(UserIndex).Stats.MinAGU = 0
118                 UserList(UserIndex).flags.Sed = 1

                End If
        
120             fenviarAyS = True

            End If

        End If

        'hambre
122     If UserList(UserIndex).Stats.MinHam > 0 Then
124         If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
126             UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
            Else
128             UserList(UserIndex).Counters.COMCounter = 0
130             UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10

132             If UserList(UserIndex).Stats.MinHam <= 0 Then
134                 UserList(UserIndex).Stats.MinHam = 0
136                 UserList(UserIndex).flags.Hambre = 1

                End If

138             fenviarAyS = True

            End If

        End If

        
        Exit Sub

HambreYSed_Err:
140     Call RegistrarError(Err.Number, Err.Description, "General.HambreYSed", Erl)
142     Resume Next
        
End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
        
        On Error GoTo Sanar_Err
        
        ' Desnudo no regenera vida
100     If UserList(UserIndex).flags.Desnudo = 1 Then Exit Sub
        
102     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

        Dim mashit As Integer

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
104     If UserList(UserIndex).Stats.MinHp < UserList(UserIndex).Stats.MaxHp Then
106         If UserList(UserIndex).flags.RegeneracionHP = 1 Then
108             Intervalo = 400

            End If
    
110         If UserList(UserIndex).Counters.HPCounter < Intervalo Then
112             UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
            Else
114             mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
        
116             UserList(UserIndex).Counters.HPCounter = 0
118             UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp + mashit

120             If UserList(UserIndex).Stats.MinHp > UserList(UserIndex).Stats.MaxHp Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
122             Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
124             EnviarStats = True

            End If

        End If

        
        Exit Sub

Sanar_Err:
126     Call RegistrarError(Err.Number, Err.Description, "General.Sanar", Erl)
128     Resume Next
        
End Sub

Public Sub CargaNpcsDat(Optional ByVal ActualizarNPCsExistentes As Boolean = False)
        
            On Error GoTo CargaNpcsDat_Err
        
            ' Leemos el NPCs.dat y lo almacenamos en la memoria.
100         Set LeerNPCs = New clsIniReader
102         Call LeerNPCs.Initialize(DatPath & "NPCs.dat")
        
            ' Cargamos la lista de NPC's hostiles disponibles para spawnear.
104         Call CargarSpawnList
    
            ' Actualizamos la informacion de los NPC's ya spawneados.
106         If ActualizarNPCsExistentes Then
    
                Dim i As Long
108             For i = 1 To NumNPCs
    
110                 If NpcList(i).flags.NPCActive Then
112                     Call OpenNPC(CInt(i), False, True)
                    End If
    
114                 DoEvents
    
116             Next i
    
            End If
        
            Exit Sub

CargaNpcsDat_Err:
118         Call RegistrarError(Err.Number, Err.Description, "General.CargaNpcsDat", Erl)
120         Resume Next
        
End Sub

Sub PasarSegundo()

        On Error GoTo ErrHandler

        Dim i    As Long

        Dim h    As Byte

        Dim Mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
    
100     If CuentaRegresivaTimer > 0 Then
102         If CuentaRegresivaTimer > 1 Then
104             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(CuentaRegresivaTimer - 1 & " segundos...!", FontTypeNames.FONTTYPE_GUILD))
            Else
106             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ya!!!", FontTypeNames.FONTTYPE_FIGHT))

            End If

108         CuentaRegresivaTimer = CuentaRegresivaTimer - 1

        End If
    
110     For i = 1 To LastUser

112         If UserList(i).flags.Silenciado = 1 Then
114             UserList(i).flags.SegundosPasados = UserList(i).flags.SegundosPasados + 1

116             If UserList(i).flags.SegundosPasados = 60 Then
118                 UserList(i).flags.MinutosRestantes = UserList(i).flags.MinutosRestantes - 1
120                 UserList(i).flags.SegundosPasados = 0

                End If
            
122             If UserList(i).flags.MinutosRestantes = 0 Then
124                 UserList(i).flags.SegundosPasados = 0
126                 UserList(i).flags.Silenciado = 0
128                 UserList(i).flags.MinutosRestantes = 0
130                 Call WriteConsoleMsg(i, "Has sido liberado del silencio.", FontTypeNames.FONTTYPE_SERVER)

                End If

            End If

132         With UserList(i)
                
                Call DuracionPociones(i)
134             If .flags.invisible = 1 Then Call EfectoInvisibilidad(i)
138             If .flags.Paralizado = 1 Then Call EfectoParalisisUser(i)
140             If .flags.Inmovilizado = 1 Then Call EfectoInmoUser(i)
142             If .flags.Ceguera = 1 Then Call EfectoCeguera(i)
144             If .flags.Estupidez = 1 Then Call EfectoEstupidez(i)
146             If .flags.Maldicion = 1 Then Call EfectoMaldicionUser(i)
148             If .flags.VelocidadHechizada > 0 Then Call EfectoVelocidadUser(i)

150             If .flags.UltimoMensaje > 0 Then
                    .Counters.RepetirMensaje = .Counters.RepetirMensaje + 1
                    If .Counters.RepetirMensaje >= 3 Then
                        .flags.UltimoMensaje = 0
                        .Counters.RepetirMensaje = 0
                    End If
                End If
                
                If .Counters.CuentaRegresiva >= 0 Then
                    If .Counters.CuentaRegresiva > 0 Then
                        Call WriteConsoleMsg(i, ">>>  " & .Counters.CuentaRegresiva & "  <<<", FontTypeNames.FONTTYPE_New_Gris)
                    Else
                        Call WriteConsoleMsg(i, ">>> YA! <<<", FontTypeNames.FONTTYPE_FIGHT)
                        Call WriteStopped(i, False)
                    End If
                    
                    .Counters.CuentaRegresiva = .Counters.CuentaRegresiva - 1
                End If
        
            End With
        
152         If UserList(i).flags.Portal > 1 Then
154             UserList(i).flags.Portal = UserList(i).flags.Portal - 1
        
156             If UserList(i).flags.Portal = 1 Then
158                 Mapa = UserList(i).flags.PortalM
160                 X = UserList(i).flags.PortalX
162                 Y = UserList(i).flags.PortalY
164                 Call SendData(SendTarget.toMap, UserList(i).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.TpVerde, 0))
166                 Call SendData(SendTarget.toMap, UserList(i).flags.PortalM, PrepareMessageLightFXToFloor(X, Y, 0, 105))

168                 If MapData(Mapa, X, Y).TileExit.Map > 0 Then
170                     MapData(Mapa, X, Y).TileExit.Map = 0
172                     MapData(Mapa, X, Y).TileExit.X = 0
174                     MapData(Mapa, X, Y).TileExit.Y = 0

                    End If

176                 MapData(Mapa, X, Y).Particula = 0
178                 MapData(Mapa, X, Y).TimeParticula = 0
180                 MapData(Mapa, X, Y).Particula = 0
182                 MapData(Mapa, X, Y).TimeParticula = 0
184                 UserList(i).flags.Portal = 0
186                 UserList(i).flags.PortalM = 0
188                 UserList(i).flags.PortalY = 0
190                 UserList(i).flags.PortalX = 0
192                 UserList(i).flags.PortalMDestino = 0
194                 UserList(i).flags.PortalYDestino = 0
196                 UserList(i).flags.PortalXDestino = 0

                End If

            End If
        
198         If UserList(i).Counters.TiempoDeMapeo > 0 Then
200             UserList(i).Counters.TiempoDeMapeo = UserList(i).Counters.TiempoDeMapeo - 1
            End If
        
        
202         If UserList(i).Counters.TiempoDeInmunidad > 0 Then
204             UserList(i).Counters.TiempoDeInmunidad = UserList(i).Counters.TiempoDeInmunidad - 1
206             If UserList(i).Counters.TiempoDeInmunidad = 0 Then
208                 UserList(i).flags.Inmunidad = 0
                End If
            End If
        
210         If UserList(i).flags.Subastando Then
212             UserList(i).Counters.TiempoParaSubastar = UserList(i).Counters.TiempoParaSubastar - 1

214             If UserList(i).Counters.TiempoParaSubastar = 0 Then
216                 Call CancelarSubasta

                End If

            End If

218         If UserList(i).flags.UserLogged Then

                'Cerrar usuario
220             If UserList(i).Counters.Saliendo Then
                    '  If UserList(i).flags.Muerto = 1 Then UserList(i).Counters.Salir = 0
222                 UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                    ' Call WriteConsoleMsg(i, "Se saldrá del juego en " & UserList(i).Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
224                 Call WriteLocaleMsg(i, "203", FontTypeNames.FONTTYPE_INFO, UserList(i).Counters.Salir)

226                 If UserList(i).Counters.Salir <= 0 Then
228                     Call WriteConsoleMsg(i, "Gracias por jugar Argentum 20.", FontTypeNames.FONTTYPE_INFO)
230                     Call WriteDisconnect(i)
                    
232                     Call CloseSocket(i)

                    End If

                End If

            End If

234     Next i

        ' **********************************
        ' **********  Invasiones  **********
        ' **********************************
        For i = 1 To UBound(Invasiones)
            With Invasiones(i)

                ' Si la invasión está activa
                If .Activa Then
                    .TimerSpawn = .TimerSpawn + 1

                    ' Comprobamos si hay que spawnear NPCs
                    If .TimerSpawn >= .IntervaloSpawn Then
                        Call InvasionSpawnNPC(i)
                        .TimerSpawn = 0
                    End If
                    
                    ' ------------------------------------
                    
                    .TimerMostrarInfo = .TimerMostrarInfo + 1
                    
                    ' Comprobamos si hay que mostrar la info
                    If .TimerMostrarInfo >= 5 Then
                        Call EnviarInfoInvasion(i)
                        .TimerMostrarInfo = 0
                    End If
                End If
            
            End With
        Next
        ' **********************************

        Exit Sub

ErrHandler:
236     Call LogError("Error en PasarSegundo. Err: " & Err.Description & " - " & Err.Number & " - UserIndex: " & i)

238     Resume Next

End Sub
 
Public Function ReiniciarAutoUpdate() As Double
        
        On Error GoTo ReiniciarAutoUpdate_Err
        

100     ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

        
        Exit Function

ReiniciarAutoUpdate_Err:
102     Call RegistrarError(Err.Number, Err.Description, "General.ReiniciarAutoUpdate", Erl)
104     Resume Next
        
End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
        'WorldSave
        
        On Error GoTo ReiniciarServidor_Err
        
100     Call DoBackUp

        'Guardar Pjs
102     Call GuardarUsuarios
    
104     If EjecutarLauncher Then Shell App.Path & "\launcher.exe" & " megustalanoche*"

        'Chauuu
106     Unload frmMain

        
        Exit Sub

ReiniciarServidor_Err:
108     Call RegistrarError(Err.Number, Err.Description, "General.ReiniciarServidor", Erl)
110     Resume Next
        
End Sub
 
Sub GuardarUsuarios()
        
        On Error GoTo GuardarUsuarios_Err
        
100     haciendoBK = True
    
102     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
104     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
        Dim i As Long
        
106     For i = 1 To LastUser

108         If UserList(i).flags.UserLogged Then
110             Call FlushBuffer(i)
            End If

112     Next i

114     For i = 1 To LastUser

116         If UserList(i).flags.UserLogged Then
120             Call SaveUser(i)

            End If

122     Next i
    
124     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
126     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

128     haciendoBK = False

        
        Exit Sub

GuardarUsuarios_Err:
130     Call RegistrarError(Err.Number, Err.Description, "General.GuardarUsuarios", Erl)
132     Resume Next
        
End Sub

Sub InicializaEstadisticas()
        
        On Error GoTo InicializaEstadisticas_Err
        

        Dim Ta As Long

100     Ta = GetTickCount()

102     Call EstadisticasWeb.Inicializa(frmMain.hwnd)
104     Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
106     Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
108     Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
110     Call EstadisticasWeb.Informar(RECORD_USUARIOS, RecordUsuarios)

        
        Exit Sub

InicializaEstadisticas_Err:
112     Call RegistrarError(Err.Number, Err.Description, "General.InicializaEstadisticas", Erl)
114     Resume Next
        
End Sub

Public Sub FreeNPCs()
        
        On Error GoTo FreeNPCs_Err
        

        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all NPC Indexes
        '***************************************************
        Dim LoopC As Long
    
        ' Free all NPC indexes
100     For LoopC = 1 To MaxNPCs
102         NpcList(LoopC).flags.NPCActive = False
104     Next LoopC

        
        Exit Sub

FreeNPCs_Err:
106     Call RegistrarError(Err.Number, Err.Description, "General.FreeNPCs", Erl)
108     Resume Next
        
End Sub

Public Sub FreeCharIndexes()
        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all char indexes
        '***************************************************
        ' Free all char indexes (set them all to 0)
        
        On Error GoTo FreeCharIndexes_Err
        
100     Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))

        
        Exit Sub

FreeCharIndexes_Err:
102     Call RegistrarError(Err.Number, Err.Description, "General.FreeCharIndexes", Erl)
104     Resume Next
        
End Sub

Function RandomString(cb As Integer, Optional ByVal OnlyUpper As Boolean = False) As String
        
        On Error GoTo RandomString_Err
        

100     Randomize Time

        Dim rgch As String

102     rgch = "abcdefghijklmnopqrstuvwxyz"
    
104     If OnlyUpper Then
106         rgch = UCase(rgch)
        Else
108         rgch = rgch & UCase(rgch)

        End If
    
110     rgch = rgch & "0123456789"  ' & "#@!~$()-_"

        Dim i As Long

112     For i = 1 To cb
114         RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Next

        
        Exit Function

RandomString_Err:
116     Call RegistrarError(Err.Number, Err.Description, "General.RandomString", Erl)
118     Resume Next
        
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

        On Error GoTo errHnd

        Dim lPos As Long

        Dim lX   As Long

        Dim iAsc As Integer
    
        '1er test: Busca un simbolo @
100     lPos = InStr(sString, "@")

102     If (lPos <> 0) Then

            '2do test: Busca un simbolo . después de @ + 1
104         If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
            '3er test: Recorre todos los caracteres y los valída
106         For lX = 0 To Len(sString) - 1

108             If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
110                 iAsc = Asc(mid$(sString, (lX + 1), 1))

112                 If Not CMSValidateChar_(iAsc) Then Exit Function

                End If

114         Next lX
        
            'Finale
116         CheckMailString = True

        End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
        
        On Error GoTo CMSValidateChar__Err
        
100     CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

        
        Exit Function

CMSValidateChar__Err:
102     Call RegistrarError(Err.Number, Err.Description, "General.CMSValidateChar_", Erl)
104     Resume Next
        
End Function

Public Function Tilde(ByRef data As String) As String
    
        On Error GoTo Tilde_Err
    

100     Tilde = UCase$(data)
 
102     Tilde = Replace$(Tilde, "Á", "A")
104     Tilde = Replace$(Tilde, "É", "E")
106     Tilde = Replace$(Tilde, "Í", "I")
108     Tilde = Replace$(Tilde, "Ó", "O")
110     Tilde = Replace$(Tilde, "Ú", "U")
        
    
        Exit Function

Tilde_Err:
112     Call RegistrarError(Err.Number, Err.Description, "Mod_General.Tilde", Erl)
114     Resume Next
    
End Function

Public Sub CerrarServidor()
        
    'Save stats!!!
    Call Statistics.DumpStatistics
    Call frmMain.QuitarIconoSystray
    
    ' Limpieza del socket del servidor.
    Call LimpiaWsApi
    
    Dim LoopC As Long
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).ConnID <> -1 Then
            Call CloseSocket(LoopC)
        End If
    Next
    
    If Database_Enabled Then Database_Close
 
    Call LimpiarModuloLimpieza
    
    'Log
    Dim n As Integer: n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & Time & " server cerrado."
    Close #n
    
    End
   
End Sub

Function max(ByVal a As Double, ByVal b As Double) As Double
        
        On Error GoTo max_Err
    
        

100     If a > b Then
102         max = a
        Else
104         max = b
        End If

        
        Exit Function

max_Err:
106     Call RegistrarError(Err.Number, Err.Description, "General.max", Erl)

        
End Function

Function min(ByVal a As Double, ByVal b As Double) As Double
        
        On Error GoTo min_Err
    
        

100     If a < b Then
102         min = a
        Else
104         min = b
        End If

        
        Exit Function

min_Err:
106     Call RegistrarError(Err.Number, Err.Description, "General.min", Erl)

        
End Function

Public Function PonerPuntos(ByVal Numero As Long) As String
    
        On Error GoTo PonerPuntos_Err
    

        Dim i     As Integer

        Dim Cifra As String
 
100     Cifra = str(Numero)
102     Cifra = Right$(Cifra, Len(Cifra) - 1)

104     For i = 0 To 4

106         If Len(Cifra) - 3 * i >= 3 Then
108             If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
110                 PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos

                End If

            Else

112             If Len(Cifra) - 3 * i > 0 Then
114                 PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos

                End If

                Exit For

            End If

        Next
 
116     PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
 
    
        Exit Function

PonerPuntos_Err:
118     Call RegistrarError(Err.Number, Err.Description, "ModLadder.PonerPuntos", Erl)
120     Resume Next
    
End Function

' Autor: WyroX
Function CalcularPromedioVida(ByVal UserIndex As Integer) As Double

100     With UserList(UserIndex)
102         If .Stats.ELV = 1 Then
                ' Siempre estamos promedio al lvl 1
104             CalcularPromedioVida = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
            Else
106             CalcularPromedioVida = (.Stats.MaxHp - .Stats.UserAtributos(eAtributos.Constitucion)) / (.Stats.ELV - 1)
            End If
        End With

End Function

' Adaptado desde https://stackoverflow.com/questions/29325069/how-to-generate-random-numbers-biased-towards-one-value-in-a-range/29325222#29325222
' By WyroX
Function RandomIntBiased(ByVal min As Double, ByVal max As Double, ByVal Bias As Double, ByVal Influence As Double) As Double

        On Error GoTo handle

        Dim RandomRango As Double, Mix As Double
    
100     RandomRango = Rnd * (max - min) + min
102     Mix = Rnd * Influence
    
104     RandomIntBiased = RandomRango * (1 - Mix) + Bias * Mix
    
        Exit Function
    
handle:
106     Call RegistrarError(Err.Number, Err.Description, "General.RandomIntBiased")
108     RandomIntBiased = Bias

End Function

'Very efficient function for testing whether this code is running in the IDE or compiled
'https://www.vbforums.com/showthread.php?231468-VB-Detect-if-you-are-running-in-the-IDE&p=5413357&viewfull=1#post5413357
Public Function RunningInVB(Optional ByRef b As Boolean = True) As Boolean
    If b Then Debug.Assert Not RunningInVB(RunningInVB) Else b = True
End Function

' WyroX: Mensaje a todo el mundo
Public Sub MensajeGlobal(texto As String, Fuente As FontTypeNames)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(texto, Fuente))
End Sub

' WyroX: Devuelve si X e Y están dentro del Rectangle
Public Function InsideRectangle(R As Rectangle, ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X < R.X1 Then Exit Function
    If X > R.X2 Then Exit Function
    If Y < R.Y1 Then Exit Function
    If Y > R.Y2 Then Exit Function
    InsideRectangle = True
End Function

' Fuente: https://stackoverflow.com/questions/1378604/end-process-from-task-manager-using-vb-6-code (ultima respuesta)
Public Function GetProcess(ByVal processName As String) As Byte
    
    Dim oService As Object
    Dim servicename As String
    Dim processCount As Byte
    
    Dim oWMI As Object: Set oWMI = GetObject("winmgmts:")
    Dim oServices As Object: Set oServices = oWMI.InstancesOf("win32_process")

    For Each oService In oServices

        servicename = LCase$(Trim$(CStr(oService.name)))

        If InStrB(1, servicename, LCase$(processName), vbBinaryCompare) > 0 Then
            
            ' Para matar un proceso adentro de este loop usar.
            'oService.Terminate
            
            processCount = processCount + 1
            
        End If

    Next
    
    GetProcess = processCount

End Function
