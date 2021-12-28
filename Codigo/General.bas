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

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Global LeerNPCs As New clsIniManager

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer)
        
        On Error GoTo DarCuerpoDesnudo_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/14/07
        'Da cuerpo desnudo a un usuario
        '***************************************************
        Dim CuerpoDesnudo As Integer

100     Select Case UserList(UserIndex).genero

            Case e_Genero.Hombre

102             Select Case UserList(UserIndex).raza

                    Case e_Raza.Humano
104                     CuerpoDesnudo = 21 'ok

106                 Case e_Raza.Drow
108                     CuerpoDesnudo = 32 ' ok

110                 Case e_Raza.Elfo
112                     CuerpoDesnudo = 510 'Revisar

114                 Case e_Raza.Gnomo
116                     CuerpoDesnudo = 508 'Revisar

118                 Case e_Raza.Enano
120                     CuerpoDesnudo = 53 'ok

122                 Case e_Raza.Orco
124                     CuerpoDesnudo = 248 ' ok

                End Select

126         Case e_Genero.Mujer

128             Select Case UserList(UserIndex).raza

                    Case e_Raza.Humano
130                     CuerpoDesnudo = 39 'ok

132                 Case e_Raza.Drow
134                     CuerpoDesnudo = 40 'ok

136                 Case e_Raza.Elfo
138                     CuerpoDesnudo = 511 'Revisar

140                 Case e_Raza.Gnomo
142                     CuerpoDesnudo = 509 'Revisar

144                 Case e_Raza.Enano
146                     CuerpoDesnudo = 60 ' ok

148                 Case e_Raza.Orco
150                     CuerpoDesnudo = 249 'ok

                End Select

        End Select

152     UserList(UserIndex).Char.Body = CuerpoDesnudo

154     UserList(UserIndex).flags.Desnudo = 1

        
        Exit Sub

DarCuerpoDesnudo_Err:
156     Call TraceError(Err.Number, Err.Description, "General.DarCuerpoDesnudo", Erl)

        
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
100     b = b And e_Block.ALL_SIDES

102     If toMap Then
104         Call SendData(SendTarget.toMap, sndIndex, PrepareMessage_BlockPosition(X, Y, b))
        Else
106         Call Write_BlockPosition(sndIndex, X, Y, b)
        End If

        
        Exit Sub

Bloquear_Err:
108     Call TraceError(Err.Number, Err.Description, "General.Bloquear", Erl)

        
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
        
106     ModPuerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo

108     Select Case ModPuerta
        
            Case 0
                ' Bloqueos superiores
110             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
112             Call Bloquear(toMap, sndIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)

                ' Bloqueos inferiores
114             Call Bloquear(toMap, sndIndex, X, Y + 1, MapData(Map, X, Y + 1).Blocked)
116             Call Bloquear(toMap, sndIndex, X - 1, Y + 1, MapData(Map, X - 1, Y + 1).Blocked)

118         Case 1
                ' para palancas o teclas sin modicar bloqueos en X,Y
                
120         Case 2
                ' Bloqueos superiores
122             Call Bloquear(toMap, sndIndex, X, Y - 1, MapData(Map, X, Y - 1).Blocked)
124             Call Bloquear(toMap, sndIndex, X - 1, Y - 1, MapData(Map, X - 1, Y - 1).Blocked)
126             Call Bloquear(toMap, sndIndex, X + 1, Y - 1, MapData(Map, X + 1, Y - 1).Blocked)
                ' Bloqueos inferiores
128             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
130             Call Bloquear(toMap, sndIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
132             Call Bloquear(toMap, sndIndex, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                
134         Case 3
                ' Bloqueos superiores
136             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
138             Call Bloquear(toMap, sndIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
140             Call Bloquear(toMap, sndIndex, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                ' Bloqueos inferiores
142             Call Bloquear(toMap, sndIndex, X, Y + 1, MapData(Map, X, Y + 1).Blocked)
144             Call Bloquear(toMap, sndIndex, X - 1, Y + 1, MapData(Map, X - 1, Y + 1).Blocked)
146             Call Bloquear(toMap, sndIndex, X + 1, Y + 1, MapData(Map, X + 1, Y + 1).Blocked)

148         Case 4
                ' Bloqueos superiores
150             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
                ' Bloqueos inferiores
152             Call Bloquear(toMap, sndIndex, X, Y + 1, MapData(Map, X, Y + 1).Blocked)

154         Case 5 'Ver WyroX
                ' Bloqueos vertical ver ReyarB
156             Call Bloquear(toMap, sndIndex, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
158             Call Bloquear(toMap, sndIndex, X + 1, Y - 1, MapData(Map, X + 1, Y - 1).Blocked)

                ' Bloqueos horizontal
160             Call Bloquear(toMap, sndIndex, X, Y - 2, MapData(Map, X, Y - 2).Blocked)
162             Call Bloquear(toMap, sndIndex, X - 1, Y - 2, MapData(Map, X - 1, Y - 2).Blocked)


164         Case 6 ' Ver WyroX
                ' Bloqueos superiores ver ReyarB
166             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
168             Call Bloquear(toMap, sndIndex, X, Y - 1, MapData(Map, X, Y - 1).Blocked)

                ' Bloqueos inferiores
170             Call Bloquear(toMap, sndIndex, X, Y - 2, MapData(Map, X, Y - 2).Blocked)
172             Call Bloquear(toMap, sndIndex, X + 1, Y - 2, MapData(Map, X + 1, Y - 2).Blocked)

        End Select

        Exit Sub

MostrarBloqueosPuerta_Err:
174     Call TraceError(Err.Number, Err.Description, "General.MostrarBloqueosPuerta", Erl)
        
End Sub

Sub BloquearPuerta(ByVal Map As Integer, _
                   ByVal X As Integer, _
                   ByVal Y As Integer, _
                   ByVal Bloquear As Boolean)
        
        On Error GoTo BloquearPuerta_Err
        Dim ModPuerta As Integer
        
        'ver reyarb
100     ModPuerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo

102     Select Case ModPuerta
        
            Case 0 'puerta 2 tiles

                ' Bloqueos superiores
104             MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or e_Block.NORTH, MapData(Map, X, Y).Blocked And Not e_Block.NORTH)
106             MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or e_Block.NORTH, MapData(Map, X - 1, Y).Blocked And Not e_Block.NORTH)

                ' Cambio bloqueos inferiores
108             MapData(Map, X, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X, Y + 1).Blocked And Not e_Block.SOUTH)
110             MapData(Map, X - 1, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X - 1, Y + 1).Blocked And Not e_Block.SOUTH)

112         Case 1
                ' para palancas o teclas sin modicar bloqueos en X,Y

114         Case 2 ' puerta 3 tiles 1 arriba
                ' Bloqueos superiores
116             MapData(Map, X, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X, Y - 1).Blocked Or e_Block.NORTH, MapData(Map, X, Y - 1).Blocked And Not e_Block.NORTH)
118             MapData(Map, X - 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y - 1).Blocked Or e_Block.NORTH, MapData(Map, X - 1, Y - 1).Blocked And Not e_Block.NORTH)
120             MapData(Map, X + 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y - 1).Blocked Or e_Block.NORTH, MapData(Map, X + 1, Y - 1).Blocked And Not e_Block.NORTH)
                ' Cambio bloqueos inferiores
122             MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or e_Block.SOUTH, MapData(Map, X, Y).Blocked And Not e_Block.SOUTH)
124             MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or e_Block.SOUTH, MapData(Map, X - 1, Y).Blocked And Not e_Block.SOUTH)
126             MapData(Map, X + 1, Y).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y).Blocked Or e_Block.SOUTH, MapData(Map, X + 1, Y).Blocked And Not e_Block.SOUTH)
                
128         Case 3 ' puerta 3 tiles
                ' Bloqueos superiores
130             MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or e_Block.NORTH, MapData(Map, X, Y).Blocked And Not e_Block.NORTH)
132             MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or e_Block.NORTH, MapData(Map, X - 1, Y).Blocked And Not e_Block.NORTH)
134             MapData(Map, X + 1, Y).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y).Blocked Or e_Block.NORTH, MapData(Map, X + 1, Y).Blocked And Not e_Block.NORTH)
                ' Cambio bloqueos inferiores
136             MapData(Map, X, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X, Y + 1).Blocked And Not e_Block.SOUTH)
138             MapData(Map, X - 1, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X - 1, Y + 1).Blocked And Not e_Block.SOUTH)
140             MapData(Map, X + 1, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X + 1, Y + 1).Blocked And Not e_Block.SOUTH)
        
142         Case 4 'puerta 1 tiles
                ' Bloqueos superiores
144             MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or e_Block.NORTH, MapData(Map, X, Y).Blocked And Not e_Block.NORTH)
                ' Cambio bloqueos inferiores
146             MapData(Map, X, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X, Y + 1).Blocked And Not e_Block.SOUTH)
                
148         Case 5 'Ver WyroX
                ' Bloqueos  vertical ver ReyarB
150             MapData(Map, X + 1, Y).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y).Blocked Or e_Block.ALL_SIDES, MapData(Map, X + 1, Y).Blocked And Not e_Block.ALL_SIDES)
152             MapData(Map, X + 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y - 1).Blocked Or e_Block.ALL_SIDES, MapData(Map, X + 1, Y - 1).Blocked And Not e_Block.ALL_SIDES)
                
                ' Cambio horizontal
154             MapData(Map, X, Y - 2).Blocked = IIf(Bloquear, MapData(Map, X, Y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, X, Y - 2).Blocked And Not e_Block.ALL_SIDES)
156             MapData(Map, X - 1, Y - 2).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, X - 1, Y - 2).Blocked And Not e_Block.ALL_SIDES)


158         Case 6 ' Ver Wyrox
                ' Bloqueos vertical ver ReyarB
160             MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or e_Block.ALL_SIDES, MapData(Map, X - 1, Y).Blocked And Not e_Block.ALL_SIDES)
162             MapData(Map, X - 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y - 1).Blocked Or e_Block.ALL_SIDES, MapData(Map, X - 1, Y - 1).Blocked And Not e_Block.ALL_SIDES)
                
                ' Cambio bloqueos Puerta abierta
164             MapData(Map, X, Y - 2).Blocked = IIf(Bloquear, MapData(Map, X, Y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, X, Y - 2).Blocked And Not e_Block.ALL_SIDES)
166             MapData(Map, X + 1, Y + 2).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, X + 1, Y - 2).Blocked And Not e_Block.ALL_SIDES)

                
        End Select

        ' Mostramos a todos
168     Call MostrarBloqueosPuerta(True, Map, X, Y)
        
        Exit Sub

BloquearPuerta_Err:
170     Call TraceError(Err.Number, Err.Description, "General.BloquearPuerta", Erl)
        
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
110     Call TraceError(Err.Number, Err.Description, "General.HayCosta", Erl)

        
End Function

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayAgua_Err
        

100     With MapData(Map, X, Y)
102         If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
104             HayAgua = (.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
                    (.Graphic(1) >= 124 And .Graphic(1) <= 139) Or _
                    (.Graphic(1) >= 24223 And .Graphic(1) <= 24238) Or _
                    (.Graphic(1) >= 24303 And .Graphic(1) <= 24318) Or _
                    (.Graphic(1) >= 468 And .Graphic(1) <= 483) Or _
                    (.Graphic(1) >= 44668 And .Graphic(1) <= 44683) Or _
                    (.Graphic(1) >= 24143 And .Graphic(1) <= 24158) Or _
                    (.Graphic(1) >= 12628 And .Graphic(1) <= 12643) Or _
                    (.Graphic(1) >= 2948 And .Graphic(1) <= 2963)
            Else
106             HayAgua = False
    
            End If
        End With

        
        Exit Function

HayAgua_Err:
108     Call TraceError(Err.Number, Err.Description, "General.HayAgua", Erl)

        
End Function

Function EsArbol(ByVal GrhIndex As Long) As Boolean
        
        On Error GoTo EsArbol_Err
    
        
100     EsArbol = GrhIndex = 7000 Or GrhIndex = 7001 Or GrhIndex = 7002 Or GrhIndex = 641 Or GrhIndex = 26075 Or GrhIndex = 643 Or GrhIndex = 644 Or _
           GrhIndex = 647 Or GrhIndex = 26076 Or GrhIndex = 7020 Or _
           GrhIndex = 11903 Or GrhIndex = 11904 Or GrhIndex = 11905 Or GrhIndex = 11906 Or GrhIndex = 12160 Or GrhIndex = 7222 Or GrhIndex = 7223 Or _
           GrhIndex = 7224 Or GrhIndex = 7225 Or GrhIndex = 7226 Or _
           GrhIndex = 26077 Or GrhIndex = 26079 Or GrhIndex = 735 Or GrhIndex = 32343 Or GrhIndex = 32344 Or GrhIndex = 26080 Or GrhIndex = 26081 Or _
           GrhIndex = 32345 Or GrhIndex = 32346 Or GrhIndex = 32347 Or GrhIndex = 32348 Or GrhIndex = 32349 Or GrhIndex = 32350 Or GrhIndex = 32351 Or _
           GrhIndex = 32352 Or GrhIndex = 14961 Or GrhIndex = 14950 Or GrhIndex = 14951 Or GrhIndex = 14952 Or GrhIndex = 14953 Or GrhIndex = 14954 Or _
           GrhIndex = 14955 Or GrhIndex = 14956 Or GrhIndex = 14957 Or GrhIndex = 14958 Or GrhIndex = 14959 Or GrhIndex = 14962 Or GrhIndex = 14963 Or _
           GrhIndex = 14964 Or GrhIndex = 14967 Or GrhIndex = 14968 Or GrhIndex = 14969 Or GrhIndex = 14970 Or GrhIndex = 14971 Or GrhIndex = 14972 Or _
           GrhIndex = 14973 Or GrhIndex = 14974 Or GrhIndex = 14975 Or GrhIndex = 14976 Or GrhIndex = 14978 Or GrhIndex = 14980 Or GrhIndex = 14982 Or _
           GrhIndex = 14983 Or GrhIndex = 14984 Or GrhIndex = 14985 Or GrhIndex = 14987 Or GrhIndex = 14988 Or GrhIndex = 26078 Or GrhIndex = 26192 Or _
           GrhIndex = 15698 Or GrhIndex = 14504 Or GrhIndex = 15697 Or _
           GrhIndex = 15110 Or GrhIndex = 15109 Or GrhIndex = 15108 Or _
           GrhIndex = 7220 Or GrhIndex = 50990 Or GrhIndex = 55626 Or GrhIndex = 55627 Or GrhIndex = 55630 Or GrhIndex = 55632 Or GrhIndex = 55633 Or GrhIndex = 55635 Or GrhIndex = 55638 Or _
           GrhIndex = 463 Or GrhIndex = 1880 Or GrhIndex = 1121 Or GrhIndex = 1878 Or GrhIndex = 12584 Or GrhIndex = 50985 Or GrhIndex = 15510 Or GrhIndex = 14775 Or GrhIndex = 14687 Or _
           GrhIndex = 9513 Or GrhIndex = 9514 Or GrhIndex = 9515 Or GrhIndex = 9518 Or GrhIndex = 9519 Or GrhIndex = 9520 Or GrhIndex = 9529

        
        Exit Function

EsArbol_Err:
102     Call TraceError(Err.Number, Err.Description, "General.EsArbol", Erl)

        
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
110     Call TraceError(Err.Number, Err.Description, "General.HayLava", Erl)

        
End Function

Sub ApagarFogatas()

        'Ladder /ApagarFogatas
        On Error GoTo ErrHandler

        Dim obj As t_Obj
100         obj.ObjIndex = FOGATA_APAG
102         obj.amount = 1

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


Private Sub InicializarConstantes()
        
        On Error GoTo InicializarConstantes_Err
    
        
    
100     LastBackup = Format(Now, "Short Time")
102     minutos = Format(Now, "Short Time")
    
104     IniPath = App.Path & "\"

106     ListaRazas(e_Raza.Humano) = "Humano"
108     ListaRazas(e_Raza.Elfo) = "Elfo"
110     ListaRazas(e_Raza.Drow) = "Elfo Oscuro"
112     ListaRazas(e_Raza.Gnomo) = "Gnomo"
114     ListaRazas(e_Raza.Enano) = "Enano"
        ListaRazas(e_Raza.Orco) = "Orco"
    
116     ListaClases(e_Class.Mage) = "Mago"
118     ListaClases(e_Class.Cleric) = "Clérigo"
120     ListaClases(e_Class.Warrior) = "Guerrero"
122     ListaClases(e_Class.Assasin) = "Asesino"
124     ListaClases(e_Class.Bard) = "Bardo"
126     ListaClases(e_Class.Druid) = "Druida"
128     ListaClases(e_Class.Paladin) = "Paladín"
130     ListaClases(e_Class.Hunter) = "Cazador"
132     ListaClases(e_Class.Trabajador) = "Trabajador"
134     ListaClases(e_Class.Pirat) = "Pirata"
136     ListaClases(e_Class.Thief) = "Ladrón"
138     ListaClases(e_Class.Bandit) = "Bandido"
    
140     SkillsNames(e_Skill.Magia) = "Magia"
142     SkillsNames(e_Skill.Robar) = "Robar"
144     SkillsNames(e_Skill.Tacticas) = "Destreza en combate"
146     SkillsNames(e_Skill.Armas) = "Combate con armas"
148     SkillsNames(e_Skill.Meditar) = "Meditar"
150     SkillsNames(e_Skill.Apuñalar) = "Apuñalar"
152     SkillsNames(e_Skill.Ocultarse) = "Ocultarse"
154     SkillsNames(e_Skill.Supervivencia) = "Supervivencia"
156     SkillsNames(e_Skill.Comerciar) = "Comercio"
158     SkillsNames(e_Skill.Defensa) = "Defensa con escudo"
160     SkillsNames(e_Skill.liderazgo) = "Liderazgo"
162     SkillsNames(e_Skill.Proyectiles) = "Armas a distancia"
164     SkillsNames(e_Skill.Wrestling) = "Combate sin armas"
166     SkillsNames(e_Skill.Navegacion) = "Navegación"
168     SkillsNames(e_Skill.equitacion) = "Equitación"
170     SkillsNames(e_Skill.Resistencia) = "Resistencia mágica"
172     SkillsNames(e_Skill.Talar) = "Tala"
174     SkillsNames(e_Skill.Pescar) = "Pesca"
176     SkillsNames(e_Skill.Mineria) = "Minería"
178     SkillsNames(e_Skill.Herreria) = "Herrería"
180     SkillsNames(e_Skill.Carpinteria) = "Carpintería"
182     SkillsNames(e_Skill.Alquimia) = "Alquimia"
184     SkillsNames(e_Skill.Sastreria) = "Sastrería"
186     SkillsNames(e_Skill.Domar) = "Domar"
   
188     ListaAtributos(e_Atributos.Fuerza) = "Fuerza"
190     ListaAtributos(e_Atributos.Agilidad) = "Agilidad"
192     ListaAtributos(e_Atributos.Inteligencia) = "Inteligencia"
194     ListaAtributos(e_Atributos.Constitucion) = "Constitución"
196     ListaAtributos(e_Atributos.Carisma) = "Carisma"
    
200     IniPath = App.Path & "\"
    
        'Bordes del mapa
202     MinXBorder = XMinMapSize + (XWindow \ 2)
204     MaxXBorder = XMaxMapSize - (XWindow \ 2)
206     MinYBorder = YMinMapSize + (YWindow \ 2)
208     MaxYBorder = YMaxMapSize - (YWindow \ 2)
    
        
        Exit Sub

InicializarConstantes_Err:
210     Call TraceError(Err.Number, Err.Description, "General.InicializarConstantes", Erl)

        
End Sub

Sub Main()

        On Error GoTo Handler

        ' Me fijo si ya hay un proceso llamado server.exe abierto
100     If GetProcess(App.EXEName & ".exe") > 1 Then
            ' Si lo hay, pregunto si lo queremos cerrar.
102         If MsgBox("Se ha encontrado mas de 1 instancia abierta de esta aplicación, ¿Desea continuar?", vbYesNo) = vbNo Then
104             End
            End If
        End If

        Dim f As Date
    
110     Call ChDir(App.Path)
112     Call ChDrive(App.Path)

114     Call InicializarConstantes
    
116     frmCargando.Show
        LastCountUsersOnline = -1
    
118     frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    
120     frmCargando.Label1(2).Caption = "Iniciando Arrays..."

        cuentaregresivaOrcos = 300
        
122     Call LoadGuildsDB
    
126     Call loadAdministrativeUsers

        '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
128     frmCargando.Label1(2).Caption = "Cargando Server.ini"
    
130     MaxUsers = 0
132     Call LoadSini
137     Call LoadMD5
135     Call LoadPacketRatePolicy
133     Call LoadPrivateKey
138     Call LoadConfiguraciones
140     Call LoadIntervalos
142     Call CargarForbidenWords
144     Call CargaApuestas
146     Call CargarSpawnList
148     Call LoadMotd
150     Call CargarListaNegraUsuarios
        Call initBase64Chars
    
152     frmCargando.Label1(2).Caption = "Conectando base de datos y limpiando usuarios logueados"
    
        ' ************************* Base de Datos ********************
        'Conecto base de datos
154     Call Database_Connect
        
        Call Database_Connect_Async
    
        ' Construimos las querys grandes
156     Call Contruir_Querys

        ' ******************* FIN - Base de Datos ********************

        '*************************************************
164     frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
166     Call CargaNpcsDat
        '*************************************************
    
168     frmCargando.Label1(2).Caption = "Cargando Obj.Dat"

170     Call LoadOBJData
        
172     frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
174     Call CargarHechizos
        
176     frmCargando.Label1(2).Caption = "Cargando Objetos de Herrería"
178     Call LoadArmasHerreria
180     Call LoadArmadurasHerreria
    
182     frmCargando.Label1(2).Caption = "Cargando Objetos de Carpintería"
184     Call LoadObjCarpintero
    
186     frmCargando.Label1(2).Caption = "Cargando Objetos de Alquimista"
188     Call LoadObjAlquimista
    
190     frmCargando.Label1(2).Caption = "Cargando Objetos de Sastre"
192     Call LoadObjSastre
    
194     frmCargando.Label1(2).Caption = "Cargando Pesca"
196     Call LoadPesca
    
198     frmCargando.Label1(2).Caption = "Cargando Recursos Especiales"
200     Call LoadRecursosEspeciales

202     frmCargando.Label1(2).Caption = "Cargando Rangos de Faccion"
204     Call LoadRangosFaccion

206     frmCargando.Label1(2).Caption = "Cargando Recompensas de Faccion"
208     Call LoadRecompensasFaccion
    
210     frmCargando.Label1(2).Caption = "Cargando Balance.dat"
212     Call LoadBalance    '4/01/08 Pablo ToxicWaste
    
214     frmCargando.Label1(2).Caption = "Cargando Ciudades.dat"
216     Call CargarCiudades
    
218     If BootDelBackUp Then
220         frmCargando.Label1(2).Caption = "Cargando WorldBackup"
222         Call CargarBackUp
        Else
224         frmCargando.Label1(2).Caption = "Cargando Mapas"
226         Call LoadMapData
        End If
        
        frmCargando.Label1(2).Caption = "Cargando donadores"
        Call CargarDonadores
        
228     Call InitPathFinding

230     frmCargando.Label1(2).Caption = "Cargando informacion de eventos"
232     Call CargarInfoRetos
234     Call CargarInfoEventos
    
        ' Pretorianos
236     frmCargando.Label1(2).Caption = "Cargando Pretorianos.dat"
        'Call LoadPretorianData

242     frmCargando.Label1(2).Caption = "Cargando Baneos Temporales"
244     Call LoadBans
    
246     frmCargando.Label1(2).Caption = "Cargando Quests"
252     Call LoadQuests

254     EstadoGlobal = False
    
        Call ResetLastLogout

        'Comentado porque hay worldsave en ese mapa!
        'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
        '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
        Dim LoopC As Integer
    
        'Resetea las conexiones de los usuarios
258     For LoopC = 1 To MaxUsers
262         UserList(LoopC).ConnIDValida = False
268     Next LoopC
    
        '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
270     With frmMain
272         .Minuto.Enabled = True
274         .TimerGuardarUsuarios.Enabled = True
276         .TimerGuardarUsuarios.Interval = IntervaloTimerGuardarUsuarios
278         .tPiqueteC.Enabled = True
280         .GameTimer.Enabled = True
282         .Segundo.Enabled = True
284         .KillLog.Enabled = True
286         .TIMER_AI.Enabled = True
            .T_UsersOnline.Enabled = True
            .t_Extraer.Enabled = True
            .t_Extraer.Interval = IntervaloTrabajarExtraer
        End With
    
290     Subasta.SubastaHabilitada = True
292     Subasta.HaySubastaActiva = False
294     Call ResetMeteo
    
        ' ----------------------------------------------------
        '           Configuracion de los sockets
        ' ----------------------------------------------------
296     Call InitializePacketList

        #If AntiExternos = 1 Then
300         Call Security.Initialize
        #End If
        
302     Call modNetwork.Listen(MaxUsers, "0.0.0.0", CStr(Puerto))

312     If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
        ' ----------------------------------------------------
        '           Configuracion de los sockets
        ' ----------------------------------------------------
    
314     Call GetHoraActual
    
316     HoraMundo = GetTickCount() - DuracionDia \ 2

318     frmCargando.Visible = False
320     Unload frmCargando

        'Log
        Dim n As Integer
322     n = FreeFile
324     Open App.Path & "\logs\Main.log" For Append Shared As #n
326     Print #n, Date & " " & Time & " server iniciado " & App.Major & "." & App.Minor & "." & App.Revision
328     Close #n
    
        'Ocultar
330     Call frmMain.InitMain(HideMe)
    
332     tInicioServer = GetTickCount()
        
        While (True)
        
            Call modNetwork.Tick(GetElapsed())
            
            DoEvents
            
            ' Unlock main loop for maximum throughput but it can hog weak CPUs.
            #If UNLOCK_CPU = 0 Then
                Call Sleep(1)
            #End If
            
        Wend
        
        Exit Sub
        
Handler:
334     Call RegistrarError(Err.Number, Err.Description, "General.Main", Erl)


End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
        '*****************************************************************
        'Se fija si existe el archivo
        '*****************************************************************
        
        On Error GoTo FileExist_Err
        
100     FileExist = LenB(dir$(File, FileType)) <> 0

        
        Exit Function

FileExist_Err:
102     Call TraceError(Err.Number, Err.Description, "General.FileExist", Erl)

        
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

        Dim currentPos As Long

        Dim delimiter  As String * 1
    
100     delimiter = Chr$(SepASCII)
    
102     For i = 1 To Pos
104         LastPos = currentPos
106         currentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
108     Next i
    
110     If currentPos = 0 Then
112         ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
        Else
114         ReadField = mid$(Text, LastPos + 1, currentPos - LastPos - 1)

        End If

        
        Exit Function

ReadField_Err:
116     Call TraceError(Err.Number, Err.Description, "General.ReadField", Erl)

        
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
        
        On Error GoTo MapaValido_Err
        
100     MapaValido = Map >= 1 And Map <= NumMaps

        
        Exit Function

MapaValido_Err:
102     Call TraceError(Err.Number, Err.Description, "General.MapaValido", Erl)

        
End Function

Sub MostrarNumUsers()

        On Error GoTo MostrarNumUsers_Err
        
        If (LastCountUsersOnline <> NumUsers) Then
100         Call SendData(SendTarget.ToAll, 0, PrepareMessageOnlineUser(NumUsers))

102         frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers
        
            LastCountUsersOnline = NumUsers
            
104         Call SetUsersLoggedDatabase(NumUsers)
        End If
        
        If (LastRecordUsuarios <> RecordUsuarios) Then
            LastRecordUsuarios = RecordUsuarios
            
            Call SetRecordUsersDatabase(LastRecordUsuarios)
        End If

        Exit Sub

MostrarNumUsers_Err:
106     Call TraceError(Err.Number, Err.Description, "General.MostrarNumUsers", Erl)

        
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

Public Sub LogError(Desc As String)

100     Dim nfile As Integer: nfile = FreeFile ' obtenemos un canal
    
102     Open App.Path & "\logs\errores.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & Desc
106     Close #nfile

        Exit Sub

End Sub

Public Sub LogPerformance(Desc As String)

100     Dim nfile As Integer: nfile = FreeFile ' obtenemos un canal
    
102     Open App.Path & "\logs\Performance.log" For Append Shared As #nfile
104         Print #nfile, Date & " " & Time & " " & Desc
106     Close #nfile

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

Public Sub LogClanes(ByVal str As String)
        
        On Error GoTo LogClanes_Err
        

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        
        Exit Sub

LogClanes_Err:
108     Call TraceError(Err.Number, Err.Description, "General.LogClanes", Erl)

        
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
108     Call TraceError(Err.Number, Err.Description, "General.LogIP", Erl)

        
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
108     Call TraceError(Err.Number, Err.Description, "General.LogDesarrollo", Erl)

        
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
108     Print #nfile, "Item: " & ObjData(ObjIndex).Name & " (" & ObjIndex & ") Cantidad: " & Cantidad & vbNewLine
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

Public Sub LogEdicionPaquete(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal

102     Open App.Path & "\logs\EdicionPaquete.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & texto
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogMacroServidor(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal

102     Open App.Path & "\logs\MacroServidor.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & texto
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogMacroCliente(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal

102     Open App.Path & "\logs\MacroCliente.log" For Append Shared As #nfile
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


Sub Restart()
        
        On Error GoTo Restart_Err
        
        'Se asegura de que los sockets estan cerrados e ignora cualquier err

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

        Dim LoopC As Long

102     Call modNetwork.Disconnect

104     For LoopC = 1 To MaxUsers
106         Call CloseSocket(LoopC)
        Next

        'Initialize statistics!!
        'Call Statistics.Initialize

116     ReDim UserList(1 To MaxUsers) As t_User

118     For LoopC = 1 To MaxUsers
122         UserList(LoopC).ConnIDValida = False
128     Next LoopC

130     LastUser = 0
132     NumUsers = 0

134     Call FreeNPCs
136     Call FreeCharIndexes

138     Call LoadSini
        Call LoadMD5
133     Call LoadPrivateKey
140     Call LoadIntervalos
142     Call LoadOBJData
144     Call LoadPesca
146     Call LoadRecursosEspeciales

148     Call LoadMapData

150     Call CargarHechizos
        
        Call modNetwork.Listen(MaxUsers, "0.0.0.0", CStr(Puerto))
        
152     If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

        'Log it
        Dim n As Integer

154     n = FreeFile
156     Open App.Path & "\logs\Main.log" For Append Shared As #n
158     Print #n, Date & " " & Time & " servidor reiniciado."
160     Close #n

        'Ocultar
162     Call frmMain.InitMain(HideMe)
    
        Exit Sub

Restart_Err:
164     Call TraceError(Err.Number, Err.Description, "General.Restart", Erl)

        
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
106     Call TraceError(Err.Number, Err.Description, "General.Intemperie", Erl)

        
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
112     Call TraceError(Err.Number, Err.Description, "General.TiempoInvocacion", Erl)

        
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoFrio_Err
        
100     If Not Intemperie(UserIndex) Then Exit Sub
        
102     With UserList(UserIndex)
            
104         If .Invent.ArmourEqpObjIndex > 0 Then
                ' WyroX: Ropa invernal
106             If ObjData(.Invent.ArmourEqpObjIndex).Invernal Then Exit Sub
            End If
            
108         If .Counters.Frio < IntervaloFrio Then
110             .Counters.Frio = .Counters.Frio + 1
            Else

112             If MapInfo(.Pos.Map).terrain = Nieve Then
114                 Call WriteConsoleMsg(UserIndex, "¡Estás muriendo de frío, abrígate o morirás!", e_FontTypeNames.FONTTYPE_INFO)

                    ' WyroX: Sin ropa perdés vida más rápido que con una ropa no-invernal
                    Dim MinDaño As Integer, MaxDaño As Integer
116                 If .flags.Desnudo = 0 Then
118                     MinDaño = 17
120                     MaxDaño = 23
                    Else
122                     MinDaño = 27
124                     MaxDaño = 33
                    End If

                    ' WyroX: Agrego aleatoriedad
                    Dim Daño As Integer
126                 Daño = Porcentaje(.Stats.MaxHp, RandomNumber(MinDaño, MaxDaño))

128                 .Stats.MinHp = .Stats.MinHp - Daño
            
130                 If .Stats.MinHp < 1 Then

132                     Call WriteConsoleMsg(UserIndex, "¡Has muerto de frío!", e_FontTypeNames.FONTTYPE_INFO)

134                     Call UserDie(UserIndex)

                    Else
136                     Call WriteUpdateHP(UserIndex)
                    End If
                End If
        
138             .Counters.Frio = 0

            End If
        
        End With
        
        Exit Sub

EfectoFrio_Err:
140     Call TraceError(Err.Number, Err.Description, "General.EfectoFrio", Erl)


        
End Sub

Public Sub EfectoStamina(ByVal UserIndex As Integer)

    Dim HambreOSed As Boolean
    Dim bEnviarStats_HP As Boolean
    Dim bEnviarStats_STA As Boolean
    
100 With UserList(UserIndex)
102     HambreOSed = .Stats.MinHam = 0 Or .Stats.MinAGU = 0
    
104     If Not HambreOSed Then 'Si no tiene hambre ni sed
106         If .Stats.MinHp < .Stats.MaxHp Then
108             Call Sanar(UserIndex, bEnviarStats_HP, IIf(.flags.Descansar, SanaIntervaloDescansar, SanaIntervaloSinDescansar))
            End If
        End If
                                
110     If .flags.Desnudo = 0 And Not HambreOSed Then
112         If Not Lloviendo Or Not Intemperie(UserIndex) Then
114             Call RecStamina(UserIndex, bEnviarStats_STA, IIf(.flags.Descansar, StaminaIntervaloDescansar, StaminaIntervaloSinDescansar))
            End If
        Else
116         If Lloviendo And Intemperie(UserIndex) Then
118             Call PierdeEnergia(UserIndex, bEnviarStats_STA, IntervaloPerderStamina * 0.5)
            Else
120             Call PierdeEnergia(UserIndex, bEnviarStats_STA, IIf(.flags.Descansar, IntervaloPerderStamina * 2, IntervaloPerderStamina))
            End If
        End If
        
122     If .flags.Descansar Then
            'termina de descansar automaticamente
124         If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
126             Call WriteRestOK(UserIndex)
128             Call WriteConsoleMsg(UserIndex, "Has terminado de descansar.", e_FontTypeNames.FONTTYPE_INFO)
130             .flags.Descansar = False
            End If
        
        End If
        
132     If bEnviarStats_STA Then
134         Call WriteUpdateSta(UserIndex)
        End If
        
136     If bEnviarStats_HP Then
138         Call WriteUpdateHP(UserIndex)
        End If
    End With
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
108                 Call WriteConsoleMsg(UserIndex, "¡Quítate de la lava, te estás quemando!", e_FontTypeNames.FONTTYPE_INFO)
110                 .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHp, 5)
            
112                 If .Stats.MinHp < 1 Then
114                     Call WriteConsoleMsg(UserIndex, "¡Has muerto quemado!", e_FontTypeNames.FONTTYPE_INFO)
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
122     Call TraceError(Err.Number, Err.Description, "General.EfectoLava", Erl)


        
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
    
        
        Dim Barco As t_ObjData
    
100     With UserList(UserIndex)
102         If .Counters.Mimetismo < IntervaloInvisible Then
104             .Counters.Mimetismo = .Counters.Mimetismo + 1

            Else
                'restore old char
106             Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", e_FontTypeNames.FONTTYPE_INFO)
            
108             If .flags.Navegando Then
110                 Call EquiparBarco(UserIndex)
                Else
112                 .Char.Body = .CharMimetizado.Body
114                 .Char.Head = .CharMimetizado.Head
116                 .Char.CascoAnim = .CharMimetizado.CascoAnim
118                 .Char.ShieldAnim = .CharMimetizado.ShieldAnim
120                 .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                End If
                
122             .Counters.Mimetismo = 0
124             .flags.Mimetizado = e_EstadoMimetismo.Desactivado
            
126             With .Char
128                 Call ChangeUserChar(UserIndex, .Body, .Head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
130                 Call RefreshCharStatus(UserIndex)
                End With
                
            End If
            
        End With
        
        Exit Sub

EfectoMimetismo_Err:
132     Call TraceError(Err.Number, Err.Description, "General.EfectoMimetismo", Erl)

        
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoInvisibilidad_Err
        

100     If UserList(UserIndex).Counters.Invisibilidad > 0 Then
102         UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad - 1
        Else
104         UserList(UserIndex).Counters.Invisibilidad = 0
106         UserList(UserIndex).flags.invisible = 0

108         If UserList(UserIndex).flags.Oculto = 0 Then
                ' Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", e_FontTypeNames.FONTTYPE_INFO)
110             Call WriteLocaleMsg(UserIndex, "307", e_FontTypeNames.FONTTYPE_INFO)
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
114             Call WriteContadores(UserIndex)

            End If

        End If

        
        Exit Sub

EfectoInvisibilidad_Err:
116     Call TraceError(Err.Number, Err.Description, "General.EfectoInvisibilidad", Erl)

        
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
106     Call TraceError(Err.Number, Err.Description, "General.EfectoParalisisNpc", Erl)

        
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
106     Call TraceError(Err.Number, Err.Description, "General.EfectoInmovilizadoNpc", Erl)

        
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
110     Call TraceError(Err.Number, Err.Description, "General.EfectoCeguera", Erl)

        
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
110     Call TraceError(Err.Number, Err.Description, "General.EfectoEstupidez", Erl)

        
End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoParalisisUser_Err
        
        With UserList(UserIndex)
100         If .Counters.Paralisis > 0 Then
102             .Counters.Paralisis = .Counters.Paralisis - 1
            Else
104             .flags.Paralizado = 0
    
                If .clase = e_Class.Warrior Or .clase = e_Class.Hunter Or .clase = e_Class.Thief Or .clase = e_Class.Pirat Then
                    .Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
                'UserList(UserIndex).Flags.AdministrativeParalisis = 0
106             Call WriteParalizeOK(UserIndex)
    
            End If
        End With

        
        Exit Sub

EfectoParalisisUser_Err:
108     Call TraceError(Err.Number, Err.Description, "General.EfectoParalisisUser", Erl)

        
End Sub

Public Sub EfectoVelocidadUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoVelocidadUser_Err
        

100     If UserList(UserIndex).Counters.velocidad > 0 Then
102         UserList(UserIndex).Counters.velocidad = UserList(UserIndex).Counters.velocidad - 1
        Else
104         UserList(UserIndex).flags.VelocidadHechizada = 0
106         Call ActualizarVelocidadDeUsuario(UserIndex)
        End If

        Exit Sub

EfectoVelocidadUser_Err:
108     Call TraceError(Err.Number, Err.Description, "General.EfectoVelocidadUser", Erl)

        
End Sub

Public Sub EfectoMaldicionUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoMaldicionUser_Err
        

100     If UserList(UserIndex).Counters.Maldicion > 0 Then
102         UserList(UserIndex).Counters.Maldicion = UserList(UserIndex).Counters.Maldicion - 1
    
        Else
104         UserList(UserIndex).flags.Maldicion = 0
106         Call WriteConsoleMsg(UserIndex, "¡La magia perdió su efecto! Ya puedes atacar.", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
        End If

        
        Exit Sub

EfectoMaldicionUser_Err:
108     Call TraceError(Err.Number, Err.Description, "General.EfectoMaldicionUser", Erl)

        
End Sub

Public Sub EfectoInmoUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoInmoUser_Err
        
        With UserList(UserIndex)
100         If .Counters.Inmovilizado > 0 Then
102             .Counters.Inmovilizado = .Counters.Inmovilizado - 1
            Else
104             .flags.Inmovilizado = 0

                If .clase = e_Class.Warrior Or .clase = e_Class.Hunter Or .clase = e_Class.Thief Or .clase = e_Class.Pirat Then
                    .Counters.TiempoDeInmunidadParalisisNoMagicas = 4
                End If
106             Call WriteInmovilizaOK(UserIndex)
    
            End If
        End With

        
        Exit Sub

EfectoInmoUser_Err:
108     Call TraceError(Err.Number, Err.Description, "General.EfectoInmoUser", Erl)

        
End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
            On Error GoTo RecStamina_Err

            Dim trigger As Byte
            Dim Suerte As Integer

100         With UserList(UserIndex)
102             trigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger

104             If trigger = 1 And trigger = 2 And trigger = 4 Then Exit Sub

106             If .Stats.MinSta < .Stats.MaxSta Then

108                 If .Counters.STACounter < Intervalo Then
110                     .Counters.STACounter = .Counters.STACounter + 1
                        Exit Sub

                    End If

112                 .Counters.STACounter = 0

114                 If .Counters.Trabajando > 0 Then Exit Sub  'Trabajando no sube energía. (ToxicWaste)

116                 EnviarStats = True

118                 Select Case .Stats.UserSkills(e_Skill.Supervivencia)
                        Case 0 To 10
120                         Suerte = 5
122                     Case 11 To 20
124                         Suerte = 7
126                     Case 21 To 30
128                         Suerte = 9
130                     Case 31 To 40
132                         Suerte = 11
134                     Case 41 To 50
136                         Suerte = 13
138                     Case 51 To 60
140                         Suerte = 15
142                     Case 61 To 70
144                         Suerte = 17
146                     Case 71 To 80
148                         Suerte = 19
150                     Case 81 To 90
152                         Suerte = 21
154                     Case 91 To 99
156                         Suerte = 23
158                     Case 100
160                         Suerte = 25
                    End Select

162                 If .flags.RegeneracionSta = 1 Then Suerte = 45
                    
                    Dim NuevaStamina As Long
164                     NuevaStamina = .Stats.MinSta + RandomNumber(1, CInt(Porcentaje(.Stats.MaxSta, Suerte)))
                    
                    ' Jopi: Prevenimos overflow al acotar la stamina que se puede recuperar en cualquier caso.
                    ' Cuando te editabas la energia con el GM causaba este error.
166                 If NuevaStamina < 32000 Then
168                     .Stats.MinSta = NuevaStamina
                    Else
170                     .Stats.MinSta = 32000
                    End If

172                 If .Stats.MinSta > .Stats.MaxSta Then
174                     .Stats.MinSta = .Stats.MaxSta
                    End If

                End If
            End With

            Exit Sub

RecStamina_Err:
176         Call TraceError(Err.Number, Err.Description, "General.RecStamina", Erl)


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
    
112                 Cantidad = RandomNumber(1, Porcentaje(.Stats.MaxSta, (MAXSKILLPOINTS * 1.5 - .Stats.UserSkills(e_Skill.Supervivencia)) * 0.25))
114                 .Stats.MinSta = .Stats.MinSta - Cantidad
    
116                 If .Stats.MinSta < 0 Then
118                     .Stats.MinSta = 0
                    End If
    
                End If
    
            End If

        End With
        
        Exit Sub

RecStamina_Err:
120     Call TraceError(Err.Number, Err.Description, "General.PierdeEnergia", Erl)

        
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)

        On Error GoTo EfectoVeneno_Err

        Dim damage As Long

100     If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
102         UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
        Else
104         Call CancelExit(UserIndex)

106         With UserList(UserIndex)
              'Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas morirás.", e_FontTypeNames.FONTTYPE_VENENO)
108           Call WriteLocaleMsg(UserIndex, "47", e_FontTypeNames.FONTTYPE_VENENO)
110           Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, e_ParticulasIndex.Envenena, 30, False))
112           .Counters.Veneno = 0

              ' El veneno saca un porcentaje de vida random.
114           damage = RandomNumber(3, 5)
116           damage = (1 + damage * .Stats.MaxHp \ 100) ' Redondea para arriba
118           .Stats.MinHp = UserList(UserIndex).Stats.MinHp - damage

120           If .ChatCombate = 1 Then
                  ' "El veneno te ha causado ¬1 puntos de daño."
122               Call WriteLocaleMsg(UserIndex, "390", e_FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(damage))
              End If

124           If UserList(UserIndex).Stats.MinHp < 1 Then
126               Call UserDie(UserIndex)
              Else
128               Call WriteUpdateHP(UserIndex)
              End If
            End With

        End If

        Exit Sub

EfectoVeneno_Err:
130     Call TraceError(Err.Number, Err.Description, "General.EfectoVeneno", Erl)


End Sub

Public Sub EfectoAhogo(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoAhogo_Err
        
100     If Not UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then Exit Sub
        
        Dim n As Integer

102     If RequiereOxigeno(UserList(UserIndex).Pos.Map) Then
104         If UserList(UserIndex).Counters.Ahogo < 70 Then
106             UserList(UserIndex).Counters.Ahogo = UserList(UserIndex).Counters.Ahogo + 1
            Else
108             Call WriteConsoleMsg(UserIndex, "Te estás ahogando, si no consigues oxígeno morirás.", e_FontTypeNames.FONTTYPE_EJECUCION)
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 205, 30, False))
110             UserList(UserIndex).Counters.Ahogo = 0
112             n = RandomNumber(150, 200)
114             UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - n

116             If UserList(UserIndex).Stats.MinHp < 1 Then
118                 Call UserDie(UserIndex)
                Else
120                 Call WriteUpdateHP(UserIndex)
                End If

            End If

        Else
122         UserList(UserIndex).flags.Ahogandose = 0

        End If

        
        Exit Sub

EfectoAhogo_Err:
124     Call TraceError(Err.Number, Err.Description, "General.EfectoAhogo", Erl)

        
End Sub

' El incineramiento tiene una logica particular, que es hacer daño sostenido en el tiempo.
Public Sub EfectoIncineramiento(ByVal UserIndex As Integer)
            On Error GoTo EfectoIncineramiento_Err

            Dim damage As Integer

100         With UserList(UserIndex)

                ' 5 Mini intervalitos, dentro del intervalo total de incineracion
102             If .Counters.Incineracion Mod (IntervaloIncineracion \ 5) = 0 Then
                    ' "Te estás incinerando, si no te curas morirás.
104                 Call WriteLocaleMsg(UserIndex, "392", e_FontTypeNames.FONTTYPE_FIGHT)
                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, e_ParticulasIndex.Incinerar, 30, False))
106                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 73, 0))

108                 damage = RandomNumber(35, 45)
110                 .Stats.MinHp = .Stats.MinHp - damage

112                 If .ChatCombate = 1 Then
                        ' "El fuego te ha causado ¬1 puntos de daño."
114                     Call WriteLocaleMsg(UserIndex, "391", e_FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(damage))
                    End If

116                 If UserList(UserIndex).Stats.MinHp < 1 Then
118                     Call UserDie(UserIndex)
                    Else
120                     Call WriteUpdateHP(UserIndex)
                    End If
                End If

122             .Counters.Incineracion = .Counters.Incineracion + 1

124             If .Counters.Incineracion > IntervaloIncineracion Then
                    ' Se termino la incineracion
126                 .flags.Incinerado = 0
128                 .Counters.Incineracion = 0
                    Exit Sub

                End If
            End With

            Exit Sub

EfectoIncineramiento_Err:
130         Call TraceError(Err.Number, Err.Description, "General.EfectoIncineramiento", Erl)


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
116     Call TraceError(Err.Number, Err.Description, "General.DuracionPociones", Erl)

        
End Sub

Public Function HambreYSed(ByVal UserIndex As Integer) As Boolean
         
        On Error GoTo HambreYSed_Err
        

100     If (UserList(UserIndex).flags.Privilegios And e_PlayerType.user) = 0 Then Exit Function

        'Sed
102     If UserList(UserIndex).Stats.MinAGU > 0 Then
104         If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
106             UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
            Else
108             UserList(UserIndex).Counters.AGUACounter = 0
110             UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
        
112             If UserList(UserIndex).Stats.MinAGU <= 0 Then
114                 UserList(UserIndex).Stats.MinAGU = 0

                End If
        
118             HambreYSed = True

            End If

        End If

        'hambre
120     If UserList(UserIndex).Stats.MinHam > 0 Then
122         If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
124             UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
            Else
126             UserList(UserIndex).Counters.COMCounter = 0
128             UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10

130             If UserList(UserIndex).Stats.MinHam <= 0 Then
132                 UserList(UserIndex).Stats.MinHam = 0

                End If

136             HambreYSed = True

            End If

        End If

        
        Exit Function

HambreYSed_Err:
138     Call TraceError(Err.Number, Err.Description, "General.HambreYSed", Erl)

        
End Function

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
        
        On Error GoTo Sanar_Err
        
        ' Desnudo no regenera vida
100     If UserList(UserIndex).flags.Desnudo = 1 Then Exit Sub
        
102     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

        Dim mashit As Integer

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
104         If UserList(UserIndex).flags.RegeneracionHP = 1 Then
106             Intervalo = 400

            End If
    
108         If UserList(UserIndex).Counters.HPCounter < Intervalo Then
110             UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
            Else
112             mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
        
114             UserList(UserIndex).Counters.HPCounter = 0
116             UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp + mashit

118             If UserList(UserIndex).Stats.MinHp > UserList(UserIndex).Stats.MaxHp Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
120             Call WriteConsoleMsg(UserIndex, "Has sanado.", e_FontTypeNames.FONTTYPE_INFO)
122             EnviarStats = True

            End If
        Exit Sub

Sanar_Err:
124     Call TraceError(Err.Number, Err.Description, "General.Sanar", Erl)

        
End Sub

Public Sub CargaNpcsDat(Optional ByVal ActualizarNPCsExistentes As Boolean = False)
        
            On Error GoTo CargaNpcsDat_Err
        
            ' Leemos el NPCs.dat y lo almacenamos en la memoria.
100         Set LeerNPCs = New clsIniManager
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
118         Call TraceError(Err.Number, Err.Description, "General.CargaNpcsDat", Erl)

        
End Sub

Sub PasarSegundo()

        On Error GoTo ErrHandler

        Dim i    As Long

        Dim h    As Byte

        Dim Mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
        
        
        
70      If cuentaregresivaOrcos > 0 Then
74          cuentaregresivaOrcos = cuentaregresivaOrcos - 1
        Else
76            cuentaregresivaOrcos = 300
78            Call TimerQuestOrco
        End If
        
        
100     If CuentaRegresivaTimer > 0 Then
102         If CuentaRegresivaTimer > 1 Then
104             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(CuentaRegresivaTimer - 1 & " segundos...!", e_FontTypeNames.FONTTYPE_GUILD))
            Else
106             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ya!!!", e_FontTypeNames.FONTTYPE_FIGHT))

            End If

108         CuentaRegresivaTimer = CuentaRegresivaTimer - 1

        End If
    
110     For i = 1 To LastUser

112         With UserList(i)

114             If .flags.UserLogged Then
116                 If .flags.Silenciado = 1 Then
118                     .flags.SegundosPasados = .flags.SegundosPasados + 1
        
120                     If .flags.SegundosPasados = 60 Then
122                         .flags.MinutosRestantes = .flags.MinutosRestantes - 1
124                         .flags.SegundosPasados = 0
        
                        End If
                    
126                     If .flags.MinutosRestantes = 0 Then
128                         .flags.SegundosPasados = 0
130                         .flags.Silenciado = 0
132                         .flags.MinutosRestantes = 0
134                         Call WriteConsoleMsg(i, "Has sido liberado del silencio.", e_FontTypeNames.FONTTYPE_SERVER)
        
                        End If
        
                    End If
                    
                     If .Counters.TimerCentinela >= 0 Then
                        .Counters.TimerCentinela = .Counters.TimerCentinela - 1

                        If .Counters.TimerCentinela = 0 Then
                            .flags.CentinelaOK = True
                            .Counters.TimerCentinela = -1
                        End If
                    End If

                    
136                 If .flags.Muerto = 0 Then
138                     Call DuracionPociones(i)
140                     Call EfectoOxigeno(i)
142                     If .flags.invisible = 1 Then Call EfectoInvisibilidad(i)
144                     If .flags.Paralizado = 1 Then Call EfectoParalisisUser(i)
146                     If .flags.Inmovilizado = 1 Then Call EfectoInmoUser(i)
148                     If .flags.Ceguera = 1 Then Call EfectoCeguera(i)
150                     If .flags.Estupidez = 1 Then Call EfectoEstupidez(i)
152                     If .flags.Maldicion = 1 Then Call EfectoMaldicionUser(i)
154                     If .flags.VelocidadHechizada > 0 Then Call EfectoVelocidadUser(i)
    
156                     If HambreYSed(i) Then
158                         Call WriteUpdateHungerAndThirst(i)
                        End If
                    
                    Else
160                     If .flags.Traveling <> 0 Then Call TravelingEffect(i)
                    End If
        
162                 If .Counters.TimerBarra > 0 Then
164                     .Counters.TimerBarra = .Counters.TimerBarra - 1
                        
166                     If .Counters.TimerBarra = 0 Then
        
168                         Select Case .Accion.TipoAccion
                                Case e_AccionBarra.Hogar
170                                 Call HomeArrival(i)
                            End Select
                            
182                         .Accion.Particula = 0
184                         .Accion.TipoAccion = e_AccionBarra.CancelarAccion
186                         .Accion.HechizoPendiente = 0
188                         .Accion.RunaObj = 0
190                         .Accion.ObjSlot = 0
192                         .Accion.AccionPendiente = False
                            
                        End If
                    End If
        
194                 If .flags.UltimoMensaje > 0 Then
196                     .Counters.RepetirMensaje = .Counters.RepetirMensaje + 1
198                     If .Counters.RepetirMensaje >= 3 Then
200                         .flags.UltimoMensaje = 0
202                         .Counters.RepetirMensaje = 0
                        End If
                    End If
                    
204                 If .Counters.CuentaRegresiva >= 0 Then
206                     If .Counters.CuentaRegresiva > 0 Then
208                         Call WriteConsoleMsg(i, ">>>  " & .Counters.CuentaRegresiva & "  <<<", e_FontTypeNames.FONTTYPE_New_Gris)
                        Else
210                         Call WriteConsoleMsg(i, ">>> YA! <<<", e_FontTypeNames.FONTTYPE_FIGHT)
212                         Call WriteStopped(i, False)
                        End If
                        
214                     .Counters.CuentaRegresiva = .Counters.CuentaRegresiva - 1
                    End If
    
216                 If .flags.Portal > 1 Then
218                     .flags.Portal = .flags.Portal - 1
                
220                     If .flags.Portal = 1 Then
222                         Mapa = .flags.PortalM
224                         X = .flags.PortalX
226                         Y = .flags.PortalY
228                         Call SendData(SendTarget.toMap, .flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, e_ParticulasIndex.TpVerde, 0))
230                         Call SendData(SendTarget.toMap, .flags.PortalM, PrepareMessageLightFXToFloor(X, Y, 0, 105))
        
232                         If MapData(Mapa, X, Y).TileExit.Map > 0 Then
234                             MapData(Mapa, X, Y).TileExit.Map = 0
236                             MapData(Mapa, X, Y).TileExit.X = 0
238                             MapData(Mapa, X, Y).TileExit.Y = 0
        
                            End If
        
240                         MapData(Mapa, X, Y).Particula = 0
242                         MapData(Mapa, X, Y).TimeParticula = 0
244                         MapData(Mapa, X, Y).Particula = 0
246                         MapData(Mapa, X, Y).TimeParticula = 0
248                         .flags.Portal = 0
250                         .flags.PortalM = 0
252                         .flags.PortalY = 0
254                         .flags.PortalX = 0
256                         .flags.PortalMDestino = 0
258                         .flags.PortalYDestino = 0
260                         .flags.PortalXDestino = 0
        
                        End If
        
                    End If
                
262                 If .Counters.EnCombate > 0 Then
264                     .Counters.EnCombate = .Counters.EnCombate - 1
                    End If
                
                    If .Counters.TiempoDeInmunidadParalisisNoMagicas > 0 Then
                        .Counters.TiempoDeInmunidadParalisisNoMagicas = .Counters.TiempoDeInmunidadParalisisNoMagicas - 1
                    End If
                
266                 If .Counters.TiempoDeInmunidad > 0 Then
268                     .Counters.TiempoDeInmunidad = .Counters.TiempoDeInmunidad - 1
270                     If .Counters.TiempoDeInmunidad = 0 Then
272                         .flags.Inmunidad = 0
                        End If
                    End If
                
274                 If .flags.Subastando Then
276                     .Counters.TiempoParaSubastar = .Counters.TiempoParaSubastar - 1
        
278                     If .Counters.TiempoParaSubastar = 0 Then
280                         Call CancelarSubasta
                        End If
                    End If
        
                    'Cerrar usuario
282                 If .Counters.Saliendo Then
                        '  If .flags.Muerto = 1 Then .Counters.Salir = 0
284                     .Counters.Salir = .Counters.Salir - 1
                        ' Call WriteConsoleMsg(i, "Se saldrá del juego en " & .Counters.Salir & " segundos...", e_FontTypeNames.FONTTYPE_INFO)
286                     Call WriteLocaleMsg(i, "203", e_FontTypeNames.FONTTYPE_INFO, .Counters.Salir)
        
288                     If .Counters.Salir <= 0 Then
290                         Call WriteConsoleMsg(i, "Gracias por jugar Argentum 20.", e_FontTypeNames.FONTTYPE_INFO)
292                         Call WriteDisconnect(i)
                            
294                         Call CloseSocket(i)
        
                        End If
        
                    End If

                End If ' If UserLogged

            End With
302     Next i

        ' **********************************
        ' **********  Invasiones  **********
        ' **********************************
304     For i = 1 To UBound(Invasiones)
306         With Invasiones(i)

                ' Si la invasión está activa
308             If .Activa Then
310                 .TimerSpawn = .TimerSpawn + 1

                    ' Comprobamos si hay que spawnear NPCs
312                 If .TimerSpawn >= .IntervaloSpawn Then
314                     Call InvasionSpawnNPC(i)
316                     .TimerSpawn = 0
                    End If
                    
                    ' ------------------------------------
                    
318                 .TimerMostrarInfo = .TimerMostrarInfo + 1
                    
                    ' Comprobamos si hay que mostrar la info
320                 If .TimerMostrarInfo >= 5 Then
322                     Call EnviarInfoInvasion(i)
324                     .TimerMostrarInfo = 0
                    End If
                End If
            
            End With
        Next
        ' **********************************

        Exit Sub

ErrHandler:
326     Call TraceError(Err.Number, Err.Description, "General.PasarSegundo", Erl)

End Sub

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
108     Call TraceError(Err.Number, Err.Description, "General.ReiniciarServidor", Erl)

        
End Sub

Sub ForzarActualizar()
    
        On Error Resume Next
    
        Dim i As Long

100     For i = 1 To LastUser

102         If UserList(i).ConnIDValida Then
        
104             Call WriteForceUpdate(i)
    
            End If
    
106     Next i
    
End Sub
 
Sub GuardarUsuarios()
        
        On Error GoTo GuardarUsuarios_Err
        
100     haciendoBK = True
    
102     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
104     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Grabando Personajes", e_FontTypeNames.FONTTYPE_SERVER))
    
        Dim i As Long
        
106     For i = 1 To LastUser

108         If UserList(i).flags.UserLogged Then
110             Call modNetwork.Poll
            End If

112     Next i

114     For i = 1 To LastUser

116         If UserList(i).flags.UserLogged Then

118              Call SaveUser(i)

            End If

120     Next i
    
122     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Personajes Grabados", e_FontTypeNames.FONTTYPE_SERVER))
124     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

126     haciendoBK = False

        
        Exit Sub

GuardarUsuarios_Err:
128     Call TraceError(Err.Number, Err.Description, "General.GuardarUsuarios", Erl)

        
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
106     Call TraceError(Err.Number, Err.Description, "General.FreeNPCs", Erl)

        
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
102     Call TraceError(Err.Number, Err.Description, "General.FreeCharIndexes", Erl)

        
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
116     Call TraceError(Err.Number, Err.Description, "General.RandomString", Erl)

        
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
102     Call TraceError(Err.Number, Err.Description, "General.CMSValidateChar_", Erl)

        
End Function

Public Function Tilde(ByRef Data As String) As String
    
        On Error GoTo Tilde_Err
    

100     Tilde = UCase$(Data)
 
102     Tilde = Replace$(Tilde, "Á", "A")
104     Tilde = Replace$(Tilde, "É", "E")
106     Tilde = Replace$(Tilde, "Í", "I")
108     Tilde = Replace$(Tilde, "Ó", "O")
110     Tilde = Replace$(Tilde, "Ú", "U")
        
    
        Exit Function

Tilde_Err:
112     Call TraceError(Err.Number, Err.Description, "Mod_General.Tilde", Erl)

    
End Function

Public Sub CerrarServidor()
        
        ' Force-dump logs
        Call frmServidor.cmdDumpLogs_Click
        
        'Save stats!!!
102     Call frmMain.QuitarIconoSystray
    
        ' Limpieza del socket del servidor.
104     Call modNetwork.Disconnect
    
        Dim LoopC As Long
106     For LoopC = 1 To MaxUsers
108         If UserList(LoopC).ConnIDValida Then
110             Call CloseSocket(LoopC)
            End If
        Next
    
112     If Database_Enabled Then Database_Close
 
        'Log
116     Dim n As Integer: n = FreeFile
118     Open App.Path & "\logs\Main.log" For Append Shared As #n
120     Print #n, Date & " " & Time & " server cerrado."
122     Close #n
    
124     End
   
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
106     Call TraceError(Err.Number, Err.Description, "General.max", Erl)

        
End Function

Function Min(ByVal a As Double, ByVal b As Double) As Double
        
        On Error GoTo min_Err
    
        

100     If a < b Then
102         Min = a
        Else
104         Min = b
        End If

        
        Exit Function

min_Err:
106     Call TraceError(Err.Number, Err.Description, "General.min", Erl)

        
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
118     Call TraceError(Err.Number, Err.Description, "ModLadder.PonerPuntos", Erl)

    
End Function

' Autor: WyroX
Function CalcularPromedioVida(ByVal UserIndex As Integer) As Double

100     With UserList(UserIndex)
102         If .Stats.ELV = 1 Then
                ' Siempre estamos promedio al lvl 1
104             CalcularPromedioVida = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5
            Else
106             CalcularPromedioVida = (.Stats.MaxHp - .Stats.UserAtributos(e_Atributos.Constitucion)) / (.Stats.ELV - 1)
            End If
        End With

End Function

' Adaptado desde https://stackoverflow.com/questions/29325069/how-to-generate-random-numbers-biased-towards-one-value-in-a-range/29325222#29325222
' By WyroX
Function RandomIntBiased(ByVal Min As Double, ByVal max As Double, ByVal Bias As Double, ByVal Influence As Double) As Double

        On Error GoTo handle

        Dim RandomRango As Double, Mix As Double
    
100     RandomRango = Rnd * (max - Min) + Min
102     Mix = Rnd * Influence
    
104     RandomIntBiased = RandomRango * (1 - Mix) + Bias * Mix
    
        Exit Function
    
handle:
106     Call TraceError(Err.Number, Err.Description, "General.RandomIntBiased")
108     RandomIntBiased = Bias

End Function

'Very efficient function for testing whether this code is running in the IDE or compiled
'https://www.vbforums.com/showthread.php?231468-VB-Detect-if-you-are-running-in-the-IDE&p=5413357&viewfull=1#post5413357
Public Function RunningInVB(Optional ByRef b As Boolean = True) As Boolean
100     If b Then Debug.Assert Not RunningInVB(RunningInVB) Else b = True
End Function

' WyroX: Mensaje a todo el mundo
Public Sub MensajeGlobal(texto As String, Fuente As e_FontTypeNames)
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(texto, Fuente))
End Sub

' WyroX: Devuelve si X e Y están dentro del Rectangle
Public Function InsideRectangle(R As t_Rectangle, ByVal X As Integer, ByVal Y As Integer) As Boolean
100     If X < R.X1 Then Exit Function
102     If X > R.X2 Then Exit Function
104     If Y < R.Y1 Then Exit Function
106     If Y > R.Y2 Then Exit Function
108     InsideRectangle = True
End Function

' Fuente: https://stackoverflow.com/questions/1378604/end-process-from-task-manager-using-vb-6-code (ultima respuesta)
Public Function GetProcess(ByVal processName As String) As Byte
    
        Dim oService As Object
        Dim servicename As String
        Dim processCount As Byte
    
100     Dim oWMI As Object: Set oWMI = GetObject("winmgmts:")
102     Dim oServices As Object: Set oServices = oWMI.InstancesOf("win32_process")

104     For Each oService In oServices

106         servicename = LCase$(Trim$(CStr(oService.Name)))

108         If InStrB(1, servicename, LCase$(processName), vbBinaryCompare) > 0 Then
            
                ' Para matar un proceso adentro de este loop usar.
                'oService.Terminate
            
110             processCount = processCount + 1
            
            End If

        Next
    
112     GetProcess = processCount

End Function

Public Function EsMapaInterdimensional(ByVal Map As Integer) As Boolean
        Dim i As Integer
100     For i = 1 To UBound(MapasInterdimensionales)
102         If Map = MapasInterdimensionales(i) Then
104             EsMapaInterdimensional = True
                Exit Function
            End If
        Next
End Function

Public Function IsValidIPAddress(ByVal IP As String) As Boolean

        On Error GoTo Handler

        Dim varAddress As Variant, n As Long, lCount As Long
100     varAddress = Split(IP, ".", 4, vbTextCompare)

102     If IsArray(varAddress) Then

104         For n = LBound(varAddress) To UBound(varAddress)
106             lCount = lCount + 1
108             varAddress(n) = CByte(varAddress(n))
            Next
        
110         IsValidIPAddress = (lCount = 4)

        End If

Handler:

End Function

Function Ceil(X As Variant) As Variant
        
        On Error GoTo Ceil_Err
        
100     Ceil = IIf(Fix(X) = X, X, Fix(X) + 1)
        
        Exit Function

Ceil_Err:
102     Call TraceError(Err.Number, Err.Description & "Ceil_Err", Erl)


        
End Function

Function Clamp(X As Variant, a As Variant, b As Variant) As Variant
        
        On Error GoTo Clamp_Err
        
100     Clamp = IIf(X < a, a, IIf(X > b, b, X))
        
        Exit Function

Clamp_Err:
102     Call TraceError(Err.Number, Err.Description & "Clamp_Err", Erl)


        
End Function

Private Function GetElapsed() As Single
    Static sTime1     As Currency
    Static sTime2     As Currency
    Static sFrequency As Currency
    
    'Get the timer frequency
    If sFrequency = 0 Then
        Call QueryPerformanceFrequency(sFrequency)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(sTime1)

     'Calculate elapsed time
    GetElapsed = ((sTime1 - sTime2) / sFrequency * 1000)
    
    'Get next end time
    Call QueryPerformanceCounter(sTime2)
End Function


