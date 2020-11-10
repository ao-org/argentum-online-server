Attribute VB_Name = "PathFinding"
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

'#######################################################
'PathFinding Module
'Coded By Gulfas Morgolock
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'
'Ore is an excellent engine for introducing you not only
'to online game programming but also to general
'game programming. I am convinced that Aaron Perkings, creator
'of ORE, did a great work. He made possible that a lot of
'people enjoy for no fee games made with his engine, and
'for me, this is something great.
'
'I'd really like to contribute to this work, and all the
'projects of free ore-based MMORPGs that are on the net.
'
'I did some basic improvements on the AI of the NPCs, I
'added pathfinding, so now, the npcs are able to avoid
'obstacles. I believe that this improvement was essential
'for the engine.
'
'I'd like to see this as my contribution to ORE project,
'I hope that someone finds this source code useful.
'So, please feel free to do whatever you want with my
'pathfinging module.
'
'I'd really appreciate that if you find this source code
'useful you mention my nickname on the credits of your
'program. But there is no obligation ;).
'
'.........................................................
'Note:
'There is a little problem, ORE refers to map arrays in a
'different manner that my pathfinding routines. When I wrote
'these routines, I did it without thinking in ORE, so in my
'program I refer to maps in the usual way I do it.
'
'For example, suppose we have:
'Map(1 to Y,1 to X) as MapBlock
'I usually use the first coordinate as Y, and
'the second one as X.
'
'ORE refers to maps in converse way, for example:
'Map(1 to X,1 to Y) as MapBlock. As you can see the
'roles of first and second coordinates are different
'that my routines
'
'#######################################################

Option Explicit

Private Const ROWS   As Integer = 100

Private Const COLUMS As Integer = 100

Private Const MAXINT As Integer = 1000

Private Type tIntermidiateWork

    Known As Boolean
    DistV As Integer
    PrevV As tVertice

End Type

Dim TmpArray(1 To ROWS, 1 To COLUMS) As tIntermidiateWork

Dim TilePosY                         As Integer

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer)
        
        On Error GoTo Limites_Err
        
100     Limites = vcolu >= 1 And vcolu <= COLUMS And vfila >= 1 And vfila <= ROWS

        
        Exit Function

Limites_Err:
        Call RegistrarError(Err.Number, Err.description, "PathFinding.Limites", Erl)
        Resume Next
        
End Function

Private Function IsWalkable(ByVal Map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo IsWalkable_Err
        
100     IsWalkable = MapData(Map, row, Col).Blocked = 0 And MapData(Map, row, Col).NpcIndex = 0

102     If MapData(Map, row, Col).UserIndex <> 0 Then
104         If MapData(Map, row, Col).UserIndex <> Npclist(NpcIndex).PFINFO.TargetUser Then IsWalkable = False

        End If

        
        Exit Function

IsWalkable_Err:
        Call RegistrarError(Err.Number, Err.description, "PathFinding.IsWalkable", Erl)
        Resume Next
        
End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef T() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)
        
        On Error GoTo ProcessAdjacents_Err
        

        Dim V As tVertice

        Dim j As Integer

        'Look to North
100     j = vfila - 1

102     If Limites(j, vcolu) Then
104         If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then

                'Nos aseguramos que no hay un camino más corto
106             If T(j, vcolu).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
108                 T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
110                 T(j, vcolu).PrevV.x = vcolu
112                 T(j, vcolu).PrevV.Y = vfila
                    'Mete el vertice en la cola
114                 V.x = vcolu
116                 V.Y = j
118                 Call Push(V)

                End If

            End If

        End If

120     j = vfila + 1

        'look to south
122     If Limites(j, vcolu) Then
124         If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then

                'Nos aseguramos que no hay un camino más corto
126             If T(j, vcolu).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
128                 T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
130                 T(j, vcolu).PrevV.x = vcolu
132                 T(j, vcolu).PrevV.Y = vfila
                    'Mete el vertice en la cola
134                 V.x = vcolu
136                 V.Y = j
138                 Call Push(V)

                End If

            End If

        End If

        'look to west
140     If Limites(vfila, vcolu - 1) Then
142         If IsWalkable(MapIndex, vfila, vcolu - 1, NpcIndex) Then

                'Nos aseguramos que no hay un camino más corto
144             If T(vfila, vcolu - 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
146                 T(vfila, vcolu - 1).DistV = T(vfila, vcolu).DistV + 1
148                 T(vfila, vcolu - 1).PrevV.x = vcolu
150                 T(vfila, vcolu - 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
152                 V.x = vcolu - 1
154                 V.Y = vfila
156                 Call Push(V)

                End If

            End If

        End If

        'look to east
158     If Limites(vfila, vcolu + 1) Then
160         If IsWalkable(MapIndex, vfila, vcolu + 1, NpcIndex) Then

                'Nos aseguramos que no hay un camino más corto
162             If T(vfila, vcolu + 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
164                 T(vfila, vcolu + 1).DistV = T(vfila, vcolu).DistV + 1
166                 T(vfila, vcolu + 1).PrevV.x = vcolu
168                 T(vfila, vcolu + 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
170                 V.x = vcolu + 1
172                 V.Y = vfila
174                 Call Push(V)

                End If

            End If

        End If
   
        
        Exit Sub

ProcessAdjacents_Err:
        Call RegistrarError(Err.Number, Err.description, "PathFinding.ProcessAdjacents", Erl)
        Resume Next
        
End Sub

Public Sub SeekPath(ByVal NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
        '############################################################
        'This Sub seeks a path from the npclist(npcindex).pos
        'to the location NPCList(NpcIndex).PFINFO.Target.
        'The optional parameter MaxSteps is the maximum of steps
        'allowed for the path.
        '############################################################
        
        On Error GoTo SeekPath_Err
        

        Dim cur_npc_pos As tVertice

        Dim tar_npc_pos As tVertice

        Dim V           As tVertice

        Dim NpcMap      As Integer

        Dim steps       As Integer

100     NpcMap = Npclist(NpcIndex).Pos.Map

102     steps = 0

104     cur_npc_pos.x = Npclist(NpcIndex).Pos.Y
106     cur_npc_pos.Y = Npclist(NpcIndex).Pos.x

108     tar_npc_pos.x = Npclist(NpcIndex).PFINFO.Target.x '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.X
110     tar_npc_pos.Y = Npclist(NpcIndex).PFINFO.Target.Y '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.Y

112     Call InitializeTable(TmpArray, cur_npc_pos)
114     Call InitQueue

        'We add the first vertex to the Queue
116     Call Push(cur_npc_pos)

118     Do While (Not IsEmpty)

120         If steps > MaxSteps Then Exit Do
122         V = Pop

124         If V.x = tar_npc_pos.x And V.Y = tar_npc_pos.Y Then Exit Do
126         Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.x, NpcIndex)
        Loop

128     Call MakePath(NpcIndex)

        
        Exit Sub

SeekPath_Err:
        Call RegistrarError(Err.Number, Err.description, "PathFinding.SeekPath", Erl)
        Resume Next
        
End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
        '#######################################################
        'Builds the path previously calculated
        '#######################################################
        
        On Error GoTo MakePath_Err
        

        Dim Pasos As Integer

        Dim miV   As tVertice

        Dim i     As Integer

100     Pasos = TmpArray(Npclist(NpcIndex).PFINFO.Target.Y, Npclist(NpcIndex).PFINFO.Target.x).DistV
102     Npclist(NpcIndex).PFINFO.PathLenght = Pasos

104     If Pasos = MAXINT Then
            'MsgBox "There is no path."
106         Npclist(NpcIndex).PFINFO.NoPath = True
108         Npclist(NpcIndex).PFINFO.PathLenght = 0
            Exit Sub

        End If

110     ReDim Npclist(NpcIndex).PFINFO.Path(0 To Pasos) As tVertice

112     miV.x = Npclist(NpcIndex).PFINFO.Target.x
114     miV.Y = Npclist(NpcIndex).PFINFO.Target.Y

116     For i = Pasos To 1 Step -1
118         Npclist(NpcIndex).PFINFO.Path(i) = miV
120         miV = TmpArray(miV.Y, miV.x).PrevV
122     Next i

124     Npclist(NpcIndex).PFINFO.CurPos = 1
126     Npclist(NpcIndex).PFINFO.NoPath = False
   
        
        Exit Sub

MakePath_Err:
        Call RegistrarError(Err.Number, Err.description, "PathFinding.MakePath", Erl)
        Resume Next
        
End Sub

Private Sub InitializeTable(ByRef T() As tIntermidiateWork, ByRef S As tVertice, Optional ByVal MaxSteps As Integer = 30)
        '#########################################################
        'Initialize the array where we calculate the path
        '#########################################################
        
        On Error GoTo InitializeTable_Err
        

        Dim j As Integer, K As Integer

        Const anymap = 1

100     For j = S.Y - MaxSteps To S.Y + MaxSteps
102         For K = S.x - MaxSteps To S.x + MaxSteps

104             If InMapBounds(anymap, j, K) Then
106                 T(j, K).Known = False
108                 T(j, K).DistV = MAXINT
110                 T(j, K).PrevV.x = 0
112                 T(j, K).PrevV.Y = 0

                End If

            Next
        Next

114     T(S.Y, S.x).Known = False
116     T(S.Y, S.x).DistV = 0

        
        Exit Sub

InitializeTable_Err:
        Call RegistrarError(Err.Number, Err.description, "PathFinding.InitializeTable", Erl)
        Resume Next
        
End Sub

Function FindDirectionEAO(a As WorldPos, b As WorldPos, Optional PuedeAgu As Boolean) As Byte
        
        On Error GoTo FindDirectionEAO_Err
        
 
        Dim R  As Byte
 
        'Mejoras:
        'Ahora los NPC puden doblar las esquinas, y pasar a los lados de los arboles, _
         Tambien cuando te persiguen en ves de ir en forma orizontal y despues en vertical, te van sigsagueando.
 
        'A = NPCPOS
        'B = UserPos
 
        'Esto es para que el NPC retroceda en caso de no poder seguir adelante, en ese caso se retrocede.
 
        'Lo que no pueden hacer los Npcs, es rodear cosas, EJ:
 
        '
        ' *******************
        ' *                 *
        ' *                 *
        ' * B              *
        ' ******     ********
        '   A  <------- El npc se va a quedar loco tratando de pasar de frente en ves de rodearlo.
        '
        '
        'Saluda: <-.Siameze.->
 
        Dim PV As Integer
 
        'r = RandomNumber(1, 2)
100     R = RandomNumber(1, 2)
 
102     If a.x > b.x And a.Y > b.Y Then
104         FindDirectionEAO = IIf(R = 1, NORTH, WEST)
   
106     ElseIf a.x < b.x And a.Y < b.Y Then
108         FindDirectionEAO = IIf(R = 1, SOUTH, EAST)
   
110     ElseIf a.x < b.x And a.Y > b.Y Then
112         FindDirectionEAO = IIf(R = 1, NORTH, EAST)
   
114     ElseIf a.x > b.x And a.Y < b.Y Then
116         FindDirectionEAO = IIf(R = 1, SOUTH, WEST)
   
118     ElseIf a.x = b.x Then
120         FindDirectionEAO = IIf(a.Y < b.Y, SOUTH, NORTH)
   
122     ElseIf a.Y = b.Y Then
124         FindDirectionEAO = IIf(a.x < b.x, EAST, WEST)
 
        Else
 
126         FindDirectionEAO = 0 ' this is imposible!
   
        End If
 
128     If Distancia(a, b) > 1 Then
 
130         Select Case FindDirectionEAO
 
                Case NORTH

132                 If Not LegalPos(a.Map, a.x, a.Y - 1, PuedeAgu) Then
 
134                     If a.x > b.x Then
136                         FindDirectionEAO = WEST
138                     ElseIf a.x < b.x Then
140                         FindDirectionEAO = EAST
                        Else
142                         FindDirectionEAO = IIf(R > 1, WEST, EAST)

                        End If

144                     PV = 1
       
                    End If
 
146             Case SOUTH

148                 If Not LegalPos(a.Map, a.x, a.Y + 1, PuedeAgu) Then
 
150                     If a.x > b.x Then
152                         FindDirectionEAO = WEST
154                     ElseIf a.x < b.x Then
156                         FindDirectionEAO = EAST
                        Else
158                         FindDirectionEAO = IIf(R > 1, WEST, EAST)

                        End If

160                     PV = 1
 
                    End If
       
162             Case WEST

164                 If Not LegalPos(a.Map, a.x - 1, a.Y, PuedeAgu) Then
 
166                     If a.Y > b.Y Then
168                         FindDirectionEAO = NORTH
170                     ElseIf a.Y < b.Y Then
172                         FindDirectionEAO = SOUTH
                        Else
174                         FindDirectionEAO = IIf(R > 1, NORTH, SOUTH)

                        End If

176                     PV = 1

                    End If
       
178             Case EAST

180                 If Not LegalPos(a.Map, a.x + 1, a.Y, PuedeAgu) Then
182                     If a.Y > b.Y Then
184                         FindDirectionEAO = NORTH
186                     ElseIf a.Y < b.Y Then
188                         FindDirectionEAO = SOUTH
                        Else
190                         FindDirectionEAO = IIf(R > 1, NORTH, SOUTH)

                        End If

192                     PV = 1
   
                    End If
       
            End Select
 
194         If PV = 2 Then
 
196             Select Case FindDirectionEAO

                    Case EAST

198                     If Not LegalPos(a.Map, a.x + 1, a.Y) Then FindDirectionEAO = WEST
       
200                 Case WEST

202                     If Not LegalPos(a.Map, a.x - 1, a.Y) Then FindDirectionEAO = EAST
           
204                 Case NORTH

206                     If Not LegalPos(a.Map, a.x, a.Y - 1) Then FindDirectionEAO = SOUTH
       
208                 Case SOUTH

210                     If Not LegalPos(a.Map, a.x, a.Y + 1) Then FindDirectionEAO = NORTH
       
                End Select
   
            End If
 
        End If
 
        
        Exit Function

FindDirectionEAO_Err:
        Call RegistrarError(Err.Number, Err.description, "PathFinding.FindDirectionEAO", Erl)
        Resume Next
        
End Function
