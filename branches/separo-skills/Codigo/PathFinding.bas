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
    Limites = vcolu >= 1 And vcolu <= COLUMS And vfila >= 1 And vfila <= ROWS

End Function

Private Function IsWalkable(ByVal Map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
    IsWalkable = MapData(Map, row, Col).Blocked = 0 And MapData(Map, row, Col).NpcIndex = 0

    If MapData(Map, row, Col).UserIndex <> 0 Then
        If MapData(Map, row, Col).UserIndex <> Npclist(NpcIndex).PFINFO.TargetUser Then IsWalkable = False

    End If

End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef T() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)

    Dim V As tVertice

    Dim j As Integer

    'Look to North
    j = vfila - 1

    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then

            'Nos aseguramos que no hay un camino más corto
            If T(j, vcolu).DistV = MAXINT Then
                'Actualizamos la tabla de calculos intermedios
                T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                T(j, vcolu).PrevV.x = vcolu
                T(j, vcolu).PrevV.Y = vfila
                'Mete el vertice en la cola
                V.x = vcolu
                V.Y = j
                Call Push(V)

            End If

        End If

    End If

    j = vfila + 1

    'look to south
    If Limites(j, vcolu) Then
        If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then

            'Nos aseguramos que no hay un camino más corto
            If T(j, vcolu).DistV = MAXINT Then
                'Actualizamos la tabla de calculos intermedios
                T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                T(j, vcolu).PrevV.x = vcolu
                T(j, vcolu).PrevV.Y = vfila
                'Mete el vertice en la cola
                V.x = vcolu
                V.Y = j
                Call Push(V)

            End If

        End If

    End If

    'look to west
    If Limites(vfila, vcolu - 1) Then
        If IsWalkable(MapIndex, vfila, vcolu - 1, NpcIndex) Then

            'Nos aseguramos que no hay un camino más corto
            If T(vfila, vcolu - 1).DistV = MAXINT Then
                'Actualizamos la tabla de calculos intermedios
                T(vfila, vcolu - 1).DistV = T(vfila, vcolu).DistV + 1
                T(vfila, vcolu - 1).PrevV.x = vcolu
                T(vfila, vcolu - 1).PrevV.Y = vfila
                'Mete el vertice en la cola
                V.x = vcolu - 1
                V.Y = vfila
                Call Push(V)

            End If

        End If

    End If

    'look to east
    If Limites(vfila, vcolu + 1) Then
        If IsWalkable(MapIndex, vfila, vcolu + 1, NpcIndex) Then

            'Nos aseguramos que no hay un camino más corto
            If T(vfila, vcolu + 1).DistV = MAXINT Then
                'Actualizamos la tabla de calculos intermedios
                T(vfila, vcolu + 1).DistV = T(vfila, vcolu).DistV + 1
                T(vfila, vcolu + 1).PrevV.x = vcolu
                T(vfila, vcolu + 1).PrevV.Y = vfila
                'Mete el vertice en la cola
                V.x = vcolu + 1
                V.Y = vfila
                Call Push(V)

            End If

        End If

    End If
   
End Sub

Public Sub SeekPath(ByVal NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
    '############################################################
    'This Sub seeks a path from the npclist(npcindex).pos
    'to the location NPCList(NpcIndex).PFINFO.Target.
    'The optional parameter MaxSteps is the maximum of steps
    'allowed for the path.
    '############################################################

    Dim cur_npc_pos As tVertice

    Dim tar_npc_pos As tVertice

    Dim V           As tVertice

    Dim NpcMap      As Integer

    Dim steps       As Integer

    NpcMap = Npclist(NpcIndex).Pos.Map

    steps = 0

    cur_npc_pos.x = Npclist(NpcIndex).Pos.Y
    cur_npc_pos.Y = Npclist(NpcIndex).Pos.x

    tar_npc_pos.x = Npclist(NpcIndex).PFINFO.Target.x '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.X
    tar_npc_pos.Y = Npclist(NpcIndex).PFINFO.Target.Y '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.Y

    Call InitializeTable(TmpArray, cur_npc_pos)
    Call InitQueue

    'We add the first vertex to the Queue
    Call Push(cur_npc_pos)

    Do While (Not IsEmpty)

        If steps > MaxSteps Then Exit Do
        V = Pop

        If V.x = tar_npc_pos.x And V.Y = tar_npc_pos.Y Then Exit Do
        Call ProcessAdjacents(NpcMap, TmpArray, V.Y, V.x, NpcIndex)
    Loop

    Call MakePath(NpcIndex)

End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
    '#######################################################
    'Builds the path previously calculated
    '#######################################################

    Dim Pasos As Integer

    Dim miV   As tVertice

    Dim i     As Integer

    Pasos = TmpArray(Npclist(NpcIndex).PFINFO.Target.Y, Npclist(NpcIndex).PFINFO.Target.x).DistV
    Npclist(NpcIndex).PFINFO.PathLenght = Pasos

    If Pasos = MAXINT Then
        'MsgBox "There is no path."
        Npclist(NpcIndex).PFINFO.NoPath = True
        Npclist(NpcIndex).PFINFO.PathLenght = 0
        Exit Sub

    End If

    ReDim Npclist(NpcIndex).PFINFO.Path(0 To Pasos) As tVertice

    miV.x = Npclist(NpcIndex).PFINFO.Target.x
    miV.Y = Npclist(NpcIndex).PFINFO.Target.Y

    For i = Pasos To 1 Step -1
        Npclist(NpcIndex).PFINFO.Path(i) = miV
        miV = TmpArray(miV.Y, miV.x).PrevV
    Next i

    Npclist(NpcIndex).PFINFO.CurPos = 1
    Npclist(NpcIndex).PFINFO.NoPath = False
   
End Sub

Private Sub InitializeTable(ByRef T() As tIntermidiateWork, ByRef S As tVertice, Optional ByVal MaxSteps As Integer = 30)
    '#########################################################
    'Initialize the array where we calculate the path
    '#########################################################

    Dim j As Integer, K As Integer

    Const anymap = 1

    For j = S.Y - MaxSteps To S.Y + MaxSteps
        For K = S.x - MaxSteps To S.x + MaxSteps

            If InMapBounds(anymap, j, K) Then
                T(j, K).Known = False
                T(j, K).DistV = MAXINT
                T(j, K).PrevV.x = 0
                T(j, K).PrevV.Y = 0

            End If

        Next
    Next

    T(S.Y, S.x).Known = False
    T(S.Y, S.x).DistV = 0

End Sub

Function FindDirectionEAO(a As WorldPos, b As WorldPos, Optional PuedeAgu As Boolean) As Byte
 
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
    R = RandomNumber(1, 2)
 
    If a.x > b.x And a.Y > b.Y Then
        FindDirectionEAO = IIf(R = 1, NORTH, WEST)
   
    ElseIf a.x < b.x And a.Y < b.Y Then
        FindDirectionEAO = IIf(R = 1, SOUTH, EAST)
   
    ElseIf a.x < b.x And a.Y > b.Y Then
        FindDirectionEAO = IIf(R = 1, NORTH, EAST)
   
    ElseIf a.x > b.x And a.Y < b.Y Then
        FindDirectionEAO = IIf(R = 1, SOUTH, WEST)
   
    ElseIf a.x = b.x Then
        FindDirectionEAO = IIf(a.Y < b.Y, SOUTH, NORTH)
   
    ElseIf a.Y = b.Y Then
        FindDirectionEAO = IIf(a.x < b.x, EAST, WEST)
 
    Else
 
        FindDirectionEAO = 0 ' this is imposible!
   
    End If
 
    If Distancia(a, b) > 1 Then
 
        Select Case FindDirectionEAO
 
            Case NORTH

                If Not LegalPos(a.Map, a.x, a.Y - 1, PuedeAgu) Then
 
                    If a.x > b.x Then
                        FindDirectionEAO = WEST
                    ElseIf a.x < b.x Then
                        FindDirectionEAO = EAST
                    Else
                        FindDirectionEAO = IIf(R > 1, WEST, EAST)

                    End If

                    PV = 1
       
                End If
 
            Case SOUTH

                If Not LegalPos(a.Map, a.x, a.Y + 1, PuedeAgu) Then
 
                    If a.x > b.x Then
                        FindDirectionEAO = WEST
                    ElseIf a.x < b.x Then
                        FindDirectionEAO = EAST
                    Else
                        FindDirectionEAO = IIf(R > 1, WEST, EAST)

                    End If

                    PV = 1
 
                End If
       
            Case WEST

                If Not LegalPos(a.Map, a.x - 1, a.Y, PuedeAgu) Then
 
                    If a.Y > b.Y Then
                        FindDirectionEAO = NORTH
                    ElseIf a.Y < b.Y Then
                        FindDirectionEAO = SOUTH
                    Else
                        FindDirectionEAO = IIf(R > 1, NORTH, SOUTH)

                    End If

                    PV = 1

                End If
       
            Case EAST

                If Not LegalPos(a.Map, a.x + 1, a.Y, PuedeAgu) Then
                    If a.Y > b.Y Then
                        FindDirectionEAO = NORTH
                    ElseIf a.Y < b.Y Then
                        FindDirectionEAO = SOUTH
                    Else
                        FindDirectionEAO = IIf(R > 1, NORTH, SOUTH)

                    End If

                    PV = 1
   
                End If
       
        End Select
 
        If PV = 2 Then
 
            Select Case FindDirectionEAO

                Case EAST

                    If Not LegalPos(a.Map, a.x + 1, a.Y) Then FindDirectionEAO = WEST
       
                Case WEST

                    If Not LegalPos(a.Map, a.x - 1, a.Y) Then FindDirectionEAO = EAST
           
                Case NORTH

                    If Not LegalPos(a.Map, a.x, a.Y - 1) Then FindDirectionEAO = SOUTH
       
                Case SOUTH

                    If Not LegalPos(a.Map, a.x, a.Y + 1) Then FindDirectionEAO = NORTH
       
            End Select
   
        End If
 
    End If
 
End Function
