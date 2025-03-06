Attribute VB_Name = "PathFinding"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Private Type t_IntermidiateWork
    Closed As Boolean
    Distance As Integer
    Previous As t_Position
    EstimatedTotalDistance As Single

End Type

Private OpenVertices(1000)                                            As t_Position
Private VertexCount                                                   As Integer
Private Table(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As t_IntermidiateWork
Private DirOffset(e_Heading.NORTH To e_Heading.WEST)                  As t_Position
Private ClosestVertex                                                 As t_Position
Private ClosestDistance                                               As Single
Private Const MAXINT                                                  As Integer = 32767
' WyroX: Usada para mover memoria... VB6 es un desastre en cuanto a contenedores dinámicos
Private Declare Sub MoveMemory _
                Lib "Kernel32" _
                Alias "RtlMoveMemory" (pDest As Any, _
                                       pSource As Any, _
                                       ByVal length As Long)

Public Sub InitPathFinding()

        On Error GoTo InitPathFinding_Err

        Dim Heading As e_Heading, DirH As Integer

100     For Heading = e_Heading.NORTH To e_Heading.WEST
105         DirOffset(Heading).x = (2 - DirH) * (DirH Mod 2)
110         DirOffset(Heading).y = (DirH - 1) * (1 - (DirH Mod 2))
115         DirH = DirH + 1
        Next
        Exit Sub
InitPathFinding_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.InitPathFinding", Erl)

End Sub

Public Sub FollowPath(ByVal NpcIndex As Integer)

        On Error GoTo FollowPath_Err

        Dim nextPos As t_WorldPos

100     With NpcList(NpcIndex)

            If (.pathFindingInfo.PathLength > UBound(.pathFindingInfo.Path)) Then ' Fix temporal para que no explote el LOG
                Exit Sub

            End If

105         nextPos.Map = .pos.Map
110         nextPos.x = .pathFindingInfo.Path(.pathFindingInfo.PathLength).x
115         nextPos.y = .pathFindingInfo.Path(.pathFindingInfo.PathLength).y
120         Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.pos, nextPos))
125         .pathFindingInfo.PathLength = .pathFindingInfo.PathLength - 1

        End With

        Exit Sub
FollowPath_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.FollowPath", Erl)

End Sub

Private Function InsideLimits(ByVal x As Integer, ByVal y As Integer)

        On Error GoTo InsideLimits_Err

100     InsideLimits = x >= XMinMapSize And x <= XMaxMapSize And y >= YMinMapSize And y <= YMaxMapSize
        Exit Function
InsideLimits_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.InsideLimits", Erl)

End Function

Private Function IsWalkable(ByVal NpcIndex As Integer, _
                            ByVal x As Integer, _
                            ByVal y As Integer, _
                            ByVal Heading As e_Heading) As Boolean

    On Error GoTo ErrHandler

    Dim Map As Integer
1   Map = NpcList(NpcIndex).pos.Map

    With MapData(Map, x, y)

        ' Otro NPC
2       If .NpcIndex Then Exit Function

        ' Usuario
3       If .UserIndex And .UserIndex <> NpcList(NpcIndex).TargetUser.ArrayIndex Then Exit Function

        ' Traslado
4       If .TileExit.Map Then Exit Function

        ' Agua
5       If .Blocked And FLAG_AGUA Then
            If NpcList(NpcIndex).flags.AguaValida = 0 Then Exit Function
            ' Tierra
        Else

6           If NpcList(NpcIndex).flags.TierraInvalida <> 0 Then Exit Function

        End If

        ' Trigger inválido para NPCs
        If .trigger = e_Trigger.POSINVALIDA Then

            ' Si no es mascota
8           If Not IsValidNpcRef(NpcList(NpcIndex).MaestroNPC) Then Exit Function

        End If

        ' Tile bloqueado
        If NpcList(NpcIndex).npcType <> e_NPCType.GuardiaReal And NpcList(NpcIndex).npcType <> e_NPCType.GuardiasCaos Then
9           If .Blocked And 2 ^ (Heading - 1) Then
                Exit Function

            End If

        Else

10          If (.Blocked And 2 ^ (Heading - 1)) And Not HayPuerta(Map, x + 1, y) And Not HayPuerta(Map, x, y) And Not HayPuerta(Map, x + 1, y - 1) And Not HayPuerta(Map, x, y - 1) Then Exit Function

        End If

    End With

11  IsWalkable = True
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "PathFinding.IsWalkable", Erl)

End Function

Private Sub ProcessAdjacent(ByVal NpcIndex As Integer, _
                            ByVal CurX As Integer, _
                            ByVal CurY As Integer, _
                            ByVal Heading As e_Heading, _
                            ByRef EndPos As t_Position)

    On Error GoTo ErrHandler

    Dim x As Integer, y As Integer, DistanceFromStart As Integer, EstimatedDistance As Single

    With DirOffset(Heading)
1       x = CurX + .x
2       y = CurY + .y

    End With

    With Table(x, y)

        ' Si ya está cerrado, salimos
        If .Closed Then Exit Sub

        ' Nos quedamos en el campo de visión del NPC
        If InsideLimits(x, y) Then

            ' Si puede atravesar el tile al siguiente
3           If IsWalkable(NpcIndex, x, y, Heading) Then
                ' Calculamos la distancia hasta este vértice
4               DistanceFromStart = Table(CurX, CurY).Distance + 1

                ' Si no habíamos visitado este vértice
                If .Distance = MAXINT Then
                    ' Lo metemos en la cola
5                   Call OpenVertex(x, y)
                    ' Si ya lo habíamos visitado, nos fijamos si este camino es más corto
                ElseIf DistanceFromStart > .Distance Then
                    ' Es más largo, salimos
                    Exit Sub

                End If

                ' Guardamos la distancia desde el inicio
6               .Distance = DistanceFromStart
                ' La distancia estimada al objetivo
7               EstimatedDistance = EuclideanDistance(x, y, EndPos)
                ' La distancia total estimada
8               .EstimatedTotalDistance = DistanceFromStart + EstimatedDistance
                ' Y la posición de la que viene
9               .Previous.x = CurX
10              .Previous.y = CurY

                ' Si la distancia total estimada es la menor hasta ahora
                If EstimatedDistance < ClosestDistance Then
11                  ClosestDistance = EstimatedDistance
12                  ClosestVertex.x = x
13                  ClosestVertex.y = y

                End If

            End If

        End If

    End With

    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "PathFinding.ProcessAdjacent", Erl)

End Sub

Public Function SeekPath(ByVal NpcIndex As Integer, _
                         Optional ByVal Closest As Boolean) As Boolean

        ' Busca un camino desde la posición del NPC a la posición en .PFINFO.Target
        ' El parámetro Closest indica que en caso de que no exista un camino completo, se debe retornar el camino parcial hasta la posición más cercana al objetivo.
        ' Si Closest = True, la función devuelve True si puede moverse al menos un tile. Si Closest = False, devuelve True si se encontró un camino completo.
        ' El camino se almacena en .PFINFO.Path
        On Error GoTo SeekPath_Err

        Dim PosNPC           As t_Position
        Dim PosTarget        As t_Position
        Dim Heading          As e_Heading, Vertex As t_Position
        Dim MaxDistance      As Integer, Index As Integer
        Dim MinTotalDistance As Integer, BestVertexIndex As Integer
        Dim UserIndex        As Integer 'no es necesario
        Dim pasos            As Long
        pasos = 0

        'Ya estamos en la posición.
        If UserIndex > 0 Then
            If NPCHasAUserInFront(NpcIndex, UserIndex) Then
                SeekPath = False
                Exit Function

            End If

        End If

100     With NpcList(NpcIndex)
105         PosNPC.x = .pos.x
110         PosNPC.y = .pos.y
            ' Posición objetivo
115         PosTarget.x = .pathFindingInfo.destination.x
120         PosTarget.y = .pathFindingInfo.destination.y
            ' Inicializar contenedores para el algoritmo
125         Call InitializeTable(Table, PosNPC, .pathFindingInfo.RangoVision)
130         VertexCount = 0
            ' Añadimos la posición inicial a la lista
135         Call OpenVertexV(PosNPC)
            ' Distancia máxima a calcular (distancia en tiles al target + inteligencia del NPC)
140         MaxDistance = TileDistance(PosNPC, PosTarget) + .pathFindingInfo.Inteligencia
            ' Distancia euclideana desde la posición inicial hasta la final
145         Table(PosNPC.x, PosNPC.y).EstimatedTotalDistance = EuclideanDistanceV(PosNPC, PosTarget)

            ' Ya estamos en la posicion
            If (Table(PosNPC.x, PosNPC.y).EstimatedTotalDistance = 0) Then
                SeekPath = False
                Exit Function

            End If

            ' Distancia posición inicial
150         Table(PosNPC.x, PosNPC.y).Distance = 0
            ' Distancia mínima
155         ClosestDistance = Table(PosNPC.x, PosNPC.y).EstimatedTotalDistance
160         ClosestVertex.x = PosNPC.x
165         ClosestVertex.y = PosNPC.y

        End With

        ' Loop principal del algoritmo
170     Do While (VertexCount > 0 And pasos < 300)
            pasos = pasos + 1
175         MinTotalDistance = MAXINT

            ' Buscamos en la cola la posición con menor distancia total
180         For Index = 0 To VertexCount - 1

185             With OpenVertices(Index)

190                 If Table(.x, .y).EstimatedTotalDistance < MinTotalDistance Then
195                     MinTotalDistance = Table(.x, .y).EstimatedTotalDistance
200                     BestVertexIndex = Index

                    End If

                End With

            Next
205         Vertex = OpenVertices(BestVertexIndex)

210         With Vertex

                ' Si es la posición objetivo
215             If .x = PosTarget.x And .y = PosTarget.y Then
                    ' Reconstruímos el trayecto
220                 Call MakePath(NpcIndex, .x, .y)
                    ' Salimos
225                 SeekPath = True
                    Exit Function

                End If

                ' Eliminamos la posición de la cola
230             Call CloseVertex(BestVertexIndex)
                ' Cerramos la posición actual
235             Table(.x, .y).Closed = True

                ' Si aún podemos seguir procesando más lejos
240             If Table(.x, .y).Distance < MaxDistance Then

                    ' Procesamos adyacentes
245                 For Heading = e_Heading.NORTH To e_Heading.WEST
250                     Call ProcessAdjacent(NpcIndex, .x, .y, Heading, PosTarget)
                    Next

                End If

            End With

        Loop

        ' No hay más nodos por procesar. O bien no existe un camino válido o el NPC no es suficientemente inteligente.
        ' Si debemos retornar la posición más cercana al objetivo
255     If Closest Then

            ' Si se recorrió al menos un tile
260         If ClosestVertex.x <> PosNPC.x Or ClosestVertex.y <> PosNPC.y Then
                ' Reconstruímos el camino desde la posición más cercana al objetivo
265             Call MakePath(NpcIndex, ClosestVertex.x, ClosestVertex.y)
270             SeekPath = True
                Exit Function

            End If

        End If

        ' Llegados a este punto, invalidamos el Path del NPC
275     NpcList(NpcIndex).pathFindingInfo.PathLength = 0
        Exit Function
SeekPath_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.SeekPath", Erl)

End Function

Private Sub MakePath(ByVal NpcIndex As Integer, ByVal x As Integer, ByVal y As Integer)

        On Error GoTo MakePath_Err

100     With NpcList(NpcIndex)
            ' Obtenemos la distancia total del camino
105         .pathFindingInfo.PathLength = Table(x, y).Distance

            Dim step As Integer

            ' Asignamos las coordenadas del resto camino, el final queda al inicio del array
110         For step = 1 To UBound(.pathFindingInfo.Path) ' .pathFindingInfo.PathLength TODO

115             With .pathFindingInfo.Path(step)
120                 .x = x
125                 .y = y

                End With

                If x > 0 And y > 0 Then

130                 With Table(x, y)
135                     x = .Previous.x
140                     y = .Previous.y

                    End With

                End If

            Next

        End With

        Exit Sub
MakePath_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.MakePath", Erl)

End Sub

Private Sub InitializeTable(ByRef Table() As t_IntermidiateWork, _
                            ByRef PosNPC As t_Position, _
                            ByVal RangoVision As Single)

        ' Inicializar la tabla de posiciones para calcular el camino.
        ' Solo limpiamos el campo de visión del NPC.
        On Error GoTo InitializeTable_Err

        Dim x As Integer, y As Integer

100     For y = PosNPC.y - RangoVision To PosNPC.y + RangoVision
105         For x = PosNPC.x - RangoVision To PosNPC.x + RangoVision

110             If InsideLimits(x, y) Then
115                 Table(x, y).Closed = False
120                 Table(x, y).Distance = MAXINT

                End If

            Next
        Next
        Exit Sub
InitializeTable_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.InitializeTable", Erl)

End Sub

Private Function TileDistance(ByRef Vertex1 As t_Position, _
                              ByRef Vertex2 As t_Position) As Integer

        On Error GoTo TileDistance_Err

100     TileDistance = Abs(Vertex1.x - Vertex2.x) + Abs(Vertex1.y - Vertex2.y)
        Exit Function
TileDistance_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.TileDistance", Erl)

End Function

Private Function EuclideanDistance(ByVal x As Integer, _
                                   ByVal y As Integer, _
                                   ByRef Vertex As t_Position) As Single

        On Error GoTo EuclideanDistance_Err

        Dim dx As Integer, dY As Integer
100     dx = Vertex.x - x
105     dY = Vertex.y - y
110     EuclideanDistance = Sqr(dx * dx + dY * dY)
        Exit Function
EuclideanDistance_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.EuclideanDistance", Erl)

End Function

Private Function EuclideanDistanceV(ByRef Vertex1 As t_Position, _
                                    ByRef Vertex2 As t_Position) As Single

        On Error GoTo EuclideanDistanceV_Err

        Dim dx As Integer, dY As Integer
100     dx = Vertex1.x - Vertex2.x
105     dY = Vertex1.y - Vertex2.y
110     EuclideanDistanceV = Sqr(dx * dx + dY * dY)
        Exit Function
EuclideanDistanceV_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.EuclideanDistanceV", Erl)

End Function

Private Sub OpenVertex(ByVal x As Integer, ByVal y As Integer)

        On Error GoTo OpenVertex_Err

100     With OpenVertices(VertexCount)
105         .x = x: .y = y

        End With

110     VertexCount = VertexCount + 1
        Exit Sub
OpenVertex_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.OpenVertex", Erl)

End Sub

Private Sub OpenVertexV(ByRef Vertex As t_Position)

        On Error GoTo OpenVertexV_Err

100     OpenVertices(VertexCount) = Vertex
105     VertexCount = VertexCount + 1
        Exit Sub
OpenVertexV_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.OpenVertexV", Erl)

End Sub

Private Sub CloseVertex(ByVal Index As Integer)

        On Error GoTo CloseVertex_Err

100     VertexCount = VertexCount - 1
105     Call MoveMemory(OpenVertices(Index), OpenVertices(Index + 1), Len(OpenVertices(0)) * (VertexCount - Index))
        Exit Sub
CloseVertex_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.CloseVertex", Erl)

End Sub

' Las posiciones se pasan ByRef pero NO SE MODIFICAN.
Public Function GetHeadingFromWorldPos(ByRef currentPos As t_WorldPos, _
                                       ByRef nextPos As t_WorldPos) As e_Heading

        On Error GoTo GetHeadingFromWorldPos_Err

        Dim dx As Integer, dY As Integer
100     dx = nextPos.x - currentPos.x
105     dY = nextPos.y - currentPos.y

110     If dx < 0 Then
115         GetHeadingFromWorldPos = e_Heading.WEST
120     ElseIf dx > 0 Then
125         GetHeadingFromWorldPos = e_Heading.EAST
130     ElseIf dY < 0 Then
135         GetHeadingFromWorldPos = e_Heading.NORTH
        Else
140         GetHeadingFromWorldPos = e_Heading.SOUTH

        End If

        Exit Function
GetHeadingFromWorldPos_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.GetHeadingFromWorldPos", Erl)

End Function
