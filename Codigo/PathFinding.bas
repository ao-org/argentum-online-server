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
'  Usada para mover memoria... VB6 es un desastre en cuanto a contenedores dinámicos
Private Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal length As Long)

Public Sub InitPathFinding()
    On Error GoTo InitPathFinding_Err
    Dim Heading As e_Heading, DirH As Integer
    For Heading = e_Heading.NORTH To e_Heading.WEST
        DirOffset(Heading).x = (2 - DirH) * (DirH Mod 2)
        DirOffset(Heading).y = (DirH - 1) * (1 - (DirH Mod 2))
        DirH = DirH + 1
    Next
    Exit Sub
InitPathFinding_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.InitPathFinding", Erl)
End Sub

Public Sub FollowPath(ByVal NpcIndex As Integer)
    On Error GoTo FollowPath_Err
    Dim nextPos As t_WorldPos
    With NpcList(NpcIndex)
        If (.pathFindingInfo.PathLength > UBound(.pathFindingInfo.Path)) Then ' Fix temporal para que no explote el LOG
            Exit Sub
        End If
        nextPos.Map = .pos.Map
        nextPos.x = .pathFindingInfo.Path(.pathFindingInfo.PathLength).x
        nextPos.y = .pathFindingInfo.Path(.pathFindingInfo.PathLength).y
        Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.pos, nextPos))
        .pathFindingInfo.PathLength = .pathFindingInfo.PathLength - 1
    End With
    Exit Sub
FollowPath_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.FollowPath", Erl)
End Sub

Private Function InsideLimits(ByVal x As Integer, ByVal y As Integer)
    On Error GoTo InsideLimits_Err
    InsideLimits = x >= XMinMapSize And x <= XMaxMapSize And y >= YMinMapSize And y <= YMaxMapSize
    Exit Function
InsideLimits_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.InsideLimits", Erl)
End Function

Private Function IsWalkable(ByVal NpcIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Heading As e_Heading) As Boolean
    On Error GoTo ErrHandler
    Dim Map As Integer
    Map = NpcList(NpcIndex).pos.Map
    With MapData(Map, x, y)
        ' Otro NPC
        If .NpcIndex Then Exit Function
        ' Usuario
        If .UserIndex And .UserIndex <> NpcList(NpcIndex).TargetUser.ArrayIndex Then Exit Function
        ' Traslado
        If .TileExit.Map Then Exit Function
        ' Agua
        If .Blocked And FLAG_AGUA Then
            If NpcList(NpcIndex).flags.AguaValida = 0 Then Exit Function
            ' Tierra
        Else
            If NpcList(NpcIndex).flags.TierraInvalida <> 0 Then Exit Function
        End If
        ' Trigger inválido para NPCs
        If .trigger = e_Trigger.POSINVALIDA Then
            ' Si no es mascota
            If Not IsValidNpcRef(NpcList(NpcIndex).MaestroNPC) Then Exit Function
        End If
        ' Tile bloqueado
        If NpcList(NpcIndex).npcType <> e_NPCType.GuardiaReal And NpcList(NpcIndex).npcType <> e_NPCType.GuardiasCaos Then
            If .Blocked And 2 ^ (Heading - 1) Then
                Exit Function
            End If
        Else
            If (.Blocked And 2 ^ (Heading - 1)) And Not HayPuerta(Map, x + 1, y) And Not HayPuerta(Map, x, y) And Not HayPuerta(Map, x + 1, y - 1) And Not HayPuerta(Map, x, y - _
                    1) Then Exit Function
        End If
    End With
    IsWalkable = True
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "PathFinding.IsWalkable", Erl)
End Function

Private Sub ProcessAdjacent(ByVal NpcIndex As Integer, ByVal CurX As Integer, ByVal CurY As Integer, ByVal Heading As e_Heading, ByRef EndPos As t_Position)
    On Error GoTo ErrHandler
    Dim x As Integer, y As Integer, DistanceFromStart As Integer, EstimatedDistance As Single
    With DirOffset(Heading)
        x = CurX + .x
        y = CurY + .y
    End With
    With Table(x, y)
        ' Si ya está cerrado, salimos
        If .Closed Then Exit Sub
        ' Nos quedamos en el campo de visión del NPC
        If InsideLimits(x, y) Then
            ' Si puede atravesar el tile al siguiente
            If IsWalkable(NpcIndex, x, y, Heading) Then
                ' Calculamos la distancia hasta este vértice
                DistanceFromStart = Table(CurX, CurY).Distance + 1
                ' Si no habíamos visitado este vértice
                If .Distance = MAX_INTEGER Then
                    ' Lo metemos en la cola
                    Call OpenVertex(x, y)
                    ' Si ya lo habíamos visitado, nos fijamos si este camino es más corto
                ElseIf DistanceFromStart > .Distance Then
                    ' Es más largo, salimos
                    Exit Sub
                End If
                ' Guardamos la distancia desde el inicio
                .Distance = DistanceFromStart
                ' La distancia estimada al objetivo
                EstimatedDistance = EuclideanDistance(x, y, EndPos)
                ' La distancia total estimada
                .EstimatedTotalDistance = DistanceFromStart + EstimatedDistance
                ' Y la posición de la que viene
                .Previous.x = CurX
                .Previous.y = CurY
                ' Si la distancia total estimada es la menor hasta ahora
                If EstimatedDistance < ClosestDistance Then
                    ClosestDistance = EstimatedDistance
                    ClosestVertex.x = x
                    ClosestVertex.y = y
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "PathFinding.ProcessAdjacent", Erl)
End Sub

Public Function SeekPath(ByVal NpcIndex As Integer, Optional ByVal Closest As Boolean) As Boolean
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
    With NpcList(NpcIndex)
        PosNPC.x = .pos.x
        PosNPC.y = .pos.y
        ' Posición objetivo
        PosTarget.x = .pathFindingInfo.destination.x
        PosTarget.y = .pathFindingInfo.destination.y
        ' Inicializar contenedores para el algoritmo
        Call InitializeTable(Table, PosNPC, .pathFindingInfo.RangoVision)
        VertexCount = 0
        ' Añadimos la posición inicial a la lista
        Call OpenVertexV(PosNPC)
        ' Distancia máxima a calcular (distancia en tiles al target + inteligencia del NPC)
        MaxDistance = Min(MAX_PATH_LENGTH, TileDistance(PosNPC, PosTarget) + .pathFindingInfo.RangoVision)
        ' Distancia euclideana desde la posición inicial hasta la final
        Table(PosNPC.x, PosNPC.y).EstimatedTotalDistance = EuclideanDistanceV(PosNPC, PosTarget)
        ' Ya estamos en la posicion
        If (Table(PosNPC.x, PosNPC.y).EstimatedTotalDistance = 0) Then
            SeekPath = False
            Exit Function
        End If
        ' Distancia posición inicial
        Table(PosNPC.x, PosNPC.y).Distance = 0
        ' Distancia mínima
        ClosestDistance = Table(PosNPC.x, PosNPC.y).EstimatedTotalDistance
        ClosestVertex.x = PosNPC.x
        ClosestVertex.y = PosNPC.y
    End With
    ' Loop principal del algoritmo
    Dim max_steps As Integer
    max_steps = SvrConfig.GetValue("NPC_PATHFINDING_MAX_STEPS")
    Debug.Assert max_steps < MAX_PATH_LENGTH
    Do While (VertexCount > 0 And pasos < max_steps)
        pasos = pasos + 1
        MinTotalDistance = MAX_INTEGER
        ' Buscamos en la cola la posición con menor distancia total
        For Index = 0 To VertexCount - 1
            With OpenVertices(Index)
                If Table(.x, .y).EstimatedTotalDistance < MinTotalDistance Then
                    MinTotalDistance = Table(.x, .y).EstimatedTotalDistance
                    BestVertexIndex = Index
                End If
            End With
        Next
        Vertex = OpenVertices(BestVertexIndex)
        With Vertex
            ' Si es la posición objetivo
            If .x = PosTarget.x And .y = PosTarget.y Then
                ' Reconstruímos el trayecto
                Call MakePath(NpcIndex, .x, .y)
                ' Salimos
                SeekPath = True
                Exit Function
            End If
            ' Eliminamos la posición de la cola
            Call CloseVertex(BestVertexIndex)
            ' Cerramos la posición actual
            Table(.x, .y).Closed = True
            ' Si aún podemos seguir procesando más lejos
            If Table(.x, .y).Distance < MaxDistance Then
                ' Procesamos adyacentes
                For Heading = e_Heading.NORTH To e_Heading.WEST
                    Call ProcessAdjacent(NpcIndex, .x, .y, Heading, PosTarget)
                Next
            End If
        End With
    Loop
    ' No hay más nodos por procesar. O bien no existe un camino válido o el NPC no es suficientemente inteligente.
    ' Si debemos retornar la posición más cercana al objetivo
    If Closest Then
        ' Si se recorrió al menos un tile
        If ClosestVertex.x <> PosNPC.x Or ClosestVertex.y <> PosNPC.y Then
            ' Reconstruímos el camino desde la posición más cercana al objetivo
            Call MakePath(NpcIndex, ClosestVertex.x, ClosestVertex.y)
            SeekPath = True
            Exit Function
        End If
    End If
    ' Llegados a este punto, invalidamos el Path del NPC
    NpcList(NpcIndex).pathFindingInfo.PathLength = 0
    Exit Function
SeekPath_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.SeekPath", Erl)
End Function

Private Sub MakePath(ByVal NpcIndex As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo MakePath_Err
    With NpcList(NpcIndex)
        ' Obtenemos la distancia total del camino
        .pathFindingInfo.PathLength = Table(x, y).Distance
        Dim step As Integer
        ' Asignamos las coordenadas del resto camino, el final queda al inicio del array
        For step = 1 To UBound(.pathFindingInfo.Path) ' .pathFindingInfo.PathLength TODO
            With .pathFindingInfo.Path(step)
                .x = x
                .y = y
            End With
            If x > 0 And y > 0 Then
                With Table(x, y)
                    x = .Previous.x
                    y = .Previous.y
                End With
            End If
        Next
    End With
    Exit Sub
MakePath_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.MakePath", Erl)
End Sub

Private Sub InitializeTable(ByRef Table() As t_IntermidiateWork, ByRef PosNPC As t_Position, ByVal RangoVision As Single)
    ' Inicializar la tabla de posiciones para calcular el camino.
    ' Solo limpiamos el campo de visión del NPC.
    On Error GoTo InitializeTable_Err
    Dim x As Integer, y As Integer
    For y = PosNPC.y - RangoVision To PosNPC.y + RangoVision
        For x = PosNPC.x - RangoVision To PosNPC.x + RangoVision
            If InsideLimits(x, y) Then
                Table(x, y).Closed = False
                Table(x, y).Distance = MAX_INTEGER
            End If
        Next
    Next
    Exit Sub
InitializeTable_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.InitializeTable", Erl)
End Sub

Private Function TileDistance(ByRef Vertex1 As t_Position, ByRef Vertex2 As t_Position) As Integer
    On Error GoTo TileDistance_Err
    TileDistance = Abs(Vertex1.x - Vertex2.x) + Abs(Vertex1.y - Vertex2.y)
    Exit Function
TileDistance_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.TileDistance", Erl)
End Function

Private Function EuclideanDistance(ByVal x As Integer, ByVal y As Integer, ByRef Vertex As t_Position) As Single
    On Error GoTo EuclideanDistance_Err
    Dim dx As Integer, dY As Integer
    dx = Vertex.x - x
    dY = Vertex.y - y
    EuclideanDistance = Sqr(dx * dx + dY * dY)
    Exit Function
EuclideanDistance_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.EuclideanDistance", Erl)
End Function

Private Function EuclideanDistanceV(ByRef Vertex1 As t_Position, ByRef Vertex2 As t_Position) As Single
    On Error GoTo EuclideanDistanceV_Err
    Dim dx As Integer, dY As Integer
    dx = Vertex1.x - Vertex2.x
    dY = Vertex1.y - Vertex2.y
    EuclideanDistanceV = Sqr(dx * dx + dY * dY)
    Exit Function
EuclideanDistanceV_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.EuclideanDistanceV", Erl)
End Function

Private Sub OpenVertex(ByVal x As Integer, ByVal y As Integer)
    On Error GoTo OpenVertex_Err
    With OpenVertices(VertexCount)
        .x = x: .y = y
    End With
    VertexCount = VertexCount + 1
    Exit Sub
OpenVertex_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.OpenVertex", Erl)
End Sub

Private Sub OpenVertexV(ByRef Vertex As t_Position)
    On Error GoTo OpenVertexV_Err
    OpenVertices(VertexCount) = Vertex
    VertexCount = VertexCount + 1
    Exit Sub
OpenVertexV_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.OpenVertexV", Erl)
End Sub

Private Sub CloseVertex(ByVal Index As Integer)
    On Error GoTo CloseVertex_Err
    VertexCount = VertexCount - 1
    Call MoveMemory(OpenVertices(Index), OpenVertices(Index + 1), Len(OpenVertices(0)) * (VertexCount - Index))
    Exit Sub
CloseVertex_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.CloseVertex", Erl)
End Sub

' Las posiciones se pasan ByRef pero NO SE MODIFICAN.
Public Function GetHeadingFromWorldPos(ByRef currentPos As t_WorldPos, ByRef nextPos As t_WorldPos) As e_Heading
    On Error GoTo GetHeadingFromWorldPos_Err
    Dim dx As Integer, dY As Integer
    dx = nextPos.x - currentPos.x
    dY = nextPos.y - currentPos.y
    If dx < 0 Then
        GetHeadingFromWorldPos = e_Heading.WEST
    ElseIf dx > 0 Then
        GetHeadingFromWorldPos = e_Heading.EAST
    ElseIf dY < 0 Then
        GetHeadingFromWorldPos = e_Heading.NORTH
    Else
        GetHeadingFromWorldPos = e_Heading.SOUTH
    End If
    Exit Function
GetHeadingFromWorldPos_Err:
    Call TraceError(Err.Number, Err.Description, "PathFinding.GetHeadingFromWorldPos", Erl)
End Function
