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
'.........................................................

'###########################################################################
' CHANGES
'
' 27/03/2021 WyroX: Fixed inverted coordinates and changed algorithm to A*
'###########################################################################


Option Explicit

Private Type tIntermidiateWork
    Closed As Boolean
    Distance As Integer
    Previous As Position
    EstimatedTotalDistance As Single
End Type

Private OpenVertices(1000) As Position
Private VertexCount As Integer

Private Table(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As tIntermidiateWork

Private DirOffset(eHeading.NORTH To eHeading.WEST) As Position

Private ClosestVertex As Position
Private ClosestDistance As Single

Private Const MAXINT As Integer = 32767

' WyroX: Usada para mover memoria... VB6 es un desastre en cuanto a contenedores dinámicos
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal Length As Long)

Public Sub InitPathFinding()

    Dim Heading As eHeading, DirH As Integer
        
    For Heading = eHeading.NORTH To eHeading.WEST
        DirOffset(Heading).X = (2 - DirH) * (DirH Mod 2)
        DirOffset(Heading).Y = (DirH - 1) * (1 - (DirH Mod 2))
        DirH = DirH + 1
    Next

End Sub

Public Sub FollowPath(ByVal NpcIndex As Integer)
    Dim nextPos As WorldPos
    
    With NpcList(NpcIndex)
        nextPos.Map = .Pos.Map
        nextPos.X = .pathFindingInfo.Path(.pathFindingInfo.PathLength).X
        nextPos.Y = .pathFindingInfo.Path(.pathFindingInfo.PathLength).Y
        
        Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.Pos, nextPos))
        .pathFindingInfo.PathLength = .pathFindingInfo.PathLength - 1
    
    End With
    
End Sub

Private Function InsideLimits(ByVal X As Integer, ByVal Y As Integer)
    InsideLimits = X >= XMinMapSize And X <= XMaxMapSize And Y >= YMinMapSize And Y <= YMaxMapSize
End Function

Private Function IsWalkable(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As eHeading) As Boolean
        
    On Error GoTo ErrHandler
    
    Dim Map As Integer
    Map = NpcList(NpcIndex).Pos.Map
    
    With MapData(Map, X, Y)

        ' Otro NPC
        If .NpcIndex Then Exit Function
        
        ' Usuario
        If .UserIndex And .UserIndex <> NpcList(NpcIndex).Target Then Exit Function

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
        If .trigger = eTrigger.POSINVALIDA Then
            ' Si no es mascota
            If NpcList(NpcIndex).MaestroNPC = 0 Then Exit Function
        End If
    
        ' Tile bloqueado
        'If .Blocked And 2 ^ (Heading - 1) Then Exit Function (fix nuevos bloqueos)
        If .Blocked And eBlock.ALL_SIDES Then Exit Function
        
    End With
    
    IsWalkable = True

ErrHandler:
        
End Function

Private Sub ProcessAdjacent(ByVal NpcIndex As Integer, ByVal CurX As Integer, ByVal CurY As Integer, ByVal Heading As eHeading, ByRef EndPos As Position)

    On Error GoTo ErrHandler
    
    Dim X As Integer, Y As Integer, DistanceFromStart As Integer, EstimatedDistance As Single
    
    With DirOffset(Heading)
        X = CurX + .X
        Y = CurY + .Y
    End With
    
    With Table(X, Y)

        ' Si ya está cerrado, salimos
        If .Closed Then Exit Sub
    
        ' Nos quedamos en el campo de visión del NPC
        If InsideLimits(X, Y) Then
        
            ' Si puede atravesar el tile al siguiente
            If IsWalkable(NpcIndex, X, Y, Heading) Then
            
                ' Calculamos la distancia hasta este vértice
                DistanceFromStart = Table(CurX, CurY).Distance + 1
    
                ' Si no habíamos visitado este vértice
                If .Distance = MAXINT Then
                    ' Lo metemos en la cola
                    Call OpenVertex(X, Y)
                    
                ' Si ya lo habíamos visitado, nos fijamos si este camino es más corto
                ElseIf DistanceFromStart > .Distance Then
                    ' Es más largo, salimos
                    Exit Sub
                End If
    
                ' Guardamos la distancia desde el inicio
                .Distance = DistanceFromStart
                
                ' La distancia estimada al objetivo
                EstimatedDistance = EuclideanDistance(X, Y, EndPos)
                
                ' La distancia total estimada
                .EstimatedTotalDistance = DistanceFromStart + EstimatedDistance
                
                ' Y la posición de la que viene
                .Previous.X = CurX
                .Previous.Y = CurY
                
                ' Si la distancia total estimada es la menor hasta ahora
                If EstimatedDistance < ClosestDistance Then
                    ClosestDistance = EstimatedDistance
                    ClosestVertex.X = X
                    ClosestVertex.Y = Y
                End If
                
            End If
            
        End If

    End With

ErrHandler:

End Sub

Public Function SeekPath(ByVal NpcIndex As Integer, Optional ByVal Closest As Boolean) As Boolean
    ' Busca un camino desde la posición del NPC a la posición en .PFINFO.Target
    ' El parámetro Closest indica que en caso de que no exista un camino completo, se debe retornar el camino parcial hasta la posición más cercana al objetivo.
    ' Si Closest = True, la función devuelve True si puede moverse al menos un tile. Si Closest = False, devuelve True si se encontró un camino completo.
    ' El camino se almacena en .PFINFO.Path
    
    Dim PosNPC As Position
    Dim PosTarget As Position
    Dim Heading As eHeading, Vertex As Position
    Dim MaxDistance As Integer, index As Integer
    Dim MinTotalDistance As Integer, BestVertexIndex As Integer

    With NpcList(NpcIndex)
        PosNPC.X = .Pos.X
        PosNPC.Y = .Pos.Y
    
        ' Posición objetivo
        PosTarget.X = .pathFindingInfo.Destination.X
        PosTarget.Y = .pathFindingInfo.Destination.Y

        ' Inicializar contenedores para el algoritmo
        Call InitializeTable(Table, PosNPC, .pathFindingInfo.RangoVision)
        VertexCount = 0
        
        ' Añadimos la posición inicial a la lista
        Call OpenVertexV(PosNPC)
        
        ' Distancia máxima a calcular (distancia en tiles al target + inteligencia del NPC)
        MaxDistance = TileDistance(PosNPC, PosTarget) + .pathFindingInfo.Inteligencia
        
        ' Distancia euclideana desde la posición inicial hasta la final
        Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance = EuclideanDistanceV(PosNPC, PosTarget)
        
        ' Distancia posición inicial
        Table(PosNPC.X, PosNPC.Y).Distance = 0
        
        ' Distancia mínima
        ClosestDistance = Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance
        ClosestVertex.X = PosNPC.X
        ClosestVertex.Y = PosNPC.Y
        
    End With
    

    ' Loop principal del algoritmo
    Do While (VertexCount > 0)
    
    
        MinTotalDistance = MAXINT
        
        ' Buscamos en la cola la posición con menor distancia total
        For index = 0 To VertexCount - 1
        
            With OpenVertices(index)
            
                If Table(.X, .Y).EstimatedTotalDistance < MinTotalDistance Then
                    MinTotalDistance = Table(.X, .Y).EstimatedTotalDistance
                    BestVertexIndex = index
                End If
                
            End With
            
        Next
        
        Vertex = OpenVertices(BestVertexIndex)
        With Vertex
            ' Si es la posición objetivo
            If .X = PosTarget.X And .Y = PosTarget.Y Then
            
                ' Reconstruímos el trayecto
                Call MakePath(NpcIndex, .X, .Y)
                
                ' Salimos
                SeekPath = True
                Exit Function
                
            End If

            ' Eliminamos la posición de la cola
            Call CloseVertex(BestVertexIndex)

            ' Cerramos la posición actual
            Table(.X, .Y).Closed = True

            ' Si aún podemos seguir procesando más lejos
            If Table(.X, .Y).Distance < MaxDistance Then
            
                ' Procesamos adyacentes
                For Heading = eHeading.NORTH To eHeading.WEST
                    Call ProcessAdjacent(NpcIndex, .X, .Y, Heading, PosTarget)
                Next
                
            End If
            
        End With
        
    Loop
    
    ' No hay más nodos por procesar. O bien no existe un camino válido o el NPC no es suficientemente inteligente.
    
    ' Si debemos retornar la posición más cercana al objetivo
    If Closest Then
    
        ' Si se recorrió al menos un tile
        If ClosestVertex.X <> PosNPC.X Or ClosestVertex.Y <> PosNPC.Y Then
        
            ' Reconstruímos el camino desde la posición más cercana al objetivo
            Call MakePath(NpcIndex, ClosestVertex.X, ClosestVertex.Y)
            
            SeekPath = True
            Exit Function
            
        End If
        
    End If
    
    
    ' Llegados a este punto, invalidamos el Path del NPC
    NpcList(NpcIndex).pathFindingInfo.PathLength = 0

End Function

Private Sub MakePath(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer)

    With NpcList(NpcIndex)
        ' Obtenemos la distancia total del camino
        .pathFindingInfo.PathLength = Table(X, Y).Distance

        Dim Step As Integer
        
        ' Asignamos las coordenadas del resto camino, el final queda al inicio del array
        For Step = 1 To .pathFindingInfo.PathLength
        
            With .pathFindingInfo.Path(Step)
                .X = X
                .Y = Y
            End With

            With Table(X, Y)
                X = .Previous.X
                Y = .Previous.Y
            End With
            
        Next

    End With
   
End Sub

Private Sub InitializeTable(ByRef Table() As tIntermidiateWork, ByRef PosNPC As Position, ByVal RangoVision As Single)
    ' Inicializar la tabla de posiciones para calcular el camino.
    ' Solo limpiamos el campo de visión del NPC.

    Dim X As Integer, Y As Integer

    For Y = PosNPC.Y - RangoVision To PosNPC.Y + RangoVision
        For X = PosNPC.X - RangoVision To PosNPC.X + RangoVision
        
            If InsideLimits(X, Y) Then
                Table(X, Y).Closed = False
                Table(X, Y).Distance = MAXINT
            End If
            
        Next
    Next

End Sub

Private Function TileDistance(ByRef Vertex1 As Position, ByRef Vertex2 As Position) As Integer
    TileDistance = Abs(Vertex1.X - Vertex2.X) + Abs(Vertex1.Y - Vertex2.Y)
End Function

Private Function EuclideanDistance(ByVal X As Integer, ByVal Y As Integer, ByRef Vertex As Position) As Single
    Dim dX As Integer, dY As Integer
    dX = Vertex.X - X
    dY = Vertex.Y - Y
    EuclideanDistance = Sqr(dX * dX + dY * dY)
End Function

Private Function EuclideanDistanceV(ByRef Vertex1 As Position, ByRef Vertex2 As Position) As Single
    Dim dX As Integer, dY As Integer
    dX = Vertex1.X - Vertex2.X
    dY = Vertex1.Y - Vertex2.Y
    EuclideanDistanceV = Sqr(dX * dX + dY * dY)
End Function

Private Sub OpenVertex(ByVal X As Integer, ByVal Y As Integer)
    With OpenVertices(VertexCount)
        .X = X: .Y = Y
    End With
    VertexCount = VertexCount + 1
End Sub

Private Sub OpenVertexV(ByRef Vertex As Position)
    OpenVertices(VertexCount) = Vertex
    VertexCount = VertexCount + 1
End Sub

Private Sub CloseVertex(ByVal index As Integer)
    VertexCount = VertexCount - 1
    Call MoveMemory(OpenVertices(index), OpenVertices(index + 1), Len(OpenVertices(0)) * (VertexCount - index))
End Sub

' Las posiciones se pasan ByRef pero NO SE MODIFICAN.
Public Function GetHeadingFromWorldPos(ByRef currentPos As WorldPos, ByRef nextPos As WorldPos) As eHeading
    Dim dX As Integer, dY As Integer
    
    dX = nextPos.X - currentPos.X
    dY = nextPos.Y - currentPos.Y
    
    If dX < 0 Then
        GetHeadingFromWorldPos = eHeading.WEST
    ElseIf dX > 0 Then
        GetHeadingFromWorldPos = eHeading.EAST
    ElseIf dY < 0 Then
        GetHeadingFromWorldPos = eHeading.NORTH
    Else
        GetHeadingFromWorldPos = eHeading.SOUTH
    End If

End Function
