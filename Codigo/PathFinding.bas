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
        
100     For Heading = eHeading.NORTH To eHeading.WEST
102         DirOffset(Heading).X = (2 - DirH) * (DirH Mod 2)
104         DirOffset(Heading).Y = (DirH - 1) * (1 - (DirH Mod 2))
106         DirH = DirH + 1
        Next

End Sub

Public Sub FollowPath(ByVal NpcIndex As Integer)
        Dim nextPos As WorldPos
    
100     With NpcList(NpcIndex)
102         nextPos.Map = .Pos.Map
104         nextPos.X = .pathFindingInfo.Path(.pathFindingInfo.PathLength).X
106         nextPos.Y = .pathFindingInfo.Path(.pathFindingInfo.PathLength).Y
        
108         Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.Pos, nextPos))
110         .pathFindingInfo.PathLength = .pathFindingInfo.PathLength - 1
    
        End With
    
End Sub

Private Function InsideLimits(ByVal X As Integer, ByVal Y As Integer)
100     InsideLimits = X >= XMinMapSize And X <= XMaxMapSize And Y >= YMinMapSize And Y <= YMaxMapSize
End Function

Private Function IsWalkable(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As eHeading) As Boolean
        
        On Error GoTo ErrHandler
    
        Dim Map As Integer
100     Map = NpcList(NpcIndex).Pos.Map
    
102     With MapData(Map, X, Y)

            ' Otro NPC
104         If .NpcIndex Then Exit Function
        
            ' Usuario
106         If .UserIndex And .UserIndex <> NpcList(NpcIndex).Target Then Exit Function

            ' Traslado
108         If .TileExit.Map Then Exit Function

            ' Agua
110         If .Blocked And FLAG_AGUA Then
112             If NpcList(NpcIndex).flags.AguaValida = 0 Then Exit Function
            ' Tierra
            Else
114             If NpcList(NpcIndex).flags.TierraInvalida <> 0 Then Exit Function
            End If
        
            ' Trigger inválido para NPCs
116         If .trigger = eTrigger.POSINVALIDA Then
                ' Si no es mascota
118             If NpcList(NpcIndex).MaestroNPC = 0 Then Exit Function
            End If
    
            ' Tile bloqueado
            'If .Blocked And 2 ^ (Heading - 1) Then Exit Function (fix nuevos bloqueos)
120         If .Blocked And eBlock.ALL_SIDES Then Exit Function
        
        End With
    
122     IsWalkable = True

ErrHandler:
        
End Function

Private Sub ProcessAdjacent(ByVal NpcIndex As Integer, ByVal CurX As Integer, ByVal CurY As Integer, ByVal Heading As eHeading, ByRef EndPos As Position)

        On Error GoTo ErrHandler
    
        Dim X As Integer, Y As Integer, DistanceFromStart As Integer, EstimatedDistance As Single
    
100     With DirOffset(Heading)
102         X = CurX + .X
104         Y = CurY + .Y
        End With
    
106     With Table(X, Y)

            ' Si ya está cerrado, salimos
108         If .Closed Then Exit Sub
    
            ' Nos quedamos en el campo de visión del NPC
110         If InsideLimits(X, Y) Then
        
                ' Si puede atravesar el tile al siguiente
112             If IsWalkable(NpcIndex, X, Y, Heading) Then
            
                    ' Calculamos la distancia hasta este vértice
114                 DistanceFromStart = Table(CurX, CurY).Distance + 1
    
                    ' Si no habíamos visitado este vértice
116                 If .Distance = MAXINT Then
                        ' Lo metemos en la cola
118                     Call OpenVertex(X, Y)
                    
                    ' Si ya lo habíamos visitado, nos fijamos si este camino es más corto
120                 ElseIf DistanceFromStart > .Distance Then
                        ' Es más largo, salimos
                        Exit Sub
                    End If
    
                    ' Guardamos la distancia desde el inicio
122                 .Distance = DistanceFromStart
                
                    ' La distancia estimada al objetivo
124                 EstimatedDistance = EuclideanDistance(X, Y, EndPos)
                
                    ' La distancia total estimada
126                 .EstimatedTotalDistance = DistanceFromStart + EstimatedDistance
                
                    ' Y la posición de la que viene
128                 .Previous.X = CurX
130                 .Previous.Y = CurY
                
                    ' Si la distancia total estimada es la menor hasta ahora
132                 If EstimatedDistance < ClosestDistance Then
134                     ClosestDistance = EstimatedDistance
136                     ClosestVertex.X = X
138                     ClosestVertex.Y = Y
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

100     With NpcList(NpcIndex)
102         PosNPC.X = .Pos.X
104         PosNPC.Y = .Pos.Y
    
            ' Posición objetivo
106         PosTarget.X = .pathFindingInfo.destination.X
108         PosTarget.Y = .pathFindingInfo.destination.Y

            ' Inicializar contenedores para el algoritmo
110         Call InitializeTable(Table, PosNPC, .pathFindingInfo.RangoVision)
112         VertexCount = 0
        
            ' Añadimos la posición inicial a la lista
114         Call OpenVertexV(PosNPC)
        
            ' Distancia máxima a calcular (distancia en tiles al target + inteligencia del NPC)
116         MaxDistance = TileDistance(PosNPC, PosTarget) + .pathFindingInfo.Inteligencia
        
            ' Distancia euclideana desde la posición inicial hasta la final
118         Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance = EuclideanDistanceV(PosNPC, PosTarget)
        
            ' Distancia posición inicial
120         Table(PosNPC.X, PosNPC.Y).Distance = 0
        
            ' Distancia mínima
122         ClosestDistance = Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance
124         ClosestVertex.X = PosNPC.X
126         ClosestVertex.Y = PosNPC.Y
        
        End With
    

        ' Loop principal del algoritmo
128     Do While (VertexCount > 0)
    
    
130         MinTotalDistance = MAXINT
        
            ' Buscamos en la cola la posición con menor distancia total
132         For index = 0 To VertexCount - 1
        
134             With OpenVertices(index)
            
136                 If Table(.X, .Y).EstimatedTotalDistance < MinTotalDistance Then
138                     MinTotalDistance = Table(.X, .Y).EstimatedTotalDistance
140                     BestVertexIndex = index
                    End If
                
                End With
            
            Next
        
142         Vertex = OpenVertices(BestVertexIndex)
144         With Vertex
                ' Si es la posición objetivo
146             If .X = PosTarget.X And .Y = PosTarget.Y Then
            
                    ' Reconstruímos el trayecto
148                 Call MakePath(NpcIndex, .X, .Y)
                
                    ' Salimos
150                 SeekPath = True
                    Exit Function
                
                End If

                ' Eliminamos la posición de la cola
152             Call CloseVertex(BestVertexIndex)

                ' Cerramos la posición actual
154             Table(.X, .Y).Closed = True

                ' Si aún podemos seguir procesando más lejos
156             If Table(.X, .Y).Distance < MaxDistance Then
            
                    ' Procesamos adyacentes
158                 For Heading = eHeading.NORTH To eHeading.WEST
160                     Call ProcessAdjacent(NpcIndex, .X, .Y, Heading, PosTarget)
                    Next
                
                End If
            
            End With
        
        Loop
    
        ' No hay más nodos por procesar. O bien no existe un camino válido o el NPC no es suficientemente inteligente.
    
        ' Si debemos retornar la posición más cercana al objetivo
162     If Closest Then
    
            ' Si se recorrió al menos un tile
164         If ClosestVertex.X <> PosNPC.X Or ClosestVertex.Y <> PosNPC.Y Then
        
                ' Reconstruímos el camino desde la posición más cercana al objetivo
166             Call MakePath(NpcIndex, ClosestVertex.X, ClosestVertex.Y)
            
168             SeekPath = True
                Exit Function
            
            End If
        
        End If
    
    
        ' Llegados a este punto, invalidamos el Path del NPC
170     NpcList(NpcIndex).pathFindingInfo.PathLength = 0

End Function

Private Sub MakePath(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer)

100     With NpcList(NpcIndex)
            ' Obtenemos la distancia total del camino
102         .pathFindingInfo.PathLength = Table(X, Y).Distance

            Dim Step As Integer
        
            ' Asignamos las coordenadas del resto camino, el final queda al inicio del array
104         For Step = 1 To .pathFindingInfo.PathLength
        
106             With .pathFindingInfo.Path(Step)
108                 .X = X
110                 .Y = Y
                End With

112             With Table(X, Y)
114                 X = .Previous.X
116                 Y = .Previous.Y
                End With
            
            Next

        End With
   
End Sub

Private Sub InitializeTable(ByRef Table() As tIntermidiateWork, ByRef PosNPC As Position, ByVal RangoVision As Single)
        ' Inicializar la tabla de posiciones para calcular el camino.
        ' Solo limpiamos el campo de visión del NPC.

        Dim X As Integer, Y As Integer

100     For Y = PosNPC.Y - RangoVision To PosNPC.Y + RangoVision
102         For X = PosNPC.X - RangoVision To PosNPC.X + RangoVision
        
104             If InsideLimits(X, Y) Then
106                 Table(X, Y).Closed = False
108                 Table(X, Y).Distance = MAXINT
                End If
            
            Next
        Next

End Sub

Private Function TileDistance(ByRef Vertex1 As Position, ByRef Vertex2 As Position) As Integer
100     TileDistance = Abs(Vertex1.X - Vertex2.X) + Abs(Vertex1.Y - Vertex2.Y)
End Function

Private Function EuclideanDistance(ByVal X As Integer, ByVal Y As Integer, ByRef Vertex As Position) As Single
        Dim dX As Integer, dY As Integer
100     dX = Vertex.X - X
102     dY = Vertex.Y - Y
104     EuclideanDistance = Sqr(dX * dX + dY * dY)
End Function

Private Function EuclideanDistanceV(ByRef Vertex1 As Position, ByRef Vertex2 As Position) As Single
        Dim dX As Integer, dY As Integer
100     dX = Vertex1.X - Vertex2.X
102     dY = Vertex1.Y - Vertex2.Y
104     EuclideanDistanceV = Sqr(dX * dX + dY * dY)
End Function

Private Sub OpenVertex(ByVal X As Integer, ByVal Y As Integer)
100     With OpenVertices(VertexCount)
102         .X = X: .Y = Y
        End With
104     VertexCount = VertexCount + 1
End Sub

Private Sub OpenVertexV(ByRef Vertex As Position)
100     OpenVertices(VertexCount) = Vertex
102     VertexCount = VertexCount + 1
End Sub

Private Sub CloseVertex(ByVal index As Integer)
100     VertexCount = VertexCount - 1
102     Call MoveMemory(OpenVertices(index), OpenVertices(index + 1), Len(OpenVertices(0)) * (VertexCount - index))
End Sub

' Las posiciones se pasan ByRef pero NO SE MODIFICAN.
Public Function GetHeadingFromWorldPos(ByRef currentPos As WorldPos, ByRef nextPos As WorldPos) As eHeading
        Dim dX As Integer, dY As Integer
    
100     dX = nextPos.X - currentPos.X
102     dY = nextPos.Y - currentPos.Y
    
104     If dX < 0 Then
106         GetHeadingFromWorldPos = eHeading.WEST
108     ElseIf dX > 0 Then
110         GetHeadingFromWorldPos = eHeading.EAST
112     ElseIf dY < 0 Then
114         GetHeadingFromWorldPos = eHeading.NORTH
        Else
116         GetHeadingFromWorldPos = eHeading.SOUTH
        End If

End Function
