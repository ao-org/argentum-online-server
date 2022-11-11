Attribute VB_Name = "PathFinding"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
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

Private Type t_IntermidiateWork
    Closed As Boolean
    Distance As Integer
    Previous As t_Position
    EstimatedTotalDistance As Single
End Type

Private OpenVertices(1000) As t_Position
Private VertexCount As Integer

Private Table(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As t_IntermidiateWork

Private DirOffset(e_Heading.NORTH To e_Heading.WEST) As t_Position

Private ClosestVertex As t_Position
Private ClosestDistance As Single

Private Const MAXINT As Integer = 32767

' WyroX: Usada para mover memoria... VB6 es un desastre en cuanto a contenedores dinámicos
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal Length As Long)

Public Sub InitPathFinding()
        
        On Error GoTo InitPathFinding_Err

        Dim Heading As e_Heading, DirH As Integer
        
100     For Heading = e_Heading.NORTH To e_Heading.WEST
105         DirOffset(Heading).X = (2 - DirH) * (DirH Mod 2)
110         DirOffset(Heading).Y = (DirH - 1) * (1 - (DirH Mod 2))
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
            
105         nextPos.Map = .Pos.Map
110         nextPos.X = .pathFindingInfo.Path(.pathFindingInfo.PathLength).X
115         nextPos.Y = .pathFindingInfo.Path(.pathFindingInfo.PathLength).Y
        
120         Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.Pos, nextPos))
125         .pathFindingInfo.PathLength = .pathFindingInfo.PathLength - 1
    
        End With
      
        Exit Sub

FollowPath_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.FollowPath", Erl)
       
End Sub

Private Function InsideLimits(ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo InsideLimits_Err
        
100     InsideLimits = X >= XMinMapSize And X <= XMaxMapSize And Y >= YMinMapSize And Y <= YMaxMapSize
        
        Exit Function

InsideLimits_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.InsideLimits", Erl)

End Function

Private Function IsWalkable(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As e_Heading) As Boolean
        
    On Error GoTo ErrHandler
    
    Dim Map As Integer
1    Map = NpcList(NpcIndex).Pos.Map
    
    With MapData(Map, X, Y)

        ' Otro NPC
2        If .NpcIndex Then Exit Function
        
        ' Usuario
3        If .UserIndex And .UserIndex <> NpcList(npcIndex).TargetUser.ArrayIndex Then Exit Function

        ' Traslado
4        If .TileExit.Map Then Exit Function

        ' Agua
5        If .Blocked And FLAG_AGUA Then
            If NpcList(NpcIndex).flags.AguaValida = 0 Then Exit Function
        ' Tierra
        Else
6            If NpcList(NpcIndex).flags.TierraInvalida <> 0 Then Exit Function
        End If
        
        ' Trigger inválido para NPCs
        If .trigger = e_Trigger.POSINVALIDA Then
            ' Si no es mascota
8            If Not IsValidNpcRef(NpcList(NpcIndex).MaestroNPC) Then Exit Function
        End If
    
        ' Tile bloqueado
        If NpcList(NpcIndex).NPCtype <> e_NPCType.GuardiaReal And NpcList(NpcIndex).NPCtype <> e_NPCType.GuardiasCaos Then
9            If .Blocked And 2 ^ (Heading - 1) Then
                Exit Function
            End If
        Else
10            If (.Blocked And 2 ^ (Heading - 1)) And Not HayPuerta(Map, X + 1, Y) And Not HayPuerta(Map, X, Y) And Not HayPuerta(Map, X + 1, Y - 1) And Not HayPuerta(Map, X, Y - 1) Then Exit Function
        End If
            
    End With
    
11  IsWalkable = True
    
    Exit Function
    
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "PathFinding.IsWalkable", Erl)
    
End Function

Private Sub ProcessAdjacent(ByVal NpcIndex As Integer, ByVal CurX As Integer, ByVal CurY As Integer, ByVal Heading As e_Heading, ByRef EndPos As t_Position)

    On Error GoTo ErrHandler
    
    Dim X As Integer, Y As Integer, DistanceFromStart As Integer, EstimatedDistance As Single
    
    With DirOffset(Heading)
1        X = CurX + .X
2        Y = CurY + .Y
    End With
    
    With Table(X, Y)

        ' Si ya está cerrado, salimos
        If .Closed Then Exit Sub
    
        ' Nos quedamos en el campo de visión del NPC
        If InsideLimits(X, Y) Then
        
            ' Si puede atravesar el tile al siguiente
3            If IsWalkable(NpcIndex, X, Y, Heading) Then
            
                ' Calculamos la distancia hasta este vértice
4                DistanceFromStart = Table(CurX, CurY).Distance + 1
    
                ' Si no habíamos visitado este vértice
                If .Distance = MAXINT Then
                    ' Lo metemos en la cola
5                    Call OpenVertex(X, Y)
                    
                ' Si ya lo habíamos visitado, nos fijamos si este camino es más corto
                ElseIf DistanceFromStart > .Distance Then
                    ' Es más largo, salimos
                    Exit Sub
                End If
    
                ' Guardamos la distancia desde el inicio
6                .Distance = DistanceFromStart
                
                ' La distancia estimada al objetivo
7                EstimatedDistance = EuclideanDistance(X, Y, EndPos)
                
                ' La distancia total estimada
8                .EstimatedTotalDistance = DistanceFromStart + EstimatedDistance
                
                ' Y la posición de la que viene
9                .Previous.X = CurX
10                .Previous.Y = CurY
                
                ' Si la distancia total estimada es la menor hasta ahora
                If EstimatedDistance < ClosestDistance Then
11                    ClosestDistance = EstimatedDistance
12                    ClosestVertex.X = X
13                    ClosestVertex.Y = Y
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
        
        Dim PosNPC As t_Position
        Dim PosTarget As t_Position
        Dim Heading As e_Heading, Vertex As t_Position
        Dim MaxDistance As Integer, Index As Integer
        Dim MinTotalDistance As Integer, BestVertexIndex As Integer
        Dim UserIndex As Integer 'no es necesario
        Dim pasos As Long
        
        pasos = 0
        'Ya estamos en la posición.
        If UserIndex > 0 Then
            If NPCHasAUserInFront(NpcIndex, UserIndex) Then
                SeekPath = False
                Exit Function
            End If
        End If
        
        
100     With NpcList(NpcIndex)
105         PosNPC.X = .Pos.X
110         PosNPC.Y = .Pos.Y
    
            ' Posición objetivo
115         PosTarget.X = .pathFindingInfo.destination.X
120         PosTarget.Y = .pathFindingInfo.destination.Y
            
            ' Inicializar contenedores para el algoritmo
125         Call InitializeTable(Table, PosNPC, .pathFindingInfo.RangoVision)
130         VertexCount = 0
        
            ' Añadimos la posición inicial a la lista
135         Call OpenVertexV(PosNPC)
        
            ' Distancia máxima a calcular (distancia en tiles al target + inteligencia del NPC)
140         MaxDistance = TileDistance(PosNPC, PosTarget) + .pathFindingInfo.Inteligencia
        
            ' Distancia euclideana desde la posición inicial hasta la final
145         Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance = EuclideanDistanceV(PosNPC, PosTarget)
            
            ' Ya estamos en la posicion
            If (Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance = 0) Then
                SeekPath = False
                Exit Function
            End If
            
            ' Distancia posición inicial
150         Table(PosNPC.X, PosNPC.Y).Distance = 0
        
            ' Distancia mínima
155         ClosestDistance = Table(PosNPC.X, PosNPC.Y).EstimatedTotalDistance
160         ClosestVertex.X = PosNPC.X
165         ClosestVertex.Y = PosNPC.Y
        
        End With

        ' Loop principal del algoritmo
170     Do While (VertexCount > 0 And pasos < 300)
            
            pasos = pasos + 1
175         MinTotalDistance = MAXINT
        
            ' Buscamos en la cola la posición con menor distancia total
180         For Index = 0 To VertexCount - 1
        
185             With OpenVertices(Index)
            
190                 If Table(.X, .Y).EstimatedTotalDistance < MinTotalDistance Then
195                     MinTotalDistance = Table(.X, .Y).EstimatedTotalDistance
200                     BestVertexIndex = Index
                    End If
                
                End With
            
            Next
        
205         Vertex = OpenVertices(BestVertexIndex)
210         With Vertex
                ' Si es la posición objetivo
215             If .X = PosTarget.X And .Y = PosTarget.Y Then
            
                    ' Reconstruímos el trayecto
220                 Call MakePath(NpcIndex, .X, .Y)
                
                    ' Salimos
225                 SeekPath = True
                    Exit Function
                
                End If

                ' Eliminamos la posición de la cola
230             Call CloseVertex(BestVertexIndex)

                ' Cerramos la posición actual
235             Table(.X, .Y).Closed = True

                ' Si aún podemos seguir procesando más lejos
240             If Table(.X, .Y).Distance < MaxDistance Then
            
                    ' Procesamos adyacentes
245                 For Heading = e_Heading.NORTH To e_Heading.WEST
250                     Call ProcessAdjacent(NpcIndex, .X, .Y, Heading, PosTarget)
                    Next
                
                End If
            
            End With
        
        Loop
    
        ' No hay más nodos por procesar. O bien no existe un camino válido o el NPC no es suficientemente inteligente.
    
        ' Si debemos retornar la posición más cercana al objetivo
255     If Closest Then
    
            ' Si se recorrió al menos un tile
260         If ClosestVertex.X <> PosNPC.X Or ClosestVertex.Y <> PosNPC.Y Then
        
                ' Reconstruímos el camino desde la posición más cercana al objetivo
265             Call MakePath(NpcIndex, ClosestVertex.X, ClosestVertex.Y)
            
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

Private Sub MakePath(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo MakePath_Err
 
100     With NpcList(NpcIndex)
            ' Obtenemos la distancia total del camino
105         .pathFindingInfo.PathLength = Table(X, Y).Distance

            Dim Step As Integer
        
            ' Asignamos las coordenadas del resto camino, el final queda al inicio del array
110         For Step = 1 To UBound(.pathFindingInfo.Path) ' .pathFindingInfo.PathLength TODO
        
115             With .pathFindingInfo.Path(Step)
120                 .X = X
125                 .Y = Y
                End With
                If X > 0 And Y > 0 Then
130                 With Table(X, Y)
135                     X = .Previous.X
140                     Y = .Previous.Y
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

        Dim X As Integer, Y As Integer

100     For Y = PosNPC.Y - RangoVision To PosNPC.Y + RangoVision
105         For X = PosNPC.X - RangoVision To PosNPC.X + RangoVision
        
110             If InsideLimits(X, Y) Then
115                 Table(X, Y).Closed = False
120                 Table(X, Y).Distance = MAXINT
                End If
            
            Next
        Next

        
        Exit Sub

InitializeTable_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.InitializeTable", Erl)
  
End Sub

Private Function TileDistance(ByRef Vertex1 As t_Position, ByRef Vertex2 As t_Position) As Integer
        
        On Error GoTo TileDistance_Err
        
100     TileDistance = Abs(Vertex1.X - Vertex2.X) + Abs(Vertex1.Y - Vertex2.Y)
        
        Exit Function

TileDistance_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.TileDistance", Erl)
  
End Function

Private Function EuclideanDistance(ByVal X As Integer, ByVal Y As Integer, ByRef Vertex As t_Position) As Single
        
        On Error GoTo EuclideanDistance_Err
        
        Dim dX As Integer, dY As Integer
100     dX = Vertex.X - X
105     dY = Vertex.Y - Y
110     EuclideanDistance = Sqr(dX * dX + dY * dY)
        
        Exit Function

EuclideanDistance_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.EuclideanDistance", Erl)
      
End Function

Private Function EuclideanDistanceV(ByRef Vertex1 As t_Position, ByRef Vertex2 As t_Position) As Single
        
        On Error GoTo EuclideanDistanceV_Err
        
        Dim dX As Integer, dY As Integer
100     dX = Vertex1.X - Vertex2.X
105     dY = Vertex1.Y - Vertex2.Y
110     EuclideanDistanceV = Sqr(dX * dX + dY * dY)
        
        Exit Function

EuclideanDistanceV_Err:
        Call TraceError(Err.Number, Err.Description, "PathFinding.EuclideanDistanceV", Erl)

End Function

Private Sub OpenVertex(ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo OpenVertex_Err
        
100     With OpenVertices(VertexCount)
105         .X = X: .Y = Y
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
Public Function GetHeadingFromWorldPos(ByRef currentPos As t_WorldPos, ByRef nextPos As t_WorldPos) As e_Heading
        
        On Error GoTo GetHeadingFromWorldPos_Err
        
        Dim dX As Integer, dY As Integer
    
100     dX = nextPos.X - currentPos.X
105     dY = nextPos.Y - currentPos.Y
    
110     If dX < 0 Then
115         GetHeadingFromWorldPos = e_Heading.WEST
120     ElseIf dX > 0 Then
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
