Attribute VB_Name = "AI"
Option Explicit

' WyroX: Hardcodeada de la vida...
Public Const FUEGOFATUO      As Integer = 964
Public Const ELEMENTAL_VIENTO      As Integer = 963
Public Const ELEMENTAL_FUEGO      As Integer = 962

'Damos a los NPCs el mismo rango de vison que un PJ
Public Const RANGO_VISION_X  As Byte = 11
Public Const RANGO_VISION_Y  As Byte = 9


Public Sub NpcAI(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler

100     With NpcList(NpcIndex)
102         Select Case .Movement
                Case e_TipoAI.Estatico
                    ' Es un NPC estatico, no hace nada.
                    Exit Sub

104             Case e_TipoAI.MueveAlAzar
106                 If .Hostile = 1 Then
108                     Call PerseguirUsuarioCercano(NpcIndex)
                    Else
110                     Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)
                    End If

112             Case e_TipoAI.NpcDefensa
114                 Call SeguirAgresor(NpcIndex)

116             Case e_TipoAI.NpcAtacaNpc
118                 Call AI_NpcAtacaNpc(NpcIndex)

120             Case e_TipoAI.SigueAmo
122                 Call SeguirAmo(NpcIndex)

124             Case e_TipoAI.Caminata
126                 Call HacerCaminata(NpcIndex)

128             Case e_TipoAI.Invasion
130                 Call MovimientoInvasion(NpcIndex)

132             Case e_TipoAI.GuardiaPersigueNpc
134                 Call AI_GuardiaPersigueNpc(NpcIndex)


            End Select

        End With

        Exit Sub

ErrorHandler:
    
136     Call LogError("NPC.AI " & NpcList(NpcIndex).Name & " " & NpcList(NpcIndex).MaestroNPC & " mapa:" & NpcList(NpcIndex).Pos.Map & " x:" & NpcList(NpcIndex).Pos.X & " y:" & NpcList(NpcIndex).Pos.Y & " Mov:" & NpcList(NpcIndex).Movement & " TargU:" & NpcList(NpcIndex).Target & " TargN:" & NpcList(NpcIndex).TargetNPC)

138     Dim MiNPC As t_Npc: MiNPC = NpcList(NpcIndex)
    
140     Call QuitarNPC(NpcIndex)
142     Call ReSpawnNpc(MiNPC)

End Sub

Private Sub PerseguirUsuarioCercano(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler

        Dim i            As Long
        Dim UserIndex    As Integer
        Dim npcEraPasivo As Boolean
        Dim agresor      As Integer
        Dim minDistancia As Integer
        Dim minDistanciaAtacable As Integer
        Dim enemigoCercano As Integer
        Dim enemigoAtacableMasCercano As Integer
    
        ' Numero muy grande para que siempre haya un mÃƒÂ­nimo
100     minDistancia = 32000
102     minDistanciaAtacable = 32000

104     With NpcList(NpcIndex)
106         npcEraPasivo = .flags.OldHostil = 0
108         .Target = 0
110         .TargetNPC = 0

112         If .flags.AttackedBy <> vbNullString Then
114           agresor = NameIndex(.flags.AttackedBy)
            End If

            ' Busco algun objetivo en el area.
116         For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
118             UserIndex = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)

120             If EsObjetivoValido(NpcIndex, UserIndex) Then

                    ' Busco el mas cercano, sea atacable o no.
122                 If Distancia(UserList(UserIndex).Pos, .Pos) < minDistancia Then
124                     enemigoCercano = UserIndex
126                     minDistancia = Distancia(UserList(UserIndex).Pos, .Pos)
                    End If

                    ' Busco el mas cercano que sea atacable.
128                 If (UsuarioAtacableConMagia(UserIndex) Or UsuarioAtacableConMelee(NpcIndex, UserIndex)) And Distancia(UserList(UserIndex).Pos, .Pos) < minDistanciaAtacable Then
130                     enemigoAtacableMasCercano = UserIndex
132                     minDistanciaAtacable = Distancia(UserList(UserIndex).Pos, .Pos)
                    End If

                End If

134         Next i

            ' Al terminar el `for`, puedo tener un maximo de tres objetivos distintos.
            ' Por prioridad, vamos a decidir estas cosas en orden.

136         If npcEraPasivo Then
                ' Significa que alguien le pego, y esta en modo agresivo trantando de darle.
                ' El unico objetivo que importa aca es el atacante; los demas son ignorados.
138             If EnRangoVision(NpcIndex, agresor) Then .Target = agresor

            Else ' El NPC es hostil siempre, le quiere pegar a alguien.

140             If minDistanciaAtacable > 0 And enemigoAtacableMasCercano > 0 Then ' Hay alguien atacable cerca
142                 .Target = enemigoAtacableMasCercano
144             ElseIf enemigoCercano > 0 Then ' Hay alguien cerca, pero no es atacable
146                 .Target = enemigoCercano
                End If

            End If

            ' Si el NPC tiene un objetivo
148         If .Target > 0 Then
150             Call AI_AtacarUsuarioObjetivo(NpcIndex)
            Else
152             If .NPCtype <> e_NPCType.GuardiaReal And .NPCtype <> e_NPCType.GuardiasCaos Then
154                 Call RestoreOldMovement(NpcIndex)
                    ' No encontro a nadie cerca, camina unos pasos en cualquier direccion.
156                 Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)
                   
                Else
158                 If Distancia(.Pos, .Orig) > 0 Then
160                     Call AI_CaminarConRumbo(NpcIndex, .Orig)
                    Else
162                     If .Char.Heading <> e_Heading.SOUTH Then
164                         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, e_Heading.SOUTH)
                        End If
                    End If
                End If
            End If

        End With

        Exit Sub

ErrorHandler:
166     Call TraceError(Err.Number, Err.Description, "AI_NPC.PerseguirUsuarioCercano", Erl)

End Sub

' Cuando un NPC no tiene target y se puede mover libremente pero cerca de su lugar de origen.
' La mayoria de los NPC deberian mantenerse cerca de su posicion de origen, algunos quedaran quietos
' en su posicion y otros se moveran libremente cerca de su posicion de origen.
Private Sub AI_CaminarSinRumboCercaDeOrigen(ByVal NpcIndex As Integer)
        On Error GoTo AI_CaminarSinRumboCercaDeOrigen_Err

100     With NpcList(NpcIndex)
102         If .flags.Paralizado > 0 Or .flags.Inmovilizado > 0 Then
104             Call AnimacionIdle(NpcIndex, True)
106         ElseIf Distancia(.Pos, .Orig) > 4 Then
108             Call AI_CaminarConRumbo(NpcIndex, .Orig)
110         ElseIf RandomNumber(1, 6) = 3 Then
112             Call MoveNPCChar(NpcIndex, CByte(RandomNumber(e_Heading.NORTH, e_Heading.WEST)))
            Else
114             Call AnimacionIdle(NpcIndex, True)
            End If

        End With

        Exit Sub

AI_CaminarSinRumboCercaDeOrigen_Err:
116     Call TraceError(Err.Number, Err.Description, "AI.AI_CaminarSinRumboCercaDeOrigen_Err", Erl)

        
End Sub

' Cuando un NPC no tiene target y se tiene que mover libremente
Private Sub AI_CaminarSinRumbo(ByVal NpcIndex As Integer)

        On Error GoTo AI_CaminarSinRumbo_Err

100     With NpcList(NpcIndex)

102         If RandomNumber(1, 6) = 3 And .flags.Paralizado = 0 And .flags.Inmovilizado = 0 Then
104             Call MoveNPCChar(NpcIndex, CByte(RandomNumber(e_Heading.NORTH, e_Heading.WEST)))
            Else
106             Call AnimacionIdle(NpcIndex, True)

            End If

        End With

        Exit Sub

AI_CaminarSinRumbo_Err:
108     Call TraceError(Err.Number, Err.Description, "AI.AI_CaminarSinRumbo", Erl)

        
End Sub

Private Sub AI_CaminarConRumbo(ByVal NpcIndex As Integer, ByRef rumbo As t_WorldPos)
        On Error GoTo AI_CaminarConRumbo_Err
    
100     If NpcList(NpcIndex).flags.Paralizado Or NpcList(NpcIndex).flags.Inmovilizado Then
102         Call AnimacionIdle(NpcIndex, True)
            Exit Sub
        End If
    
104     With NpcList(NpcIndex).pathFindingInfo
            ' Si no tiene un camino calculado o si el destino cambio
106         If .PathLength = 0 Or .destination.X <> rumbo.X Or .destination.Y <> rumbo.Y Then
108             .destination.X = rumbo.X
110             .destination.Y = rumbo.Y

                ' Recalculamos el camino
112             If SeekPath(NpcIndex, True) Then
                    ' Si consiguo un camino
114                 Call FollowPath(NpcIndex)
                End If
            Else ' Avanzamos en el camino
116             Call FollowPath(NpcIndex)
            End If

        End With

        Exit Sub

AI_CaminarConRumbo_Err:
        Dim errorDescription As String
118     errorDescription = Err.Description & vbNewLine & " NpcIndex: " & NpcIndex & " NPCList.size= " & UBound(NpcList)
120     Call TraceError(Err.Number, errorDescription, "AI.AI_CaminarConRumbo", Erl)

End Sub


Private Sub AI_AtacarUsuarioObjetivo(ByVal AtackerNpcIndex As Integer)
        On Error GoTo ErrorHandler

        Dim AtacaConMagia As Boolean
        Dim AtacaMelee As Boolean
        Dim EstaPegadoAlUsuario As Boolean
        Dim tHeading As Byte
    
100     With NpcList(AtackerNpcIndex)
102         If .Target = 0 Then Exit Sub
        
104         EstaPegadoAlUsuario = (Distancia(.Pos, UserList(.Target).Pos) <= 1)
106         AtacaConMagia = (.flags.LanzaSpells And IntervaloPermiteLanzarHechizo(AtackerNpcIndex) And (RandomNumber(1, 100) <= 50 Or Not EstaPegadoAlUsuario))
108         AtacaMelee = (EstaPegadoAlUsuario And UsuarioAtacableConMelee(AtackerNpcIndex, .Target) And .flags.Paralizado = 0 And Not AtacaConMagia)

110         If AtacaConMagia Then
                ' Le lanzo un Hechizo
112             Call NpcLanzaUnSpell(AtackerNpcIndex)
114         ElseIf AtacaMelee Then
                ' Se da vuelta y enfrenta al Usuario
116             tHeading = GetHeadingFromWorldPos(.Pos, UserList(.Target).Pos)
118             Call AnimacionIdle(AtackerNpcIndex, True)
120             Call ChangeNPCChar(AtackerNpcIndex, .Char.Body, .Char.Head, tHeading)

                ' Le pego al Usuario
122             Call NpcAtacaUser(AtackerNpcIndex, .Target, tHeading)
            End If

124         If UsuarioAtacableConMagia(.Target) Or UsuarioAtacableConMelee(AtackerNpcIndex, .Target) Then
                ' Si no tiene un camino pero esta pegado al usuario, no queremos gastar tiempo calculando caminos.
126             If .pathFindingInfo.PathLength = 0 And EstaPegadoAlUsuario Then Exit Sub
            
128             Call AI_CaminarConRumbo(AtackerNpcIndex, UserList(.Target).Pos)
            Else
130             Call AI_CaminarSinRumboCercaDeOrigen(AtackerNpcIndex)
            End If
        End With

        Exit Sub

ErrorHandler:
132     Call TraceError(Err.Number, Err.Description, "AIv2.AI_AtacarUsuarioObjetivo", Erl)

End Sub

Public Sub AI_GuardiaPersigueNpc(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler
        Dim targetPos As t_WorldPos
        
100     With NpcList(NpcIndex)
        
102          If .TargetNPC > 0 Then
104             targetPos = NpcList(.TargetNPC).Pos
                
106             If Distancia(.Pos, targetPos) <= 1 Then
108                 Call SistemaCombate.NpcAtacaNpc(NpcIndex, .TargetNPC, False)
                End If
                
110             If DistanciaRadial(.Orig, targetPos) <= (DIAMETRO_VISION_GUARDIAS_NPCS \ 2) Then
112                 If NpcList(.TargetNPC).Target = 0 Then
114                     Call AI_CaminarConRumbo(NpcIndex, targetPos)
116                 ElseIf UserList(NpcList(.TargetNPC).Target).flags.NPCAtacado <> .TargetNPC Then
118                     Call AI_CaminarConRumbo(NpcIndex, targetPos)
                    Else
120                     .TargetNPC = 0
122                     Call AI_CaminarConRumbo(NpcIndex, .Orig)
                    End If
                Else
124                 .TargetNPC = 0
126                 Call AI_CaminarConRumbo(NpcIndex, .Orig)
                End If
                
            Else
128             .TargetNPC = BuscarNpcEnArea(NpcIndex)
130             If Distancia(.Pos, .Orig) > 0 Then
132                 Call AI_CaminarConRumbo(NpcIndex, .Orig)
                Else
134                 Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, e_Heading.SOUTH)
                End If
            End If
            
            
        End With
        
        Exit Sub
        
        
ErrorHandler:
136     Call TraceError(Err.Number, Err.Description, "AIv2.AI_GuardiaAtacaNpc", Erl)


End Sub

Private Function DistanciaRadial(OrigenPos As t_WorldPos, DestinoPos As t_WorldPos) As Long
100     DistanciaRadial = max(Abs(OrigenPos.X - DestinoPos.X), Abs(OrigenPos.Y - DestinoPos.Y))
End Function

Private Function BuscarNpcEnArea(ByVal NpcIndex As Integer) As Integer
        
        On Error GoTo BuscarNpcEnArea
        
        Dim X As Byte, Y As Byte
       
100    With NpcList(NpcIndex)
       
102         For X = (.Orig.X - (DIAMETRO_VISION_GUARDIAS_NPCS \ 2)) To (.Orig.X + (DIAMETRO_VISION_GUARDIAS_NPCS \ 2))
104             For Y = (.Orig.Y - (DIAMETRO_VISION_GUARDIAS_NPCS \ 2)) To (.Orig.Y + (DIAMETRO_VISION_GUARDIAS_NPCS \ 2))
                
106                 If MapData(.Orig.Map, X, Y).NpcIndex > 0 And NpcIndex <> MapData(.Orig.Map, X, Y).NpcIndex Then
                        Dim foundNpc As Integer
108                     foundNpc = MapData(.Orig.Map, X, Y).NpcIndex
                        
110                     If NpcList(foundNpc).Hostile Then
                        
112                         If NpcList(foundNpc).Target = 0 Then
114                             BuscarNpcEnArea = MapData(.Orig.Map, X, Y).NpcIndex
                                Exit Function
116                         ElseIf UserList(NpcList(foundNpc).Target).flags.NPCAtacado <> foundNpc Then
118                             BuscarNpcEnArea = MapData(.Orig.Map, X, Y).NpcIndex
                                Exit Function
                            End If
                            
                        End If
                        
                    End If
                    
120             Next Y
122         Next X

        End With
        
124     BuscarNpcEnArea = 0
        
        Exit Function

BuscarNpcEnArea:
126     Call TraceError(Err.Number, Err.Description, "Extra.BuscarNpcEnArea", Erl)

        
End Function


Public Sub AI_NpcAtacaNpc(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler
    
        Dim targetPos As t_WorldPos
    
100     With NpcList(NpcIndex)
102         If .TargetNPC > 0 Then
104             targetPos = NpcList(.TargetNPC).Pos
            
106             If InRangoVisionNPC(NpcIndex, targetPos.X, targetPos.Y) Then
                   ' Me fijo si el NPC esta al lado del Objetivo
108                If Distancia(.Pos, targetPos) = 1 And .flags.Paralizado = 0 Then
110                    Call SistemaCombate.NpcAtacaNpc(NpcIndex, .TargetNPC)
                   End If
               
112                If .TargetNPC <> vbNull And .TargetNPC > 0 Then
114                    Call AI_CaminarConRumbo(NpcIndex, targetPos)
                   End If
               
                   Exit Sub
                End If
            End If
           
116         Call RestoreOldMovement(NpcIndex)
 
        End With
                
        Exit Sub
                
ErrorHandler:
118     Call TraceError(Err.Number, Err.Description, "AIv2.AI_NpcAtacaNpc", Erl)

End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
        ' La IA que se ejecuta cuando alguien le pega al maestro de una Mascota/Elemental
        ' o si atacas a los NPCs con Movement = e_TipoAI.NpcDefensa
        ' A diferencia de IrUsuarioCercano(), aca no buscamos objetivos cercanos en el area
        ' porque ya establecemos como objetivo a el usuario que ataco a los NPC con este tipo de IA

        On Error GoTo SeguirAgresor_Err


100     If EsObjetivoValido(NpcIndex, NpcList(NpcIndex).Target) Then
102         Call AI_AtacarUsuarioObjetivo(NpcIndex)
        Else
104         Call RestoreOldMovement(NpcIndex)

        End If

        Exit Sub

SeguirAgresor_Err:
106     Call TraceError(Err.Number, Err.Description, "AI.SeguirAgresor", Erl)


End Sub

Public Sub SeguirAmo(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler
        
100     With NpcList(NpcIndex)
        
102         If .MaestroUser = 0 Or Not .flags.Follow Then Exit Sub
        
            ' Si la mascota no tiene objetivo establecido.
104         If .Target = 0 And .TargetNPC = 0 Then
            
106             If EnRangoVision(NpcIndex, .MaestroUser) Then
108                 If UserList(.MaestroUser).flags.Muerto = 0 And _
                        UserList(.MaestroUser).flags.invisible = 0 And _
                        UserList(.MaestroUser).flags.Oculto = 0 And _
                        Distancia(.Pos, UserList(.MaestroUser).Pos) > 3 Then
                    
                        ' Caminamos cerca del usuario
110                     Call AI_CaminarConRumbo(NpcIndex, UserList(.MaestroUser).Pos)
                        Exit Sub
                    
                    End If
                End If
                
112             Call AI_CaminarSinRumbo(NpcIndex)
            End If
        End With
    
        Exit Sub

ErrorHandler:
114     Call TraceError(Err.Number, Err.Description, "AIv2.SeguirAmo", Erl)

End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

        On Error GoTo RestoreOldMovement_Err

100     With NpcList(NpcIndex)
102         .Target = 0
104         .TargetNPC = 0
        
            ' Si el NPC no tiene maestro, reseteamos el movimiento que tenia antes.
106         If .MaestroUser = 0 Then
108             .Movement = .flags.OldMovement
110             .Hostile = .flags.OldHostil
112             .flags.AttackedBy = vbNullString
            Else
            
                ' Si tiene maestro, hacemos que lo siga.
114             Call FollowAmo(NpcIndex)
            
            End If

        End With

        Exit Sub

RestoreOldMovement_Err:
116     Call TraceError(Err.Number, Err.Description, "AI.RestoreOldMovement", Erl)


End Sub

Private Sub HacerCaminata(ByVal NpcIndex As Integer)
        On Error GoTo Handler
    
        Dim Destino As t_WorldPos
        Dim Heading As e_Heading
        Dim NextTile As t_WorldPos
        Dim MoveChar As Integer
        Dim PudoMover As Boolean

100     With NpcList(NpcIndex)
    
102         Destino.Map = .Pos.Map
104         Destino.X = .Orig.X + .Caminata(.CaminataActual).Offset.X
106         Destino.Y = .Orig.Y + .Caminata(.CaminataActual).Offset.Y

            ' Si todaviï¿½a no llego al destino
108         If .Pos.X <> Destino.X Or .Pos.Y <> Destino.Y Then
        
                ' Tratamos de acercarnos (podemos pisar npcs, usuarios o triggers)
110             Heading = GetHeadingFromWorldPos(.Pos, Destino)
                ' Obtengo la posicion segun el heading
112             NextTile = .Pos
114             Call HeadtoPos(Heading, NextTile)
            
                ' Si hay un NPC
116             MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).NpcIndex
118             If MoveChar Then
                    ' Lo movemos hacia un lado
120                 Call MoveNpcToSide(MoveChar, Heading)
                End If
            
                ' Si hay un user
122             MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).UserIndex
124             If MoveChar Then
                    ' Si no esta muerto o es admin invisible (porque a esos los atraviesa)
126                 If UserList(MoveChar).flags.AdminInvisible = 0 Or UserList(MoveChar).flags.Muerto = 0 Then
                        ' Lo movemos hacia un lado
128                     Call MoveUserToSide(MoveChar, Heading)
                    End If
                End If
            
                ' Movemos al NPC de la caminata
130             PudoMover = MoveNPCChar(NpcIndex, Heading)
            
                ' Si no pudimos moverlo, hacemos como si hubiese llegado a destino... para evitar que se quede atascado
132             If Not PudoMover Or Distancia(.Pos, Destino) = 0 Then
            
                    ' Llegamos a destino, ahora esperamos el tiempo necesario para continuar
134                 .Contadores.IntervaloMovimiento = GetTickCount + .Caminata(.CaminataActual).Espera - .IntervaloMovimiento
                
                    ' Pasamos a la siguiente caminata
136                 .CaminataActual = .CaminataActual + 1
                
                    ' Si pasamos el ultimo, volvemos al primero
138                 If .CaminataActual > UBound(.Caminata) Then
140                     .CaminataActual = 1
                    End If
                
                End If
            
            ' Si por alguna razÃƒÂ³n estamos en el destino, seguimos con la siguiente caminata
            Else
        
142             .CaminataActual = .CaminataActual + 1
            
                ' Si pasamos el ultimo, volvemos al primero
144             If .CaminataActual > UBound(.Caminata) Then
146                 .CaminataActual = 1
                End If
            
            End If
    
        End With
    
        Exit Sub
    
Handler:
148     Call TraceError(Err.Number, Err.Description, "AI.HacerCaminata", Erl)

End Sub

Private Sub MovimientoInvasion(ByVal NpcIndex As Integer)
        On Error GoTo Handler
    
100     With NpcList(NpcIndex)
            Dim SpawnBox As t_SpawnBox
102         SpawnBox = Invasiones(.flags.InvasionIndex).SpawnBoxes(.flags.SpawnBox)
    
            ' Calculamos la distancia a la muralla y generamos una posicion de destino
            Dim DistanciaMuralla As Integer, Destino As t_WorldPos
104         Destino = .Pos
        
106         If SpawnBox.Heading = e_Heading.EAST Or SpawnBox.Heading = e_Heading.WEST Then
108             DistanciaMuralla = Abs(.Pos.X - SpawnBox.CoordMuralla)
110             Destino.X = SpawnBox.CoordMuralla
            Else
112             DistanciaMuralla = Abs(.Pos.Y - SpawnBox.CoordMuralla)
114             Destino.Y = SpawnBox.CoordMuralla
            End If

            ' Si todavia esta lejos de la muralla
116         If DistanciaMuralla > 1 Then
        
                ' Tratamos de acercarnos (sin pisar)
                Dim Heading As e_Heading
118             Heading = GetHeadingFromWorldPos(.Pos, Destino)
            
                ' Nos aseguramos que la posicion nueva esta dentro del rectangulo valido
                Dim NextTile As t_WorldPos
120             NextTile = .Pos
122             Call HeadtoPos(Heading, NextTile)
            
                ' Si la posicion nueva queda fuera del rectangulo valido
124             If Not InsideRectangle(SpawnBox.LegalBox, NextTile.X, NextTile.Y) Then
                    ' Invertimos la direccion de movimiento
126                 Heading = InvertHeading(Heading)
                End If
            
                ' Movemos el NPC
128             Call MoveNPCChar(NpcIndex, Heading)
        
            ' Si esta pegado a la muralla
            Else
        
                ' Chequeamos el intervalo de ataque
130             If Not IntervaloPermiteAtacarNPC(NpcIndex) Then
                    Exit Sub
                End If
            
                ' Nos aseguramos que mire hacia la muralla
132             If .Char.Heading <> SpawnBox.Heading Then
134                 Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, SpawnBox.Heading)
                End If
            
                ' Sonido de ataque (si tiene)
136             If .flags.Snd1 > 0 Then
138                 Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
                End If
            
                ' Sonido de impacto
140             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
                ' Dañamos la muralla
142             Call HacerDañoMuralla(.flags.InvasionIndex, RandomNumber(.Stats.MinHIT, .Stats.MaxHit))  ' TODO: Defensa de la muralla? No hace falta creo...

            End If
    
        End With

        Exit Sub
    
Handler:
        Dim errorDescription As String
144     errorDescription = Err.Description & vbNewLine & "NpcId=" & NpcList(NpcIndex).Numero & " InvasionIndex:" & NpcList(NpcIndex).flags.InvasionIndex & " SpawnBox:" & NpcList(NpcIndex).flags.SpawnBox & vbNewLine
146     Call TraceError(Err.Number, errorDescription, "AI.MovimientoInvasion", Erl)
End Sub

' El NPC elige un hechizo al azar dentro de su listado, con un potencial Target.
' Depdendiendo el tipo de spell que elije, se elije un target distinto que puede ser:
' - El .Target, el NPC mismo o area.
Private Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer)

        On Error GoTo NpcLanzaUnSpell_Err

        ' Elegir hechizo, dependiendo del hechizo lo tiro sobre NPC, sobre Target o Sobre area (cerca de user o NPC si no tiene)
        Dim SpellIndex As Integer
        Dim Target     As Integer
        Dim PuedeDanarAlUsuario As Boolean

100     If Not IntervaloPermiteLanzarHechizo(NpcIndex) Then Exit Sub

102     Target = NpcList(NpcIndex).Target
104     SpellIndex = NpcList(NpcIndex).Spells(RandomNumber(1, NpcList(NpcIndex).flags.LanzaSpells))
106     PuedeDanarAlUsuario = UserList(Target).flags.NoMagiaEfecto = 0 And NpcList(NpcIndex).flags.Paralizado = 0
    
108     Select Case Hechizos(SpellIndex).Target

            Case e_TargetType.uUsuarios

110             If UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
112                 Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)

114                 If UserList(Target).flags.AtacadoPorNpc = 0 Then
116                     UserList(Target).flags.AtacadoPorNpc = NpcIndex

                    End If

                End If

118         Case e_TargetType.uNPC

120             If Hechizos(SpellIndex).AutoLanzar = 1 Then
122                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)

124             ElseIf NpcList(NpcIndex).TargetNPC > 0 Then
126                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcList(NpcIndex).TargetNPC, SpellIndex)

                End If

128         Case e_TargetType.uUsuariosYnpc

130             If Hechizos(SpellIndex).AutoLanzar = 1 Then
132                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)

134             ElseIf UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
136                 Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)

138                 If UserList(Target).flags.AtacadoPorNpc = 0 Then
140                     UserList(Target).flags.AtacadoPorNpc = NpcIndex

                    End If

142             ElseIf NpcList(NpcIndex).TargetNPC > 0 Then
144                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcList(NpcIndex).TargetNPC, SpellIndex)

                End If

146         Case e_TargetType.uTerreno
148             Call NpcLanzaSpellSobreArea(NpcIndex, SpellIndex)

        End Select

        Exit Sub

NpcLanzaUnSpell_Err:
150     Call TraceError(Err.Number, Err.Description, "AI.NpcLanzaUnSpell", Erl)



End Sub

Private Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
        On Error GoTo NpcLanzaUnSpellSobreNpc_Err
    
100     With NpcList(NpcIndex)
        
102         If Not IntervaloPermiteLanzarHechizo(NpcIndex) Then Exit Sub
104         If .Pos.Map <> NpcList(TargetNPC).Pos.Map Then Exit Sub
    
            Dim K As Integer
106             K = RandomNumber(1, .flags.LanzaSpells)

108         Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, .Spells(K))
    
        End With
     
        Exit Sub

NpcLanzaUnSpellSobreNpc_Err:
110     Call TraceError(Err.Number, Err.Description, "AI.NpcLanzaUnSpellSobreNpc", Erl)


End Sub


' ---------------------------------------------------------------------------------------------------
'                                       HELPERS
' ---------------------------------------------------------------------------------------------------

Private Function EsObjetivoValido(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
100     If UserIndex = 0 Then Exit Function

        ' Esta condicion debe ejecutarse independiemente de el modo de busqueda.
102     EsObjetivoValido = ( _
          EnRangoVision(NpcIndex, UserIndex) And _
          EsEnemigo(NpcIndex, UserIndex) And _
          UserList(UserIndex).flags.Muerto = 0 And _
          UserList(UserIndex).flags.EnConsulta = 0 And _
          Not EsGM(UserIndex))

End Function

Private Function EsEnemigo(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

        On Error GoTo EsEnemigo_Err


100     If NpcIndex = 0 Or UserIndex = 0 Then Exit Function

102     With NpcList(NpcIndex)

104         If .flags.AttackedBy <> vbNullString Then
106             EsEnemigo = (UserIndex = NameIndex(.flags.AttackedBy))
108             If EsEnemigo Then Exit Function
            End If

110         Select Case .flags.AIAlineacion
                Case e_Alineacion.Real
112                 EsEnemigo = (Status(UserIndex) Mod 2) <> 1

114             Case e_Alineacion.Caos
116                 EsEnemigo = (Status(UserIndex) Mod 2) <> 0

118             Case e_Alineacion.ninguna
120                 EsEnemigo = True
                    ' Ok. No hay nada especial para hacer, cualquiera puede ser enemigo!

            End Select

        End With

        Exit Function

EsEnemigo_Err:
122     Call TraceError(Err.Number, Err.Description, "AI.EsEnemigo", Erl)


End Function

Private Function EnRangoVision(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

        On Error GoTo EnRangoVision_Err

        Dim userPos As t_WorldPos
        Dim NpcPos As t_WorldPos
        Dim Limite_X As Byte, Limite_Y As Byte

        ' Si alguno es cero, devolve false
100     If NpcIndex = 0 Or UserIndex = 0 Then Exit Function

102     Limite_X = IIf(NpcList(NpcIndex).Distancia <> 0, NpcList(NpcIndex).Distancia, RANGO_VISION_X)
104     Limite_Y = IIf(NpcList(NpcIndex).Distancia <> 0, NpcList(NpcIndex).Distancia, RANGO_VISION_Y)

106     userPos = UserList(UserIndex).Pos
108     NpcPos = NpcList(NpcIndex).Pos

110     EnRangoVision = ( _
          (userPos.Map = NpcPos.Map) And _
          (Abs(userPos.X - NpcPos.X) <= Limite_X) And _
          (Abs(userPos.Y - NpcPos.Y) <= Limite_Y) _
        )


        Exit Function

EnRangoVision_Err:
112     Call TraceError(Err.Number, Err.Description, "AI.EnRangoVision", Erl)


End Function

Private Function UsuarioAtacableConMagia(ByVal targetUserIndex As Integer) As Boolean

        On Error GoTo UsuarioAtacableConMagia_Err

100     If targetUserIndex = 0 Then Exit Function

102     With UserList(targetUserIndex)
104       UsuarioAtacableConMagia = ( _
            .flags.Muerto = 0 And _
            .flags.invisible = 0 And _
            .flags.Inmunidad = 0 And _
            .flags.Oculto = 0 And _
            .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And _
            Not EsGM(targetUserIndex) And _
            Not .flags.EnConsulta)
        End With


        Exit Function

UsuarioAtacableConMagia_Err:
106     Call TraceError(Err.Number, Err.Description, "AI.UsuarioAtacableConMagia", Erl)


End Function

Private Function UsuarioAtacableConMelee(ByVal NpcIndex As Integer, ByVal targetUserIndex As Integer) As Boolean

        On Error GoTo UsuarioAtacableConMelee_Err

100     If targetUserIndex = 0 Then Exit Function

        Dim EstaPegadoAlUser As Boolean
    
102     With UserList(targetUserIndex)
    
104       EstaPegadoAlUser = Distancia(NpcList(NpcIndex).Pos, .Pos) = 1

106       UsuarioAtacableConMelee = ( _
            .flags.Muerto = 0 And _
            .flags.Inmunidad = 0 And _
            (EstaPegadoAlUser Or (Not EstaPegadoAlUser And (.flags.invisible + .flags.Oculto) = 0)) And _
            .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And _
            Not EsGM(targetUserIndex) And _
            Not .flags.EnConsulta)
        End With

        Exit Function

UsuarioAtacableConMelee_Err:
108     Call TraceError(Err.Number, Err.Description, "AI.UsuarioAtacableConMelee", Erl)


End Function

