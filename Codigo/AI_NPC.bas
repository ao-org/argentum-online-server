Attribute VB_Name = "AI"
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

' WyroX: Hardcodeada de la vida...
Public Const FUEGOFATUO      As Integer = 964
Public Const ELEMENTAL_VIENTO      As Integer = 963
Public Const ELEMENTAL_FUEGO      As Integer = 962

'Damos a los NPCs el mismo rango de vison que un PJ
Public Const RANGO_VISION_X  As Byte = 11
Public Const RANGO_VISION_Y  As Byte = 9
Public Sub NpcDummyUpdate(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex)
        Debug.Assert .npcType = DummyTarget
        If .Stats.MinHp < .Stats.MaxHp Then
            .Contadores.UltimoAtaque = .Contadores.UltimoAtaque - 1
            If .Contadores.UltimoAtaque <= 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageTextOverChar(.Stats.MaxHp - .Stats.MinHp, .Char.charindex, vbGreen))
                .Stats.MinHp = .Stats.MaxHp
            End If
        End If
    End With
End Sub

Public Sub NpcAI(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler
        'Debug.Print "NPC: " & NpcList(NpcIndex).Name
100     With NpcList(NpcIndex)
102         Select Case .Movement
                Case e_TipoAI.Estatico
                    ' Es un NPC estatico, no hace nada.
                    Exit Sub

104             Case e_TipoAI.MueveAlAzar
106                 If .Hostile = 1 Then
                        If .AttackRange <= 1 Then
108                         Call PerseguirUsuarioCercano(NpcIndex)
                        Else
                            Call AI_RangeAttack(NpcIndex)
                        End If
                    Else
110                     Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)
                    End If
                Case e_TipoAI.FixedInPos
                    If .Hostile = 1 Then
                        Call AttackFromPos(NpcIndex)
                    End If
112             Case e_TipoAI.NpcDefensa
114                 Call SeguirAgresor(NpcIndex)

116             Case e_TipoAI.NpcAtacaNpc
118                 Call AI_NpcAtacaNpc(NpcIndex)

120             Case e_TipoAI.SigueAmo
122                 Call SeguirAmo(NpcIndex)

                Case e_TipoAI.SupportAndAttack
                    Call AI_SupportAndAttackNpc(NpcIndex)
                
                Case e_TipoAI.BGTankBehavior
                    Call AI_BgTankBehavior(NpcIndex)
                Case e_TipoAI.BGSupportBehavior
                    Call AI_BGSupportBehavior(NpcIndex)
                Case e_TipoAI.BGRangedBehavior
                    Call AI_BGRangedBehavior(NpcIndex)
                Case e_TipoAI.BGBossBehavior
                    Call AI_BGBossBehavior(NpcIndex)
                Case e_TipoAI.BGBossReturnToOrigin
                    Call AI_BGBossReturnToOrigin(NpcIndex)
                    
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
    
136     Call LogError("NPC.AI " & NpcList(NpcIndex).Name & " " & NpcList(NpcIndex).MaestroNPC.ArrayIndex & " mapa:" & NpcList(NpcIndex).Pos.map & " x:" & NpcList(NpcIndex).Pos.X & " y:" & NpcList(NpcIndex).Pos.y & " Mov:" & NpcList(NpcIndex).Movement & " TargU:" & NpcList(NpcIndex).TargetUser.ArrayIndex & " TargN:" & NpcList(NpcIndex).TargetNPC.ArrayIndex)

138     Dim MiNPC As t_Npc: MiNPC = NpcList(NpcIndex)
    
140     Call QuitarNPC(NpcIndex, eAiResetNpc)
142     Call ReSpawnNpc(MiNPC)

End Sub

Private Sub PerseguirUsuarioCercano(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler

        Dim i            As Long
        Dim UserIndex    As Integer
        Dim UserIndexFront As Integer
        Dim npcEraPasivo As Boolean
        Dim agresor      As t_UserReference
        Dim minDistancia As Integer
        Dim minDistanciaAtacable As Integer
        Dim enemigoCercano As Integer
        Dim enemigoAtacableMasCercano As Integer
    
        ' Numero muy grande para que siempre haya un mÃƒÂ­nimo
100     minDistancia = 32000
102     minDistanciaAtacable = 32000

104     With NpcList(NpcIndex)
106         npcEraPasivo = .flags.OldHostil = 0
            If Not IsSet(.flags.StatusMask, eTaunted) Then
108             Call SetUserRef(.targetUser, 0)
110             Call ClearNpcRef(.TargetNPC)


112             If .flags.AttackedBy <> vbNullString Then
114                 agresor = NameIndex(.flags.AttackedBy)
                End If
            
                If UserIndex > 0 And UserIndexFront > 0 Then
                
                    If NPCHasAUserInFront(npcIndex, UserIndexFront) And EsEnemigo(npcIndex, UserIndexFront) Then
                        enemigoAtacableMasCercano = UserIndexFront
                        minDistanciaAtacable = 1
                        minDistancia = 1
                    End If
                Else
                    ' Busco algun objetivo en el area.
116                 For i = 1 To ModAreas.ConnGroups(.pos.map).CountEntrys
118                     UserIndex = ModAreas.ConnGroups(.pos.map).UserEntrys(i)
        
120                     If EsObjetivoValido(npcIndex, UserIndex) Then
                            ' Busco el mas cercano, sea atacable o no.
122                         If Distancia(UserList(UserIndex).pos, .pos) < minDistancia And Not (UserList(UserIndex).flags.invisible > 0 Or UserList(UserIndex).flags.Oculto) Then
124                             enemigoCercano = UserIndex
126                             minDistancia = Distancia(UserList(UserIndex).pos, .pos)
                            End If
                            
                            ' Busco el mas cercano que sea atacable.
128                         If (UsuarioAtacableConMagia(UserIndex) Or UsuarioAtacableConMelee(npcIndex, UserIndex)) And Distancia(UserList(UserIndex).pos, .pos) < minDistanciaAtacable Then
130                             enemigoAtacableMasCercano = UserIndex
132                             minDistanciaAtacable = Distancia(UserList(UserIndex).pos, .pos)
                            End If
        
                        End If
        
134                 Next i
                End If
    
                ' Al terminar el `for`, puedo tener un maximo de tres objetivos distintos.
                ' Por prioridad, vamos a decidir estas cosas en orden.
    
136             If npcEraPasivo Then
                    ' Significa que alguien le pego, y esta en modo agresivo trantando de darle.
                    ' El unico objetivo que importa aca es el atacante; los demas son ignorados.
138                 If EnRangoVision(NpcIndex, agresor.ArrayIndex) Then
                        Call SetUserRef(.TargetUser, agresor.ArrayIndex)
                    Else
                        .flags.AttackedBy = vbNullString
                    End If
    
                Else ' El NPC es hostil siempre, le quiere pegar a alguien.
    
140                 If minDistanciaAtacable > 0 And enemigoAtacableMasCercano > 0 Then ' Hay alguien atacable cerca
142                     Call SetUserRef(.targetUser, enemigoAtacableMasCercano)
144                 ElseIf enemigoCercano > 0 Then ' Hay alguien cerca, pero no es atacable
146                     Call SetUserRef(.targetUser, enemigoCercano)
                    End If
    
                End If
            End If
            ' Si el NPC tiene un objetivo
148         If IsValidUserRef(.TargetUser) Then
                'asignamos heading nuevo al NPC según el Target del nuevo usuario: .Char.Heading, si la distancia es <= 1
                If NPCs.CanMove(.Contadores, .flags) Then
                    Call ChangeNPCChar(npcIndex, .Char.body, .Char.head, GetHeadingFromWorldPos(.pos, UserList(.TargetUser.ArrayIndex).pos))
                End If
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

Private Sub AttackFromPos(ByVal NpcIndex As Integer)

    With NpcList(NpcIndex)
        Dim NearTarget As Integer
        If Not IsSet(.flags.StatusMask, eTaunted) Or Not IsValidUserRef(.targetUser) Then
            Dim Distance As Single
            NearTarget = SelectNearestUser(NpcIndex, Distance)
        Else
            NearTarget = .targetUser.ArrayIndex
        End If
        If NearTarget > 0 Then
            Call SetUserRef(.targetUser, NearTarget)
            Dim TargetHeading As Integer
            TargetHeading = GetHeadingFromWorldPos(.pos, UserList(NearTarget).pos)
            If TargetHeading <> .Char.Heading Then
                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, TargetHeading)
            End If
            If .flags.LanzaSpells And _
                IntervaloPermiteLanzarHechizo(NpcIndex) Then
                If NpcLanzaSpellInmovilizado(NpcIndex, .targetUser.ArrayIndex) Then
                    Call NpcLanzaUnSpell(NpcIndex)
                End If
            ElseIf NPCHasAUserInFront(NpcIndex, NearTarget) Then
                Call NpcAtacaUser(NpcIndex, NearTarget, .Char.Heading)
            End If
        Else
            Call AnimacionIdle(NpcIndex, True)
            If .flags.OldHostil = 0 Then
                Call RestoreOldMovement(NpcIndex)
            End If
        End If
    End With
    
End Sub

Public Function SelectNearestUser(ByVal NpcIndex As Integer, ByRef NearestTargetDistance As Single) As Integer
    Dim i As Integer
    Dim UserIndex As Integer
    NearestTargetDistance = 0
    With NpcList(NpcIndex)
        For i = 1 To ModAreas.ConnGroups(.pos.Map).CountEntrys
            UserIndex = ModAreas.ConnGroups(.pos.Map).UserEntrys(i)
            ' Find nearest user
            If UserMod.IsVisible(UserList(UserIndex)) And UserCanAttackNpc(UserIndex, NpcIndex).CanAttack And _
                NPCs.CanAttackUser(NpcIndex, UserIndex) = e_AttackInteractionResult.eCanAttack Then
                Dim UserDistance As Double
                UserDistance = Distance(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, .pos.x, .pos.y)
                If SelectNearestUser = 0 Or UserDistance < NearestTargetDistance Then
                    SelectNearestUser = UserIndex
                    NearestTargetDistance = UserDistance
                End If
            End If
        Next i
    End With
End Function

Public Function SelectNearestNpc(ByVal NpcIndex, ByRef NearestTargetDistance As Single) As Integer
    Dim i As Integer
    Dim TargetIndex As Integer
    With NpcList(NpcIndex)
        For i = 0 To ModAreas.ConnGroups(.pos.Map).RegisteredNpc - 1
            TargetIndex = ModAreas.ConnGroups(.pos.Map).NpcForAi(i)
            ' Find nearest user
            If NPCs.CanAttackNpc(NpcIndex, TargetIndex) = eCanAttack Then
                Dim NpcDistance As Double
                NpcDistance = Distance(NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y, .pos.x, .pos.y)
                If NearestTargetDistance = 0 Or NpcDistance < NearestTargetDistance Then
                    SelectNearestNpc = TargetIndex
                    NearestTargetDistance = NpcDistance
                End If
            End If
        Next i
    End With
End Function

Public Function SelectCurrentTarget(ByVal NpcIndex, ByVal NearestUser As Integer) As t_AnyReference
    Dim CurrentTarget As t_AnyReference
    With NpcList(NpcIndex)
        If IsSet(.flags.StatusMask, eTaunted) Then
            Exit Function
        End If
        Call ClearUserRef(.targetUser)
        If .flags.AttackedBy <> vbNullString Then
            .targetUser = NameIndex(.flags.AttackedBy)
            If Not IsValidUserRef(.targetUser) Then
                Call ClearUserRef(.targetUser)
            ElseIf Not EnRangoVision(NpcIndex, .targetUser.ArrayIndex) Then
                Call ClearUserRef(.targetUser)
            End If
        End If
        If NearestUser > 0 And Not IsValidUserRef(.targetUser) Then
            Call SetUserRef(.targetUser, NearestUser)
        End If
        If Not CastUserToAnyRef(.targetUser, CurrentTarget) Then
            Call CastNpcToAnyRef(.TargetNPC, CurrentTarget)
        End If
    End With
    SelectCurrentTarget = CurrentTarget
End Function

Public Sub AI_RangeAttack(ByVal NpcIndex As Integer)
    On Error GoTo AI_RangeAttack_Err
    
        Dim CurrentTarget As t_AnyReference
        Dim TargetPos As t_WorldPos
100     With NpcList(NpcIndex)
            Dim NearestUser As Integer
            Dim NearestTargetDistance As Single
            NearestUser = SelectNearestUser(NpcIndex, NearestTargetDistance)
            CurrentTarget = SelectCurrentTarget(NpcIndex, NearestUser)
            'perform attack
128         If IsValidRef(CurrentTarget) And NPCs.CanAttack(.Contadores, .flags) Then
130             TargetPos = GetPosition(CurrentTarget)
132             If Distance(.pos.x, .pos.y, TargetPos.x, TargetPos.y) <= .AttackRange Then
136                 If NpcCanAttack(NpcIndex, CurrentTarget) = eCanAttack And CurrentTarget.RefType = eUser Then
                        If NpcAtacaUser(NpcIndex, CurrentTarget.ArrayIndex, .Char.Heading) And .ProjectileType > 0 Then
                            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, _
                                          PrepareCreateProjectile(.pos.x, .pos.y, TargetPos.x, TargetPos.y, .ProjectileType))
                        End If
                    End If
                Else
                    Call AI_CaminarConRumbo(NpcIndex, TargetPos)
                End If
            End If
            'perform movement
            If NPCs.CanMove(.Contadores, .flags) Then
                If NearestUser > 0 And NearestTargetDistance < .PreferedRange Then
                    Dim Direction As t_Vector
                    Dim TargetMapPos As t_WorldPos
                    Direction = GetDirection(UserList(NearestUser).pos, .pos)
                    TargetMapPos = PreferedTileForDirection(Direction, .pos)
                    Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.pos, TargetMapPos))
                ElseIf Math.Round(NearestTargetDistance) = .PreferedRange Then
                    'do nothing, look at pos?
                ElseIf IsValidRef(CurrentTarget) And Distance(.pos.x, .pos.y, TargetPos.x, TargetPos.y) > .PreferedRange Then
                    Call AI_CaminarConRumbo(NpcIndex, TargetPos)
                ElseIf Distancia(.pos, .Orig) > 0 Then 'return to origin
160                 Call AI_CaminarConRumbo(NpcIndex, .Orig)
                ElseIf .Char.Heading <> e_Heading.SOUTH Then
164                 Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, e_Heading.SOUTH)
                End If
            End If
    End With
    Exit Sub
AI_RangeAttack_Err:
    Call TraceError(Err.Number, Err.Description, "AI_NPC.AI_RangeAttack", Erl)
End Sub
' Cuando un NPC no tiene target y se puede mover libremente pero cerca de su lugar de origen.
' La mayoria de los NPC deberian mantenerse cerca de su posicion de origen, algunos quedaran quietos
' en su posicion y otros se moveran libremente cerca de su posicion de origen.
Private Sub AI_CaminarSinRumboCercaDeOrigen(ByVal NpcIndex As Integer)
        On Error GoTo AI_CaminarSinRumboCercaDeOrigen_Err

100     With NpcList(NpcIndex)
102         If Not NPCs.CanMove(.Contadores, .flags) Then
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

102         If RandomNumber(1, 6) = 3 And NPCs.CanMove(.Contadores, .flags) Then
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
    
100     If Not NPCs.CanMove(NpcList(npcIndex).Contadores, NpcList(npcIndex).flags) Then
102         Call AnimacionIdle(NpcIndex, True)
            Exit Sub
        End If
        If NpcList(NpcIndex).pos.x = rumbo.x And NpcList(NpcIndex).pos.y = rumbo.y Then
            NpcList(NpcIndex).pathFindingInfo.PathLength = 0
            Call AnimacionIdle(NpcIndex, True)
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
Private Function NpcLanzaSpellInmovilizado(ByVal NpcIndex As Integer, ByVal tIndex As Integer) As Boolean
        
    NpcLanzaSpellInmovilizado = False
    
    With NpcList(NpcIndex)
        If Not NPCs.CanMove(.Contadores, .flags) Then
            Select Case .Char.Heading
                Case e_Heading.NORTH
                    If .Pos.X = UserList(tIndex).Pos.X And .Pos.Y > UserList(tIndex).Pos.Y Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
                    
                Case e_Heading.EAST
                    If .Pos.Y = UserList(tIndex).Pos.Y And .Pos.X < UserList(tIndex).Pos.X Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
                
                Case e_Heading.SOUTH
                    If .Pos.X = UserList(tIndex).Pos.X And .Pos.Y < UserList(tIndex).Pos.Y Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
                
                Case e_Heading.WEST
                    If .Pos.Y = UserList(tIndex).Pos.Y And .Pos.X > UserList(tIndex).Pos.X Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
            End Select
        Else
            NpcLanzaSpellInmovilizado = True
        End If
    End With
    
End Function

Public Function ComputeNextHeadingPos(ByVal NpcIndex As Integer) As t_WorldPos
On Error Resume Next
With NpcList(NpcIndex)
    ComputeNextHeadingPos.Map = .Pos.Map
    ComputeNextHeadingPos.X = .Pos.X
    ComputeNextHeadingPos.Y = .Pos.Y
    
    Select Case .Char.Heading
        Case e_Heading.NORTH
            ComputeNextHeadingPos.Y = ComputeNextHeadingPos.Y - 1
        Exit Function
        
        Case e_Heading.SOUTH
            ComputeNextHeadingPos.Y = ComputeNextHeadingPos.Y + 1
        Exit Function
        
        Case e_Heading.EAST
            ComputeNextHeadingPos.X = ComputeNextHeadingPos.X + 1
        Exit Function
        
        Case e_Heading.WEST
            ComputeNextHeadingPos.X = ComputeNextHeadingPos.X - 1
        Exit Function
        
    End Select
End With
End Function

Public Function NPCHasAUserInFront(ByVal NpcIndex As Integer, ByRef UserIndex As Integer) As Boolean
    On Error Resume Next
    Dim NextPosNPC As t_WorldPos
    
    If UserList(UserIndex).flags.Muerto = 1 Then
        NPCHasAUserInFront = False
        Exit Function
    End If
    
    
    
    NextPosNPC = ComputeNextHeadingPos(NpcIndex)
    UserIndex = MapData(NextPosNPC.Map, NextPosNPC.X, NextPosNPC.Y).UserIndex
    NPCHasAUserInFront = (UserIndex > 0)
End Function


Private Sub AI_AtacarUsuarioObjetivo(ByVal AtackerNpcIndex As Integer)
        On Error GoTo ErrorHandler

        Dim AtacaConMagia As Boolean
        Dim AtacaMelee As Boolean
        Dim EstaPegadoAlUsuario As Boolean
        Dim tHeading As Byte
        Dim NextPosNPC As t_WorldPos
        Dim AtacaAlDelFrente As Boolean
        
        AtacaAlDelFrente = False
100     With NpcList(AtackerNpcIndex)
102         If Not IsValidUserRef(.TargetUser) Then Exit Sub
        
104         EstaPegadoAlUsuario = (Distancia(.pos, UserList(.TargetUser.ArrayIndex).pos) <= 1)
106         AtacaConMagia = .flags.LanzaSpells And _
                            IntervaloPermiteLanzarHechizo(AtackerNpcIndex) And _
                            (RandomNumber(1, 100) <= 50)
             
108         AtacaMelee = EstaPegadoAlUsuario And UsuarioAtacableConMelee(AtackerNpcIndex, .targetUser.ArrayIndex)
            AtacaMelee = AtacaMelee And (.flags.LanzaSpells > 0 And _
                        ((UserList(.targetUser.ArrayIndex).flags.invisible > 0 Or UserList(.targetUser.ArrayIndex).flags.Oculto > 0)) Or (IsFeatureEnabled("Magic_and_Punch") And Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eDontHitVisiblePlayers)))
            AtacaMelee = AtacaMelee Or .flags.LanzaSpells = 0
            AtacaMelee = AtacaMelee And NPCs.CanAttack(.Contadores, .flags)
            
            
            ' Se da vuelta y enfrenta al Usuario
109         tHeading = GetHeadingFromWorldPos(.pos, UserList(.TargetUser.ArrayIndex).pos)
            
110         If AtacaConMagia Then
                ' Le lanzo un Hechizo
                If NpcLanzaSpellInmovilizado(AtackerNpcIndex, .TargetUser.ArrayIndex) Then
                    Call ChangeNPCChar(AtackerNpcIndex, .Char.Body, .Char.Head, tHeading)
112                 Call NpcLanzaUnSpell(AtackerNpcIndex)
                End If
114         ElseIf AtacaMelee Then
                Dim ChangeHeading As Boolean
                ChangeHeading = (.flags.Inmovilizado > 0 And tHeading = .Char.Heading) Or NPCs.CanMove(.Contadores, .flags)
                
                Dim UserIndexFront As Integer
                NextPosNPC = ComputeNextHeadingPos(AtackerNpcIndex)
                UserIndexFront = MapData(NextPosNPC.Map, NextPosNPC.X, NextPosNPC.Y).UserIndex
                AtacaAlDelFrente = (UserIndexFront > 0)
                
                If AtacaAlDelFrente And NPCs.CanAttack(.Contadores, .flags) Then
                    Call AnimacionIdle(AtackerNpcIndex, True)
                    If UserIndexFront > 0 Then
                        If UserList(UserIndexFront).flags.Muerto = 0 Then
                            If EsEnemigo(AtackerNpcIndex, UserIndexFront) Then
                                Call NpcAtacaUser(AtackerNpcIndex, UserIndexFront, tHeading)
                            End If
                        End If
                    End If

                End If
            End If

124         If UsuarioAtacableConMagia(.TargetUser.ArrayIndex) Or UsuarioAtacableConMelee(AtackerNpcIndex, .TargetUser.ArrayIndex) Then
                ' Si no tiene un camino pero esta pegado al usuario, no queremos gastar tiempo calculando caminos.
126             If .pathFindingInfo.PathLength = 0 And EstaPegadoAlUsuario Then Exit Sub
            
128             Call AI_CaminarConRumbo(AtackerNpcIndex, UserList(.TargetUser.ArrayIndex).pos)
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
        
102          If IsValidNpcRef(.TargetNPC) Then
104             targetPos = NpcList(.TargetNPC.ArrayIndex).Pos
106             If Distancia(.Pos, targetPos) <= 1 Then
108                 Call SistemaCombate.NpcAtacaNpc(NpcIndex, .TargetNPC.ArrayIndex, False)
                End If
                
110             If DistanciaRadial(.Orig, targetPos) <= (DIAMETRO_VISION_GUARDIAS_NPCS \ 2) Then
112                 If Not IsValidUserRef(NpcList(.TargetNPC.ArrayIndex).TargetUser) Then
114                     Call AI_CaminarConRumbo(NpcIndex, targetPos)
116                 ElseIf UserList(NpcList(.TargetNPC.ArrayIndex).TargetUser.ArrayIndex).flags.NPCAtacado.ArrayIndex <> .TargetNPC.ArrayIndex Then
118                     Call AI_CaminarConRumbo(NpcIndex, targetPos)
                    Else
120                     Call ClearNpcRef(.TargetNPC)
122                     Call AI_CaminarConRumbo(NpcIndex, .Orig)
                    End If
                Else
124                 Call ClearNpcRef(.TargetNPC)
126                 Call AI_CaminarConRumbo(NpcIndex, .Orig)
                End If
            Else
128             Call SetNpcRef(.TargetNPC, BuscarNpcEnArea(NpcIndex))
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

Public Sub AI_SupportAndAttackNpc(ByVal NpcIndex As Integer)
        On Error GoTo ErrorHandler
        
100     With NpcList(NpcIndex)
            If IntervaloPermiteLanzarHechizo(NpcIndex) Then
                Call TrySupportThenAttackSpells(NpcIndex)
            End If
            Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)
        End With
        Exit Sub
ErrorHandler:
136     Call TraceError(Err.Number, Err.Description, "AIv2.AI_SupportAndAttackNpc", Erl)
End Sub

Public Sub AI_BgTankBehavior(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
    With NpcList(NpcIndex)
        Dim CurrentTarget As t_AnyReference
        Dim NearestTarget As Integer
        Dim NearestTargetDistance As Single
        NearestTarget = SelectNearestUser(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eUser)
        NearestTarget = SelectNearestNpc(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eNpc)
        Dim TargetPos As t_WorldPos
        TargetPos = ModReferenceUtils.GetPosition(CurrentTarget)
        If IsValidRef(CurrentTarget) And InRangoVisionNPC(NpcIndex, TargetPos.x, TargetPos.y) Then
            If CurrentTarget.RefType = eUser Then
                Call SetUserRef(.targetUser, CurrentTarget.ArrayIndex)
                AI_AtacarUsuarioObjetivo (NpcIndex)
            Else
                Call SetNpcRef(.TargetNPC, CurrentTarget.ArrayIndex)
                Call AI_NpcAtacaNpc(NpcIndex, False)
            End If
        Else
            Call AI_CaminarConRumbo(NpcIndex, GoToNextWp(NpcIndex))
        End If
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.AI_BgTankBehavior", Erl)
End Sub

Public Sub AI_BGSupportBehavior(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
    With NpcList(NpcIndex)
        If IntervaloPermiteLanzarHechizo(NpcIndex) Then
            Call TrySupportThenAttackSpells(NpcIndex)
        End If
        Dim CurrentTarget As t_AnyReference
        Dim NearestTarget As Integer
        Dim NearestTargetDistance As Single
        NearestTarget = SelectNearestUser(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eUser)
        NearestTarget = SelectNearestNpc(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eNpc)
        Dim TargetPos As t_WorldPos
        TargetPos = ModReferenceUtils.GetPosition(CurrentTarget)
        If NPCs.CanMove(.Contadores, .flags) Then
            If CurrentTarget.ArrayIndex > 0 And NearestTargetDistance < .PreferedRange Then
                Dim Direction As t_Vector
                Dim TargetMapPos As t_WorldPos
                Direction = GetDirection(TargetPos, .pos)
                TargetMapPos = PreferedTileForDirection(Direction, .pos)
                Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.pos, TargetMapPos))
            Else
                If IsValidRef(CurrentTarget) And InRangoVisionNPC(NpcIndex, TargetPos.x, TargetPos.y) Then
                Else
                    Call AI_CaminarConRumbo(NpcIndex, GoToNextWp(NpcIndex))
                End If
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.AI_BGSupportBehavior", Erl)
End Sub

Public Sub AI_BGRangedBehavior(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
    With NpcList(NpcIndex)
        Dim CurrentTarget As t_AnyReference
        Dim NearestTarget As Integer
        Dim NearestTargetDistance As Single
        NearestTarget = SelectNearestUser(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eUser)
        NearestTarget = SelectNearestNpc(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eNpc)
        Dim TargetPos As t_WorldPos
        TargetPos = ModReferenceUtils.GetPosition(CurrentTarget)
        If Not InRangoVisionNPC(NpcIndex, TargetPos.x, TargetPos.y) Then
            Call ClearRef(CurrentTarget)
        End If
        'perform attack
        If CanPerformAttackAction(NpcIndex, .IntervaloAtaque) Then
            If IsValidRef(CurrentTarget) And NPCs.CanAttack(.Contadores, .flags) Then
                If Distance(.pos.x, .pos.y, TargetPos.x, TargetPos.y) <= .AttackRange Then
                    If NpcCanAttack(NpcIndex, CurrentTarget) = eCanAttack Then
                        If CurrentTarget.RefType = eUser Then
                            Call NpcAtacaUser(NpcIndex, CurrentTarget.ArrayIndex, .Char.Heading)
                        Else
                            Call NpcPerformAttackNpc(NpcIndex, CurrentTarget.ArrayIndex)
                        End If
                        If .ProjectileType > 0 Then
                            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, _
                                          PrepareCreateProjectile(.pos.x, .pos.y, TargetPos.x, TargetPos.y, .ProjectileType))
                        End If
                    End If
                ElseIf InRangoVisionNPC(NpcIndex, TargetPos.x, TargetPos.y) Then
                    Call AI_CaminarConRumbo(NpcIndex, TargetPos)
                End If
            End If
        End If
        'perform movement
        If NPCs.CanMove(.Contadores, .flags) Then
            If CurrentTarget.ArrayIndex > 0 And NearestTargetDistance < .PreferedRange Then
                Dim Direction As t_Vector
                Dim TargetMapPos As t_WorldPos
                Direction = GetDirection(TargetPos, .pos)
                TargetMapPos = PreferedTileForDirection(Direction, .pos)
                Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.pos, TargetMapPos))
            ElseIf Not IsValidRef(CurrentTarget) Then
                Call AI_CaminarConRumbo(NpcIndex, GoToNextWp(NpcIndex))
            End If
        End If
        
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.AI_BGRangedBehavior", Erl)
End Sub

Public Sub AI_BGBossBehavior(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
    With NpcList(NpcIndex)
        Dim CurrentTarget As t_AnyReference
        Dim NearestTarget As Integer
        Dim NearestTargetDistance As Single
        NearestTarget = SelectNearestUser(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eUser)
        NearestTarget = SelectNearestNpc(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eNpc)
        Dim TargetPos As t_WorldPos
        TargetPos = ModReferenceUtils.GetPosition(CurrentTarget)
        Dim DistanceFromOrigin As Integer
        DistanceFromOrigin = distance(.Orig.x, .Orig.y, .pos.x, .pos.y)
        If DistanceFromOrigin > 10 Then
            .Movement = BGBossReturnToOrigin
            Call IncreaseSingle(NpcList(NpcIndex).Modifiers.MovementSpeed, 0.5)
            Call UpdateNpcSpeed(NpcIndex)
            Exit Sub
        End If
        If IsValidRef(CurrentTarget) And InRangoVisionNPC(NpcIndex, TargetPos.x, TargetPos.y) Then
            If CurrentTarget.RefType = eUser Then
                Call SetUserRef(.TargetUser, CurrentTarget.ArrayIndex)
                AI_AtacarUsuarioObjetivo (NpcIndex)
            Else
                Call SetNpcRef(.TargetNPC, CurrentTarget.ArrayIndex)
                Call AI_NpcAtacaNpc(NpcIndex, False)
            End If
        Else
            Call AI_CaminarConRumbo(NpcIndex, .Orig)
        End If
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.AI_BGBossBehavior", Erl)
End Sub

Public Sub AI_BGBossReturnToOrigin(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
    With NpcList(NpcIndex)
        Call AI_CaminarConRumbo(NpcIndex, .Orig)
        
        If .pathFindingInfo.PathLength = 0 Then
           .pos = .Orig
           .Movement = BGBossBehavior
           Call IncreaseSingle(NpcList(NpcIndex).Modifiers.MovementSpeed, -0.5)
           Call UpdateNpcSpeed(NpcIndex)
        End If
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.AI_BGBossBehavior", Erl)
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
112                         If Not IsValidUserRef(NpcList(foundNpc).TargetUser) Then
114                             BuscarNpcEnArea = MapData(.Orig.Map, X, Y).NpcIndex
                                Exit Function
116                         ElseIf UserList(NpcList(foundNpc).TargetUser.ArrayIndex).flags.NPCAtacado.ArrayIndex <> foundNpc Then
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


Public Sub AI_NpcAtacaNpc(ByVal NpcIndex As Integer, Optional ByVal ChangeTargetMovement As Boolean = True)
        On Error GoTo ErrorHandler
        Dim targetPos As t_WorldPos
    
100     With NpcList(NpcIndex)
102         If IsValidNpcRef(.TargetNPC) Then
104             targetPos = NpcList(.TargetNPC.ArrayIndex).Pos
            
106             If InRangoVisionNPC(NpcIndex, targetPos.X, targetPos.Y) Then
                   ' Me fijo si el NPC esta al lado del Objetivo
108                If Distancia(.Pos, targetPos) = 1 And NPCs.CanAttack(.Contadores, .flags) Then
110                    Call SistemaCombate.NpcAtacaNpc(NpcIndex, .TargetNPC.ArrayIndex, ChangeTargetMovement)
                   End If
               
112                If IsValidNpcRef(.TargetNPC) Then
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

        
100     If IsValidUserRef(NpcList(npcIndex).TargetUser) And EsObjetivoValido(npcIndex, NpcList(npcIndex).TargetUser.ArrayIndex) Then
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
        
102         If Not IsValidUserRef(.MaestroUser) Or Not .flags.Follow Then Exit Sub
        
            ' Si la mascota no tiene objetivo establecido.
104         If Not IsValidUserRef(.TargetUser) And Not IsValidNpcRef(.TargetNPC) Then
            
106             If EnRangoVision(npcIndex, .MaestroUser.ArrayIndex) Then
108                 If UserList(.MaestroUser.ArrayIndex).flags.Muerto = 0 And _
                        UserList(.MaestroUser.ArrayIndex).flags.invisible = 0 And _
                        UserList(.MaestroUser.ArrayIndex).flags.Oculto = 0 And _
                        Distancia(.pos, UserList(.MaestroUser.ArrayIndex).pos) > 3 Then
                    
                        ' Caminamos cerca del usuario
110                     Call AI_CaminarConRumbo(npcIndex, UserList(.MaestroUser.ArrayIndex).pos)
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
102         Call SetUserRef(.TargetUser, 0)
104         Call ClearNpcRef(.TargetNPC)
        
            ' Si el NPC no tiene maestro, reseteamos el movimiento que tenia antes.
106         If Not IsValidUserRef(.MaestroUser) Then
108             Call SetMovement(NpcIndex, .flags.OldMovement)
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

        If Not IsValidUserRef(NpcList(npcIndex).TargetUser) Then Exit Sub
102     Target = NpcList(npcIndex).TargetUser.ArrayIndex
104     SpellIndex = NpcList(NpcIndex).Spells(RandomNumber(1, NpcList(NpcIndex).flags.LanzaSpells)).SpellIndex
106     PuedeDanarAlUsuario = UserList(Target).flags.NoMagiaEfecto = 0 And NpcList(NpcIndex).flags.Paralizado = 0
        
        If SpellIndex = 0 Then Exit Sub
    
108     Select Case Hechizos(SpellIndex).Target
            Case e_TargetType.uUsuarios
110             If UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
112                 Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)
114                 If Not IsValidNpcRef(UserList(Target).flags.AtacadoPorNpc) Then
116                     Call SetNpcRef(UserList(Target).flags.AtacadoPorNpc, NpcIndex)
                    End If
                End If

118         Case e_TargetType.uNPC
120             If Hechizos(SpellIndex).AutoLanzar = 1 Then
122                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)
124             ElseIf IsValidNpcRef(NpcList(NpcIndex).TargetNPC) Then
126                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcList(NpcIndex).TargetNPC.ArrayIndex, SpellIndex)
                End If
                
128         Case e_TargetType.uUsuariosYnpc
130             If Hechizos(SpellIndex).AutoLanzar = 1 Then
132                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)
134             ElseIf UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
136                 Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)
138                 If Not IsValidNpcRef(UserList(Target).flags.AtacadoPorNpc) Then
140                     Call SetNpcRef(UserList(Target).flags.AtacadoPorNpc, NpcIndex)
                    End If
142             ElseIf IsValidNpcRef(NpcList(NpcIndex).TargetNPC) Then
144                 Call NpcLanzaSpellSobreNpc(NpcIndex, NpcList(NpcIndex).TargetNPC.ArrayIndex, SpellIndex)
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

108         Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, .Spells(K).SpellIndex)
    
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
102     EsObjetivoValido = EnRangoVision(NpcIndex, UserIndex)
        EsObjetivoValido = EsObjetivoValido And EsEnemigo(NpcIndex, UserIndex)
        EsObjetivoValido = EsObjetivoValido And UserList(UserIndex).flags.Muerto = 0
        EsObjetivoValido = EsObjetivoValido And UserList(UserIndex).flags.EnConsulta = 0
        Dim EsAdmin As Boolean: EsAdmin = EsGM(UserIndex) And Not UserList(UserIndex).flags.AdminPerseguible
        EsObjetivoValido = EsObjetivoValido And Not EsAdmin

End Function

Private Function EsEnemigo(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

        On Error GoTo EsEnemigo_Err


100     If NpcIndex = 0 Or UserIndex = 0 Then Exit Function

        EsEnemigo = True

102     With NpcList(NpcIndex)
            ' Si el NPC tiene un atacante
104         If .flags.AttackedBy <> vbNullString Then
                ' Si el usuario actual es el atacante
106             EsEnemigo = (UserIndex = NameIndex(.flags.AttackedBy).ArrayIndex)
108             If EsEnemigo Then Exit Function
                ' Si no es el atacante, preguntamos si el NPC puede atacarlo
109             EsEnemigo = CanAttackNotOwner(NpcIndex, UserIndex)
            End If

110         Select Case .flags.AIAlineacion
                Case e_Alineacion.Real
112                 EsEnemigo = EsEnemigo And (Status(UserIndex) Mod 2) <> 1

114             Case e_Alineacion.Caos
116                 EsEnemigo = EsEnemigo And (Status(UserIndex) Mod 2) <> 0

118             Case e_Alineacion.ninguna
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
            Not (EsGM(targetUserIndex) And Not UserList(targetUserIndex).flags.AdminPerseguible) And _
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
            Not (EsGM(targetUserIndex) And Not UserList(targetUserIndex).flags.AdminPerseguible) And _
            Not .flags.EnConsulta)
        End With

        Exit Function

UsuarioAtacableConMelee_Err:
108     Call TraceError(Err.Number, Err.Description, "AI.UsuarioAtacableConMelee", Erl)


End Function

Private Function CanCastSpell(ByRef npc As t_Npc, ByVal Slot As Integer) As Boolean
     CanCastSpell = GlobalFrameTime - npc.Spells(Slot).LastUse > (npc.Spells(Slot).Cd * 1000)
End Function

Public Function GetAvailableSpellEffects(ByVal NpcIndex As Integer) As Long
    Dim SpellIndex As Integer
    With NpcList(NpcIndex)
        For SpellIndex = 1 To .flags.LanzaSpells
            If CanCastSpell(NpcList(NpcIndex), SpellIndex) Then
                Call SetMask(GetAvailableSpellEffects, Hechizos(.Spells(SpellIndex).SpellIndex).Effects)
            End If
        Next SpellIndex
    End With
End Function

' To help npc they need to be listed in the map for ai with the setting: AddToMapAiList=1
Private Function SelectSupportSpellAndTarget(ByVal NpcIndex As Integer, ByRef Target As t_AnyReference, ByVal AvailableSpellEffect As Long) As Integer
    With NpcList(NpcIndex)
        If Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers Or e_BehaviorFlags.eHelpNpc) Then
            Call ClearRef(Target)
            Exit Function
        End If
        Dim i As Integer
        Dim SpellIndex As Integer
        Dim TargetIndex As Integer
        Dim UserDistance As Double
        If IsSet(AvailableSpellEffect, e_SpellEffects.RemoveParalysis) Then
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers) Then
                For i = 1 To ModAreas.ConnGroups(.pos.Map).CountEntrys
                    TargetIndex = ModAreas.ConnGroups(.pos.Map).UserEntrys(i)
                    ' Find nearest user
                    If CanSeeUser(NpcIndex, TargetIndex) And _
                        NPCs.CanHelpUser(NpcIndex, TargetIndex) = eInteractionOk And _
                        (UserList(TargetIndex).flags.Inmovilizado Or UserList(TargetIndex).flags.Paralizado) Then
                        UserDistance = Distance(UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y, .pos.x, .pos.y)
                        If UserDistance < .SpellRange Then
                            For SpellIndex = 1 To .flags.LanzaSpells
                                If CanCastSpell(NpcList(NpcIndex), SpellIndex) And _
                                   IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.RemoveParalysis) Then
                                    SelectSupportSpellAndTarget = SpellIndex
                                    Call SetRef(Target, TargetIndex, eUser)
                                    Exit Function
                                End If
                            Next SpellIndex
                        End If
                    End If
                Next i
            End If
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eHelpNpc) And Not IsValidRef(Target) Then
                For i = 0 To ModAreas.ConnGroups(.pos.Map).RegisteredNpc - 1
                    TargetIndex = ModAreas.ConnGroups(.pos.Map).NpcForAi(i)
                    ' Find nearest npc
                    If NPCs.CanHelpNpc(NpcIndex, TargetIndex) = eInteractionOk And _
                        (NpcList(TargetIndex).flags.Inmovilizado Or NpcList(TargetIndex).flags.Paralizado) Then
                        UserDistance = Distance(NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y, .pos.x, .pos.y)
                        If UserDistance < .SpellRange Then
                            For SpellIndex = 1 To .flags.LanzaSpells
                                If CanCastSpell(NpcList(NpcIndex), SpellIndex) And _
                                   IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.RemoveParalysis) Then
                                    SelectSupportSpellAndTarget = SpellIndex
                                    Call SetRef(Target, TargetIndex, eNpc)
                                    Exit Function
                                End If
                            Next SpellIndex
                        End If
                    End If
                Next i
            End If
        End If
        
        If IsSet(AvailableSpellEffect, e_SpellEffects.eDoHeal) Then
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers) Then
                For i = 1 To ModAreas.ConnGroups(.pos.Map).CountEntrys
                    TargetIndex = ModAreas.ConnGroups(.pos.Map).UserEntrys(i)
                    ' Find nearest user
                    If CanSeeUser(NpcIndex, TargetIndex) And _
                        UserList(TargetIndex).Stats.MinHp < UserList(TargetIndex).Stats.MaxHp And _
                        NPCs.CanHelpUser(NpcIndex, TargetIndex) = eInteractionOk Then
                        UserDistance = Distance(UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y, .pos.x, .pos.y)
                        If UserDistance < .SpellRange Then
                            For SpellIndex = 1 To .flags.LanzaSpells
                                If CanCastSpell(NpcList(NpcIndex), SpellIndex) And _
                                   IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.eDoHeal) Then
                                    SelectSupportSpellAndTarget = SpellIndex
                                    Call SetRef(Target, TargetIndex, eUser)
                                    Exit Function
                                End If
                            Next SpellIndex
                        End If
                    End If
                Next i
            End If
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eHelpNpc) Then
                For i = 0 To ModAreas.ConnGroups(.pos.Map).RegisteredNpc - 1
                    TargetIndex = ModAreas.ConnGroups(.pos.Map).NpcForAi(i)
                    ' Find nearest user
                    If NpcList(TargetIndex).Stats.MinHp < NpcList(TargetIndex).Stats.MaxHp And _
                        NPCs.CanHelpNpc(NpcIndex, TargetIndex) = eInteractionOk Then
                        UserDistance = Distance(NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y, .pos.x, .pos.y)
                        If UserDistance < .SpellRange Then
                            For SpellIndex = 1 To .flags.LanzaSpells
                                If CanCastSpell(NpcList(NpcIndex), SpellIndex) And _
                                   IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.eDoHeal) Then
                                    SelectSupportSpellAndTarget = SpellIndex
                                    Call SetRef(Target, TargetIndex, eNpc)
                                    Exit Function
                                End If
                            Next SpellIndex
                        End If
                    End If
                Next i
            End If
        End If
    End With
End Function

Private Function SelectAttackSpellAndTarget(ByVal NpcIndex As Integer, ByRef Target As t_AnyReference, ByVal AvailableSpellEffect As Long) As Integer
    With NpcList(NpcIndex)
        If Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers Or e_BehaviorFlags.eAttackNpc) Then
            Exit Function
        End If
        Dim TargetIndex As Integer
        Dim i As Integer
        Dim UserDistance As Double
        Dim SpellIndex As Integer
        If IsSet(AvailableSpellEffect, e_SpellEffects.Paralize Or e_SpellEffects.Immobilize) Then
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers) Then
                If Not IsValidRef(Target) Then
                    For i = 1 To ModAreas.ConnGroups(.pos.Map).CountEntrys
                        TargetIndex = ModAreas.ConnGroups(.pos.Map).UserEntrys(i)
                        ' Find nearest user
                        If CanSeeUser(NpcIndex, TargetIndex) And _
                            NPCs.CanAttackUser(NpcIndex, TargetIndex) = eInteractionOk Then
                            UserDistance = Distance(UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y, .pos.x, .pos.y)
                            If UserDistance < .SpellRange Then
                                Call SetRef(Target, TargetIndex, eUser)
                                Exit For
                            End If
                        End If
                    Next i
                End If
            End If
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackNpc) Then
                If Not IsValidRef(Target) Then
                    For i = 0 To ModAreas.ConnGroups(.pos.Map).RegisteredNpc - 1
                        TargetIndex = ModAreas.ConnGroups(.pos.Map).NpcForAi(i)
                        ' Find nearest user
                        If NPCs.CanAttackNpc(NpcIndex, TargetIndex) = eInteractionOk Then
                            UserDistance = Distance(NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y, .pos.x, .pos.y)
                            If UserDistance < .SpellRange Then
                                Call SetRef(Target, TargetIndex, eNpc)
                                Exit For
                            End If
                        End If
                    Next i
                End If
            End If
            If IsValidRef(Target) Then
                For SpellIndex = 1 To .flags.LanzaSpells
                    If CanCastSpell(NpcList(NpcIndex), SpellIndex) And _
                       IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.Paralize Or e_SpellEffects.Immobilize) Then
                        SelectAttackSpellAndTarget = SpellIndex
                        Exit Function
                    End If
                Next SpellIndex
            End If
        End If
        If IsSet(AvailableSpellEffect, e_SpellEffects.eDoDamage) Then
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers) Then
                If Not IsValidRef(Target) Then
                    For i = 1 To ModAreas.ConnGroups(.pos.Map).CountEntrys
                        TargetIndex = ModAreas.ConnGroups(.pos.Map).UserEntrys(i)
                        ' Find nearest user
                        If CanSeeUser(NpcIndex, TargetIndex) And _
                            NPCs.CanAttackUser(NpcIndex, TargetIndex) = eInteractionOk Then
                            UserDistance = Distance(UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y, .pos.x, .pos.y)
                            If UserDistance < .SpellRange Then
                                Call SetRef(Target, TargetIndex, eUser)
                                Exit For
                            End If
                        End If
                    Next i
                End If
            End If
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackNpc) Then
                If Not IsValidRef(Target) Then
                    For i = 0 To ModAreas.ConnGroups(.pos.Map).RegisteredNpc - 1
                        TargetIndex = ModAreas.ConnGroups(.pos.Map).NpcForAi(i)
                        ' Find nearest user
                        If NPCs.CanAttackNpc(NpcIndex, TargetIndex) = eInteractionOk Then
                            UserDistance = Distance(NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y, .pos.x, .pos.y)
                            If UserDistance < .SpellRange Then
                                Call SetRef(Target, TargetIndex, eNpc)
                                Exit For
                            End If
                        End If
                    Next i
                End If
            End If
            If IsValidRef(Target) Then
                For SpellIndex = 1 To .flags.LanzaSpells
                    If CanCastSpell(NpcList(NpcIndex), SpellIndex) And _
                       IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.eDoDamage) Then
                        SelectAttackSpellAndTarget = SpellIndex
                        Exit Function
                    End If
                Next SpellIndex
            End If
        End If
    End With
End Function

Public Function TryCastHelpSpell(ByVal NpcIndex As Integer, ByVal AvailableSpellEffect As Long) As Boolean
    Dim CurrentTarget As t_AnyReference
    Dim SpellIndex As Integer
    SpellIndex = SelectSupportSpellAndTarget(NpcIndex, CurrentTarget, AvailableSpellEffect)
    If SpellIndex > 0 And IsValidRef(CurrentTarget) Then
        If CurrentTarget.RefType = eUser Then
            Call NpcLanzaSpellSobreUser(NpcIndex, CurrentTarget.ArrayIndex, NpcList(NpcIndex).Spells(SpellIndex).SpellIndex)
        Else
            Call NpcLanzaSpellSobreNpc(NpcIndex, CurrentTarget.ArrayIndex, NpcList(NpcIndex).Spells(SpellIndex).SpellIndex)
        End If
        NpcList(NpcIndex).Spells(SpellIndex).LastUse = GlobalFrameTime
        TryCastHelpSpell = True
    End If
End Function

Public Function TryCastAttackSpell(ByVal NpcIndex As Integer, ByVal AvailableSpellEffect As Long) As Boolean
    Dim CurrentTarget As t_AnyReference
    Dim SpellIndex As Integer
    CurrentTarget = SelectCurrentTarget(NpcIndex, 0)
    SpellIndex = SelectAttackSpellAndTarget(NpcIndex, CurrentTarget, AvailableSpellEffect)
    If SpellIndex > 0 And IsValidRef(CurrentTarget) Then
        If CurrentTarget.RefType = eUser Then
            Call NpcLanzaSpellSobreUser(NpcIndex, CurrentTarget.ArrayIndex, NpcList(NpcIndex).Spells(SpellIndex).SpellIndex)
        Else
            Call NpcLanzaSpellSobreNpc(NpcIndex, CurrentTarget.ArrayIndex, NpcList(NpcIndex).Spells(SpellIndex).SpellIndex)
        End If
        NpcList(NpcIndex).Spells(SpellIndex).LastUse = GlobalFrameTime
        TryCastAttackSpell = True
    End If
End Function

Public Function TrySupportThenAttackSpells(ByVal NpcIndex As Integer) As Boolean
    Dim AvailableSpellEffects As Long
    AvailableSpellEffects = GetAvailableSpellEffects(NpcIndex)
    With NpcList(NpcIndex)
        If IsSet(.flags.StatusMask, eTaunted) Then
            TrySupportThenAttackSpells = TryCastAttackSpell(NpcIndex, AvailableSpellEffects)
            Exit Function
        End If
        TrySupportThenAttackSpells = TryCastHelpSpell(NpcIndex, AvailableSpellEffects)
        If Not TrySupportThenAttackSpells Then
            TrySupportThenAttackSpells = TryCastAttackSpell(NpcIndex, AvailableSpellEffects)
        End If
    End With
End Function

Public Function GoToNextWp(ByVal NpcIndex As Integer) As t_WorldPos
    Dim TargetPos As t_WorldPos
    TargetPos = NpcList(NpcIndex).pos
    Call GetNextWaypointForNpc(NpcIndex, TargetPos.x, TargetPos.y)
    GoToNextWp = TargetPos
End Function

Public Sub SetMovement(ByVal NpcIndex As Integer, ByVal NewMovement As e_TipoAI)
    NpcList(NpcIndex).Movement = NewMovement
    If IsValidUserRef(NpcList(NpcIndex).MaestroUser) Then
        If NewMovement = e_TipoAI.Estatico Or NewMovement = SigueAmo Then
            Call SetBlockTileState(NpcIndex, True)
        Else
            Call SetBlockTileState(NpcIndex, False)
        End If
    End If
End Sub
