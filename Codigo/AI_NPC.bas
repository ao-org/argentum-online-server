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
'  Hardcodeada de la vida...
'Damos a los NPCs el mismo rango de vison que un PJ
Public Const RANGO_VISION_X As Byte = DEFAULT_NPC_VISION_RANGE_X
Public Const RANGO_VISION_Y As Byte = DEFAULT_NPC_VISION_RANGE_Y

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
    With NpcList(NpcIndex)
        Select Case .Movement
            Case e_TipoAI.Estatico
                ' Es un NPC estatico, no hace nada.
                Exit Sub
            Case e_TipoAI.MueveAlAzar
                If .Hostile = 1 Then
                    If .AttackRange <= 1 Then
                        Call PerseguirUsuarioCercano(NpcIndex)
                    Else
                        Call AI_RangeAttack(NpcIndex)
                    End If
                Else
                    Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)
                End If
            Case e_TipoAI.FixedInPos
                If .Hostile = 1 Then
                    Call AttackFromPos(NpcIndex)
                End If
            Case e_TipoAI.NpcDefensa
                Call SeguirAgresor(NpcIndex)
            Case e_TipoAI.NpcAtacaNpc
                Call AI_NpcAtacaNpc(NpcIndex)
            Case e_TipoAI.SigueAmo
                Call SeguirAmo(NpcIndex)
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
            Case e_TipoAI.Caminata
                Call HacerCaminata(NpcIndex)
            Case e_TipoAI.Invasion
                Call MovimientoInvasion(NpcIndex)
            Case e_TipoAI.GuardiaPersigueNpc
                Call AI_GuardiaPersigueNpc(NpcIndex)
        End Select
    End With
    Exit Sub
ErrorHandler:
    Call LogError("NPC.AI " & NpcList(NpcIndex).name & " " & NpcList(NpcIndex).MaestroNPC.ArrayIndex & " mapa:" & NpcList(NpcIndex).pos.Map & " x:" & NpcList(NpcIndex).pos.x & _
            " y:" & NpcList(NpcIndex).pos.y & " Mov:" & NpcList(NpcIndex).Movement & " TargU:" & NpcList(NpcIndex).TargetUser.ArrayIndex & " TargN:" & NpcList( _
            NpcIndex).TargetNPC.ArrayIndex)
    Dim MiNPC As t_Npc: MiNPC = NpcList(NpcIndex)
    Call QuitarNPC(NpcIndex, eAiResetNpc)
    Call ReSpawnNpc(MiNPC)
End Sub

Private Sub PerseguirUsuarioCercano(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim i                         As Long
    Dim UserIndex                 As Integer
    Dim UserIndexFront            As Integer
    Dim npcEraPasivo              As Boolean
    Dim agresor                   As t_UserReference
    Dim minDistancia              As Integer
    Dim minDistanciaAtacable      As Integer
    Dim enemigoCercano            As Integer
    Dim enemigoAtacableMasCercano As Integer
    minDistancia = MAX_INTEGER
    minDistanciaAtacable = MAX_INTEGER
    With NpcList(NpcIndex)
        npcEraPasivo = .flags.OldHostil = 0
        If Not IsSet(.flags.StatusMask, eTaunted) Then
            Call SetUserRef(.TargetUser, 0)
            Call ClearNpcRef(.TargetNPC)
            If .flags.AttackedBy <> vbNullString Then
                agresor = NameIndex(.flags.AttackedBy)
            End If
            If UserIndex > 0 And UserIndexFront > 0 Then
                If NPCHasAUserInFront(NpcIndex, UserIndexFront) And EsEnemigo(NpcIndex, UserIndexFront) Then
                    enemigoAtacableMasCercano = UserIndexFront
                    minDistanciaAtacable = 1
                    minDistancia = 1
                End If
            Else
                ' Busco algun objetivo en el area.
                For i = 1 To ModAreas.ConnGroups(.pos.Map).CountEntrys
                    UserIndex = ModAreas.ConnGroups(.pos.Map).UserEntrys(i)
                    If EsObjetivoValido(NpcIndex, UserIndex) Then
                        ' Busco el mas cercano, sea atacable o no.
                        If Distancia(UserList(UserIndex).pos, .pos) < minDistancia And Not (UserList(UserIndex).flags.invisible > 0 Or UserList(UserIndex).flags.Oculto) Then
                            enemigoCercano = UserIndex
                            minDistancia = Distancia(UserList(UserIndex).pos, .pos)
                        End If
                        ' Busco el mas cercano que sea atacable.
                        If (UsuarioAtacableConMagia(UserIndex) Or UsuarioAtacableConMelee(NpcIndex, UserIndex)) And Distancia(UserList(UserIndex).pos, .pos) < _
                                minDistanciaAtacable Then
                            enemigoAtacableMasCercano = UserIndex
                            minDistanciaAtacable = Distancia(UserList(UserIndex).pos, .pos)
                        End If
                    End If
                Next i
            End If
            ' Al terminar el `for`, puedo tener un maximo de tres objetivos distintos.
            ' Por prioridad, vamos a decidir estas cosas en orden.
            If npcEraPasivo Then
                ' Significa que alguien le pego, y esta en modo agresivo trantando de darle.
                ' El unico objetivo que importa aca es el atacante; los demas son ignorados.
                If EnRangoVision(NpcIndex, agresor.ArrayIndex) Then
                    Call SetUserRef(.TargetUser, agresor.ArrayIndex)
                Else
                    .flags.AttackedBy = vbNullString
                End If
            Else ' El NPC es hostil siempre, le quiere pegar a alguien.
                If minDistanciaAtacable > 0 And enemigoAtacableMasCercano > 0 Then ' Hay alguien atacable cerca
                    Call SetUserRef(.TargetUser, enemigoAtacableMasCercano)
                ElseIf enemigoCercano > 0 Then ' Hay alguien cerca, pero no es atacable
                    Call SetUserRef(.TargetUser, enemigoCercano)
                End If
            End If
        End If
        ' Si el NPC tiene un objetivo
        If IsValidUserRef(.TargetUser) Then
            'asignamos heading nuevo al NPC según el Target del nuevo usuario: .Char.Heading, si la distancia es <= 1
            If NPCs.CanMove(.Contadores, .flags) Then
                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, GetHeadingFromWorldPos(.pos, UserList(.TargetUser.ArrayIndex).pos))
            End If
            Call AI_AtacarUsuarioObjetivo(NpcIndex)
        Else
            If .npcType <> e_NPCType.GuardiaReal And .npcType <> e_NPCType.GuardiasCaos Then
                Call RestoreOldMovement(NpcIndex)
                ' No encontro a nadie cerca, camina unos pasos en cualquier direccion.
                Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)
            Else
                If Distancia(.pos, .Orig) > 0 Then
                    Call AI_CaminarConRumbo(NpcIndex, .Orig)
                Else
                    If .Char.Heading <> e_Heading.SOUTH Then
                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, e_Heading.SOUTH)
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AI_NPC.PerseguirUsuarioCercano", Erl)
End Sub

Private Sub AttackFromPos(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex)
        Dim NearTarget As Integer
        If Not IsSet(.flags.StatusMask, eTaunted) Or Not IsValidUserRef(.TargetUser) Then
            Dim Distance As Single
            NearTarget = SelectNearestUser(NpcIndex, Distance)
        Else
            NearTarget = .TargetUser.ArrayIndex
        End If
        If NearTarget > 0 Then
            Call SetUserRef(.TargetUser, NearTarget)
            Dim TargetHeading As Integer
            TargetHeading = GetHeadingFromWorldPos(.pos, UserList(NearTarget).pos)
            If TargetHeading <> .Char.Heading Then
                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, TargetHeading)
            End If
            If .flags.LanzaSpells And IntervaloPermiteLanzarHechizo(NpcIndex) Then
                If NpcLanzaSpellInmovilizado(NpcIndex, .TargetUser.ArrayIndex) Then
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
    Dim i         As Integer
    Dim UserIndex As Integer
    NearestTargetDistance = 0
    With NpcList(NpcIndex)
        For i = 1 To ModAreas.ConnGroups(.pos.Map).CountEntrys
            UserIndex = ModAreas.ConnGroups(.pos.Map).UserEntrys(i)
            ' Find nearest user
            If UserMod.IsVisible(UserList(UserIndex)) And UserCanAttackNpc(UserIndex, NpcIndex).CanAttack And NPCs.CanAttackUser(NpcIndex, UserIndex) = _
                    e_AttackInteractionResult.eCanAttack Then
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
    Dim i           As Integer
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
        Call ClearUserRef(.TargetUser)
        If .flags.AttackedBy <> vbNullString Then
            .TargetUser = NameIndex(.flags.AttackedBy)
            If Not IsValidUserRef(.TargetUser) Then
                Call ClearUserRef(.TargetUser)
            ElseIf Not EnRangoVision(NpcIndex, .TargetUser.ArrayIndex) Then
                Call ClearUserRef(.TargetUser)
            End If
        End If
        If NearestUser > 0 And Not IsValidUserRef(.TargetUser) Then
            Call SetUserRef(.TargetUser, NearestUser)
        End If
        If Not CastUserToAnyRef(.TargetUser, CurrentTarget) Then
            Call CastNpcToAnyRef(.TargetNPC, CurrentTarget)
        End If
    End With
    SelectCurrentTarget = CurrentTarget
End Function

Public Sub AI_RangeAttack(ByVal NpcIndex As Integer)
    On Error GoTo AI_RangeAttack_Err
    Dim CurrentTarget As t_AnyReference
    Dim TargetPos     As t_WorldPos
    With NpcList(NpcIndex)
        Dim NearestUser           As Integer
        Dim NearestTargetDistance As Single
        NearestUser = SelectNearestUser(NpcIndex, NearestTargetDistance)
        CurrentTarget = SelectCurrentTarget(NpcIndex, NearestUser)
        'perform attack
        If IsValidRef(CurrentTarget) And NPCs.CanAttack(.Contadores, .flags) Then
            TargetPos = GetPosition(CurrentTarget)
            If Distance(.pos.x, .pos.y, TargetPos.x, TargetPos.y) <= .AttackRange Then
                If NpcCanAttack(NpcIndex, CurrentTarget) = eCanAttack And CurrentTarget.RefType = eUser Then
                    If NpcAtacaUser(NpcIndex, CurrentTarget.ArrayIndex, .Char.Heading) And .ProjectileType > 0 Then
                        Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareCreateProjectile(.pos.x, .pos.y, TargetPos.x, TargetPos.y, .ProjectileType))
                    End If
                End If
            Else
                Call AI_CaminarConRumbo(NpcIndex, TargetPos)
            End If
        End If
        'perform movement
        If NPCs.CanMove(.Contadores, .flags) Then
            If NearestUser > 0 And NearestTargetDistance < .PreferedRange Then
                Dim Direction    As t_Vector
                Dim TargetMapPos As t_WorldPos
                Direction = GetDirection(UserList(NearestUser).pos, .pos)
                TargetMapPos = PreferedTileForDirection(Direction, .pos)
                Call MoveNPCChar(NpcIndex, GetHeadingFromWorldPos(.pos, TargetMapPos))
            ElseIf Math.Round(NearestTargetDistance) = .PreferedRange Then
                'do nothing, look at pos?
            ElseIf IsValidRef(CurrentTarget) And Distance(.pos.x, .pos.y, TargetPos.x, TargetPos.y) > .PreferedRange Then
                Call AI_CaminarConRumbo(NpcIndex, TargetPos)
            ElseIf Distancia(.pos, .Orig) > 0 Then 'return to origin
                Call AI_CaminarConRumbo(NpcIndex, .Orig)
            ElseIf .Char.Heading <> e_Heading.SOUTH Then
                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, e_Heading.SOUTH)
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
    With NpcList(NpcIndex)
        If Not NPCs.CanMove(.Contadores, .flags) Then
            Call AnimacionIdle(NpcIndex, True)
        ElseIf Distancia(.pos, .Orig) > 4 Then
            Call AI_CaminarConRumbo(NpcIndex, .Orig)
        ElseIf RandomNumber(1, 6) = 3 Then
            Call MoveNPCChar(NpcIndex, CByte(RandomNumber(e_Heading.NORTH, e_Heading.WEST)))
        Else
            Call AnimacionIdle(NpcIndex, True)
        End If
    End With
    Exit Sub
AI_CaminarSinRumboCercaDeOrigen_Err:
    Call TraceError(Err.Number, Err.Description, "AI.AI_CaminarSinRumboCercaDeOrigen_Err", Erl)
End Sub

' Cuando un NPC no tiene target y se tiene que mover libremente
Private Sub AI_CaminarSinRumbo(ByVal NpcIndex As Integer)
    On Error GoTo AI_CaminarSinRumbo_Err
    With NpcList(NpcIndex)
        If RandomNumber(1, 6) = 3 And NPCs.CanMove(.Contadores, .flags) Then
            Call MoveNPCChar(NpcIndex, CByte(RandomNumber(e_Heading.NORTH, e_Heading.WEST)))
        Else
            Call AnimacionIdle(NpcIndex, True)
        End If
    End With
    Exit Sub
AI_CaminarSinRumbo_Err:
    Call TraceError(Err.Number, Err.Description, "AI.AI_CaminarSinRumbo", Erl)
End Sub

Private Sub AI_CaminarConRumbo(ByVal NpcIndex As Integer, ByRef rumbo As t_WorldPos)
    On Error GoTo AI_CaminarConRumbo_Err
    If NpcList(NpcIndex).TargetUser.ArrayIndex = 0 Then
        Call NpcClearTargetUnreachable(NpcIndex)
    End If
    If Not NPCs.CanMove(NpcList(NpcIndex).Contadores, NpcList(NpcIndex).flags) Then
        Call AnimacionIdle(NpcIndex, True)
        Exit Sub
    End If
    If NpcList(NpcIndex).pos.x = rumbo.x And NpcList(NpcIndex).pos.y = rumbo.y Then
        Call NpcClearTargetUnreachable(NpcIndex)
        NpcList(NpcIndex).pathFindingInfo.PathLength = 0
        Call AnimacionIdle(NpcIndex, True)
        Exit Sub
    End If
    With NpcList(NpcIndex).pathFindingInfo
        ' Si no tiene un camino calculado o si el destino cambio
        If .PathLength = 0 Or .destination.x <> rumbo.x Or .destination.y <> rumbo.y Then
            .destination.x = rumbo.x
            .destination.y = rumbo.y
            ' Recalculamos el camino
            If SeekPath(NpcIndex, True) Then
                ' Si consiguo un camino
                Call NpcClearTargetUnreachable(NpcIndex)
                Call FollowPath(NpcIndex)
            Else
                ' Cannot find path
                If NpcList(NpcIndex).Hostile = 1 And NpcList(NpcIndex).TargetUser.ArrayIndex <> 0 Then
                    NpcList(NpcIndex).pathFindingInfo.RangoVision = Min(SvrConfig.GetValue("NPC_MAX_VISION_RANGE"), NpcList(NpcIndex).pathFindingInfo.RangoVision + _
                            PATH_VISION_DELTA)
                End If
                If NpcList(NpcIndex).TargetUser.ArrayIndex <> 0 And NpcList(NpcIndex).flags.LanzaSpells = 0 Then
                    Call NpcMarkTargetUnreachable(NpcIndex)
                End If
                ' Si no hay camino, pasar a estado idle
                Call AnimacionIdle(NpcIndex, True)
            End If
        Else ' Avanzamos en el camino
            Call NpcClearTargetUnreachable(NpcIndex)
            Call FollowPath(NpcIndex)
        End If
    End With
    Exit Sub
AI_CaminarConRumbo_Err:
    Dim errorDescription As String
    errorDescription = Err.Description & vbNewLine & " NpcIndex: " & NpcIndex & " NPCList.size= " & UBound(NpcList)
    Call TraceError(Err.Number, errorDescription, "AI.AI_CaminarConRumbo", Erl)
End Sub

Private Sub NpcMarkTargetUnreachable(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex)
        If Not .pathFindingInfo.TargetUnreachable Then
            .pathFindingInfo.TargetUnreachable = True
            .pathFindingInfo.PreviousAttackable = .Attackable
            .Attackable = 0
        End If
    End With
End Sub

Private Sub NpcClearTargetUnreachable(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex)
        If .pathFindingInfo.TargetUnreachable Then
            .Attackable = .pathFindingInfo.PreviousAttackable
            .pathFindingInfo.TargetUnreachable = False
        End If
    End With
End Sub

Private Function NpcLanzaSpellInmovilizado(ByVal NpcIndex As Integer, ByVal tIndex As Integer) As Boolean
    NpcLanzaSpellInmovilizado = False
    With NpcList(NpcIndex)
        If Not NPCs.CanMove(.Contadores, .flags) Then
            Select Case .Char.Heading
                Case e_Heading.NORTH
                    If .pos.x = UserList(tIndex).pos.x And .pos.y > UserList(tIndex).pos.y Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
                Case e_Heading.EAST
                    If .pos.y = UserList(tIndex).pos.y And .pos.x < UserList(tIndex).pos.x Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
                Case e_Heading.SOUTH
                    If .pos.x = UserList(tIndex).pos.x And .pos.y < UserList(tIndex).pos.y Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function
                    End If
                Case e_Heading.WEST
                    If .pos.y = UserList(tIndex).pos.y And .pos.x > UserList(tIndex).pos.x Then
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
        ComputeNextHeadingPos.Map = .pos.Map
        ComputeNextHeadingPos.x = .pos.x
        ComputeNextHeadingPos.y = .pos.y
        Select Case .Char.Heading
            Case e_Heading.NORTH
                ComputeNextHeadingPos.y = ComputeNextHeadingPos.y - 1
                Exit Function
            Case e_Heading.SOUTH
                ComputeNextHeadingPos.y = ComputeNextHeadingPos.y + 1
                Exit Function
            Case e_Heading.EAST
                ComputeNextHeadingPos.x = ComputeNextHeadingPos.x + 1
                Exit Function
            Case e_Heading.WEST
                ComputeNextHeadingPos.x = ComputeNextHeadingPos.x - 1
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
    UserIndex = MapData(NextPosNPC.Map, NextPosNPC.x, NextPosNPC.y).UserIndex
    NPCHasAUserInFront = (UserIndex > 0)
End Function

Private Sub AI_AtacarUsuarioObjetivo(ByVal AtackerNpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim AtacaConMagia       As Boolean
    Dim AtacaMelee          As Boolean
    Dim EstaPegadoAlUsuario As Boolean
    Dim tHeading            As Byte
    Dim NextPosNPC          As t_WorldPos
    Dim AtacaAlDelFrente    As Boolean
    AtacaAlDelFrente = False
    With NpcList(AtackerNpcIndex)
        If Not IsValidUserRef(.TargetUser) Then Exit Sub
        EstaPegadoAlUsuario = (Distancia(.pos, UserList(.TargetUser.ArrayIndex).pos) <= 1)
        AtacaConMagia = .flags.LanzaSpells And IntervaloPermiteLanzarHechizo(AtackerNpcIndex)
        AtacaMelee = EstaPegadoAlUsuario And UsuarioAtacableConMelee(AtackerNpcIndex, .TargetUser.ArrayIndex)
        AtacaMelee = AtacaMelee And (.flags.LanzaSpells > 0 And ((UserList(.TargetUser.ArrayIndex).flags.invisible > 0 Or UserList(.TargetUser.ArrayIndex).flags.Oculto > 0)) Or ( _
                IsFeatureEnabled("Magic_and_Punch") And Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eDontHitVisiblePlayers)))
        AtacaMelee = AtacaMelee Or .flags.LanzaSpells = 0
        AtacaMelee = AtacaMelee And NPCs.CanAttack(.Contadores, .flags)
        ' Se da vuelta y enfrenta al Usuario
        tHeading = GetHeadingFromWorldPos(.pos, UserList(.TargetUser.ArrayIndex).pos)
        If AtacaConMagia Then
            ' Le lanzo un Hechizo
            If NpcLanzaSpellInmovilizado(AtackerNpcIndex, .TargetUser.ArrayIndex) Then
                Call ChangeNPCChar(AtackerNpcIndex, .Char.body, .Char.head, tHeading)
                Call NpcLanzaUnSpell(AtackerNpcIndex)
            End If
        ElseIf AtacaMelee Then
            Dim ChangeHeading As Boolean
            ChangeHeading = (.flags.Inmovilizado > 0) Or NPCs.CanMove(.Contadores, .flags)
            If ChangeHeading Then
                NpcList(AtackerNpcIndex).Char.Heading = tHeading
                Call ChangeNPCChar(AtackerNpcIndex, .Char.body, .Char.head, .Char.Heading)
            End If
            Dim UserIndexFront As Integer
            NextPosNPC = ComputeNextHeadingPos(AtackerNpcIndex)
            UserIndexFront = MapData(NextPosNPC.Map, NextPosNPC.x, NextPosNPC.y).UserIndex
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
        If UsuarioAtacableConMagia(.TargetUser.ArrayIndex) Or UsuarioAtacableConMelee(AtackerNpcIndex, .TargetUser.ArrayIndex) Then
            ' Si no tiene un camino pero esta pegado al usuario, no queremos gastar tiempo calculando caminos.
            If .pathFindingInfo.PathLength = 0 And EstaPegadoAlUsuario Then Exit Sub
            Call AI_CaminarConRumbo(AtackerNpcIndex, UserList(.TargetUser.ArrayIndex).pos)
        Else
            Call AI_CaminarSinRumboCercaDeOrigen(AtackerNpcIndex)
        End If
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.AI_AtacarUsuarioObjetivo", Erl)
End Sub

Public Sub AI_GuardiaPersigueNpc(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    Dim TargetPos As t_WorldPos
    With NpcList(NpcIndex)
        If IsValidNpcRef(.TargetNPC) Then
            TargetPos = NpcList(.TargetNPC.ArrayIndex).pos
            If Distancia(.pos, TargetPos) <= 1 Then
                Call SistemaCombate.NpcAtacaNpc(NpcIndex, .TargetNPC.ArrayIndex, False)
            End If
            If DistanciaRadial(.Orig, TargetPos) <= (DIAMETRO_VISION_GUARDIAS_NPCS \ 2) Then
                If Not IsValidUserRef(NpcList(.TargetNPC.ArrayIndex).TargetUser) Then
                    Call AI_CaminarConRumbo(NpcIndex, TargetPos)
                ElseIf UserList(NpcList(.TargetNPC.ArrayIndex).TargetUser.ArrayIndex).flags.NPCAtacado.ArrayIndex <> .TargetNPC.ArrayIndex Then
                    Call AI_CaminarConRumbo(NpcIndex, TargetPos)
                Else
                    Call ClearNpcRef(.TargetNPC)
                    Call AI_CaminarConRumbo(NpcIndex, .Orig)
                End If
            Else
                Call ClearNpcRef(.TargetNPC)
                Call AI_CaminarConRumbo(NpcIndex, .Orig)
            End If
        Else
            Call SetNpcRef(.TargetNPC, BuscarNpcEnArea(NpcIndex))
            If Distancia(.pos, .Orig) > 0 Then
                Call AI_CaminarConRumbo(NpcIndex, .Orig)
            Else
                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, e_Heading.SOUTH)
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.AI_GuardiaAtacaNpc", Erl)
End Sub

Public Sub AI_SupportAndAttackNpc(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    With NpcList(NpcIndex)
        If IntervaloPermiteLanzarHechizo(NpcIndex) Then
            Call TrySupportThenAttackSpells(NpcIndex)
        End If
        Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.AI_SupportAndAttackNpc", Erl)
End Sub

Public Sub AI_BgTankBehavior(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    With NpcList(NpcIndex)
        Dim CurrentTarget         As t_AnyReference
        Dim NearestTarget         As Integer
        Dim NearestTargetDistance As Single
        NearestTarget = SelectNearestUser(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eUser)
        NearestTarget = SelectNearestNpc(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eNpc)
        Dim TargetPos As t_WorldPos
        TargetPos = ModReferenceUtils.GetPosition(CurrentTarget)
        If IsValidRef(CurrentTarget) And InRangoVisionNPC(NpcIndex, TargetPos.x, TargetPos.y) Then
            If CurrentTarget.RefType = eUser Then
                Call SetUserRef(.TargetUser, CurrentTarget.ArrayIndex)
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
        Dim CurrentTarget         As t_AnyReference
        Dim NearestTarget         As Integer
        Dim NearestTargetDistance As Single
        NearestTarget = SelectNearestUser(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eUser)
        NearestTarget = SelectNearestNpc(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eNpc)
        Dim TargetPos As t_WorldPos
        TargetPos = ModReferenceUtils.GetPosition(CurrentTarget)
        If NPCs.CanMove(.Contadores, .flags) Then
            If CurrentTarget.ArrayIndex > 0 And NearestTargetDistance < .PreferedRange Then
                Dim Direction    As t_Vector
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
        Dim CurrentTarget         As t_AnyReference
        Dim NearestTarget         As Integer
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
                            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareCreateProjectile(.pos.x, .pos.y, TargetPos.x, TargetPos.y, .ProjectileType))
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
                Dim Direction    As t_Vector
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
        Dim CurrentTarget         As t_AnyReference
        Dim NearestTarget         As Integer
        Dim NearestTargetDistance As Single
        NearestTarget = SelectNearestUser(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eUser)
        NearestTarget = SelectNearestNpc(NpcIndex, NearestTargetDistance)
        If NearestTarget > 0 Then Call SetRef(CurrentTarget, NearestTarget, eNpc)
        Dim TargetPos As t_WorldPos
        TargetPos = ModReferenceUtils.GetPosition(CurrentTarget)
        Dim DistanceFromOrigin As Integer
        DistanceFromOrigin = Distance(.Orig.x, .Orig.y, .pos.x, .pos.y)
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
    DistanciaRadial = max(Abs(OrigenPos.x - DestinoPos.x), Abs(OrigenPos.y - DestinoPos.y))
End Function

Private Function BuscarNpcEnArea(ByVal NpcIndex As Integer) As Integer
    On Error GoTo BuscarNpcEnArea
    Dim x As Byte, y As Byte
    With NpcList(NpcIndex)
        For x = (.Orig.x - (DIAMETRO_VISION_GUARDIAS_NPCS \ 2)) To (.Orig.x + (DIAMETRO_VISION_GUARDIAS_NPCS \ 2))
            For y = (.Orig.y - (DIAMETRO_VISION_GUARDIAS_NPCS \ 2)) To (.Orig.y + (DIAMETRO_VISION_GUARDIAS_NPCS \ 2))
                If MapData(.Orig.Map, x, y).NpcIndex > 0 And NpcIndex <> MapData(.Orig.Map, x, y).NpcIndex Then
                    Dim foundNpc As Integer
                    foundNpc = MapData(.Orig.Map, x, y).NpcIndex
                    If NpcList(foundNpc).Hostile Then
                        If Not IsValidUserRef(NpcList(foundNpc).TargetUser) Then
                            BuscarNpcEnArea = MapData(.Orig.Map, x, y).NpcIndex
                            Exit Function
                        ElseIf UserList(NpcList(foundNpc).TargetUser.ArrayIndex).flags.NPCAtacado.ArrayIndex <> foundNpc Then
                            BuscarNpcEnArea = MapData(.Orig.Map, x, y).NpcIndex
                            Exit Function
                        End If
                    End If
                End If
            Next y
        Next x
    End With
    BuscarNpcEnArea = 0
    Exit Function
BuscarNpcEnArea:
    Call TraceError(Err.Number, Err.Description, "Extra.BuscarNpcEnArea", Erl)
End Function

Public Sub AI_NpcAtacaNpc(ByVal NpcIndex As Integer, Optional ByVal ChangeTargetMovement As Boolean = True)
    On Error GoTo ErrorHandler
    Dim TargetPos As t_WorldPos
    With NpcList(NpcIndex)
        If IsValidNpcRef(.TargetNPC) Then
            TargetPos = NpcList(.TargetNPC.ArrayIndex).pos
            If InRangoVisionNPC(NpcIndex, TargetPos.x, TargetPos.y) Then
                ' Me fijo si el NPC esta al lado del Objetivo
                If Distancia(.pos, TargetPos) = 1 And NPCs.CanAttack(.Contadores, .flags) Then
                    Call SistemaCombate.NpcAtacaNpc(NpcIndex, .TargetNPC.ArrayIndex, ChangeTargetMovement)
                End If
                If IsValidNpcRef(.TargetNPC) Then
                    Call AI_CaminarConRumbo(NpcIndex, TargetPos)
                End If
                Exit Sub
            End If
        End If
        Call RestoreOldMovement(NpcIndex)
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.AI_NpcAtacaNpc", Erl)
End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
    ' La IA que se ejecuta cuando alguien le pega al maestro de una Mascota/Elemental
    ' o si atacas a los NPCs con Movement = e_TipoAI.NpcDefensa
    ' A diferencia de IrUsuarioCercano(), aca no buscamos objetivos cercanos en el area
    ' porque ya establecemos como objetivo a el usuario que ataco a los NPC con este tipo de IA
    On Error GoTo SeguirAgresor_Err
    If IsValidUserRef(NpcList(NpcIndex).TargetUser) And EsObjetivoValido(NpcIndex, NpcList(NpcIndex).TargetUser.ArrayIndex) Then
        Call AI_AtacarUsuarioObjetivo(NpcIndex)
    Else
        Call RestoreOldMovement(NpcIndex)
    End If
    Exit Sub
SeguirAgresor_Err:
    Call TraceError(Err.Number, Err.Description, "AI.SeguirAgresor", Erl)
End Sub

Public Sub SeguirAmo(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
    With NpcList(NpcIndex)
        If Not IsValidUserRef(.MaestroUser) Or Not .flags.Follow Then Exit Sub
        ' Si la mascota no tiene objetivo establecido.
        If Not IsValidUserRef(.TargetUser) And Not IsValidNpcRef(.TargetNPC) Then
            If EnRangoVision(NpcIndex, .MaestroUser.ArrayIndex) Then
                If UserList(.MaestroUser.ArrayIndex).flags.Muerto = 0 And UserList(.MaestroUser.ArrayIndex).flags.invisible = 0 And UserList( _
                        .MaestroUser.ArrayIndex).flags.Oculto = 0 And Distancia(.pos, UserList(.MaestroUser.ArrayIndex).pos) > 3 Then
                    ' Caminamos cerca del usuario
                    Call AI_CaminarConRumbo(NpcIndex, UserList(.MaestroUser.ArrayIndex).pos)
                    Exit Sub
                End If
            End If
            Call AI_CaminarSinRumbo(NpcIndex)
        End If
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "AIv2.SeguirAmo", Erl)
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
    On Error GoTo RestoreOldMovement_Err
    With NpcList(NpcIndex)
        Call SetUserRef(.TargetUser, 0)
        Call ClearNpcRef(.TargetNPC)
        If .pathFindingInfo.RangoVision <> .pathFindingInfo.OriginalVision Then
            'Restore old range
            .pathFindingInfo.RangoVision = .pathFindingInfo.OriginalVision
        End If
        ' Si el NPC no tiene maestro, reseteamos el movimiento que tenia antes.
        If Not IsValidUserRef(.MaestroUser) Then
            Call SetMovement(NpcIndex, .flags.OldMovement)
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
        Else
            ' Si tiene maestro, hacemos que lo siga.
            Call FollowAmo(NpcIndex)
        End If
    End With
    Exit Sub
RestoreOldMovement_Err:
    Call TraceError(Err.Number, Err.Description, "AI.RestoreOldMovement", Erl)
End Sub

Private Sub HacerCaminata(ByVal NpcIndex As Integer)
    On Error GoTo Handler
    Dim Destino   As t_WorldPos
    Dim Heading   As e_Heading
    Dim NextTile  As t_WorldPos
    Dim MoveChar  As Integer
    Dim PudoMover As Boolean
    With NpcList(NpcIndex)
        Destino.Map = .pos.Map
        Destino.x = .Orig.x + .Caminata(.CaminataActual).offset.x
        Destino.y = .Orig.y + .Caminata(.CaminataActual).offset.y
        ' Si todaviï¿½a no llego al destino
        If .pos.x <> Destino.x Or .pos.y <> Destino.y Then
            ' Tratamos de acercarnos (podemos pisar npcs, usuarios o triggers)
            Heading = GetHeadingFromWorldPos(.pos, Destino)
            ' Obtengo la posicion segun el heading
            NextTile = .pos
            Call HeadtoPos(Heading, NextTile)
            ' Si hay un NPC
            MoveChar = MapData(NextTile.Map, NextTile.x, NextTile.y).NpcIndex
            If MoveChar Then
                ' Lo movemos hacia un lado
                Call MoveNpcToSide(MoveChar, Heading)
            End If
            ' Si hay un user
            MoveChar = MapData(NextTile.Map, NextTile.x, NextTile.y).UserIndex
            If MoveChar Then
                ' Si no esta muerto o es admin invisible (porque a esos los atraviesa)
                If UserList(MoveChar).flags.AdminInvisible = 0 Or UserList(MoveChar).flags.Muerto = 0 Then
                    ' Lo movemos hacia un lado
                    Call MoveUserToSide(MoveChar, Heading)
                End If
            End If
            ' Movemos al NPC de la caminata
            PudoMover = MoveNPCChar(NpcIndex, Heading)
            ' Si no pudimos moverlo, hacemos como si hubiese llegado a destino... para evitar que se quede atascado
            If Not PudoMover Or Distancia(.pos, Destino) = 0 Then
                ' Llegamos a destino, ahora esperamos el tiempo necesario para continuar
                .Contadores.IntervaloMovimiento = AddMod32(GetTickCountRaw(), .Caminata(.CaminataActual).Espera)
                ' Pasamos a la siguiente caminata
                .CaminataActual = .CaminataActual + 1
                ' Si pasamos el ultimo, volvemos al primero
                If .CaminataActual > UBound(.Caminata) Then
                    .CaminataActual = 1
                End If
            End If
            ' Si por alguna razÃƒÂ³n estamos en el destino, seguimos con la siguiente caminata
        Else
            .CaminataActual = .CaminataActual + 1
            ' Si pasamos el ultimo, volvemos al primero
            If .CaminataActual > UBound(.Caminata) Then
                .CaminataActual = 1
            End If
        End If
    End With
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "AI.HacerCaminata", Erl)
End Sub

Private Sub MovimientoInvasion(ByVal NpcIndex As Integer)
    On Error GoTo Handler
    With NpcList(NpcIndex)
        Dim SpawnBox As t_SpawnBox
        SpawnBox = Invasiones(.flags.InvasionIndex).SpawnBoxes(.flags.SpawnBox)
        ' Calculamos la distancia a la muralla y generamos una posicion de destino
        Dim DistanciaMuralla As Integer, Destino As t_WorldPos
        Destino = .pos
        If SpawnBox.Heading = e_Heading.EAST Or SpawnBox.Heading = e_Heading.WEST Then
            DistanciaMuralla = Abs(.pos.x - SpawnBox.CoordMuralla)
            Destino.x = SpawnBox.CoordMuralla
        Else
            DistanciaMuralla = Abs(.pos.y - SpawnBox.CoordMuralla)
            Destino.y = SpawnBox.CoordMuralla
        End If
        ' Si todavia esta lejos de la muralla
        If DistanciaMuralla > 1 Then
            ' Tratamos de acercarnos (sin pisar)
            Dim Heading As e_Heading
            Heading = GetHeadingFromWorldPos(.pos, Destino)
            ' Nos aseguramos que la posicion nueva esta dentro del rectangulo valido
            Dim NextTile As t_WorldPos
            NextTile = .pos
            Call HeadtoPos(Heading, NextTile)
            ' Si la posicion nueva queda fuera del rectangulo valido
            If Not InsideRectangle(SpawnBox.LegalBox, NextTile.x, NextTile.y) Then
                ' Invertimos la direccion de movimiento
                Heading = InvertHeading(Heading)
            End If
            ' Movemos el NPC
            Call MoveNPCChar(NpcIndex, Heading)
            ' Si esta pegado a la muralla
        Else
            ' Chequeamos el intervalo de ataque
            If Not IntervaloPermiteAtacarNPC(NpcIndex) Then
                Exit Sub
            End If
            ' Nos aseguramos que mire hacia la muralla
            If .Char.Heading <> SpawnBox.Heading Then
                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, SpawnBox.Heading)
            End If
            ' Sonido de ataque (si tiene)
            If .flags.Snd1 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .pos.x, .pos.y))
            End If
            ' Sonido de impacto
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_IMPACTO, .pos.x, .pos.y))
            ' Dañamos la muralla
            Call HacerDañoMuralla(.flags.InvasionIndex, RandomNumber(.Stats.MinHIT, .Stats.MaxHit))  ' TODO: Defensa de la muralla? No hace falta creo...
        End If
    End With
    Exit Sub
Handler:
    Dim errorDescription As String
    errorDescription = Err.Description & vbNewLine & "NpcId=" & NpcList(NpcIndex).Numero & " InvasionIndex:" & NpcList(NpcIndex).flags.InvasionIndex & " SpawnBox:" & NpcList( _
            NpcIndex).flags.SpawnBox & vbNewLine
    Call TraceError(Err.Number, errorDescription, "AI.MovimientoInvasion", Erl)
End Sub

' El NPC elige un hechizo al azar dentro de su listado, con un potencial Target.
' Depdendiendo el tipo de spell que elije, se elije un target distinto que puede ser:
' - El .Target, el NPC mismo o area.
Private Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer)
    On Error GoTo NpcLanzaUnSpell_Err
    ' Elegir hechizo, dependiendo del hechizo lo tiro sobre NPC, sobre Target o Sobre area (cerca de user o NPC si no tiene)
    Dim SpellIndex          As Integer
    Dim Target              As Integer
    Dim PuedeDanarAlUsuario As Boolean
    If Not IsValidUserRef(NpcList(NpcIndex).TargetUser) Then Exit Sub
    Target = NpcList(NpcIndex).TargetUser.ArrayIndex
    ' Compute how far the user is from the npcs
    Dim dst_userx As Integer
    Dim dst_usery As Integer
    dst_userx = Abs(NpcList(NpcIndex).pos.x - UserList(Target).pos.x)
    dst_usery = Abs(NpcList(NpcIndex).pos.y - UserList(Target).pos.y)
    Debug.Assert NpcList(NpcIndex).pos.Map = UserList(Target).pos.Map
    If dst_userx > SvrConfig.GetValue("NPC_SPELL_RANGE_X") Then Exit Sub
    If dst_usery > SvrConfig.GetValue("NPC_SPELL_RANGE_Y") Then Exit Sub
    SpellIndex = NpcList(NpcIndex).Spells(RandomNumber(1, NpcList(NpcIndex).flags.LanzaSpells)).SpellIndex
    PuedeDanarAlUsuario = UserList(Target).flags.NoMagiaEfecto = 0 And NpcList(NpcIndex).flags.Paralizado = 0
    If SpellIndex = 0 Then Exit Sub
    Select Case Hechizos(SpellIndex).Target
        Case e_TargetType.uUsuarios
            If UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
                Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)
                If Not IsValidNpcRef(UserList(Target).flags.AtacadoPorNpc) Then
                    Call SetNpcRef(UserList(Target).flags.AtacadoPorNpc, NpcIndex)
                End If
            End If
        Case e_TargetType.uNPC
            If Hechizos(SpellIndex).AutoLanzar = 1 Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)
            ElseIf IsValidNpcRef(NpcList(NpcIndex).TargetNPC) Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, NpcList(NpcIndex).TargetNPC.ArrayIndex, SpellIndex)
            End If
        Case e_TargetType.uUsuariosYnpc
            If Hechizos(SpellIndex).AutoLanzar = 1 Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)
            ElseIf UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
                Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)
                If Not IsValidNpcRef(UserList(Target).flags.AtacadoPorNpc) Then
                    Call SetNpcRef(UserList(Target).flags.AtacadoPorNpc, NpcIndex)
                End If
            ElseIf IsValidNpcRef(NpcList(NpcIndex).TargetNPC) Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, NpcList(NpcIndex).TargetNPC.ArrayIndex, SpellIndex)
            End If
        Case e_TargetType.uTerreno
            Call NpcLanzaSpellSobreArea(NpcIndex, SpellIndex)
    End Select
    Exit Sub
NpcLanzaUnSpell_Err:
    Call TraceError(Err.Number, Err.Description, "AI.NpcLanzaUnSpell", Erl)
End Sub

' ---------------------------------------------------------------------------------------------------
'                                       HELPERS
' ---------------------------------------------------------------------------------------------------
Private Function EsObjetivoValido(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    If UserIndex = 0 Then Exit Function
    ' Esta condicion debe ejecutarse independiemente de el modo de busqueda.
    EsObjetivoValido = EnRangoVision(NpcIndex, UserIndex)
    EsObjetivoValido = EsObjetivoValido And EsEnemigo(NpcIndex, UserIndex)
    EsObjetivoValido = EsObjetivoValido And UserList(UserIndex).flags.Muerto = 0
    EsObjetivoValido = EsObjetivoValido And UserList(UserIndex).flags.EnConsulta = 0
    Dim EsAdmin As Boolean: EsAdmin = EsGM(UserIndex) And Not UserList(UserIndex).flags.AdminPerseguible
    EsObjetivoValido = EsObjetivoValido And Not EsAdmin
End Function

Private Function EsEnemigo(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    On Error GoTo EsEnemigo_Err
    If NpcIndex = 0 Or UserIndex = 0 Then Exit Function
    EsEnemigo = True
    With NpcList(NpcIndex)
        ' Si el NPC tiene un atacante
        If .flags.AttackedBy <> vbNullString Then
            ' Si el usuario actual es el atacante
            EsEnemigo = (UserIndex = NameIndex(.flags.AttackedBy).ArrayIndex)
            If EsEnemigo Then Exit Function
            ' Si no es el atacante, preguntamos si el NPC puede atacarlo
            EsEnemigo = CanAttackNotOwner(NpcIndex, UserIndex)
        End If
        Select Case .flags.AIAlineacion
            Case e_Alineacion.Real
                EsEnemigo = EsEnemigo And (Status(UserIndex) Mod 2) <> 1
            Case e_Alineacion.Caos
                EsEnemigo = EsEnemigo And (Status(UserIndex) Mod 2) <> 0
            Case e_Alineacion.ninguna
                ' Ok. No hay nada especial para hacer, cualquiera puede ser enemigo!
        End Select
    End With
    Exit Function
EsEnemigo_Err:
    Call TraceError(Err.Number, Err.Description, "AI.EsEnemigo", Erl)
End Function

Private Function EnRangoVision(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    On Error GoTo EnRangoVision_Err
    Dim userPos  As t_WorldPos
    Dim NpcPos   As t_WorldPos
    Dim Limite_X As Byte, Limite_Y As Byte
    ' Si alguno es cero, devolve false
    If NpcIndex = 0 Or UserIndex = 0 Then Exit Function
    Limite_X = IIf(NpcList(NpcIndex).Distancia <> 0, NpcList(NpcIndex).Distancia, RANGO_VISION_X)
    Limite_Y = IIf(NpcList(NpcIndex).Distancia <> 0, NpcList(NpcIndex).Distancia, RANGO_VISION_Y)
    userPos = UserList(UserIndex).pos
    NpcPos = NpcList(NpcIndex).pos
    EnRangoVision = ((userPos.Map = NpcPos.Map) And (Abs(userPos.x - NpcPos.x) <= Limite_X) And (Abs(userPos.y - NpcPos.y) <= Limite_Y))
    Exit Function
EnRangoVision_Err:
    Call TraceError(Err.Number, Err.Description, "AI.EnRangoVision", Erl)
End Function

Private Function UsuarioAtacableConMagia(ByVal targetUserIndex As Integer) As Boolean
    On Error GoTo UsuarioAtacableConMagia_Err
    If targetUserIndex = 0 Then Exit Function
    With UserList(targetUserIndex)
        UsuarioAtacableConMagia = (.flags.Muerto = 0 And .flags.invisible = 0 And .flags.Inmunidad = 0 And .flags.Oculto = 0 And .flags.Mimetizado < _
                e_EstadoMimetismo.FormaBichoSinProteccion And Not (EsGM(targetUserIndex) And Not UserList(targetUserIndex).flags.AdminPerseguible) And Not .flags.EnConsulta)
    End With
    Exit Function
UsuarioAtacableConMagia_Err:
    Call TraceError(Err.Number, Err.Description, "AI.UsuarioAtacableConMagia", Erl)
End Function

Private Function UsuarioAtacableConMelee(ByVal NpcIndex As Integer, ByVal targetUserIndex As Integer) As Boolean
    On Error GoTo UsuarioAtacableConMelee_Err
    If targetUserIndex = 0 Then Exit Function
    Dim EstaPegadoAlUser As Boolean
    With UserList(targetUserIndex)
        EstaPegadoAlUser = Distancia(NpcList(NpcIndex).pos, .pos) = 1
        UsuarioAtacableConMelee = (.flags.Muerto = 0 And .flags.Inmunidad = 0 And (EstaPegadoAlUser Or (Not EstaPegadoAlUser And (.flags.invisible + .flags.Oculto) = 0)) And _
                .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And Not (EsGM(targetUserIndex) And Not UserList(targetUserIndex).flags.AdminPerseguible) And Not _
                .flags.EnConsulta)
    End With
    Exit Function
UsuarioAtacableConMelee_Err:
    Call TraceError(Err.Number, Err.Description, "AI.UsuarioAtacableConMelee", Erl)
End Function

Private Function CanCastSpell(ByRef Npc As t_Npc, ByVal Slot As Integer) As Boolean
    CanCastSpell = GlobalFrameTime - Npc.Spells(Slot).lastUse > (Npc.Spells(Slot).Cd * 1000)
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
        Dim i            As Integer
        Dim SpellIndex   As Integer
        Dim TargetIndex  As Integer
        Dim UserDistance As Double
        If IsSet(AvailableSpellEffect, e_SpellEffects.RemoveParalysis) Then
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers) Then
                For i = 1 To ModAreas.ConnGroups(.pos.Map).CountEntrys
                    TargetIndex = ModAreas.ConnGroups(.pos.Map).UserEntrys(i)
                    ' Find nearest user
                    If CanSeeUser(TargetIndex) And NPCs.CanHelpUser(NpcIndex, TargetIndex) = eInteractionOk And (UserList(TargetIndex).flags.Inmovilizado Or UserList( _
                            TargetIndex).flags.Paralizado) Then
                        UserDistance = Distance(UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y, .pos.x, .pos.y)
                        If UserDistance < .SpellRange Then
                            For SpellIndex = 1 To .flags.LanzaSpells
                                If CanCastSpell(NpcList(NpcIndex), SpellIndex) And IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.RemoveParalysis) Then
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
                    If NPCs.CanHelpNpc(NpcIndex, TargetIndex) = eInteractionOk And (NpcList(TargetIndex).flags.Inmovilizado Or NpcList(TargetIndex).flags.Paralizado) Then
                        UserDistance = Distance(NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y, .pos.x, .pos.y)
                        If UserDistance < .SpellRange Then
                            For SpellIndex = 1 To .flags.LanzaSpells
                                If CanCastSpell(NpcList(NpcIndex), SpellIndex) And IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.RemoveParalysis) Then
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
                    If CanSeeUser(TargetIndex) And UserList(TargetIndex).Stats.MinHp < UserList(TargetIndex).Stats.MaxHp And NPCs.CanHelpUser(NpcIndex, TargetIndex) = _
                            eInteractionOk Then
                        UserDistance = Distance(UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y, .pos.x, .pos.y)
                        If UserDistance < .SpellRange Then
                            For SpellIndex = 1 To .flags.LanzaSpells
                                If CanCastSpell(NpcList(NpcIndex), SpellIndex) And IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.eDoHeal) Then
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
                    If NpcList(TargetIndex).Stats.MinHp < NpcList(TargetIndex).Stats.MaxHp And NPCs.CanHelpNpc(NpcIndex, TargetIndex) = eInteractionOk Then
                        UserDistance = Distance(NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y, .pos.x, .pos.y)
                        If UserDistance < .SpellRange Then
                            For SpellIndex = 1 To .flags.LanzaSpells
                                If CanCastSpell(NpcList(NpcIndex), SpellIndex) And IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.eDoHeal) Then
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
        Dim TargetIndex  As Integer
        Dim i            As Integer
        Dim UserDistance As Double
        Dim SpellIndex   As Integer
        If IsSet(AvailableSpellEffect, e_SpellEffects.Paralize Or e_SpellEffects.Immobilize) Then
            If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers) Then
                If Not IsValidRef(Target) Then
                    For i = 1 To ModAreas.ConnGroups(.pos.Map).CountEntrys
                        TargetIndex = ModAreas.ConnGroups(.pos.Map).UserEntrys(i)
                        ' Find nearest user
                        If CanSeeUser(TargetIndex) And NPCs.CanAttackUser(NpcIndex, TargetIndex) = eInteractionOk Then
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
                    If CanCastSpell(NpcList(NpcIndex), SpellIndex) And IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.Paralize Or _
                            e_SpellEffects.Immobilize) Then
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
                        If CanSeeUser(TargetIndex) And NPCs.CanAttackUser(NpcIndex, TargetIndex) = eInteractionOk Then
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
                    If CanCastSpell(NpcList(NpcIndex), SpellIndex) And IsSet(Hechizos(.Spells(SpellIndex).SpellIndex).Effects, e_SpellEffects.eDoDamage) Then
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
    Dim SpellIndex    As Integer
    SpellIndex = SelectSupportSpellAndTarget(NpcIndex, CurrentTarget, AvailableSpellEffect)
    If SpellIndex > 0 And IsValidRef(CurrentTarget) Then
        If CurrentTarget.RefType = eUser Then
            Call NpcLanzaSpellSobreUser(NpcIndex, CurrentTarget.ArrayIndex, NpcList(NpcIndex).Spells(SpellIndex).SpellIndex)
        Else
            Call NpcLanzaSpellSobreNpc(NpcIndex, CurrentTarget.ArrayIndex, NpcList(NpcIndex).Spells(SpellIndex).SpellIndex)
        End If
        NpcList(NpcIndex).Spells(SpellIndex).lastUse = GlobalFrameTime
        TryCastHelpSpell = True
    End If
End Function

Public Function TryCastAttackSpell(ByVal NpcIndex As Integer, ByVal AvailableSpellEffect As Long) As Boolean
    Dim CurrentTarget As t_AnyReference
    Dim SpellIndex    As Integer
    CurrentTarget = SelectCurrentTarget(NpcIndex, 0)
    SpellIndex = SelectAttackSpellAndTarget(NpcIndex, CurrentTarget, AvailableSpellEffect)
    If SpellIndex > 0 And IsValidRef(CurrentTarget) Then
        If CurrentTarget.RefType = eUser Then
            Call NpcLanzaSpellSobreUser(NpcIndex, CurrentTarget.ArrayIndex, NpcList(NpcIndex).Spells(SpellIndex).SpellIndex)
        Else
            Call NpcLanzaSpellSobreNpc(NpcIndex, CurrentTarget.ArrayIndex, NpcList(NpcIndex).Spells(SpellIndex).SpellIndex)
        End If
        NpcList(NpcIndex).Spells(SpellIndex).lastUse = GlobalFrameTime
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
