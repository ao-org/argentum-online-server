VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClanPretoriano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsClanPretoriano - Handles the Praeorians NPCs.
'
' Designed (original version) by Mariano Barrou (El Oso)
' Redesigned by ZaMa
'**************************************************************

'**************************************************************************
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
'**************************************************************************
Option Explicit

'####################
' Pretorians Clan Type (Define wether it is default or custom clan)
Public Enum ePretorianType
    Default = 1
    Custom = 2
End Enum

'####################
' Pretorians Map Vars (Only used for the default map)
Private LeftSpawnPos  As WorldPos
Private RightSpawnPos As WorldPos

'####################
'Actions and spells

' King
Private Enum eKingSpells

    LittleHeal = 1
    RemoveParalisis = 2
    BlindAttack = 3
    StupidityAttack = 4
    HealPoison = 5

End Enum

Private Enum eKingAction

    HealAllyParalisis = 1
    HealAllyPoison
    HealAlly
    ChaseTarget
    MassiveAttack
    RetreatAndHeal

End Enum

' Healer
Private Enum eHealerAction

    HealAllyParalisis = 1
    ParalizePet
    ParalizeUser
    AttackUser
    HealAlly
    RescueFarAlly

End Enum

Private Enum eHealerSpells

    ParalizeUser = 1
    RemoveParalisis = 2
    ParalizeNpc = 3
    Heal = 4
    Tormenta = 5

End Enum

' SpellCaster
Private Enum eSpellCasterAction

    RemoveInvi = 1
    Attack

End Enum

Private Enum eSpellCasterSpells

    Apocalipsis = 1
    RemoInvi = 2

End Enum

' Thief
Private Enum eThiefAction

    Attack = 1
    Steal
    None

End Enum

Private Enum eThiefSpells

    Arrow = 1
    Paralisis = 2

End Enum

' Indicates wether the clan is active or not
Private ClanActive                   As Boolean

' Max distance for the thief to steal
Private Const THIEF_STEAL_DISTANCE   As Integer = 2

' Countdown to break the SpellCaster's wand
Private WandBreakCounter             As Integer

' Number from wich starts countdown
Private Const MAX_WAND_BREAK_VALUE   As Integer = 6

' Min HP needed for SpellCaster to start countdown
Private Const START_COUNTDOWN_MIN_HP As Integer = 750

' Time in which the king hasn't done anything
Private KingIdleTime                 As Long

' Time needed for respawning an ally (Aprox 5 min)
Private Const ALLY_RESPAWN_TIME      As Long = 2000

' Max distance to follow a user, or separate from team
Private Const MAX_DISTANCE           As Integer = 14

' Number of total clan members
Private Const NRO_PRETORIANOS        As Integer = 9

' Indicates clan map and king position
Private CenterPos                    As WorldPos

' Indicates if the clan is spawned after being defeated
Private RespawnClan                  As Boolean

' The number of clan members still alive
Private ClanMembersAlive             As Integer

Private Type tPretorianData

    NpcIndex As Integer
    NPCAI As ePretorianAI

End Type

' Indice del clan pretoriano
Private ClanIndex     As Integer

' Contiene los index de los miembros del clan.
Private Pretorianos() As tPretorianData

Public Function SpawnClan(ByVal Mapa As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer, _
                          ByVal PretoClanIndex As Integer, _
                          Optional ByVal Respawning As Boolean = False) As Boolean

    '********************************************************
    'Author: EL OSO
    'Inicializa el clan Pretoriano.
    'Last Modify Date: 21/09/2010
    '22/06/2006: (Nacho) - Seteamos cuantos NPCs creamos
    '21/09/2010: ZaMa - Reescribi la funcion, ahora se le pasa las coordenadas de respawn
    '********************************************************
    On Error GoTo ErrHandler

    Dim nPos       As WorldPos

    Dim PretoIndex As Integer
    
    ' Initial pos
    CenterPos.Map = Mapa
    CenterPos.X = X
    CenterPos.Y = Y
    
    ReDim Pretorianos(1 To NRO_PRETORIANOS) As tPretorianData
    
    ClanIndex = PretoClanIndex
    
    ' Check valid spawn place only if it's the first time
    If Not Respawning Then

        ' The clan only respawns on pretorian map
        If Mapa = MAPA_PRETORIANO Then
            RespawnClan = True
            Call InitializePretoriansVars
            
            ' Default coordinates, override admin choise
            CenterPos = LeftSpawnPos
            
            ' Check if it's a valid area to respawn
        Else

            If Not IsValidSpawnArea() Then Exit Function

        End If

    End If

    nPos = CenterPos

    ' Spawn clan members
    ClanMembersAlive = NRO_PRETORIANOS
    
    ' Spawn King
    PretoIndex = PretoIndex + 1
    Call SpawnPretorian(nPos, ePretorianAI.King, PretoIndex, True)
    
    ' Spawn 2 Healers
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X + 3
    Call SpawnPretorian(nPos, ePretorianAI.Healer, PretoIndex)
    
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X - 6
    Call SpawnPretorian(nPos, ePretorianAI.Healer, PretoIndex)
    
    ' Spawn 3 Sword Masters
    PretoIndex = PretoIndex + 1
    nPos.Y = nPos.Y + 3
    Call SpawnPretorian(nPos, ePretorianAI.SwordMaster, PretoIndex)
    
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X + 3
    Call SpawnPretorian(nPos, ePretorianAI.SwordMaster, PretoIndex)
    
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X + 3
    Call SpawnPretorian(nPos, ePretorianAI.SwordMaster, PretoIndex)
    
    ' Sapwn 1 Shooter
    PretoIndex = PretoIndex + 1
    nPos.Y = nPos.Y - 6
    nPos.X = nPos.X - 1
    Call SpawnPretorian(nPos, ePretorianAI.Shooter, PretoIndex)
    
    ' Spawn 1 SpellCaster
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X - 4
    Call SpawnPretorian(nPos, ePretorianAI.SpellCaster, PretoIndex)
    
    ' Spawn 1 Thief
    PretoIndex = PretoIndex + 1
    Call SpawnPretorian(nPos, ePretorianAI.Thief, PretoIndex)

    ' Reset wand break counter
    WandBreakCounter = MAX_WAND_BREAK_VALUE
    
    ' Now is active
    ClanActive = True
    
    ' Clan creation success
    SpawnClan = True

    Exit Function

ErrHandler:
    Call LogError("Error en CrearClanPretoriano. Error: " & Err.Number & " - " & Err.description)

End Function

Private Function IsValidSpawnArea() As Boolean
        '********************************************************
        'Author: ZaMa
        'Checks if it's a clean area to respawn the clan.
        'Required conditions:
        '-> A 12x12 square free of blocks and in map bounds.
        '-> Must be a non-water terrain
        'Last Modify Date: 21/09/2010
        '********************************************************
        
        On Error GoTo IsValidSpawnArea_Err
        
    
        Dim LoopX As Long

        Dim LoopY As Long
    
100     For LoopX = CenterPos.X - 6 To CenterPos.X + 6
102         For LoopY = CenterPos.Y - 6 To CenterPos.Y + 6
            
104             If Not InMapBounds(CenterPos.Map, LoopX, LoopY) Then Exit Function
            
106             If (MapData(CenterPos.Map, LoopX, LoopY).Blocked And FLAG_AGUA) <> 0 Then Exit Function
            
108         Next LoopY
110     Next LoopX
    
112     IsValidSpawnArea = True
    
        
        Exit Function

IsValidSpawnArea_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.IsValidSpawnArea", Erl)
        Resume Next
        
End Function

Private Sub InitializePretoriansVars()
        '********************************************************
        'Author: ZaMa
        'Initialize the variables/const used in pretorian map
        'Last Modify Date: 21/09/2010
        '********************************************************
        
        On Error GoTo InitializePretoriansVars_Err
        

100     LeftSpawnPos.Map = MAPA_PRETORIANO
102     LeftSpawnPos.X = 35
104     LeftSpawnPos.Y = 25
    
106     RightSpawnPos.Map = MAPA_PRETORIANO
108     RightSpawnPos.X = 67
110     RightSpawnPos.Y = 25
    
        
        Exit Sub

InitializePretoriansVars_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.InitializePretoriansVars", Erl)
        Resume Next
        
End Sub

Private Sub SpawnPretorian(ByRef SpawnPos As WorldPos, _
                           ByVal PretorianAI As Byte, _
                           ByVal PretoIndex As Integer, _
                           Optional ByVal IsKing As Boolean = False)
        '********************************************************
        'Author: ZaMa
        'Spawns the Pretorian in the given position.
        'Last Modify Date: 21/09/2010
        '********************************************************
        
        On Error GoTo SpawnPretorian_Err
        

        Dim FinalPos       As WorldPos
        Dim NpcIndex       As Integer
        'Dim Head           As Integer
        Dim PretoDatNumber As Integer

100     PretoDatNumber = RandomNumber(PretorianAIOffset(PretorianAI), PretorianAIOffset(PretorianAI + 1) - 1)
                            
        ' Odd numbers are tall races
102     If (PretoDatNumber And 1) = 1 Then
104         'Head = RandomTallHead
        Else
106         'Head = RandomShortHead
        End If
                      
        ' The King has to spawn always in the given position
108     If IsKing Then
        
            ' Check if there's any npc in the king's position
            Dim OtherNpcIndex As Integer

110         OtherNpcIndex = MapData(SpawnPos.Map, SpawnPos.X, SpawnPos.Y).NpcIndex
        
            ' Found one
112         If OtherNpcIndex <> 0 Then
                ' Move it to the closest legal position
114             Call ClosestLegalPos(SpawnPos, FinalPos)

116             If FinalPos.X <> 0 And FinalPos.Y <> 0 Then
118                 Call SendData(SendTarget.ToNPCArea, OtherNpcIndex, PrepareMessageCharacterMove(Npclist(OtherNpcIndex).Char.CharIndex, FinalPos.X, FinalPos.Y))
                    ' Update map and npc pos
120                 MapData(SpawnPos.Map, SpawnPos.X, SpawnPos.Y).NpcIndex = 0
122                 Npclist(OtherNpcIndex).Pos = FinalPos
124                 MapData(FinalPos.Map, FinalPos.X, FinalPos.Y).NpcIndex = OtherNpcIndex
            
                    ' King's position is now empty
126                 FinalPos = SpawnPos
            
                    ' Couldn't find a legal pos => Remove it
                Else
128                 Call QuitarNPC(OtherNpcIndex)

                End If

            Else
130             FinalPos = SpawnPos

            End If
        
            ' Other clan member doesn't have to spawn in the exact position
        Else
132         Call ClosestLegalPos(SpawnPos, FinalPos, False, True)

        End If
                            
        ' Create and Spawn npc
134     NpcIndex = CrearNPC(PretorianDatNumbers(PretoDatNumber), FinalPos.Map, FinalPos)
        'NpcIndex = CrearNPC(PretorianDatNumbers(PretoDatNumber), FinalPos.Map, FinalPos, Head)
    
        ' Store Index
136     Pretorianos(PretoIndex).NpcIndex = NpcIndex
138     Pretorianos(PretoIndex).NPCAI = PretorianAI
    
        ' Asign ClanIndex
140     Npclist(NpcIndex).ClanIndex = ClanIndex
    
        
        Exit Sub

SpawnPretorian_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.SpawnPretorian", Erl)
        Resume Next
        
End Sub

Private Function RandomTallHead()
        '********************************************************
        'Author: ZaMa
        'Gets a random tall race's Head.
        'Last Modify Date: 21/09/2010
        '********************************************************
        
        On Error GoTo RandomTallHead_Err
        
    
        ' 3 tall races
        Dim raza As Integer

100     raza = RandomNumber(1, 3)
        
        ' Male
102     If (RandomNumber(0, 1) = 1) Then
        
            ' Human
104         If raza = 1 Then
106             'RandomTallHead = RandomNumber(HUMANO_H_PRIMER_CABEZA, HUMANO_H_ULTIMA_CABEZA)
                ' Elf
108         ElseIf raza = 2 Then
110             'RandomTallHead = RandomNumber(ELFO_H_PRIMER_CABEZA, ELFO_H_ULTIMA_CABEZA)
                ' Drow
            Else
112             'RandomTallHead = RandomNumber(DROW_H_PRIMER_CABEZA, DROW_H_ULTIMA_CABEZA)

            End If
        
            ' Female
        Else
        
            ' Human
114         If raza = 1 Then
116             'RandomTallHead = RandomNumber(HUMANO_M_PRIMER_CABEZA, HUMANO_M_ULTIMA_CABEZA)
                ' Elf
118         ElseIf raza = 2 Then
120             'RandomTallHead = RandomNumber(ELFO_M_PRIMER_CABEZA, ELFO_M_ULTIMA_CABEZA)
                ' Drow
            Else
122             'RandomTallHead = RandomNumber(DROW_M_PRIMER_CABEZA, DROW_M_ULTIMA_CABEZA)

            End If
        
        End If
    
        
        Exit Function

RandomTallHead_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.RandomTallHead", Erl)
        Resume Next
        
End Function

Private Function RandomShortHead()
        '********************************************************
        'Author: ZaMa
        'Gets a random short race's Head.
        'Last Modify Date: 21/09/2010
        '********************************************************
        
        On Error GoTo RandomShortHead_Err
        

        ' 2 short races
        Dim raza As Integer

100     raza = RandomNumber(1, 2)
        
        ' Male
102     If (RandomNumber(0, 1) = 1) Then
        
            ' Dwarf
104         If raza = 1 Then
106             'RandomShortHead = RandomNumber(ENANO_H_PRIMER_CABEZA, ENANO_H_ULTIMA_CABEZA)
                ' Gnome
            Else
108             'RandomShortHead = RandomNumber(GNOMO_H_PRIMER_CABEZA, GNOMO_H_ULTIMA_CABEZA)

            End If
        
            ' Female
        Else
        
            ' Dwarf
110         If raza = 1 Then
112             'RandomShortHead = RandomNumber(ENANO_M_PRIMER_CABEZA, ENANO_M_ULTIMA_CABEZA)
                ' Gnome
            Else
114             'RandomShortHead = RandomNumber(GNOMO_M_PRIMER_CABEZA, GNOMO_M_ULTIMA_CABEZA)

            End If
        
        End If

        
        Exit Function

RandomShortHead_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.RandomShortHead", Erl)
        Resume Next
        
End Function

Public Sub PerformPretorianAI(ByVal NpcIndex As Integer)
        '********************************************************
        'Author: ZaMa
        'Performs Pretorian's AI.
        'Last Modify Date: 21/09/2010
        '********************************************************
        
        On Error GoTo PerformPretorianAI_Err

        Dim PretorianAI As Byte

100     PretorianAI = GetPretorianAI(NpcIndex)
    
102     Select Case PretorianAI
        
            Case ePretorianAI.King
104             Call AI_King(NpcIndex)
            
106         Case ePretorianAI.Healer
108             Call AI_Healer(NpcIndex)
            
110         Case ePretorianAI.SpellCaster
112             Call AI_SpellCaster(NpcIndex)
            
114         Case ePretorianAI.SwordMaster
116             Call AI_SwordMaster(NpcIndex)
            
118         Case ePretorianAI.Shooter
120             Call AI_Shooter(NpcIndex)
            
122         Case ePretorianAI.Thief
124             Call AI_Thief(NpcIndex)
            
        End Select
    
        
        Exit Sub

PerformPretorianAI_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.PerformPretorianAI", Erl)
        Resume Next
        
End Sub

Private Function GetPretorianAI(ByVal NpcIndex As Integer) As Byte
        '********************************************************
        'Author: ZaMa
        'Returns the Pretorian's AI for the given Npc.
        'Last Modify Date: 21/09/2010
        '********************************************************
        
        On Error GoTo GetPretorianAI_Err
        
    
        Dim counter As Long
    
100     For counter = 1 To NRO_PRETORIANOS

102         If Pretorianos(counter).NpcIndex = NpcIndex Then
104             GetPretorianAI = Pretorianos(counter).NPCAI
                Exit Function

            End If

106     Next counter
    
        
        Exit Function

GetPretorianAI_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.GetPretorianAI", Erl)
        Resume Next
        
End Function

Public Sub AI_King(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    '***************************************************
    On Error GoTo ErrHandler

    Dim BestTarget As Integer

    Dim Action     As Byte
    
    BestTarget = KingBestTarget(NpcIndex, Action)
    
    Call KingPerformAction(NpcIndex, BestTarget, Action)
    
    Exit Sub

ErrHandler:
    LogError ("Error en AI_King. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function KingBestTarget(ByVal NpcIndex As Integer, _
                                ByRef Accion As Byte) As Integer

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Pick the best target according to the following criteria:
    '1) First try to heal status effects on allies (paralisis, poison, damaged) if any is alive
    '2) If no ally is alive, then start a melee atack to near users
    '3) If a user tries to run, then blind and make him stupid
    '4) if no one is near, then heal himself.
    '***************************************************
    On Error GoTo ErrHandler
    
    Dim BestTarget As Integer

    ' There should be more than one in order to support
    If ClanMembersAlive <> 1 Then

        ' Chooses target only if can attack
        If IntervaloPermiteAtacarNPC(NpcIndex) Then
            ' First checks if any ally is paralized in order to remove its paralizis
            BestTarget = AllyParalyzed(NpcIndex, False)

            If BestTarget <> 0 Then

                ' 20% of probability of doing it
                If RandomNumber(1, 100) < 21 Then
                    Accion = eKingAction.HealAllyParalisis

                End If

            Else
                ' Now checks if any ally is poisoned in order to remove its poison
                BestTarget = AllyPoisoned(NpcIndex)

                If BestTarget <> 0 Then
                    Accion = eKingAction.HealAllyPoison
                Else
                    ' If no target found, then try to heal injured allies
                    BestTarget = AllyInjured(NpcIndex)

                    If BestTarget <> 0 Then
                        Accion = eKingAction.HealAlly

                    End If

                End If

            End If

        End If
    
        ' King's alone
    Else
        
        Dim Mapa               As Integer

        Dim NPCPosX            As Integer

        Dim NPCPosY            As Integer
        
        Dim UserIndex          As Integer

        Dim counter            As Long
        
        Dim BestTargetDistance As Integer

        Dim Distance           As Integer
        
        With Npclist(NpcIndex).Pos
            Mapa = .Map
            NPCPosX = .X
            NPCPosY = .Y

        End With
        
        Dim CounterStart As Long

        Dim CounterEnd   As Long

        Dim CounterStep  As Long
        
        ' To avoid that all attack the same target
        CounterStep = RandomNumber(0, 1)

        If CounterStep = 1 Then
            CounterStart = 1
            CounterEnd = ModAreas.ConnGroups(Mapa).CountEntrys
        Else
            CounterStart = ModAreas.ConnGroups(Mapa).CountEntrys
            CounterEnd = 1
            CounterStep = -1

        End If
        
        ' Search for the best user target
        For counter = CounterStart To CounterEnd Step CounterStep
        
            UserIndex = ModAreas.ConnGroups(Mapa).UserEntrys(counter)
            
            'Is it in it's range of vision??
            If InVisionRange(UserIndex, NPCPosX, NPCPosY) Then

                ' Can be atacked? If it's blinded, doesn't count.
                If UserAtacable(UserIndex) And UserList(UserIndex).flags.Ceguera = 0 Then

                    ' if previous user choosen, select the better
                    If BestTarget <> 0 Then
                        ' Has priority to attack the nearest
                        Distance = UserDistance(UserIndex, NPCPosX, NPCPosY)
                        
                        If Distance < BestTargetDistance Then
                            BestTarget = UserIndex
                            BestTargetDistance = Distance

                        End If

                    Else
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NPCPosX, NPCPosY)

                    End If

                End If

            End If
                
        Next counter
        
        ' Any target found?
        If BestTarget <> 0 Then

            ' Is reachable?
            If UserReachable(NpcIndex, BestTarget) Then
                ' Chase it
                Accion = eKingAction.ChaseTarget
            
                ' The target is trying to escape => Stop him with a massive status attack
            Else
                Accion = eKingAction.MassiveAttack

            End If
            
            ' No target found => Retreat and heal
        Else
            Accion = eKingAction.RetreatAndHeal

        End If

    End If

    KingBestTarget = BestTarget

    Exit Function

ErrHandler:
    LogError ("Error en KingBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub KingPerformAction(ByVal NpcIndex As Integer, _
                              ByVal BestTarget As Integer, _
                              ByVal Accion As Byte)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Performs the required action.
    '***************************************************
    On Error GoTo ErrHandler

    Dim IdleTime As Long

    With Npclist(NpcIndex)

        Select Case Accion
            
            Case eKingAction.HealAllyParalisis
                Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eKingSpells.RemoveParalisis))
            
            Case eKingAction.HealAllyPoison
                Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eKingSpells.HealPoison))
            
            Case eKingAction.HealAlly
                Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eKingSpells.LittleHeal))
                
                ' If time for respawning an ally has passed, then try to do it
                If KingIdleTime > ALLY_RESPAWN_TIME Then Call KingReviveAlly(NpcIndex)
                IdleTime = KingIdleTime + 1
                
            Case eKingAction.ChaseTarget, eKingAction.MassiveAttack
                
                If Accion = eKingAction.ChaseTarget Then
                    Call GreedyWalkTo(NpcIndex, UserList(BestTarget).Pos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
                Else
                    Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eKingSpells.StupidityAttack), True)
                    Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eKingSpells.BlindAttack), True)
                    
                    ' King Message
                    Call WriteConsoleMsg(BestTarget, "El rey pretoriano te ha vuelto estupido.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(BestTarget, "El rey pretoriano te ha vuelto ciego ", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(BestTarget, "A la distancia escuchas las siguientes palabras: Cobarde, no eres digno de luchar conmigo si escapas! ", FontTypeNames.FONTTYPE_VENENO)

                End If
            
                Call KingMeleeAttack(NpcIndex)
            
            Case eKingAction.RetreatAndHeal
                Call ReturnToCenter(NpcIndex)

                If .Stats.MinHp <> .Stats.MaxHp Then
                    Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, .Spells(eKingSpells.LittleHeal))
                End If
                
                ' If time for respawning an ally has passed, then try to do it
                If KingIdleTime > ALLY_RESPAWN_TIME Then Call KingReviveAlly(NpcIndex)
                IdleTime = KingIdleTime + 1
                
        End Select
    
    End With
    
    ' Resets if anyone is around, else increases
    KingIdleTime = IdleTime
    Exit Sub

ErrHandler:
    LogError ("Error en KingPerformAction. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub KingMeleeAttack(NpcIndex)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    'King performes a melee attack ignoring attack intervals
    '***************************************************
    On Error GoTo ErrHandler

    Dim headingloop As Byte

    Dim nPos        As WorldPos
    
    Dim UserIndex   As Integer
    
    With Npclist(NpcIndex)
    
        ' Check the four directions
        For headingloop = eHeading.NORTH To eHeading.WEST
        
            nPos = .Pos
            
            Call HeadtoPos(headingloop, nPos)
            
            ' In map Limits?
            If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
            
                UserIndex = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex

                If UserIndex > 0 Then
                    If UserAtacable(UserIndex, False) Then
                        If NpcAtacaUser(NpcIndex, UserIndex, headingloop) Then
                            Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                        End If

                    End If

                End If

            End If

        Next headingloop
        
    End With
    
    Exit Sub

ErrHandler:
    LogError ("Error en KingMeleeAttack. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub KingReviveAlly(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 26/09/2010
        'King tries to revive an ally.
        '***************************************************
        
        On Error GoTo KingReviveAlly_Err
        
    
        Dim AllyIndex As Integer

        Dim nPos      As WorldPos
    
100     With Npclist(NpcIndex)

            ' Has to be in the center to revive
102         If .Pos.X = CenterPos.X Then
104             If .Pos.Y = CenterPos.Y Then
                
                    ' Any ally dead?
106                 AllyIndex = AllyDead()

108                 If AllyIndex <> 0 Then
110                     nPos = CenterPos
112                     nPos.X = nPos.X + 1
114                     Call SpawnPretorian(nPos, Pretorianos(AllyIndex).NPCAI, AllyIndex)
116                     ClanMembersAlive = ClanMembersAlive + 1

                    End If
                
                    ' Reset iddle time
118                 KingIdleTime = 0
                
                End If

            End If

        End With

        
        Exit Sub

KingReviveAlly_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.KingReviveAlly", Erl)
        Resume Next
        
End Sub

Public Sub AI_Healer(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unkown
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler
    
    Dim BestTarget As Integer

    Dim Accion     As Byte
    
    ' Chooses target only if can attack
    If IntervaloPermiteAtacarNPC(NpcIndex) Then
        ' First checks if any ally is paralized in order to remove its paralizis
        BestTarget = AllyParalyzed(NpcIndex, False)

        If BestTarget <> 0 Then
            Accion = eHealerAction.HealAllyParalisis
        Else
            ' Now picks the best target to attack, either user or pet
            BestTarget = HealerBestTarget(NpcIndex, Accion)
            
            ' If no target found, then try to heal injured allies
            If BestTarget = 0 Then
                BestTarget = AllyInjured(NpcIndex)
                
                ' Found an injured Ally
                If BestTarget <> 0 Then
                    Accion = eHealerAction.HealAlly

                End If
                
            End If

        End If

    End If
    
    ' Search for far paralized allies, if not better choise selected
    If BestTarget = 0 Then
        
        'if it's paralized, it has no sense doing it
        If Npclist(NpcIndex).flags.Paralizado = 0 Then
        
            BestTarget = AllyParalyzed(NpcIndex, True)

            If BestTarget <> 0 Then
                Accion = eHealerAction.RescueFarAlly

            End If
            
        End If

    End If
    
    Dim CanMove As Boolean
    
    ' Performes the best action acording to the chosen target
    Call HealerPerformAction(NpcIndex, BestTarget, Accion, CanMove)
    
    ' Moves only if in danger
    If CanMove Then Call HealerMove(NpcIndex)
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en AI_Healer. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function HealerBestTarget(ByVal NpcIndex As Integer, _
                                  ByRef Accion As Byte) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 09/07/2010
    'Pick the best target according to the following criteria:
    '1) "hoaxed" friends MUST be released
    '2) enemy shall be annihilated no matter what
    '3) party healing if no threats
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex           As Integer

    Dim counter             As Long
    
    Dim PetIndex            As Integer
    
    Dim BestTarget          As Integer

    Dim BestTargetInvisible As Boolean
    
    Dim NpcX                As Integer

    Dim NpcY                As Integer
    
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    ' To avoid that all attack the same target
    CounterStep = RandomNumber(0, 1)

    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
    Else
        CounterStart = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
        CounterEnd = 1
        CounterStep = -1

    End If
    
    For counter = CounterStart To CounterEnd Step CounterStep
    
        UserIndex = ModAreas.ConnGroups(CenterPos.Map).UserEntrys(counter)
        
        ' Can be atacked (doesn't matter if invisible or hidden)?
        If UserAtacable(UserIndex, False) Then
                
            ' Checks if ther's a non-paralized user pet in range. If so, then paralize it.
            PetIndex = CheckNearUserPets(NpcIndex, UserIndex)

            If PetIndex <> 0 Then
                BestTarget = PetIndex
                Accion = eHealerAction.ParalizePet
                Exit For

            End If
                
            'Is it in it's range of vision??
            If InVisionRange(UserIndex, NpcX, NpcY) Then
                
                With UserList(UserIndex)
                    
                    ' It's Paralized?
                    If .flags.Paralizado = 0 Then
                        
                        ' It's visible?
                        If (.flags.invisible = 0) And (.flags.Oculto = 0) Then
                            
                            BestTarget = UserIndex
                            Accion = eHealerAction.ParalizeUser
                            
                        End If
                        
                        ' User Paralized, if not better target found then attack it
                    ElseIf BestTarget = 0 Then

                        BestTarget = UserIndex
                        Accion = eHealerAction.AttackUser
                        
                    End If
                    
                End With
                
            End If
        
        End If
        
    Next counter
    
    HealerBestTarget = BestTarget
        
    Exit Function

ErrHandler:
    LogError ("Error en HealerBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub HealerPerformAction(ByVal NpcIndex As Integer, _
                                ByVal BestTarget As Integer, _
                                ByVal Accion As Byte, _
                                ByRef CanMove As Boolean)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 09/07/2010
        'Performs the required action. Determines if the npc can move after perfoming the selected action.
        '***************************************************
        
        On Error GoTo HealerPerformAction_Err
        
    
100     With Npclist(NpcIndex)

102         Select Case Accion
            
                Case eHealerAction.HealAllyParalisis
104                 Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eHealerSpells.RemoveParalisis))
106                 CanMove = False
                
108             Case eHealerAction.ParalizePet
110                 Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eHealerSpells.ParalizeNpc))
112                 CanMove = True
                
114             Case eHealerAction.ParalizeUser
116                 'Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePalabrasMagicas(.Spells(eHealerSpells.ParalizeUser), .Char.CharIndex))
118                 Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eHealerSpells.ParalizeUser), True)
120                 CanMove = False
                
122             Case eHealerAction.AttackUser
124                 Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eHealerSpells.Tormenta), True)
126                 CanMove = False
                
128             Case eHealerAction.HealAlly
130                 Call NpcLanzaSpellSobreNpc(NpcIndex, BestTarget, .Spells(eHealerSpells.Heal))
132                 CanMove = True
                
134             Case eHealerAction.RescueFarAlly
136                 Call GreedyWalkTo(NpcIndex, .Pos.Map, Npclist(BestTarget).Pos.X, Npclist(BestTarget).Pos.Y)
138                 CanMove = False
                 
140             Case Else
142                 CanMove = True
                
            End Select

        End With

        
        Exit Sub

HealerPerformAction_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.HealerPerformAction", Erl)
        Resume Next
        
End Sub

Private Sub HealerMove(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 09/07/2010
        'Doesn't move after casting some spells. Try to avoid swordmasters.
        '***************************************************
        
        On Error GoTo HealerMove_Err
        
    
100     With Npclist(NpcIndex)
        
            ' Can't move if paralized
102         If .flags.Paralizado = 1 Then Exit Sub
    
            ' Only applies to clan map
104         If .Pos.Map <> CenterPos.Map Then Exit Sub
        
            ' if nobody's near or it's idle then returns near the king
106         Call ReturnToCenter(NpcIndex)
        
            ' Runs away from close users
108         If TargetClose(.Pos) Then
110             Call ReturnToCenter(NpcIndex)

            End If
        
        End With

        
        Exit Sub

HealerMove_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.HealerMove", Erl)
        Resume Next
        
End Sub

Public Sub AI_SpellCaster(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler
    
    Dim BestTarget As Integer

    Dim Accion     As Byte
    
    ' It sacrifices itself giving a fatal blow if close to death
    If SpellCasterSacrifice(NpcIndex) Then Exit Sub
    
    ' If healthy then choose best target
    BestTarget = SpellCasterBestTarget(NpcIndex, Accion)
    
    ' Performes the best action acoording to the chosen target
    If BestTarget <> 0 Then
        Call SpellCasterPerformAction(NpcIndex, BestTarget, Accion)

    End If
    
    ' Moves only if in danger
    Call SpellCasterMove(NpcIndex, BestTarget)
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en AI_SpellCaster. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function SpellCasterSacrifice(ByVal NpcIndex As Integer) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 05/07/2010
        'Checks if the required conditions to sacrify are fullfiled, if so then does it.
        'The wand break counter is stored in the boat slot, asuming it'd never be used for any other purpose.
        '***************************************************
        
        On Error GoTo SpellCasterSacrifice_Err
        
    
100     With Npclist(NpcIndex)
    
102         If .Stats.MinHp <= START_COUNTDOWN_MIN_HP Then
104             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARGRANDE, INFINITE_LOOPS))
        
106             If IntervaloPermiteAtacarNPC(NpcIndex) Then
                
108                 WandBreakCounter = WandBreakCounter - 1
110                 Call SpellCasterDestroyWand(NpcIndex, eSpellCasterSpells.Apocalipsis)
                
                    '  Wand broken => dies
112                 If WandBreakCounter = 0 Then Call MuereNpc(NpcIndex, 0)
                
                End If
            
114             SpellCasterSacrifice = True
            
            Else

                'Restore wand break counter
116             If WandBreakCounter <> MAX_WAND_BREAK_VALUE Then
118                 WandBreakCounter = MAX_WAND_BREAK_VALUE
120                 Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))

                End If

            End If

        End With
    
        
        Exit Function

SpellCasterSacrifice_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.SpellCasterSacrifice", Erl)
        Resume Next
        
End Function

Private Function SpellCasterBestTarget(ByVal NpcIndex As Integer, _
                                       ByRef Accion As Byte) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Pick the best target according to the following criteria:
    '1) Invisible enemies can be detected sometimes
    '2) A wizard's mission is background spellcasting attack
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex           As Integer

    Dim counter             As Long
    
    Dim BestTarget          As Integer

    Dim NpcX                As Integer

    Dim NpcY                As Integer

    Dim BestTargetInvisible As Boolean
        
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    ' To avoid that all attack the same target
    CounterStep = RandomNumber(0, 1)

    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
    Else
        CounterStart = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
        CounterEnd = 1
        CounterStep = -1

    End If
    
    For counter = CounterStart To CounterEnd Step CounterStep
    
        UserIndex = ModAreas.ConnGroups(CenterPos.Map).UserEntrys(counter)
        
        'Is it in it's range of vision??
        If InVisionRange(UserIndex, NpcX, NpcY) Then
            
            ' Can be atacked?
            If UserAtacable(UserIndex, False) Then
                
                With UserList(UserIndex)

                    ' It's invisible?
                    If (.flags.invisible = 1) Or (.flags.Oculto = 1) Then
                        
                        ' Try to remove invisibility
                        If (RandomNumber(1, 100) <= 35) Then
                            BestTarget = UserIndex
                            Accion = eSpellCasterAction.RemoveInvi
                            Exit For

                        End If
                        
                        ' Paralized invisible users are good target!
                        If .flags.Paralizado = 1 Then
                            BestTarget = UserIndex
                            BestTargetInvisible = True
                            Accion = eSpellCasterAction.Attack

                        End If
                    
                        ' Visible but paralized?
                    ElseIf (.flags.Paralizado = 1) Then
                        
                        ' If not found an invisible and paralized target, then it's a good one
                        If Not BestTargetInvisible Or BestTarget = 0 Then
                            BestTarget = UserIndex
                            Accion = eSpellCasterAction.Attack

                        End If
                    
                        ' Visible and can move, if not better found then choose it
                    ElseIf BestTarget = 0 Then

                        BestTarget = UserIndex
                        Accion = eSpellCasterAction.Attack
                        
                    End If
                    
                End With
                
            End If
        
        End If
        
    Next counter
    
    SpellCasterBestTarget = BestTarget
    
    Exit Function

ErrHandler:
    LogError ("Error en SpellCasterBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub SpellCasterPerformAction(ByVal NpcIndex As Integer, _
                                     ByVal BestTarget As Integer, _
                                     ByVal Accion As Byte)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 05/07/2010
        'Performs the required action.
        '***************************************************
        
        On Error GoTo SpellCasterPerformAction_Err
        
    
100     With Npclist(NpcIndex)

102         Select Case Accion
            
                Case eSpellCasterAction.Attack
104                 Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eSpellCasterSpells.Apocalipsis), True)
            
106             Case eSpellCasterAction.RemoveInvi
108                 Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eSpellCasterSpells.RemoInvi), True)
                
            End Select

        End With
    
        
        Exit Sub

SpellCasterPerformAction_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.SpellCasterPerformAction", Erl)
        Resume Next
        
End Sub

Private Sub SpellCasterMove(ByVal NpcIndex As Integer, ByVal BestTarget As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 05/07/2010
        'Doesn't move unless nobody is near or there is no best target.
        '***************************************************
        
        On Error GoTo SpellCasterMove_Err
        
    
100     With Npclist(NpcIndex)
        
            ' Can't move if paralized
102         If .flags.Paralizado = 1 Then Exit Sub
        
            ' Only applies to clan map
104         If .Pos.Map <> CenterPos.Map Then Exit Sub
        
            ' if nobody's near or it's idle then returns near the king
106         If (BestTarget = 0) And (IntervaloPermiteAtacarNPC(NpcIndex)) Then Call ReturnToCenter(NpcIndex)
        
            ' Runs away from close users
108         If TargetClose(.Pos) Then
110             Call ReturnToCenter(NpcIndex)

            End If
        
        End With

        
        Exit Sub

SpellCasterMove_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.SpellCasterMove", Erl)
        Resume Next
        
End Sub

Private Sub SpellCasterDestroyWand(ByVal NpcIndex As Integer, ByVal SpellIndex As Integer)

    '***************************************************
    'Author: Unkown
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler

    ''sonidos: 30 y 32, y no los cambien sino termina siendo muy chistoso...
    ''Para el FX utiliza el del hechizos(SpellIndex)
    
    With Npclist(NpcIndex)

        Select Case WandBreakCounter

            Case 5
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("Rahma", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 4
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("vortax", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 3
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("Zill", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 2
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("yaka E'nta", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 1
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("Korata!!", .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 0
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead(vbNullString, .Char.CharIndex, vbGreen))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))
                
                Call DealWandDamage(NpcIndex, .Spells(SpellIndex))

        End Select
        
    End With
    
    Exit Sub

ErrHandler:
    LogError ("Error en SpellCasterDestroyWand. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub DealWandDamage(ByVal NpcIndex As Integer, ByVal SpellIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 17/06/2011
    'Description: Deals massive wand damage.
    '17/06/2011: Amraphen - User's hp is updated once he gets hit by wand damage.
    '***************************************************
    On Error GoTo ErrHandler

    Dim UserIndex As Integer

    Dim counter   As Long
    
    Dim Distancia As Integer

    Dim Danio     As Integer
    
    Dim X         As Integer

    Dim Y         As Integer
    
    With Npclist(NpcIndex)
        
        X = .Pos.X
        Y = .Pos.Y
        
        For counter = 1 To ModAreas.ConnGroups(CenterPos.Map).CountEntrys
            
            UserIndex = ModAreas.ConnGroups(CenterPos.Map).UserEntrys(counter)
            
            If UserIndex <> 0 Then
            
                With UserList(UserIndex)
                    
                    If UserAtacable(UserIndex, False, False) Then
                        
                        Distancia = UserDistance(UserIndex, X, Y)
                        Danio = Abs(Int(880 / (Distancia ^ (3 / 7))))
                        
                        .Stats.MinHp = .Stats.MinHp - Danio
                        
                        Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).name & " te ha quitado " & Danio & " puntos de vida al romper su vara.", FontTypeNames.FONTTYPE_FIGHT)
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).wav, .Pos.X, .Pos.Y))
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                        
                        If .Stats.MinHp < 1 Then
                            .Stats.MinHp = 0
                            Call UserDie(UserIndex)
                        Else
                            'Updates user's hp:
                            Call WriteUpdateHP(UserIndex)

                            ' Also hit pets
                            If .NroMascotas > 0 Then Call DealWandDamageToPets(NpcIndex, UserIndex, SpellIndex)
                            
                        End If
                        
                    End If
                    
                End With
            
            End If
            
        Next counter
        
    End With
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en DealWandDamage. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub DealWandDamageToPets(ByVal NpcIndex As Integer, _
                                 ByVal UserIndex As Integer, _
                                 ByVal SpellIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/09/2010
    'Deals massive wand damage to user pets.
    '***************************************************
    On Error GoTo ErrHandler

    Dim PetIndex   As Integer

    Dim PetCounter As Long

    Dim Distancia  As Integer

    Dim Danio      As Integer
    
    Dim NpcX       As Integer

    Dim NpcY       As Integer
    
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    
    For PetCounter = 1 To MAXMASCOTAS
                                    
        PetIndex = UserList(UserIndex).MascotasIndex(PetCounter)

        If PetIndex > 0 Then

            With Npclist(PetIndex)
           
                Distancia = NpcDistance(PetIndex, NpcX, NpcY)
                Danio = Abs(Int(880 / (Distancia ^ (3 / 7))))

                .Stats.MinHp = .Stats.MinHp - Danio
                
                Call SendData(SendTarget.ToNPCArea, PetIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).wav, .Pos.X, .Pos.Y))
                Call SendData(SendTarget.ToNPCArea, PetIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                
                If .Stats.MinHp < 1 Then
                    .Stats.MinHp = 0
                    Call MuereNpc(PetIndex, 0)

                End If
                
            End With

        End If
        
    Next PetCounter
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en DealWandDamageToPets. Error: " & Err.Number & " - " & Err.description)

End Sub

Public Sub AI_SwordMaster(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler

    Dim BestTarget As Integer
    
    ' First choose best target
    BestTarget = SwordMasterBestTarget(NpcIndex)
    
    ' Then moves, close to target user if in range
    Call SwordMasterMove(NpcIndex, BestTarget)
    
    ' Finally, attacks sorrounding targets
    Call SwordMasterAttack(NpcIndex)
    
    Exit Sub

ErrHandler:
    LogError ("Error en AI_SwordMaster. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function SwordMasterBestTarget(ByVal NpcIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/06/2010
    'Picks the best target according to the following criteria:
    '1) The nearest enemy will be attacked.
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex          As Integer

    Dim counter            As Long
    
    Dim BestTarget         As Integer

    Dim BestTargetDistance As Integer
    
    Dim NpcX               As Integer

    Dim NpcY               As Integer

    Dim Distance           As Integer
    
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    ' To avoid that all attack the same target
    CounterStep = RandomNumber(0, 1)

    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
    Else
        CounterStart = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
        CounterEnd = 1
        CounterStep = -1

    End If
    
    For counter = CounterStart To CounterEnd Step CounterStep
    
        UserIndex = ModAreas.ConnGroups(CenterPos.Map).UserEntrys(counter)
        
        'Is it in it's range of vision??
        If InVisionRange(UserIndex, NpcX, NpcY) Then
            
            ' Can be atacked?
            If UserAtacable(UserIndex) Then

                ' Checks limits in order to avoid to separate it from the rest of the clan
                If UserReachable(NpcIndex, UserIndex) Then
                
                    If BestTarget <> 0 Then
                        
                        ' Has priority to attack the nearest
                        Distance = UserDistance(UserIndex, NpcX, NpcY)
                        
                        If Distance < BestTargetDistance Then
                            BestTarget = UserIndex
                            BestTargetDistance = Distance

                        End If
                        
                    Else
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NpcX, NpcY)

                    End If
                
                End If
            
            End If
        
        End If
        
    Next counter
    
    SwordMasterBestTarget = BestTarget
    
    Exit Function

ErrHandler:
    LogError ("Error en SwordMasterBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub SwordMasterMove(ByVal NpcIndex As Integer, ByVal BestTarget As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'La idea es que no lo "alejen" del rey y despues queden
    'lejos de la batalla cuando matan a un enemigo o este
    'sale del area de combate (tipica forma de separar un clan)
    '***************************************************

    On Error GoTo ErrHandler

    With Npclist(NpcIndex)
    
        ' If paralized can't move
        If .flags.Paralizado = 1 Then Exit Sub
        
        ' Only applies to clan map
        If .Pos.Map <> CenterPos.Map Then Exit Sub
        
        ' If it's far from protection then returns to center
        If FarFromTeam(NpcIndex) Then
            Call ReturnToCenter(NpcIndex)
            Exit Sub

        End If
        
        ' If someone's in range, then go close to him
        If BestTarget > 0 Then
            
            With UserList(BestTarget)
                Call GreedyWalkTo(NpcIndex, CenterPos.Map, .Pos.X, .Pos.Y)

            End With
            
            'Returns to center if no target found
        Else
            Call ReturnToCenter(NpcIndex)

        End If
        
    End With
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en SwordMasterMove. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub SwordMasterAttack(ByVal NpcIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/06/2010
    'Searchs and Attacks the sorrounding target
    '***************************************************

    On Error GoTo ErrHandler

    Dim headingloop As Byte

    Dim NpcPos      As WorldPos

    Dim UserIndex   As Integer
    
    With Npclist(NpcIndex)
    
        ' Check the four directions
        For headingloop = eHeading.NORTH To eHeading.WEST
        
            NpcPos = .Pos
            Call HeadtoPos(headingloop, NpcPos)
            
            UserIndex = MapData(NpcPos.Map, NpcPos.X, NpcPos.Y).UserIndex
            
            If UserIndex > 0 Then
                If UserAtacable(UserIndex, False) Then
                    If NpcAtacaUser(NpcIndex, UserIndex, headingloop) Then
                        Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)
                        Exit For

                    End If

                End If

            End If
            
        Next headingloop
        
    End With

    Exit Sub
    
ErrHandler:
    LogError ("Error en SwordMasterAttack. Error: " & Err.Number & " - " & Err.description)

End Sub

Public Sub AI_Shooter(ByVal NpcIndex As Integer)
        
        On Error GoTo AI_Shooter_Err
        

        '***************************************************
        'Author: ZaMa
        'Last Modification: -
        '***************************************************
        Dim BestTarget As Integer
    
        ' First choose best target
100     BestTarget = ShooterBestTarget(NpcIndex)
    
        ' Found a target => attack!
102     If BestTarget > 0 Then
            ' Attack with arrow
104         Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, Npclist(NpcIndex).Spells(1))

        End If
    
        ' Then moves, close to target user if attacking one
106     Call ShooterMove(NpcIndex, BestTarget)
    
        Exit Sub
    
ErrHandler:
108     LogError ("Error en AI_Shooter. Error: " & Err.Number & " - " & Err.description)

        
        Exit Sub

AI_Shooter_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.AI_Shooter", Erl)
        Resume Next
        
End Sub

Private Function ShooterBestTarget(ByVal NpcIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'Picks the best target according to the following criteria:
    '1) Magic clases ARE dangerous, but they are weak too, they're
    '   our primary target
    '2) in any other case, our nearest enemy will be attacked
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex          As Integer

    Dim counter            As Long
    
    Dim BestTarget         As Integer

    Dim BestTargetDistance As Integer
    
    Dim NpcX               As Integer

    Dim NpcY               As Integer

    Dim Distance           As Integer
        
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
        
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    ' To avoid that all attack the same target
    CounterStep = RandomNumber(0, 1)

    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
    Else
        CounterStart = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
        CounterEnd = 1
        CounterStep = -1

    End If
    
    For counter = CounterStart To CounterEnd Step CounterStep
    
        UserIndex = ModAreas.ConnGroups(CenterPos.Map).UserEntrys(counter)
        
        'Is it in it's range of vision??
        If InVisionRange(UserIndex, NpcX, NpcY) Then
            
            ' Can be atacked?
            If UserAtacable(UserIndex) Then
                
                ' Has Priority to attack magic clases
                If EsClaseMagica(UserIndex) Then
                    BestTarget = UserIndex
                    Exit For
                Else

                    If BestTarget <> 0 Then
                        
                        ' Has priority to attack the nearest
                        Distance = UserDistance(UserIndex, NpcX, NpcY)
                        
                        If Distance < BestTargetDistance Then
                            BestTarget = UserIndex
                            BestTargetDistance = Distance

                        End If
                        
                    Else
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NpcX, NpcY)

                    End If
                    
                End If
            
            End If
        
        End If
        
    Next counter
    
    ShooterBestTarget = BestTarget
        
    Exit Function

ErrHandler:
    LogError ("Error en ShooterBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub ShooterMove(ByVal NpcIndex As Integer, ByVal BestTarget As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'Vamos a setear el hold on del cazador en el medio entre el rey
    'y el atacante. De esta manera se lo podra atacar aun asi esta lejos
    'pero sin alejarse del rango de los an hoax vorps de los
    'clerigos o rey. A menos q este paralizado, claro
    '***************************************************

    On Error GoTo ErrHandler

    With Npclist(NpcIndex)
    
        ' If paralized can't move
        If .flags.Paralizado = 1 Then Exit Sub
        
        ' Only applies to clan map
        If .Pos.Map <> CenterPos.Map Then Exit Sub

        ' If it's far from protection then returns to center
        If FarFromTeam(NpcIndex) Then
            Call ReturnToCenter(NpcIndex)
            Exit Sub

        End If
        
        ' Don't go too far from center
        If Abs(CenterPos.X - .Pos.X) > 4 Or Abs(CenterPos.Y - .Pos.Y) > 4 Then
            Call ReturnToCenter(NpcIndex)
            Exit Sub

        End If
        
        ' If has attacked someone, then go close to him
        If BestTarget > 0 Then
            
            Call GreedyWalkTo(NpcIndex, CenterPos.Map, CenterPos.X + ((UserList(BestTarget).Pos.X - CenterPos.X) / 2), CenterPos.Y + ((UserList(BestTarget).Pos.Y - CenterPos.Y) / 2))
            Exit Sub
        Else
        
            ' Search for aproaching enemies
            Dim counter   As Long

            Dim UserIndex As Integer
            
            For counter = 1 To ModAreas.ConnGroups(CenterPos.Map).CountEntrys
            
                UserIndex = ModAreas.ConnGroups(CenterPos.Map).UserEntrys(counter)
                
                'Is it in extended range of vision from center position??
                If InVisionRange(UserIndex, CenterPos.X, CenterPos.Y, True) Then
                    
                    ' Can be atacked?
                    If UserAtacable(UserIndex) Then
                        
                        With UserList(UserIndex)
                            Call GreedyWalkTo(NpcIndex, CenterPos.Map, .Pos.X, .Pos.Y)

                        End With
                        
                        Exit Sub

                    End If
                    
                End If

            Next counter

        End If
        
        'Returns to center if no target found
        If IntervaloPermiteAtacarNPC(NpcIndex) Then Call ReturnToCenter(NpcIndex)
         
    End With
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en ShooterMove. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub AI_Thief(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler

    Dim BestTarget As Integer

    Dim Action     As Byte
    
    ' First choose best target
    BestTarget = ThiefBestTarget(NpcIndex, Action)
    
    ' Perform's thief action acording to its priorities
    Call ThiefPerfomAction(NpcIndex, BestTarget, Action)
    
    Exit Sub

ErrHandler:
    LogError ("Error en AI_Thief. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function ThiefBestTarget(ByVal NpcIndex As Integer, _
                                 ByRef Action As Byte) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/06/2010
    'Picks the best target according to the following criteria:
    '1) Has priority to attack hunters.
    '2) Targets the nearer user.
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex          As Integer

    Dim counter            As Long
    
    Dim BestTarget         As Integer

    Dim BestTargetDistance As Integer
    
    Dim NpcX               As Integer

    Dim NpcY               As Integer

    Dim Distance           As Integer
    
    NpcX = Npclist(NpcIndex).Pos.X
    NpcY = Npclist(NpcIndex).Pos.Y
    
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    ' To avoid that all attack the same target
    CounterStep = RandomNumber(0, 1)

    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
    Else
        CounterStart = ModAreas.ConnGroups(CenterPos.Map).CountEntrys
        CounterEnd = 1
        CounterStep = -1

    End If
    
    For counter = CounterStart To CounterEnd Step CounterStep
    
        UserIndex = ModAreas.ConnGroups(CenterPos.Map).UserEntrys(counter)
        
        'Is it in it's range of vision??
        If InVisionRange(UserIndex, NpcX, NpcY) Then
            
            ' Can be atacked?
            If UserAtacable(UserIndex, False) Then

                ' Checks limits in order to avoid to separate it from the rest of the clan
                If UserReachable(NpcIndex, UserIndex) Then
                
                    ' Is it a hunter? => Has priority
                    If UserList(UserIndex).clase = eClass.Hunter Then
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NpcX, NpcY)
                        Exit For

                    End If
                    
                    If BestTarget <> 0 Then
                        
                        ' Has priority to attack the nearest
                        Distance = UserDistance(UserIndex, NpcX, NpcY)
                        
                        If Distance < BestTargetDistance Then
                            BestTarget = UserIndex
                            BestTargetDistance = Distance

                        End If
                        
                    Else
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NpcX, NpcY)

                    End If
                
                End If
            
            End If
        
        End If
        
    Next counter
    
    ' Couldn't find anyone
    If BestTarget = 0 Then
        Action = eThiefAction.None
    Else

        ' if close to target => Snatch/Steal
        If BestTargetDistance <= THIEF_STEAL_DISTANCE Then
            Action = eThiefAction.Steal
        
            ' Too far => Attack while aproaching
        Else
            Action = eThiefAction.Attack

        End If

    End If
    
    ThiefBestTarget = BestTarget
    
    Exit Function

ErrHandler:
    LogError ("Error en ThiefBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub ThiefPerfomAction(ByVal NpcIndex As Integer, _
                              ByVal BestTarget As Integer, _
                              ByVal Action As Byte)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 28/10/2010
    '28/10/2010: ZaMa - Now thief doesn't steal/snatch admins.
    '***************************************************
    On Error GoTo ErrHandler
    
    Dim WeaponEqpSlot   As Byte

    Dim MunicionEqpSlot As Byte
    
    With Npclist(NpcIndex)
        
        Select Case Action
            
            Case eThiefAction.Attack, eThiefAction.Steal
                
                ' Can't do anything
                If IntervaloPermiteAtacarNPC(NpcIndex) Then
                
                    ' Attack with arrow
                    If Action = eThiefAction.Attack Then
                        Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eThiefSpells.Arrow), True)
                        
                        ' Can paralize
                        If RandomNumber(1, 100) < 10 Then
                            Call NpcLanzaSpellSobreUser(NpcIndex, BestTarget, .Spells(eThiefSpells.Paralisis), True)
                            Call WriteConsoleMsg(BestTarget, "" & .name & " te ha paralizado con su golpe!!", FontTypeNames.FONTTYPE_VENENO)
                        
                            ' Can unequip weapon
                        ElseIf RandomNumber(1, 100) < 22 Then
                            
                            WeaponEqpSlot = UserList(BestTarget).Invent.WeaponEqpSlot

                            If WeaponEqpSlot <> 0 Then Call Desequipar(BestTarget, WeaponEqpSlot)
                            
                        End If
                        
                        ' Try to snatch/Steal (except for admins)
                    ElseIf Not EsGM(BestTarget) Then
                        
                        ' Can snatch weapon or arrows
                        If RandomNumber(1, 100) < 16 Then
                            
                            WeaponEqpSlot = UserList(BestTarget).Invent.WeaponEqpSlot

                            If WeaponEqpSlot <> 0 Then
                                
                                If SnatchItem(NpcIndex, BestTarget, WeaponEqpSlot) Then
                                    Call WriteConsoleMsg(BestTarget, "" & .name & " te ha arrebatado tu arma!!", FontTypeNames.FONTTYPE_VENENO)

                                End If
                                
                            Else
                                MunicionEqpSlot = UserList(BestTarget).Invent.MunicionEqpSlot

                                If MunicionEqpSlot <> 0 Then
                                
                                    If SnatchItem(NpcIndex, BestTarget, MunicionEqpSlot) Then
                                        Call WriteConsoleMsg(BestTarget, "" & .name & " te ha arrebatado tus municiones!!", FontTypeNames.FONTTYPE_VENENO)

                                    End If
                                    
                                End If

                            End If
                        
                            ' Can steal the items of a slot
                        ElseIf RandomNumber(1, 100) < 16 Then
                        
                            Dim slot As Byte

                            slot = RandomNumber(1, UserList(BestTarget).CurrentInventorySlots)
                            
                            If ThiefStealITem(NpcIndex, BestTarget, slot) Then
                                Call WriteConsoleMsg(BestTarget, "" & .name & " te esta robando!!", FontTypeNames.FONTTYPE_VENENO)

                            End If
                            
                        End If

                    End If
                    
                End If
                
                ' Has chances of becoming invisible (if visible)
                If .flags.invisible = 0 Then
                    
                    If RandomNumber(1, 100) < 25 Then Call ThiefTurnInvisible(NpcIndex, True)
                
                    ' Is invisible, can turn visible.
                Else

                    If RandomNumber(1, 100) < 13 Then Call ThiefTurnInvisible(NpcIndex, False)

                End If
                
                ' Get closer to target
                Call GreedyWalkTo(NpcIndex, CenterPos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
                
            Case eThiefAction.None
                
                ' Turns visible if no threat
                If .flags.invisible = 1 Then Call ThiefTurnInvisible(NpcIndex, False)
                
                Call ReturnToCenter(NpcIndex)
                
        End Select
        
    End With

    Exit Sub

ErrHandler:
    LogError ("Error en ThiefPerfomAction. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function SnatchItem(ByVal NpcIndex As Integer, _
                            ByVal targetIndex As Integer, _
                            ByVal slot As Byte) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/09/2010
    'Snatchs slot items from user. No validations are made bacause it's either weapon or munition.
    '***************************************************
    On Error GoTo ErrHandler

    Dim DropObj As obj

    Dim nPos    As WorldPos
    
    With UserList(targetIndex)
        DropObj.Amount = .Invent.Object(slot).Amount
        DropObj.ObjIndex = .Invent.Object(slot).ObjIndex
    
        ' Search for suitable place to drop item
        Call Tilelibre(.Pos, nPos, DropObj, False, True)
        
        ' Found any?
        If nPos.X <> 0 And nPos.Y <> 0 Then
            ' Drop it
            Call MakeObj(DropObj, nPos.Map, nPos.X, nPos.Y)
            Call QuitarUserInvItem(targetIndex, slot, DropObj.Amount)
            Call UpdateUserInv(False, targetIndex, slot)
            
            ' Suceed
            SnatchItem = True

        End If

    End With
    
    Exit Function
    
ErrHandler:
    LogError ("Error en SnatchItem. Error: " & Err.Number & " - " & Err.description)

End Function

Private Function ThiefStealITem(ByVal NpcIndex As Integer, _
                                ByVal targetIndex As Integer, _
                                ByVal slot As Byte) As Boolean
        
        On Error GoTo ThiefStealITem_Err
        

        '***************************************************
        'Author: ZaMa
        'Last Modification: 25/09/2010
        'Steal slot items from user. If no space in thief's inventory, drop it on floor.
        '22/10/2010: ZaMa - Now thief doesn't try to stel slots with no items in it.
        '***************************************************
        Dim StolenObj As obj

        Dim NroItems  As Integer
    
        ' If doesn't have any item, the doesn't do anything
100     If UserList(targetIndex).Invent.Object(slot).ObjIndex = 0 Then Exit Function
    
        ' If not stealable, then don't do anything
102     If Not ObjEsRobable(targetIndex, slot) Then Exit Function

        ' If thief's inventory is full, then snatch it
104     NroItems = Npclist(NpcIndex).Invent.NroItems

106     If NroItems = 20 Then
        
108         ThiefStealITem = SnatchItem(NpcIndex, targetIndex, slot)
             
            ' Steal it
        Else
        
110         StolenObj.Amount = UserList(targetIndex).Invent.Object(slot).Amount
112         StolenObj.ObjIndex = UserList(targetIndex).Invent.Object(slot).ObjIndex
        
            ' Add it to Thief inventory
114         With Npclist(NpcIndex)
116             NroItems = NroItems + 1
118             .Invent.Object(NroItems).Amount = StolenObj.Amount
120             .Invent.Object(NroItems).ObjIndex = StolenObj.ObjIndex
122             .Invent.NroItems = NroItems

            End With
        
            ' Take it from user inventory
124         Call QuitarUserInvItem(targetIndex, slot, StolenObj.Amount)
126         Call UpdateUserInv(False, targetIndex, slot)
        
128         ThiefStealITem = True
        
        End If
    
        
        Exit Function

ThiefStealITem_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.ThiefStealITem", Erl)
        Resume Next
        
End Function

Private Sub ThiefTurnInvisible(ByVal NpcIndex As Integer, ByVal TurnVisible As Boolean)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 26/09/2010
        'Turn thief visible/Invisible.
        '***************************************************
        
        On Error GoTo ThiefTurnInvisible_Err
        
    
100     With Npclist(NpcIndex)
    
102         If TurnVisible Then
104             .flags.invisible = 1
            Else
106             .flags.invisible = 0

            End If
        
108         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSetInvisible(.Char.CharIndex, TurnVisible))
    
        End With
    
        
        Exit Sub

ThiefTurnInvisible_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.ThiefTurnInvisible", Erl)
        Resume Next
        
End Sub

Private Function TargetClose(ByRef Pos As WorldPos) As Boolean
        
        On Error GoTo TargetClose_Err
        

        '***************************************************
        'Author: ZaMa
        'Last Modification: 05/07/2010
        'Checks if there is any attackable user sorrounding the given position.
        '***************************************************
        Dim UserIndex   As Integer
    
        Dim nPos        As WorldPos

        Dim headingloop As Byte
    
100     For headingloop = eHeading.NORTH To eHeading.WEST
        
102         nPos = Pos
104         Call HeadtoPos(headingloop, nPos)
        
106         UserIndex = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
        
108         If UserIndex <> 0 Then
110             If UserAtacable(UserIndex) Then
112                 TargetClose = True
                    Exit Function

                End If

            End If

114     Next headingloop
    
        
        Exit Function

TargetClose_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.TargetClose", Erl)
        Resume Next
        
End Function

Private Function CheckNearUserPets(ByVal NpcIndex As Integer, _
                                   ByVal UserIndex As Integer) As Integer
        
        On Error GoTo CheckNearUserPets_Err
        

        '***************************************************
        'Author: ZaMa
        'Last Modification: 09/07/2010
        'Checks if there is a non-paralized user pet in cleric vision range.
        '***************************************************
        Dim PetCounter As Long

        Dim PetIndex   As Integer
    
100     With UserList(UserIndex)
        
102         If .NroMascotas = 0 Then Exit Function
        
104         For PetCounter = 1 To MAXMASCOTAS
        
106             PetIndex = .MascotasIndex(PetCounter)

108             If PetIndex > 0 Then
            
110                 With Npclist(PetIndex)
                    
                        ' If it is visible for the Pretorian
112                     If InVisionRangeNpc(NpcIndex, .Pos.X, .Pos.Y) Then
                        
                            ' If not paralized then it's a suitable target
114                         If .flags.Paralizado = 0 Then
116                             CheckNearUserPets = PetIndex
                                Exit Function

                            End If
                    
                        End If
                    
                    End With
               
                End If
            
118         Next PetCounter
        
        End With

        
        Exit Function

CheckNearUserPets_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.CheckNearUserPets", Erl)
        Resume Next
        
End Function

Private Function EsClaseMagica(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unkown
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        EsClaseMagica = .clase = eClass.Mage Or .clase = eClass.Cleric Or .clase = eClass.Druid Or .clase = eClass.Bard

    End With
    
    Exit Function

ErrHandler:
    LogError ("Error en EsClaseMagica. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub GreedyWalkTo(ByVal NpcIndex As Integer, _
                         ByVal TargetMap As Integer, _
                         ByVal TargetX As Integer, _
                         ByVal TargetY As Integer)

    '***************************************************
    'Author: Unknown
    '  Este procedimiento es llamado cada vez que un NPC deba ir
    '  a otro lugar en el mismo TargetMap. Utiliza una tecnica
    '  de programacion greedy no deterministica.
    '  Cada paso azaroso que me acerque al destino, es un buen paso.
    '  Si no hay mejor paso valido, entonces hay que volver atras y reintentar.
    '  Si no puedo moverme, me considero piketeado
    '  La funcion es larga, pero es O(1) - orden algoritmico temporal constante
    'Last Modification: 26/09/2010
    'Rapsodius - Changed Mod by And for speed
    '26/09/2010: ZaMa - Now make movements as normal npcs, which allows to kick caspers and invisible admins.
    '***************************************************
    On Error GoTo ErrHandler

    Dim NpcX      As Integer

    Dim NpcY      As Integer

    Dim RandomDir As Integer
    
    With Npclist(NpcIndex).Pos

        If .Map <> TargetMap Then Exit Sub
        
        NpcX = .X
        NpcY = .Y

    End With
    
    ' Arrived to destination
    If (NpcX = TargetX And NpcY = TargetY) Then Exit Sub
    
    ' Try to reach target
    If (NpcX > TargetX) Then
    
        ' Target is Down-Left
        If (NpcY < TargetY) Then
            
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move down
            
                If MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub

                End If
                
            Else

                ''random first move
                If MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub

                End If
                
            End If
            
            ' Target is Up-Left
        ElseIf (NpcY > TargetY) Then
        
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move up
            
                If MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub

                End If
                
            Else

                ''random first move
                If MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub

                End If
                
            End If
            
            ' Target is Straight Left
        Else
        
            If MoveNPCChar(NpcIndex, eHeading.WEST) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                Exit Sub
            Else
                ' If moves to east, algorithm fails because starts a loop
                Call MoveFailed(NpcIndex)

            End If
            
        End If
    
    ElseIf (NpcX < TargetX) Then
        
        ' Target is Down-Right
        If (NpcY < TargetY) Then
            
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move down
            
                If MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub

                End If
                
            Else    ''random first move
                
                If MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub

                End If
                
            End If
        
            ' Target is Up-Right
        ElseIf (NpcY > TargetY) Then
        
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move up
            
                If MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub

                End If
                
            Else
            
                If MoveNPCChar(NpcIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                    Exit Sub

                End If
                
            End If
        
            ' Target is Straight Right
        Else
        
            If MoveNPCChar(NpcIndex, eHeading.EAST) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                Exit Sub
            Else
                ' If moves to west, algorithm fails because starts a loop
                Call MoveFailed(NpcIndex)

            End If
            
        End If
    
        ' Target is straight Up/Down
    Else
    
        ' Target is straight Up
        If (NpcY > TargetY) Then
        
            If MoveNPCChar(NpcIndex, eHeading.NORTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                Exit Sub
            Else
                ' If moves to south, algorithm fails because starts a loop
                Call MoveFailed(NpcIndex)

            End If
            
            ' Target is straight Down
        Else
        
            If MoveNPCChar(NpcIndex, eHeading.SOUTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.EAST) Then
                Exit Sub
            ElseIf MoveNPCChar(NpcIndex, eHeading.WEST) Then
                Exit Sub
            Else
                ' If moves to north, algorithm fails because starts a loop
                Call MoveFailed(NpcIndex)

            End If
            
        End If
        
    End If
    
    Exit Sub

ErrHandler:
    LogError ("Error en GreedyWalkTo. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub MoveFailed(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 26/06/2010
        'Npc shouts a message overhead because cannot move (to avoid a loop)
        '***************************************************
        
        On Error GoTo MoveFailed_Err
        

100     With Npclist(NpcIndex)

102         If IntervaloPermiteAtacarNPC(NpcIndex) Then
104             If (RandomNumber(1, 100) > 95) Then
106                 Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("Maldito bastardo, Ven aqui!", .Char.CharIndex, vbYellow))

                End If

            End If

        End With

        
        Exit Sub

MoveFailed_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.MoveFailed", Erl)
        Resume Next
        
End Sub

Private Sub ReturnToCenter(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 19/09/2010
        'Returns to the center, where the king is.
        '***************************************************
        
        On Error GoTo ReturnToCenter_Err
        
    
100     Call GreedyWalkTo(NpcIndex, CenterPos.Map, CenterPos.X, CenterPos.Y)
    
        
        Exit Sub

ReturnToCenter_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.ReturnToCenter", Erl)
        Resume Next
        
End Sub

Private Function FarFromTeam(ByVal NpcIndex) As Boolean
        '***************************************************
        'Author: Unknown & ZaMa
        'Last Modification: 19/09/2010
        'Checks if Npc is far from cleric protection ring
        '***************************************************
        
        On Error GoTo FarFromTeam_Err
        
    
100     If Npclist(NpcIndex).Pos.Map <> CenterPos.Map Then Exit Function
    
102     FarFromTeam = (Abs(Npclist(NpcIndex).Pos.Y - CenterPos.Y) > MAX_DISTANCE) Or (Abs(Npclist(NpcIndex).Pos.X - CenterPos.X) > MAX_DISTANCE)
    
        Exit Function

ErrHandler:
104     LogError ("Error en FarFromTeam. Error: " & Err.Number & " - " & Err.description)

        
        Exit Function

FarFromTeam_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.FarFromTeam", Erl)
        Resume Next
        
End Function

Private Function UserReachable(ByVal NpcIndex As Integer, _
                               ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    'Ignores users who are too far in order to avoid being separated from the rest of the team.
    '***************************************************
    On Error GoTo ErrHandler
        
    If Npclist(NpcIndex).Pos.Map <> CenterPos.Map Then Exit Function

    UserReachable = (Abs(UserList(UserIndex).Pos.X - CenterPos.X) < MAX_DISTANCE) And (Abs(UserList(UserIndex).Pos.Y - CenterPos.Y) < MAX_DISTANCE)
    
    Exit Function
    
ErrHandler:
    LogError ("Error en UserReachable. Error: " & Err.Number & " - " & Err.description)

End Function

Private Function InVisionRange(ByVal UserIndex As Integer, _
                               ByVal X As Integer, _
                               ByVal Y As Integer, _
                               Optional ByVal ExtendedRange As Boolean = False) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 24/06/2010
        'Checks whether user is in vision range or not
        '***************************************************
        
        On Error GoTo InVisionRange_Err
        

        Dim Rango As Integer

100     Rango = val(IIf(ExtendedRange, 2, 1))
    
102     With UserList(UserIndex)
104         InVisionRange = Abs(.Pos.X - X) <= RANGO_VISION_X * Rango And Abs(.Pos.Y - Y) <= RANGO_VISION_Y * Rango

        End With
    
        
        Exit Function

InVisionRange_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.InVisionRange", Erl)
        Resume Next
        
End Function

Private Function InVisionRangeNpc(ByVal NpcIndex As Integer, _
                                  ByVal X As Integer, _
                                  ByVal Y As Integer, _
                                  Optional ByVal ExtendedRange As Boolean = False) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 24/06/2010
        'Checks whether npc is in vision range or not
        '***************************************************
        
        On Error GoTo InVisionRangeNpc_Err
        
    
        Dim Rango As Integer

100     Rango = val(IIf(ExtendedRange, 2, 1))

102     With Npclist(NpcIndex)
104         InVisionRangeNpc = Abs(.Pos.X - X) <= (RANGO_VISION_X * Rango) And Abs(.Pos.Y - Y) <= (RANGO_VISION_Y * Rango)
        End With
    
        
        Exit Function

InVisionRangeNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.InVisionRangeNpc", Erl)
        Resume Next
        
End Function

Private Function UserAtacable(ByVal UserIndex As Integer, _
                              Optional ByVal CheckVisibility As Boolean = True, _
                              Optional ByVal AttackAdmin As Boolean = False) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 05/10/2010
        'DEtermines whether the user can be atacked or not
        '05/10/2010: ZaMa - Now doesn't allow to attack admins sometimes.
        '***************************************************
        
        On Error GoTo UserAtacable_Err
        

100     With UserList(UserIndex).flags
102         UserAtacable = .AdminInvisible = 0 And .Muerto = 0
                       
104         If CheckVisibility Then
106             UserAtacable = UserAtacable And .Oculto = 0 And .invisible = 0

            End If
        
108         If Not AttackAdmin Then
110             UserAtacable = UserAtacable And (Not EsGM(UserIndex))

            End If
        
        End With
                        
        
        Exit Function

UserAtacable_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.UserAtacable", Erl)
        Resume Next
        
End Function

Private Function UserDistance(ByVal UserIndex As Integer, _
                              ByVal X As Integer, _
                              ByVal Y As Integer) As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 24/06/2010
        'Calculates the user distance according to the given coords.
        '***************************************************
        
        On Error GoTo UserDistance_Err
        

100     With UserList(UserIndex)
102         UserDistance = Abs(X - .Pos.X) + Abs(Y - .Pos.Y)

        End With
    
        
        Exit Function

UserDistance_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.UserDistance", Erl)
        Resume Next
        
End Function

Private Function NpcDistance(ByVal NpcIndex As Integer, _
                             ByVal X As Integer, _
                             ByVal Y As Integer) As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 24/06/2010
        'Calculates the npc distance according to the given coords.
        '***************************************************
        
        On Error GoTo NpcDistance_Err
        

100     With Npclist(NpcIndex)
102         NpcDistance = Abs(X - .Pos.X) + Abs(Y - .Pos.Y)

        End With
    
        
        Exit Function

NpcDistance_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.NpcDistance", Erl)
        Resume Next
        
End Function

Public Sub MuerePretoriano(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 27/06/2010
        'Eliminates the pretorian from the current alive ones, and respawn the clan if in pretorian's default map.
        '***************************************************
        
        On Error GoTo MuerePretoriano_Err
        

        Dim PretoIndex As Integer
    
100     For PretoIndex = 1 To NRO_PRETORIANOS

            ' Remove npc index
102         If Pretorianos(PretoIndex).NpcIndex = NpcIndex Then
104             Pretorianos(PretoIndex).NpcIndex = 0
106             ClanMembersAlive = ClanMembersAlive - 1
            
                ' Entire clan has been defeated
108             If ClanMembersAlive = 0 Then

                    ' Respawn it? (Only allowed in pretorian default map)
110                 If RespawnClan Then
                    
                        Dim NewSpawnPos As WorldPos
                
                        ' Switch respawn place (alternate between two places)
112                     If CenterPos.X = LeftSpawnPos.X Then
114                         NewSpawnPos = RightSpawnPos
                        Else
116                         NewSpawnPos = LeftSpawnPos

                        End If
                    
                        ' Spawn it
118                     Call SpawnClan(NewSpawnPos.Map, NewSpawnPos.X, NewSpawnPos.Y, ClanIndex, True)
                    Else
120                     ClanActive = False

                    End If

                End If
            
                Exit Sub

            End If

122     Next PretoIndex
    
        
        Exit Sub

MuerePretoriano_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.MuerePretoriano", Erl)
        Resume Next
        
End Sub

Private Function AllyParalyzed(ByVal NpcIndex As Integer, _
                               ByVal ExtendedRange As Boolean) As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 27/06/2010
        'Returns the index of the paralized ally if exists one.
        '***************************************************
        
        On Error GoTo AllyParalyzed_Err
        
   
        Dim PretoIndex As Integer

        Dim AllyIndex  As Integer
    
100     For PretoIndex = 1 To NRO_PRETORIANOS
        
102         AllyIndex = Pretorianos(PretoIndex).NpcIndex

104         If AllyIndex <> 0 Then
106             If Npclist(AllyIndex).flags.Paralizado = 1 Then
108                 If InVisionRangeNpc(AllyIndex, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, ExtendedRange) Then
110                     AllyParalyzed = AllyIndex
                        Exit Function

                    End If

                End If

            End If
        
112     Next PretoIndex
 
        
        Exit Function

AllyParalyzed_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.AllyParalyzed", Erl)
        Resume Next
        
End Function

Private Function AllyInjured(ByVal NpcIndex As Integer) As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 09/07/2010
        'Returns the index of the first found injured ally if exists one.
        '***************************************************
        
        On Error GoTo AllyInjured_Err
        
   
        Dim PretoIndex As Integer

        Dim AllyIndex  As Integer
    
100     For PretoIndex = 1 To NRO_PRETORIANOS
        
102         AllyIndex = Pretorianos(PretoIndex).NpcIndex

104         If AllyIndex <> 0 Then

106             With Npclist(AllyIndex)

108                 If .Stats.MinHp < .Stats.MaxHp Then
110                     If InVisionRangeNpc(NpcIndex, .Pos.X, .Pos.Y, False) Then
112                         AllyInjured = AllyIndex
                            Exit Function

                        End If

                    End If

                End With

            End If
        
114     Next PretoIndex
 
        
        Exit Function

AllyInjured_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.AllyInjured", Erl)
        Resume Next
        
End Function

Private Function AllyPoisoned(ByVal NpcIndex As Integer) As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 19/09/2010
        'Returns the index of the poisoned ally if exists one.
        '***************************************************
        
        On Error GoTo AllyPoisoned_Err
        
   
        Dim PretoIndex As Integer

        Dim AllyIndex  As Integer
    
100     For PretoIndex = 1 To NRO_PRETORIANOS
        
102         AllyIndex = Pretorianos(PretoIndex).NpcIndex

104         If AllyIndex <> 0 Then
106             If Npclist(AllyIndex).flags.Envenenado = 1 Then
108                 If InVisionRangeNpc(AllyIndex, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
110                     AllyPoisoned = AllyIndex
                        Exit Function

                    End If

                End If

            End If
        
112     Next PretoIndex
 
        
        Exit Function

AllyPoisoned_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.AllyPoisoned", Erl)
        Resume Next
        
End Function

Private Function AllyDead() As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 19/09/2010
        'Returns the Index of the first dead member found, if exists.
        '***************************************************
        
        On Error GoTo AllyDead_Err
        

        Dim PretoIndex As Integer

        Dim AllyIndex  As Integer
    
100     For PretoIndex = 1 To NRO_PRETORIANOS
        
102         If Pretorianos(PretoIndex).NpcIndex = 0 Then
104             AllyDead = PretoIndex
                Exit Function

            End If
        
106     Next PretoIndex

        
        Exit Function

AllyDead_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.AllyDead", Erl)
        Resume Next
        
End Function

Public Function CanAtackMember(ByVal NpcIndex As Integer) As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 26/09/2010
        'Returns true if given NpcIndex belongs to a normal clan member, o if king's alone.
        '***************************************************
        
        On Error GoTo CanAtackMember_Err
        
    
        ' King?
100     If Pretorianos(1).NpcIndex = NpcIndex Then

            ' any member left? => Can't attack
102         If ClanMembersAlive <> 1 Then Exit Function

        End If
    
104     CanAtackMember = True
    
        
        Exit Function

CanAtackMember_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.CanAtackMember", Erl)
        Resume Next
        
End Function

Public Property Get ClanMap() As Integer
        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/10/2010
        'Returns the clan map
        '***************************************************
        
        On Error GoTo ClanMap_Err
        
100     ClanMap = CenterPos.Map

        
        Exit Property

ClanMap_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.ClanMap", Erl)
        Resume Next
        
End Property

Public Property Get Active() As Boolean
        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/10/2010
        'Returns true if clan is active.
        '***************************************************
        
        On Error GoTo Active_Err
        
100     Active = ClanActive

        
        Exit Property

Active_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.Active", Erl)
        Resume Next
        
End Property

Public Sub DeleteClan()
        '***************************************************
        'Author: ZaMa
        'Last Modification: 29/10/2010
        'Kill all alive members.
        '***************************************************
        
        On Error GoTo DeleteClan_Err
        

        Dim PretoIndex As Long

        Dim NpcIndex   As Integer
    
100     For PretoIndex = 1 To NRO_PRETORIANOS
        
102         NpcIndex = Pretorianos(PretoIndex).NpcIndex

104         If NpcIndex <> 0 Then Call QuitarNPC(NpcIndex)
    
106         Pretorianos(PretoIndex).NpcIndex = 0
        
108     Next PretoIndex
    
110     ClanActive = False
    
        
        Exit Sub

DeleteClan_Err:
        Call RegistrarError(Err.Number, Err.description, "clsClanPretoriano.DeleteClan", Erl)
        Resume Next
        
End Sub


