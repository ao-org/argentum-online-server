Attribute VB_Name = "NPCs"
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
Public Const MaxRespawn             As Integer = 255
Public Const NpcIndexHeapSize       As Integer = 10000
Public RespawnList(1 To MaxRespawn) As t_Npc
Private IdNpcLibres                 As t_IndexHeap
Option Explicit

Public Sub InitializeNpcIndexHeap(Optional ByVal Size As Integer = NpcIndexHeapSize)
    On Error GoTo ErrHandler_InitizlizeNpcIndex
    ReDim IdNpcLibres.IndexInfo(Size)
    Dim i As Integer
    For i = 1 To Size
        IdNpcLibres.IndexInfo(i) = Size - (i - 1)
    Next i
    IdNpcLibres.currentIndex = Size
    Exit Sub
ErrHandler_InitizlizeNpcIndex:
    Call TraceError(Err.Number, Err.Description, "NPCs.InitializeNpcIndexHeap", Erl)
End Sub

Public Function IsValidNpcRef(ByRef NpcRef As t_NpcReference) As Boolean
    IsValidNpcRef = False
    If NpcRef.ArrayIndex < LBound(NpcList) Or NpcRef.ArrayIndex > UBound(NpcList) Then
        Exit Function
    End If
    If NpcList(NpcRef.ArrayIndex).VersionId <> NpcRef.VersionId Then
        Exit Function
    End If
    IsValidNpcRef = True
End Function

Public Function SetNpcRef(ByRef NpcRef As t_NpcReference, ByVal Index As Integer) As Boolean
    SetNpcRef = False
    NpcRef.ArrayIndex = Index
    If Index < LBound(NpcList) Or NpcRef.ArrayIndex > UBound(NpcList) Then
        Exit Function
    End If
    NpcRef.VersionId = NpcList(Index).VersionId
    SetNpcRef = True
End Function

Public Sub ClearNpcRef(ByRef NpcRef As t_NpcReference)
    NpcRef.ArrayIndex = 0
    NpcRef.VersionId = -1
End Sub

Public Sub IncreaseNpcVersionId(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex)
        If .VersionId > 32760 Then
            .VersionId = 0
        Else
            .VersionId = .VersionId + 1
        End If
    End With
End Sub

Public Function ReleaseNpc(ByVal NpcIndex As Integer, ByVal Reason As e_DeleteSource) As Boolean
    On Error GoTo ErrHandler
    If Not NpcList(NpcIndex).flags.NPCActive Then
        Call TraceError(Err.Number, "Trying to release the id twice, last reset reason: " & NpcList(NpcIndex).LastReset & " current reason " & Reason, "NPCs.ReleaseNpc", Erl)
        ReleaseNpc = False
        Exit Function
    End If
    NpcList(NpcIndex).flags.NPCActive = False
    NpcList(NpcIndex).LastReset = Reason
    Call IncreaseNpcVersionId(NpcIndex)
    IdNpcLibres.currentIndex = IdNpcLibres.currentIndex + 1
    Debug.Assert IdNpcLibres.currentIndex <= NpcIndexHeapSize
    IdNpcLibres.IndexInfo(IdNpcLibres.currentIndex) = NpcIndex
    ReleaseNpc = True
    Exit Function
ErrHandler:
    ReleaseNpc = False
    Call TraceError(Err.Number, Err.Description, "NPCs.ReleaseNpc", Erl)
End Function

Public Function GetAvailableNpcIndex() As Integer
    GetAvailableNpcIndex = IdNpcLibres.currentIndex
End Function

Public Function GetNextAvailableNpc() As Integer
    On Error GoTo ErrHandler
    If (IdNpcLibres.currentIndex = 0) Then
        GetNextAvailableNpc = 0
        Exit Function
    End If
    GetNextAvailableNpc = IdNpcLibres.IndexInfo(IdNpcLibres.currentIndex)
    IdNpcLibres.currentIndex = IdNpcLibres.currentIndex - 1
    If NpcList(GetNextAvailableNpc).flags.NPCActive Then
        Call TraceError(Err.Number, "Trying to active the same id twice", "NPCs.GetNextAvailableNpc", Erl)
    End If
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "NPCs.GetNextAvailableNpc", Erl)
End Function

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
    On Error GoTo QuitarMascotaNpc_Err
    NpcList(Maestro).Mascotas = NpcList(Maestro).Mascotas - 1
    Exit Sub
QuitarMascotaNpc_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.QuitarMascotaNpc", Erl)
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim MiNPC       As t_Npc
    Dim EraCriminal As Byte
    Dim TiempoRespw As Long
    Dim i           As Long, j As Long
    Dim Indice      As Integer
    ' Objetivo de pruebas nunca muere
    If NpcList(NpcIndex).npcType = DummyTarget Then
        Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageChatOverHead("¡¡Auch!!", NpcList(NpcIndex).Char.charindex, vbRed))
        If UBound(NpcList(NpcIndex).Char.Animation) > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(NpcList(NpcIndex).Char.charindex, NpcList(NpcIndex).Char.Animation(1)))
        End If
        NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MaxHp
        NpcList(NpcIndex).Stats.shield = 0
        Exit Sub
    End If
    MiNPC = NpcList(NpcIndex)
    TiempoRespw = NpcList(NpcIndex).Contadores.IntervaloRespawn
    ' Es NPC de la invasión?
    If MiNPC.flags.InvasionIndex Then
        Call MuereNpcInvasion(MiNPC.flags.InvasionIndex, MiNPC.flags.IndexInInvasion)
    End If
    If NpcList(NpcIndex).ShowKillerConsole > 0 Then
        'Msg1986=¬1 ha muerto en manos de ¬2
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg("1986", NpcList(NpcIndex).Name & "¬" & UserList(UserIndex).Name, e_FontTypeNames.FONTTYPE_GLOBAL))
    End If
    'Quitamos el npc
    If MiNPC.flags.IsGlobalQuestBoss Then
        GlobalQuestInfo(MiNPC.flags.IsGlobalQuestBoss).IsBossAlive = False
    End If
    Call QuitarNPC(NpcIndex, eDie)
    If UserIndex > 0 Then ' Lo mato un usuario?
        If MiNPC.flags.Snd3 > 0 Then
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.pos.x, MiNPC.pos.y))
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("28", MiNPC.pos.x, MiNPC.pos.y))
        End If
        Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
        UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
        ' El user que lo mato tiene mascotas?
        If UserList(UserIndex).NroMascotas > 0 Then
            ' Me fijo si alguna de sus mascotas le estaba pegando al NPC
            For i = 1 To UBound(UserList(UserIndex).MascotasIndex)
                If UserList(UserIndex).MascotasIndex(i).ArrayIndex > 0 Then
                    If IsValidNpcRef(UserList(UserIndex).MascotasIndex(i)) Then
                        If NpcList(UserList(UserIndex).MascotasIndex(i).ArrayIndex).TargetNPC.ArrayIndex = NpcIndex Then
                            Call FollowAmo(UserList(UserIndex).MascotasIndex(i).ArrayIndex)
                        End If
                    Else
                        Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
                    End If
                End If
            Next
        End If
        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteLocaleMsg(UserIndex, 184, e_FontTypeNames.FONTTYPE_DIOS)
        End If
        Call IncrementLongCounter(UserList(UserIndex).Stats.NPCsMuertos, "NPCsMuertos")
        If IsValidUserRef(MiNPC.MaestroUser) Then Exit Sub
        Call SubirSkill(UserIndex, e_Skill.Supervivencia)
        If MiNPC.flags.ExpCount > 0 Then
            If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + MiNPC.flags.ExpCount
                If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
                Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(MiNPC.flags.ExpCount), UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, RGB(0, 169, 255))
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
            End If
            MiNPC.flags.ExpCount = 0
        End If
        EraCriminal = Status(UserIndex)
        If MiNPC.GiveEXPClan > 0 Then
            If UserList(UserIndex).GuildIndex > 0 Then
                Call modGuilds.CheckClanExp(UserIndex, MiNPC.GiveEXPClan * SvrConfig.GetValue("ExpMult"))
            End If
        End If
        For i = 1 To MAXUSERQUESTS
            With UserList(UserIndex).QuestStats.Quests(i)
                If .QuestIndex Then
                    If QuestList(.QuestIndex).RequiredNPCs Then
                        For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                            If QuestList(.QuestIndex).RequiredNPC(j).NpcIndex = MiNPC.Numero Then
                                If QuestList(.QuestIndex).RequiredNPC(j).amount >= .NPCsKilled(j) Then
                                    .NPCsKilled(j) = .NPCsKilled(j) + 1 '
                                    Call WriteLocaleMsg(UserIndex, 1623, e_FontTypeNames.FONTTYPE_INFOIAO, MiNPC.name & "¬" & .NPCsKilled(j) & "¬" & QuestList( _
                                            .QuestIndex).RequiredNPC(j).Amount) 'Msg1623=¬1 matados/as: ¬2 de ¬3
                                    Call WriteChatOverHead(UserIndex, "NOCONSOLA*" & .NPCsKilled(j) & "/" & QuestList(.QuestIndex).RequiredNPC(j).amount & " " & MiNPC.name, _
                                            UserList(UserIndex).Char.charindex, RGB(180, 180, 180))
                                Else
                                    Call WriteLocaleMsg(UserIndex, 1624, e_FontTypeNames.FONTTYPE_INFOIAO, MiNPC.name & "¬" & QuestList(.QuestIndex).nombre) 'Msg1624=Ya has matado todos los ¬1 que la misión ¬2 requería. Revisa si ya estás listo para recibir la recompensa.
                                    Call WriteChatOverHead(UserIndex, "NOCONSOLA*" & QuestList(.QuestIndex).RequiredNPC(j).amount & "/" & QuestList(.QuestIndex).RequiredNPC( _
                                            j).amount & " " & MiNPC.name, UserList(UserIndex).Char.charindex, RGB(180, 180, 180))
                                End If
                            End If
                        Next j
                    End If
                    UserList(UserIndex).flags.ModificoQuests = True
                End If
            End With
        Next i
        'Tiramos el oro
        Call NPCTirarOro(MiNPC, UserIndex)
        Call DropObjQuest(MiNPC, UserIndex)
        'Item Magico!
        Call NpcDropeo(MiNPC, UserIndex)
        Call DropFromGlobalDropTable(MiNPC, UserIndex)
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC)
    End If ' UserIndex > 0
    ' Mascotas y npcs de entrenamiento no respawnean
    If MiNPC.MaestroNPC.ArrayIndex > 0 Or IsValidUserRef(MiNPC.MaestroUser) Then Exit Sub
    If NpcIndex = npc_index_evento Then
        BusquedaNpcActiva = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg("1549", vbNullString, e_FontTypeNames.FONTTYPE_CITIZEN)) ' Msg1549=Evento> El NPC ha sido asesinado.
        npc_index_evento = 0
    End If
    'ReSpawn o no
    If TiempoRespw = 0 Then
        Call ReSpawnNpc(MiNPC)
    Else
        MiNPC.flags.NPCActive = True
        Indice = ObtenerIndiceRespawn
        RespawnList(Indice) = MiNPC
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description & "->" & Erl(), "NPCs.MuereNpc", Erl())
End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags
    On Error GoTo ResetNpcFlags_Err
    With NpcList(NpcIndex).flags
        .AfectaParalisis = 0
        .ImmuneToSpells = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedTime = 0
        .AttackedFirstBy = vbNullString
        .backup = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .LanzaSpells = 0
        .GolpeExacto = 0
        .invisible = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .SndRespawn = 0
        .TierraInvalida = 0
        Call ResetMask(.BehaviorFlags)
        Call SetMask(.BehaviorFlags, e_BehaviorFlags.eAttackNpc Or e_BehaviorFlags.eAttackUsers)
        .AIAlineacion = e_Alineacion.ninguna
        .NPCIdle = False
        Call ClearNpcRef(.Summoner)
    End With
    Exit Sub
ResetNpcFlags_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcFlags", Erl)
End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)
    On Error GoTo ResetNpcCounters_Err
    NpcList(NpcIndex).Contadores.Paralisis = 0
    NpcList(NpcIndex).Contadores.TiempoExistencia = 0
    NpcList(NpcIndex).Contadores.IntervaloMovimiento = 0
    NpcList(NpcIndex).Contadores.IntervaloAtaque = 0
    NpcList(NpcIndex).Contadores.IntervaloLanzarHechizo = 0
    NpcList(NpcIndex).Contadores.IntervaloRespawn = 0
    NpcList(NpcIndex).Contadores.CriaturasInvocadas = 0
    Exit Sub
ResetNpcCounters_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcCounters", Erl)
End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
    On Error GoTo ResetNpcCharInfo_Err
    NpcList(NpcIndex).Char.body = 0
    NpcList(NpcIndex).Char.CascoAnim = 0
    NpcList(NpcIndex).Char.charindex = 0
    NpcList(NpcIndex).Char.FX = 0
    NpcList(NpcIndex).Char.head = 0
    NpcList(NpcIndex).Char.Heading = 0
    NpcList(NpcIndex).Char.loops = 0
    NpcList(NpcIndex).Char.ShieldAnim = 0
    NpcList(NpcIndex).Char.WeaponAnim = 0
    NpcList(NpcIndex).Char.CartAnim = 0
    Exit Sub
ResetNpcCharInfo_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcCharInfo", Erl)
End Sub

Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
    On Error GoTo ResetNpcCriatures_Err
    Dim j As Integer
    For j = 1 To NpcList(NpcIndex).NroCriaturas
        NpcList(NpcIndex).Criaturas(j).NpcIndex = 0
        NpcList(NpcIndex).Criaturas(j).NpcName = vbNullString
    Next j
    NpcList(NpcIndex).NroCriaturas = 0
    Exit Sub
ResetNpcCriatures_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcCriatures", Erl)
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
    On Error GoTo ResetExpresiones_Err
    Dim j As Integer
    For j = 1 To NpcList(NpcIndex).NroExpresiones
        NpcList(NpcIndex).Expresiones(j) = vbNullString
    Next j
    NpcList(NpcIndex).NroExpresiones = 0
    Exit Sub
ResetExpresiones_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.ResetExpresiones", Erl)
End Sub

Sub ResetDrop(ByVal NpcIndex As Integer)
    On Error GoTo ResetDrop_Err
    Dim j As Integer
    For j = 1 To NpcList(NpcIndex).NumQuiza
        NpcList(NpcIndex).QuizaDropea(j) = 0
    Next j
    NpcList(NpcIndex).NumQuiza = 0
    Exit Sub
ResetDrop_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.ResetDrop", Erl)
End Sub

Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
    On Error GoTo ResetNpcMainInfo_Err
    With (NpcList(NpcIndex))
        .Attackable = 0
        .pathFindingInfo.TargetUnreachable = False
        .pathFindingInfo.PreviousAttackable = 0
        .pathFindingInfo.PathLength = 0
        .pathFindingInfo.StrafeOffset.x = 0
        .pathFindingInfo.StrafeOffset.y = 0
        .pathFindingInfo.StrafeExpiresAt = 0
        .pathFindingInfo.OrbitDirection = 0
        .pathFindingInfo.OrbitReevaluateAt = 0
        .pathFindingInfo.NextPathRecomputeAt = 0
        .Comercia = 0
        .GiveEXP = 0
        .GiveEXPClan = 0
        .GiveGLD = 0
        .Hostile = 0
        .InvReSpawn = 0
        .nivel = 0
        Call ClearNpcRef(.MaestroNPC)
        .Mascotas = 0
        .Movement = 0
        .name = "NPC SIN INICIAR"
        .npcType = 0
        .Numero = 0
        .Orig.Map = 0
        .Orig.x = 0
        .Orig.y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .pos.Map = 0
        .pos.x = 0
        .pos.y = 0
        Call SetUserRef(.TargetUser, 0)
        Call ClearNpcRef(.TargetNPC)
        .TipoItems = 0
        .Veneno = 0
        .Desc = vbNullString
        .NumDropQuest = 0
        If IsValidUserRef(.MaestroUser) Then Call QuitarMascota(.MaestroUser.ArrayIndex, NpcIndex)
        If IsValidNpcRef(.MaestroNPC) Then Call QuitarMascotaNpc(.MaestroNPC.ArrayIndex)
        Call SetUserRef(.MaestroUser, 0)
        Call ClearNpcRef(.MaestroNPC)
        .CaminataActual = 0
        Dim j As Integer
        For j = 1 To .flags.LanzaSpells
            .Spells(j).SpellIndex = 0
        Next j
        Call ClearEffectList(.EffectOverTime)
        Call ClearModifiers(.Modifiers)
    End With
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)
    Call ResetDrop(NpcIndex)
    Exit Sub
ResetNpcMainInfo_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcMainInfo", Erl)
End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer, ByVal releaseReason As e_DeleteSource)
    On Error GoTo ErrHandler
    If Not ReleaseNpc(NpcIndex, releaseReason) Then
        Exit Sub
    End If
    If IsValidNpcRef(NpcList(NpcIndex).flags.Summoner) Then
        If NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Contadores.CriaturasInvocadas > 0 Then
            'Resto 1 Npc invocado al invocador
            NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Contadores.CriaturasInvocadas = NpcList(NpcList( _
                    NpcIndex).flags.Summoner.ArrayIndex).Contadores.CriaturasInvocadas - 1
            'También lo saco de la lista
            Dim LoopC As Long
            For LoopC = 1 To NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Stats.CantidadInvocaciones
                If NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Stats.NpcsInvocados(LoopC).ArrayIndex = NpcIndex Then
                    Call ClearNpcRef(NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Stats.NpcsInvocados(LoopC))
                    Exit For
                End If
            Next LoopC
        End If
    ElseIf NpcList(NpcIndex).Contadores.CriaturasInvocadas > 0 Then
        Dim i As Long
        For i = 1 To NpcList(NpcIndex).Stats.CantidadInvocaciones
            If IsValidNpcRef(NpcList(NpcIndex).Stats.NpcsInvocados(i)) Then
                Call MuereNpc(NpcList(NpcIndex).Stats.NpcsInvocados(i).ArrayIndex, 0)
            End If
        Next i
    End If
    If InMapBounds(NpcList(NpcIndex).pos.Map, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y) Then
        Call EraseNPCChar(NpcIndex)
    End If
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    Call ResetNpcMainInfo(NpcIndex)
    If NpcIndex = LastNPC Then
        Do Until NpcList(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If
    Exit Sub
ErrHandler:
    Call LogError("Error en QuitarNPC")
End Sub

Function TestSpawnTrigger(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo TestSpawnTrigger_Err
    TestSpawnTrigger = MapData(Map, x, y).trigger < 1 Or (MapData(Map, x, y).trigger > 3 And MapData(Map, x, y).trigger < 12)
    Exit Function
TestSpawnTrigger_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.TestSpawnTrigger", Erl)
End Function

Public Function CrearNPC(NroNPC As Integer, Mapa As Integer, OrigPos As t_WorldPos, Optional ByVal CustomHead As Integer)
    'Crea un NPC del tipo NRONPC
    On Error GoTo CrearNPC_Err
    Dim NpcIndex    As Integer
    Dim Iteraciones As Long
    Dim PuedeAgua   As Boolean
    Dim PuedeTierra As Boolean
    Dim Map         As Integer
    Dim x           As Integer
    Dim y           As Integer
    NpcIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    If NpcIndex = 0 Then Exit Function
    With NpcList(NpcIndex)
        ' Cabeza customizada
        If CustomHead <> 0 Then .Char.head = CustomHead
        PuedeAgua = .flags.AguaValida = 1
        PuedeTierra = .flags.TierraInvalida = 0
        'Necesita ser respawned en un lugar especifico
        If .flags.RespawnOrigPos And InMapBounds(OrigPos.Map, OrigPos.x, OrigPos.y) Then
            Map = OrigPos.Map
            x = OrigPos.x
            y = OrigPos.y
            .Orig = OrigPos
            .pos = OrigPos
        Else
            ' Primera búsqueda: buscamos una posición ideal hasta llegar al máximo de iteraciones
            Do
                .pos.Map = Mapa
                .pos.x = RandomNumber(MinXBorder + 2, MaxXBorder - 2) 'Obtenemos posicion al azar en x
                .pos.y = RandomNumber(MinYBorder + 2, MaxYBorder - 2) 'Obtenemos posicion al azar en y
                .pos = ClosestLegalPosNPC(NpcIndex, 10, , True)     'Nos devuelve la posicion valida mas cercana
                Iteraciones = Iteraciones + 1
            Loop While .pos.x = 0 And .pos.y = 0 And Iteraciones < MAXSPAWNATTEMPS
            ' Si no encontramos una posición válida en la primera instancia
            If Iteraciones >= MAXSPAWNATTEMPS Then
                ' Hacemos una búsqueda exhaustiva partiendo desde el centro del mapa
                .pos.Map = Mapa
                .pos.x = (XMaxMapSize - XMinMapSize) \ 2
                .pos.y = (YMaxMapSize - YMinMapSize) \ 2
                .pos = ClosestLegalPosNPC(NpcIndex, (XMaxMapSize - XMinMapSize) \ 2, , True)
                ' Si sigue fallando
                If .pos.x = 0 And .pos.y = 0 Then
                    ' Hacemos una última búsqueda exhaustiva, ignorando los usuarios
                    .pos.Map = Mapa
                    .pos.x = (XMaxMapSize - XMinMapSize) \ 2
                    .pos.y = (YMaxMapSize - YMinMapSize) \ 2
                    .pos = ClosestLegalPosNPC(NpcIndex, (XMaxMapSize - XMinMapSize) \ 2, True)
                    ' Si falló, borramos el NPC y salimos
                    If .pos.x = 0 And .pos.y = 0 Then
                        Call QuitarNPC(NpcIndex, eFailToFindSpawnPos)
                        Exit Function
                    End If
                End If
            End If
            'asignamos las nuevas coordenas
            Map = .pos.Map
            x = .pos.x
            y = .pos.y
            'Y tambien asignamos su posicion original, para tener una posicion de retorno.
            .Orig.Map = .pos.Map
            .Orig.x = .pos.x
            .Orig.y = .pos.y
        End If
    End With
    'Crea el NPC
    Call MakeNPCChar(True, Map, NpcIndex, Map, x, y)
    CrearNPC = NpcIndex
    Exit Function
CrearNPC_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.CrearNPC", Erl)
End Function

Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo MakeNPCChar_Err
    With NpcList(NpcIndex)
        Dim charindex As Integer
        If .Char.charindex = 0 Then
            charindex = NextOpenCharIndex
            .Char.charindex = charindex
            CharList(charindex) = NpcIndex
        End If
        MapData(Map, x, y).NpcIndex = NpcIndex
        Dim Simbolo As Byte
        Dim GG      As String
        Dim tmpByte As Byte
        GG = IIf(.showName > 0, .name & .SubName, vbNullString)
        If Not toMap Then
            If .NumQuest > 0 Then
                Dim q             As Byte
                Dim HayFinalizada As Boolean
                Dim HayDisponible As Boolean
                Dim HayPendiente  As Boolean
                For q = 1 To .NumQuest
                    tmpByte = TieneQuest(sndIndex, .QuestNumber(q))
                    If tmpByte Then
                        If FinishQuestCheck(sndIndex, .QuestNumber(q), tmpByte) Then
                            Simbolo = 3
                            HayFinalizada = True
                        Else
                            HayPendiente = True
                            Simbolo = 4
                        End If
                    Else
                        If UserDoneQuest(sndIndex, .QuestNumber(q)) Or Not UserDoneQuest(sndIndex, QuestList(.QuestNumber(q)).RequiredQuest) Or UserList(sndIndex).Stats.ELV < _
                                QuestList(.QuestNumber(q)).RequiredLevel Or UserList(sndIndex).clase = QuestList(.QuestNumber(q)).RequiredClass Then
                            Simbolo = 2
                        Else
                            Simbolo = 1
                            HayDisponible = True
                        End If
                    End If
                Next q
                'Para darle prioridad a ciertos simbolos
                If HayDisponible Then
                    Simbolo = 1
                End If
                If HayPendiente Then
                    Simbolo = 4
                End If
                If HayFinalizada Then
                    Simbolo = 3
                End If
                'Para darle prioridad a ciertos simbolos
            End If
            Dim body As Integer
            'Si está muerto el usuario y en zona insegura
            If UserList(sndIndex).flags.Muerto = 1 And MapInfo(UserList(sndIndex).pos.Map).Seguro = 0 Then
                'Solamente mando el body si es de tipo revividor.
                If .npcType = e_NPCType.Revividor Then
                    body = IIf(.flags.NPCIdle, .Char.BodyIdle, .Char.body)
                Else
                    body = 0
                End If
            Else
                body = IIf(.flags.NPCIdle, .Char.BodyIdle, .Char.body)
            End If
            If UserList(sndIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 90 Then
                Call WriteCharacterCreate(sndIndex, body, .Char.head, .Char.Heading, .Char.charindex, x, y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, _
                        .Char.CartAnim, 0, GG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .Char.speeding, IIf(.MaestroUser.ArrayIndex = sndIndex, 2, 1), 0, 0, 0, 0, .Stats.MinHp, _
                        .Stats.MaxHp, 0, 0, Simbolo, .flags.NPCIdle, , , .flags.team, , .Char.Ataque1)
            Else
                Call WriteCharacterCreate(sndIndex, body, .Char.head, .Char.Heading, .Char.charindex, x, y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, _
                        .Char.CartAnim, 0, GG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .Char.speeding, IIf(.MaestroUser.ArrayIndex = sndIndex, 2, 1), 0, 0, 0, 0, 0, 0, 0, 0, Simbolo, _
                        .flags.NPCIdle, , , .flags.team, , .Char.Ataque1)
            End If
            If IsSet(.flags.StatusMask, e_StatusMask.eDontBlockTile) Then
                Call SendData(ToIndex, sndIndex, PrepareUpdateCharValue(.Char.charindex, e_CharValue.eDontBlockTile, True))
            End If
        Else
            Call AgregarNpc(NpcIndex)
        End If
    End With
    Exit Sub
MakeNPCChar_Err:
    Dim errNumber As Long
    Dim errDescription As String
    Dim contextInfo As String

    errNumber = Err.Number
    errDescription = Err.Description

    contextInfo = "Params: toMap=" & CStr(toMap) & ", sndIndex=" & sndIndex & ", NpcIndex=" & NpcIndex & ", Map=" & Map & _
                  ", x=" & x & ", y=" & y

    Dim npcLowerBound As Long
    Dim npcUpperBound As Long

    On Error Resume Next
    npcLowerBound = LBound(NpcList)
    npcUpperBound = UBound(NpcList)

    If NpcIndex >= npcLowerBound And NpcIndex <= npcUpperBound Then
        contextInfo = contextInfo & ", NpcName=" & NpcList(NpcIndex).name & NpcList(NpcIndex).SubName
        contextInfo = contextInfo & ", CharIndex=" & NpcList(NpcIndex).Char.charindex
        contextInfo = contextInfo & ", NPCType=" & NpcList(NpcIndex).npcType
    Else
        contextInfo = contextInfo & ", NpcIndexOutOfBounds=True (Bounds " & npcLowerBound & "-" & npcUpperBound & ")"
    End If

    On Error GoTo 0

    Call TraceError(errNumber, errDescription & " | " & contextInfo, "NPCs.MakeNPCChar", Erl)
End Sub

Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal body As Integer, ByVal head As Integer, ByVal Heading As e_Heading)
    On Error GoTo ChangeNPCChar_Err
    With NpcList(NpcIndex)
        If NpcIndex > 0 Then
            If .flags.NPCIdle Then
                body = .Char.BodyIdle
            End If
            .Char.head = head
            .Char.Heading = Heading
            If .Char.charindex > 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageCharacterChange(body, head, Heading, .Char.charindex, 0, 0, 0, 0, 0, 0, 0, .flags.NPCIdle, False))
            End If
        End If
    End With
    Exit Sub
ChangeNPCChar_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.ChangeNPCChar", Erl)
End Sub

Sub EraseNPCChar(ByVal NpcIndex As Integer)
    On Error GoTo EraseNPCChar_Err
    If NpcList(NpcIndex).Char.charindex <> 0 Then CharList(NpcList(NpcIndex).Char.charindex) = 0
    If NpcList(NpcIndex).Char.charindex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    Call RemoveNpc(NpcIndex)
    'Quitamos del mapa
    MapData(NpcList(NpcIndex).pos.Map, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y).NpcIndex = 0
    'Actualizamos los clientes
    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(5, NpcList(NpcIndex).Char.charindex, True))
    'Update la lista npc
    NpcList(NpcIndex).Char.charindex = 0
    'update NumChars
    NumChars = NumChars - 1
    Exit Sub
EraseNPCChar_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.EraseNPCChar", Erl)
End Sub

Public Sub TranslateNpcChar(ByVal NpcIndex As Integer, ByRef NewPos As t_WorldPos, ByVal Speed As Long)
    On Error GoTo TranslateNpcChar_Err
    With NpcList(NpcIndex)
        If MapData(.pos.Map, NewPos.x, NewPos.y).UserIndex Then
            Call SwapTargetUserPos(MapData(.pos.Map, NewPos.x, NewPos.y).UserIndex, .pos)
        End If
        'Update map and user pos
        MapData(.pos.Map, .pos.x, .pos.y).NpcIndex = 0
        Dim PrevPos As t_WorldPos
        PrevPos = .pos
        .pos = NewPos
        MapData(.pos.Map, NewPos.x, NewPos.y).NpcIndex = NpcIndex
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareCharacterTranslate(.Char.charindex, NewPos.x, NewPos.y, Speed))
        Call CheckUpdateNeededNpc(NpcIndex, GetHeadingFromWorldPos(PrevPos, NewPos))
    End With
    Exit Sub
TranslateNpcChar_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.TranslateNpcChar", Erl)
End Sub

Public Function MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte) As Boolean
    On Error GoTo errh
    Dim nPos      As t_WorldPos
    Dim UserIndex As Integer
    Dim esGuardia As Boolean
    With NpcList(NpcIndex)
        If Not NPCs.CanMove(.Contadores, .flags) Then Exit Function
        nPos = .pos
        Call HeadtoPos(nHeading, nPos)
        esGuardia = .npcType = e_NPCType.GuardiaReal Or .npcType = e_NPCType.GuardiasCaos
        ' es una posicion legal
        If LegalWalkNPC(nPos.Map, nPos.x, nPos.y, nHeading, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IsValidUserRef(.MaestroUser), , esGuardia) Then
            UserIndex = MapData(.pos.Map, nPos.x, nPos.y).UserIndex
            ' Si hay un usuario a donde se mueve el npc, entonces esta muerto o es un gm invisible
            If UserIndex > 0 Then
                With UserList(UserIndex)
                    ' Actualizamos posicion y mapa
                    MapData(.pos.Map, .pos.x, .pos.y).UserIndex = 0
                    .pos.x = NpcList(NpcIndex).pos.x
                    .pos.y = NpcList(NpcIndex).pos.y
                    MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex
                    ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.charindex, .pos.x, .pos.y))
                    Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))
                End With
            End If
            ' Solo NPCs hum
            If NpcList(NpcIndex).Humanoide Or NpcList(NpcIndex).npcType = e_NPCType.GuardiaReal Or NpcList(NpcIndex).npcType = e_NPCType.GuardiasCaos Or NpcList( _
                    NpcIndex).npcType = e_NPCType.GuardiaNpc Then
                If HayPuerta(nPos.Map, nPos.x, nPos.y) Then
                    Call AccionParaPuertaNpc(nPos.Map, nPos.x, nPos.y, NpcIndex)
                ElseIf HayPuerta(nPos.Map, nPos.x + 1, nPos.y) Then
                    Call AccionParaPuertaNpc(nPos.Map, nPos.x + 1, nPos.y, NpcIndex)
                ElseIf HayPuerta(nPos.Map, nPos.x + 1, nPos.y - 1) Then
                    Call AccionParaPuertaNpc(nPos.Map, nPos.x + 1, nPos.y - 1, NpcIndex)
                ElseIf HayPuerta(nPos.Map, nPos.x, nPos.y - 1) Then
                    Call AccionParaPuertaNpc(nPos.Map, nPos.x, nPos.y - 1, NpcIndex)
                End If
            End If
            Call AnimacionIdle(NpcIndex, False)
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.x, nPos.y))
            'Update map and user pos
            MapData(.pos.Map, .pos.x, .pos.y).NpcIndex = 0
            .pos = nPos
            .Char.Heading = nHeading
            MapData(.pos.Map, nPos.x, nPos.y).NpcIndex = NpcIndex
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
            If Not MapData(.pos.Map, nPos.x, nPos.y).Trap Is Nothing Then
                Call ModMap.ActivateTrap(NpcIndex, eNpc, .pos.Map, nPos.x, nPos.y)
            End If
            ' Npc has moved
            MoveNPCChar = True
        End If
    End With
    Exit Function
errh:
    LogError ("Error en move npc " & NpcIndex & ". Error: " & Err.Number & " - " & Err.Description)
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer, ByVal VenenoNivel As Byte)
    On Error GoTo NpcEnvenenarUser_Err
    Dim n As Integer
    n = RandomNumber(1, 100)
    If n < 30 Then
        UserList(UserIndex).flags.Envenenado = VenenoNivel
        'Msg182=¡¡La criatura te ha envenenado!!
        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteLocaleMsg(UserIndex, 182, e_FontTypeNames.FONTTYPE_FIGHT)
        End If
    End If
    Exit Sub
NpcEnvenenarUser_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.NpcEnvenenarUser", Erl)
End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, _
                  pos As t_WorldPos, _
                  ByVal FX As Boolean, _
                  ByVal Respawn As Boolean, _
                  Optional Avisar As Boolean = False, _
                  Optional ByVal MaestroUser As Integer = 0, _
                  Optional ByVal SpellWav As Integer = SND_WARP) As Integer
    On Error GoTo SpawnNpc_Err
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 23/01/2007
    '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
    '***************************************************
    Dim NewPos      As t_WorldPos
    Dim altpos      As t_WorldPos
    Dim nIndex      As Integer
    Dim PuedeAgua   As Boolean
    Dim PuedeTierra As Boolean
    Dim Map         As Integer
    Dim x           As Integer
    Dim y           As Integer
    nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice
    If nIndex = 0 Then
        SpawnNpc = 0
        Exit Function
    End If
    PuedeAgua = NpcList(nIndex).flags.AguaValida = 1
    PuedeTierra = NpcList(nIndex).flags.TierraInvalida = 0
    Call ClosestLegalPos(pos, NewPos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
    'Si X e Y son iguales a 0 significa que no se encontro posicion valida
    If NewPos.x <> 0 And NewPos.y <> 0 Then
        'Asignamos las nuevas coordenas solo si son validas
        NpcList(nIndex).pos.Map = NewPos.Map
        NpcList(nIndex).pos.x = NewPos.x
        NpcList(nIndex).pos.y = NewPos.y
    Else
        Call QuitarNPC(nIndex, eFailToFindSpawnPos)
        SpawnNpc = 0
        Exit Function
    End If
    'asignamos las nuevas coordenas
    Map = NewPos.Map
    x = NpcList(nIndex).pos.x
    y = NpcList(nIndex).pos.y
    '  Asignamos el dueño
    Call SetUserRef(NpcList(nIndex).MaestroUser, MaestroUser)
    NpcList(nIndex).Orig = NpcList(nIndex).pos
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, x, y)
    If FX And SpellWav = SND_WARP Then
        Call SendData(SendTarget.ToNPCAliveArea, nIndex, PrepareMessagePlayWave(SND_WARP, x, y))
        Call SendData(SendTarget.ToNPCAliveArea, nIndex, PrepareMessageCreateFX(NpcList(nIndex).Char.charindex, e_GraphicEffects.ModernGmWarp, 0))
    Else
        Call SendData(SendTarget.ToNPCAliveArea, nIndex, PrepareMessageCreateFX(NpcList(nIndex).Char.charindex, e_GraphicEffects.ModernGmWarp, 0))
    End If
    If Avisar Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg("1548", NpcList(nIndex).name & "¬" & get_map_name(Map), e_FontTypeNames.FONTTYPE_CITIZEN)) '  Msg1548=¬1 ha aparecido en ¬2, todo indica que puede tener una gran recompensa para el que logre sobrevivir a él.
    End If
    SpawnNpc = nIndex
    Exit Function
SpawnNpc_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.SpawnNpc", Erl)
End Function

Sub ReSpawnNpc(MiNPC As t_Npc)
    On Error GoTo ReSpawnNpc_Err
    If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.pos.Map, MiNPC.Orig)
    Exit Sub
ReSpawnNpc_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.ReSpawnNpc", Erl)
End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As Integer
    On Error GoTo NPCHostiles_Err
    Dim NpcIndex As Integer
    Dim cont     As Integer
    'Contador
    cont = 0
    For NpcIndex = 1 To LastNPC
        '¿esta vivo?
        If NpcList(NpcIndex).flags.NPCActive And NpcList(NpcIndex).pos.Map = Map And NpcList(NpcIndex).Hostile = 1 Then
            cont = cont + 1
        End If
    Next NpcIndex
    NPCHostiles = cont
    Exit Function
NPCHostiles_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.NPCHostiles", Erl)
End Function

Sub NPCTirarOro(MiNPC As t_Npc, ByVal UserIndex As Integer)
    On Error GoTo NPCTirarOro_Err
    If UserIndex = 0 Then Exit Sub
    If MiNPC.GiveGLD > 0 Then
        Dim Oro As Long
        Oro = MiNPC.GiveGLD * SvrConfig.GetValue("GoldMult")
        Dim MiObj As t_Obj
        MiObj.ObjIndex = iORO
        While (Oro > 0)
            If Oro > MAX_INVENTORY_OBJS Then
                MiObj.amount = MAX_INVENTORY_OBJS
                Oro = Oro - MAX_INVENTORY_OBJS
            Else
                MiObj.amount = Oro
                Oro = 0
            End If
            Call TirarItemAlPiso(MiNPC.pos, MiObj, MiNPC.flags.AguaValida = 1)
        Wend
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageFxPiso("87", MiNPC.pos.x, MiNPC.pos.y))
    End If
    Exit Sub
NPCTirarOro_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.NPCTirarOro", Erl)
End Sub

Function UpdateNpcSpeed(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex)
        If .IntervaloMovimiento = 0 Then
            .IntervaloMovimiento = 380
            .Char.speeding = IntervaloNPCAI / 330
        Else
            .Char.speeding = 210 / .IntervaloMovimiento
        End If
        .Char.speeding = .Char.speeding * max(0, (1 + .Modifiers.MovementSpeed))
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSpeedingACT(.Char.charindex, .Char.speeding))
    End With
End Function

Function GetNpcSpeedModifiers(ByVal NpcIndex As Integer) As Single
    GetNpcSpeedModifiers = max(0, (1 + NpcList(NpcIndex).Modifiers.MovementSpeed))
End Function

Function GetNpcName(ByVal NpcNumber As Integer) As String
    On Error GoTo ErrHandler
    If NpcInfoCacheInitialized Then
        If NpcNumber >= LBound(NpcInfoCache) And NpcNumber <= UBound(NpcInfoCache) Then
            If NpcInfoCache(NpcNumber).Exists Then
                GetNpcName = NpcInfoCache(NpcNumber).name
                Exit Function
            End If
        End If
    End If
    If Not LeerNPCs Is Nothing Then
        GetNpcName = LeerNPCs.GetValue("NPC" & NpcNumber, "Name")
    End If
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "NPCs.GetNpcName", Erl)
End Function

Public Sub BuildNpcInfoCache()
    On Error GoTo ErrHandler
    If LeerNPCs Is Nothing Then Exit Sub
    NpcInfoCacheInitialized = False
    ReDim NpcInfoCache(1 To MaxNPCs)
    Dim i As Long
    For i = 1 To MaxNPCs
        NpcInfoCache(i).Exists = False
    Next i
    Dim nodeCount As Long
    nodeCount = LeerNPCs.NodesCount
    For i = 0 To nodeCount - 1
        Dim nodeName As String
        nodeName = LeerNPCs.GetNode(i)
        If Left$(nodeName, 3) = "NPC" Then
            Dim npcNumber As Long
            npcNumber = Val(Mid$(nodeName, 4))
            If npcNumber >= LBound(NpcInfoCache) And npcNumber <= UBound(NpcInfoCache) Then
                Call LoadNpcInfoIntoCache(CInt(npcNumber))
            End If
        End If
    Next i
    NpcInfoCacheInitialized = True
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "NPCs.BuildNpcInfoCache", Erl)
End Sub

Private Sub LoadNpcInfoIntoCache(ByVal NpcNumber As Integer)
    On Error GoTo ErrHandler
    Dim SectionName As String
    SectionName = "NPC" & NpcNumber
    If LeerNPCs Is Nothing Then Exit Sub
    If Not LeerNPCs.KeyExists(SectionName) Then Exit Sub
    Dim LoopC As Long
    Dim ln As String
    Dim aux As String
    Dim Field() As String
    Dim AnimacionesCount As Integer
    Dim MaxInventory As Long
    Dim cant As Long
    With NpcInfoCache(NpcNumber)
        .Exists = True
        .TestOnly = Val(LeerNPCs.GetValue(SectionName, "TESTONLY"))
        .RequireToggle = LeerNPCs.GetValue(SectionName, "REQUIRETOGGLE")
        .name = LeerNPCs.GetValue(SectionName, "Name")
        .SubName = LeerNPCs.GetValue(SectionName, "SubName")
        .Desc = LeerNPCs.GetValue(SectionName, "Desc")
        .nivel = Val(LeerNPCs.GetValue(SectionName, "Nivel"))
        .Movement = Val(LeerNPCs.GetValue(SectionName, "Movement"))
        .AguaValida = Val(LeerNPCs.GetValue(SectionName, "AguaValida"))
        .TierraInvalida = Val(LeerNPCs.GetValue(SectionName, "TierraInValida"))
        .Faccion = Val(LeerNPCs.GetValue(SectionName, "Faccion"))
        .ElementalTags = Val(LeerNPCs.GetValue(SectionName, "ElementalTags"))
        .IsGlobalQuestBoss = val(LeerNPCs.GetValue(SectionName, "IsGlobalQuestBoss"))
        .npcType = Val(LeerNPCs.GetValue(SectionName, "NpcType"))
        .Body = Val(LeerNPCs.GetValue(SectionName, "Body"))
        .Head = Val(LeerNPCs.GetValue(SectionName, "Head"))
        .Heading = Val(LeerNPCs.GetValue(SectionName, "Heading"))
        .BodyIdle = Val(LeerNPCs.GetValue(SectionName, "BodyIdle"))
        .Ataque1 = Val(LeerNPCs.GetValue(SectionName, "Ataque1"))
        .CastAnimation = Val(LeerNPCs.GetValue(SectionName, "CastAnimation"))
        AnimacionesCount = Val(LeerNPCs.GetValue(SectionName, "Animaciones"))
        .AnimacionesCount = AnimacionesCount
        If AnimacionesCount > 0 Then
            ReDim .Animaciones(1 To AnimacionesCount)
            For LoopC = 1 To AnimacionesCount
                .Animaciones(LoopC) = Val(LeerNPCs.GetValue(SectionName, "Anim" & LoopC))
            Next LoopC
        Else
            Erase .Animaciones
        End If
        .WeaponAnim = Val(LeerNPCs.GetValue(SectionName, "Arma"))
        .ShieldAnim = Val(LeerNPCs.GetValue(SectionName, "Escudo"))
        .CascoAnim = Val(LeerNPCs.GetValue(SectionName, "Casco"))
        .CartAnim = Val(LeerNPCs.GetValue(SectionName, "Cart"))
        .Attackable = Val(LeerNPCs.GetValue(SectionName, "Attackable"))
        .Comercia = Val(LeerNPCs.GetValue(SectionName, "Comercia"))
        .Craftea = Val(LeerNPCs.GetValue(SectionName, "Craftea"))
        .Hostile = Val(LeerNPCs.GetValue(SectionName, "Hostile"))
        .AttackRange = Val(LeerNPCs.GetValue(SectionName, "AttackRange"))
        .ProjectileType = Val(LeerNPCs.GetValue(SectionName, "ProjectileType"))
        .PreferedRange = Val(LeerNPCs.GetValue(SectionName, "PreferedRange"))
        .GiveEXP = Val(LeerNPCs.GetValue(SectionName, "GiveEXP"))
        .Distancia = Val(LeerNPCs.GetValue(SectionName, "Distancia"))
        .GiveEXPClan = Val(LeerNPCs.GetValue(SectionName, "GiveEXPClan"))
        .Veneno = Val(LeerNPCs.GetValue(SectionName, "Veneno"))
        .Domable = Val(LeerNPCs.GetValue(SectionName, "Domable"))
        .AttackableByEveryone = Val(LeerNPCs.GetValue(SectionName, "AttackableByEveryone", 0))
        .MapEntryPrice = Val(LeerNPCs.GetValue(SectionName, "MapEntryPrice", 0))
        .MapTargetEntry = Val(LeerNPCs.GetValue(SectionName, "MapTargetEntry", 1))
        .MapTargetEntryX = Val(LeerNPCs.GetValue(SectionName, "MapTargetEntryX", 50))
        .MapTargetEntryY = Val(LeerNPCs.GetValue(SectionName, "MapTargetEntryY", 50))
        .ArenaEnabled = Val(LeerNPCs.GetValue(SectionName, "ArenaEnabled", 0))
        .GiveGLD = Val(LeerNPCs.GetValue(SectionName, "GiveGLD"))
        .PoderAtaque = Val(LeerNPCs.GetValue(SectionName, "PoderAtaque"))
        .PoderEvasion = Val(LeerNPCs.GetValue(SectionName, "PoderEvasion"))
        .InvReSpawn = Val(LeerNPCs.GetValue(SectionName, "InvReSpawn"))
        .ShowName = Val(LeerNPCs.GetValue(SectionName, "ShowName"))
        .GobernadorDe = Val(LeerNPCs.GetValue(SectionName, "GobernadorDe"))
        .SoundOpen = Val(LeerNPCs.GetValue(SectionName, "SoundOpen"))
        .SoundClose = Val(LeerNPCs.GetValue(SectionName, "SoundClose"))
        .IntervaloAtaque = Val(LeerNPCs.GetValue(SectionName, "IntervaloAtaque"))
        .IntervaloMovimiento = Val(LeerNPCs.GetValue(SectionName, "IntervaloMovimiento"))
        .IntervaloLanzarHechizo = Val(LeerNPCs.GetValue(SectionName, "IntervaloLanzarHechizo"))
        .IntervaloRespawnMin = Val(LeerNPCs.GetValue(SectionName, "IntervaloRespawnMin"))
        .IntervaloRespawnMax = Val(LeerNPCs.GetValue(SectionName, "IntervaloRespawn"))
        .InformarRespawn = Val(LeerNPCs.GetValue(SectionName, "InformarRespawn"))
        .QuizaProb = Val(LeerNPCs.GetValue(SectionName, "QuizaProb"))
        .MinTameLevel = Val(LeerNPCs.GetValue(SectionName, "MinTameLevel", 1))
        .OnlyForGuilds = Val(LeerNPCs.GetValue(SectionName, "OnlyForGuilds", 0))
        .ShowKillerConsole = Val(LeerNPCs.GetValue(SectionName, "ShowKillerConsole", 0))
        .StatsMaxHp = Val(LeerNPCs.GetValue(SectionName, "MaxHP"))
        .StatsMinHp = Val(LeerNPCs.GetValue(SectionName, "MinHP"))
        .StatsMaxHit = Val(LeerNPCs.GetValue(SectionName, "MaxHIT"))
        .StatsMinHit = Val(LeerNPCs.GetValue(SectionName, "MinHIT"))
        .StatsDef = Val(LeerNPCs.GetValue(SectionName, "DEF"))
        .StatsDefM = Val(LeerNPCs.GetValue(SectionName, "DEFm"))
        .MagicResistance = Val(LeerNPCs.GetValue(SectionName, "MagicResistance"))
        .MagicDef = Val(LeerNPCs.GetValue(SectionName, "MagicDef"))
        .CantidadInvocaciones = Val(LeerNPCs.GetValue(SectionName, "CantidadInvocaciones"))
        .MagicBonus = Val(LeerNPCs.GetValue(SectionName, "MagicBonus"))
        .AIAlineacion = Val(LeerNPCs.GetValue(SectionName, "Alineacion"))
        MaxInventory = MAX_INVENTORY_SLOTS
        .InventoryCount = Val(LeerNPCs.GetValue(SectionName, "NROITEMS"))
        If .InventoryCount > MaxInventory Then .InventoryCount = MaxInventory
        If .InventoryCount > 0 Then
            ReDim .InventoryItems(1 To .InventoryCount)
            For LoopC = 1 To .InventoryCount
                ln = LeerNPCs.GetValue(SectionName, "Obj" & LoopC)
                .InventoryItems(LoopC).ObjIndex = Val(ReadField(1, ln, 45))
                .InventoryItems(LoopC).amount = Val(ReadField(2, ln, 45))
            Next LoopC
        Else
            Erase .InventoryItems
        End If
        .Humanoide = Val(LeerNPCs.GetValue(SectionName, "Humanoide"))
        .LanzaSpells = Val(LeerNPCs.GetValue(SectionName, "LanzaSpells"))
        If .LanzaSpells > 0 Then
            ReDim .Spells(1 To .LanzaSpells)
            .SpellRange = Val(LeerNPCs.GetValue(SectionName, "RangoSpell"))
            For LoopC = 1 To .LanzaSpells
                .Spells(LoopC).SpellIndex = Val(LeerNPCs.GetValue(SectionName, "Sp" & LoopC))
                .Spells(LoopC).Cd = Val(LeerNPCs.GetValue(SectionName, "Cd" & LoopC))
            Next LoopC
        Else
            .SpellRange = 0
            Erase .Spells
        End If
        If .npcType = e_NPCType.Entrenador Then
            .NroCriaturas = Val(LeerNPCs.GetValue(SectionName, "NroCriaturas"))
            If .NroCriaturas > 0 Then
                ReDim .Criaturas(1 To .NroCriaturas)
                For LoopC = 1 To .NroCriaturas
                    .Criaturas(LoopC).NpcIndex = Val(LeerNPCs.GetValue(SectionName, "CI" & LoopC))
                    .Criaturas(LoopC).NpcName = LeerNPCs.GetValue(SectionName, "CN" & LoopC)
                    .Criaturas(LoopC).tmpIndex = 0
                    .Criaturas(LoopC).PuedeInvocar = False
                Next LoopC
            Else
                Erase .Criaturas
            End If
        Else
            .NroCriaturas = 0
            Erase .Criaturas
        End If
        .RestriccionAtaque = Val(LeerNPCs.GetValue(SectionName, "RestriccionDeAtaque"))
        .RestriccionAyuda = Val(LeerNPCs.GetValue(SectionName, "RestriccionDeAyuda"))
        .RespawnValue = Val(LeerNPCs.GetValue(SectionName, "ReSpawn"))
        .DontHitVisiblePlayers = Val(LeerNPCs.GetValue(SectionName, "DontHitVisiblePlayers"))
        .AddToMapAiList = Val(LeerNPCs.GetValue(SectionName, "AddToMapAiList"))
        .DisplayCastMessage = Val(LeerNPCs.GetValue(SectionName, "DisplayCastMessage"))
        .Team = Val(LeerNPCs.GetValue(SectionName, "Team"))
        .Backup = Val(LeerNPCs.GetValue(SectionName, "BackUp"))
        .RespawnOrigPos = Val(LeerNPCs.GetValue(SectionName, "OrigPos"))
        .AfectaParalisis = Val(LeerNPCs.GetValue(SectionName, "AfectaParalisis"))
        .ImmuneToSpells = Val(LeerNPCs.GetValue(SectionName, "ImmuneToSpells", 0))
        .GolpeExacto = Val(LeerNPCs.GetValue(SectionName, "GolpeExacto"))
        .TranslationInmune = Val(LeerNPCs.GetValue(SectionName, "TranslationInmune"))
        .Snd1 = Val(LeerNPCs.GetValue(SectionName, "Snd1"))
        .Snd2 = Val(LeerNPCs.GetValue(SectionName, "Snd2"))
        .Snd3 = Val(LeerNPCs.GetValue(SectionName, "Snd3"))
        .SndRespawn = val(LeerNPCs.GetValue(SectionName, "SndRespawn"))
        aux = LeerNPCs.GetValue(SectionName, "NROEXP")
        If LenB(aux) = 0 Then
            .NroExp = 0
            Erase .Expresiones
        Else
            .NroExp = Val(aux)
            If .NroExp > 0 Then
                ReDim .Expresiones(1 To .NroExp)
                For LoopC = 1 To .NroExp
                    .Expresiones(LoopC) = LeerNPCs.GetValue(SectionName, "Exp" & LoopC)
                Next LoopC
            Else
                Erase .Expresiones
            End If
        End If
        .NumQuiza = Val(LeerNPCs.GetValue(SectionName, "NumQuiza"))
        If .NumQuiza > 0 Then
            ReDim .QuizaDropea(1 To .NumQuiza)
            For LoopC = 1 To .NumQuiza
                .QuizaDropea(LoopC) = LeerNPCs.GetValue(SectionName, "QuizaDropea" & LoopC)
            Next LoopC
        Else
            Erase .QuizaDropea
        End If
        aux = LeerNPCs.GetValue(SectionName, "NumQuest")
        If LenB(aux) = 0 Then
            .NumQuest = 0
            Erase .QuestNumber
        Else
            .NumQuest = Val(aux)
            If .NumQuest > 0 Then
                ReDim .QuestNumber(1 To .NumQuest)
                For LoopC = 1 To .NumQuest
                    .QuestNumber(LoopC) = Val(LeerNPCs.GetValue(SectionName, "QuestNumber" & LoopC))
                Next LoopC
            Else
                Erase .QuestNumber
            End If
        End If
        .NumDropQuest = Val(LeerNPCs.GetValue(SectionName, "NumDropQuest"))
        If .NumDropQuest > 0 Then
            ReDim .DropQuest(1 To .NumDropQuest)
            For LoopC = 1 To .NumDropQuest
                ln = LeerNPCs.GetValue(SectionName, "DropQuest" & LoopC)
                .DropQuest(LoopC).QuestIndex = Val(ReadField(1, ln, Asc("-")))
                .DropQuest(LoopC).ObjIndex = Val(ReadField(2, ln, Asc("-")))
                .DropQuest(LoopC).amount = Val(ReadField(3, ln, Asc("-")))
                .DropQuest(LoopC).Probabilidad = Val(ReadField(4, ln, Asc("-")))
            Next LoopC
        Else
            Erase .DropQuest
        End If
        .PathFindingVision = Val(LeerNPCs.GetValue(SectionName, "Distancia", RANGO_VISION_X))
        aux = LeerNPCs.GetValue(SectionName, "NumDestinos")
        If LenB(aux) = 0 Then
            .NumDestinos = 0
            Erase .Dest
        Else
            .NumDestinos = Val(aux)
            If .NumDestinos > 0 Then
                ReDim .Dest(1 To .NumDestinos)
                For LoopC = 1 To .NumDestinos
                    .Dest(LoopC) = LeerNPCs.GetValue(SectionName, "Dest" & LoopC)
                Next LoopC
            Else
                Erase .Dest
            End If
        End If
        .Interface = Val(LeerNPCs.GetValue(SectionName, "Interface"))
        .TipoItems = Val(LeerNPCs.GetValue(SectionName, "TipoItems"))
        .PuedeInvocar = Val(LeerNPCs.GetValue(SectionName, "PuedeInvocar"))
        cant = Val(LeerNPCs.GetValue(SectionName, "CaminataLen"))
        .CaminataLen = cant
        If cant > 0 Then
            ReDim .Caminata(1 To cant)
            For LoopC = 1 To cant
                Field = Split(LeerNPCs.GetValue(SectionName, "Caminata" & LoopC), ":")
                If UBound(Field) >= 2 Then
                    .Caminata(LoopC).OffsetX = Val(Field(0))
                    .Caminata(LoopC).OffsetY = Val(Field(1))
                    .Caminata(LoopC).Espera = Val(Field(2))
                Else
                    .Caminata(LoopC).OffsetX = 0
                    .Caminata(LoopC).OffsetY = 0
                    .Caminata(LoopC).Espera = 0
                End If
            Next LoopC
        Else
            Erase .Caminata
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "NPCs.LoadNpcInfoIntoCache", Erl)
End Sub

Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn As Boolean = True, Optional ByVal Reload As Boolean = False) As Integer
    On Error GoTo OpenNPC_Err
    Dim NpcIndex As Integer
    Dim Info As t_NpcInfoCache
    Dim LoopC As Long
    Dim cant As Long

    If Not NpcInfoCacheInitialized Then
        If LeerNPCs Is Nothing Then
            OpenNPC = 0
            Exit Function
        End If
        Call BuildNpcInfoCache
        If Not NpcInfoCacheInitialized Then
            OpenNPC = 0
            Exit Function
        End If
    End If

    If NpcNumber < LBound(NpcInfoCache) Or NpcNumber > UBound(NpcInfoCache) Then
        OpenNPC = 0
        Exit Function
    End If

    Info = NpcInfoCache(NpcNumber)

    If Not Info.Exists Then
        OpenNPC = 0
        Exit Function
    End If
#If DEBUGGING = 0 Then
    If Info.TestOnly > 0 Then Exit Function
#End If
    If Info.RequireToggle <> "" Then
        If Not IsFeatureEnabled(Info.RequireToggle) Then Exit Function
    End If

    NpcIndex = GetNextAvailableNpc
    If NpcIndex > MaxNPCs Then
        OpenNPC = 0
        Exit Function
    End If

    With NpcList(NpcIndex)
        .Numero = NpcNumber
        .name = Info.name
        .SubName = Info.SubName
        .Desc = Info.Desc
        .nivel = Info.nivel
        Call SetMovement(NpcIndex, Info.Movement)
        .flags.OldMovement = .Movement
        .flags.AguaValida = Info.AguaValida
        .flags.IsGlobalQuestBoss = Info.IsGlobalQuestBoss
        .flags.TierraInvalida = Info.TierraInvalida
        .flags.Faccion = Info.Faccion
        .flags.ElementalTags = Info.ElementalTags
        .npcType = Info.npcType
        .Char.body = Info.Body
        .Char.head = Info.Head
        .Char.Heading = Info.Heading
        .Char.BodyIdle = Info.BodyIdle
        .Char.Ataque1 = Info.Ataque1
        .Char.CastAnimation = Info.CastAnimation
        If .Char.BodyIdle > 0 Then
            .flags.NPCIdle = True
        End If
        If Info.AnimacionesCount > 0 Then
            ReDim .Char.Animation(1 To Info.AnimacionesCount)
            For LoopC = 1 To Info.AnimacionesCount
                .Char.Animation(LoopC) = Info.Animaciones(LoopC)
            Next LoopC
        Else
            ReDim .Char.Animation(0)
        End If
        .Char.WeaponAnim = Info.WeaponAnim
        .Char.ShieldAnim = Info.ShieldAnim
        .Char.CascoAnim = Info.CascoAnim
        .Char.CartAnim = Info.CartAnim
        .Attackable = Info.Attackable
        .Comercia = Info.Comercia
        .Craftea = Info.Craftea
        .Hostile = Info.Hostile
        .flags.OldHostil = .Hostile
        .AttackRange = Info.AttackRange
        .ProjectileType = Info.ProjectileType
        .PreferedRange = Info.PreferedRange
        .GiveEXP = Info.GiveEXP
        .Distancia = Info.Distancia
        .GiveEXPClan = Info.GiveEXPClan
        .flags.ExpCount = .GiveEXP
        .Veneno = Info.Veneno
        .flags.Domable = Info.Domable
        .flags.AttackableByEveryone = Info.AttackableByEveryone
        .flags.MapEntryPrice = Info.MapEntryPrice
        .flags.MapTargetEntry = Info.MapTargetEntry
        .flags.MapTargetEntryX = Info.MapTargetEntryX
        .flags.MapTargetEntryY = Info.MapTargetEntryY
        .flags.ArenaEnabled = Info.ArenaEnabled
        .GiveGLD = Info.GiveGLD
        .PoderAtaque = Info.PoderAtaque
        .PoderEvasion = Info.PoderEvasion
        .InvReSpawn = Info.InvReSpawn
        .showName = Info.ShowName
        .GobernadorDe = Info.GobernadorDe
        .SoundOpen = Info.SoundOpen
        .SoundClose = Info.SoundClose
        .IntervaloAtaque = Info.IntervaloAtaque
        .IntervaloMovimiento = Info.IntervaloMovimiento
        .IntervaloLanzarHechizo = Info.IntervaloLanzarHechizo
        .Contadores.IntervaloRespawn = RandomNumber(Info.IntervaloRespawnMin, Info.IntervaloRespawnMax)
        .InformarRespawn = Info.InformarRespawn
        .QuizaProb = Info.QuizaProb
        .MinTameLevel = Info.MinTameLevel
        .OnlyForGuilds = Info.OnlyForGuilds
        .ShowKillerConsole = Info.ShowKillerConsole
        If .IntervaloMovimiento = 0 Then
            .IntervaloMovimiento = 380
            .Char.speeding = IntervaloNPCAI / 330
        Else
            .Char.speeding = 210 / .IntervaloMovimiento
        End If
        If .IntervaloLanzarHechizo = 0 Then
            .IntervaloLanzarHechizo = 8000
        End If
        If .IntervaloAtaque = 0 Then
            .IntervaloAtaque = 2000
        End If
        .Stats.MaxHp = Info.StatsMaxHp
        .Stats.MinHp = Info.StatsMinHp
        .Stats.MaxHit = Info.StatsMaxHit
        .Stats.MinHIT = Info.StatsMinHit
        .Stats.def = Info.StatsDef
        .Stats.defM = Info.StatsDefM
        .Stats.MagicResistance = Info.MagicResistance
        .Stats.MagicDef = Info.MagicDef
        .Stats.CantidadInvocaciones = Info.CantidadInvocaciones
        .Stats.MagicBonus = Info.MagicBonus
        If .Stats.CantidadInvocaciones > 0 Then
            ReDim .Stats.NpcsInvocados(1 To .Stats.CantidadInvocaciones)
            For LoopC = 1 To .Stats.CantidadInvocaciones
                Call ClearNpcRef(.Stats.NpcsInvocados(LoopC))
            Next LoopC
        End If
        .flags.AIAlineacion = Info.AIAlineacion
        .invent.NroItems = Info.InventoryCount
        If .invent.NroItems > UBound(.invent.Object) Then
            Debug.Print "Error in NPC " & .name & " configuration, too many items in inventory! Max possible is " & UBound(.invent.Object)
        End If
        Debug.Assert .invent.NroItems <= UBound(.invent.Object)
        .invent.NroItems = Min(UBound(.invent.Object), .invent.NroItems)
        .Humanoide = CBool(Info.Humanoide)
        For LoopC = 1 To .invent.NroItems
            .invent.Object(LoopC).ObjIndex = Info.InventoryItems(LoopC).ObjIndex
            .invent.Object(LoopC).amount = Info.InventoryItems(LoopC).amount
        Next LoopC
        .flags.LanzaSpells = Info.LanzaSpells
        If .flags.LanzaSpells > 0 Then
            ReDim .Spells(1 To .flags.LanzaSpells)
            .SpellRange = Info.SpellRange
        Else
            Erase .Spells
            .SpellRange = 0
        End If
        For LoopC = 1 To .flags.LanzaSpells
            .Spells(LoopC).SpellIndex = Info.Spells(LoopC).SpellIndex
            .Spells(LoopC).Cd = Info.Spells(LoopC).Cd
            .Spells(LoopC).lastUse = 0
        Next LoopC
        If .npcType = e_NPCType.Entrenador Then
            .NroCriaturas = Info.NroCriaturas
            If .NroCriaturas > 0 Then
                ReDim .Criaturas(1 To .NroCriaturas)
                For LoopC = 1 To .NroCriaturas
                    .Criaturas(LoopC).NpcIndex = Info.Criaturas(LoopC).NpcIndex
                    .Criaturas(LoopC).NpcName = Info.Criaturas(LoopC).NpcName
                    .Criaturas(LoopC).tmpIndex = 0
                    .Criaturas(LoopC).PuedeInvocar = False
                Next LoopC
            Else
                Erase .Criaturas
            End If
        Else
            .NroCriaturas = 0
            Erase .Criaturas
        End If
        Call ResetMask(.flags.StatusMask)
        .flags.NPCActive = True
        Call ResetMask(.flags.BehaviorFlags)
        Select Case Info.RestriccionAtaque
            Case 0
                Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eAttackNpc)
                Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers)
            Case 1
                Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers)
            Case 2
                Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eAttackNpc)
        End Select
        Select Case Info.RestriccionAyuda
            Case 1
                Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eHelpNpc)
            Case 2
                Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers)
            Case 3
                Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eHelpNpc)
                Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers)
        End Select
        If Respawn Then
            .flags.Respawn = Info.RespawnValue
        Else
            .flags.Respawn = 1
        End If
        If Info.DontHitVisiblePlayers > 0 Then Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eDontHitVisiblePlayers)
        If Info.AddToMapAiList > 0 Then Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eConsideredByMapAi)
        If Info.DisplayCastMessage > 0 Then Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eDisplayCastMessage)
        .flags.team = Info.Team
        .flags.backup = Info.Backup
        .flags.RespawnOrigPos = Info.RespawnOrigPos
        .flags.AfectaParalisis = Info.AfectaParalisis
        .flags.ImmuneToSpells = Info.ImmuneToSpells
        .flags.GolpeExacto = Info.GolpeExacto
        If Info.TranslationInmune > 0 Then Call SetMask(.flags.EffectInmunity, e_Inmunities.eTranslation)
        .flags.Snd1 = Info.Snd1
        .flags.Snd2 = Info.Snd2
        .flags.Snd3 = Info.Snd3
        .flags.SndRespawn = Info.SndRespawn
        .NroExpresiones = Info.NroExp
        If .NroExpresiones > 0 Then
            ReDim .Expresiones(1 To .NroExpresiones)
            For LoopC = 1 To .NroExpresiones
                .Expresiones(LoopC) = Info.Expresiones(LoopC)
            Next LoopC
        Else
            Erase .Expresiones
        End If
        .NumQuiza = Info.NumQuiza
        If .NumQuiza > 0 Then
            ReDim .QuizaDropea(1 To .NumQuiza)
            For LoopC = 1 To .NumQuiza
                .QuizaDropea(LoopC) = Info.QuizaDropea(LoopC)
            Next LoopC
        Else
            Erase .QuizaDropea
        End If
        .NumQuest = Info.NumQuest
        If .NumQuest > 0 Then
            ReDim .QuestNumber(1 To .NumQuest)
            For LoopC = 1 To .NumQuest
                .QuestNumber(LoopC) = Info.QuestNumber(LoopC)
            Next LoopC
        Else
            Erase .QuestNumber
        End If
        .NumDropQuest = Info.NumDropQuest
        If .NumDropQuest > 0 Then
            ReDim .DropQuest(1 To .NumDropQuest)
            For LoopC = 1 To .NumDropQuest
                .DropQuest(LoopC) = Info.DropQuest(LoopC)
            Next LoopC
        Else
            Erase .DropQuest
        End If
        .pathFindingInfo.RangoVision = Info.PathFindingVision
        .pathFindingInfo.OriginalVision = .pathFindingInfo.RangoVision
        .pathFindingInfo.TargetUnreachable = False
        .pathFindingInfo.PreviousAttackable = .Attackable
        .pathFindingInfo.StrafeOffset.x = 0
        .pathFindingInfo.StrafeOffset.y = 0
        .pathFindingInfo.StrafeExpiresAt = 0
        .pathFindingInfo.OrbitDirection = 0
        .pathFindingInfo.OrbitReevaluateAt = 0
        .pathFindingInfo.NextPathRecomputeAt = 0
        ReDim .pathFindingInfo.Path(1 To MAX_PATH_LENGTH)
        .NumDestinos = Info.NumDestinos
        If .NumDestinos > 0 Then
            ReDim .dest(1 To .NumDestinos)
            For LoopC = 1 To .NumDestinos
                .dest(LoopC) = Info.Dest(LoopC)
            Next LoopC
        Else
            Erase .dest
        End If
        .Interface = Info.Interface
        .TipoItems = Info.TipoItems
        .PuedeInvocar = Info.PuedeInvocar
        If NumUsers > 0 Then
            Call AnimacionIdle(NpcIndex, True)
        End If
        If .Movement = Caminata Then
            cant = Info.CaminataLen
            If cant = 0 Then
                Call SetMovement(NpcIndex, Estatico)
            Else
                ReDim .Caminata(1 To cant)
                For LoopC = 1 To cant
                    .Caminata(LoopC).offset.x = Info.Caminata(LoopC).OffsetX
                    .Caminata(LoopC).offset.y = Info.Caminata(LoopC).OffsetY
                    .Caminata(LoopC).Espera = Info.Caminata(LoopC).Espera
                Next LoopC
                .CaminataActual = 1
            End If
        Else
            Erase .Caminata
        End If
    End With

    If Reload = False Then
        If NpcIndex > LastNPC Then LastNPC = NpcIndex
        NumNPCs = NumNPCs + 1
    End If

    OpenNPC = NpcIndex
    Exit Function
OpenNPC_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.OpenNPC", Erl)
End Function

Function NpcSellsItem(ByVal NpcNumber As Integer, ByVal NroObjeto As Integer) As Boolean
    On Error GoTo NpcSellsItem_Err
    Dim LoopC As Long

    If Not NpcInfoCacheInitialized Then
        If LeerNPCs Is Nothing Then Exit Function
        Call BuildNpcInfoCache
        If Not NpcInfoCacheInitialized Then Exit Function
    End If

    If NpcNumber < LBound(NpcInfoCache) Or NpcNumber > UBound(NpcInfoCache) Then Exit Function
    If Not NpcInfoCache(NpcNumber).Exists Then Exit Function

    With NpcInfoCache(NpcNumber)
        For LoopC = 1 To .InventoryCount
            If NroObjeto = .InventoryItems(LoopC).ObjIndex Then
                NpcSellsItem = True
                Exit Function
            End If
        Next LoopC
    End With

    NpcSellsItem = False
    Exit Function
NpcSellsItem_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.NpcSellsItem", Erl)
End Function

Sub DoFollow(ByVal NpcIndex As Integer, ByVal username As String)
    On Error GoTo DoFollow_Err
    With NpcList(NpcIndex)
        If .flags.Follow Then
            .flags.AttackedBy = vbNullString
            Call SetUserRef(.TargetUser, 0)
            .flags.Follow = False
            Call SetMovement(NpcIndex, .flags.OldMovement)
            .Hostile = .flags.OldHostil
        Else
            Dim player As t_UserReference
            player = NameIndex(username)
            If IsValidUserRef(player) Then
                .flags.AttackedBy = username
                .flags.AttackedTime = GlobalFrameTime
                .TargetUser = player
                .flags.Follow = True
                Call SetMovement(NpcIndex, e_TipoAI.NpcDefensa)
                .Hostile = 0
            End If
        End If
    End With
    Exit Sub
DoFollow_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.DoFollow", Erl)
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
    On Error GoTo FollowAmo_Err
    With NpcList(NpcIndex)
        .flags.Follow = True
        Call SetMovement(NpcIndex, e_TipoAI.SigueAmo)
        .Hostile = 0
        Call ClearUserRef(.TargetUser)
        Call ClearNpcRef(.TargetNPC)
    End With
    Exit Sub
FollowAmo_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.FollowAmo", Erl)
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
    On Error GoTo AllFollowAmo_Err
    Dim j As Long
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j).ArrayIndex > 0 Then
            If IsValidNpcRef(UserList(UserIndex).MascotasIndex(j)) Then
                Call FollowAmo(UserList(UserIndex).MascotasIndex(j).ArrayIndex)
            Else
                Call ClearNpcRef(UserList(UserIndex).MascotasIndex(j))
            End If
        End If
    Next j
    Exit Sub
AllFollowAmo_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.AllFollowAmo", Erl)
End Sub

Public Function ObtenerIndiceRespawn() As Integer
    On Error GoTo ErrHandler
    Dim LoopC As Integer
    For LoopC = 1 To MaxRespawn
        If Not RespawnList(LoopC).flags.NPCActive Then Exit For
    Next LoopC
    ObtenerIndiceRespawn = LoopC
    Exit Function
ErrHandler:
    Call LogError("Error en ObtenerIndiceRespawn")
End Function

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    On Error GoTo QuitarMascota_Err
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i).ArrayIndex = NpcIndex Then
            Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
            UserList(UserIndex).MascotasType(i) = 0
            UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
            UserList(UserIndex).flags.ModificoMascotas = True
            Exit For
        End If
    Next i
    Exit Sub
QuitarMascota_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.QuitarMascota", Erl)
End Sub

Sub AnimacionIdle(ByVal NpcIndex As Integer, ByVal Show As Boolean)
    On Error GoTo Handler
    With NpcList(NpcIndex)
        If .Char.BodyIdle = 0 Then Exit Sub
        If .flags.NPCIdle = Show Then Exit Sub
        .flags.NPCIdle = Show
        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, .Char.Heading)
    End With
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "NPCs.AnimacionIdle", Erl)
End Sub

Sub WarpNpcChar(ByVal NpcIndex As Integer, ByVal Map As Byte, ByVal x As Integer, ByVal y As Integer, Optional ByVal FX As Boolean = False)
    Dim NuevaPos  As t_WorldPos
    Dim FuturePos As t_WorldPos
    Call EraseNPCChar(NpcIndex)
    FuturePos.Map = Map
    FuturePos.x = x
    FuturePos.y = y
    Call ClosestLegalPos(FuturePos, NuevaPos, True, True)
    If NuevaPos.Map = 0 Or NuevaPos.x = 0 Or NuevaPos.y = 0 Then
        Debug.Print "Error al tepear NPC"
        Call QuitarNPC(NpcIndex, eFailedToWarp)
    Else
        NpcList(NpcIndex).pos = NuevaPos
        Call MakeNPCChar(True, 0, NpcIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
        If FX Then                                    'FX
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessagePlayWave(SND_WARP, NuevaPos.x, NuevaPos.y))
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageCreateFX(NpcList(NpcIndex).Char.charindex, e_GraphicEffects.ModernGmWarp, 0))
        End If
    End If
End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado. Se usa para mover NPCs del camino de otro char.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveNpcToSide(ByVal NpcIndex As Integer, ByVal Heading As e_Heading)
    On Error GoTo Handler
    With NpcList(NpcIndex)
        ' Elegimos un lado al azar
        Dim r As Integer
        r = RandomNumber(0, 1) * 2 - 1 ' -1 o 1
        ' Roto el heading original hacia ese lado
        Heading = Rotate_Heading(Heading, r)
        ' Intento moverlo para ese lado
        If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        ' Si falló, intento moverlo para el lado opuesto
        Heading = InvertHeading(Heading)
        If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
        Dim NuevaPos As t_WorldPos
        Call ClosestLegalPos(.pos, NuevaPos, .flags.AguaValida, .flags.TierraInvalida = 0)
        Call WarpNpcChar(NpcIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
    End With
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "NPCs.MoveNpcToSide", Erl)
End Sub

Public Sub DummyTargetAttacked(ByVal NpcIndex As Integer)
    With NpcList(NpcIndex)
        .Contadores.UltimoAtaque = 30
        If RandomNumber(1, 5) = 1 Then
            If UBound(.Char.Animation) > 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(.Char.charindex, .Char.Animation(1)))
            End If
        End If
    End With
End Sub

Public Sub KillRandomNpc()
    Dim validNpc As Boolean: validNpc = False
    Dim NpcIndex As Integer: NpcIndex = 0
    If GetAvailableNpcIndex > 8000 Or GetAvailableNpcIndex = 0 Then
        Exit Sub
    End If
    Do While Not validNpc
        NpcIndex = RandomNumber(1, 10000)
        If NpcList(NpcIndex).flags.NPCActive And NpcList(NpcIndex).Hostile > 0 Then
            validNpc = True
        End If
    Loop
    Call MuereNpc(NpcIndex, 0)
End Sub

Public Function CanMove(ByRef counter As t_NpcCounters, ByRef flags As t_NPCFlags) As Boolean
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    CanMove = (flags.Inmovilizado + flags.Paralizado = 0) And Not flags.TranslationActive
    CanMove = CanMove And DeadlinePassed(nowRaw, counter.StunEndTime)
End Function

Public Function CanAttack(ByRef counter As t_NpcCounters, ByRef flags As t_NPCFlags) As Boolean
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    CanAttack = (flags.Paralizado = 0) And DeadlinePassed(nowRaw, counter.StunEndTime)
End Function

Public Sub StunNPc(ByRef Counters As t_NpcCounters)
    Counters.StunEndTime = AddMod32(GetTickCountRaw(), NpcStunTime)
End Sub

Public Function ModifyHealth(ByVal NpcIndex As Integer, ByVal amount As Long, Optional ByVal MinValue = 0) As Boolean
    With NpcList(NpcIndex)
        ModifyHealth = False
        .Stats.MinHp = .Stats.MinHp + amount
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
        End If
        If .Stats.MinHp <= MinValue Then
            .Stats.MinHp = MinValue
            ModifyHealth = True
        End If
        Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageNpcUpdateHP(NpcIndex))
    End With
End Function

Public Function DoDamageOrHeal(ByVal NpcIndex As Integer, _
                               ByVal SourceIndex As Integer, _
                               ByVal SourceType As e_ReferenceType, _
                               ByVal amount As Long, _
                               ByVal DamageSourceType As e_DamageSourceType, _
                               ByVal DamageSourceIndex As Integer, _
                               Optional ByVal DamageColor As Long = vbRed) As e_DamageResult
    On Error GoTo DoDamageOrHeal_Err
    Dim DamageStr As String
    Dim Color     As Long
    If amount > 0 Then
        Color = vbGreen
    Else
        Color = DamageColor
    End If
    If amount < 0 Then
        If SourceType = eUser Then
            DamageStr = PonerPuntos(Math.Abs(amount))
            If UserList(SourceIndex).ChatCombate = 1 Then
                Call WriteLocaleMsg(SourceIndex, 382, e_FontTypeNames.FONTTYPE_FIGHT, DamageStr)
            End If
        End If
        amount = EffectsOverTime.TargetApplyDamageReduction(NpcList(NpcIndex).EffectOverTime, amount, SourceIndex, SourceType, DamageSourceType)
        Call EffectsOverTime.TargetWasDamaged(NpcList(NpcIndex).EffectOverTime, SourceIndex, SourceType, DamageSourceType)
        Call RegisterNpcDamageStrafe(NpcIndex, SourceIndex, SourceType)
    End If
    DamageStr = PonerPuntos(Math.Abs(amount))
    With NpcList(NpcIndex)
        Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageTextOverChar(DamageStr, .Char.charindex, Color))
        ' Mascotas dan experiencia al amo
        If SourceType = eNpc And amount < 0 Then
            If IsValidUserRef(NpcList(SourceIndex).MaestroUser) Then
                Call CalcularDarExp(NpcList(SourceIndex).MaestroUser.ArrayIndex, NpcIndex, -amount)
                ' NPC de invasión
                If .flags.InvasionIndex Then
                    Call SumarScoreInvasion(NpcList(SourceIndex).flags.InvasionIndex, NpcList(SourceIndex).MaestroUser.ArrayIndex, -amount)
                End If
            End If
        ElseIf SourceType = eUser And amount < 0 Then
            ' NPC de invasión
            If NpcList(NpcIndex).flags.InvasionIndex Then
                Call SumarScoreInvasion(NpcList(NpcIndex).flags.InvasionIndex, SourceIndex, -amount)
            End If
            Call CalcularDarExp(SourceIndex, NpcIndex, -amount)
        End If
        If NPCs.ModifyHealth(NpcIndex, amount) Then
            DoDamageOrHeal = eDead
            CustomScenarios.NpcDie (NpcIndex)
            If SourceType = eUser Then
                Call CustomScenarios.PlayerKillNpc(.pos.Map, NpcIndex, SourceIndex, DamageSourceType, DamageSourceIndex)
                Call MuereNpc(NpcIndex, SourceIndex)
            Else
                If IsValidUserRef(NpcList(SourceIndex).MaestroUser) Then
                    Call PlayerKillNpc(NpcList(NpcIndex).pos.Map, NpcIndex, NpcList(SourceIndex).MaestroUser.ArrayIndex, e_pet, DamageSourceIndex)
                    Call FollowAmo(SourceIndex)
                    Call MuereNpc(NpcIndex, NpcList(SourceIndex).MaestroUser.ArrayIndex)
                Else
                    Call MuereNpc(NpcIndex, -1)
                End If
            End If
            Exit Function
        End If
    End With
    DoDamageOrHeal = eStillAlive
    Exit Function
DoDamageOrHeal_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.DoDamageOrHeal_Err", Erl)
End Function

Public Function UserCanAttackNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As t_AttackInteractionResult
    On Error GoTo UserCanAttackNpc_Err
    UserCanAttackNpc.CanAttack = False
    UserCanAttackNpc.TurnPK = False
    'Estas muerto?
    If UserList(UserIndex).flags.Muerto = 1 Then
        UserCanAttackNpc.Result = eDeathAttacker
        Exit Function
    End If
    If UserList(UserIndex).flags.Montado = 1 Then
        UserCanAttackNpc.Result = eMounted
        Exit Function
    End If
    If UserList(UserIndex).flags.Inmunidad = 1 Then
        UserCanAttackNpc.Result = eCreatureInmunity
        Exit Function
    End If
    'Solo administradores, dioses y usuarios pueden atacar a NPC's (PARA TESTING)
    If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 And NpcList(NpcIndex).npcType <> e_NPCType.DummyTarget Then
        UserCanAttackNpc.Result = eInvalidPrivilege
        Exit Function
    End If
    ' No podes atacar si estas en consulta
    If UserList(UserIndex).flags.EnConsulta Then
        UserCanAttackNpc.Result = eTalkWithMaster
        Exit Function
    End If
    'Es una criatura atacable?
    If NpcList(NpcIndex).Attackable = 0 Or NpcList(NpcIndex).pathFindingInfo.TargetUnreachable Then
        UserCanAttackNpc.Result = eInmuneNpc
        Exit Function
    End If
    'Es valida la distancia a la cual estamos atacando?
    If Abs(UserList(UserIndex).pos.x - NpcList(NpcIndex).pos.x) > RANGO_VISION_X Or Abs(UserList(UserIndex).pos.y - NpcList(NpcIndex).pos.y) > RANGO_VISION_Y Then
        UserCanAttackNpc.Result = eOutOfRange
        Exit Function
    End If
    Dim IsPet As Boolean
    IsPet = IsValidUserRef(NpcList(NpcIndex).MaestroUser)
    'Si el usuario pertenece a una faccion
    If esArmada(UserIndex) Or esCaos(UserIndex) Then
        ' Y el NPC pertenece a la misma faccion
        If NpcList(NpcIndex).flags.Faccion = UserList(UserIndex).Faccion.Status Then
            UserCanAttackNpc.Result = eSameFaction
            Exit Function
        End If
        ' Si es una mascota, checkeamos en el Maestro
        If IsPet Then
            If UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Faccion.Status = UserList(UserIndex).Faccion.Status Then
                UserCanAttackNpc.Result = eSameFaction
                Exit Function
            End If
        End If
    End If
    If IsPet Then
        If UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Grupo.Id = UserList(UserIndex).Grupo.Id And UserList(UserIndex).Grupo.Id > 0 Then
            UserCanAttackNpc.Result = eSameGroup
            Exit Function
        End If
        If esArmada(UserIndex) Then
            If UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Faccion.Status = Ciudadano Or UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Faccion.Status = Armada Then
                UserCanAttackNpc.Result = eSameFaction
                Exit Function
            End If
        End If
    End If
    If Status(UserIndex) = Ciudadano Then
        If IsPet And NpcList(NpcIndex).MaestroUser.ArrayIndex = UserIndex Then
            UserCanAttackNpc.Result = eOwnPet
            Exit Function
        End If
    End If
    If UserList(UserIndex).flags.CurrentTeam <> 0 Then
        If UserList(UserIndex).flags.CurrentTeam = NpcList(NpcIndex).flags.team Then
            UserCanAttackNpc.Result = eSameTeam
            Exit Function
        End If
    End If
    If IsPet And MapInfo(UserList(UserIndex).pos.Map).Seguro Then
        UserCanAttackNpc.Result = eSafeArea
        Exit Function
    End If
    If Status(UserIndex) = Ciudadano Then
        If NpcList(NpcIndex).flags.Faccion = Armada Or NpcList(NpcIndex).flags.Faccion = consejo Then
            If UserList(UserIndex).flags.Seguro Then
                UserCanAttackNpc.Result = eRemoveSafe
                Exit Function
            Else
                UserCanAttackNpc.Result = eAttackSameFaction
                UserCanAttackNpc.TurnPK = True
                UserCanAttackNpc.CanAttack = True
                Exit Function
            End If
        End If
    End If
    If Status(UserIndex) = e_Facciones.Ciudadano Or Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
        'Es el NPC mascota de alguien?
        If IsPet Then
            Select Case UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Faccion.Status
                Case e_Facciones.Armada
                    If UserList(UserIndex).flags.Seguro Then
                        UserCanAttackNpc.Result = eRemoveSafe
                        Exit Function
                    Else
                        UserCanAttackNpc.TurnPK = True
                        UserCanAttackNpc.CanAttack = True
                        UserCanAttackNpc.Result = eAttackPetSameFaction
                        Exit Function
                    End If
                Case e_Facciones.Ciudadano
                    If UserList(UserIndex).flags.Seguro Then
                        UserCanAttackNpc.Result = eRemoveSafe
                        Exit Function
                    Else
                        UserCanAttackNpc.TurnPK = True
                        UserCanAttackNpc.CanAttack = True
                        UserCanAttackNpc.Result = eAttackPetSameFaction
                        Exit Function
                    End If
                Case Else
                    UserCanAttackNpc.Result = eCanAttack
                    UserCanAttackNpc.CanAttack = True
                    Exit Function
            End Select
        End If
        If NpcList(NpcIndex).flags.AttackableByEveryone = 0 Then
            If NpcList(NpcIndex).flags.team = 0 Then
                Dim CurrentOwnerIndex As Integer: CurrentOwnerIndex = GetOwnedBy(NpcIndex)
                If CurrentOwnerIndex <> 0 Then
                    If CurrentOwnerIndex <> UserIndex And IsValidNpcRef(UserList(CurrentOwnerIndex).flags.NPCAtacado) Then
                        If UserList(CurrentOwnerIndex).flags.NPCAtacado.ArrayIndex = NpcIndex And UserList(CurrentOwnerIndex).flags.Muerto = 0 And (Status(CurrentOwnerIndex) = _
                                Ciudadano Or Status(CurrentOwnerIndex) = Armada Or Status(CurrentOwnerIndex) = consejo) And (UserList(UserIndex).GuildIndex = 0 Or UserList( _
                                UserIndex).GuildIndex <> UserList(CurrentOwnerIndex).GuildIndex) And (UserList(UserIndex).Grupo.EnGrupo = False Or UserList(UserIndex).Grupo.Id _
                                <> UserList(CurrentOwnerIndex).Grupo.Id) Then
                            If UserList(UserIndex).flags.Seguro Then
                                UserCanAttackNpc.Result = eRemoveSafeCitizenNpc
                                Exit Function
                            Else
                                UserCanAttackNpc.TurnPK = True
                                UserCanAttackNpc.CanAttack = True
                                UserCanAttackNpc.Result = eAttackCitizenNpc
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    UserCanAttackNpc.CanAttack = True
    UserCanAttackNpc.Result = eCanAttack
    Exit Function
UserCanAttackNpc_Err:
    Call TraceError(Err.Number, Err.Description, "Npcs.UserCanAttackNpc", Erl)
End Function

Public Function Inmovilize(ByVal SourceIndex As Integer, ByVal TargetIndex As Integer, ByVal Time As Integer, ByVal FX As Integer) As Boolean
    With NpcList(TargetIndex)
        Call NPCAtacado(TargetIndex, SourceIndex)
        .flags.Inmovilizado = 1
        .Contadores.Inmovilizado = Time
        .flags.Paralizado = 0
        .Contadores.Paralisis = 0
        Call AnimacionIdle(TargetIndex, True)
        Call SendData(SendTarget.ToNPCAliveArea, TargetIndex, PrepareMessageCreateFX(.Char.charindex, FX, 0, .pos.x, .pos.y))
        Inmovilize = True
    End With
End Function

Public Function GetPhysicalDamageModifier(ByRef Npc As t_Npc) As Single
    GetPhysicalDamageModifier = max(1 + Npc.Modifiers.PhysicalDamageBonus, 0)
End Function

Public Function GetMagicDamageModifier(ByRef Npc As t_Npc) As Single
    GetMagicDamageModifier = max(1 + Npc.Modifiers.MagicDamageBonus, 0)
End Function

Public Function GetMagicDamageReduction(ByRef Npc As t_Npc) As Single
    GetMagicDamageReduction = max(1 - Npc.Modifiers.MagicDamageReduction, 0)
End Function

Public Function GetPhysicDamageReduction(ByRef Npc As t_Npc) As Single
    GetPhysicDamageReduction = max(1 - Npc.Modifiers.PhysicalDamageReduction, 0)
End Function

Public Function CanAttackUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As e_AttackInteractionResult
    With NpcList(NpcIndex)
        If Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers) Then
            CanAttackUser = eNotEnougthPrivileges
            Exit Function
        End If
        If UserList(UserIndex).Stats.MinHp = 0 Then
            CanAttackUser = eDeathTarget
            Exit Function
        End If
        If EsGM(UserIndex) Then
            If UserList(UserIndex).flags.EnConsulta Then
                CanAttackUser = eNotEnougthPrivileges
                Exit Function
            End If
            If Not UserList(UserIndex).flags.AdminPerseguible Then
                CanAttackUser = eNotEnougthPrivileges
                Exit Function
            End If
            If UserList(UserIndex).flags.invisible Then
                CanAttackUser = eNotEnougthPrivileges
                Exit Function
            End If
        End If
        Dim AttackerFaction As e_Facciones
        If IsValidUserRef(.MaestroUser) Then
            AttackerFaction = UserList(.MaestroUser.ArrayIndex).Faccion.Status
        Else
            AttackerFaction = .flags.Faccion
        End If
        If Not FactionCanAttackFaction(AttackerFaction, UserList(UserIndex).Faccion.Status) Then
            CanAttackUser = eSameFaction
            Exit Function
        End If
        If .flags.team <> 0 Then
            If .flags.team = UserList(UserIndex).flags.CurrentTeam Then
                CanAttackUser = eSameTeam
                Exit Function
            End If
        End If
    End With
    CanAttackUser = eCanAttack
End Function

Public Function CanHelpUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As e_InteractionResult
    With NpcList(NpcIndex)
        If Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers) Then
            CanHelpUser = eCantHelpUsers
            Exit Function
        End If
        Dim SourceFaction As e_Facciones
        If IsValidUserRef(.MaestroUser) Then
            SourceFaction = UserList(.MaestroUser.ArrayIndex).Faccion.Status
        Else
            SourceFaction = .flags.Faccion
        End If
        CanHelpUser = FactionCanHelpFaction(SourceFaction, UserList(UserIndex).Faccion.Status)
        If CanHelpUser <> e_InteractionResult.eInteractionOk Then
            Exit Function
        End If
        If UserList(UserIndex).flags.CurrentTeam > 0 Then
            Dim NpcTeam As Byte
            NpcTeam = .flags.team
            If NpcTeam <> UserList(UserIndex).flags.CurrentTeam Then
                CanHelpUser = eCantHelpUsers
                Exit Function
            End If
        End If
        CanHelpUser = eInteractionOk
    End With
End Function

Public Function CanHelpNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer) As e_InteractionResult
    CanHelpNpc = eInteractionOk
    If NpcList(NpcIndex).flags.team > 0 And NpcList(NpcIndex).flags.team <> NpcList(TargetNPC).flags.team Then
        CanHelpNpc = eOposingFaction
        Exit Function
    End If
End Function

Public Function CanAttackNpc(ByVal NpcIndex As Integer, ByVal TargetIndex As Integer) As e_AttackInteractionResult
    If NpcIndex = TargetIndex Then
        CanAttackNpc = eSameFaction
        Exit Function
    End If
    With NpcList(NpcIndex)
        If NpcList(TargetIndex).Attackable = 0 Then
            CanAttackNpc = eNotEnougthPrivileges
            Exit Function
        End If
        If Not NpcList(TargetIndex).flags.NPCActive Then
            CanAttackNpc = eNotEnougthPrivileges
            Exit Function
        End If
        If Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackNpc) Then
            CanAttackNpc = eNotEnougthPrivileges
            Exit Function
        End If
        Dim TargetFaction         As e_Facciones
        Dim AttackerFaction       As e_Facciones
        Dim AttackerIsFreeCrature As Boolean
        Dim TargetIsFreeCrature   As Boolean
        If IsValidUserRef(NpcList(TargetIndex).MaestroUser) Then
            TargetFaction = UserList(NpcList(TargetIndex).MaestroUser.ArrayIndex).Faccion.Status
            TargetIsFreeCrature = False
        Else
            TargetFaction = NpcList(TargetIndex).flags.Faccion
            TargetIsFreeCrature = True
        End If
        If IsValidUserRef(.MaestroUser) Then
            AttackerFaction = UserList(.MaestroUser.ArrayIndex).Faccion.Status
            AttackerIsFreeCrature = False
        Else
            AttackerFaction = .flags.Faccion
            AttackerIsFreeCrature = True
        End If
        If Not FactionCanAttackFaction(AttackerFaction, TargetFaction) Then
            CanAttackNpc = eSameFaction
            Exit Function
        End If
        If .flags.team = NpcList(TargetIndex).flags.team Then
            CanAttackNpc = eSameTeam
            Exit Function
        End If
    End With
    CanAttackNpc = eCanAttack
End Function

Public Function GetEvasionBonus(ByRef Npc As t_Npc) As Integer
    GetEvasionBonus = Npc.Modifiers.EvasionBonus
End Function

Public Function GetHitBonus(ByRef Npc As t_Npc) As Integer
    GetHitBonus = Npc.Modifiers.HitBonus
End Function

'Defines the healing bonus when using a potion, a spell or any other healing source
Public Function GetSelfHealingBonus(ByRef Npc As t_Npc) As Single
    GetSelfHealingBonus = max(1 + Npc.Modifiers.SelfHealingBonus, 0)
End Function

'Defines bonus when healing someone with magic
Public Function GetMagicHealingBonus(ByRef Npc As t_Npc) As Single
    GetMagicHealingBonus = max(1 + Npc.Modifiers.MagicHealingBonus, 0)
End Function

Public Function CanSeeUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    CanSeeUser = UserMod.IsVisible(UserList(UserIndex))
End Function

Public Function CanPerformAttackAction(ByVal NpcIndex As Integer, ByVal AttackInterval As Long)
    With NpcList(NpcIndex)
        CanPerformAttackAction = GlobalFrameTime - .Contadores.IntervaloLanzarHechizo > AttackInterval And GlobalFrameTime - .Contadores.IntervaloAtaque > AttackInterval
    End With
End Function

Public Function GetLinearDamageBonus(ByVal NpcIndex As Integer) As Integer
    GetLinearDamageBonus = NpcList(NpcIndex).Modifiers.PhysicalDamageLinearBonus
End Function

Public Sub SetBlockTileState(ByVal NpcIndex As Integer, ByVal Block As Boolean)
    Dim CurrentValue As Boolean
    With NpcList(NpcIndex)
        CurrentValue = IsSet(.flags.StatusMask, e_StatusMask.eDontBlockTile)
        If CurrentValue = Block Then Exit Sub
        If Block Then
            Call SetMask(.flags.StatusMask, e_StatusMask.eDontBlockTile)
        Else
            Call UnsetMask(.flags.StatusMask, e_StatusMask.eDontBlockTile)
        End If
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareUpdateCharValue(.Char.charindex, e_CharValue.eDontBlockTile, Block))
    End With
End Sub

Public Function GetDefenseBonus(ByVal NpcIndex As Integer) As Integer
    GetDefenseBonus = NpcList(NpcIndex).Modifiers.DefenseBonus
End Function

' Retorna el usuario que esta atacando al NPC actualmente (medido con tiempo)
Public Function GetOwnedBy(ByVal NpcIndex As Integer) As Integer
    GetOwnedBy = 0
    With NpcList(NpcIndex).flags
        If .AttackedBy = vbNullString Then Exit Function
        If GlobalFrameTime - .AttackedTime > IntervaloNpcOwner Then Exit Function
        Dim Attacker As t_UserReference: Attacker = NameIndex(.AttackedBy)
        If Not IsValidUserRef(Attacker) Then Exit Function
        GetOwnedBy = Attacker.ArrayIndex
    End With
End Function

' Retorna si un NPC puede atacar un usuario diferente al que lo esta atacando
Public Function CanAttackNotOwner(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    Dim AttackResult As t_AttackInteractionResult
    AttackResult = UserCanAttackNpc(UserIndex, NpcIndex)
    ' Si el usuario puede atacar al NPC
    If AttackResult.CanAttack Then
        ' Lo atacamos solo si puede atacar sin hacerse PK (no lo forzamos a hacerse PK)
        CanAttackNotOwner = Not AttackResult.TurnPK
    Else
        ' En caso que el usuario no pueda atacar al NPC, este debe ignorarlo a el
        ' Excepto que no pueda atacar por los siguientes motivos: esta montado, esta fuera de su campo de vision
        CanAttackNotOwner = AttackResult.Result = eMounted Or AttackResult.Result = eOutOfRange
    End If
End Function
