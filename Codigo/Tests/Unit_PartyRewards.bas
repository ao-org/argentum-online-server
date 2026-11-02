Attribute VB_Name = "Unit_PartyRewards"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
#If UNIT_TEST = 1 Then
Public SynchronousPersistence As Boolean
Private capturing As Boolean
Private packets As Collection
Private packetUsers As Collection
Private fixtureVersion As Integer
Private victim As t_Npc

Public Function CapturePacket(ByVal recipient As Integer, ByRef queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    Dim writer As Network.Writer, bytes() As Byte, payload As Variant
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes): Call writer.Clear()
    payload = bytes: Call packets.Add(payload): Call packetUsers.Add(recipient)
    queued = True: CapturePacket = True
End Function

Private Sub Expect(ByVal value As Boolean, ByVal label As String)
    If Not value Then Err.Raise vbObjectError + 851, "Party rewards", label
End Sub

Private Function CountPackets(ByVal packetId As Long, Optional ByVal recipient As Integer = 0) As Long
    Dim i As Long, bytes() As Byte
    For i = 1 To packets.Count
        bytes = packets(i)
        If UBound(bytes) >= 1 Then
            If CLng(bytes(0)) + 256& * bytes(1) = packetId Then
                If recipient = 0 Or packetUsers(i) = recipient Then CountPackets = CountPackets + 1
            End If
        End If
    Next i
End Function

Private Sub AcceptQuest(ByVal userIndex As Integer, ByVal slot As Integer, ByVal quest As Integer)
    With UserList(userIndex).QuestStats.Quests(slot)
        .QuestIndex = quest: .Dirty = False
        ReDim .NPCsKilled(1 To 1): ReDim .NPCsTarget(1 To 1)
    End With
End Sub

Private Sub Fixture()
    Call Unit_NpcBow.InitializeTestWorld()
    Call ResetTurnInTestCounters()
    Call SetFeatureToggle("collectible_cards", False)
    ReDim QuestList(0 To 3): ReDim NpcInfoCache(1 To 20)
    fixtureVersion = fixtureVersion + 1
    Dim i As Integer, q As Integer, blank As t_Npc
    victim = blank
    For q = 1 To 3
        With QuestList(q)
            .nombre = "Party quest " & q: .TalkTo = 11
            .RequiredNPCs = 1: ReDim .RequiredNPC(1 To 1)
            .RequiredNPC(1).NpcIndex = 10: .RequiredNPC(1).Amount = 2
            .RequiredTargetNPCs = 1: ReDim .RequiredTargetNPC(1 To 1)
            .RequiredTargetNPC(1).NpcIndex = 10: .RequiredTargetNPC(1).Amount = 1
        End With
    Next q
    QuestList(3).RequiredNPC(1).NpcIndex = 12
    NpcInfoCache(11).Exists = True: NpcInfoCache(11).Name = "Recipient"
    LastUser = 4
    For i = 1 To 4
        With UserList(i)
            .Id = 200 + i: .VersionId = fixtureVersion: .Name = "party fixture " & i
            .flags.UserLogged = True: .ConnectionDetails.ConnIDValida = True
            .pos.Map = 1: .pos.x = 50: .pos.y = 50
            .Stats.ELV = 10: .Stats.MinHp = 100: .Stats.MaxHp = 100
            .Stats.UserAtributos(e_Atributos.Carisma) = 0
            .Stats.UserSkills(e_Skill.liderazgo) = 0
            .CurrentInventorySlots = 3
            .Grupo.EnGrupo = True: .Grupo.Id = 77
        End With
        Call AcceptQuest(i, 1, 1)
    Next i
    ' Versioned references must be captured after every user has its session ID.
    For i = 1 To 4
        Call SetUserRef(UserList(i).Grupo.Lider, 3)
    Next i
    With UserList(3).Grupo
        .CantidadMiembros = 3
        Call SetUserRef(.Miembros(1), 3)
        Call SetUserRef(.Miembros(2), 1)
        Call SetUserRef(.Miembros(3), 2)
    End With
    UserList(3).pos.x = 80
    For i = 1 To STAT_MAXELV
        ExpLevelUp(i) = MAXEXP
    Next i
    With victim
        .Numero = 10: .Name = "Required creature": .npcType = Comun
        .pos.Map = 1: .pos.x = 51: .pos.y = 50
        .Stats.MaxHp = 100: .Stats.MinHp = 100: .GiveEXP = 100
        .flags.ExpCount = 100: .flags.Respawn = 1
        ReDim .Char.Animation(0 To 0)
    End With
    Call ResetQuestTurnInRecipients()
    Set packets = New Collection: Set packetUsers = New Collection: capturing = True
End Sub

Private Sub Drops(Optional ByVal amount As Integer = 3, Optional ByVal required As Integer = 5)
    With QuestList(1)
        .RequiredOBJs = 1: ReDim .RequiredOBJ(1 To 1)
        .RequiredOBJ(1).ObjIndex = 50: .RequiredOBJ(1).Amount = required
    End With
    With victim
        .NumDropQuest = 1: ReDim .DropQuest(1 To 1)
        .DropQuest(1).QuestIndex = 1: .DropQuest(1).ObjIndex = 50
        .DropQuest(1).Amount = amount: .DropQuest(1).Probabilidad = 1
    End With
    ObjData(50).Name = "Quest item"
End Sub

Private Sub InstallVictim(ByVal index As Integer)
    NpcList(index) = victim
    NpcList(index).flags.NPCActive = True: NpcList(index).VersionId = 1
    NpcList(index).Char.charindex = index
    CharList(index) = index
    MapData(51, 50, 1).NpcIndex = index
End Sub

Private Sub PersistenceRoundTrip()
    On Error GoTo Failed
    Dim original As ADODB.Connection, database As New ADODB.Connection
    Dim oldAsync As Byte, i As Integer, blankQuests As t_QuestStats, blankInventory As t_Inventario
    Set original = Connection: oldAsync = Current_async
    ' Existing SQLite ODBC dependency, private in-memory schema; never writes the
    ' validator/game database. Execute is synchronous only inside this test scope.
    Call database.Open("DRIVER={SQLite3 ODBC Driver};DATABASE=:memory:")
    Call database.Execute("CREATE TABLE quest (user_id INTEGER, number INTEGER, quest_id INTEGER, npcs TEXT, npcstarget TEXT, PRIMARY KEY(user_id,number))")
    Call database.Execute("CREATE TABLE inventory_item (user_id INTEGER, number INTEGER, item_id INTEGER, amount INTEGER, is_equipped INTEGER, elemental_tags INTEGER, PRIMARY KEY(user_id,number))")
    Set Connection = database: SynchronousPersistence = True
    For i = 1 To 2
        Call SavePartyQuestRewardsForTest(i)
    Next i
    For i = 1 To 2
        UserList(i).QuestStats = blankQuests: UserList(i).invent = blankInventory
        Call Expect(LoadPartyQuestRewardsForTest(i), "production inventory reload")
        Call Expect(UserList(i).QuestStats.Quests(1).QuestIndex = 1 And UserList(i).QuestStats.Quests(1).NPCsKilled(1) = 1, "quest serializer/loader round trip")
        Call Expect(CountInventoryObjects(i, 50) = 3, "normal inventory serializer/loader round trip")
    Next i
    SynchronousPersistence = False: Set Connection = original: Current_async = oldAsync
    Call database.Close()
    Exit Sub
Failed:
    Dim message As String: message = Err.Description
    SynchronousPersistence = False: Set Connection = original: Current_async = oldAsync
    On Error Resume Next
    Call database.Close()
    On Error GoTo 0
    Err.Raise vbObjectError + 852, "Party persistence", message
End Sub

Private Sub SetXpMultiplier(ByVal multiplier As Integer)
    Dim config As New clsIniManager
    Call config.Initialize(App.Path & "\Example.Configuracion.ini")
    Call config.ChangeValue("CONFIGURACIONES", "ExpMult", CStr(multiplier))
    Set SvrConfig = New ServerConfig
    Call SvrConfig.LoadSettings(App.Path & "\Example.Configuracion.ini", config)
End Sub

Private Sub RestoreFixtureParty(ByVal groupId As Long)
    Dim i As Integer
    For i = 1 To 3
        With UserList(i).Grupo
            .EnGrupo = True: .Id = groupId
            Call SetUserRef(.Lider, 3)
        End With
    Next i
    With UserList(3).Grupo
        .CantidadMiembros = 3
        Call SetUserRef(.Miembros(1), 3)
        Call SetUserRef(.Miembros(2), 1)
        Call SetUserRef(.Miembros(3), 2)
    End With
End Sub

Private Sub ExpectFairPair(ByVal total As Long)
    Call Expect(UserList(1).Stats.Exp + UserList(2).Stats.Exp = total, "exact cumulative pool accounting")
    Call Expect(Abs(UserList(1).Stats.Exp - UserList(2).Stats.Exp) <= 1, "cumulative difference at most one")
End Sub

Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Fixture()
    Dim recipients As t_PartyRewardRecipients, i As Integer, n As Integer, pet As Integer
    Dim before As Long, beforeFirst As Long, beforeSecond As Long, beforeFourth As Long
    Dim lowest As Long, highest As Long, iteration As Integer, member As Integer, attacker As Integer
    Dim result As e_DamageResult, config As clsIniManager
    Select Case index
    Case 1
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 1, "killer kill")
        Call Expect(UserList(2).QuestStats.Quests(1).NPCsKilled(1) = 1, "support without damage shares kill")
        Call Expect(UserList(3).QuestStats.Quests(1).NPCsKilled(1) = 0, "absent leader excluded")
        Call Expect(UserList(2).QuestStats.Quests(1).NPCsTarget(1) = 0, "no visit sharing")
    Case 2
        Call AcceptQuest(2, 2, 2): Call AcceptQuest(2, 3, 3)
        UserList(1).QuestStats.Quests(1).QuestIndex = 0
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(UserList(2).QuestStats.Quests(1).NPCsKilled(1) = 1 And UserList(2).QuestStats.Quests(2).NPCsKilled(1) = 1, "multiple accepted quests")
        Call Expect(UserList(2).QuestStats.Quests(3).NPCsKilled(1) = 0, "exact unrelated template")
        Call Expect(UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 0, "no auto acceptance")
    Case 3, 4
        UserList(2).pos.x = 50 + 19 + (index - 3)
        Call CollectPartyRewardRecipients(1, recipients, 1)
        Call Expect(recipients.Count = 5 - index, "19 qualifies, 20 does not")
    Case 5 To 11
        Select Case index
        Case 5: UserList(2).pos.Map = 2
        Case 6: UserList(2).flags.Muerto = 1
        Case 7: UserList(2).ConnectionDetails.ConnIDValida = False
        Case 8: UserList(2).flags.UserLogged = False
        Case 9: UserList(2).VersionId = fixtureVersion + 1
        Case 10: UserList(2).Grupo.Id = 78
        Case 11: Call SetUserRef(UserList(2).Grupo.Lider, 4)
        End Select
        Call CollectPartyRewardRecipients(1, recipients, 1)
        Call Expect(recipients.Count = 1 And recipients.Users(1).ArrayIndex = 1, "excluded invalid member")
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(UserList(2).QuestStats.Quests(1).NPCsKilled(1) = 0, "no invalid quest award")
    Case 12
        UserList(3).pos.x = 50
        With UserList(3).Grupo
            .CantidadMiembros = 6
            Call SetUserRef(.Miembros(4), 2): Call SetUserRef(.Miembros(5), 1): Call SetUserRef(.Miembros(6), 3)
        End With
        Call CollectPartyRewardRecipients(1, recipients, 1)
        Call Expect(recipients.Count = 3, "leader and duplicate roster deduped")
        Call AwardNpcPartyQuests(1, victim)
        For i = 1 To 3
            Call Expect(UserList(i).QuestStats.Quests(1).NPCsKilled(1) = 1, "one award per identity")
        Next i
    Case 13
        UserList(1).Grupo.Lider.VersionId = fixtureVersion - 1
        Call CollectPartyRewardRecipients(1, recipients)
        Call Expect(recipients.Count = 1 And Not recipients.IsParty, "bad party preserves solo")
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 1, "solo quest fallback")
    Case 14
        Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 50 And UserList(2).Stats.Exp = 50 And UserList(3).Stats.Exp = 0, "no absent-leader divisor")
        Call Expect(victim.flags.ExpCount = 0, "pool consumed once")
        Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 50, "no duplicate remainder payment")
    Case 15
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        Call CalcularDarExp(1, n, 21)
        Call Expect(NpcList(n).flags.ExpCount = 79, "damage consumes base21")
        Call Expect(UserList(1).Stats.Exp = 11 And UserList(2).Stats.Exp = 10, "base remainder allocated without manufacturing XP")
        victim = NpcList(n)
        Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp + UserList(2).Stats.Exp = 100, "damage plus remainder conserve base pool")
    Case 16
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        UserList(1).ConnectionDetails.ConnIDValida = False
        Call CalcularDarExp(1, n, 25)
        Call Expect(NpcList(n).flags.ExpCount = 100, "invalid credit cannot consume pool")
    Case 17
        Call Drops()
        UserList(1).invent.Object(1).ObjIndex = 50: UserList(1).invent.Object(1).Amount = 4
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(CountInventoryObjects(1, 50) = 5 And CountInventoryObjects(2, 50) = 3, "independent capped personal grants")
        Call Expect(CountPackets(ServerPacketID.ePlayWave, 1) = 1 And CountPackets(ServerPacketID.ePlayWave, 2) = 1, "positive grant sound each")
    Case 18
        Call Drops(3, 5)
        victim.NumDropQuest = 2: ReDim Preserve victim.DropQuest(1 To 2)
        victim.DropQuest(2) = victim.DropQuest(1)
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(CountInventoryObjects(1, 50) = 5 And CountInventoryObjects(2, 50) = 5, "overlapping entries recheck missing count")
    Case 19
        Call Drops()
        UserList(1).CurrentInventorySlots = 1
        UserList(1).invent.Object(1).ObjIndex = 51: UserList(1).invent.Object(1).Amount = GetMaxInvOBJ()
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(CountInventoryObjects(1, 50) = 0 And CountInventoryObjects(2, 50) = 3, "one full inventory does not block peer")
        Call Expect(CountPackets(ServerPacketID.ePlayWave, 1) = 0, "failed insertion has no success audio")
        UserList(1).invent.Object(1).ObjIndex = 0: UserList(1).invent.Object(1).Amount = 0
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(CountInventoryObjects(1, 50) = 3, "later kill remains eligible")
    Case 20
        Call Drops(5, 10002)
        UserList(1).CurrentInventorySlots = 1
        UserList(1).invent.Object(1).ObjIndex = 50: UserList(1).invent.Object(1).Amount = GetMaxInvOBJ() - 1
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(CountInventoryObjects(1, 50) = GetMaxInvOBJ(), "partial final stack capacity")
        before = packets.Count
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(CountInventoryObjects(1, 50) = GetMaxInvOBJ(), "full stack never overflows")
        Call Expect(CountPackets(ServerPacketID.ePlayWave, 1) = 1, "only actual stack insertion sounds")
    Case 21
        For i = 1 To 2
            UserList(i).QuestStats.Quests(1).NPCsKilled(1) = 2
        Next i
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(packets.Count = 0 And Not UserList(1).flags.ModificoQuests, "completed requirements produce no repeated notice or presentation")
    Case 22
        victim.npcType = DummyTarget
        Call AwardNpcPartyQuests(1, victim): Call AwardNpcDeathExperience(1, victim)
        Call Expect(packets.Count = 0 And UserList(1).Stats.Exp = 0, "training exclusion")
        victim.npcType = Comun: Call SetUserRef(victim.MaestroUser, 4)
        Call AwardNpcPartyQuests(1, victim): Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(2).QuestStats.Quests(1).NPCsKilled(1) = 0 And UserList(2).Stats.Exp = 0, "player-owned victim exclusion")
    Case 23
        Call AwardNpcPartyQuests(-1, victim): Call AwardNpcDeathExperience(-1, victim)
        Call Expect(packets.Count = 0 And victim.flags.ExpCount = 100, "environment has no player reward")
    Case 24
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        Call Drops(): NpcList(n) = victim: Call InstallVictim(n)
        Call MuereNpc(n, 1)
        Call Expect(Not NpcList(n).flags.NPCActive, "real death retires live slot")
        Call Expect(UserList(2).QuestStats.Quests(1).NPCsKilled(1) = 1 And CountInventoryObjects(2, 50) = 3, "death uses saved template/map/drop data")
        Call Expect(UserList(1).Stats.Exp + UserList(2).Stats.Exp = 100, "real death shares saved remaining base")
    Case 25, 26
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        pet = GetNextAvailableNpc()
        NpcList(pet).flags.NPCActive = True
        If index = 25 Then Call SetUserRef(NpcList(pet).MaestroUser, 1)
        NpcList(n).Stats.MinHp = 1: NpcList(n).flags.ExpCount = 0
        result = NPCs.DoDamageOrHeal(n, pet, eNpc, -1, e_pet, 0)
        Call Expect(result = eDead, "production pet/environment kill path")
        If index = 25 Then
            Call Expect(UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 1 And UserList(2).QuestStats.Quests(1).NPCsKilled(1) = 1, "pet owner party gets one kill")
        Else
            Call Expect(UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 0 And UserList(2).QuestStats.Quests(1).NPCsKilled(1) = 0, "unowned NPC kill no player reward")
        End If
    Case 27
        Call Drops(1, 1)
        QuestList(1).RequiredNPC(1).Amount = 1: QuestList(1).RequiredTargetNPCs = 0
        For i = 1 To 2
            Call SendQuestTurnInSnapshot(i, True)
        Next i
        Call ProcessQuestTurnInUpdates()
        Call ResetTurnInTestCounters()
        Set packets = New Collection: Set packetUsers = New Collection
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(EvaluationCount = 0, "no intermediate readiness evaluation")
        Call ProcessQuestTurnInUpdates()
        Call Expect(EvaluationCount = 2, "one final readiness evaluation each")
        Call Expect(CountPackets(ServerPacketID.eQuestTurnIns, 1) = 1 And CountPackets(ServerPacketID.eQuestTurnIns, 2) = 1, "both members receive ready state")
        Call Expect(CountPackets(ServerPacketID.eQuestFeedback, 1) = 1 And CountPackets(ServerPacketID.eQuestFeedback, 2) = 1, "one completed feedback state each")
    Case 28
        Call Drops()
        Call UpdateSavedQuests(2): Call UpdateSavedInventory(2)
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(HaveQuestsChanged(2) And HasInventoryChanged(2), "normal persistence detects changes")
        Call Expect(UserList(2).QuestStats.Quests(1).Dirty And UserList(2).flags.ModificoInventario, "normal dirty hooks")
        Call UpdateSavedQuests(2): Call UpdateSavedInventory(2)
        Call Expect(UserList(2).Persist.LastQuests(1).NPCsKilled(1) = 1 And UserList(2).Persist.LastInventory(1).Amount = 3, "existing saved projections contain actual progress/items")
        Call Expect(Not HaveQuestsChanged(2) And Not HasInventoryChanged(2), "saved projections stable")
    Case 29
        victim.nivel = 4
        UserList(1).Stats.UserSkills(e_Skill.liderazgo) = 100
        Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = CLng(CLng(50 * 0.9) * SvrConfig.GetValue("LeadershipExpPartyBonus")), "level penalty then leadership once")
        Call Expect(UserList(2).Stats.Exp = 45, "per-recipient penalty")
    Case 30
        UserList(2).Stats.ELV = STAT_MAXELV
        Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 50 And UserList(2).Stats.Exp = 0 And victim.flags.ExpCount = 0, "existing max-level participation/no gain")
    Case 31
        UserList(1).Grupo.EnGrupo = False
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        Call CalcularDarExp(1, n, 20)
        victim = NpcList(n): Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 100 And UserList(2).Stats.Exp = 0, "solo damage and remainder unchanged")
    Case 32
        victim.pos.Map = 2
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(packets.Count = 0, "quest rewards require defeated NPC saved map")
    Case 33
        UserList(3).Grupo.Id = 78
        Call CollectPartyRewardRecipients(1, recipients)
        Call Expect(recipients.Count = 1 And Not recipients.IsParty, "leader identity mismatch preserves solo")
    Case 34
        Call SetFeatureToggle("collectible_cards", True)
        victim.CollectibleCardIndex = 1
        UserList(1).AccountCollectibleCardQuantities(1) = 1
        NpcList(1).CollectibleCardIndex = 2
        Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = CLng(50 * GetCardExpBonusForCard(1, 1)), "saved NPC card bonus, not recycled live slot")
        Call Expect(UserList(2).Stats.Exp = 50, "cards per recipient")
    Case 35
        Set config = New clsIniManager
        Call config.Initialize(App.Path & "\Example.Configuracion.ini")
        Call config.ChangeValue("CONFIGURACIONES", "ExpMult", "3")
        Set SvrConfig = New ServerConfig
        Call SvrConfig.LoadSettings(App.Path & "\Example.Configuracion.ini", config)
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        Call CalcularDarExp(1, n, 21)
        victim = NpcList(n): Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 153 And UserList(2).Stats.Exp = 147, "damage and remainder apply multiplier exactly once")
    Case 36
        UserList(1).Stats.Exp = MAXEXP - 10
        ExpLevelUp(10) = MAXEXP + 1
        victim.flags.ExpCount = MAXEXP
        Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = MAXEXP, "XP cap before integer addition")
    Case 37
        Call Drops()
        Call AcceptQuest(2, 1, 2)
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(CountInventoryObjects(1, 50) = 3 And CountInventoryObjects(2, 50) = 0, "quest-specific personal drop requires acceptance")
    Case 38
        Call Drops()
        UserList(2).invent.Object(1).ObjIndex = 50: UserList(2).invent.Object(1).Amount = 5
        Call AwardNpcPartyQuests(1, victim)
        Call Expect(CountInventoryObjects(2, 50) = 5 And CountPackets(ServerPacketID.ePlayWave, 2) = 0, "complete item requirement grants/sounds nothing")
    Case 39
        Call SetUserRef(UserList(3).Grupo.Lider, 4)
        Call CollectPartyRewardRecipients(1, recipients)
        Call Expect(recipients.Count = 1 And Not recipients.IsParty, "leader must identify itself")
    Case 40
        UserList(3).flags.UserLogged = False: UserList(3).ConnectionDetails.ConnIDValida = False
        Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 50 And UserList(2).Stats.Exp = 50, "valid absent leader identity need not receive XP")
    Case 41
        UserList(1).flags.Muerto = 1
        Call AwardNpcPartyQuests(1, victim): Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 0 And UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 0, "dead credit is not a recipient")
        Call Expect(UserList(2).Stats.Exp = 100 And UserList(2).QuestStats.Quests(1).NPCsKilled(1) = 1, "valid dead source retains live party credit")
    Case 42
        Call Drops()
        Call AwardNpcPartyQuests(1, victim)
        Call PersistenceRoundTrip()
    Case 43
        Call SetXpMultiplier(2)
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        For iteration = 1 To 100
            Call CalcularDarExp(1, n, 1)
            Call Expect(NpcList(n).flags.ExpCount = 100 - iteration, "one base unit debit per hit")
        Next iteration
        Call Expect(UserList(1).Stats.Exp = 100 And UserList(2).Stats.Exp = 100, "100 single-base hits at multiplier2 split100/100")
    Case 44
        victim.GiveEXP = 300: victim.Stats.MaxHp = 300: victim.Stats.MinHp = 300: victim.flags.ExpCount = 300
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        For iteration = 1 To 100
            Call CalcularDarExp(1, n, 3)
            Call ExpectFairPair(3& * iteration)
            Call Expect(NpcList(n).flags.ExpCount = 300 - 3& * iteration, "three-unit award pool debit")
        Next iteration
    Case 45
        ReDim Preserve UserList(1 To 6)
        For member = 1 To 6
            With UserList(member)
                .Id = 200 + member: .VersionId = fixtureVersion
                .flags.UserLogged = True: .ConnectionDetails.ConnIDValida = True
                .pos.Map = 1: .pos.x = 50: .pos.y = 50
                .Stats.ELV = 10: .Grupo.EnGrupo = True: .Grupo.Id = 77
                Call SetUserRef(.Grupo.Lider, 3)
            End With
            Call SetUserRef(UserList(3).Grupo.Miembros(member), member)
        Next member
        UserList(3).Grupo.CantidadMiembros = 6
        victim.GiveEXP = 120: victim.Stats.MaxHp = 120: victim.Stats.MinHp = 120: victim.flags.ExpCount = 120
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        For iteration = 1 To 120
            Call CalcularDarExp(1, n, 1)
            lowest = UserList(1).Stats.Exp: highest = lowest: before = 0
            For member = 1 To 6
                before = before + UserList(member).Stats.Exp
                If UserList(member).Stats.Exp < lowest Then lowest = UserList(member).Stats.Exp
                If UserList(member).Stats.Exp > highest Then highest = UserList(member).Stats.Exp
            Next member
            Call Expect(before = iteration And highest - lowest <= 1, "six recipients conserve/fairly rotate small awards")
        Next iteration
        For member = 1 To 6
            Call Expect(UserList(member).Stats.Exp = 20, "no six-member starvation")
        Next member
    Case 46
        victim.GiveEXP = 120: victim.Stats.MaxHp = 120: victim.Stats.MinHp = 120: victim.flags.ExpCount = 120
        n = GetNextAvailableNpc(): Call InstallVictim(n)
        For iteration = 1 To 120
            attacker = 2
            If iteration Mod 3 = 0 Then attacker = 1
            Call SetUserRef(UserList(3).Grupo.Miembros(2), attacker)
            Call SetUserRef(UserList(3).Grupo.Miembros(3), 3 - attacker)
            Call CalcularDarExp(attacker, n, 1)
            Call ExpectFairPair(iteration)
        Next iteration
        Call Expect(NpcList(n).flags.ExpCount = 0, "changing attacker/roster never changes debit")
    Case 47
        n = GetNextAvailableNpc()
        For iteration = 1 To 20
            victim.GiveEXP = 3: victim.Stats.MaxHp = 3: victim.Stats.MinHp = 3: victim.flags.ExpCount = 3
            Call InstallVictim(n)
            Call CalcularDarExp(1, n, 1)
            Call Expect(NpcList(n).flags.ExpCount = 2, "mixed award damage debit")
            victim = NpcList(n)
            Call AwardNpcDeathExperience(1, victim)
            Call Expect(victim.flags.ExpCount = 0, "mixed award remainder consumed")
            Call ExpectFairPair(3& * iteration)
        Next iteration
    Case 48
        victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        For iteration = 1 To 5
            Select Case iteration
            Case 1: UserList(2).flags.Muerto = 1
            Case 2: UserList(2).flags.Muerto = 0: UserList(2).pos.x = 70
            Case 3: UserList(2).pos.x = 50: UserList(2).pos.Map = 2
            Case 4: UserList(2).pos.Map = 1: UserList(2).ConnectionDetails.ConnIDValida = False
            Case 5: UserList(2).ConnectionDetails.ConnIDValida = True: UserList(2).Grupo.Id = 78
            End Select
            victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
            Call Expect(UserList(2).Stats.Exp = 0, "ineligible member earns no unit/debt")
        Next iteration
        UserList(2).Grupo.Id = 77
        victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 6 And UserList(2).Stats.Exp = 1, "return resumes rotation without catch-up payment")
        UserList(3).Grupo.CantidadMiembros = 4
        Call SetUserRef(UserList(3).Grupo.Miembros(4), 4)
        beforeFirst = UserList(1).Stats.Exp: beforeSecond = UserList(2).Stats.Exp
        For iteration = 1 To 9
            victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Next iteration
        Call Expect(UserList(1).Stats.Exp - beforeFirst = 3 And UserList(2).Stats.Exp - beforeSecond = 3 And UserList(4).Stats.Exp = 3, "joined eligible member rotates fairly")
        UserList(2).Grupo.EnGrupo = False
        beforeFirst = UserList(1).Stats.Exp: beforeSecond = UserList(2).Stats.Exp: beforeFourth = UserList(4).Stats.Exp
        For iteration = 1 To 6
            victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Next iteration
        Call Expect(UserList(1).Stats.Exp - beforeFirst = 3 And UserList(4).Stats.Exp - beforeFourth = 3 And UserList(2).Stats.Exp = beforeSecond, "departed member receives no allocated XP")
    Case 49
        victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Call FinalizarGrupo(3): Call GroupCreateSuccess(3)
        Call RestoreFixtureParty(UserList(3).Grupo.Id)
        victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 2 And UserList(2).Stats.Exp = 0, "new party does not inherit old remainder cursor")
    Case 50
        victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Call IncreaseVersionId(3)
        Call RestoreFixtureParty(77)
        victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 2 And UserList(2).Stats.Exp = 0, "leader reconnect resets party-session cursor")
    Case 51, 52
        victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Call IncreaseVersionId(2)
        If index = 52 Then UserList(2).Id = 999
        Call SetUserRef(UserList(3).Grupo.Miembros(3), 2)
        victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 2 And UserList(2).Stats.Exp = 0, "reconnected/reused slot does not inherit old next identity")
        victim.flags.ExpCount = 1: Call AwardNpcDeathExperience(1, victim)
        Call Expect(UserList(1).Stats.Exp = 2 And UserList(2).Stats.Exp = 1, "new session participates without owed XP")
    Case 53
        For iteration = 1 To 20
            victim.flags.ExpCount = 1
            n = GetNextAvailableNpc(): Call InstallVictim(n)
            Call MuereNpc(n, 1)
            Call Expect(Not NpcList(n).flags.NPCActive, "real repeated kill releases slot")
            Call ExpectFairPair(iteration)
        Next iteration
    End Select
    capturing = False
    Call UnitTesting.Progress("Party rewards " & index & " passed")
    Call UnitTesting.RunTest("Party rewards " & index, True)
    Exit Sub
Failed:
    Dim failure As String: failure = Err.Description
    capturing = False
    Call UnitTesting.Progress("Party rewards " & index & " failed: " & failure)
    Call UnitTesting.RunTestError("Party rewards " & index, failure)
End Sub

Public Function TestSuite() As Boolean
    On Error GoTo Failed
    Dim before As Long, i As Integer, oldExp(1 To STAT_MAXELV) As Long, oldCards As Boolean
    before = UnitTesting.GetFailedTests(): oldCards = IsFeatureEnabled("collectible_cards")
    For i = 1 To STAT_MAXELV
        oldExp(i) = ExpLevelUp(i)
    Next i
    For i = 1 To 53
        Call RunCase(i)
    Next i
    For i = 1 To STAT_MAXELV
        ExpLevelUp(i) = oldExp(i)
    Next i
    Call SetFeatureToggle("collectible_cards", oldCards)
    TestSuite = UnitTesting.GetFailedTests() = before
    Exit Function
Failed:
    capturing = False
    Call UnitTesting.RunTestError("Party rewards suite", Err.Description)
End Function
#End If
