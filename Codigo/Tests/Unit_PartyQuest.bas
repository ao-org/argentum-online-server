Attribute VB_Name = "Unit_PartyQuest"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private packetUsers As Collection
Private database As ADODB.Connection
Private oldConnection As ADODB.Connection
Private fileName As String
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
    If Not value Then Err.Raise vbObjectError + 861, "Party quest", label & "; save=" & LastPartySaveFailureForTest
End Sub
Private Function Scalar(ByVal sql As String) As Long
    Dim rows As ADODB.Recordset
    Set rows = database.Execute(sql)
    Scalar = rows.Fields(0).value
    Call rows.Close()
End Function
Private Sub Fixture()
    ' The production loop delivers ADO completion events at DoEvents. Drain prior
    ' suites' real legacy writes before requiring an immediate isolated save.
    Dim waitStarted As Long
    waitStarted = GetTickCountRaw()
    Do While Not PartyQuestLegacyWritesComplete()
        DoEvents
        If TicksElapsed(waitStarted, GetTickCountRaw()) > 1000 Then Err.Raise vbObjectError + 861, , "Legacy write fence did not drain"
    Loop
    Set Builder = New cStringBuilder
    Call InitializeNpcIndexHeap()
    MapData(51, 50, 1).NpcIndex = 0
    Call ResetPartyQuestTests(): Call ResetTurnInTestCounters()
    Set oldConnection = Connection
    fileName = Environ$("TEMP") & "\ao-party-hunt-test-" & GetTickCountRaw() & ".db"
    Set database = New ADODB.Connection
    database.CursorLocation = adUseClient
    Call database.Open("Driver={SQLite3 ODBC Driver};Database=" & fileName & ";")
    Set Connection = database
    Call database.Execute("PRAGMA foreign_keys=ON")
    Call database.BeginTrans()
    Call database.Execute("CREATE TABLE user(id INTEGER PRIMARY KEY AUTOINCREMENT,gold INTEGER NOT NULL DEFAULT 0)")
    Call database.Execute("CREATE TABLE quest(user_id INTEGER REFERENCES user(id) ON DELETE CASCADE,number INTEGER,quest_id INTEGER,npcs TEXT,npcstarget TEXT,PRIMARY KEY(user_id,number))")
    Call database.Execute("CREATE TABLE quest_done(user_id INTEGER REFERENCES user(id) ON DELETE CASCADE,quest_id INTEGER,PRIMARY KEY(user_id,quest_id))")
    ' Exercise the production migration, not a parallel test schema.
    Dim f As Integer, sql As String, statement As Variant
    f = FreeFile: Open App.Path & "\ScriptsDB\20261103-01-party-quest.sql" For Binary Access Read As #f
    sql = Space$(LOF(f)): Get #f, , sql: Close #f
    For Each statement In Split(sql, ";")
        If Len(Trim$(statement)) > 0 Then Call database.Execute(CStr(statement))
    Next statement
    Call BuildCharacterSaveSchema()
    ReDim UserList(1 To 7): ReDim QuestList(0 To 2): ReDim NpcInfoCache(1 To 500)
    fixtureVersion = fixtureVersion + 1
    Dim i As Integer, blankNpc As t_Npc
    For i = 1 To 3: NpcList(i) = blankNpc: CharList(i) = 0: Next i
    i = GetNextAvailableNpc()
    For i = 1 To 7
        With UserList(i)
            .AccountID = 700 + i: .Id = 700 + i: .VersionId = fixtureVersion: .name = "Hunter" & i
            .flags.UserLogged = True: .ConnectionDetails.ConnIDValida = True
            .flags.Privilegios = e_PlayerType.User: .Stats.ELV = 10
            .Stats.MinHp = 100: .Stats.MaxHp = 100: .Stats.GLD = 10
            .pos.Map = 1: .pos.x = 50: .pos.y = 50: .Char.charindex = i
            .CurrentInventorySlots = 3
        End With
        Call database.Execute("INSERT INTO user(id,account_id,name,gold) VALUES(" & UserList(i).Id & "," & UserList(i).AccountID & ",'Hunter" & i & "',10)")
        Call database.Execute("INSERT INTO account(id) VALUES(" & UserList(i).AccountID & ")")
    Next i
    Call database.CommitTrans()
    Unit_PartyRewards.SynchronousPersistence = True
    For i = 1 To 2
        With QuestList(i)
            .nombre = "The Pack Hunt": .RequiredLevel = 1: .TalkTo = 432
            .RequiredNPCs = 1: ReDim .RequiredNPC(1 To 1)
            .RequiredNPC(1).NpcIndex = 500: .RequiredNPC(1).amount = 6
        End With
    Next i
    QuestList(1).PartyMinimum = 2: QuestList(1).RewardGLD = 360
    QuestList(2).nombre = "Personal wolf hunt"
    NpcInfoCache(432).Exists = True: NpcInfoCache(432).Name = "Alaric von Nix"
    NpcInfoCache(500).Exists = True: NpcInfoCache(500).Name = "Lobo"
    With NpcList(1)
        .Numero = 432: .VersionId = fixtureVersion: .NpcType = e_NPCType.Quest: .flags.NPCActive = True
        .pos = UserList(1).pos: .NumQuest = 2: ReDim .QuestNumber(1 To 2)
        .QuestNumber(1) = 1: .QuestNumber(2) = 2: .Char.charindex = 50
    End With
    Dim blankState As t_Npc
    victim = blankState: victim.Numero = 500: victim.pos = UserList(1).pos: victim.Name = "Lobo"
    Call SetFeatureToggle("collectible_cards", False)
    LastUser = 7
    Dim config As New clsIniManager
    Call config.Initialize(App.Path & "\Example.Configuracion.ini")
    Call config.ChangeValue("CONFIGURACIONES", "OroMult", "1")
    Set SvrConfig = New ServerConfig
    Call SvrConfig.LoadSettings(App.Path & "\Example.Configuracion.ini", config)
    Set packets = New Collection: Set packetUsers = New Collection: capturing = True
    Call Group(1, 2)
    Call Expect(InitializePartyQuestPersistence(), "initialize durable run ID range on login/startup, not kill")
End Sub
Private Sub BuildCharacterSaveSchema()
    Call Contruir_Querys()
    Dim fields As Object, text As String, item As Variant, column As String
    Set fields = CreateObject("Scripting.Dictionary")
    fields.CompareMode = vbTextCompare
    fields("id") = True: fields("gold") = True: fields("party_quest_save_revision") = True
    text = Mid$(QUERY_LOAD_MAINPJ, 8): text = Left$(text, InStr(text, " FROM ") - 1)
    For Each item In Split(text, ",")
        column = Trim$(CStr(item))
        If Not fields.Exists(column) Then
            Call database.Execute("ALTER TABLE user ADD COLUMN [" & column & "] NUMERIC DEFAULT 0")
            fields(column) = True
        End If
    Next item
    text = Mid$(QUERY_UPDATE_MAINPJ, InStr(QUERY_UPDATE_MAINPJ, "SET ") + 4)
    text = Left$(text, InStr(text, "WHERE ") - 1)
    For Each item In Split(text, ",")
        column = Trim$(Left$(CStr(item), InStr(CStr(item), "=") - 1))
        If Not fields.Exists(column) Then
            Call database.Execute("ALTER TABLE user ADD COLUMN [" & column & "] NUMERIC DEFAULT 0")
            fields(column) = True
        End If
    Next item
    Call database.Execute("CREATE TABLE account(id INTEGER PRIMARY KEY,last_ip TEXT,is_active_patron INTEGER DEFAULT 0,offline_patron_credits INTEGER DEFAULT 0)")
    Call database.Execute("CREATE TABLE spell(user_id INTEGER REFERENCES user(id) ON DELETE CASCADE,number INTEGER,spell_id INTEGER,PRIMARY KEY(user_id,number))")
    Call database.Execute("CREATE TABLE inventory_item(user_id INTEGER REFERENCES user(id) ON DELETE CASCADE,number INTEGER,item_id INTEGER,amount INTEGER,is_equipped INTEGER,elemental_tags TEXT,PRIMARY KEY(user_id,number))")
    Call database.Execute("CREATE TABLE bank_item(user_id INTEGER REFERENCES user(id) ON DELETE CASCADE,number INTEGER,item_id INTEGER,amount INTEGER,elemental_tags TEXT,PRIMARY KEY(user_id,number))")
    Call database.Execute("CREATE TABLE skillpoint(user_id INTEGER REFERENCES user(id) ON DELETE CASCADE,number INTEGER,value INTEGER,PRIMARY KEY(user_id,number))")
    Call database.Execute("CREATE TABLE pet(user_id INTEGER REFERENCES user(id) ON DELETE CASCADE,number INTEGER,pet_id INTEGER,PRIMARY KEY(user_id,number))")
    Call database.Execute("CREATE TABLE inventory_item_skins(user_id INTEGER REFERENCES user(id) ON DELETE CASCADE,skin_id INTEGER,type_skin INTEGER,skin_equipped INTEGER,PRIMARY KEY(user_id,skin_id))")
End Sub

Private Function LastSharedCount(ByVal userIndex As Integer) As Integer
    Dim i As Long, bytes() As Byte, packet As New Network.Reader, ignored As Long
    LastSharedCount = -1
    For i = packets.count To 1 Step -1
        If packetUsers(i) = userIndex Then
            bytes = packets(i): Call packet.SetData(bytes)
            If packet.ReadInt16() = ServerPacketID.ePartyQuestState Then
                ignored = packet.ReadInt32(): ignored = packet.ReadInt32()
                ignored = packet.ReadInt8(): ignored = packet.ReadInt8()
                ignored = packet.ReadInt32(): ignored = packet.ReadInt32()
                ignored = packet.ReadInt16(): ignored = packet.ReadInt16(): ignored = packet.ReadInt16(): ignored = packet.ReadInt16()
                LastSharedCount = packet.ReadInt16(): Exit Function
            End If
        End If
    Next i
End Function

Private Sub ReloadCharacter(ByVal userIndex As Integer)
    Dim id As Long, account As Long, blankState As t_User, why As String
    id = UserList(userIndex).Id: account = UserList(userIndex).AccountID
    Call PartyQuestLogout(userIndex)
    UserList(userIndex) = blankState
    With UserList(userIndex)
        .Id = id: .AccountID = account: .VersionId = fixtureVersion + 1
    End With
    Call Expect(LoadCharacterFromDB(userIndex, why), "production full character load: " & why)
    UserList(userIndex).flags.UserLogged = True: UserList(userIndex).ConnectionDetails.ConnIDValida = True
    UserList(userIndex).CurrentInventorySlots = 3
End Sub

Private Sub Group(ByVal leader As Integer, ByVal count As Integer)
    Dim i As Integer
    For i = leader To leader + count - 1
        With UserList(i).Grupo
            .EnGrupo = True: .Id = 77 + leader: Call SetUserRef(.Lider, leader)
        End With
    Next i
    With UserList(leader).Grupo
        .CantidadMiembros = count
        For i = 1 To count: Call SetUserRef(.Miembros(i), leader + i - 1): Next i
    End With
End Sub
Private Sub Propose(ByVal leader As Integer)
    Dim writer As New Network.Writer, data() As Byte
    Call SetNpcRef(UserList(leader).flags.TargetNPC, 1)
    Call writer.WriteInt8(1): Call writer.GetData(data): Set Protocol.reader = New Network.Reader: Call Protocol.reader.SetData(data)
    Call Protocol.HandleQuestAccept(leader)
End Sub
Private Sub Action(ByVal u As Integer, ByVal request As Long, ByVal action As Byte, ByVal identity As Long, Optional ByVal trailing As Boolean = False)
    ' Production packet dispatcher including bounds, authentication and maintained ID.
    Dim writer As New Network.Writer, data() As Byte, incoming As New Network.Reader
    Call writer.WriteInt16(ClientPacketID.ePartyQuestAction): Call writer.WriteInt32(request)
    Call writer.WriteInt8(action): Call writer.WriteInt32(identity): Call writer.WriteInt16(1)
    If trailing Then Call writer.WriteInt8(99)
    Call writer.GetData(data): Call incoming.SetData(data)
    Mapping(1).PacketCount = 0
    Call Protocol.HandleIncomingData(1, incoming, u)
End Sub
Private Sub StartDuo(Optional ByVal leader As Integer = 1)
    Call Propose(leader)
    Dim proposal As Long: proposal = PartyQuestProposalForTest(leader)
    Call Expect(proposal > 0, "normal quest accept proposes instead of enrolling")
    Call Expect(Not TieneQuest(leader, 1), "proposal is not acceptance")
    Call Action(leader + 1, 1, HuntAccept, proposal)
    Call Action(leader, 1, HuntStart, proposal)
    Call Expect(TieneQuest(leader, 1) And TieneQuest(leader + 1, 1), "accepted roster receives ordinary quest projections")
End Sub
Private Sub KillWolf(Optional ByVal credited As Integer = 1)
    Call AwardNpcPartyQuests(credited, victim)
End Sub
Private Sub RestartCache()
    Call ResetPartyQuestTests()
    Call database.Close()
    Call database.Open("Driver={SQLite3 ODBC Driver};Database=" & fileName & ";")
    Dim i As Integer
    For i = 1 To 7
        If UserList(i).flags.UserLogged Then Call Expect(LoadPartyQuest(i), "restart reloads durable projection")
    Next i
End Sub
Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Fixture()
    Dim i As Integer, proposal As Long, slot As Byte, rows As ADODB.Recordset
    Select Case index
    Case 1
        Call StartDuo()
        With UserList(1).QuestStats.Quests(2)
            .QuestIndex = 2: ReDim .NPCsKilled(1 To 1)
        End With
        For i = 1 To 3: Call KillWolf(): Next i
        PartyQuestSqlCount = 0: Call KillWolf()
        Call Expect(PartyQuestProgressForTest(1) = 4 And PartyQuestProgressForTest(2) = 4, "fighter and zero-damage healer share4/6")
        Call Expect(PartyQuestMaskForTest(1, 4) = 3, "qualified roster mask recorded in memory")
        Call Expect(PartyQuestSqlCount = 0, "kill and personal progress perform zero SQL")
        Call Expect(LastSharedCount(1) = 4 And LastSharedCount(2) = 4, "both clients receive immediate authoritative counter4")
        Call Expect(UserList(1).QuestStats.Quests(2).NPCsKilled(1) = 4, "personal wolf quest still progresses independently")
    Case 2
        Call Group(1, 3): Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntAccept, proposal): Call Action(3, 1, HuntAccept, proposal): Call Action(1, 1, HuntStart, proposal)
        Call KillWolf()
        Call Expect(PartyQuestProgressForTest(3) = 1 And PartyQuestMaskForTest(1, 1) = 7, "three witnesses increment exactly once")
        Call SaveCharacterDB(1)
        Call Expect(Scalar("SELECT count(*) FROM party_quest_kill") = 1, "normal save persists one kill")
    Case 3
        Call StartDuo()
        For i = 1 To 4: Call KillWolf(): Next i
        Call SaveCharacterDB(2, True)
        Call PartyQuestLogout(2): UserList(2).flags.UserLogged = False
        Call KillWolf(): Call Expect(PartyQuestProgressForTest(1) = 4, "solo cannot advance after healer disconnect")
        UserList(2).VersionId = UserList(2).VersionId + 1: UserList(2).flags.UserLogged = True
        Call Expect(LoadPartyQuest(2), "reconnect loads same persistent run")
        Call Group(1, 2): UserList(1).Grupo.Id = 999: UserList(2).Grupo.Id = 999
        Call KillWolf()
        Call Expect(PartyQuestProgressForTest(1) = 5 And PartyQuestProgressForTest(2) = 5, "regrouped new party resumes 5/6")
    Case 4
        Call Group(3, 2): Call StartDuo(1): Call StartDuo(3)
        Call KillWolf(1): Call KillWolf(3): Call KillWolf(3)
        Call Expect(PartyQuestProgressForTest(1) = 1 And PartyQuestProgressForTest(3) = 2, "two parties have independent counters")
    Case 5
        Call StartDuo(): Call Group(1, 3): Call KillWolf()
        Call Expect(Not TieneQuest(3, 1) And PartyQuestProgressForTest(3) = 0, "late joiner cannot inherit enrollment")
        Call ClaimPartyQuest(3, 1): Call Expect(UserList(3).Stats.GLD = 10, "late joiner has no claim")
    Case 6, 7
        Call StartDuo()
        For i = 1 To 6: Call KillWolf(): Next i
        Call Expect(PartyQuestReady(1, 1) And PartyQuestReady(2, 1), "both participating members ready")
        Call ClaimPartyQuest(1, 1): Call ClaimPartyQuest(1, 1)
        Call Expect(UserList(1).Stats.GLD = 370 And UserList(2).Stats.GLD = 10, "independent exactly-once first claim")
        If index = 7 Then
            Call SaveCharacterDB(1, True)
            Call RestartCache()
            Call Expect(Scalar("SELECT count(*) FROM quest_done WHERE user_id=701") = 1, "durable claim repairs stale async completion projection")
            Call Expect(Scalar("SELECT gold FROM user WHERE id=701") = 370, "claimed gold survives actual database reopen")
        End If
        Call SetNpcRef(UserList(2).flags.TargetNPC, 1)
        Call ClaimPartyQuest(2, 1): Call ClaimPartyQuest(2, 1)
        Call SaveCharacterDB(2, True)
        Call Expect(UserList(2).Stats.GLD = 370 And Scalar("SELECT count(*) FROM party_quest_member WHERE status=1") = 2, "second claim survives partial-claim restart")
    Case 8
        Call StartDuo(): Call KillWolf(): Call SaveCharacterDB(1, True): Call RestartCache()
        Call Expect(PartyQuestProgressForTest(1) = 1 And PartyQuestProgressForTest(2) = 1, "active counter survives actual database close/reopen")
    Case 9
        Call StartDuo()
        With UserList(1).QuestStats.Quests(2)
            .QuestIndex = 2: ReDim .NPCsKilled(1 To 1)
        End With
        Call KillWolf()
        Call Expect(UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 1 And UserList(1).QuestStats.Quests(2).NPCsKilled(1) = 1, "shared and personal progress independently")
    Case 10
        Call Group(1, 3): Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntAccept, proposal): Call Action(1, 1, HuntStart, proposal)
        Call Expect(Not TieneQuest(3, 1), "unanswered member never silently enrolls")
    Case 11
        Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntDecline, proposal): Call Action(1, 1, HuntStart, proposal)
        Call Expect(Not TieneQuest(1, 1) And Scalar("SELECT count(*) FROM party_quest_run") = 0, "decline blocks duo launch")
    Case 12
        Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntAccept, proposal): UserList(2).pos.x = 70
        Call Action(1, 1, HuntStart, proposal)
        Call Expect(Not TieneQuest(1, 1), "launch revalidates distance")
    Case 13
        Call StartDuo(): Call AbandonPartyQuest(1): Call KillWolf(2)
        Call Expect(Not TieneQuest(1, 1) And PartyQuestProgressForTest(2) = 0, "abandon removes eligibility without blocking future enrollment")
        Call SaveCharacterDB(1, True)
        Call RestartCache(): Call Expect(Not TieneQuest(1, 1), "abandon persists")
    Case 14 To 18
        Call StartDuo()
        Select Case index
            Case 14: UserList(2).flags.Muerto = 1
            Case 15: UserList(2).ConnectionDetails.ConnIDValida = False
            Case 16: UserList(2).pos.Map = 2
            Case 17: UserList(2).pos.x = 70
            Case 18: UserList(2).Grupo.Id = 99
        End Select
        Call KillWolf(): Call Expect(PartyQuestProgressForTest(1) = 0, "ineligible second member cannot count")
    Case 19
        Call StartDuo(): UserList(2).pos.x = 69: Call KillWolf()
        Call Expect(PartyQuestProgressForTest(1) = 1, "distance19 qualifies")
    Case 20
        Call StartDuo(): victim.npcType = DummyTarget: Call KillWolf()
        Call Expect(PartyQuestProgressForTest(1) = 0, "training victim excluded")
        victim.npcType = 0: Call SetUserRef(victim.MaestroUser, 3): Call KillWolf()
        Call Expect(PartyQuestProgressForTest(1) = 0, "owned victim excluded")
    Case 21
        Call StartDuo(): victim.Numero = 499: Call KillWolf()
        Call Expect(PartyQuestProgressForTest(1) = 0, "exact template matching")
    Case 22
        Call StartDuo()
        For i = 1 To 7: Call KillWolf(): Next i
        Call Expect(PartyQuestProgressForTest(1) = 6, "counter capped at requirement in memory")
        Call SaveCharacterDB(1)
        Call Expect(Scalar("SELECT count(*) FROM party_quest_kill") = 6, "save persists exactly six counted kills")
    Case 23
        Call StartDuo(): Call ClaimPartyQuest(1, 1)
        Call Expect(UserList(1).Stats.GLD = 10 And Scalar("SELECT count(*) FROM party_quest_member WHERE status=1") = 0, "incomplete claim rejected")
    Case 24
        Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntAccept, proposal): Call Action(1, 1, HuntStart, proposal): Call Action(1, 1, HuntStart, proposal)
        Call SaveCharacterDB(1)
        Call Expect(Scalar("SELECT count(*) FROM party_quest_run") = 1, "duplicate start does not allocate run")
    Case 25
        Call StartDuo()
        HoldPartySaves = True: Call SaveCharacterDB(1)
        Call HoldPartySnapshotForTest(1)
        For i = 1 To 6: Call KillWolf(): Next i
        Call ClaimPartyQuest(1, 1)
        UserList(1).Stats.GLD = UserList(1).Stats.GLD + 40 - 15
        Call ClaimPartyQuest(1, 1)
        HoldPartySaves = False: Call SaveCharacterDB(1)
        Call Expect(CompleteHeldPartySnapshotForTest(), "older production snapshot completes harmlessly")
        Call Expect(Scalar("SELECT gold FROM user WHERE id=701") = 395 And UserList(1).Stats.GLD = 395, "stale save cannot erase claim or unrelated gold")
        Call RestartCache()
        Call Expect(Scalar("SELECT count(*) FROM party_quest_member WHERE character_id=701 AND status=1") = 1, "one persisted claim")
    Case 26
        Call StartDuo()
        For i = 1 To 6: Call KillWolf(): Next i
        Call ClaimPartyQuest(1, 1)
        ObjData(50).name = "Fixture item"
        UserList(1).invent.Object(1).ObjIndex = 50: UserList(1).invent.Object(1).amount = 7
        UserList(1).Stats.Exp = 123
        Call database.Execute("CREATE TRIGGER deny_save BEFORE UPDATE ON user WHEN OLD.id=701 BEGIN SELECT RAISE(ABORT,'fixture'); END")
        Call SaveCharacterDB(1, True)
        Call Expect(UserList(1).Stats.GLD = 370 And HasPendingPartySaveForTest(701), "failed save retains claim and whole character snapshot")
        Call Expect(Scalar("SELECT gold FROM user WHERE id=701") = 10 And Scalar("SELECT count(*) FROM party_quest_run") = 0, "claim, wallet and run roll back atomically")
        Call ClaimPartyQuest(1, 1): UserList(1).Stats.GLD = UserList(1).Stats.GLD - 13 + 9
        Call database.Execute("DROP TRIGGER deny_save")
        Call RetryPartyQuestSaves()
        Call Expect(Scalar("SELECT gold FROM user WHERE id=701") = 366, "retry captures latest unrelated spending and earnings")
        Call Expect(Scalar("SELECT exp FROM user WHERE id=701") = 123 And Scalar("SELECT amount FROM inventory_item WHERE user_id=701 AND number=1") = 7, "inventory and unrelated character changes recover")
    Case 27
        Call Group(1, 3): Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntAccept, proposal): Call Action(3, 1, HuntAccept, proposal): Call Action(1, 1, HuntStart, proposal)
        UserList(3).pos.x = 90
        For i = 1 To 6: Call KillWolf(): Next i
        Call Expect(Not PartyQuestReady(3, 1) And PartyQuestProgressForTest(3) = 6, "nonparticipant shares counter but cannot claim")
    Case 28
        Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntAccept, proposal + 1): Call Action(1, 1, HuntStart, proposal)
        Call Expect(Not TieneQuest(1, 1), "stale invitation cannot enroll")
    Case 29
        Call StartDuo(): Call PartyQuestLogout(2)
        UserList(2).Id = 999: UserList(2).VersionId = UserList(2).VersionId + 1
        Call Expect(LoadPartyQuest(2), "new character loads")
        Call Expect(Not TieneQuest(2, 1), "connection slot reuse cannot inherit fixed character roster")
    Case 30
        Call SetNpcRef(UserList(1).flags.TargetNPC, 1)
        Call Expect(QuestAcceptanceFailure(1, 1, 2, QuestList(2)) = QuestAcceptAllowed, "Alaric personal offer unaffected by party requirement")
    Case 34
        Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntAccept, proposal, True): Call Action(1, 1, HuntStart, proposal)
        Call Expect(Not TieneQuest(1, 1), "trailing action bytes do not enroll a player")
    Case 33
        For i = 1 To 20
            Call Propose(1): proposal = PartyQuestProposalForTest(1)
            Call Action(2, i * 2, HuntAccept, proposal): Call Action(1, i * 2, HuntStart, proposal)
            Call Expect(TieneQuest(1, 1) And TieneQuest(2, 1), "cache reuse permits a fresh hunt after abandon")
            Call AbandonPartyQuest(1): Call AbandonPartyQuest(2)
            Call SaveCharacterDB(1): Call SaveCharacterDB(2)
        Next i
        Call Expect(Scalar("SELECT count(*) FROM party_quest_run") = 20, "bounded cache does not discard durable history")
    Case 31, 32
        Call StartDuo()
        Dim npcIndex As Integer, petIndex As Integer, result As e_DamageResult
        npcIndex = GetNextAvailableNpc(): petIndex = GetNextAvailableNpc()
        NpcList(npcIndex) = victim
        With NpcList(npcIndex)
            .flags.NPCActive = True: .VersionId = fixtureVersion
            .Char.charindex = npcIndex: .Stats.MinHp = 1: .Stats.MaxHp = 1
            .pos.x = 51: .flags.ExpCount = 0
            ReDim .Char.Animation(0 To 0)
        End With
        CharList(npcIndex) = npcIndex: MapData(51, 50, 1).NpcIndex = npcIndex
        NpcList(petIndex).flags.NPCActive = True
        If index = 31 Then Call SetUserRef(NpcList(petIndex).MaestroUser, 1)
        result = NPCs.DoDamageOrHeal(npcIndex, petIndex, eNpc, -1, e_pet, 0)
        If index = 31 Then
            Call Expect(PartyQuestProgressForTest(1) = 1 And PartyQuestProgressForTest(2) = 1, "pet owner attributes one shared kill through real death path")
        Else
            Call Expect(PartyQuestProgressForTest(1) = 0, "unowned NPC kill does not credit party hunt")
        End If
    Case 35
        Call StartDuo()
        For i = 1 To 4: Call KillWolf(): Next i
        HoldPartySaves = True: Call SaveCharacterDB(2, True): Call HoldPartySnapshotForTest(2)
        Call PartyQuestBeforeSlotReset(2): UserList(2).flags.UserLogged = False
        Call KillWolf(): Call Expect(PartyQuestProgressForTest(1) = 4, "solo kill after logout cannot advance")
        HoldPartySaves = False: Call ReloadCharacter(2): Call Group(1, 2)
        UserList(1).Grupo.Id = 999: UserList(2).Grupo.Id = 999
        Call KillWolf(): Call SaveChangesInUser(1)
        Call Expect(CompleteHeldPartySnapshotForTest(), "old4/6 save completes after newer5/6")
        Call Expect(Scalar("SELECT progress FROM party_quest_run") = 5, "old run write cannot overwrite5/6")
        Call RestartCache(): Call ReloadCharacter(1): Call ReloadCharacter(2)
        Call Expect(PartyQuestProgressForTest(1) = 5 And PartyQuestProgressForTest(2) = 5, "production reload restores shared5/6")
    Case 36
        Call StartDuo()
        For i = 1 To 4: Call KillWolf(): Next i
        Call SaveCharacterDB(1, True): Call PartyQuestBeforeSlotReset(1): UserList(1).flags.UserLogged = False
        Call SaveCharacterDB(2, True): Call PartyQuestBeforeSlotReset(2): UserList(2).flags.UserLogged = False
        Call RestartCache(): Call ReloadCharacter(1): Call ReloadCharacter(2)
        Call Expect(PartyQuestProgressForTest(1) = 4 And PartyQuestProgressForTest(2) = 4, "all offline, database reopen, full character reload preserves run")
    Case 37
        Call StartDuo()
        For i = 1 To 4: Call KillWolf(): Next i
        Call database.Execute("CREATE TRIGGER deny_run BEFORE INSERT ON party_quest_run BEGIN SELECT RAISE(ABORT,'fixture'); END")
        Call SaveCharacterDB(2, True): Call PartyQuestBeforeSlotReset(2): UserList(2).flags.UserLogged = False
        Call SaveCharacterDB(1, True): Call PartyQuestBeforeSlotReset(1): UserList(1).flags.UserLogged = False
        Call Expect(HasPendingPartySaveForTest(701) And HasPendingPartySaveForTest(702), "last member failure retains owned snapshots")
        Dim replacement As t_User
        UserList(1) = replacement
        UserList(1).Id = 999: UserList(1).VersionId = fixtureVersion + 10: UserList(1).Stats.GLD = 77
        UserList(1).flags.UserLogged = True: UserList(1).ConnectionDetails.ConnIDValida = True
        Call database.Execute("INSERT INTO user(id,gold) VALUES(999,77)")
        Call database.Execute("DROP TRIGGER deny_run")
        Call RetryPartyQuestSaves(7)
        Call Expect(Scalar("SELECT progress FROM party_quest_run") = 4, "offline retry persists correct run")
        Call Expect(UserList(1).Stats.GLD = 77 And Scalar("SELECT gold FROM user WHERE id=999") = 77, "reused connection slot untouched")
        UserList(3).Id = 701: UserList(3).AccountID = 701
        Call ReloadCharacter(3): Call Expect(PartyQuestProgressForTest(3) = 4, "original character resumes in another slot")
    Case 38
        Call StartDuo(): Call KillWolf(): Call SaveCharacterDB(1)
        Call ResetPartyQuestTests()
        FailPartyRosterRead = True
        Call Expect(Not LoadPartyQuest(1), "first roster read fails")
        Call Expect(PartyQuestCacheCountForTest() = 0, "partial run was never published")
        Call Expect(LoadPartyQuest(1) And PartyQuestProgressForTest(1) = 1, "same-process retry restores full run")
        Call Expect(PartyQuestMemberCountForTest(1) = 2, "complete roster survives failed read")
    Case 39, 40
        Call StartDuo()
        For i = 1 To 6: Call KillWolf(): Next i
        If index = 39 Then Call ClaimPartyQuest(1, 1) Else Call AbandonPartyQuest(1)
        Call SaveCharacterDB(1, True): Call PartyQuestBeforeSlotReset(1): UserList(1).flags.UserLogged = False
        Call database.Execute("DELETE FROM user WHERE id=701")
        Call Expect(Scalar("SELECT count(*) FROM user WHERE id=701") = 0, "character deletion works with foreign keys enforced")
        Call Expect(Scalar("SELECT count(*) FROM party_quest_member WHERE character_id=701 AND ordinal=1 AND user_id IS NULL") = 1, "tombstone preserves ordinal1")
        Call Expect(Scalar("SELECT qualified_mask FROM party_quest_kill WHERE ordinal=1") = 3, "participation mask is unchanged")
        Call RestartCache(): Call ReloadCharacter(2)
        Call Expect(PartyQuestReady(2, 1) And PartyQuestMemberCountForTest(2) = 2, "teammate retains progress and claim rights")
        Call SetNpcRef(UserList(2).flags.TargetNPC, 1): Call ClaimPartyQuest(2, 1): Call SaveCharacterDB(2)
        Call Expect(UserList(2).Stats.GLD = 370, "teammate claims independently after deletion")
    Case 41
        Call StartDuo()
        For i = 1 To 4: Call KillWolf(): Next i
        HoldPartySaves = True: Call SaveCharacterDB(2, True)
        PartyQuestSqlCount = 0: Call KillWolf()
        Call Expect(PartyQuestSqlCount = 0 And PartyQuestProgressForTest(1) = 5, "queued save never sends kill SQL")
        Call PartyQuestBeforeSlotReset(2): UserList(2).flags.UserLogged = False
        HoldPartySaves = False: Call RetryPartyQuestSaves()
        Call Expect(Scalar("SELECT progress FROM party_quest_run") = 5, "pending save coalesces to latest run at retry boundary")
    Case 42
        Call StartDuo(): Call KillWolf()
        HoldPartySaves = True: Call SaveCharacterDB(1, True)
        Call Expect(Not FlushPartyQuestShutdown(), "shutdown cannot report success while save pending")
        HoldPartySaves = False
        Call Expect(FlushPartyQuestShutdown(), "orderly shutdown flushes owned work")
        Call Expect(Scalar("SELECT progress FROM party_quest_run") = 1, "shutdown persisted latest dirty run")
    Case 43
        Dim fence As New clsHuntAsyncFence, probe As New ADODB.Command, deadline As Long
        Call fence.Attach(database)
        Set probe.ActiveConnection = database: probe.CommandText = "SELECT 1": probe.CommandType = adCmdText
        Call probe.Execute(, , adAsyncExecute)
        deadline = GetTickCountRaw()
        Do While fence.Pending <> 0 And TicksElapsed(deadline, GetTickCountRaw()) < 1000
            DoEvents
        Loop
        Call Expect(fence.Started > 0 And fence.Completed > 0 And fence.Pending = 0, "real ADO command completion fence fires")
    Case 44
        Call StartDuo()
        For i = 1 To 4: Call KillWolf(): Next i
        HoldPartySaves = True: Call SaveCharacterDB(2): Call HoldPartySnapshotForTest(2)
        Call KillWolf(): HoldPartySaves = False
        Call Expect(CompleteHeldPartySnapshotForTest(), "old snapshot can commit before latest save")
        Call Expect(Scalar("SELECT progress FROM party_quest_run") = 4 And PartyQuestDirtyForTest(1), "old commit cannot clear newer dirty revision")
        Call SaveChangesInUser(1)
        Call Expect(Scalar("SELECT progress FROM party_quest_run") = 5 And Not PartyQuestDirtyForTest(1), "latest normal save acknowledges5/6")
    Case 45
        Call StartDuo()
        For i = 1 To 6: Call KillWolf(): Next i
        Call ClaimPartyQuest(1, 1)
        UserList(1).Stats.GLD = UserList(1).Stats.GLD + 25
        UserList(2).Stats.GLD = UserList(2).Stats.GLD + 7
        Call database.Execute("CREATE TRIGGER deny_claim_wallet BEFORE UPDATE ON user WHEN NEW.id=701 BEGIN SELECT RAISE(ABORT,'fixture'); END")
        Call SaveChangesInUser(2)
        Call Expect(Scalar("SELECT count(*) FROM party_quest_run") = 0 And Scalar("SELECT gold FROM user WHERE id=702") = 10, "teammate save rolls back with claimant wallet")
        Call Expect(UserList(1).Stats.GLD = 395 And Not PartyQuestReady(1, 1), "failed save never reopens claim or replaces live wallet")
        Call database.Execute("DROP TRIGGER deny_claim_wallet")
        Call RetryPartyQuestSaves(7)
        Call Expect(Scalar("SELECT count(*) FROM party_quest_member WHERE status=1") = 1 And Scalar("SELECT gold FROM user WHERE id=701") = 395, "teammate save commits claimant and wallet together")
        Call Expect(Scalar("SELECT gold FROM user WHERE id=702") = 17, "teammate unrelated wallet is preserved")
        Call RestartCache(): Call ReloadCharacter(1)
        Call Expect(UserList(1).Stats.GLD = 395 And Not PartyQuestReady(1, 1), "claim remains once after full reload")
    End Select
    Call CleanupFixture()
    Call UnitTesting.Progress("Party hunt " & index & " passed")
    Call UnitTesting.RunTest("Party hunt " & index, True)
    Exit Sub
Failed:
    Dim failure As String: failure = Err.Description
    Resume ReportFailure
ReportFailure:
    Call CleanupFixture()
    Call UnitTesting.Progress("Party hunt " & index & " failed: " & failure)
    Call UnitTesting.RunTestError("Party hunt " & index, failure)
End Sub
Private Sub CleanupFixture()
    ' A pooled ODBC file handle must not mask the actual assertion failure.
    On Error Resume Next
    capturing = False
    Unit_PartyRewards.SynchronousPersistence = False
    HoldPartySaves = False
    If Not database Is Nothing Then
        If database.State <> adStateClosed Then Call database.Close()
    End If
    Set Connection = oldConnection: Set database = Nothing
    If Len(fileName) > 0 Then Kill fileName
    Err.Clear
End Sub
Public Function TestSuite() As Boolean
    Dim i As Integer, before As Long, started As Long
    Call Unit_NpcBow.InitializeTestWorld()
    started = GetTickCountRaw()
    before = UnitTesting.GetFailedTests()
    #If PYMMO = 1 Then
        Call UnitTesting.Progress("Party hunt native validator configuration PYMMO=1")
    #Else
        Call UnitTesting.Progress("Party hunt native validator configuration PYMMO=0")
    #End If
    For i = 1 To 45: Call RunCase(i): Next i
    Call UnitTesting.Progress("Party hunt 45 cases duration ms=" & TicksElapsed(started, GetTickCountRaw()))
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
