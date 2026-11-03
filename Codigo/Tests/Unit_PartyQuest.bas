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
    If Not value Then Err.Raise vbObjectError + 861, "Party quest", label
End Sub
Private Function Scalar(ByVal sql As String) As Long
    Dim rows As ADODB.Recordset
    Set rows = database.Execute(sql)
    Scalar = rows.Fields(0).value
    Call rows.Close()
End Function
Private Sub Fixture()
    Call InitializeNpcIndexHeap()
    MapData(51, 50, 1).NpcIndex = 0
    Call ResetPartyQuestTests(): Call ResetTurnInTestCounters()
    Set oldConnection = Connection
    fileName = Environ$("TEMP") & "\ao-party-hunt-test-" & GetTickCountRaw() & ".db"
    Set database = New ADODB.Connection
    Call database.Open("Driver={SQLite3 ODBC Driver};Database=" & fileName & ";")
    Set Connection = database
    Call database.Execute("CREATE TABLE user(id INTEGER PRIMARY KEY,gold INTEGER NOT NULL DEFAULT 0)")
    Call database.Execute("CREATE TABLE quest(user_id INTEGER,number INTEGER,quest_id INTEGER,npcs TEXT,npcstarget TEXT,PRIMARY KEY(user_id,number))")
    Call database.Execute("CREATE TABLE quest_done(user_id INTEGER,quest_id INTEGER,PRIMARY KEY(user_id,quest_id))")
    ' Exercise the production migration, not a parallel test schema.
    Dim f As Integer, sql As String, statement As Variant
    f = FreeFile: Open App.Path & "\ScriptsDB\20261103-01-party-quest.sql" For Binary Access Read As #f
    sql = Space$(LOF(f)): Get #f, , sql: Close #f
    For Each statement In Split(sql, ";")
        If Len(Trim$(statement)) > 0 Then Call database.Execute(CStr(statement))
    Next statement
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
        Call database.Execute("INSERT INTO user(id,gold) VALUES(" & UserList(i).Id & ",10)")
    Next i
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
        Call StartDuo(): Call KillWolf()
        Call Expect(PartyQuestProgressForTest(1) = 1 And PartyQuestProgressForTest(2) = 1, "fighter and zero-damage healer share 1/6")
        Call Expect(Scalar("SELECT qualified_mask FROM party_quest_kill") = 3, "qualified persistent roster recorded")
    Case 2
        Call Group(1, 3): Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntAccept, proposal): Call Action(3, 1, HuntAccept, proposal): Call Action(1, 1, HuntStart, proposal)
        Call KillWolf()
        Call Expect(PartyQuestProgressForTest(3) = 1 And Scalar("SELECT count(*) FROM party_quest_kill") = 1, "three witnesses increment exactly once")
    Case 3
        Call StartDuo()
        For i = 1 To 4: Call KillWolf(): Next i
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
            Call database.Execute("DELETE FROM quest_done WHERE user_id=701")
            Call RestartCache()
            Call Expect(Scalar("SELECT count(*) FROM quest_done WHERE user_id=701") = 1, "durable claim repairs stale async completion projection")
            Call Expect(Scalar("SELECT gold FROM user WHERE id=701") = 370, "claimed gold survives actual database reopen")
        End If
        Call SetNpcRef(UserList(2).flags.TargetNPC, 1)
        Call ClaimPartyQuest(2, 1): Call ClaimPartyQuest(2, 1)
        Call Expect(UserList(2).Stats.GLD = 370 And Scalar("SELECT count(*) FROM party_quest_member WHERE status=1") = 2, "second claim survives partial-claim restart")
    Case 8
        Call StartDuo(): Call KillWolf(): Call RestartCache()
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
        Call Expect(PartyQuestProgressForTest(1) = 6 And Scalar("SELECT count(*) FROM party_quest_kill") = 6, "counter capped at requirement")
    Case 23
        Call StartDuo(): Call ClaimPartyQuest(1, 1)
        Call Expect(UserList(1).Stats.GLD = 10 And Scalar("SELECT count(*) FROM party_quest_member WHERE status=1") = 0, "incomplete claim rejected")
    Case 24
        Call Propose(1): proposal = PartyQuestProposalForTest(1)
        Call Action(2, 1, HuntAccept, proposal): Call Action(1, 1, HuntStart, proposal): Call Action(1, 1, HuntStart, proposal)
        Call Expect(Scalar("SELECT count(*) FROM party_quest_run") = 1, "duplicate start does not allocate run")
    Case 25
        Call StartDuo()
        For i = 1 To 6: Call KillWolf(): Next i
        Dim savedSnapshot As t_User, assignments As String, assignment As Variant, column As String
        savedSnapshot = UserList(1)
        Call ClaimPartyQuest(1, 1)
        Call Contruir_Querys()
        assignments = Mid$(QUERY_UPDATE_MAINPJ, InStr(QUERY_UPDATE_MAINPJ, "SET ") + 4)
        assignments = Left$(assignments, InStr(assignments, "WHERE ") - 1)
        For Each assignment In Split(assignments, ",")
            column = Trim$(Left$(CStr(assignment), InStr(CStr(assignment), "=") - 1))
            If column <> "gold" Then Call database.Execute("ALTER TABLE user ADD COLUMN [" & column & "] TEXT DEFAULT 0")
        Next assignment
        Unit_PartyRewards.SynchronousPersistence = True
        ' Execute the actual production save builder/parameters using a pre-claim snapshot.
        Call SavePartyQuestGoldSnapshotForTest(savedSnapshot)
        Call Expect(Scalar("SELECT gold FROM user WHERE id=701") = 370, "production pre-claim save cannot overwrite reward")
        UserList(1).Stats.GLD = 365
        Call SavePartyQuestGoldSnapshotForTest(UserList(1))
        Call Expect(Scalar("SELECT gold FROM user WHERE id=701") = 365, "production post-claim save does not duplicate reward")
        Unit_PartyRewards.SynchronousPersistence = False
    Case 26
        Call StartDuo()
        For i = 1 To 6: Call KillWolf(): Next i
        Call database.Execute("CREATE TRIGGER deny_claim BEFORE UPDATE OF status ON party_quest_member WHEN NEW.status=1 BEGIN SELECT RAISE(ABORT,'fixture'); END")
        Call ClaimPartyQuest(1, 1)
        Call Expect(UserList(1).Stats.GLD = 10 And PartyQuestReady(1, 1), "failed transaction leaves retryable unclaimed state")
        Call database.Execute("DROP TRIGGER deny_claim")
        Call ClaimPartyQuest(1, 1): Call Expect(UserList(1).Stats.GLD = 370, "retry grants exactly once")
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
    For i = 1 To 34: Call RunCase(i): Next i
    Call UnitTesting.Progress("Party hunt 34 cases duration ms=" & TicksElapsed(started, GetTickCountRaw()))
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
