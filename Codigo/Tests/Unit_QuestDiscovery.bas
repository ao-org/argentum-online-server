Attribute VB_Name = "Unit_QuestDiscovery"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private attempts As Long
Private rejectAt As Long
Public Function CapturePacket(ByRef queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    Dim bytes() As Byte, writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    queued = True
    If UBound(bytes) >= 1 Then
        If CLng(bytes(0)) + 256& * bytes(1) = ServerPacketID.eQuestDiscovery Then
            attempts = attempts + 1
            queued = attempts <> rejectAt
            If queued Then
                Dim payload As Variant: payload = bytes
                Call packets.Add(payload)
            End If
        End If
    End If
    CapturePacket = True
End Function
Private Sub Expect(ByVal value As Boolean, ByVal label As String)
    If Not value Then Err.Raise vbObjectError + 849, "Quest discovery", label
End Sub
Private Sub Fixture()
    Call Unit_NpcBow.InitializeTestWorld()
    Call ResetTurnInTestCounters()
    ReDim QuestList(0 To 300)
    ReDim NpcInfoCache(1 To 20)
    ReDim GlobalQuestInfo(1 To 1)
    With UserList(1)
        .Id = 140: .VersionId = 1: .flags.UserLogged = True
        .Stats.ELV = 10: .clase = e_Class.Mage: .Faccion.Status = e_Facciones.Ciudadano
        .QuestStats.NumQuestsDone = 0
        Dim slot As Integer
        For slot = 1 To MAXUSERQUESTS
            .QuestStats.Quests(slot).QuestIndex = 0
        Next slot
    End With
    QuestList(1).nombre = "Eligible quest"
    With NpcInfoCache(10)
        .Exists = True: .NumQuest = 1
        ReDim .QuestNumber(1 To 1): .QuestNumber(1) = 1
    End With
    Call ResetQuestDiscoveryIndex()
    DiscoveryEvaluationCount = 0
    Set packets = New Collection
    attempts = 0: rejectAt = 0: capturing = True
End Sub
Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Fixture()
    Dim bytes() As Byte, reader As New Network.Reader, firstSequence As Long, q As Integer, count As Integer
    Dim checks As Long, before As Long, flags As Byte, page As Integer
    Select Case index
    Case 1
        Call Expect(IsQuestDiscoverable(1, 1), "valid offer")
        Call Expect(CanUserAcceptQuest(1, 0, 1, QuestList(1)), "acceptance shares predicate")
        Call Expect(PublishQuestDiscovery(1, True), "initial snapshot queued")
        bytes = packets(1): Call reader.SetData(bytes)
        Call Expect(reader.ReadInt16() = ServerPacketID.eQuestDiscovery, "packet id")
        firstSequence = reader.ReadInt32()
        Call Expect(firstSequence > 0 And reader.ReadInt16() = 0, "transaction identity")
        Call Expect(reader.ReadInt8() = 1 And reader.ReadInt16() = 1, "complete one-offer replacement")
        Call Expect(reader.ReadInt16() = 10 And reader.ReadInt16() = 1, "template and quest identities")
        Call Expect(reader.ReadString8() = "Eligible quest" And reader.GetAvailable() = 0, "exact title and packet length")
    Case 2
        QuestList(1).Trabajador = True
        Call Expect(Not IsQuestDiscoverable(1, 1), "worker restriction")
        UserList(1).clase = e_Class.Trabajador
        Call Expect(IsQuestDiscoverable(1, 1), "worker permitted")
        QuestList(1).Trabajador = False
        QuestList(1).RequiredClassesCount = 1
        ReDim QuestList(1).RequiredClass(1 To 1)
        QuestList(1).RequiredClass(1) = e_Class.Mage
        Call Expect(Not IsQuestDiscoverable(1, 1), "class restriction")
        UserList(1).clase = e_Class.Mage
        Call Expect(IsQuestDiscoverable(1, 1), "required class permitted")
    Case 3
        QuestList(1).RequiredLevel = 11
        Call Expect(Not IsQuestDiscoverable(1, 1), "minimum level")
        QuestList(1).RequiredLevel = 0: QuestList(1).LimitLevel = 9
        Call Expect(Not IsQuestDiscoverable(1, 1), "maximum level")
        QuestList(1).LimitLevel = 0: QuestList(1).RequiredSkill.SkillType = 1
        QuestList(1).RequiredSkill.RequiredValue = 20: UserList(1).Stats.UserSkills(1) = 0
        Call Expect(Not IsQuestDiscoverable(1, 1), "required skill")
        UserList(1).Stats.UserSkills(1) = 20
        Call Expect(IsQuestDiscoverable(1, 1), "skill attained")
    Case 4
        QuestList(1).RequiredQuest = 2
        Call Expect(Not IsQuestDiscoverable(1, 1), "prerequisite hidden")
        Call AddDoneQuest(1, 2)
        Call Expect(IsQuestDiscoverable(1, 1), "prerequisite complete")
        Call AddDoneQuest(1, 1)
        Call Expect(Not IsQuestDiscoverable(1, 1), "non-repeatable complete")
        QuestList(1).Repetible = 1
        Call Expect(IsQuestDiscoverable(1, 1), "repeatable reacceptance")
        UserList(1).QuestStats.Quests(1).QuestIndex = 1
        Call Expect(Not IsQuestDiscoverable(1, 1), "active repeatable excluded")
    Case 5
        QuestList(1).PermittedFactions = e_QuestPermittedFactions.RoyalArmy
        Call Expect(Not IsQuestDiscoverable(1, 1), "faction restriction")
        UserList(1).Faccion.Status = e_Facciones.Armada
        Call Expect(IsQuestDiscoverable(1, 1), "faction permitted")
        For q = 1 To MAXUSERQUESTS
            UserList(1).QuestStats.Quests(q).QuestIndex = 2
        Next q
        Call Expect(Not IsQuestDiscoverable(1, 1), "no free quest slot")
        UserList(1).QuestStats.Quests(1).QuestIndex = 0
        Call Expect(IsQuestDiscoverable(1, 1), "free slot restored")
    Case 6
        QuestList(1).GlobalQuestIndex = 1: QuestList(1).GlobalQuestThresholdNeeded = 10
        Call ResetQuestDiscoveryIndex()
        Call Expect(Not IsQuestDiscoverable(1, 1), "global threshold")
        Call Expect(PublishQuestDiscovery(1), "empty initial snapshot")
        before = DiscoveryEvaluationCount
        GlobalQuestInfo(1).GatheringGlobalCounter = 1
        Call RefreshGlobalQuestTurnIns(1): Call ProcessQuestTurnInUpdates()
        Call Expect(DiscoveryEvaluationCount = before And packets.Count = 1, "below-threshold contributions do not rescan offers")
        GlobalQuestInfo(1).GatheringGlobalCounter = 10
        Call RefreshGlobalQuestTurnIns(1): Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 2, "unaccepted global offer refreshed")
        Call RefreshGlobalQuestTurnIns(1): Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 2, "unchanged global signature suppressed")
    Case 7
        Call SendQuestTurnInSnapshot(1): Call SendQuestTurnInSnapshot(1): Call ProcessQuestTurnInUpdates()
        Call Expect(DiscoveryEvaluationCount = 1 And packets.Count = 1, "dirty mutations coalesced")
        before = DiscoveryEvaluationCount
        Call SendQuestTurnInSnapshot(1): Call ProcessQuestTurnInUpdates()
        Call Expect(DiscoveryEvaluationCount = before And packets.Count = 1, "unrelated mutation skips eligibility and packet")
        UserList(1).Stats.ELV = 11
        Call SendQuestTurnInSnapshot(1): Call ProcessQuestTurnInUpdates()
        Call Expect(DiscoveryEvaluationCount = before + 1 And packets.Count = 1, "changed input with identical availability sends nothing")
        Call AddDoneQuest(1, 1): Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 2, "completion invalidates availability")
        bytes = packets(2): Call reader.SetData(bytes)
        Call reader.ReadInt16(): Call reader.ReadInt32(): Call reader.ReadInt16(): Call reader.ReadInt8()
        Call Expect(reader.ReadInt16() = 0, "empty replacement clears old offer")
    Case 8
        rejectAt = 1
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates(1000)
        Call Expect(packets.Count = 0, "failed enqueue not cached")
        Call ProcessQuestTurnInUpdates(1100): Call Expect(attempts = 1, "retry bounded")
        Call ProcessQuestTurnInUpdates(1250): Call Expect(packets.Count = 1, "retry succeeds")
        Call SendQuestTurnInSnapshot(1): Call ProcessQuestTurnInUpdates(1500)
        Call Expect(packets.Count = 1, "successful cache suppresses duplicate")
    Case 9
        rejectAt = 1: Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates(1000)
        UserList(1).VersionId = 2: UserList(1).Id = 141
        Call ProcessQuestTurnInUpdates(1300): Call Expect(packets.Count = 0, "slot reuse cannot inherit dirty work")
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates(1600)
        Call Expect(packets.Count = 1, "new session forced snapshot")
    Case 10
        NpcInfoCache(10).NumQuest = 255: ReDim NpcInfoCache(10).QuestNumber(1 To 255)
        For q = 1 To 255
            NpcInfoCache(10).QuestNumber(q) = q: QuestList(q).nombre = "Quest " & q
        Next q
        Call ResetQuestDiscoveryIndex()
        rejectAt = 2
        Call Expect(Not PublishQuestDiscovery(1, True), "partial page failure")
        Call Expect(packets.Count = 1, "first page accepted only")
        rejectAt = 0
        Call Expect(PublishQuestDiscovery(1, True), "retry restarts complete transaction")
        Call Expect(packets.Count = 3, "two retry pages, no silent truncation")
        For page = 2 To 3
            bytes = packets(page): Call reader.SetData(bytes)
            Call reader.ReadInt16(): Call reader.ReadInt32()
            Call Expect(reader.ReadInt16() = page - 2, "ordered pages")
            flags = reader.ReadInt8(): count = reader.ReadInt16()
            If page = 2 Then Call Expect(flags = 0 And count = 128, "full nonterminal page")
            If page = 3 Then Call Expect(flags = 1 And count = 127, "terminal remainder")
            For q = 1 To count
                Call reader.ReadInt16(): Call reader.ReadInt16(): Call reader.ReadString8()
            Next q
            Call Expect(reader.GetAvailable() = 0, "exact page layout")
        Next page
    Case 11
        QuestList(1).GlobalQuestIndex = 1
        ReDim GlobalQuestInfo(1).IsBossAlive(1 To 1)
        GlobalQuestInfo(1).IsActive = False
        Call Expect(Not IsQuestDiscoverable(1, 1), "inactive global event")
        GlobalQuestInfo(1).IsActive = True
        Call Expect(IsQuestDiscoverable(1, 1), "active event without boss")
        GlobalQuestInfo(1).IsBossAlive(1) = True
        Call Expect(Not IsQuestDiscoverable(1, 1), "live boss gate")
        Call ResetQuestDiscoveryIndex()
        Call Expect(PublishQuestDiscovery(1), "blocked initial event snapshot")
        GlobalQuestInfo(1).IsBossAlive(1) = False
        Call RefreshGlobalQuestTurnIns(1): Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 2, "boss-state change refreshes discovery")
    End Select
    capturing = False
    Call UnitTesting.RunTest("Quest discovery " & index, True)
    Exit Sub
Failed:
    capturing = False
    Call UnitTesting.RunTestError("Quest discovery " & index, Err.Description)
End Sub
Public Function TestSuite() As Boolean
    Dim before As Long, i As Integer
    before = UnitTesting.GetFailedTests()
    For i = 1 To 11
        Call RunCase(i)
    Next i
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
