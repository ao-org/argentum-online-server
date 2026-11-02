Attribute VB_Name = "Unit_QuestTurnIn"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private feedbackPackets As Collection
Private rejectQueue As Boolean
Private mMeasurements As String
Public Function Measurements() As String
    Measurements = mMeasurements
End Function
Public Function CapturePacket(ByRef queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    Dim bytes() As Byte, writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    queued = Not rejectQueue
    If UBound(bytes) >= 1 Then
        If CLng(bytes(0)) + 256& * bytes(1) = ServerPacketID.eQuestTurnIns And queued Then
            Dim value As Variant: value = bytes
            Call packets.Add(value)
        End If
    End If
    If UBound(bytes) >= 1 And queued Then
        If CLng(bytes(0)) + 256& * bytes(1) = ServerPacketID.eQuestFeedback Then
            Dim feedback As Variant: feedback = bytes
            Call feedbackPackets.Add(feedback)
        End If
    End If
    CapturePacket = True
End Function
Private Sub Expect(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 845, "Quest turn-in", label
End Sub
Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Unit_NpcBow.InitializeTestWorld()
    Call ResetTurnInTestCounters()
    ReDim QuestList(0 To 2)
    ReDim NpcInfoCache(1 To 20)
    QuestList(1).nombre = "Quest": QuestList(1).TalkTo = 11
    QuestList(1).RequiredNPCs = 1
    ReDim QuestList(1).RequiredNPC(1 To 1)
    QuestList(1).RequiredNPC(1).NpcIndex = 10
    QuestList(1).RequiredNPC(1).amount = 8
    UserList(1).Id = 100: UserList(1).VersionId = 7
    UserList(1).flags.UserLogged = True
    UserList(1).QuestStats.Quests(1).QuestIndex = 1
    ReDim UserList(1).QuestStats.Quests(1).NPCsKilled(1 To 1)
    UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 8
    UserList(1).CurrentInventorySlots = 3
    NpcInfoCache(11).Exists = True: NpcInfoCache(11).name = "Recipient"
    NpcInfoCache(12).Exists = True: NpcInfoCache(12).name = "Other recipient"
    NpcInfoCache(12).NumQuest = 1
    ReDim NpcInfoCache(12).QuestNumber(1 To 1)
    NpcInfoCache(12).QuestNumber(1) = 1
    Call ResetQuestTurnInRecipients()
    Set packets = New Collection: Set feedbackPackets = New Collection
    capturing = True: rejectQueue = False
    Dim bytes() As Byte, reader As New Network.Reader
    Select Case index
    Case 1
        Call SendQuestTurnInSnapshot(1, True)
        Call ProcessQuestTurnInUpdates()
        bytes = packets(1): Call reader.SetData(bytes)
        Call Expect(reader.ReadInt16() = ServerPacketID.eQuestTurnIns, "packet ID")
        Call Expect(reader.ReadInt8() = 1 And reader.ReadInt16() = 1, "ready accepted quest")
        Call Expect(reader.ReadString8() = "Quest" And reader.ReadInt8() = 2, "two valid recipients")
        Call Expect(reader.ReadInt16() = 11 And reader.ReadString8() = "Recipient", "TalkTo differs from killed NPC")
        Call Expect(reader.ReadInt16() = 12 And reader.ReadString8() = "Other recipient", "offering NPC also valid")
        Call Expect(reader.GetAvailable() = 0, "exact packet bytes")
    Case 2
        QuestList(1).RequiredOBJs = 1
        ReDim QuestList(1).RequiredOBJ(1 To 1)
        QuestList(1).RequiredOBJ(1).ObjIndex = 50: QuestList(1).RequiredOBJ(1).amount = 2
        Call SendQuestTurnInSnapshot(1, True)
        Call ProcessQuestTurnInUpdates()
        bytes = packets(1): Call Expect(bytes(2) = 0, "kills alone do not imply readiness")
        UserList(1).invent.Object(1).ObjIndex = 50: UserList(1).invent.Object(1).amount = 2
        Call WriteChangeInventorySlot(1, 1)
        Call ProcessQuestTurnInUpdates()
        bytes = packets(2): Call Expect(bytes(2) = 1, "real inventory gain refreshes readiness")
        UserList(1).invent.Object(1).amount = 1
        Call WriteChangeInventorySlot(1, 1)
        Call ProcessQuestTurnInUpdates()
        bytes = packets(3): Call Expect(bytes(2) = 0, "real inventory loss clears readiness")
    Case 3
        Call SendQuestTurnInSnapshot(1, True)
        Call ProcessQuestTurnInUpdates()
        Call SendQuestTurnInSnapshot(1)
        Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 1, "unchanged state is not resent")
        Call CleanQuestSlot(1, 1)
        Call ProcessQuestTurnInUpdates()
        bytes = packets(2): Call Expect(bytes(2) = 0, "confirmed removal clears ready state")
    Case 4
        rejectQueue = True: Call SendQuestTurnInSnapshot(1, True)
        Call ProcessQuestTurnInUpdates(1000)
        rejectQueue = False: Call SendQuestTurnInSnapshot(1)
        Call ProcessQuestTurnInUpdates(1300)
        Call Expect(packets.Count = 1, "failed enqueue does not prevent retry")
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 7
        Call SendQuestTurnInSnapshot(1)
        Call ProcessQuestTurnInUpdates(1600)
        bytes = packets(2): Call Expect(bytes(2) = 0, "lost requirement clears readiness")
    Case 5
        QuestList(1).RewardOBJs = 1
        UserList(1).CurrentInventorySlots = 0
        Call SendQuestTurnInSnapshot(1, True)
        Call ProcessQuestTurnInUpdates()
        bytes = packets(1): Call Expect(bytes(2) = 0, "reward admission uses real eligibility")
    Case 6
        ReDim QuestList(0 To 5)
        ReDim NpcInfoCache(1 To 100)
        Dim q As Integer, npc As Integer, n As Integer
        For q = 1 To 5
            QuestList(q).nombre = "Quest"
            UserList(1).QuestStats.Quests(q).QuestIndex = q
        Next q
        For npc = 1 To 100
            NpcInfoCache(npc).Exists = True: NpcInfoCache(npc).name = "Recipient"
            NpcInfoCache(npc).NumQuest = 5
            ReDim NpcInfoCache(npc).QuestNumber(1 To 5)
            For q = 1 To 5
                NpcInfoCache(npc).QuestNumber(q) = q
            Next q
        Next npc
        Call ResetQuestTurnInRecipients()
        Call SendQuestTurnInSnapshot(1, True)
        Call ProcessQuestTurnInUpdates()
        bytes = packets(1): Call reader.SetData(bytes)
        Call Expect(reader.ReadInt16() = ServerPacketID.eQuestTurnIns And reader.ReadInt8() = 5, "all readiness retained")
        For q = 1 To 5
            Call Expect(reader.ReadInt16() = q And reader.ReadString8() = "Quest", "quest identity")
            n = reader.ReadInt8()
            If q <= 3 Then
                Call Expect(n = 100, "complete recipient list")
            Else
                Call Expect(n = 0, "packet budget reports unknown instead of partial recipients")
            End If
            For npc = 1 To n
                Call Expect(reader.ReadInt16() = npc And reader.ReadString8() = "Recipient", "recipient bytes")
            Next npc
        Next q
        Call Expect(reader.GetAvailable() = 0, "bounded payload consumed")
    Case 7
        LastUser = 4: ReDim GlobalQuestInfo(1 To 1)
        Call FinishGlobalQuest(1, 1, 0, 0)
        Call RefreshGlobalQuestTurnIns(0)
        Call RefreshGlobalQuestTurnIns(-1)
        Call RefreshGlobalQuestTurnIns(2)
        Call ProcessQuestTurnInUpdates()
        Call Expect(GlobalScanCount = 0 And EvaluationCount = 0, "non-global and invalid indexes do not scan or evaluate")
    Case 8
        Dim i As Integer
        For i = 1 To 64
            Call WriteChangeInventorySlot(1, 1)
        Next i
        Call Expect(EvaluationCount = 0 And packets.Count = 0, "bulk writes defer presentation")
        Call ProcessQuestTurnInUpdates()
        Call Expect(EvaluationCount = 1 And packets.Count = 1, "bulk writes evaluate final state once")
        Call SendQuestTurnInSnapshot(1)
        Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 1, "unchanged fingerprint suppresses snapshot")
    Case 9
        rejectQueue = True
        Call SendQuestTurnInSnapshot(1, True)
        Call ProcessQuestTurnInUpdates(1000)
        Call ProcessQuestTurnInUpdates(1100)
        Call Expect(EvaluationCount = 1 And packets.Count = 0, "failed enqueue does not busy loop")
        rejectQueue = False
        Call ProcessQuestTurnInUpdates(1250)
        Call Expect(EvaluationCount = 2 And packets.Count = 1, "pending work retries after bounded delay")
    Case 10
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates()
        Set packets = New Collection: Call ResetTurnInTestCounters()
        Call WriteChangeInventorySlot(1, 1)
        Call CleanQuestSlot(1, 1)
        Call WriteChangeInventorySlot(1, 2)
        Call ProcessQuestTurnInUpdates()
        bytes = packets(1)
        Call Expect(EvaluationCount = 1 And packets.Count = 1 And bytes(2) = 0, "reward and completion batch cannot restore completed guidance")
    Case 11
        Call SendQuestTurnInSnapshot(1, True)
        UserList(1).VersionId = 8: UserList(1).Id = 200
        Call ProcessQuestTurnInUpdates()
        Call Expect(EvaluationCount = 0 And packets.Count = 0, "old identity work is discarded")
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates()
        Call Expect(EvaluationCount = 1 And packets.Count = 1, "new identity can synchronize")
    Case 12
        Dim batch As Long, started As Long, beforeMs As Double, afterMs As Double, beforeCount As Long
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates()
        Call ResetTurnInTestCounters(): Set packets = New Collection
        started = GetTickCountRaw()
        For batch = 1 To 1000
            For i = 1 To 64
                Call EvaluateTurnInImmediatelyForTest(1)
            Next i
        Next batch
        beforeMs = TicksElapsed(started, GetTickCountRaw()): beforeCount = EvaluationCount
        Call ResetTurnInTestCounters(): started = GetTickCountRaw()
        For batch = 1 To 1000
            For i = 1 To 64
                Call SendQuestTurnInSnapshot(1)
            Next i
            Call ProcessQuestTurnInUpdates()
        Next batch
        afterMs = TicksElapsed(started, GetTickCountRaw())
        Call Expect(beforeCount = 64000 And EvaluationCount = 1000, "bulk benchmark counts")
        mMeasurements = "Bulk inventory: 1 user, 1000 batches x64 mutations; immediate=" & beforeCount & " evaluations/" & beforeMs & " ms; batched=" & EvaluationCount & " evaluations/" & afterMs & " ms"
    Case 13
        LastUser = 4: ReDim GlobalQuestInfo(1 To 1)
        QuestList(1).GlobalQuestIndex = 1: QuestList(1).GlobalQuestThresholdNeeded = 1
        Dim u As Integer, beforeScans As Long
        For u = 1 To 4
            UserList(u).Id = 100 + u: UserList(u).VersionId = 7: UserList(u).flags.UserLogged = True
            UserList(u).QuestStats.Quests(1).QuestIndex = 1
            ReDim UserList(u).QuestStats.Quests(1).NPCsKilled(1 To 1)
            UserList(u).QuestStats.Quests(1).NPCsKilled(1) = 8
        Next u
        Call ResetTurnInTestCounters(): started = GetTickCountRaw()
        For batch = 1 To 20
            GlobalQuestInfo(1).GatheringGlobalCounter = batch Mod 2
            Call LegacyGlobalRefreshForTest(1): Call LegacyGlobalRefreshForTest(1)
        Next batch
        beforeMs = TicksElapsed(started, GetTickCountRaw()): beforeCount = EvaluationCount: beforeScans = GlobalScanCount
        Call ResetTurnInTestCounters(): started = GetTickCountRaw()
        For batch = 1 To 20
            GlobalQuestInfo(1).GatheringGlobalCounter = batch Mod 2
            Call RefreshGlobalQuestTurnIns(1): Call RefreshGlobalQuestTurnIns(1)
            Call ProcessQuestTurnInUpdates()
        Next batch
        afterMs = TicksElapsed(started, GetTickCountRaw())
        Call Expect(beforeScans = 40 And beforeCount = 160 And GlobalScanCount = 20 And EvaluationCount = 80, "global benchmark counts")
        mMeasurements = mMeasurements & vbCrLf & "Global events: 4 users, 20 threshold changes x2 notifications; immediate=" & beforeScans & " scans/" & beforeCount & " evaluations/" & beforeMs & " ms; batched=" & GlobalScanCount & " scans/" & EvaluationCount & " evaluations/" & afterMs & " ms"
        Call RefreshGlobalQuestTurnIns(1): Call ProcessQuestTurnInUpdates()
        Call Expect(GlobalScanCount = 20 And EvaluationCount = 80, "unchanged global state does no work")
    Case 14
        QuestList(1).RequiredTargetNPCs = 1
        ReDim QuestList(1).RequiredTargetNPC(1 To 1)
        QuestList(1).RequiredTargetNPC(1).NpcIndex = 11: QuestList(1).RequiredTargetNPC(1).amount = 1
        ReDim UserList(1).QuestStats.Quests(1).NPCsTarget(1 To 1)
        UserList(1).QuestStats.Quests(1).NPCsTarget(1) = 1
        QuestList(1).RequiredOBJs = 1: ReDim QuestList(1).RequiredOBJ(1 To 1)
        QuestList(1).RequiredOBJ(1).ObjIndex = 50: QuestList(1).RequiredOBJ(1).amount = 5
        UserList(1).invent.Object(1).ObjIndex = 50: UserList(1).invent.Object(1).amount = 3
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates()
        bytes = feedbackPackets(1): Call reader.SetData(bytes)
        Call Expect(reader.ReadInt16() = ServerPacketID.eQuestFeedback, "feedback packet ID")
        Call Expect(reader.ReadInt32() > 0 And reader.ReadInt8() = 0 And reader.ReadInt8() = 1 And reader.ReadInt8() = 1, "initial synchronization silent")
        Call Expect(reader.ReadInt16() = 1 And reader.ReadInt32() > 0 And reader.ReadString8() = "Quest", "quest lifecycle identity")
        Call Expect(reader.ReadInt8() = 0 And reader.ReadInt16() = 3, "missing item prevents readiness, complete goals retained")
        Call Expect(reader.ReadInt8() = 1 And reader.ReadInt16() = 1 And reader.ReadInt16() = 10 And reader.ReadInt32() = 8 And reader.ReadInt32() = 8, "final kill count retained")
        Call Expect(reader.ReadInt8() = 2 And reader.ReadInt16() = 1 And reader.ReadInt16() = 11 And reader.ReadInt32() = 1 And reader.ReadInt32() = 1, "visit count")
        Call Expect(reader.ReadInt8() = 3 And reader.ReadInt16() = 1 And reader.ReadInt16() = 50 And reader.ReadInt32() = 3 And reader.ReadInt32() = 5, "actual item count")
        UserList(1).invent.Object(1).amount = 5
        Call WriteChangeInventorySlot(1, 1): Call ProcessQuestTurnInUpdates()
        Call Expect(feedbackPackets.Count = 2, "item changes publish updated counts")
        Call SendQuestTurnInSnapshot(1): Call ProcessQuestTurnInUpdates()
        Call Expect(feedbackPackets.Count = 2, "duplicate feedback suppressed")
    Case 15
        UserList(1).flags.TargetNPC.ArrayIndex = 1: NpcList(1).Char.charindex = 2
        QuestList(1).RewardGLD = 200000
        QuestList(1).RewardOBJs = 1: ReDim QuestList(1).RewardOBJ(1 To 1)
        QuestList(1).RewardOBJ(1).ObjIndex = 50: QuestList(1).RewardOBJ(1).amount = 3
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates()
        Call FinishQuest(1, 1, 1): Call ProcessQuestTurnInUpdates()
        Call Expect(UserList(1).QuestStats.NumQuestsDone = 1, "real turn-in completed")
        Dim packet As Variant, receiptCount As Integer, rewardCount As Integer, kind As Integer, rewardId As Integer, amount As Long
        For Each packet In feedbackPackets
            bytes = packet: Call reader.SetData(bytes)
            Call reader.ReadInt16(): Call reader.ReadInt32()
            If reader.ReadInt8() = 1 Then
                receiptCount = receiptCount + 1
                Call Expect(reader.ReadInt8() = 0 And reader.ReadInt8() = 1 And reader.ReadInt16() = 1, "correlated completion")
                Call reader.ReadInt32(): Call reader.ReadString8(): Call reader.ReadInt8(): Call reader.ReadInt16()
                rewardCount = reader.ReadInt16()
                Call Expect(rewardCount = 2, "only actual gold and item grants")
                kind = reader.ReadInt8(): rewardId = reader.ReadInt16(): amount = reader.ReadInt32()
                Call Expect(kind = 3 And rewardId = 0 And amount = UserList(1).Stats.Banco And amount = 200000 * SvrConfig.GetValue("GoldMult"), "bank gold uses actual multiplier")
                Call Expect(reader.ReadInt8() = 4 And reader.ReadInt16() = 50 And reader.ReadInt32() = CountInventoryObjects(1, 50), "actual awarded objects")
            End If
        Next packet
        Call Expect(receiptCount = 1, "one receipt after actual grant")
    Case 16
        UserList(1).flags.TargetNPC.ArrayIndex = 1: NpcList(1).Char.charindex = 2
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 0
        Call FinishQuest(1, 1, 1): Call ProcessQuestTurnInUpdates()
        Call Expect(feedbackPackets.Count = 0 And UserList(1).QuestStats.Quests(1).QuestIndex = 1, "rejected turn-in has no receipt")
    Case 17
        Dim rewards As New Collection, lifecycleId As Long
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates()
        lifecycleId = CaptureQuestFeedbackLifecycle(1, 1, 1)
        Call AddGrantedQuestReward(rewards, 1, 0, 27)
        Call AddGrantedQuestReward(rewards, 4, 50, 0)
        Call QueueQuestCompletionFeedback(1, lifecycleId, 1, rewards)
        rejectQueue = True: Call ProcessQuestTurnInUpdates(1000)
        rejectQueue = False: Call ProcessQuestTurnInUpdates(1250)
        Call Expect(feedbackPackets.Count = 2, "failed receipt retries without duplication")
        Call ProcessQuestTurnInUpdates(1500)
        Call Expect(feedbackPackets.Count = 2, "accepted receipt is not replayed")
        Call BeginQuestFeedbackLifecycle(1, 1)
        Call Expect(CaptureQuestFeedbackLifecycle(1, 1, 1) <> lifecycleId, "repeat acceptance has new lifecycle")
    End Select
    capturing = False
    Call UnitTesting.RunTest("Quest turn-in " & index, True)
    Exit Sub
Failed:
    capturing = False
    Call UnitTesting.RunTestError("Quest turn-in " & index, Err.Description)
End Sub
Public Function TestSuite() As Boolean
    Dim before As Long, i As Integer
    before = UnitTesting.GetFailedTests()
    For i = 1 To 17
        Call RunCase(i)
    Next i
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
