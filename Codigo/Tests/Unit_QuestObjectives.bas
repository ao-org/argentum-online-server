Attribute VB_Name = "Unit_QuestObjectives"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean, rejectQueue As Boolean
Private packets As Collection
Public Function CapturePacket(ByRef queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    Dim bytes() As Byte, writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    queued = Not rejectQueue
    If UBound(bytes) >= 1 Then
        If CLng(bytes(0)) + CLng(bytes(1)) * 256 = ServerPacketID.eQuestObjectiveTemplates Then
            Dim value As Variant
            value = bytes
            If queued Then Call packets.Add(value)
        End If
    End If
    CapturePacket = True
End Function
Private Sub Expect(ByVal value As Boolean, ByVal label As String)
    If Not value Then Err.Raise vbObjectError + 844, "Quest objectives", label
End Sub
Private Sub Fixture()
    Call Unit_NpcBow.InitializeTestWorld()
    Set packets = New Collection
    rejectQueue = False
    With UserList(1)
        .Id = 101: .VersionId = 5
        .flags.UserLogged = True: .ConnectionDetails.ConnIDValida = True
        .pos.Map = 1: .pos.x = 50: .pos.y = 50
    End With
    ReDim QuestList(0 To 5)
    QuestList(1).RequiredNPCs = 1
    ReDim QuestList(1).RequiredNPC(1 To 1)
    QuestList(1).RequiredNPC(1).NpcIndex = 10
    QuestList(1).RequiredNPC(1).amount = 8
    UserList(1).QuestStats.Quests(1).QuestIndex = 1
    ReDim UserList(1).QuestStats.Quests(1).NPCsKilled(1 To 1)
    UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 3
    capturing = True
    Call SendQuestObjectiveSnapshot(1, True)
    Set packets = New Collection
End Sub
Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Fixture()
    Dim ids() As Integer, count As Integer, bytes() As Byte, npc As t_Npc
    Dim q As Integer, j As Integer, serial As Integer
    Select Case index
    Case 1
        Call CollectQuestObjectiveTemplates(1, ids, count)
        Call Expect(count = 1 And ids(1) = 10, "accepted 3/8 exact NPC template")
        UserList(1).QuestStats.Quests(1).QuestIndex = 0
        Call CollectQuestObjectiveTemplates(1, ids, count)
        Call Expect(count = 0, "unaccepted definitions do not mark")
    Case 2
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 8
        Call SendQuestObjectiveSnapshot(1)
        bytes = packets(1)
        Call Expect(UBound(bytes) = 4 And bytes(2) = 0, "objective completion pushes empty replacement")
    Case 3
        UserList(1).QuestStats.Quests(2).QuestIndex = 1
        ReDim UserList(1).QuestStats.Quests(2).NPCsKilled(1 To 1)
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 8
        Call CollectQuestObjectiveTemplates(1, ids, count)
        Call Expect(count = 1 And ids(1) = 10, "another unfinished accepted objective retains single marker")
        Call CleanQuestSlot(1, 2)
        bytes = packets(1)
        Call Expect(bytes(2) = 0, "abandon/completion cleanup publishes removal")
    Case 4
        QuestList(1).RequiredTargetNPCs = 1
        ReDim QuestList(1).RequiredTargetNPC(1 To 1)
        QuestList(1).RequiredTargetNPC(1).NpcIndex = 11
        QuestList(1).RequiredTargetNPC(1).amount = 2
        ReDim UserList(1).QuestStats.Quests(1).NPCsTarget(1 To 1)
        UserList(1).QuestStats.Quests(1).NPCsTarget(1) = 1
        Call CollectQuestObjectiveTemplates(1, ids, count)
        Call Expect(count = 2 And ids(2) = -11, "click-to-visit uses exact independent template")
        UserList(1).QuestStats.Quests(1).NPCsTarget(1) = 2
        Call CollectQuestObjectiveTemplates(1, ids, count)
        Call Expect(count = 1 And ids(1) = 10, "completed visit no longer marks")
    Case 5
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = -1
        Call CollectQuestObjectiveTemplates(1, ids, count)
        Call Expect(count = 0, "invalid progress is not guessed")
    Case 6
        npc.Numero = 10: npc.Name = "Shared name"
        Call OnNpcKilledUpdateQuest(1, npc)
        Call Expect(UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 4 And packets.Count = 1, "real progress publishes current counts")
        npc.Numero = 11
        Call OnNpcKilledUpdateQuest(1, npc)
        Call Expect(UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 4, "same name is not matching identity")
        npc.Numero = 10
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 7
        Call OnNpcKilledUpdateQuest(1, npc)
        bytes = packets(2)
        Call Expect(bytes(2) = 0, "real kill completion removes projection immediately")
    Case 7
        rejectQueue = True
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 8
        Call SendQuestObjectiveSnapshot(1)
        Call Expect(packets.Count = 0, "failed enqueue not cached")
        rejectQueue = False
        Call SendQuestObjectiveSnapshot(1)
        Call Expect(packets.Count = 1, "unchanged state retries after failed enqueue")
    Case 8
        UserList(1).VersionId = 6
        Call SendQuestObjectiveSnapshot(1)
        Call Expect(packets.Count = 1, "reused user slot gets snapshot")
        Call SendQuestObjectiveSnapshot(1, True)
        Call Expect(packets.Count = 2, "login always rebuilds authoritative state")
    Case 9
        For q = 1 To 5
            UserList(1).QuestStats.Quests(q).QuestIndex = q
            QuestList(q).RequiredNPCs = 255: QuestList(q).RequiredTargetNPCs = 255
            ReDim QuestList(q).RequiredNPC(1 To 255)
            ReDim QuestList(q).RequiredTargetNPC(1 To 255)
            ReDim UserList(1).QuestStats.Quests(q).NPCsKilled(1 To 255)
            ReDim UserList(1).QuestStats.Quests(q).NPCsTarget(1 To 255)
            For j = 1 To 255
                serial = serial + 1
                QuestList(q).RequiredNPC(j).NpcIndex = serial
                QuestList(q).RequiredNPC(j).amount = 1
                serial = serial + 1
                QuestList(q).RequiredTargetNPC(j).NpcIndex = serial
                QuestList(q).RequiredTargetNPC(j).amount = 1
            Next j
        Next q
        Call SendQuestObjectiveSnapshot(1)
        bytes = packets(1)
        Call Expect(UBound(bytes) = 20429, "all 2550 possible templates fit one bounded snapshot")
    Case 12
        QuestList(1).nombre = "Fangs of Fire"
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 2
        QuestList(1).RequiredNPC(1).amount = 25
        Call SendQuestObjectiveSnapshot(1)
        bytes = packets(1)
        Dim reader As New Network.Reader
        Call reader.SetData(bytes)
        Call Expect(reader.ReadInt16() = ServerPacketID.eQuestObjectiveTemplates, "objective packet")
        Call Expect(reader.ReadInt16() = 1 And reader.ReadInt16() = 10, "template prefix")
        Call Expect(reader.ReadInt8() = 1 And reader.ReadInt16() = 1, "accepted quest identity")
        Call Expect(reader.ReadString8() = "Fangs of Fire", "authoritative quest title")
        Call Expect(reader.ReadInt16() = 1 And reader.ReadInt16() = 10, "kill objective")
        Call Expect(reader.ReadInt16() = 2 And reader.ReadInt16() = 25 And reader.GetAvailable() = 0, "exact current and required, full payload")
    Case 11
        QuestList(1).RequiredTargetNPCs = 1
        ReDim QuestList(1).RequiredTargetNPC(1 To 1)
        QuestList(1).RequiredTargetNPC(1).NpcIndex = 10
        QuestList(1).RequiredTargetNPC(1).amount = 1
        ReDim UserList(1).QuestStats.Quests(1).NPCsTarget(1 To 1)
        Call CollectQuestObjectiveTemplates(1, ids, count)
        Call Expect(count = 2 And ids(1) = 10 And ids(2) = -10, "kill and visit retain different eligibility for same template")
    Case 10
        Erase UserList(1).QuestStats.Quests(1).NPCsKilled
        Call CollectQuestObjectiveTemplates(1, ids, count)
        Call Expect(count = 0, "missing metadata fails closed")
    End Select
    capturing = False
    Call UnitTesting.RunTest("Quest objectives " & index, True)
    Exit Sub
Failed:
    capturing = False
    Call UnitTesting.ProgressError("Quest objectives " & index, Err.Number, Err.Description)
    Call UnitTesting.RunTestError("Quest objectives " & index, Err.Description)
End Sub
Public Function TestSuite() As Boolean
    Dim before As Long, i As Integer
    before = UnitTesting.GetFailedTests()
        Call UnitTesting.RunTest("Quest objective packet ID", ServerPacketID.eQuestObjectiveTemplates = 223)
    For i = 1 To 12
        Call RunCase(i)
    Next i
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
