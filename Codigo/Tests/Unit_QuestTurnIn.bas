Attribute VB_Name = "Unit_QuestTurnIn"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private rejectQueue As Boolean
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
    CapturePacket = True
End Function
Private Sub Expect(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 845, "Quest turn-in", label
End Sub
Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Unit_NpcBow.InitializeTestWorld()
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
    Set packets = New Collection: capturing = True: rejectQueue = False
    Dim bytes() As Byte, reader As New Network.Reader
    Select Case index
    Case 1
        Call SendQuestTurnInSnapshot(1, True)
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
        bytes = packets(1): Call Expect(bytes(2) = 0, "kills alone do not imply readiness")
        UserList(1).invent.Object(1).ObjIndex = 50: UserList(1).invent.Object(1).amount = 2
        Call WriteChangeInventorySlot(1, 1)
        bytes = packets(2): Call Expect(bytes(2) = 1, "real inventory gain refreshes readiness")
        UserList(1).invent.Object(1).amount = 1
        Call WriteChangeInventorySlot(1, 1)
        bytes = packets(3): Call Expect(bytes(2) = 0, "real inventory loss clears readiness")
    Case 3
        Call SendQuestTurnInSnapshot(1, True)
        Call SendQuestTurnInSnapshot(1)
        Call Expect(packets.Count = 1, "unchanged state is not resent")
        Call CleanQuestSlot(1, 1)
        bytes = packets(2): Call Expect(bytes(2) = 0, "confirmed removal clears ready state")
    Case 4
        rejectQueue = True: Call SendQuestTurnInSnapshot(1, True)
        rejectQueue = False: Call SendQuestTurnInSnapshot(1)
        Call Expect(packets.Count = 1, "failed enqueue does not prevent retry")
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 7
        Call SendQuestTurnInSnapshot(1)
        bytes = packets(2): Call Expect(bytes(2) = 0, "lost requirement clears readiness")
    Case 5
        QuestList(1).RewardOBJs = 1
        UserList(1).CurrentInventorySlots = 0
        Call SendQuestTurnInSnapshot(1, True)
        bytes = packets(1): Call Expect(bytes(2) = 0, "reward admission uses real eligibility")
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
    For i = 1 To 5
        Call RunCase(i)
    Next i
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
