Attribute VB_Name = "Unit_NpcInspection"
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
    Dim value As Variant
    value = bytes
    If queued Then Call packets.Add(value)
    CapturePacket = True
End Function
Private Sub Expect(ByVal value As Boolean, ByVal label As String)
    If Not value Then Err.Raise vbObjectError + 816, "NPC inspection", label
End Sub
Private Sub Fixture()
    Call Unit_NpcBow.InitializeTestWorld()
    Set packets = New Collection
    rejectQueue = False
    With UserList(1)
        .Id = 100: .VersionId = 5
        .flags.UserLogged = True: .ConnectionDetails.ConnIDValida = True
        .pos.Map = 1: .pos.x = 50: .pos.y = 50
        .Stats.MaxHp = 100: .Stats.MinHp = 100
    End With
    With NpcList(1)
        .flags.NPCActive = True: .flags.invisible = 0: .VersionId = 4
        .Char.charindex = 2: .Char.body = 1
        .Numero = 10: .pos.Map = 1: .pos.x = 51: .pos.y = 50
        .Stats.MaxHp = 100: .Stats.MinHp = 100: .Stats.MaxHit = 20
        .Hostile = 1: .Attackable = 1: .IntervaloAtaque = 2000
        .Veneno = 1: .AttackRange = 1: .flags.LanzaSpells = 0
    End With
    CharList(2) = 1
    ReDim QuestList(0 To 2)
    QuestList(1).RequiredNPCs = 1
    ReDim QuestList(1).RequiredNPC(1 To 1)
    QuestList(1).RequiredNPC(1).NpcIndex = 10
    QuestList(1).RequiredNPC(1).amount = 8
    QuestList(2).RequiredNPCs = 1
    ReDim QuestList(2).RequiredNPC(1 To 1)
    QuestList(2).RequiredNPC(1).NpcIndex = 11
    QuestList(2).RequiredNPC(1).amount = 2
    UserList(1).QuestStats.Quests(1).QuestIndex = 1
    ReDim UserList(1).QuestStats.Quests(1).NPCsKilled(1 To 1)
    UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 3
    UserList(1).QuestStats.Quests(2).QuestIndex = 2
    ReDim UserList(1).QuestStats.Quests(2).NPCsKilled(1 To 1)
    capturing = True
End Sub
Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Fixture()
    Call SubscribeNpcInspection(1, index, 1, 2, 10)
    Call Expect(packets.Count = 1, "initial snapshot")
    Dim bytes() As Byte
    bytes = packets(1)
    Call Expect(bytes(13) = 3 And bytes(14) = 1, "physical threat and offensive poison")
    Call Expect(bytes(15) = 1 And bytes(16) = 1 And bytes(19) = 3 And bytes(23) = 8, "exact template objective")
    Select Case index
    Case 1
        Call RefreshNpcInspection(1)
        Call Expect(packets.Count = 1, "unchanged snapshot not resent")
        UserList(1).Stats.MinHp = 1
        Call RefreshNpcInspection(1)
        Call Expect(packets.Count = 1, "current HP cannot change threat")
    Case 2
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 8
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(15) = 0, "completed kills disappear")
        QuestList(1).TalkTo = 10
        Call RefreshNpcInspection(1)
        bytes = packets(3): Call Expect(bytes(15) = 1 And bytes(18) = 2, "real ready turn-in separate from kill objective")
    Case 3
        NpcList(1).VersionId = 5
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(12) = 0, "slot reuse invalidates")
    Case 4
        UserList(1).pos.Map = 2
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(12) = 0, "map replacement invalidates")
    Case 5
        UserList(1).VersionId = 6
        Call RefreshNpcInspection(1)
        Call Expect(packets.Count = 1, "old session never receives data")
    Case 6
        NpcList(1).flags.invisible = 1
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(12) = 0, "hidden NPC data removed")
    Case 7
        NpcList(1).flags.LanzaSpells = 1
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(13) = 0 And bytes(14) = 5, "unknown spell damage is not low threat")
    Case 8
        NpcList(1).Hostile = 0
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(13) = 0, "friendly NPC not rated")
    Case 9
        rejectQueue = True
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 4
        Call RefreshNpcInspection(1)
        rejectQueue = False
        Call RefreshNpcInspection(1)
        Call Expect(packets.Count = 2, "rejected enqueue does not consume changed snapshot")
    Case 10
        QuestList(2).RequiredNPC(1).NpcIndex = 10
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(15) = 2, "multiple active quests")
        UserList(1).QuestStats.Quests(1).QuestIndex = 0
        Call RefreshNpcInspection(1)
        bytes = packets(3): Call Expect(bytes(15) = 1 And bytes(16) = 2, "abandon removes relevance")
    Case 11
        NpcList(1).Stats.MaxHit = 0
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(13) = 0, "missing physical damage is unknown")
    Case 12
        NpcList(1).Stats.MinHp = 0
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(12) = 0, "dead target is unavailable")
    Case 13
        UserList(1).Stats.MaxHp = 200
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(13) = 2, "lasting maximum HP updates threat")
    Case 14
        QuestList(1).RequiredNPCs = 33
        ReDim QuestList(1).RequiredNPC(1 To 33)
        ReDim UserList(1).QuestStats.Quests(1).NPCsKilled(1 To 33)
        Dim i As Integer
        For i = 1 To 33
            QuestList(1).RequiredNPC(i).NpcIndex = 10
            QuestList(1).RequiredNPC(i).amount = 8
        Next i
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(15) = 160 And UBound(bytes) = 367, "bounded snapshot explicitly reports overflow")
    Case 15
        rejectQueue = True
        NpcList(1).flags.invisible = 1
        Call RefreshNpcInspection(1)
        rejectQueue = False
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(12) = 0, "failed invalidation enqueue retried")
    Case 16
        Call SubscribeNpcInspection(1, 1000, 1, 0, 0)
        UserList(1).Stats.MaxHp = 200
        Call RefreshNpcInspection(1)
        Call Expect(packets.Count = 1, "unsubscribe stops updates")
    Case 17
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 8
        QuestList(1).TalkTo = 10
        QuestList(1).RewardOBJs = 1
        UserList(1).CurrentInventorySlots = 0
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(15) = 0, "full reward inventory is not ready")
        UserList(1).CurrentInventorySlots = 1
        Call RefreshNpcInspection(1)
        bytes = packets(3): Call Expect(bytes(15) = 1 And bytes(18) = 2, "free reward slot admits turn-in")
    Case 18
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 8
        QuestList(1).TalkTo = 10
        QuestList(1).RequiredClassesCount = 1
        ReDim QuestList(1).RequiredClass(1 To 1)
        QuestList(1).RequiredClass(1) = e_Class.Hunter
        UserList(1).clase = e_Class.Mage
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(15) = 0, "wrong class is not ready")
    Case 19
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = 8
        QuestList(1).RequiredTargetNPCs = 1
        ReDim QuestList(1).RequiredTargetNPC(1 To 1)
        QuestList(1).RequiredTargetNPC(1).NpcIndex = 10
        QuestList(1).RequiredTargetNPC(1).amount = 1
        ReDim UserList(1).QuestStats.Quests(1).NPCsTarget(1 To 1)
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(15) = 1 And bytes(18) = 3, "real click-to-visit objective")
        UserList(1).QuestStats.Quests(1).NPCsTarget(1) = 1
        Call RefreshNpcInspection(1)
        bytes = packets(3): Call Expect(bytes(15) = 0, "completed visit disappears")
    Case 20
        UserList(1).QuestStats.Quests(1).NPCsKilled(1) = -1
        Call RefreshNpcInspection(1)
        bytes = packets(2): Call Expect(bytes(15) = 0, "invalid progress is omitted instead of emitting malformed data")
    End Select
    capturing = False
    Call UnitTesting.RunTest("NPC inspection " & index, True)
    Exit Sub
Failed:
    capturing = False
    Call UnitTesting.ProgressError("NPC inspection " & index, Err.Number, Err.Description)
    Call UnitTesting.RunTestError("NPC inspection " & index, Err.Description)
End Sub
Public Function TestSuite() As Boolean
    Dim before As Long, i As Integer
    before = UnitTesting.GetFailedTests()
    #If PYMMO = 1 Then
        Call UnitTesting.RunTest("NPC inspection packet IDs", ServerPacketID.eNpcInspection = 221 And ClientPacketID.eNpcInspectionSubscribe = 317)
    #Else
        Call UnitTesting.RunTest("NPC inspection conditional packet IDs", ServerPacketID.eNpcInspection = 222 And ClientPacketID.eNpcInspectionSubscribe = 320)
    #End If
    Call UnitTesting.RunTest("NPC threat missing data", InspectionThreatBand(0, 10, 2000) = 0 And InspectionThreatBand(100, 0, 2000) = 0)
    Call UnitTesting.RunTest("NPC threat thresholds", InspectionThreatBand(100, 5, 2000) = 1 And InspectionThreatBand(100, 6, 2000) = 2 And InspectionThreatBand(100, 15, 2000) = 3 And InspectionThreatBand(100, 30, 2000) = 4)
    Call UnitTesting.RunTest("NPC threat attack interval", InspectionThreatBand(100, 15, 1000) = 4)
    For i = 1 To 20
        Call RunCase(i)
    Next i
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
