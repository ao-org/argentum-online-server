Attribute VB_Name = "Unit_GatheringFeedback"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Public Function CapturePacket(ByRef queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    On Error GoTo Done
    Dim bytes() As Byte, writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    If UBound(bytes) >= 1 Then
        If CLng(bytes(0)) + CLng(bytes(1)) * 256 = ServerPacketID.eGatheringFeedback Then
            Dim value As Variant
            value = bytes
            Call packets.Add(value)
        End If
    End If
Done:
    queued = True
    CapturePacket = True
End Function
Private Sub Expect(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 819, "Gathering feedback", message
End Sub
Private Sub Fixture()
    Call Unit_NpcBow.InitializeTestWorld()
    Set packets = New Collection
    With UserList(1)
        .Id = 100: .VersionId = 5
        .flags.UserLogged = True
        .pos.Map = 1: .pos.x = 50: .pos.y = 50
        .invent.EquippedWorkingToolObjIndex = 1
    End With
    Call StartAutomatedAction(51, 50, e_Skill.Talar, 1)
    Call BeginGatheringFeedback(1, 7, e_Skill.Talar)
    capturing = True
End Sub
Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Fixture()
    Dim bytes() As Byte, count As Long, reason As e_GatheringReason
    Select Case index
    Case 1
        Call SendGatheringFeedback(1, GatheringStarted)
        bytes = packets(1)
        Call Expect(UBound(bytes) = 23, "exact 24-byte result")
        Call Expect(bytes(2) = 7 And bytes(6) = 1 And bytes(10) = 1, "request sequence map")
        Call Expect(bytes(12) = e_Skill.Talar And bytes(13) = GatheringStarted And bytes(15) = 1, "accepted activity")
    Case 2
        Call SendGatheringFeedback(1, GatheringReward, GatheringNoReason, 58, 3)
        Call SendGatheringFeedback(1, GatheringReward, GatheringNoReason, 59, 7)
        bytes = packets(2)
        Call Expect(bytes(6) = 2 And bytes(16) = 59 And bytes(18) = 7, "different actual reward quantities")
    Case 3
        Call SendGatheringFeedback(1, GatheringFailedRoll)
        Call Expect(UserList(1).AutomatedAction.IsActive, "failed roll does not cancel gameplay")
        bytes = packets(1)
        Call Expect(bytes(13) = GatheringFailedRoll And bytes(15) = 1 And bytes(16) = 0 And bytes(18) = 0, "failed roll has no invented reward")
    Case 4
        Call SendGatheringFeedback(1, GatheringBlocked, GatheringStamina)
        Call ResetUserAutomatedActions(1)
        Call ResetUserAutomatedActions(1)
        Call Expect(packets.Count = 1 And Not UserList(1).AutomatedAction.IsActive, "one blocking reason not duplicate stop")
    Case 5
        Call ResetUserAutomatedActions(1)
        Call SendGatheringFeedback(1, GatheringReward, GatheringNoReason, 58, 2)
        bytes = packets(2)
        Call Expect(bytes(13) = GatheringReward And bytes(15) = 0, "late real reward cannot reactivate")
    Case 6
        UserList(1).VersionId = 6
        Call SendGatheringFeedback(1, GatheringReward, GatheringNoReason, 58, 3)
        Call Expect(packets.Count = 0, "reused user lifetime cannot receive stale feedback")
    Case 7
        Call BeginGatheringFeedback(1, 0, e_Skill.Talar)
        Call SendGatheringFeedback(1, GatheringReward, GatheringNoReason, 58, 3)
        Call Expect(packets.Count = 0, "legacy request has no correlated feedback")
    Case 8
        Call RejectGatheringRequest(1, 8, e_Skill.Mineria, GatheringTool)
        Call SendGatheringFeedback(1, GatheringFailedRoll)
        bytes = packets(2)
        Call Expect(bytes(2) = 7 And bytes(12) = e_Skill.Talar, "rejected replacement preserves actual old action context")
        Call Expect(UserList(1).AutomatedAction.IsActive And UserList(1).AutomatedAction.skill = e_Skill.Talar, "rejection does not change retry rules")
    Case 9
        UserList(1).Stats.MinSta = 0
        Call Expect(Not DecreaseUserStamina(1, 1), "real stamina gate")
        Call Expect(packets.Count = 1, "actual gate emits exactly one result")
        bytes = packets(1)
        Call Expect(bytes(13) = GatheringBlocked And bytes(14) = GatheringStamina, "actual stamina reason")
    Case 10
        UserList(1).invent.EquippedWorkingToolObjIndex = 0
        Call Expect(Not CanUserExtractWood(1, 51, 50, reason) And reason = GatheringTool, "wood tool validator")
        Call Expect(Not CanUserExtractMinerals(1, 51, 50, reason) And reason = GatheringTool, "ore tool validator")
        Call Expect(Not CanUserFish(1, 51, 50, reason) And reason = GatheringTool, "fish tool validator")
    Case 11
        Call SendGatheringFeedback(1, GatheringGroundReward, GatheringNoReason, 58, 3)
        bytes = packets(1)
        Call Expect(bytes(13) = GatheringGroundReward And bytes(18) = 3, "ground award not inventory collection")
    Case 12
        Call SendGatheringFeedback(1, GatheringReward, GatheringNoReason, 58, 0)
        Call Expect(packets.Count = 0, "zero yield cannot claim collection")
    Case 13
        Dim obj As t_Obj, destination As t_WorldPos
        obj.ObjIndex = 58: obj.amount = 3
        Call ReportGatheringDrop(1, obj, destination)
        bytes = packets(1)
        Call Expect(bytes(13) = GatheringBlocked And bytes(14) = GatheringNoDropSpace And bytes(15) = 1, "no floor placement is not an award or a stopped loop")
        Call Expect(bytes(16) = 0 And bytes(18) = 0 And UserList(1).AutomatedAction.IsActive, "no fabricated floor award")
    Case 14
        UserList(1).CurrentInventorySlots = 1
        ObjData(58).OBJType = e_OBJType.otMinerals
        obj.ObjIndex = 58: obj.amount = 3
        Call Expect(InsertGatheredItem(1, obj), "actual first insertion")
        obj.amount = 2
        Call Expect(InsertGatheredItem(1, obj), "actual stacked insertion")
        bytes = packets(2)
        Call Expect(bytes(18) = 2 And UserList(1).invent.Object(1).amount = 5, "report awarded delta, not resulting inventory total")
    Case 15
        UserList(1).CurrentInventorySlots = 1
        ObjData(58).OBJType = e_OBJType.otMinerals
        obj.ObjIndex = 58: obj.amount = CLng(GetMaxInvOBJ()) + 1
        Call InsertGatheredItem(1, obj)
        Call Expect(packets.Count = 0, "existing oversize empty-slot behavior cannot invent a full input-quantity award")
    End Select
    capturing = False
    Call UnitTesting.RunTest("Gathering feedback " & index, True)
    Exit Sub
Failed:
    capturing = False
    Call UnitTesting.ProgressError("Gathering feedback " & index, Err.Number, Err.Description)
    Call UnitTesting.RunTestError("Gathering feedback " & index, Err.Description)
End Sub
Public Function TestSuite() As Boolean
    Dim before As Long, index As Integer
    before = UnitTesting.GetFailedTests()
    #If PYMMO = 1 Then
        Call UnitTesting.RunTest("Gathering protocol IDs", ServerPacketID.eGatheringFeedback = 222 And ClientPacketID.eGatheringStart = 318)
    #Else
        Call UnitTesting.RunTest("Gathering conditional protocol IDs", ServerPacketID.eGatheringFeedback = 223 And ClientPacketID.eGatheringStart = 321)
    #End If
    For index = 1 To 15
        Call RunCase(index)
    Next index
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
