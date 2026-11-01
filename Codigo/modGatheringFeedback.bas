Attribute VB_Name = "modGatheringFeedback"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit

Public Enum e_GatheringOutcome
    GatheringStarted = 0
    GatheringReward = 1
    GatheringFailedRoll = 2
    GatheringBlocked = 3
    GatheringStopped = 4
    GatheringGroundReward = 5
End Enum
Public Enum e_GatheringReason
    GatheringNoReason = 0
    GatheringStamina = 1
    GatheringInventory = 2
    GatheringTarget = 3
    GatheringTool = 4
    GatheringRange = 5
    GatheringRestriction = 6
    GatheringDepleted = 7
    GatheringSpecialFish = 8
    GatheringNoDropSpace = 9
End Enum
Private Type GatheringContext
    request As Long
    sequence As Long
    userId As Long
    userVersion As Integer
    map As Integer
    skill As Byte
    stopped As Boolean
End Type
Private contexts() As GatheringContext
Private specialCatches() As GatheringContext
Private capacity As Long

Public Function HasGatheringFeedback(ByVal userIndex As Integer) As Boolean
    If userIndex <= 0 Or userIndex > capacity Then Exit Function
    With contexts(userIndex)
        HasGatheringFeedback = .request > 0 And .userId = UserList(userIndex).Id And .userVersion = UserList(userIndex).VersionId
    End With
End Function

Public Sub BeginGatheringFeedback(ByVal userIndex As Integer, ByVal request As Long, ByVal skill As e_Skill)
    On Error GoTo Failed
    If capacity <> UBound(UserList) Then
        capacity = UBound(UserList)
        ReDim contexts(1 To capacity)
        ReDim specialCatches(1 To capacity)
    End If
    With contexts(userIndex)
        .request = request
        .sequence = 0
        .userId = UserList(userIndex).Id
        .userVersion = UserList(userIndex).VersionId
        .map = UserList(userIndex).pos.Map
        .skill = skill
        .stopped = False
    End With
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modGatheringFeedback.Begin", Erl)
End Sub

' Called only at actual validation/award/roll branches, never from inventory sync.
' Retain the stopped context for the special-fish minigame's eventual award.
Public Sub SendGatheringFeedback(ByVal userIndex As Integer, ByVal outcome As e_GatheringOutcome, Optional ByVal reason As e_GatheringReason = GatheringNoReason, Optional ByVal objectIndex As Integer = 0, Optional ByVal amount As Long = 0, Optional ByVal keepActive As Boolean = False)
    On Error GoTo Failed
    If Not HasGatheringFeedback(userIndex) Then Exit Sub
    With contexts(userIndex)
        If outcome = GatheringStopped And .stopped Then Exit Sub
        If outcome = GatheringReward Or outcome = GatheringGroundReward Then
            If objectIndex <= 0 Or amount <= 0 Then Exit Sub
        End If
        If .sequence = &H7FFFFFFF Then Exit Sub
        .sequence = .sequence + 1
        Dim active As Boolean
        active = UserList(userIndex).AutomatedAction.IsActive And Not .stopped
        If (outcome = GatheringBlocked And Not keepActive) Or outcome = GatheringStopped Then active = False
        If Not active Then .stopped = True
        Dim writer As Network.Writer
        Set writer = Protocol_Writes.GetWriterBuffer()
        Call writer.WriteInt16(ServerPacketID.eGatheringFeedback)
        Call writer.WriteInt32(.request)
        Call writer.WriteInt32(.sequence)
        Call writer.WriteInt16(.map)
        Call writer.WriteInt8(.skill)
        Call writer.WriteInt8(outcome)
        Call writer.WriteInt8(reason)
        Call writer.WriteInt8(Abs(CInt(active)))
        Call writer.WriteInt16(objectIndex)
        Call writer.WriteInt32(amount)
        Call writer.WriteInt16(UserList(userIndex).invent.EquippedWorkingToolObjIndex)
        Call modSendData.SendData(ToIndex, userIndex)
    End With
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modGatheringFeedback.Send", Erl)
End Sub

Public Sub RejectGatheringRequest(ByVal userIndex As Integer, ByVal request As Long, ByVal skill As e_Skill, ByVal reason As e_GatheringReason)
    ' Temporarily swap feedback context; rejecting a new request must not change
    ' the existing automatic action or its correlation.
    On Error GoTo Failed
    Dim previous As GatheringContext
    If userIndex <= capacity Then previous = contexts(userIndex)
    Call BeginGatheringFeedback(userIndex, request, skill)
    Call SendGatheringFeedback(userIndex, GatheringBlocked, reason)
    contexts(userIndex) = previous
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modGatheringFeedback.Reject", Erl)
End Sub

Public Sub RememberSpecialGatheringCatch(ByVal userIndex As Integer)
    On Error GoTo Failed
    If HasGatheringFeedback(userIndex) Then specialCatches(userIndex) = contexts(userIndex)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modGatheringFeedback.RememberCatch", Erl)
End Sub

Public Sub CompleteSpecialGatheringCatch(ByVal userIndex As Integer, ByVal objectIndex As Integer, ByVal awarded As Boolean)
    On Error GoTo Failed
    If userIndex <= 0 Or userIndex > capacity Then Exit Sub
    Dim current As GatheringContext
    current = contexts(userIndex)
    contexts(userIndex) = specialCatches(userIndex)
    If awarded Then
        Call SendGatheringFeedback(userIndex, GatheringReward, GatheringNoReason, objectIndex, 1)
    Else
        Call SendGatheringFeedback(userIndex, GatheringBlocked, GatheringInventory)
    End If
    If current.request <> contexts(userIndex).request Then contexts(userIndex) = current
    specialCatches(userIndex).request = 0
    Exit Sub
Failed:
    contexts(userIndex) = current
    Call TraceError(Err.Number, Err.Description, "modGatheringFeedback.SpecialCatch", Erl)
End Sub

Public Sub ReportGatheringDrop(ByVal userIndex As Integer, ByRef obj As t_Obj, ByRef destination As t_WorldPos)
    On Error GoTo Failed
    If destination.x > 0 And destination.y > 0 Then
        Call SendGatheringFeedback(userIndex, GatheringGroundReward, GatheringNoReason, obj.ObjIndex, obj.amount)
    Else
        ' Preserve the existing retry loop even when floor placement fails.
        Call SendGatheringFeedback(userIndex, GatheringBlocked, GatheringNoDropSpace, 0, 0, True)
    End If
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modGatheringFeedback.ReportDrop", Erl)
End Sub

Private Function GatheredInventoryTotal(ByVal userIndex As Integer, ByRef obj As t_Obj) As Long
    On Error GoTo Failed
    If obj.ObjIndex = 12 Then
        GatheredInventoryTotal = UserList(userIndex).Stats.GLD
        Exit Function
    End If
    Dim slot As Integer
    For slot = 1 To UserList(userIndex).CurrentInventorySlots
        With UserList(userIndex).invent.Object(slot)
            If .ObjIndex = obj.ObjIndex And .ElementalTags = obj.ElementalTags Then
                GatheredInventoryTotal = GatheredInventoryTotal + .amount
            End If
        End With
    Next slot
    Exit Function
Failed:
    GatheredInventoryTotal = -1
    Call TraceError(Err.Number, Err.Description, "modGatheringFeedback.InventoryTotal", Erl)
End Function

' This synchronous observation wraps only the existing gathering award call.
' It is not an inventory-update observer. Preserve MeterItemEnInventario's
' return value and mutations, including existing stack-cap behavior.
Public Function InsertGatheredItem(ByVal userIndex As Integer, ByRef obj As t_Obj) As Boolean
    On Error GoTo Failed
    Dim before As Long, after As Long, reason As e_GatheringReason
    If Not HasGatheringFeedback(userIndex) Then
        InsertGatheredItem = MeterItemEnInventario(userIndex, obj)
        Exit Function
    End If
    before = GatheredInventoryTotal(userIndex, obj)
    InsertGatheredItem = MeterItemEnInventario(userIndex, obj)
    If Not InsertGatheredItem Or before < 0 Then Exit Function
    after = GatheredInventoryTotal(userIndex, obj)
    If after <= before Then Exit Function
    If Not UserList(userIndex).AutomatedAction.IsActive Then reason = GatheringDepleted
    Call SendGatheringFeedback(userIndex, GatheringReward, reason, obj.ObjIndex, after - before)
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "modGatheringFeedback.Insert", Erl)
End Function
