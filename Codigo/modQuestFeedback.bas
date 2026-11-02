Attribute VB_Name = "modQuestFeedback"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Private Const MAX_PENDING_RECEIPTS As Integer = 16
Private Type FeedbackCache
    userId As Long
    version As Integer
    known As Boolean
    sequence As Long
    nextCycle As Long
    fingerprint As String
    quests(1 To 5) As Integer
    cycles(1 To 5) As Long
    receipts As Collection
End Type
Private mUsers() As FeedbackCache
Private mCapacity As Long

Private Sub EnsureFeedbackUser(ByVal userIndex As Integer)
    If userIndex > mCapacity Then
        mCapacity = UBound(UserList)
        ReDim Preserve mUsers(1 To mCapacity)
    End If
    With mUsers(userIndex)
        If .userId <> UserList(userIndex).Id Or .version <> UserList(userIndex).VersionId Then
            .userId = UserList(userIndex).Id: .version = UserList(userIndex).VersionId
            .known = False: .fingerprint = ""
            .sequence = 0: .nextCycle = 0
            Dim slot As Integer
            For slot = 1 To MAXUSERQUESTS
                .quests(slot) = 0: .cycles(slot) = 0
            Next slot
            Set .receipts = New Collection
        End If
        If .receipts Is Nothing Then Set .receipts = New Collection
    End With
End Sub

Public Sub BeginQuestFeedbackLifecycle(ByVal userIndex As Integer, ByVal slot As Byte)
    Call EnsureFeedbackUser(userIndex)
    Dim i As Integer
    For i = 1 To MAXUSERQUESTS
        If mUsers(userIndex).quests(i) = UserList(userIndex).QuestStats.Quests(slot).QuestIndex Then mUsers(userIndex).quests(i) = 0
    Next i
    Call SendQuestTurnInSnapshot(userIndex)
End Sub

Private Function Lifecycle(ByVal userIndex As Integer, ByVal slot As Integer, ByVal quest As Integer) As Long
    Dim i As Integer, available As Integer
    With mUsers(userIndex)
        For i = 1 To MAXUSERQUESTS
            If .quests(i) = quest Then
                Lifecycle = .cycles(i): Exit Function
            End If
            If .quests(i) = 0 Then available = i
        Next i
        If available = 0 Then available = slot
        .nextCycle = .nextCycle + 1
        .quests(available) = quest: .cycles(available) = .nextCycle
        Lifecycle = .nextCycle
    End With
End Function

Public Function CaptureQuestFeedbackLifecycle(ByVal userIndex As Integer, ByVal slot As Byte, ByVal quest As Integer) As Long
    Call EnsureFeedbackUser(userIndex)
    CaptureQuestFeedbackLifecycle = Lifecycle(userIndex, slot, quest)
End Function

Private Function GoalsFor(ByVal userIndex As Integer, ByVal slot As Integer, ByVal quest As Integer) As Collection
    Dim goals As New Collection, i As Integer, progress As Long, required As Long, id As Integer, kind As Integer
    With QuestList(quest)
        For kind = 1 To 3
            Dim count As Integer
            If kind = 1 Then count = .RequiredNPCs
            If kind = 2 Then count = .RequiredTargetNPCs
            If kind = 3 Then count = .RequiredOBJs
            For i = 1 To count
                If kind = 1 Then
                    id = .RequiredNPC(i).NpcIndex: required = .RequiredNPC(i).amount
                    progress = UserList(userIndex).QuestStats.Quests(slot).NPCsKilled(i)
                ElseIf kind = 2 Then
                    id = .RequiredTargetNPC(i).NpcIndex: required = .RequiredTargetNPC(i).amount
                    progress = UserList(userIndex).QuestStats.Quests(slot).NPCsTarget(i)
                Else
                    id = .RequiredOBJ(i).ObjIndex: required = .RequiredOBJ(i).amount
                    progress = CountInventoryObjects(userIndex, id)
                End If
                If progress > required Then progress = required
                If id > 0 And required > 0 And progress >= 0 Then Call goals.Add(Array(CByte(kind), i, id, progress, required))
            Next i
        Next kind
    End With
    Set GoalsFor = goals
End Function

Private Sub WriteFeedbackHeader(ByVal writer As Network.Writer, ByVal userIndex As Integer, ByVal completion As Boolean, ByVal silent As Boolean, ByVal count As Byte)
    Call writer.WriteInt16(ServerPacketID.eQuestFeedback)
    Call writer.WriteInt32(mUsers(userIndex).sequence + 1)
    Call writer.WriteInt8(Abs(CInt(completion)))
    Call writer.WriteInt8(Abs(CInt(silent)))
    Call writer.WriteInt8(count)
End Sub

' Called only by the existing dirty readiness drain, with its already evaluated
' ready collection. No second readiness calculation and no per-slot publication.
Public Function PublishQuestFeedback(ByVal userIndex As Integer, ByVal ready As Collection, ByVal force As Boolean) As Boolean
    On Error GoTo Failed
    Call EnsureFeedbackUser(userIndex)
    Dim slot As Integer, quest As Integer, count As Byte, cycle As Long, fingerprint As String
    Dim goals As Collection, goalsBySlot As New Collection, goal As Variant, q As Variant, isReady(1 To 5) As Boolean
    Dim cached As Integer, found As Boolean
    For cached = 1 To MAXUSERQUESTS
        found = False
        For slot = 1 To MAXUSERQUESTS
            If UserList(userIndex).QuestStats.Quests(slot).QuestIndex = mUsers(userIndex).quests(cached) Then found = True
        Next slot
        If Not found Then mUsers(userIndex).quests(cached) = 0
    Next cached
    For slot = 1 To MAXUSERQUESTS
        quest = UserList(userIndex).QuestStats.Quests(slot).QuestIndex
        If quest > 0 And quest <= UBound(QuestList) Then
            cycle = Lifecycle(userIndex, slot, quest)
            For Each q In ready
                If CInt(q) = quest Then isReady(slot) = True
            Next q
            Set goals = GoalsFor(userIndex, slot, quest)
            Call goalsBySlot.Add(goals, CStr(slot))
            fingerprint = fingerprint & quest & ":" & cycle & ":" & isReady(slot) & ":" & QuestList(quest).nombre & ";"
            For Each goal In goals
                fingerprint = fingerprint & goal(0) & ":" & goal(1) & ":" & goal(2) & ":" & goal(3) & ":" & goal(4) & ";"
            Next goal
            count = count + 1
        End If
    Next slot
    Dim writer As Network.Writer, queued As Boolean
    Set writer = Protocol_Writes.GetWriterBuffer()
    If force Or Not mUsers(userIndex).known Or fingerprint <> mUsers(userIndex).fingerprint Then
        Call WriteFeedbackHeader(writer, userIndex, False, force Or Not mUsers(userIndex).known, count)
        For slot = 1 To MAXUSERQUESTS
            quest = UserList(userIndex).QuestStats.Quests(slot).QuestIndex
            If quest > 0 And quest <= UBound(QuestList) Then
                Call writer.WriteInt16(quest)
                Call writer.WriteInt32(Lifecycle(userIndex, slot, quest))
                Call writer.WriteString8(Left$(QuestList(quest).nombre, 128))
                Call writer.WriteInt8(Abs(CInt(isReady(slot))))
                Set goals = goalsBySlot(CStr(slot))
                Call writer.WriteInt16(goals.Count)
                For Each goal In goals
                    Call writer.WriteInt8(goal(0))
                    Call writer.WriteInt16(goal(1))
                    Call writer.WriteInt16(goal(2))
                    Call writer.WriteInt32(goal(3))
                    Call writer.WriteInt32(goal(4))
                Next goal
            End If
        Next slot
        Call writer.WriteInt16(0)
        Call modSendData.SendData(ToIndex, userIndex, , False, queued, True)
        If Not queued Then Exit Function
        mUsers(userIndex).sequence = mUsers(userIndex).sequence + 1
        mUsers(userIndex).known = True: mUsers(userIndex).fingerprint = fingerprint
    End If
    ' Receipts are copied at the successful grant boundary; never reconstructed
    ' from static definitions or later inventory changes. Retry in order.
    Do While mUsers(userIndex).receipts.count > 0
        Dim receipt As Collection, rewards As Collection
        Set receipt = mUsers(userIndex).receipts(1)
        Set rewards = receipt(4)
        Call WriteFeedbackHeader(writer, userIndex, True, False, 1)
        Call writer.WriteInt16(receipt(1))
        Call writer.WriteInt32(receipt(2))
        Call writer.WriteString8(receipt(3))
        Call writer.WriteInt8(0)
        Call writer.WriteInt16(0)
        Call writer.WriteInt16(rewards.Count)
        For Each goal In rewards
            Call writer.WriteInt8(goal(0))
            Call writer.WriteInt16(goal(1))
            Call writer.WriteInt32(goal(2))
        Next goal
        queued = False
        Call modSendData.SendData(ToIndex, userIndex, , False, queued, True)
        If Not queued Then Exit Function
        mUsers(userIndex).sequence = mUsers(userIndex).sequence + 1
        Call mUsers(userIndex).receipts.Remove(1)
    Loop
    PublishQuestFeedback = True
    Exit Function
Failed:
    If Not writer Is Nothing Then Call writer.Clear()
    Call TraceError(Err.Number, Err.Description, "modQuestFeedback.PublishQuestFeedback", Erl)
End Function

Public Sub AddGrantedQuestReward(ByVal rewards As Collection, ByVal kind As Byte, ByVal id As Integer, ByVal amount As Long)
    If amount <= 0 Then Exit Sub
    If rewards.Count >= 512 Then Err.Raise 6, "Quest feedback", "Reward receipt limit exceeded"
    Call rewards.Add(Array(kind, id, amount))
End Sub

Public Sub QueueQuestCompletionFeedback(ByVal userIndex As Integer, ByVal cycle As Long, ByVal quest As Integer, ByVal rewards As Collection)
    On Error GoTo Failed
    Call EnsureFeedbackUser(userIndex)
    If mUsers(userIndex).receipts.Count >= MAX_PENDING_RECEIPTS Then
        Call TraceError(6, "Pending quest receipt limit reached", "modQuestFeedback.QueueQuestCompletionFeedback", 0)
        Exit Sub
    End If
    Dim receipt As New Collection
    Call receipt.Add(quest)
    Call receipt.Add(cycle)
    Call receipt.Add(Left$(QuestList(quest).nombre, 128))
    Call receipt.Add(rewards)
    Call mUsers(userIndex).receipts.Add(receipt)
    Call SendQuestTurnInSnapshot(userIndex)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modQuestFeedback.QueueQuestCompletionFeedback", Erl)
End Sub
