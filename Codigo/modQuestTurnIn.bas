Attribute VB_Name = "modQuestTurnIn"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Private Const MAX_RECIPIENTS_PER_QUEST As Integer = 128
Private Const MAX_RECIPIENTS_PER_PACKET As Integer = 384

Private mRecipients As Object
Private Type TurnInCache
    known As Boolean
    userId As Long
    version As Integer
    fingerprint As String
    globalGates As String
    dirtyId As Long
    dirtyVersion As Integer
    force As Boolean
    retry As Boolean
    readinessDirty As Boolean
    attempted As Long
End Type
Private mCache() As TurnInCache
Private mCapacity As Long

Private mDirty As Object
Private mGlobals As Object
Private mGlobalSent As Object
#If UNIT_TEST = 1 Then
Public EvaluationCount As Long
Public GlobalScanCount As Long
Public Sub EvaluateTurnInImmediatelyForTest(ByVal userIndex As Integer)
    Call EvaluateSnapshot(userIndex)
End Sub
Public Sub LegacyGlobalRefreshForTest(ByVal index As Integer)
    Dim u As Integer, slot As Integer, q As Integer
    GlobalScanCount = GlobalScanCount + 1
    For u = 1 To LastUser
        If UserList(u).flags.UserLogged Then
            For slot = 1 To MAXUSERQUESTS
                q = UserList(u).QuestStats.Quests(slot).QuestIndex
                If q > 0 And q <= UBound(QuestList) Then
                    If QuestList(q).GlobalQuestIndex = index Then
                        Call EvaluateSnapshot(u)
                        Exit For
                    End If
                End If
            Next slot
        End If
    Next u
End Sub
Public Sub ResetTurnInTestCounters()
    EvaluationCount = 0: GlobalScanCount = 0
    Set mDirty = Nothing: Set mGlobals = Nothing: Set mGlobalSent = Nothing
End Sub
#End If

Private Function ValidGlobal(ByVal index As Integer) As Boolean
    On Error GoTo Unavailable
    If index <= 0 Then Exit Function
    ValidGlobal = index >= LBound(GlobalQuestInfo) And index <= UBound(GlobalQuestInfo)
Unavailable:
End Function

Public Sub RefreshGlobalQuestTurnIns(ByVal globalQuest As Integer)
    If Not ValidGlobal(globalQuest) Then Exit Sub
    If mGlobals Is Nothing Then Set mGlobals = CreateObject("Scripting.Dictionary")
    mGlobals(CStr(globalQuest)) = True
End Sub

Private Function GlobalSignature(ByVal index As Integer) As String
    Dim alive As Boolean
    On Error Resume Next
    alive = IsAnyGlobalQuestBossAlive(index)
    Err.Clear
    GlobalSignature = CStr(GlobalQuestInfo(index).GatheringGlobalCounter) & ":" & CStr(GlobalQuestInfo(index).IsActive) & ":" & CStr(alive)
End Function

Private Function UserGlobalGates(ByVal userIndex As Integer) As String
    Dim slot As Integer, quest As Integer, globalQuest As Integer
    For slot = 1 To MAXUSERQUESTS
        quest = UserList(userIndex).QuestStats.Quests(slot).QuestIndex
        If quest > 0 And quest <= UBound(QuestList) Then
            globalQuest = QuestList(quest).GlobalQuestIndex
            If ValidGlobal(globalQuest) Then UserGlobalGates = UserGlobalGates & quest & ":" & CStr(FinishGlobalQuestCheck(userIndex, globalQuest, QuestList(quest).GlobalQuestThresholdNeeded, False)) & ";"
        End If
    Next slot
End Function

Private Sub CollectGlobalChanges()
    If mGlobals Is Nothing Then Exit Sub
    If mGlobalSent Is Nothing Then Set mGlobalSent = CreateObject("Scripting.Dictionary")
    Dim key As Variant, signature As String, index As Integer, userIndex As Long, slot As Integer, quest As Integer
    Dim discoveryChanged As Boolean
    For Each key In mGlobals.Keys
        index = CInt(key)
        If ValidGlobal(index) Then
            signature = GlobalSignature(index)
            If Not mGlobalSent.Exists(key) Then mGlobalSent(key) = ""
            If mGlobalSent(key) <> signature Then
                discoveryChanged = InvalidateQuestDiscoveryGlobal(index)
                #If UNIT_TEST = 1 Then
                    GlobalScanCount = GlobalScanCount + 1
                #End If
                For userIndex = 1 To LastUser
                    If UserList(userIndex).flags.UserLogged Then
                        If discoveryChanged Then Call SendQuestTurnInSnapshot(CInt(userIndex), False, True)
                        For slot = 1 To MAXUSERQUESTS
                            quest = UserList(userIndex).QuestStats.Quests(slot).QuestIndex
                            If quest > 0 And quest <= UBound(QuestList) Then
                                If QuestList(quest).GlobalQuestIndex = index Then
                                    If userIndex > mCapacity Then
                                        Call SendQuestTurnInSnapshot(CInt(userIndex))
                                    ElseIf Not mCache(userIndex).known Or mCache(userIndex).userId <> UserList(userIndex).Id Or mCache(userIndex).version <> UserList(userIndex).VersionId Or mCache(userIndex).globalGates <> UserGlobalGates(CInt(userIndex)) Then
                                        Call SendQuestTurnInSnapshot(CInt(userIndex))
                                    End If
                                    Exit For
                                End If
                            End If
                        Next slot
                    End If
                Next userIndex
                mGlobalSent(key) = signature
            End If
        End If
    Next key
    Call mGlobals.RemoveAll()
End Sub

' End-of-network-poll boundary: only queued identities, once per batch. Failed
' enqueue waits at least 250 ms; no retries within this drain.
Public Sub ProcessQuestTurnInUpdates(Optional ByVal nowRaw As Long = 0)
    On Error GoTo Failed
    Call CollectGlobalChanges()
    If mDirty Is Nothing Then Exit Sub
    If nowRaw = 0 Then nowRaw = GetTickCountRaw()
    Dim key As Variant, userIndex As Integer, processed As Integer, id As Long, version As Integer, succeeded As Boolean
    For Each key In mDirty.Keys
        If processed >= 128 Then Exit For
        userIndex = CInt(key)
        With mCache(userIndex)
            If Not UserList(userIndex).flags.UserLogged Or UserList(userIndex).Id <> .dirtyId Or UserList(userIndex).VersionId <> .dirtyVersion Then
                Call ResetQuestDiscoveryUser(userIndex)
                Call mDirty.Remove(key)
            ElseIf Not .retry Or TicksElapsed(.attempted, nowRaw) >= 250 Then
                processed = processed + 1: id = .dirtyId: version = .dirtyVersion
                If .readinessDirty Then
                    succeeded = EvaluateSnapshot(userIndex, .force)
                Else
                    succeeded = PublishQuestDiscovery(userIndex, .force)
                End If
                If succeeded Then
                    If .dirtyId = id And .dirtyVersion = version Then Call mDirty.Remove(key)
                Else
                    If .dirtyId = id And .dirtyVersion = version Then
                        .retry = True: .attempted = nowRaw
                        Call mDirty.Remove(key)
                        Call mDirty.Add(key, True)
                    End If
                End If
            End If
        End With
    Next key
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modQuestTurnIn.ProcessQuestTurnInUpdates", Erl)
End Sub

' Existing mutation call sites now request a presentation update, never send it
' inside an inventory/reward operation. Force also obeys the consistent boundary.
Public Sub SendQuestTurnInSnapshot(ByVal userIndex As Integer, Optional ByVal force As Boolean = False, Optional ByVal discoveryOnly As Boolean = False)
    On Error GoTo Failed
    If userIndex < 1 Or userIndex > UBound(UserList) Then Exit Sub
    If Not UserList(userIndex).flags.UserLogged Then Exit Sub
    If userIndex > mCapacity Then
        mCapacity = UBound(UserList): ReDim Preserve mCache(1 To mCapacity)
    End If
    If mDirty Is Nothing Then Set mDirty = CreateObject("Scripting.Dictionary")
    With mCache(userIndex)
        If .dirtyId <> UserList(userIndex).Id Or .dirtyVersion <> UserList(userIndex).VersionId Or Not mDirty.Exists(CStr(userIndex)) Then
            .dirtyId = UserList(userIndex).Id: .dirtyVersion = UserList(userIndex).VersionId
            .force = force: .retry = False
            .readinessDirty = Not discoveryOnly
        Else
            .force = .force Or force
            .readinessDirty = .readinessDirty Or Not discoveryOnly
        End If
    End With
    mDirty(CStr(userIndex)) = True
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modQuestTurnIn.SendQuestTurnInSnapshot", Erl)
End Sub

Public Sub ResetQuestTurnInRecipients()
    Set mRecipients = Nothing
    Call ResetQuestDiscoveryIndex()
End Sub

Private Function RecipientsFor(ByVal quest As Integer) As Collection
    On Error GoTo Unavailable
    If mRecipients Is Nothing Then Set mRecipients = CreateObject("Scripting.Dictionary")
    Dim key As String: key = CStr(quest)
    If Not mRecipients.Exists(key) Then
        Dim recipients As New Collection, npc As Long
        For npc = LBound(NpcInfoCache) To UBound(NpcInfoCache)
            If IsQuestTurnInRecipient(quest, CInt(npc)) Then Call recipients.Add(CInt(npc))
        Next npc
        Call mRecipients.Add(key, recipients)
    End If
    Set RecipientsFor = mRecipients(key)
    Exit Function
Unavailable:
    ' Startup/test worlds may not have authored template metadata yet.
    Set RecipientsFor = New Collection
End Function

' Full replacement of ready accepted quests. Empty replacement clears readiness.
' No selection subscription, database access or distant position disclosure.
Private Function EvaluateSnapshot(ByVal userIndex As Integer, Optional ByVal force As Boolean = False) As Boolean
    On Error GoTo Failed
    If userIndex < 1 Or userIndex > UBound(UserList) Then Exit Function
    If Not UserList(userIndex).flags.UserLogged Then Exit Function
    If userIndex > mCapacity Then
        mCapacity = UBound(UserList)
        ReDim Preserve mCache(1 To mCapacity)
    End If
    #If UNIT_TEST = 1 Then
        EvaluationCount = EvaluationCount + 1
    #End If
    If Not PublishQuestDiscovery(userIndex, force) Then Exit Function
    Dim ready As New Collection, slot As Integer, quest As Integer, fingerprint As String
    For slot = 1 To MAXUSERQUESTS
        quest = UserList(userIndex).QuestStats.Quests(slot).QuestIndex
        If quest > 0 And quest <= UBound(QuestList) Then
            If CanInspectQuestTurnIn(userIndex, quest, CByte(slot)) Then
                Call ready.Add(quest)
                Dim recipients As Collection, recipient As Variant
                Set recipients = RecipientsFor(quest)
                fingerprint = fingerprint & quest & ":" & QuestList(quest).nombre & ":"
                For Each recipient In recipients
                    fingerprint = fingerprint & recipient & ":" & GetNpcName(CInt(recipient)) & ";"
                Next recipient
            End If
        End If
    Next slot
    With mCache(userIndex)
        If Not force And .known And .userId = UserList(userIndex).Id And _
            .version = UserList(userIndex).VersionId And .fingerprint = fingerprint Then
            .globalGates = UserGlobalGates(userIndex)
            EvaluateSnapshot = PublishQuestFeedback(userIndex, ready, force)
            Exit Function
        End If
        Dim writer As Network.Writer, queued As Boolean, entry As Variant, recipientTotal As Long
        Set writer = Protocol_Writes.GetWriterBuffer()
        Call writer.WriteInt16(ServerPacketID.eQuestTurnIns)
        Call writer.WriteInt8(ready.Count)
        For Each entry In ready
            quest = CInt(entry)
            Call writer.WriteInt16(quest)
            Call writer.WriteString8(Left$(QuestList(quest).nombre, 128))
            Set recipients = RecipientsFor(quest)
            ' Overflow preserves readiness but does not guess a partial destination set.
            If recipients.Count > MAX_RECIPIENTS_PER_QUEST Or recipientTotal + recipients.Count > MAX_RECIPIENTS_PER_PACKET Then
                Call writer.WriteInt8(0)
            Else
                recipientTotal = recipientTotal + recipients.Count
                Call writer.WriteInt8(recipients.Count)
                For Each recipient In recipients
                    Call writer.WriteInt16(CInt(recipient))
                    Call writer.WriteString8(Left$(GetNpcName(CInt(recipient)), 128))
                Next recipient
            End If
        Next entry
        Call modSendData.SendData(ToIndex, userIndex, , False, queued, True)
        If queued Then
            EvaluateSnapshot = PublishQuestFeedback(userIndex, ready, force)
            .globalGates = UserGlobalGates(userIndex)
            .known = True: .userId = UserList(userIndex).Id
            .version = UserList(userIndex).VersionId: .fingerprint = fingerprint
        End If
    End With
    Exit Function
Failed:
    If Not writer Is Nothing Then Call writer.Clear()
    Call TraceError(Err.Number, Err.Description, "modQuestTurnIn.SendQuestTurnInSnapshot", Erl)
End Function

' Recipient identity comes from the same template metadata used by QuestTalk.
' This function never reads live NPC positions or performs I/O.
Public Function IsQuestTurnInRecipient(ByVal quest As Integer, ByVal npc As Integer) As Boolean
    On Error GoTo Failed
    If quest <= 0 Or quest > UBound(QuestList) Then Exit Function
    If npc < LBound(NpcInfoCache) Or npc > UBound(NpcInfoCache) Then Exit Function
    If Not NpcInfoCache(npc).Exists Then Exit Function
    If QuestList(quest).TalkTo = npc Then
        IsQuestTurnInRecipient = True
        Exit Function
    End If
    Dim offer As Integer
    For offer = 1 To NpcInfoCache(npc).NumQuest
        If NpcInfoCache(npc).QuestNumber(offer) = quest Then
            IsQuestTurnInRecipient = True
            Exit Function
        End If
    Next offer
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "modQuestTurnIn.IsQuestTurnInRecipient", Erl)
End Function
