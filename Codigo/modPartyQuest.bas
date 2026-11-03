Attribute VB_Name = "modPartyQuest"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Private Const MAX_MEMBERS As Integer = 6
Private Const MAX_KILLS As Integer = 6
Private Const QUESTION_PARTY_QUEST As Integer = 64
Private Const PROPOSAL_LIFETIME As Long = 90000
Public Enum e_PartyQuestAction
    HuntRefresh = 0
    HuntPropose = 1
    HuntAccept = 2
    HuntDecline = 3
    HuntStart = 4
    HuntAbandon = 5
End Enum
Public Enum e_PartyQuestResult
    HuntState = 0
    HuntProposed = 1
    HuntNeedParty = 2
    HuntNotLeader = 3
    HuntUnavailable = 4
    HuntAlreadyActive = 5
    HuntStale = 6
    HuntBusy = 7
    HuntStarted = 8
    HuntAbandoned = 9
    HuntClaimed = 10
    HuntStorageError = 11
    HuntExpired = 12
    HuntEnrollmentClosed = 13
End Enum
Private Type HuntMember
    characterId As Long
    name As String
    online As t_UserReference
    vote As Byte ' 0 unanswered, 1 accepted, 2 declined, 3 no longer eligible
    participated As Boolean
    status As Byte ' 0 enrolled, 1 claimed, 2 abandoned
    savedStatus As Byte
End Type
Private Type HuntRun
    revision As Long
    savedRevision As Long
    killMasks(1 To MAX_KILLS) As Long
    id As Long
    quest As Integer
    leaderId As Long
    target As Integer
    required As Integer
    progress As Integer
    reward As Long
    count As Byte
    members(1 To MAX_MEMBERS) As HuntMember
End Type
Private Type HuntProposal
    id As Long
    quest As Integer
    party As Long
    leader As t_UserReference
    created As Long
    count As Byte
    members(1 To MAX_MEMBERS) As HuntMember
End Type
Private Type HuntSession
    characterId As Long
    version As Integer
    runSlot As Integer
    proposal As Integer
    request As Long
    requestKey As String
    result As e_PartyQuestResult
    revision As Long
    dirty As Boolean
    retryAt As Long
End Type
Private mRuns() As HuntRun
Private mProposals() As HuntProposal
Private mSessions() As HuntSession
Private mOnline As Object
Private mCapacity As Integer
Private mNextProposal As Long
Private mDefinition As Integer
Private mRetries As Object
Private mLeaders As Object
Private mNextRun As Long
Private mRunIdsKnown As Boolean

Private Type HuntSaveOwner
    characterId As Long
    nextSequence As Long
    committedSequence As Long
    latest As clsHuntCharacterSave
End Type
Private Type HuntSaveBatch
    characterId As Long
    sequence As Long
    count As Integer
    runs() As HuntRun
    characters As Collection
End Type
Private mSaveOwners As Object
Private mPendingSaves As Object
Private mOwners() As HuntSaveOwner
Private mBatches() As HuntSaveBatch
Private mSaveCapacity As Integer
Private mFlushing As Boolean
#If UNIT_TEST = 1 Then
Public PartyQuestSqlCount As Long
Public FailPartyRosterRead As Boolean
Public HoldPartySaves As Boolean
Public LastPartySaveFailureForTest As String
Private mHeldBatch As HuntSaveBatch
#End If

' Load/save-boundary SQL propagates errors. Gameplay mutations never call this
' helper; they dirty the shared owner for the next character save.
Private Function HuntSql(ByVal sql As String, ParamArray values() As Variant) As ADODB.Recordset
    Dim command As New ADODB.Command, value As Variant, parameter As ADODB.Parameter
    Set command.ActiveConnection = Connection
    command.CommandType = adCmdText: command.CommandText = sql
    For Each value In values
        If VarType(value) = vbString Then
            Set parameter = command.CreateParameter(, adVarWChar, adParamInput, Len(value) + 1, value)
        Else
            Set parameter = command.CreateParameter(, adInteger, adParamInput, , CLng(value))
        End If
        Call command.Parameters.Append(parameter)
    Next value
    Call CountPartyQuestSql()
    Set HuntSql = command.Execute()
End Function

Private Sub CloseHuntRows(ByRef rows As ADODB.Recordset)
    If rows Is Nothing Then Exit Sub
    If rows.State <> adStateClosed Then Call rows.Close()
    Set rows = Nothing
End Sub

Private Sub EnsureHuntSession(ByVal userIndex As Integer)
    If mCapacity < UBound(UserList) Then
        mCapacity = UBound(UserList)
        ReDim Preserve mSessions(1 To mCapacity)
        ReDim Preserve mRuns(1 To mCapacity)
        ReDim Preserve mProposals(1 To mCapacity)
    End If
    If mOnline Is Nothing Then Set mOnline = CreateObject("Scripting.Dictionary")
    If mRetries Is Nothing Then Set mRetries = CreateObject("Scripting.Dictionary")
    If mLeaders Is Nothing Then Set mLeaders = CreateObject("Scripting.Dictionary")
    With mSessions(userIndex)
        If .characterId <> UserList(userIndex).Id Or .version <> UserList(userIndex).VersionId Then
            If mOnline.Exists(CStr(.characterId)) Then
                If mOnline(CStr(.characterId)) = userIndex Then Call mOnline.Remove(CStr(.characterId))
            End If
            If mRetries.Exists(CStr(userIndex)) Then Call mRetries.Remove(CStr(userIndex))
            .characterId = UserList(userIndex).Id: .version = UserList(userIndex).VersionId
            .runSlot = 0: .proposal = 0: .request = 0: .revision = 0: .dirty = False
            .requestKey = vbNullString: .result = HuntState: .retryAt = 0
        End If
    End With
    mOnline(CStr(UserList(userIndex).Id)) = userIndex
End Sub

Private Function OnlineCharacter(ByVal characterId As Long) As Integer
    If mOnline Is Nothing Then Exit Function
    If Not mOnline.Exists(CStr(characterId)) Then Exit Function
    Dim userIndex As Integer
    userIndex = mOnline(CStr(characterId))
    If userIndex < 1 Or userIndex > UBound(UserList) Then Exit Function
    With mSessions(userIndex)
        If .characterId <> characterId Or .version <> UserList(userIndex).VersionId Then Exit Function
    End With
    If UserList(userIndex).Id <> characterId Or Not IsRewardCredit(userIndex) Then Exit Function
    OnlineCharacter = userIndex
End Function

Private Function RunMember(ByVal runSlot As Integer, ByVal characterId As Long) As Integer
    Dim i As Integer
    For i = 1 To mRuns(runSlot).count
        If mRuns(runSlot).members(i).characterId = characterId Then RunMember = i: Exit Function
    Next i
End Function

Private Function QuestSlotFor(ByVal userIndex As Integer, ByVal quest As Integer) As Byte
    Dim i As Byte
    For i = 1 To MAXUSERQUESTS
        If UserList(userIndex).QuestStats.Quests(i).QuestIndex = quest Then QuestSlotFor = i: Exit Function
    Next i
End Function

Private Function HuntDefinition() As Integer
    ' Deliberately one authored party quest in the first delivery, no ID/name heuristic.
    If mDefinition > 0 Then HuntDefinition = mDefinition: Exit Function
    Dim q As Integer
    For q = 1 To UBound(QuestList)
        If QuestList(q).PartyMinimum > 0 Then
            If QuestList(q).PartyMinimum = 2 And QuestList(q).RequiredNPCs = 1 And _
               QuestList(q).RequiredOBJs = 0 And QuestList(q).RewardOBJs = 0 And _
               QuestList(q).RewardEXP = 0 And QuestList(q).RewardSpellCount = 0 And QuestList(q).RequiredNPC(1).amount = MAX_KILLS Then
                HuntDefinition = q: mDefinition = q
            End If
            Exit Function
        End If
    Next q
End Function

Private Function CanEnroll(ByVal userIndex As Integer, ByVal quest As Integer) As Boolean
    If Not IsRewardUser(userIndex) Then Exit Function
    Call EnsureHuntSession(userIndex)
    If mSessions(userIndex).runSlot <> 0 Then Exit Function
    If FreeQuestSlot(userIndex) = 0 Then Exit Function
    CanEnroll = QuestAcceptanceFailure(userIndex, 0, quest, QuestList(quest)) = QuestAcceptAllowed
End Function

Private Function CanProposeHunt(ByVal userIndex As Integer, ByVal quest As Integer) As Boolean
    If PartyLeaderIndex(userIndex) <> userIndex Then Exit Function
    If Not CanEnroll(userIndex, quest) Then Exit Function
    Dim recipients As t_PartyRewardRecipients, i As Integer, count As Integer
    Call CollectPartyRewardRecipients(userIndex, recipients)
    For i = 1 To recipients.count
        If IsValidUserRef(recipients.Users(i)) Then
            If CanEnroll(recipients.Users(i).ArrayIndex, quest) Then count = count + 1
        End If
    Next i
    CanProposeHunt = count >= 2
End Function

Private Function MemberEligible(ByVal leader As Integer, ByVal userIndex As Integer, ByVal quest As Integer) As Boolean
    If Not CanEnroll(userIndex, quest) Then Exit Function
    If PartyLeaderIndex(userIndex) <> leader Then Exit Function
    If UserList(userIndex).pos.Map <> UserList(leader).pos.Map Then Exit Function
    MemberEligible = Distancia(UserList(leader).pos, UserList(userIndex).pos) < 20
End Function

Private Sub ProjectRun(ByVal runSlot As Integer, ByVal userIndex As Integer, ByVal newAcceptance As Boolean)
    Dim member As Integer, slot As Byte, quest As Integer
    member = RunMember(runSlot, UserList(userIndex).Id)
    If member = 0 Then Exit Sub
    If mRuns(runSlot).members(member).status <> 0 Then Exit Sub
    quest = mRuns(runSlot).quest: slot = QuestSlotFor(userIndex, quest)
    If slot = 0 Then slot = FreeQuestSlot(userIndex)
    If slot = 0 Then Err.Raise vbObjectError + 860, , "No slot for durable party quest"
    With UserList(userIndex).QuestStats.Quests(slot)
        .QuestIndex = quest: ReDim .NPCsKilled(1 To 1)
        .NPCsKilled(1) = mRuns(runSlot).progress: .Dirty = True
    End With
    UserList(userIndex).flags.ModificoQuests = True
    mSessions(userIndex).runSlot = runSlot
    Call SetUserRef(mRuns(runSlot).members(member).online, userIndex)
    If newAcceptance Then
        Call BeginQuestFeedbackLifecycle(userIndex, slot)
        Call WriteLocaleMsg(userIndex, MSG_ACEPTADO_MISION, e_TextChannel.TEXTCHANNEL_QUEST, e_FontTypeNames.FONTTYPE_INFOBOLD, quest)
    End If
End Sub

Private Function FreeRunSlot() As Integer
    Dim i As Integer, j As Integer, online As Boolean, userIndex As Integer
    For i = 1 To mCapacity
        online = False
        For j = 1 To mRuns(i).count
            userIndex = OnlineCharacter(mRuns(i).members(j).characterId)
            If userIndex > 0 And mRuns(i).members(j).status = 0 Then
                If mSessions(userIndex).runSlot = i Then online = True
            End If
        Next j
        If mRuns(i).id = 0 Then FreeRunSlot = i: Exit Function
        If Not online And mRuns(i).revision = mRuns(i).savedRevision Then FreeRunSlot = i: Exit Function
    Next i
End Function

Public Function InitializePartyQuestPersistence() As Boolean
    On Error GoTo Failed
    If mRunIdsKnown Then InitializePartyQuestPersistence = True: Exit Function
    Dim rows As ADODB.Recordset
    Set rows = HuntSql("SELECT COALESCE(MAX(id),0) FROM party_quest_run")
    mNextRun = rows.Fields(0).value
    Call CloseHuntRows(rows)
    mRunIdsKnown = True: InitializePartyQuestPersistence = True
    Exit Function
Failed:
    Call CloseHuntRows(rows)
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.InitializePartyQuestPersistence", Erl)
End Function

Private Sub ReadCompleteRun(ByVal runId As Long, ByRef loaded As HuntRun)
    On Error GoTo Failed
    Dim rows As ADODB.Recordset, i As Integer, j As Integer, leaderFound As Boolean, mask As Long
    Set rows = HuntSql("SELECT * FROM party_quest_run WHERE id=?", runId)
    If rows.EOF Then Err.Raise vbObjectError + 862, , "Missing party run"
    With loaded
        .id = runId: .quest = rows!quest_id: .leaderId = rows!leader_id
        .target = rows!target_npc: .required = rows!required: .progress = rows!progress
        .reward = rows!reward_gold: .revision = rows!revision: .savedRevision = .revision
        If .quest < 1 Or .quest > UBound(QuestList) Then Err.Raise vbObjectError + 862, , "Invalid hunt definition"
        If .required <> MAX_KILLS Or .progress < 0 Or .progress > .required Or .revision < 1 Or .target < 1 Or .reward < 1 Or .reward > MAXORO Then Err.Raise vbObjectError + 862, , "Invalid hunt state"
    End With
    Call CloseHuntRows(rows)
    #If UNIT_TEST = 1 Then
        If FailPartyRosterRead Then
            FailPartyRosterRead = False
            Err.Raise vbObjectError + 862, , "Injected roster read failure"
        End If
    #End If
    Set rows = HuntSql("SELECT * FROM party_quest_member WHERE run_id=? ORDER BY ordinal", runId)
    Do While Not rows.EOF
        loaded.count = loaded.count + 1: i = loaded.count
        If i > MAX_MEMBERS Then Err.Raise vbObjectError + 862, , "Hunt roster overflow"
        If rows!ordinal <> i Then Err.Raise vbObjectError + 862, , "Non-contiguous hunt ordinals"
        With loaded.members(i)
            .characterId = rows!character_id: .name = rows!name
            .participated = rows!participated <> 0: .status = rows!status: .savedStatus = .status: .vote = 1
            If .characterId <= 0 Or Len(.name) = 0 Or .status > 2 Then Err.Raise vbObjectError + 862, , "Invalid hunt member"
            If .characterId = loaded.leaderId Then leaderFound = True
            For j = 1 To i - 1
                If loaded.members(j).characterId = .characterId Then Err.Raise vbObjectError + 862, , "Duplicate hunt identity"
            Next j
            ' Retain tombstone identity and ordinal after user deletion.
            If IsNull(rows!user_id) And .status = 0 Then
                .status = 2: loaded.revision = loaded.revision + 1
            End If
        End With
        Call rows.MoveNext()
    Loop
    Call CloseHuntRows(rows)
    If loaded.count < 2 Or Not leaderFound Then Err.Raise vbObjectError + 862, , "Incomplete hunt roster"
    Set rows = HuntSql("SELECT ordinal,qualified_mask FROM party_quest_kill WHERE run_id=? ORDER BY ordinal", runId)
    i = 0
    Do While Not rows.EOF
        i = i + 1
        If i > loaded.progress Or rows!ordinal <> i Then Err.Raise vbObjectError + 862, , "Invalid hunt kill history"
        mask = rows!qualified_mask
        If mask < 1 Or mask >= 2 ^ loaded.count Then Err.Raise vbObjectError + 862, , "Invalid hunt participation mask"
        loaded.killMasks(i) = mask
        Call rows.MoveNext()
    Loop
    Call CloseHuntRows(rows)
    If i <> loaded.progress Then Err.Raise vbObjectError + 862, , "Incomplete hunt kill history"
    Exit Sub
Failed:
    Dim number As Long, description As String
    number = Err.Number: description = Err.Description
    Resume Cleanup
Cleanup:
    Call CloseHuntRows(rows)
    On Error GoTo 0
    Err.Raise number, "modPartyQuest.ReadCompleteRun", description
End Sub

Public Function LoadPartyQuest(ByVal userIndex As Integer) As Boolean
    On Error GoTo Failed
    Call EnsureHuntSession(userIndex)
    If Not InitializePartyQuestPersistence() Then Exit Function
    Dim rows As ADODB.Recordset, runId As Long, cache As Integer, i As Integer, j As Integer, slot As Byte, quest As Integer
    Dim loaded As HuntRun, history As Boolean
    ' A retained owner is newer than storage after a failed save / while a teammate plays.
    For i = 1 To mCapacity
        If mRuns(i).id > 0 Then
            j = RunMember(i, UserList(userIndex).Id)
            If j > 0 Then
                history = True
                If mRuns(i).members(j).status = 0 Then cache = i: Exit For
            End If
        End If
    Next i
    Set rows = HuntSql("SELECT r.quest_id,m.status,m.run_id FROM party_quest_run r JOIN party_quest_member m ON m.run_id=r.id WHERE m.character_id=?", UserList(userIndex).Id)
    Dim done As New Collection
    Do While Not rows.EOF
        history = True
        If rows!status = 1 Then Call done.Add(CInt(rows!quest_id))
        If rows!status = 0 Then runId = rows!run_id
        Call rows.MoveNext()
    Loop
    Call CloseHuntRows(rows)
    If cache = 0 And runId > 0 Then
        For i = 1 To mCapacity
            If mRuns(i).id = runId Then
                j = RunMember(i, UserList(userIndex).Id)
                If j > 0 Then
                    If mRuns(i).members(j).status = 0 Then cache = i
                    runId = 0: Exit For
                End If
            End If
        Next i
        If cache = 0 And runId > 0 Then
            ' Publish nothing until every read and validation has succeeded.
            Call ReadCompleteRun(runId, loaded)
            cache = FreeRunSlot(): If cache = 0 Then Err.Raise vbObjectError + 862, , "Hunt cache capacity reached"
        End If
    End If
    If history Or UserList(userIndex).PartyQuestSaveRevision > 0 Then
        If Not RegisterHuntSaveOwner(userIndex) Then Err.Raise vbObjectError + 862, , "Hunt save capacity reached"
    End If
    If loaded.id > 0 Then mRuns(cache) = loaded
    mSessions(userIndex).runSlot = 0
    For slot = 1 To MAXUSERQUESTS
        quest = UserList(userIndex).QuestStats.Quests(slot).QuestIndex
        If quest > 0 Then
            If QuestList(quest).PartyMinimum > 0 Then
                UserList(userIndex).QuestStats.Quests(slot).QuestIndex = 0
                UserList(userIndex).QuestStats.Quests(slot).Dirty = True
                UserList(userIndex).flags.ModificoQuests = True
            End If
        End If
    Next slot
    Dim completed As Variant
    For Each completed In done
        If Not UserDoneQuest(userIndex, CInt(completed)) Then Call AddDoneQuest(userIndex, CInt(completed))
    Next completed
    If cache > 0 Then Call ProjectRun(cache, userIndex, False)
    LoadPartyQuest = True
    Exit Function
Failed:
    Call CloseHuntRows(rows)
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.LoadPartyQuest", Erl)
End Function

Private Sub PublishRun(ByVal cache As Integer)
    Dim i As Integer, userIndex As Integer
    For i = 1 To mRuns(cache).count
        userIndex = OnlineCharacter(mRuns(cache).members(i).characterId)
        If userIndex > 0 Then
            If mRuns(cache).members(i).status = 0 Then
                Call ProjectRun(cache, userIndex, False)
                Call RefreshNpcInspection(userIndex)
                Call SendQuestObjectiveSnapshot(userIndex)
                Call WritePartyQuestState(userIndex)
            End If
        End If
    Next i
End Sub

Private Sub PublishProposal(ByVal leader As Integer)
    Dim i As Integer, u As Integer
    For i = 1 To mProposals(leader).count
        If IsValidUserRef(mProposals(leader).members(i).online) Then
            u = mProposals(leader).members(i).online.ArrayIndex
            Call WritePartyQuestState(u)
        End If
    Next i
End Sub

Public Sub RefreshPartyQuestProposal(ByVal leader As Integer)
    On Error GoTo Failed
    If leader > mCapacity Or leader < 1 Then Exit Sub
    If mProposals(leader).id = 0 Then Exit Sub
    If Not IsValidUserRef(mProposals(leader).leader) Then
        Call CloseProposal(leader, HuntExpired): Exit Sub
    End If
    If PartyLeaderIndex(leader) <> leader Or UserList(leader).Grupo.Id <> mProposals(leader).party Then
        Call CloseProposal(leader, HuntExpired): Exit Sub
    End If
    Dim i As Integer, userIndex As Integer
    For i = 1 To mProposals(leader).count
        With mProposals(leader).members(i)
            If IsValidUserRef(.online) Then
                userIndex = .online.ArrayIndex
                If Not MemberEligible(leader, userIndex, mProposals(leader).quest) Then
                    .vote = 3
                    If mSessions(userIndex).proposal = leader Then
                        mSessions(userIndex).proposal = 0
                        If UserList(userIndex).flags.pregunta = QUESTION_PARTY_QUEST Then
                            UserList(userIndex).flags.pregunta = 0: UserList(userIndex).flags.RespondiendoPregunta = False
                        End If
                        Call WritePartyQuestState(userIndex)
                    End If
                End If
            Else
                .vote = 3
            End If
        End With
    Next i
    Call PublishProposal(leader)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.RefreshPartyQuestProposal", Erl)
End Sub

Private Sub CloseProposal(ByVal leader As Integer, ByVal result As e_PartyQuestResult, Optional ByVal startedRun As Integer = 0)
    Dim i As Integer, u As Integer
    For i = 1 To mProposals(leader).count
        With mProposals(leader).members(i)
            If IsValidUserRef(.online) Then
                u = .online.ArrayIndex
                If mSessions(u).proposal = leader Then
                mSessions(u).proposal = 0
                If UserList(u).flags.pregunta = QUESTION_PARTY_QUEST Then
                    UserList(u).flags.pregunta = 0: UserList(u).flags.RespondiendoPregunta = False
                End If
                mSessions(u).result = result
                If result = HuntStarted Then
                    mSessions(u).result = HuntEnrollmentClosed
                    If startedRun > 0 Then
                        If RunMember(startedRun, UserList(u).Id) > 0 Then mSessions(u).result = HuntStarted
                    End If
                End If
                mSessions(u).dirty = True
                mRetries(CStr(u)) = True
                End If
            End If
        End With
    Next i
    mProposals(leader).id = 0
    If mLeaders.Exists(CStr(leader)) Then Call mLeaders.Remove(CStr(leader))
End Sub

Public Sub ProposePartyQuest(ByVal userIndex As Integer, ByVal quest As Integer, ByVal npcIndex As Integer)
    On Error GoTo Failed
    Call EnsureHuntSession(userIndex)
    Dim leader As Integer, recipients As t_PartyRewardRecipients, i As Integer, j As Integer, u As Integer
    mSessions(userIndex).result = HuntUnavailable
    If quest <> HuntDefinition() Or npcIndex < 1 Then GoTo Reply
    If Not IsValidNpcRef(UserList(userIndex).flags.TargetNPC) Then GoTo Reply
    If UserList(userIndex).flags.TargetNPC.ArrayIndex <> npcIndex Then GoTo Reply
    If NpcList(npcIndex).Numero <> QuestList(quest).TalkTo Then GoTo Reply
    Dim offered As Boolean
    For i = 1 To NpcList(npcIndex).NumQuest
        If NpcList(npcIndex).QuestNumber(i) = quest Then offered = True
    Next i
    If Not offered Then GoTo Reply
    If QuestAcceptanceFailure(userIndex, npcIndex, quest, QuestList(quest)) <> QuestAcceptAllowed Then GoTo Reply
    leader = PartyLeaderIndex(userIndex)
    If leader = 0 Then mSessions(userIndex).result = HuntNeedParty: GoTo Reply
    If leader <> userIndex Then mSessions(userIndex).result = HuntNotLeader: GoTo Reply
    If mSessions(userIndex).runSlot > 0 Then mSessions(userIndex).result = HuntAlreadyActive: GoTo Reply
    If Not CanEnroll(userIndex, quest) Then GoTo Reply
    If UserList(userIndex).flags.RespondiendoPregunta Then mSessions(userIndex).result = HuntBusy: GoTo Reply
    If mProposals(leader).id > 0 Then Call CloseProposal(leader, HuntExpired)
    Call CollectPartyRewardRecipients(userIndex, recipients)
    If Not recipients.IsParty Or recipients.count < 2 Then mSessions(userIndex).result = HuntNeedParty: GoTo Reply
    Dim blankState As HuntProposal, candidate As HuntMember
    mProposals(leader) = blankState
    With mProposals(leader)
        mNextProposal = mNextProposal + 1: .id = mNextProposal: .quest = quest
        .party = recipients.PartyId: .created = GetTickCountRaw(): Call SetUserRef(.leader, leader)
        For i = 1 To recipients.count
            If IsValidUserRef(recipients.Users(i)) Then
                u = recipients.Users(i).ArrayIndex
                If CanEnroll(u, quest) And mSessions(u).proposal = 0 And Not UserList(u).flags.RespondiendoPregunta Then
                    .count = .count + 1
                    With .members(.count)
                        .characterId = UserList(u).Id: .name = UserList(u).name
                        Call SetUserRef(.online, u)
                        If u = leader Then .vote = 1
                    End With
                End If
            End If
        Next i
        If .count < 2 Then .id = 0: mSessions(userIndex).result = HuntNeedParty: GoTo Reply
        ' Stable character ordering, independent of roster order / connection slot.
        For i = 2 To .count
            candidate = .members(i): j = i - 1
            Do While j > 0
                If .members(j).characterId < candidate.characterId Then Exit Do
                .members(j + 1) = .members(j): j = j - 1
            Loop
            .members(j + 1) = candidate
        Next i
        For i = 1 To .count
            u = .members(i).online.ArrayIndex
            mSessions(u).proposal = leader
            If u <> leader Then
                UserList(u).flags.pregunta = QUESTION_PARTY_QUEST
                UserList(u).flags.RespondiendoPregunta = True
            End If
        Next i
    End With
    mLeaders(CStr(leader)) = True
    mSessions(userIndex).result = HuntProposed
    Call PublishProposal(leader)
    Exit Sub
Reply:
    Call WritePartyQuestState(userIndex)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.ProposePartyQuest", Erl)
    mSessions(userIndex).result = HuntStorageError
    Call WritePartyQuestState(userIndex)
End Sub

Private Function ProposalCanStart(ByVal leader As Integer) As Boolean
    If Not IsValidUserRef(mProposals(leader).leader) Then Exit Function
    If PartyLeaderIndex(leader) <> leader Then Exit Function
    If UserList(leader).Grupo.Id <> mProposals(leader).party Then Exit Function
    Dim i As Integer, count As Integer, hasLeader As Boolean, u As Integer
    For i = 1 To mProposals(leader).count
        With mProposals(leader).members(i)
            If .vote = 1 And IsValidUserRef(.online) Then
                u = .online.ArrayIndex
                If MemberEligible(leader, u, mProposals(leader).quest) Then
                    count = count + 1
                    If u = leader Then hasLeader = True
                End If
            End If
        End With
    Next i
    ProposalCanStart = count >= 2 And hasLeader
End Function

Private Sub StartHunt(ByVal leader As Integer)
    On Error GoTo Failed
    Call RefreshPartyQuestProposal(leader)
    If Not ProposalCanStart(leader) Then mSessions(leader).result = HuntNeedParty: Exit Sub
    Dim run As HuntRun, i As Integer, u As Integer, cache As Integer
    cache = FreeRunSlot(): If cache = 0 Then mSessions(leader).result = HuntBusy: Exit Sub
    run.quest = mProposals(leader).quest: run.leaderId = UserList(leader).Id
    run.target = QuestList(run.quest).RequiredNPC(1).NpcIndex
    run.required = QuestList(run.quest).RequiredNPC(1).amount
    run.reward = QuestList(run.quest).RewardGLD * SvrConfig.GetValue("GoldMult")
    If run.reward < 1 Or run.reward > MAXORO Then Err.Raise vbObjectError + 860, , "Invalid hunt reward"
    For i = 1 To mProposals(leader).count
        If mProposals(leader).members(i).vote = 1 Then
            run.count = run.count + 1: run.members(run.count) = mProposals(leader).members(i)
        End If
    Next i
    If Not mRunIdsKnown Or mNextRun = &H7FFFFFFF Then mSessions(leader).result = HuntBusy: Exit Sub
    If Not CanReserveHuntOwners(run) Then mSessions(leader).result = HuntBusy: Exit Sub
    For i = 1 To run.count
        If Not RegisterHuntSaveOwner(run.members(i).online.ArrayIndex) Then Err.Raise vbObjectError + 862, , "Hunt owner capacity"
    Next i
    mNextRun = mNextRun + 1: run.id = mNextRun: run.revision = 1
    mRuns(cache) = run
    Call CloseProposal(leader, HuntStarted, cache)
    For i = 1 To run.count
        u = run.members(i).online.ArrayIndex
        Call ProjectRun(cache, u, True)
        Call WriteQuestListSend(u)
        Call WritePartyQuestState(u)
    Next i
    Exit Sub
Failed:
    mSessions(leader).result = HuntStorageError
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.StartHunt", Erl)
End Sub

Public Sub PartyQuestNpcDeath(ByVal credited As Integer, ByRef npc As t_Npc, ByRef recipients As t_PartyRewardRecipients)
    On Error GoTo Failed
    If credited < 1 Or credited > mCapacity Then Exit Sub
    If npc.npcType = DummyTarget Or npc.MaestroUser.ArrayIndex <> 0 Then Exit Sub
    If Not recipients.IsParty Or recipients.count < 2 Then Exit Sub
    Dim cache As Integer, member As Integer, i As Integer, mask As Long, count As Integer, u As Integer
    cache = mSessions(credited).runSlot: If cache = 0 Then Exit Sub
    If mSessions(credited).characterId <> UserList(credited).Id Or mSessions(credited).version <> UserList(credited).VersionId Then Exit Sub
    With mRuns(cache)
        member = RunMember(cache, UserList(credited).Id)
        If member = 0 Then Exit Sub
        If .members(member).status <> 0 Or .target <> npc.Numero Or .progress >= .required Then Exit Sub
        For i = 1 To recipients.count
            If IsValidUserRef(recipients.Users(i)) Then
                u = recipients.Users(i).ArrayIndex
                member = RunMember(cache, UserList(u).Id)
                If member > 0 Then
                    If .members(member).status = 0 Then
                        mask = mask Or (2 ^ (member - 1)): count = count + 1
                    End If
                End If
            End If
        Next i
        If count < 2 Then Exit Sub
        .progress = .progress + 1
        .killMasks(.progress) = mask: .revision = .revision + 1
        For member = 1 To .count
            If (mask And (2 ^ (member - 1))) <> 0 Then .members(member).participated = True
        Next member
    End With
    Call PublishRun(cache)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.PartyQuestNpcDeath", Erl)
End Sub

Public Function PartyQuestReady(ByVal userIndex As Integer, ByVal quest As Integer) As Boolean
    If userIndex < 1 Or userIndex > mCapacity Then Exit Function
    If mSessions(userIndex).characterId <> UserList(userIndex).Id Or mSessions(userIndex).version <> UserList(userIndex).VersionId Then Exit Function
    Dim cache As Integer, member As Integer
    cache = mSessions(userIndex).runSlot: If cache = 0 Then Exit Function
    member = RunMember(cache, UserList(userIndex).Id): If member = 0 Then Exit Function
    With mRuns(cache)
        PartyQuestReady = .quest = quest And .progress >= .required And .members(member).participated And .members(member).status = 0
    End With
End Function

Public Sub ClaimPartyQuest(ByVal userIndex As Integer, ByVal quest As Integer)
    On Error GoTo Failed
    If Not PartyQuestReady(userIndex, quest) Then Exit Sub
    If Not IsValidNpcRef(UserList(userIndex).flags.TargetNPC) Then Exit Sub
    Dim npcIndex As Integer, cache As Integer, member As Integer, slot As Byte
    npcIndex = UserList(userIndex).flags.TargetNPC.ArrayIndex
    If NpcList(npcIndex).Numero <> QuestList(quest).TalkTo Then Exit Sub
    If NpcList(npcIndex).pos.Map <> UserList(userIndex).pos.Map Then Exit Sub
    If Distancia(NpcList(npcIndex).pos, UserList(userIndex).pos) > 5 Then Exit Sub
    cache = mSessions(userIndex).runSlot: member = RunMember(cache, UserList(userIndex).Id)
    slot = QuestSlotFor(userIndex, quest): If slot = 0 Then Exit Sub
    Dim reward As Long, cycle As Long, granted As New Collection
    reward = mRuns(cache).reward
    If CDbl(UserList(userIndex).Stats.GLD) + reward > MAXORO Then
        mSessions(userIndex).result = HuntBusy: Call WritePartyQuestState(userIndex): Exit Sub
    End If
    cycle = CaptureQuestFeedbackLifecycle(userIndex, slot, quest)
    UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD + reward
    mRuns(cache).members(member).status = 1: mRuns(cache).revision = mRuns(cache).revision + 1
    mSessions(userIndex).runSlot = 0
    Call CleanQuestSlot(userIndex, slot): Call ArrangeUserQuests(userIndex)
    If Not UserDoneQuest(userIndex, quest) Then Call AddDoneQuest(userIndex, quest)
    Call AddGrantedQuestReward(granted, 2, 0, reward)
    Call WriteChatOverHead(userIndex, "QUESTFIN*" & quest, NpcList(npcIndex).Char.charindex, vbYellow)
    Call QueueQuestCompletionFeedback(userIndex, cycle, quest, granted)
    Call WriteUpdateGold(userIndex)
    mSessions(userIndex).result = HuntClaimed
    Call WritePartyQuestState(userIndex): Call WriteQuestListSend(userIndex)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.ClaimPartyQuest", Erl)
End Sub

Public Sub AbandonPartyQuest(ByVal userIndex As Integer)
    On Error GoTo Failed
    Call EnsureHuntSession(userIndex)
    Dim cache As Integer, member As Integer, slot As Byte, quest As Integer
    cache = mSessions(userIndex).runSlot
    If cache = 0 Then
        If mSessions(userIndex).proposal > 0 Then
            If mSessions(userIndex).proposal = userIndex Then Call CloseProposal(userIndex, HuntAbandoned)
        End If
        Call WritePartyQuestState(userIndex): Exit Sub
    End If
    quest = mRuns(cache).quest: member = RunMember(cache, UserList(userIndex).Id)
    mRuns(cache).members(member).status = 2: mRuns(cache).revision = mRuns(cache).revision + 1
    mSessions(userIndex).runSlot = 0
    slot = QuestSlotFor(userIndex, quest)
    If slot > 0 Then Call CleanQuestSlot(userIndex, slot)
    Call ArrangeUserQuests(userIndex)
    mSessions(userIndex).result = HuntAbandoned
    Call WriteQuestListSend(userIndex): Call WritePartyQuestState(userIndex): Call PublishRun(cache)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.AbandonPartyQuest", Erl)
End Sub

Public Sub HandlePartyQuestAction(ByVal userIndex As Integer)
    On Error GoTo Failed
    Dim request As Long, action As Byte, identity As Long, quest As Integer
    request = reader.ReadInt32(): action = reader.ReadInt8(): identity = reader.ReadInt32(): quest = reader.ReadInt16()
    If reader.GetAvailable() <> 0 Or identity < 0 Or quest <= 0 Then Exit Sub
    Call ExecutePartyQuestAction(userIndex, request, action, identity, quest)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.HandlePartyQuestAction", Erl)
End Sub

Public Sub ExecutePartyQuestAction(ByVal userIndex As Integer, ByVal request As Long, ByVal action As Byte, ByVal identity As Long, ByVal quest As Integer)
    On Error GoTo Failed
    If Not IsRewardCredit(userIndex) Or request <= 0 Or action > HuntAbandon Then Exit Sub
    Call EnsureHuntSession(userIndex)
    Dim key As String, leader As Integer, i As Integer
    key = CStr(action) & ":" & CStr(identity) & ":" & CStr(quest)
    If request <= mSessions(userIndex).request Then
        If request = mSessions(userIndex).request And key = mSessions(userIndex).requestKey Then Call WritePartyQuestState(userIndex)
        Exit Sub
    End If
    mSessions(userIndex).request = request: mSessions(userIndex).requestKey = key: mSessions(userIndex).result = HuntState
    Select Case action
        Case HuntRefresh
        Case HuntPropose
            If quest <> HuntDefinition() Then Exit Sub
            If IsValidNpcRef(UserList(userIndex).flags.TargetNPC) Then Call ProposePartyQuest(userIndex, quest, UserList(userIndex).flags.TargetNPC.ArrayIndex)
        Case HuntAccept, HuntDecline, HuntStart
            leader = mSessions(userIndex).proposal
            If leader < 1 Then GoTo Stale
            If identity <> mProposals(leader).id Or quest <> mProposals(leader).quest Then GoTo Stale
            If TicksElapsed(mProposals(leader).created, GetTickCountRaw()) >= PROPOSAL_LIFETIME Then
                Call CloseProposal(leader, HuntExpired): GoTo Stale
            End If
            If action = HuntStart Then
                If leader <> userIndex Then GoTo Stale
                Call StartHunt(leader)
            Else
                For i = 1 To mProposals(leader).count
                    With mProposals(leader).members(i)
                        If .characterId = UserList(userIndex).Id And IsValidUserRef(.online) Then
                            If .online.ArrayIndex = userIndex And .vote = 0 Then
                                If action = HuntAccept And MemberEligible(leader, userIndex, quest) Then .vote = 1 Else .vote = 2
                                UserList(userIndex).flags.pregunta = 0: UserList(userIndex).flags.RespondiendoPregunta = False
                            End If
                        End If
                    End With
                Next i
                Call RefreshPartyQuestProposal(leader)
            End If
        Case HuntAbandon
            If mSessions(userIndex).runSlot > 0 Then
                If mRuns(mSessions(userIndex).runSlot).id <> identity Then GoTo Stale
            ElseIf mSessions(userIndex).proposal > 0 Then
                If mProposals(mSessions(userIndex).proposal).id <> identity Then GoTo Stale
            Else
                GoTo Stale
            End If
            Call AbandonPartyQuest(userIndex)
    End Select
    Call WritePartyQuestState(userIndex)
    Exit Sub
Stale:
    mSessions(userIndex).result = HuntStale: Call WritePartyQuestState(userIndex)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.ExecutePartyQuestAction", Erl)
End Sub

Public Sub WritePartyQuestState(ByVal userIndex As Integer)
    On Error GoTo Failed
    Call EnsureHuntSession(userIndex)
    Dim quest As Integer, cache As Integer, proposal As Integer, state As Byte, i As Integer, count As Integer, member As Integer
    Dim writer As Network.Writer, queued As Boolean, reward As Long, target As Integer, required As Integer, progress As Integer, leaderId As Long, localVote As Byte
    quest = HuntDefinition(): If quest = 0 Then Exit Sub
    cache = mSessions(userIndex).runSlot: proposal = mSessions(userIndex).proposal
    target = QuestList(quest).RequiredNPC(1).NpcIndex: required = QuestList(quest).RequiredNPC(1).amount
    reward = QuestList(quest).RewardGLD * SvrConfig.GetValue("GoldMult")
    If cache > 0 Then
        state = 3: If mRuns(cache).progress >= mRuns(cache).required Then state = 4
        count = mRuns(cache).count: progress = mRuns(cache).progress: reward = mRuns(cache).reward
        target = mRuns(cache).target: required = mRuns(cache).required
        leaderId = mRuns(cache).leaderId
    ElseIf proposal > 0 Then
        state = 1: If proposal <> userIndex Then state = 2
        count = mProposals(proposal).count: leaderId = UserList(proposal).Id
    End If
    Set writer = Protocol_Writes.GetWriterBuffer()
    mSessions(userIndex).revision = mSessions(userIndex).revision + 1
    Call writer.WriteInt16(ServerPacketID.ePartyQuestState)
    Call writer.WriteInt32(mSessions(userIndex).revision)
    Call writer.WriteInt32(mSessions(userIndex).request)
    Call writer.WriteInt8(mSessions(userIndex).result)
    Call writer.WriteInt8(state)
    If proposal > 0 Then Call writer.WriteInt32(mProposals(proposal).id) Else Call writer.WriteInt32(0)
    If cache > 0 Then Call writer.WriteInt32(mRuns(cache).id) Else Call writer.WriteInt32(0)
    Call writer.WriteInt16(quest): Call writer.WriteInt16(QuestList(quest).TalkTo)
    Call writer.WriteInt16(target): Call writer.WriteInt16(required): Call writer.WriteInt16(progress)
    Call writer.WriteInt32(reward): Call writer.WriteInt32(UserList(userIndex).Id): Call writer.WriteInt32(leaderId)
    Call writer.WriteBool(CanProposeHunt(userIndex, quest))
    If proposal > 0 Then Call writer.WriteBool(ProposalCanStart(proposal)) Else Call writer.WriteBool(False)
    Call writer.WriteBool(PartyQuestReady(userIndex, quest))
    Call writer.WriteInt8(count)
    For i = 1 To count
        Dim row As HuntMember
        If cache > 0 Then row = mRuns(cache).members(i) Else row = mProposals(proposal).members(i)
        Call writer.WriteInt32(row.characterId): Call writer.WriteString8(row.name)
        Call writer.WriteInt8(row.vote): Call writer.WriteBool(row.participated): Call writer.WriteInt8(row.status)
    Next i
    Call modSendData.SendData(ToIndex, userIndex, , False, queued, True)
    mSessions(userIndex).dirty = Not queued: mSessions(userIndex).retryAt = GetTickCountRaw()
    If queued Then
        If mRetries.Exists(CStr(userIndex)) Then Call mRetries.Remove(CStr(userIndex))
    Else
        mRetries(CStr(userIndex)) = True
    End If
    Exit Sub
Failed:
    mSessions(userIndex).dirty = True: mSessions(userIndex).retryAt = GetTickCountRaw()
    mRetries(CStr(userIndex)) = True
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.WritePartyQuestState", Erl)
End Sub

Public Sub PartyQuestLogout(ByVal userIndex As Integer)
    On Error GoTo Failed
    If userIndex < 1 Or userIndex > mCapacity Then Exit Sub
    If mSessions(userIndex).proposal > 0 Then
        Dim leader As Integer
        leader = mSessions(userIndex).proposal
        If leader = userIndex Then
            Call CloseProposal(leader, HuntExpired)
        Else
            Dim i As Integer
            For i = 1 To mProposals(leader).count
                If mProposals(leader).members(i).characterId = UserList(userIndex).Id Then mProposals(leader).members(i).vote = 3
            Next i
            Call PublishProposal(leader)
        End If
    End If
    If Not mOnline Is Nothing Then
        If mOnline.Exists(CStr(UserList(userIndex).Id)) Then Call mOnline.Remove(CStr(UserList(userIndex).Id))
    End If
    Dim blankState As HuntSession
    mSessions(userIndex) = blankState
    If mRetries.Exists(CStr(userIndex)) Then Call mRetries.Remove(CStr(userIndex))
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.PartyQuestLogout", Erl)
End Sub

Public Sub ProcessPartyQuestUpdates()
    On Error GoTo Failed
    ' Only pending proposals / failed enqueues; never poll every logged-in user.
    If mOnline Is Nothing Then Exit Sub
    Dim key As Variant, u As Integer, nowRaw As Long
    nowRaw = GetTickCountRaw()
    For Each key In mLeaders.Keys
        u = CInt(key)
        If TicksElapsed(mProposals(u).created, nowRaw) >= PROPOSAL_LIFETIME Then Call CloseProposal(u, HuntExpired)
    Next key
    For Each key In mRetries.Keys
        u = CInt(key)
        If OnlineCharacter(mSessions(u).characterId) = u Then
            If TicksElapsed(mSessions(u).retryAt, nowRaw) >= 500 Then Call WritePartyQuestState(u)
        Else
            Call mRetries.Remove(key)
        End If
    Next key
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.ProcessPartyQuestUpdates", Erl)
End Sub

#If UNIT_TEST = 1 Then
Public Sub ResetPartyQuestTests()
    mCapacity = 0: mDefinition = 0: mRunIdsKnown = False: mNextRun = 0: Erase mSessions: Erase mRuns: Erase mProposals: Set mOnline = Nothing
    Set mRetries = Nothing: Set mLeaders = Nothing
    Call ResetHuntSaveOwners()
    HoldPartySaves = False: FailPartyRosterRead = False: LastPartySaveFailureForTest = vbNullString
    Dim blankBatch As HuntSaveBatch
    mHeldBatch = blankBatch
End Sub
Public Function PartyQuestProgressForTest(ByVal userIndex As Integer) As Integer
    If mSessions(userIndex).runSlot > 0 Then PartyQuestProgressForTest = mRuns(mSessions(userIndex).runSlot).progress
End Function
Public Function PartyQuestProposalForTest(ByVal userIndex As Integer) As Long
    If mSessions(userIndex).proposal > 0 Then PartyQuestProposalForTest = mProposals(mSessions(userIndex).proposal).id
End Function
#End If

Private Function RegisterHuntSaveOwner(ByVal userIndex As Integer) As Boolean
    If mSaveOwners Is Nothing Then
        Set mSaveOwners = CreateObject("Scripting.Dictionary")
        Set mPendingSaves = CreateObject("Scripting.Dictionary")
    End If
    If mSaveCapacity < mCapacity Then
        mSaveCapacity = mCapacity
        ReDim Preserve mOwners(1 To mSaveCapacity)
        ReDim Preserve mBatches(1 To mSaveCapacity)
    End If
    Dim key As String, i As Integer
    key = CStr(UserList(userIndex).Id)
    If mSaveOwners.Exists(key) Then RegisterHuntSaveOwner = True: Exit Function
    For i = 1 To mSaveCapacity
        If mOwners(i).characterId = 0 Then
            mOwners(i).characterId = UserList(userIndex).Id
            mOwners(i).nextSequence = UserList(userIndex).PartyQuestSaveRevision
            mOwners(i).committedSequence = mOwners(i).nextSequence
            mSaveOwners(key) = i: RegisterHuntSaveOwner = True: Exit Function
        End If
    Next i
End Function

Private Function CanReserveHuntOwners(ByRef run As HuntRun) As Boolean
    Dim i As Integer, needed As Integer, used As Integer
    If Not mSaveOwners Is Nothing Then used = mSaveOwners.count
    For i = 1 To run.count
        If Not PartyQuestOwnsCharacter(run.members(i).characterId) Then needed = needed + 1
    Next i
    CanReserveHuntOwners = used + needed <= mCapacity
End Function

Public Function PartyQuestOwnsCharacter(ByVal characterId As Long) As Boolean
    If mSaveOwners Is Nothing Then Exit Function
    PartyQuestOwnsCharacter = mSaveOwners.Exists(CStr(characterId))
End Function

Private Function CaptureHuntCharacter(ByVal userIndex As Integer, ByVal logout As Boolean) As clsHuntCharacterSave
    Dim owner As Integer, job As New clsHuntCharacterSave
    owner = mSaveOwners(CStr(UserList(userIndex).Id))
    If mOwners(owner).nextSequence = &H7FFFFFFF Then Err.Raise vbObjectError + 862, , "Hunt save sequence exhausted"
    mOwners(owner).nextSequence = mOwners(owner).nextSequence + 1
    Call job.Capture(userIndex, mOwners(owner).nextSequence, logout)
    Set mOwners(owner).latest = job
    Set CaptureHuntCharacter = job
End Function

Private Sub CaptureHuntBatch(ByVal owner As Integer, ByVal logout As Boolean, Optional ByVal refresh As Boolean = False)
    Dim batch As HuntSaveBatch, i As Integer, j As Integer, userIndex As Integer
    Dim job As clsHuntCharacterSave, characterId As Long, included As Object
    Set included = CreateObject("Scripting.Dictionary")
    characterId = mOwners(owner).characterId
    userIndex = OnlineCharacter(characterId)
    Set job = mOwners(owner).latest
    If refresh And userIndex > 0 Then Set job = CaptureHuntCharacter(userIndex, logout)
    If job Is Nothing Then Err.Raise vbObjectError + 862, , "Missing owned character snapshot"
    batch.characterId = characterId: batch.sequence = job.Sequence
    Set batch.characters = New Collection
    Call batch.characters.Add(job, CStr(characterId))
    included(CStr(characterId)) = True
    ReDim batch.runs(1 To mCapacity)
    For i = 1 To mCapacity
        If mRuns(i).id > 0 And mRuns(i).revision > mRuns(i).savedRevision Then
            If RunMember(i, characterId) > 0 Then
                batch.count = batch.count + 1: batch.runs(batch.count) = mRuns(i)
                For j = 1 To mRuns(i).count
                    With mRuns(i).members(j)
                        If .status = 1 And .savedStatus <> 1 And Not included.Exists(CStr(.characterId)) Then
                            ' Committing someone else's pending claim also commits their
                            ' latest whole character snapshot in the SAME transaction.
                            userIndex = OnlineCharacter(.characterId)
                            If userIndex > 0 Then
                                Set job = CaptureHuntCharacter(userIndex, False)
                            Else
                                Set job = mOwners(mSaveOwners(CStr(.characterId))).latest
                            End If
                            If job Is Nothing Then Err.Raise vbObjectError + 862, , "Pending claim has no wallet snapshot"
                            Call batch.characters.Add(job, CStr(.characterId))
                            included(CStr(.characterId)) = True
                        End If
                    End With
                Next j
            End If
        End If
    Next i
    ' Persist historical runs before a newer acceptance for the same character,
    ' so its partial unique active-membership index sees the final old status.
    Dim ordered As HuntRun, position As Integer
    For i = 2 To batch.count
        ordered = batch.runs(i): position = i
        Do While position > 1
            If batch.runs(position - 1).id < ordered.id Then Exit Do
            batch.runs(position) = batch.runs(position - 1): position = position - 1
        Loop
        batch.runs(position) = ordered
    Next i
    mBatches(owner) = batch
    mPendingSaves(CStr(owner)) = True
End Sub

Private Sub PersistHuntSnapshot(ByRef snapshot As HuntRun)
    Dim rows As ADODB.Recordset, i As Integer, stored As Long
    Set rows = HuntSql("SELECT revision FROM party_quest_run WHERE id=?", snapshot.id)
    If Not rows.EOF Then stored = rows.Fields(0).value
    Call CloseHuntRows(rows)
    If stored >= snapshot.revision Then Exit Sub
    With snapshot
        Set rows = HuntSql("INSERT INTO party_quest_run(id,revision,quest_id,leader_id,target_npc,required,progress,reward_gold) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET revision=excluded.revision,progress=excluded.progress WHERE party_quest_run.revision<excluded.revision", .id, .revision, .quest, .leaderId, .target, .required, .progress, .reward)
        Call CloseHuntRows(rows)
        For i = 1 To .count
            Set rows = HuntSql("INSERT INTO party_quest_member(run_id,ordinal,character_id,user_id,name,participated,status) VALUES(?,?,?,(SELECT id FROM user WHERE id=?),?,?,?) ON CONFLICT(run_id,character_id) DO UPDATE SET user_id=excluded.user_id,participated=excluded.participated,status=excluded.status", .id, i, .members(i).characterId, .members(i).characterId, .members(i).name, Abs(CInt(.members(i).participated)), .members(i).status)
            Call CloseHuntRows(rows)
        Next i
        For i = 1 To .progress
            Set rows = HuntSql("INSERT OR IGNORE INTO party_quest_kill(run_id,ordinal,qualified_mask) VALUES(?,?,?)", .id, i, .killMasks(i))
            Call CloseHuntRows(rows)
        Next i
    End With
End Sub

Private Function CommitHuntBatch(ByRef batch As HuntSaveBatch) As Boolean
    On Error GoTo Failed
    If mFlushing Then Exit Function
    If Not PartyQuestLegacyWritesComplete() Then
        #If UNIT_TEST = 1 Then
            LastPartySaveFailureForTest = "legacy writes pending"
        #End If
        Exit Function
    End If
    #If UNIT_TEST = 1 Then
        If HoldPartySaves Then Exit Function
    #End If
    Dim inTransaction As Boolean, rows As ADODB.Recordset, job As clsHuntCharacterSave, command As ADODB.Command
    Dim i As Integer, j As Integer, owner As Integer, stored As Long, exists As Boolean
    mFlushing = True
    Call Connection.BeginTrans(): inTransaction = True
    For i = 1 To batch.count: Call PersistHuntSnapshot(batch.runs(i)): Next i
    For Each job In batch.characters
        Set rows = HuntSql("SELECT party_quest_save_revision FROM user WHERE id=?", job.CharacterId)
        exists = Not rows.EOF
        If exists Then stored = rows.Fields(0).value
        Call CloseHuntRows(rows)
        If exists Then
            If stored < job.Sequence Then
                For Each command In job.Commands
                    Set command.ActiveConnection = Connection
                    Call CountPartyQuestSql()
                    Call command.Execute()
                Next command
                Set rows = HuntSql("UPDATE user SET party_quest_save_revision=? WHERE id=? AND party_quest_save_revision<?", job.Sequence, job.CharacterId, job.Sequence)
                Call CloseHuntRows(rows)
            End If
        End If
    Next job
    Call Connection.CommitTrans(): inTransaction = False
    ' Only a completed transaction acknowledges immutable revisions. Never clear
    ' a newer run mutation or a newer pending character snapshot.
    For i = 1 To batch.count
        For j = 1 To mCapacity
            If mRuns(j).id = batch.runs(i).id Then
                If mRuns(j).savedRevision <= batch.runs(i).revision Then
                    mRuns(j).savedRevision = batch.runs(i).revision
                    Dim member As Integer
                    For member = 1 To mRuns(j).count
                        mRuns(j).members(member).savedStatus = batch.runs(i).members(member).status
                    Next member
                End If
            End If
        Next j
    Next i
    For Each job In batch.characters
        If mSaveOwners.Exists(CStr(job.CharacterId)) Then
            owner = mSaveOwners(CStr(job.CharacterId))
            If mOwners(owner).committedSequence < job.Sequence Then mOwners(owner).committedSequence = job.Sequence
            If mOwners(owner).nextSequence = job.Sequence Then
                Call job.Acknowledge()
                Set mOwners(owner).latest = Nothing
            End If
        End If
    Next job
    CommitHuntBatch = True: mFlushing = False
    Exit Function
Failed:
    Dim number As Long, description As String
    number = Err.Number: description = Err.Description
    #If UNIT_TEST = 1 Then
        LastPartySaveFailureForTest = CStr(number) & ": " & description
        If Not command Is Nothing Then LastPartySaveFailureForTest = LastPartySaveFailureForTest & "; SQL=" & command.CommandText
        Call UnitTesting.Progress("Party save failure: " & LastPartySaveFailureForTest)
    #End If
    Resume Rollback
Rollback:
    On Error Resume Next
    If inTransaction Then Call Connection.RollbackTrans()
    Call CloseHuntRows(rows)
    mFlushing = False
    Call TraceError(number, description, "modPartyQuest.CommitHuntBatch", Erl)
End Function

Private Sub RetireCompletedHuntSaves()
    Dim key As Variant, owner As Integer, i As Integer, j As Integer, saved As Boolean
    For Each key In mPendingSaves.Keys
        owner = CInt(key): saved = mBatches(owner).sequence <= mOwners(owner).committedSequence
        For i = 1 To mBatches(owner).count
            For j = 1 To mCapacity
                If mRuns(j).id = mBatches(owner).runs(i).id Then
                    If mRuns(j).savedRevision < mBatches(owner).runs(i).revision Then saved = False
                End If
            Next j
        Next i
        If saved Then
            Call mPendingSaves.Remove(key)
            Dim blankState As HuntSaveBatch
            mBatches(owner) = blankState
            Call ReleaseHuntOwnerIfClean(owner)
        End If
    Next key
End Sub

Private Function FlushHuntOwner(ByVal owner As Integer) As Boolean
    If Not mPendingSaves.Exists(CStr(owner)) Then FlushHuntOwner = True: Exit Function
    ' Immutable even if an ADO completion event causes another save request.
    Dim snapshot As HuntSaveBatch
    snapshot = mBatches(owner)
    If CommitHuntBatch(snapshot) Then
        Call RetireCompletedHuntSaves()
        FlushHuntOwner = True
    End If
End Function

Private Sub ReleaseHuntOwnerIfClean(ByVal owner As Integer)
    If mOwners(owner).characterId = 0 Then Exit Sub
    If Not mOwners(owner).latest Is Nothing Then Exit Sub
    If mPendingSaves.Exists(CStr(owner)) Then Exit Sub
    If OnlineCharacter(mOwners(owner).characterId) > 0 Then Exit Sub
    If CharacterHasDirtyHunt(mOwners(owner).characterId) Then Exit Sub
    Call mSaveOwners.Remove(CStr(mOwners(owner).characterId))
    Dim blankState As HuntSaveOwner
    mOwners(owner) = blankState
End Sub

Private Function CharacterHasDirtyHunt(ByVal characterId As Long) As Boolean
    Dim i As Integer
    For i = 1 To mCapacity
        If mRuns(i).revision > mRuns(i).savedRevision Then
            If RunMember(i, characterId) > 0 Then CharacterHasDirtyHunt = True: Exit Function
        End If
    Next i
End Function

Public Sub SavePartyQuestCharacter(ByVal userIndex As Integer, ByVal logout As Boolean)
    On Error GoTo Failed
    Dim owner As Integer, job As clsHuntCharacterSave
    owner = mSaveOwners(CStr(UserList(userIndex).Id))
    ' Capture even after connection loss: identity is persistent, not the slot.
    Set job = CaptureHuntCharacter(userIndex, logout)
    Call CaptureHuntBatch(owner, logout)
    Call FlushHuntOwner(owner)
    Exit Sub
Failed:
    #If UNIT_TEST = 1 Then
        LastPartySaveFailureForTest = CStr(Err.Number) & ": " & Err.Description
        Call UnitTesting.Progress("Party capture failure: " & LastPartySaveFailureForTest)
    #End If
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.SavePartyQuestCharacter", Erl)
End Sub

Public Function PreparePartyQuestLogin(ByVal characterId As Long) As Boolean
    On Error GoTo Failed
    PreparePartyQuestLogin = True
    If Not PartyQuestOwnsCharacter(characterId) Then Exit Function
    Dim owner As Integer
    owner = mSaveOwners(CStr(characterId))
    ' Do not load an old wallet while a retained logout save still owns its claim.
    If mPendingSaves.Exists(CStr(owner)) Or Not mOwners(owner).latest Is Nothing Then
        Call CaptureHuntBatch(owner, True, True)
        PreparePartyQuestLogin = FlushHuntOwner(owner)
    End If
    Exit Function
Failed:
    PreparePartyQuestLogin = False
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.PreparePartyQuestLogin", Erl)
End Function

Public Sub RetryPartyQuestSaves(Optional ByVal limit As Integer = 2)
    On Error GoTo Failed
    If mPendingSaves Is Nothing Or mFlushing Then Exit Sub
    Dim key As Variant, owner As Integer, attempted As Integer
    For Each key In mPendingSaves.Keys
        owner = CInt(key)
        ' A queued, not-yet-dispatched save coalesces to the latest shared state.
        Call CaptureHuntBatch(owner, False, True)
        Call FlushHuntOwner(owner)
        attempted = attempted + 1: If attempted >= limit Then Exit For
    Next key
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.RetryPartyQuestSaves", Erl)
End Sub

Public Function FlushPartyQuestShutdown() As Boolean
    Call RetryPartyQuestSaves(mSaveCapacity)
    If mPendingSaves Is Nothing Then FlushPartyQuestShutdown = True: Exit Function
    FlushPartyQuestShutdown = mPendingSaves.count = 0
End Function

Public Sub PartyQuestBeforeSlotReset(ByVal userIndex As Integer)
    On Error GoTo Failed
    If PartyQuestOwnsCharacter(UserList(userIndex).Id) Then
        Dim owner As Integer, job As clsHuntCharacterSave
        owner = mSaveOwners(CStr(UserList(userIndex).Id))
        If Not mOwners(owner).latest Is Nothing Or CharacterHasDirtyHunt(UserList(userIndex).Id) Then
            Set job = CaptureHuntCharacter(userIndex, True)
            Call CaptureHuntBatch(owner, True)
        End If
    End If
    Call PartyQuestLogout(userIndex)
    If owner > 0 Then Call ReleaseHuntOwnerIfClean(owner)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.PartyQuestBeforeSlotReset", Erl)
End Sub

Private Sub ResetHuntSaveOwners()
    Set mSaveOwners = Nothing: Set mPendingSaves = Nothing
    Erase mOwners: Erase mBatches: mSaveCapacity = 0: mFlushing = False
End Sub

Public Sub CountPartyQuestSql()
    #If UNIT_TEST = 1 Then
        PartyQuestSqlCount = PartyQuestSqlCount + 1
    #End If
End Sub

#If UNIT_TEST = 1 Then
Public Function PartyQuestMaskForTest(ByVal userIndex As Integer, ByVal kill As Integer) As Long
    PartyQuestMaskForTest = mRuns(mSessions(userIndex).runSlot).killMasks(kill)
End Function
Public Function HasPendingPartySaveForTest(ByVal characterId As Long) As Boolean
    If Not PartyQuestOwnsCharacter(characterId) Then Exit Function
    HasPendingPartySaveForTest = mPendingSaves.Exists(CStr(mSaveOwners(CStr(characterId))))
End Function
Public Sub HoldPartySnapshotForTest(ByVal userIndex As Integer)
    mHeldBatch = mBatches(mSaveOwners(CStr(UserList(userIndex).Id)))
End Sub
Public Function CompleteHeldPartySnapshotForTest() As Boolean
    CompleteHeldPartySnapshotForTest = CommitHuntBatch(mHeldBatch)
End Function
Public Function PartyQuestDirtyForTest(ByVal userIndex As Integer) As Boolean
    Dim cache As Integer
    cache = mSessions(userIndex).runSlot
    If cache > 0 Then PartyQuestDirtyForTest = mRuns(cache).revision > mRuns(cache).savedRevision
End Function
#End If

#If UNIT_TEST = 1 Then
Public Function PartyQuestCacheCountForTest() As Integer
    Dim i As Integer
    For i = 1 To mCapacity
        If mRuns(i).id > 0 Then PartyQuestCacheCountForTest = PartyQuestCacheCountForTest + 1
    Next i
End Function
Public Function PartyQuestMemberCountForTest(ByVal userIndex As Integer) As Integer
    PartyQuestMemberCountForTest = mRuns(mSessions(userIndex).runSlot).count
End Function
#End If
