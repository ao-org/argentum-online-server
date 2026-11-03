Attribute VB_Name = "modPartyQuest"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Private Const MAX_MEMBERS As Integer = 6
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
End Type
Private Type HuntRun
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

' All writes here use the synchronous connection and propagate errors. Query/Execute
' intentionally swallow errors / dispatch asynchronously and cannot own a claim.
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
               QuestList(q).RewardEXP = 0 And QuestList(q).RewardSpellCount = 0 Then
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
        If mRuns(i).id = 0 Or Not online Then FreeRunSlot = i: Exit Function
    Next i
End Function

Public Function LoadPartyQuest(ByVal userIndex As Integer) As Boolean
    On Error GoTo Failed
    Call EnsureHuntSession(userIndex)
    mSessions(userIndex).runSlot = 0
    Dim rows As ADODB.Recordset, runId As Long, cache As Integer, i As Integer, slot As Byte, quest As Integer
    ' Quest rows can lag behind their durable run. Rebuild only this content type.
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
    Set rows = HuntSql("SELECT DISTINCT r.quest_id FROM party_quest_run r JOIN party_quest_member m ON m.run_id=r.id WHERE m.character_id=? AND m.status=1", UserList(userIndex).Id)
    Do While Not rows.EOF
        quest = rows.Fields(0).value
        If Not UserDoneQuest(userIndex, quest) Then Call AddDoneQuest(userIndex, quest)
        Call rows.MoveNext()
    Loop
    Call CloseHuntRows(rows)
    Set rows = HuntSql("SELECT run_id FROM party_quest_member WHERE character_id=? AND status=0", UserList(userIndex).Id)
    If Not rows.EOF Then runId = rows.Fields(0).value
    Call CloseHuntRows(rows)
    If runId > 0 Then
        For i = 1 To mCapacity
            If mRuns(i).id = runId Then cache = i: Exit For
        Next i
        If cache = 0 Then
            cache = FreeRunSlot(): If cache = 0 Then Err.Raise vbObjectError + 860, , "Party quest cache full"
            Dim blankState As HuntRun
            mRuns(cache) = blankState
            Set rows = HuntSql("SELECT * FROM party_quest_run WHERE id=?", runId)
            With mRuns(cache)
                .id = runId: .quest = rows!quest_id: .leaderId = rows!leader_id
                .target = rows!target_npc: .required = rows!required: .progress = rows!progress: .reward = rows!reward_gold
            End With
            Call CloseHuntRows(rows)
            Set rows = HuntSql("SELECT * FROM party_quest_member WHERE run_id=? ORDER BY ordinal", runId)
            Do While Not rows.EOF
                With mRuns(cache)
                    .count = .count + 1
                    If .count > MAX_MEMBERS Then Err.Raise vbObjectError + 860, , "Invalid persisted hunt roster"
                    With .members(.count)
                        .characterId = rows!character_id: .name = rows!name
                        .participated = rows!participated <> 0: .status = rows!status: .vote = 1
                    End With
                End With
                Call rows.MoveNext()
            Loop
            Call CloseHuntRows(rows)
        End If
        Call ProjectRun(cache, userIndex, False)
    End If
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
    Dim run As HuntRun, i As Integer, u As Integer, cache As Integer, inTransaction As Boolean, rows As ADODB.Recordset
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
    Call Connection.BeginTrans(): inTransaction = True
    Set rows = HuntSql("INSERT INTO party_quest_run(quest_id,leader_id,target_npc,required,reward_gold) VALUES(?,?,?,?,?)", run.quest, run.leaderId, run.target, run.required, run.reward)
    Call CloseHuntRows(rows)
    Set rows = HuntSql("SELECT last_insert_rowid()")
    run.id = rows.Fields(0).value: Call CloseHuntRows(rows)
    For i = 1 To run.count
        Set rows = HuntSql("INSERT INTO party_quest_member(run_id,ordinal,character_id,name) VALUES(?,?,?,?)", run.id, i, run.members(i).characterId, run.members(i).name)
        Call CloseHuntRows(rows)
    Next i
    Call Connection.CommitTrans(): inTransaction = False
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
    If inTransaction Then Call Connection.RollbackTrans()
    Call CloseHuntRows(rows)
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
        Dim rows As ADODB.Recordset, inTransaction As Boolean
        Call Connection.BeginTrans(): inTransaction = True
        Set rows = HuntSql("UPDATE party_quest_run SET progress=progress+1 WHERE id=? AND progress=? AND progress<required", .id, .progress)
        Call CloseHuntRows(rows)
        Set rows = HuntSql("SELECT changes()")
        If rows.Fields(0).value <> 1 Then Err.Raise vbObjectError + 860, , "Stale hunt counter"
        Call CloseHuntRows(rows)
        Set rows = HuntSql("INSERT INTO party_quest_kill(run_id,ordinal,qualified_mask) VALUES(?,?,?)", .id, .progress + 1, mask)
        Call CloseHuntRows(rows)
        For member = 1 To .count
            If (mask And (2 ^ (member - 1))) <> 0 Then
                Set rows = HuntSql("UPDATE party_quest_member SET participated=1 WHERE run_id=? AND character_id=? AND status=0", .id, .members(member).characterId)
                Call CloseHuntRows(rows)
            End If
        Next member
        Call Connection.CommitTrans(): inTransaction = False
        .progress = .progress + 1
        For member = 1 To .count
            If (mask And (2 ^ (member - 1))) <> 0 Then .members(member).participated = True
        Next member
    End With
    Call PublishRun(cache)
    Exit Sub
Failed:
    If inTransaction Then Call Connection.RollbackTrans()
    Call CloseHuntRows(rows)
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
    If UserList(userIndex).PartyQuestSaveBlocked Then
        Dim recovery As ADODB.Recordset
        Set recovery = HuntSql("SELECT gold,party_quest_gold_total FROM user WHERE id=?", UserList(userIndex).Id)
        UserList(userIndex).Stats.GLD = recovery!gold
        UserList(userIndex).PartyQuestGoldSeen = recovery!party_quest_gold_total
        Call CloseHuntRows(recovery)
        Dim previousCache As Integer, previousMember As Integer
        previousCache = mSessions(userIndex).runSlot
        If previousCache > 0 Then
            previousMember = RunMember(previousCache, UserList(userIndex).Id)
            Set recovery = HuntSql("SELECT status FROM party_quest_member WHERE run_id=? AND character_id=?", mRuns(previousCache).id, UserList(userIndex).Id)
            If Not recovery.EOF And previousMember > 0 Then mRuns(previousCache).members(previousMember).status = recovery.Fields(0).value
            Call CloseHuntRows(recovery)
        End If
        If Not LoadPartyQuest(userIndex) Then Exit Sub
        UserList(userIndex).PartyQuestSaveBlocked = False
        Call WriteUpdateGold(userIndex): Call WriteQuestListSend(userIndex): Call WritePartyQuestState(userIndex)
        Exit Sub
    End If
    If Not PartyQuestReady(userIndex, quest) Then Exit Sub
    If Not IsValidNpcRef(UserList(userIndex).flags.TargetNPC) Then Exit Sub
    Dim npcIndex As Integer, cache As Integer, member As Integer, slot As Byte, rows As ADODB.Recordset, inTransaction As Boolean
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
    UserList(userIndex).PartyQuestSaveBlocked = True
    Call Connection.BeginTrans(): inTransaction = True
    Set rows = HuntSql("UPDATE party_quest_member SET status=1 WHERE run_id=? AND character_id=? AND status=0 AND participated=1", mRuns(cache).id, UserList(userIndex).Id)
    Call CloseHuntRows(rows)
    Set rows = HuntSql("SELECT changes()")
    If rows.Fields(0).value <> 1 Then Err.Raise vbObjectError + 860, , "Hunt already claimed"
    Call CloseHuntRows(rows)
    ' Cumulative credit fences asynchronous pre-claim character snapshots.
    Set rows = HuntSql("UPDATE user SET party_quest_gold_total=party_quest_gold_total+?,gold=?+party_quest_gold_total+?-? WHERE id=?", reward, UserList(userIndex).Stats.GLD, reward, UserList(userIndex).PartyQuestGoldSeen, UserList(userIndex).Id)
    Call CloseHuntRows(rows)
    Set rows = HuntSql("INSERT OR IGNORE INTO quest_done(user_id,quest_id) VALUES(?,?)", UserList(userIndex).Id, quest)
    Call CloseHuntRows(rows)
    Set rows = HuntSql("DELETE FROM quest WHERE user_id=? AND quest_id=?", UserList(userIndex).Id, quest)
    Call CloseHuntRows(rows)
    Call Connection.CommitTrans(): inTransaction = False
    UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD + reward
    UserList(userIndex).PartyQuestGoldSeen = UserList(userIndex).PartyQuestGoldSeen + reward
    UserList(userIndex).PartyQuestSaveBlocked = False
    mRuns(cache).members(member).status = 1: mSessions(userIndex).runSlot = 0
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
    Dim failureNumber As Long, failureText As String
    failureNumber = Err.Number: failureText = Err.Description
    On Error Resume Next
    If inTransaction Then
        Err.Clear
        Call Connection.RollbackTrans()
        UserList(userIndex).PartyQuestSaveBlocked = Err.Number <> 0
    End If
    Call CloseHuntRows(rows)
    mSessions(userIndex).result = HuntStorageError
    Call WritePartyQuestState(userIndex)
    Call TraceError(failureNumber, failureText, "modPartyQuest.ClaimPartyQuest", Erl)
End Sub

Public Sub AbandonPartyQuest(ByVal userIndex As Integer)
    On Error GoTo Failed
    Call EnsureHuntSession(userIndex)
    Dim cache As Integer, rows As ADODB.Recordset, member As Integer, slot As Byte, quest As Integer, inTransaction As Boolean
    cache = mSessions(userIndex).runSlot
    If cache = 0 Then
        If mSessions(userIndex).proposal > 0 Then
            If mSessions(userIndex).proposal = userIndex Then Call CloseProposal(userIndex, HuntAbandoned)
        End If
        Call WritePartyQuestState(userIndex): Exit Sub
    End If
    quest = mRuns(cache).quest: member = RunMember(cache, UserList(userIndex).Id)
    Call Connection.BeginTrans(): inTransaction = True
    Set rows = HuntSql("UPDATE party_quest_member SET status=2 WHERE run_id=? AND character_id=? AND status=0", mRuns(cache).id, UserList(userIndex).Id)
    Call CloseHuntRows(rows)
    Set rows = HuntSql("DELETE FROM quest WHERE user_id=? AND quest_id=?", UserList(userIndex).Id, quest)
    Call CloseHuntRows(rows)
    Call Connection.CommitTrans(): inTransaction = False
    mRuns(cache).members(member).status = 2: mSessions(userIndex).runSlot = 0
    slot = QuestSlotFor(userIndex, quest)
    If slot > 0 Then Call CleanQuestSlot(userIndex, slot)
    Call ArrangeUserQuests(userIndex)
    mSessions(userIndex).result = HuntAbandoned
    Call WriteQuestListSend(userIndex): Call WritePartyQuestState(userIndex): Call PublishRun(cache)
    Exit Sub
Failed:
    If inTransaction Then Call Connection.RollbackTrans()
    Call CloseHuntRows(rows)
    mSessions(userIndex).result = HuntStorageError
    Call WritePartyQuestState(userIndex)
    Call TraceError(Err.Number, Err.Description, "modPartyQuest.AbandonPartyQuest", Erl)
End Sub

Public Sub HandlePartyQuestAction(ByVal userIndex As Integer)
    On Error GoTo Failed
    Dim request As Long, action As Byte, identity As Long, quest As Integer
    request = reader.ReadInt32(): action = reader.ReadInt8(): identity = reader.ReadInt32(): quest = reader.ReadInt16()
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
    mCapacity = 0: mDefinition = 0: Erase mSessions: Erase mRuns: Erase mProposals: Set mOnline = Nothing
    Set mRetries = Nothing: Set mLeaders = Nothing
End Sub
Public Function PartyQuestProgressForTest(ByVal userIndex As Integer) As Integer
    If mSessions(userIndex).runSlot > 0 Then PartyQuestProgressForTest = mRuns(mSessions(userIndex).runSlot).progress
End Function
Public Function PartyQuestProposalForTest(ByVal userIndex As Integer) As Long
    If mSessions(userIndex).proposal > 0 Then PartyQuestProposalForTest = mProposals(mSessions(userIndex).proposal).id
End Function
#End If
