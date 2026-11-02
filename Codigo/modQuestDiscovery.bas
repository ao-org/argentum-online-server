Attribute VB_Name = "modQuestDiscovery"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Public Const QUEST_DISCOVERY_PAGE_SIZE As Integer = 128
Public Const QUEST_DISCOVERY_LIMIT As Long = 32767
Private Const MAX_INDEX_LINKS As Long = 131072
Private Const MAX_PAGES_PER_DRAIN As Integer = 4
Private Const FLAG_FINAL As Byte = 1
Private Const FLAG_UNAVAILABLE As Byte = 2
Private mOffers As Object
Private mSkills As Object
Private mGlobals As Object
Private mGlobalGates As Object
Private mRevision As Long
Private mIndexOverflow As Boolean
Private Type DiscoveryCache
    userId As Long
    version As Integer
    known As Boolean
    inputs As String
    fingerprint As String
    sequence As Long
    pendingRows As Collection
    pendingInputs As String
    pendingFingerprint As String
    pendingUnavailable As Boolean
    nextRow As Long
    nextPage As Integer
End Type
Private mCache() As DiscoveryCache
Private mCapacity As Long
#If UNIT_TEST = 1 Then
Public DiscoveryEvaluationCount As Long
#End If

Private Function TraceDiscoveryNpc() As Long
    On Error GoTo Disabled
    Static initialized As Boolean, target As Long
    If Not initialized Then
        initialized = True
        Dim value As Double
        value = Val(Environ$("AO_TRACE_QUEST_DISCOVERY_NPC"))
        If value >= 1 And value <= 32767 Then target = CLng(value)
    End If
    TraceDiscoveryNpc = target
Disabled:
End Function
Private Sub TraceDiscovery(ByVal message As String)
    On Error GoTo Failed
    Static count As Integer
    If TraceDiscoveryNpc() = 0 Or count >= 64 Then Exit Sub
    count = count + 1
    Call LogInfoServidor("QuestDiscoveryTrace " & message)
    Exit Sub
Failed:
    ' Optional diagnostics cannot change publication or eligibility.
End Sub

Public Sub ResetQuestDiscoveryIndex()
    On Error GoTo Failed
    Set mOffers = Nothing
    mRevision = mRevision + 1
    Dim userIndex As Long
    For userIndex = 1 To LastUser
        If UserList(userIndex).flags.UserLogged Then Call SendQuestTurnInSnapshot(CInt(userIndex), False, True)
    Next userIndex
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modQuestDiscovery.ResetQuestDiscoveryIndex", Erl)
End Sub

Public Sub ResetQuestDiscoveryUser(ByVal userIndex As Integer)
    On Error GoTo Failed
    If userIndex < 1 Or userIndex > mCapacity Then Exit Sub
    With mCache(userIndex)
        Set .pendingRows = Nothing
        .pendingInputs = "": .pendingFingerprint = ""
        .inputs = "": .fingerprint = "": .known = False
        .userId = 0: .version = 0: .sequence = 0
    End With
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modQuestDiscovery.ResetQuestDiscoveryUser", Erl)
End Sub

Private Sub BuildIndex()
    On Error GoTo Failed
    Dim npc As Long, offer As Integer, quest As Integer, links As Long, key As String
    Set mOffers = CreateObject("Scripting.Dictionary")
    Set mSkills = CreateObject("Scripting.Dictionary")
    Set mGlobals = CreateObject("Scripting.Dictionary")
    Set mGlobalGates = CreateObject("Scripting.Dictionary")
    mIndexOverflow = False
    For npc = LBound(NpcInfoCache) To UBound(NpcInfoCache)
        If npc > 0 And npc <= 32767 And NpcInfoCache(npc).Exists Then
            For offer = 1 To NpcInfoCache(npc).NumQuest
                quest = NpcInfoCache(npc).QuestNumber(offer)
                If quest > 0 And quest <= UBound(QuestList) Then
                    key = CStr(quest)
                    If Not mOffers.Exists(key) Then
                        Dim providers As Object
                        Set providers = CreateObject("Scripting.Dictionary")
                        Call mOffers.Add(key, providers)
                        If QuestList(quest).RequiredSkill.SkillType > 0 And QuestList(quest).RequiredSkill.SkillType <= NUMSKILLS Then mSkills(CStr(QuestList(quest).RequiredSkill.SkillType)) = True
                        If QuestList(quest).GlobalQuestIndex > 0 Then
                            Dim globalKey As String, gatedQuests As Object
                            globalKey = CStr(QuestList(quest).GlobalQuestIndex)
                            If Not mGlobals.Exists(globalKey) Then
                                Set gatedQuests = CreateObject("Scripting.Dictionary")
                                Call mGlobals.Add(globalKey, gatedQuests)
                            End If
                            Set gatedQuests = mGlobals(globalKey)
                            gatedQuests(CStr(quest)) = True
                        End If
                    End If
                    Set providers = mOffers(key)
                    If Not providers.Exists(CStr(npc)) Then
                        providers(CStr(npc)) = True
                        links = links + 1
                        If links > MAX_INDEX_LINKS Then
                            mIndexOverflow = True
                            Call mOffers.RemoveAll()
                            Exit Sub
                        End If
                    End If
                End If
            Next offer
        End If
    Next npc
    Dim index As Variant
    For Each index In mGlobals.Keys
        mGlobalGates(index) = GlobalEligibilityGates(CInt(index))
    Next index
    If TraceDiscoveryNpc() > 0 Then Call TraceDiscovery("index npc=" & TraceDiscoveryNpc() & " quests=" & mOffers.Count & " overflow=" & mIndexOverflow)
    Exit Sub
Failed:
    ' Incomplete source definitions never publish a guessed partial index.
    mIndexOverflow = True
    Call mOffers.RemoveAll()
End Sub

Private Function GlobalEligibilityGates(ByVal index As Integer) As String
    On Error GoTo Unavailable
    Dim key As Variant, quests As Object, threshold As Long, gateOpen As Boolean, eventOpen As Boolean, eventChecked As Boolean
    Set quests = mGlobals(CStr(index))
    For Each key In quests.Keys
        threshold = QuestList(CInt(key)).GlobalQuestThresholdNeeded
        If threshold > 0 Then
            gateOpen = GlobalQuestInfo(index).GatheringGlobalCounter >= threshold
        Else
            If Not eventChecked Then
                eventOpen = False
                If GlobalQuestInfo(index).IsActive Then eventOpen = Not IsAnyGlobalQuestBossAlive(index)
                eventChecked = True
            End If
            gateOpen = eventOpen
        End If
        GlobalEligibilityGates = GlobalEligibilityGates & key & ":" & CStr(gateOpen) & ";"
    Next key
    Exit Function
Unavailable:
    GlobalEligibilityGates = "unavailable"
End Function

Public Function InvalidateQuestDiscoveryGlobal(ByVal index As Integer) As Boolean
    On Error GoTo Failed
    If mOffers Is Nothing Then Call BuildIndex()
    If mGlobals.Exists(CStr(index)) Then
        Dim gates As String
        gates = GlobalEligibilityGates(index)
        If Not mGlobalGates.Exists(CStr(index)) Then mGlobalGates(CStr(index)) = ""
        If mGlobalGates(CStr(index)) <> gates Then
            mGlobalGates(CStr(index)) = gates
            mRevision = mRevision + 1
            InvalidateQuestDiscoveryGlobal = True
        End If
    End If
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "modQuestDiscovery.InvalidateQuestDiscoveryGlobal", Erl)
End Function

' Distance belongs to the NPC interaction, not a map-independent discovery snapshot.
' Capacity remains the same FreeQuestSlot rule used by HandleQuestAccept.
Public Function IsQuestDiscoverable(ByVal userIndex As Integer, ByVal quest As Integer) As Boolean
    On Error GoTo Failed
    If quest <= 0 Or quest > UBound(QuestList) Then Exit Function
    If FreeQuestSlot(userIndex) = 0 Then Exit Function
    IsQuestDiscoverable = QuestAcceptanceFailure(userIndex, 0, quest, QuestList(quest)) = QuestAcceptAllowed
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "modQuestDiscovery.IsQuestDiscoverable", Erl)
End Function

Private Function InputSignature(ByVal userIndex As Integer) As String
    Dim slot As Integer, key As Variant
    With UserList(userIndex)
        InputSignature = .Stats.ELV & ":" & .clase & ":" & .Faccion.Status & ":" & .QuestStats.NumQuestsDone & ":" & mRevision & ";"
        For slot = 1 To MAXUSERQUESTS
            InputSignature = InputSignature & .QuestStats.Quests(slot).QuestIndex & ";"
        Next slot
        For Each key In mSkills.Keys
            InputSignature = InputSignature & key & ":" & .Stats.UserSkills(CInt(key)) & ";"
        Next key
    End With
End Function

' Called only by the existing dirty-user drain. False retains the same bounded
' retry/session guards as turn-in publication. Cache advances after ALL pages enqueue.
Public Function PublishQuestDiscovery(ByVal userIndex As Integer, Optional ByVal force As Boolean = False) As Boolean
    On Error GoTo Failed
    If mOffers Is Nothing Then Call BuildIndex()
    If userIndex > mCapacity Then
        mCapacity = UBound(UserList): ReDim Preserve mCache(1 To mCapacity)
    End If
    Dim inputs As String, fingerprint As String, rows As New Collection
    Dim key As Variant, npc As Variant, overflow As Boolean, providers As Object
    inputs = InputSignature(userIndex)
    With mCache(userIndex)
        If .userId <> UserList(userIndex).Id Or .version <> UserList(userIndex).VersionId Then
            .known = False: .sequence = 0
            Set .pendingRows = Nothing
            .userId = UserList(userIndex).Id: .version = UserList(userIndex).VersionId
        End If
        If .pendingRows Is Nothing Then
            If .known And Not force And .inputs = inputs Then
                PublishQuestDiscovery = True
                Exit Function
            End If
        ElseIf .pendingInputs <> inputs Then
            Set .pendingRows = Nothing
        End If
        If .pendingRows Is Nothing Then
            #If UNIT_TEST = 1 Then
                DiscoveryEvaluationCount = DiscoveryEvaluationCount + 1
            #End If
            overflow = mIndexOverflow
            If Not overflow And FreeQuestSlot(userIndex) > 0 Then
                For Each key In mOffers.Keys
                    If TraceDiscoveryNpc() > 0 Then
                        Set providers = mOffers(key)
                        If providers.Exists(CStr(TraceDiscoveryNpc())) Then Call TraceDiscovery("eligibility user=" & userIndex & " npc=" & TraceDiscoveryNpc() & " quest=" & key & " failure=" & QuestAcceptanceFailure(userIndex, 0, CInt(key), QuestList(CInt(key))) & " free-slot=" & FreeQuestSlot(userIndex))
                    End If
                    If IsQuestDiscoverable(userIndex, CInt(key)) Then
                        Set providers = mOffers(key)
                        For Each npc In providers.Keys
                            Call rows.Add(Array(CInt(npc), CInt(key)))
                            fingerprint = fingerprint & npc & ":" & key & ":" & Left$(QuestList(CInt(key)).nombre, 128) & ";"
                            If rows.Count > QUEST_DISCOVERY_LIMIT Then
                                overflow = True
                                Exit For
                            End If
                        Next npc
                    End If
                    If overflow Then Exit For
                Next key
            End If
            If overflow Then
                Set rows = New Collection
                fingerprint = "unavailable"
            End If
            If .known And Not force And .fingerprint = fingerprint Then
                .inputs = inputs: PublishQuestDiscovery = True
                Exit Function
            End If
            .sequence = .sequence + 1
            Set .pendingRows = rows
            .pendingInputs = inputs: .pendingFingerprint = fingerprint
            .pendingUnavailable = overflow: .nextRow = 1: .nextPage = 0
        End If
        Dim first As Long, last As Long, i As Long, flags As Byte, row As Variant, pages As Integer
        Dim writer As Network.Writer, queued As Boolean
        Do
            first = .nextRow
            last = first + QUEST_DISCOVERY_PAGE_SIZE - 1
            If last > .pendingRows.Count Then last = .pendingRows.Count
            flags = 0
            If last = .pendingRows.Count Then flags = FLAG_FINAL
            If .pendingUnavailable Then flags = flags Or FLAG_UNAVAILABLE
            Set writer = Protocol_Writes.GetWriterBuffer()
            Call writer.WriteInt16(ServerPacketID.eQuestDiscovery)
            Call writer.WriteInt32(.sequence)
            Call writer.WriteInt16(.nextPage)
            Call writer.WriteInt8(flags)
            Call writer.WriteInt16(last - first + 1)
            For i = first To last
                row = .pendingRows(i)
                Call writer.WriteInt16(CInt(row(0)))
                Call writer.WriteInt16(CInt(row(1)))
                Call writer.WriteString8(Left$(QuestList(CInt(row(1))).nombre, 128))
            Next i
            Call modSendData.SendData(ToIndex, userIndex, , False, queued, True)
            If Not queued Then
                If TraceDiscoveryNpc() > 0 Then Call TraceDiscovery("enqueue-failed user=" & userIndex & " sequence=" & .sequence & " page=" & .nextPage)
                Exit Function
            End If
            .nextRow = last + 1: .nextPage = .nextPage + 1: pages = pages + 1
            If .nextRow > .pendingRows.Count Then Exit Do
            If pages >= MAX_PAGES_PER_DRAIN Then Exit Function
        Loop
        If .pendingUnavailable Then Call TraceError(6, "Discovery capacity exceeded; unavailable replacement sent", "modQuestDiscovery.PublishQuestDiscovery", 0)
        .inputs = .pendingInputs: .fingerprint = .pendingFingerprint: .known = True
        If TraceDiscoveryNpc() > 0 Then Call TraceDiscovery("committed user=" & userIndex & " sequence=" & .sequence & " rows=" & .pendingRows.Count & " unavailable=" & .pendingUnavailable & " free-slot=" & FreeQuestSlot(userIndex))
        Set .pendingRows = Nothing
        .pendingInputs = "": .pendingFingerprint = ""
    End With
    PublishQuestDiscovery = True
    Exit Function
Failed:
    If Not writer Is Nothing Then Call writer.Clear()
    Call TraceError(Err.Number, Err.Description, "modQuestDiscovery.PublishQuestDiscovery", Erl)
End Function
