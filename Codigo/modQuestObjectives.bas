Attribute VB_Name = "modQuestObjectives"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit

' Positive template = kill; negative template = click-to-visit.
' Five accepted quests, two byte-sized objective lists per quest.
Public Const MAX_QUEST_OBJECTIVE_TEMPLATES As Integer = 2550
Private Type ObjectiveSnapshotCache
    known As Boolean
    userId As Long
    userVersion As Integer
    fingerprint As String
End Type
Private mCache() As ObjectiveSnapshotCache
Private mCapacity As Long
Private Type ObjectiveDetail
    quest As Integer
    npc As Integer
    progress As Integer
    required As Integer
End Type

Public Sub CollectQuestObjectiveTemplates(ByVal userIndex As Integer, ByRef templates() As Integer, ByRef count As Integer)
    Dim details() As ObjectiveDetail, detailCount As Integer
    Call CollectQuestObjectiveData(userIndex, templates, count, details, detailCount)
End Sub

Private Sub CollectQuestObjectiveData(ByVal userIndex As Integer, ByRef templates() As Integer, ByRef count As Integer, ByRef details() As ObjectiveDetail, ByRef detailCount As Integer)
    On Error GoTo CollectQuestObjectiveTemplates_Err
    count = 0
    detailCount = 0
    ReDim details(1 To MAX_QUEST_OBJECTIVE_TEMPLATES)
    ReDim templates(1 To MAX_QUEST_OBJECTIVE_TEMPLATES)
    Dim seen As New Dictionary
    Dim slot As Integer, quest As Integer, objective As Integer
    Dim npc As Integer, progress As Integer, required As Integer, kind As Integer
    For slot = 1 To MAXUSERQUESTS
        quest = UserList(userIndex).QuestStats.Quests(slot).QuestIndex
        If quest > 0 And quest <= UBound(QuestList) Then
            For kind = 1 To 2
                Dim objectives As Integer
                If kind = 1 Then
                    objectives = QuestList(quest).RequiredNPCs
                Else
                    objectives = QuestList(quest).RequiredTargetNPCs
                End If
                For objective = 1 To objectives
                    If kind = 1 Then
                        npc = QuestList(quest).RequiredNPC(objective).NpcIndex
                        required = QuestList(quest).RequiredNPC(objective).amount
                        progress = UserList(userIndex).QuestStats.Quests(slot).NPCsKilled(objective)
                    Else
                        npc = QuestList(quest).RequiredTargetNPC(objective).NpcIndex
                        required = QuestList(quest).RequiredTargetNPC(objective).amount
                        progress = UserList(userIndex).QuestStats.Quests(slot).NPCsTarget(objective)
                    End If
                    ' Same exact template/progress conditions as kill/visit and inspection.
                    If npc > 0 And progress >= 0 And progress < required Then
                        If kind = 2 Then npc = -npc
                        detailCount = detailCount + 1
                        details(detailCount).quest = quest
                        details(detailCount).npc = npc
                        details(detailCount).progress = progress
                        details(detailCount).required = required
                        If Not seen.Exists(CStr(npc)) Then
                            Call seen.Add(CStr(npc), True)
                            count = count + 1
                            templates(count) = npc
                        End If
                    End If
                Next objective
            Next kind
        End If
    Next slot
    Exit Sub
CollectQuestObjectiveTemplates_Err:
    ' Incomplete metadata must not leave a guessed partial projection.
    count = 0
    detailCount = 0
    Call TraceError(Err.Number, Err.Description, "modQuestObjectives.CollectQuestObjectiveTemplates", Erl)
End Sub

Public Sub SendQuestObjectiveSnapshot(ByVal userIndex As Integer, Optional ByVal force As Boolean = False)
    On Error GoTo SendQuestObjectiveSnapshot_Err
    If userIndex < 1 Or userIndex > UBound(UserList) Then Exit Sub
    If Not UserList(userIndex).flags.UserLogged Then Exit Sub
    Call SendQuestTurnInSnapshot(userIndex, force)
    If userIndex > mCapacity Then
        mCapacity = UBound(UserList)
        ReDim Preserve mCache(1 To mCapacity)
    End If
    Dim templates() As Integer, count As Integer, i As Integer, fingerprint As String
    Dim details() As ObjectiveDetail, detailCount As Integer
    Call CollectQuestObjectiveData(userIndex, templates, count, details, detailCount)
    Dim quests As New Dictionary, questKey As Variant, questCount As Integer
    For i = 1 To detailCount
        With details(i)
            If Not quests.Exists(CStr(.quest)) Then Call quests.Add(CStr(.quest), Left$(QuestList(.quest).nombre, 128))
            fingerprint = fingerprint & .quest & ":" & .npc & ":" & .progress & ":" & .required & ";"
        End With
    Next i
    For Each questKey In quests.Keys
        fingerprint = fingerprint & Len(CStr(quests(questKey))) & ":" & CStr(quests(questKey))
    Next questKey
    With mCache(userIndex)
        If Not force And .known And .userId = UserList(userIndex).Id And _
            .userVersion = UserList(userIndex).VersionId And .fingerprint = fingerprint Then Exit Sub
        Dim writer As Network.Writer, queued As Boolean
        Set writer = Protocol_Writes.GetWriterBuffer()
        Call writer.WriteInt16(ServerPacketID.eQuestObjectiveTemplates)
        Call writer.WriteInt16(count)
        For i = 1 To count
            Call writer.WriteInt16(templates(i))
        Next i
        ' One title per accepted quest; bounded to five titles and 2550 objectives.
        Call writer.WriteInt8(quests.Count)
        For Each questKey In quests.Keys
            questCount = 0
            For i = 1 To detailCount
                If details(i).quest = CInt(questKey) Then questCount = questCount + 1
            Next i
            Call writer.WriteInt16(CInt(questKey))
            Call writer.WriteString8(CStr(quests(questKey)))
            Call writer.WriteInt16(questCount)
            For i = 1 To detailCount
                If details(i).quest = CInt(questKey) Then
                    Call writer.WriteInt16(details(i).npc)
                    Call writer.WriteInt16(details(i).progress)
                    Call writer.WriteInt16(details(i).required)
                End If
            Next i
        Next questKey
        Call modSendData.SendData(ToIndex, userIndex, , False, queued, True)
        If queued Then
            .known = True
            .userId = UserList(userIndex).Id
            .userVersion = UserList(userIndex).VersionId
            .fingerprint = fingerprint
        End If
    End With
    Exit Sub
SendQuestObjectiveSnapshot_Err:
    If Not writer Is Nothing Then Call writer.Clear()
    Call TraceError(Err.Number, Err.Description, "modQuestObjectives.SendQuestObjectiveSnapshot", Erl)
End Sub
