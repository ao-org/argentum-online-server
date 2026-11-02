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
End Type
Private mCache() As TurnInCache
Private mCapacity As Long

Public Sub RefreshGlobalQuestTurnIns(ByVal globalQuest As Integer)
    On Error GoTo Failed
    If LastUser = 0 Then Exit Sub
    Dim userIndex As Long, slot As Integer, quest As Integer
    For userIndex = 1 To LastUser
        If UserList(userIndex).flags.UserLogged Then
            For slot = 1 To MAXUSERQUESTS
                quest = UserList(userIndex).QuestStats.Quests(slot).QuestIndex
                If quest > 0 And quest <= UBound(QuestList) Then
                    If QuestList(quest).GlobalQuestIndex = globalQuest Then
                        Call SendQuestTurnInSnapshot(CInt(userIndex))
                        Exit For
                    End If
                End If
            Next slot
        End If
    Next userIndex
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modQuestTurnIn.RefreshGlobalQuestTurnIns", Erl)
End Sub

Public Sub ResetQuestTurnInRecipients()
    Set mRecipients = Nothing
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
Public Sub SendQuestTurnInSnapshot(ByVal userIndex As Integer, Optional ByVal force As Boolean = False)
    On Error GoTo Failed
    If userIndex < 1 Or userIndex > UBound(UserList) Then Exit Sub
    If Not UserList(userIndex).flags.UserLogged Then Exit Sub
    If userIndex > mCapacity Then
        mCapacity = UBound(UserList)
        ReDim Preserve mCache(1 To mCapacity)
    End If
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
            .version = UserList(userIndex).VersionId And .fingerprint = fingerprint Then Exit Sub
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
            .known = True: .userId = UserList(userIndex).Id
            .version = UserList(userIndex).VersionId: .fingerprint = fingerprint
        End If
    End With
    Exit Sub
Failed:
    If Not writer Is Nothing Then Call writer.Clear()
    Call TraceError(Err.Number, Err.Description, "modQuestTurnIn.SendQuestTurnInSnapshot", Erl)
End Sub

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
