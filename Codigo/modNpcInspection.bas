Attribute VB_Name = "modNpcInspection"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit

Private Const INSPECTION_INTERVAL As Long = 250
Private Const MAX_ENTRIES As Integer = 32
Private Type InspectionSubscription
    request As Long
    userVersion As Integer
    userId As Long
    npcIndex As Integer
    npcVersion As Integer
    viewerMap As Integer
    characterIndex As Integer
    npcNumber As Integer
    fingerprint As String
End Type
Private mSubscriptions() As InspectionSubscription
Private mCapacity As Long
Private mLastUpdate As Long

' Physical pressure estimate, not a combat simulation. Unknown spell damage is
' deliberately not rated. Current HP, buffs, random rolls and temporary damage
' modifiers do not enter this estimate.
Public Function InspectionThreatBand(ByVal maximumHp As Long, ByVal maximumHit As Double, ByVal intervalMs As Long) As Byte
    If maximumHp <= 0 Or maximumHit <= 0 Or intervalMs <= 0 Then Exit Function
    Dim pressure As Double
    pressure = maximumHit * 2000# / intervalMs / maximumHp
    If pressure >= 0.3 Then
        InspectionThreatBand = 4
    ElseIf pressure >= 0.15 Then
        InspectionThreatBand = 3
    ElseIf pressure >= 0.06 Then
        InspectionThreatBand = 2
    Else
        InspectionThreatBand = 1
    End If
End Function

Private Function AverageArmour(ByVal objectIndex As Integer) As Double
    If objectIndex <= 0 Or objectIndex > UBound(ObjData) Then Exit Function
    AverageArmour = (CDbl(ObjData(objectIndex).MinDef) + ObjData(objectIndex).MaxDef) / 2#
End Function

Private Function ThreatFor(ByVal userIndex As Integer, ByVal npcIndex As Integer) As Byte
    With NpcList(npcIndex)
        If .Hostile = 0 Or .Attackable = 0 Or .MaestroUser.ArrayIndex <> 0 Then Exit Function
        If .flags.LanzaSpells > 0 Or .Stats.MaxHit <= 0 Or .Stats.MaxHp <= 0 Or .IntervaloAtaque <= 0 Then Exit Function
        Dim armour As Double, damage As Double
        With UserList(userIndex).invent
            armour = (AverageArmour(.EquippedHelmetObjIndex) + 5# * (AverageArmour(.EquippedArmorObjIndex) + AverageArmour(.EquippedShieldObjIndex))) / 6#
        End With
        damage = .Stats.MaxHit - armour
        ' An absorbed physical hit still has a minimum nonzero estimate; this
        ' deliberately does not claim immunity or account for poison.
        If damage < 1 Then damage = 1
        ThreatFor = InspectionThreatBand(UserList(userIndex).Stats.MaxHp, damage, .IntervaloAtaque)
    End With
End Function

Public Sub SubscribeNpcInspection(ByVal userIndex As Integer, ByVal request As Long, ByVal viewerMap As Integer, ByVal characterIndex As Integer, ByVal npcNumber As Integer)
    On Error GoTo Failed
    If mCapacity <> UBound(UserList) Then
        mCapacity = UBound(UserList)
        ReDim mSubscriptions(1 To mCapacity)
    End If
    If request <= 0 Then Exit Sub
    With mSubscriptions(userIndex)
        .request = 0
        If characterIndex = 0 Then Exit Sub
        If Not UserList(userIndex).flags.UserLogged Then Exit Sub
        If UserList(userIndex).pos.Map <> viewerMap Then Exit Sub
        Dim npcIndex As Integer
        npcIndex = ResolveHooTargetedSpellNpc(characterIndex)
        If npcIndex = 0 Then Exit Sub
        If NpcList(npcIndex).Numero <> npcNumber Then Exit Sub
        .request = request
        .userVersion = UserList(userIndex).VersionId
        .userId = UserList(userIndex).Id
        .npcIndex = npcIndex
        .npcVersion = NpcList(npcIndex).VersionId
        .viewerMap = viewerMap
        .characterIndex = characterIndex
        .npcNumber = npcNumber
        .fingerprint = vbNullString
    End With
    Call RefreshNpcInspection(userIndex)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modNpcInspection.SubscribeNpcInspection", Erl)
End Sub

Public Sub UpdateNpcInspections(ByVal currentTime As Long)
    On Error GoTo Failed
    If TicksElapsed(mLastUpdate, currentTime) < INSPECTION_INTERVAL Then Exit Sub
    mLastUpdate = currentTime
    Dim userIndex As Integer
    For userIndex = 1 To mCapacity
        If mSubscriptions(userIndex).request > 0 Then Call RefreshNpcInspection(userIndex)
    Next userIndex
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modNpcInspection.UpdateNpcInspections", Erl)
End Sub

Public Sub RefreshNpcInspection(ByVal userIndex As Integer)
    On Error GoTo Failed
    If userIndex < 1 Or userIndex > mCapacity Then Exit Sub
    With mSubscriptions(userIndex)
        If .request = 0 Then Exit Sub
        If Not UserList(userIndex).flags.UserLogged Or UserList(userIndex).VersionId <> .userVersion Or UserList(userIndex).Id <> .userId Then
            .request = 0
            Exit Sub
        End If
        Dim valid As Boolean, npcIndex As Integer
        npcIndex = .npcIndex
        valid = NpcList(npcIndex).flags.NPCActive And NpcList(npcIndex).Stats.MinHp > 0 And NpcList(npcIndex).Char.body > 0 And _
            NpcList(npcIndex).VersionId = .npcVersion And _
            NpcList(npcIndex).Char.charindex = .characterIndex And NpcList(npcIndex).Numero = .npcNumber And _
            UserList(userIndex).pos.Map = .viewerMap And NpcList(npcIndex).pos.Map = .viewerMap
        If valid Then valid = NpcBowInLoadedArea(userIndex, npcIndex) And NpcList(npcIndex).flags.invisible = 0
        If valid And UserList(userIndex).flags.Muerto <> 0 Then valid = MapInfo(.viewerMap).Seguro <> 0 Or NpcList(npcIndex).npcType = e_NPCType.Revividor
        Dim threat As Byte, traits As Byte, count As Byte, truncated As Boolean
        Dim questIds(1 To MAX_ENTRIES) As Integer, kinds(1 To MAX_ENTRIES) As Byte
        Dim current(1 To MAX_ENTRIES) As Long, required(1 To MAX_ENTRIES) As Long
        Dim slot As Integer, quest As Integer, objective As Integer, canReturn As Boolean
        If valid Then
            threat = ThreatFor(userIndex, npcIndex)
            If NpcList(npcIndex).Veneno > 0 Then traits = traits Or 1
            If NpcList(npcIndex).AttackRange > 1 And NpcList(npcIndex).ProjectileType > 0 Then traits = traits Or 2
            If NpcList(npcIndex).flags.LanzaSpells > 0 Then traits = traits Or 4
            For slot = 1 To MAXUSERQUESTS
                quest = UserList(userIndex).QuestStats.Quests(slot).QuestIndex
                If quest > 0 Then
                    For objective = 1 To QuestList(quest).RequiredNPCs
                        If QuestList(quest).RequiredNPC(objective).NpcIndex = .npcNumber Then
                            If UserList(userIndex).QuestStats.Quests(slot).NPCsKilled(objective) < QuestList(quest).RequiredNPC(objective).amount Then
                                If count = MAX_ENTRIES Then
                                    truncated = True
                                    Exit For
                                End If
                                count = count + 1
                                questIds(count) = quest: kinds(count) = 1
                                current(count) = UserList(userIndex).QuestStats.Quests(slot).NPCsKilled(objective)
                                required(count) = QuestList(quest).RequiredNPC(objective).amount
                            End If
                        End If
                    Next objective
                    For objective = 1 To QuestList(quest).RequiredTargetNPCs
                        If QuestList(quest).RequiredTargetNPC(objective).NpcIndex = .npcNumber Then
                            If UserList(userIndex).QuestStats.Quests(slot).NPCsTarget(objective) < QuestList(quest).RequiredTargetNPC(objective).amount Then
                                If count = MAX_ENTRIES Then
                                    truncated = True
                                    Exit For
                                End If
                                count = count + 1: questIds(count) = quest: kinds(count) = 3
                                current(count) = UserList(userIndex).QuestStats.Quests(slot).NPCsTarget(objective)
                                required(count) = QuestList(quest).RequiredTargetNPC(objective).amount
                            End If
                        End If
                    Next objective
                    canReturn = QuestList(quest).TalkTo = .npcNumber
                    For objective = 1 To NpcList(npcIndex).NumQuest
                        If NpcList(npcIndex).QuestNumber(objective) = quest Then canReturn = True
                    Next objective
                    If canReturn Then
                        If CanInspectQuestTurnIn(userIndex, quest, CByte(slot)) Then
                            If count < MAX_ENTRIES Then
                                count = count + 1: questIds(count) = quest: kinds(count) = 2
                            Else
                                truncated = True
                            End If
                        End If
                    End If
                End If
            Next slot
        End If
        Dim fingerprint As String, entry As Integer
        fingerprint = CStr(valid) & "|" & threat & "|" & traits & "|" & truncated
        For entry = 1 To count
            fingerprint = fingerprint & "|" & questIds(entry) & ":" & kinds(entry) & ":" & current(entry) & ":" & required(entry)
        Next entry
        If fingerprint = .fingerprint Then Exit Sub
        Dim writer As Network.Writer, queued As Boolean
        Set writer = Protocol_Writes.GetWriterBuffer()
        Call writer.WriteInt16(ServerPacketID.eNpcInspection)
        Call writer.WriteInt32(.request)
        Call writer.WriteInt16(.viewerMap)
        Call writer.WriteInt16(.characterIndex)
        Call writer.WriteInt16(.npcNumber)
        Call writer.WriteInt8(Abs(CInt(valid)))
        Call writer.WriteInt8(threat)
        Call writer.WriteInt8(traits)
        Call writer.WriteInt8(count Or (Abs(CInt(truncated)) * 128))
        For entry = 1 To count
            Call writer.WriteInt16(questIds(entry))
            Call writer.WriteInt8(kinds(entry))
            Call writer.WriteInt32(current(entry))
            Call writer.WriteInt32(required(entry))
        Next entry
        Call modSendData.SendData(ToIndex, userIndex, , False, queued, True)
        If queued Then .fingerprint = fingerprint
        If Not valid And queued Then .request = 0
    End With
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modNpcInspection.RefreshNpcInspection", Erl)
End Sub
