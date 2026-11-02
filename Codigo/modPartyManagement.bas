Attribute VB_Name = "modPartyManagement"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Public Enum e_PartyResult
    PartyState = 0
    PartySuccess = 1
    PartyNotMember = 2
    PartyNotLeader = 3
    PartyAlreadyMember = 4
    PartyUnavailable = 5
    PartyBlocked = 6
    PartyFull = 7
    PartyQuestionPending = 8
    PartyExpired = 9
    PartyIneligible = 10
    PartyStale = 11
    PartyDead = 12
    PartyLevel = 13
    PartyFaction = 14
    PartyGameMaster = 15
    PartyDeclined = 16
End Enum
Private Type PartyRequestCache
    userId As Long
    version As Integer
    request As Long
    fingerprint As String
    result As e_PartyResult
End Type
Private mRequests() As PartyRequestCache
Private mCapacity As Integer

Public Function PartyLeaderIndex(ByVal userIndex As Integer) As Integer
    On Error GoTo InvalidParty
    If userIndex < 1 Or userIndex > UBound(UserList) Then Exit Function
    With UserList(userIndex).Grupo
        If Not .EnGrupo Or .Id <= 0 Or Not IsValidUserRef(.Lider) Then Exit Function
        Dim leader As Integer, i As Integer, found As Boolean
        leader = .Lider.ArrayIndex
        If Not UserList(leader).Grupo.EnGrupo Or UserList(leader).Grupo.Id <> .Id Then Exit Function
        If Not IsValidUserRef(UserList(leader).Grupo.Lider) Then Exit Function
        If UserList(leader).Grupo.Lider.ArrayIndex <> leader Then Exit Function
        If UserList(leader).Grupo.CantidadMiembros < 1 Or UserList(leader).Grupo.CantidadMiembros > 6 Then Exit Function
        For i = 1 To UserList(leader).Grupo.CantidadMiembros
            If IsValidUserRef(UserList(leader).Grupo.Miembros(i)) Then
                If UserList(leader).Grupo.Miembros(i).ArrayIndex = userIndex Then found = True
            End If
        Next i
        If found Then PartyLeaderIndex = leader
    End With
InvalidParty:
End Function

Public Function PartyHasMultipleMembers(ByVal userIndex As Integer) As Boolean
    Dim leader As Integer
    leader = PartyLeaderIndex(userIndex)
    If leader = 0 Then Exit Function
    Dim i As Integer, j As Integer, member As Integer, count As Integer, seen(1 To 6) As Integer, duplicate As Boolean
    For i = 1 To UserList(leader).Grupo.CantidadMiembros
        If IsValidUserRef(UserList(leader).Grupo.Miembros(i)) Then
            member = UserList(leader).Grupo.Miembros(i).ArrayIndex
            If IsRewardCredit(member) And PartyLeaderIndex(member) = leader Then
                duplicate = False
                For j = 1 To count
                    If seen(j) = member Then duplicate = True
                Next j
                If Not duplicate Then
                    count = count + 1: seen(count) = member
                End If
            End If
        End If
    Next i
    PartyHasMultipleMembers = count >= 2
End Function

Public Sub ClearPartyProposal(ByVal userIndex As Integer)
    On Error GoTo ClearPartyProposal_Err
    If userIndex < 1 Or userIndex > UBound(UserList) Then Exit Sub
    Call ClearUserRef(UserList(userIndex).Grupo.PropuestaDe)
    UserList(userIndex).Grupo.ProposalPartyId = 0
    If UserList(userIndex).flags.pregunta = 1 Then
        UserList(userIndex).flags.pregunta = 0
        UserList(userIndex).flags.RespondiendoPregunta = False
    End If
    Exit Sub
ClearPartyProposal_Err:
    Call TraceError(Err.Number, Err.Description, "modPartyManagement.ClearPartyProposal", Erl)
End Sub

Public Sub WritePartyManagementState(ByVal userIndex As Integer, Optional ByVal request As Long = 0, Optional ByVal result As e_PartyResult = PartyState)
    On Error GoTo WritePartyManagementState_Err
    Call SendData(ToIndex, userIndex, PrepareUpdateGroupInfo(userIndex, request, result))
    Exit Sub
WritePartyManagementState_Err:
    Call TraceError(Err.Number, Err.Description, "modPartyManagement.WritePartyManagementState", Erl)
End Sub

Public Sub PublishManagedParty(ByVal leader As Integer)
    On Error GoTo PublishManagedParty_Err
    Dim i As Integer, member As Integer
    For i = 1 To UserList(leader).Grupo.CantidadMiembros
        If IsValidUserRef(UserList(leader).Grupo.Miembros(i)) Then
            member = UserList(leader).Grupo.Miembros(i).ArrayIndex
            If IsRewardCredit(member) And PartyLeaderIndex(member) = leader Then
                Call WritePartyManagementState(member)
                Call CompartirUbicacion(member)
                Call RefreshCharStatus(member)
            End If
        End If
    Next i
    Exit Sub
PublishManagedParty_Err:
    Call TraceError(Err.Number, Err.Description, "modPartyManagement.PublishManagedParty", Erl)
End Sub

' Same predicate for name/target invitations and acceptance. No mutation here.
Public Function PartyInviteEligibility(ByVal inviter As Integer, ByVal invited As Integer) As e_PartyResult
    PartyInviteEligibility = PartyUnavailable
    If Not IsRewardCredit(inviter) Or Not IsRewardCredit(invited) Then Exit Function
    If inviter = invited Then Exit Function
    If UserList(inviter).flags.Muerto <> 0 Or UserList(invited).flags.Muerto <> 0 Then
        PartyInviteEligibility = PartyDead: Exit Function
    End If
    If UserList(inviter).pos.Map <> UserList(invited).pos.Map Then Exit Function
    If Distancia(UserList(inviter).pos, UserList(invited).pos) > 8 Then Exit Function
    If Not CanUserSeeAdminInvisible(inviter, invited) Then Exit Function
    If EsGM(inviter) Or EsGM(invited) Then
        PartyInviteEligibility = PartyGameMaster: Exit Function
    End If
    If UserList(invited).Grupo.EnGrupo Then
        PartyInviteEligibility = PartyAlreadyMember: Exit Function
    End If
    If UserList(invited).flags.SeguroParty Then
        PartyInviteEligibility = PartyBlocked: Exit Function
    End If
    Dim firstStatus As Integer, secondStatus As Integer, levelLimit As Integer
    firstStatus = Status(inviter): secondStatus = Status(invited)
    If (firstStatus = 0 And (secondStatus = 1 Or secondStatus = 3)) Or _
       (firstStatus = 1 And (secondStatus = 0 Or secondStatus = 2)) Or _
       (firstStatus = 2 And secondStatus = 1) Or (firstStatus = 3 And (secondStatus = 0 Or secondStatus = 2)) Or _
       (firstStatus = 2 And secondStatus = 3) Then
        PartyInviteEligibility = PartyFaction: Exit Function
    End If
    levelLimit = SvrConfig.GetValue("PartyELV")
    With UserList(inviter).Stats
        If CInt(.UserSkills(e_Skill.liderazgo)) >= (15 - .UserAtributos(e_Atributos.Carisma) / 2) Then levelLimit = SvrConfig.GetValue("PartyELVwLeadership")
    End With
    If Abs(CInt(UserList(inviter).Stats.ELV) - CInt(UserList(invited).Stats.ELV)) > levelLimit Then
        PartyInviteEligibility = PartyLevel: Exit Function
    End If
    PartyInviteEligibility = PartySuccess
End Function

Public Function InviteManagedParty(ByVal inviter As Integer, ByVal invited As Integer) As e_PartyResult
    On Error GoTo Failed
    InviteManagedParty = PartyInviteEligibility(inviter, invited)
    If InviteManagedParty <> PartySuccess Then Exit Function
    If Not UserList(inviter).Grupo.EnGrupo Then Call GroupCreateSuccess(inviter)
    If PartyLeaderIndex(inviter) <> inviter Then
        InviteManagedParty = PartyNotLeader: Exit Function
    End If
    If UserList(inviter).Grupo.CantidadMiembros >= 6 Then
        InviteManagedParty = PartyFull: Exit Function
    End If
    If UserList(invited).flags.RespondiendoPregunta Then
        InviteManagedParty = PartyQuestionPending: Exit Function
    End If
    With UserList(invited)
        Call SetUserRef(.Grupo.PropuestaDe, inviter)
        .Grupo.ProposalPartyId = UserList(inviter).Grupo.Id
        .flags.pregunta = 1: .flags.RespondiendoPregunta = True
    End With
    ' Confirm the existing question enqueued before reporting invitation success.
    Dim writer As Network.Writer, queued As Boolean
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eShowPregunta)
    Call writer.WriteInt16(MSG_TE_INVITO_A_UNIRSE_A_SU_GRUPO_DESEAS_UNIRTE)
    Call writer.WriteString8(GetUserDisplayName(inviter))
    Call SendData(ToIndex, invited, , False, queued, True)
    If Not queued Then
        Call ClearPartyProposal(invited): InviteManagedParty = PartyUnavailable
    End If
    Exit Function
Failed:
    InviteManagedParty = PartyUnavailable
    Call TraceError(Err.Number, Err.Description, "modPartyManagement.InviteManagedParty", Erl)
End Function

Public Function AcceptManagedParty(ByVal invited As Integer) As e_PartyResult
    On Error GoTo Failed
    AcceptManagedParty = PartyExpired
    Dim inviter As Integer, i As Integer
    If Not IsValidUserRef(UserList(invited).Grupo.PropuestaDe) Then GoTo Finished
    inviter = UserList(invited).Grupo.PropuestaDe.ArrayIndex
    If PartyLeaderIndex(inviter) <> inviter Then GoTo Finished
    If UserList(inviter).Grupo.Id <> UserList(invited).Grupo.ProposalPartyId Then GoTo Finished
    AcceptManagedParty = PartyInviteEligibility(inviter, invited)
    If AcceptManagedParty <> PartySuccess Then GoTo Finished
    With UserList(inviter).Grupo
        If .CantidadMiembros >= 6 Then
            AcceptManagedParty = PartyFull: GoTo Finished
        End If
        For i = 1 To .CantidadMiembros
            If .Miembros(i).ArrayIndex = invited Then
                AcceptManagedParty = PartyAlreadyMember: GoTo Finished
            End If
        Next i
        .CantidadMiembros = .CantidadMiembros + 1
        Call SetUserRef(.Miembros(.CantidadMiembros), invited)
        .RosterRevision = .RosterRevision + 1
    End With
    Call ClearPartyXpRotation(invited)
    With UserList(invited).Grupo
        .EnGrupo = True: .Id = UserList(inviter).Grupo.Id
        Call SetUserRef(.Lider, inviter)
    End With
    Call PublishManagedParty(inviter)
    Call WritePartyManagementState(inviter, 0, PartySuccess)
Finished:
    Call ClearPartyProposal(invited)
    Call WritePartyManagementState(invited, 0, AcceptManagedParty)
    Exit Function
Failed:
    AcceptManagedParty = PartyExpired
    Call ClearPartyProposal(invited)
    Call TraceError(Err.Number, Err.Description, "modPartyManagement.AcceptManagedParty", Erl)
End Function

Private Sub EvictLonePartyFromRestrictedMap(ByVal userIndex As Integer)
    If Not IsRewardCredit(userIndex) Then Exit Sub
    With UserList(userIndex)
        If .pos.Map <= 0 Or .pos.Map > NumMaps Then Exit Sub
        If MapInfo(.pos.Map).OnlyGroups And Not PartyHasMultipleMembers(userIndex) Then
            If MapInfo(.pos.Map).Salida.Map > 0 Then Call WarpUserChar(userIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
        End If
    End With
End Sub

Public Sub RemoveManagedPartyMember(ByVal member As Integer)
    On Error GoTo Failed
    Dim leader As Integer, i As Integer, found As Integer
    leader = PartyLeaderIndex(member)
    If leader = 0 Then Exit Sub
    If leader = member Then
        Call DisbandManagedParty(leader): Exit Sub
    End If
    With UserList(leader).Grupo
        For i = 1 To .CantidadMiembros
            If .Miembros(i).ArrayIndex = member Then found = i
        Next i
        If found = 0 Then Exit Sub
        For i = found To .CantidadMiembros - 1
            .Miembros(i) = .Miembros(i + 1)
        Next i
        Call ClearUserRef(.Miembros(.CantidadMiembros))
        .CantidadMiembros = .CantidadMiembros - 1: .RosterRevision = .RosterRevision + 1
    End With
    Call ClearManagedMembership(member)
    If IsRewardCredit(member) Then
        Call WritePartyManagementState(member): Call RefreshCharStatus(member)
    End If
    Call PublishManagedParty(leader)
    Call EvictLonePartyFromRestrictedMap(member)
    Call EvictLonePartyFromRestrictedMap(leader)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyManagement.RemoveManagedPartyMember", Erl)
End Sub

Private Sub ClearManagedMembership(ByVal member As Integer)
    Call ClearPartyXpRotation(member)
    Call ClearPartyProposal(member)
    With UserList(member).Grupo
        .EnGrupo = False: .Id = -1: .CantidadMiembros = 0: .RosterRevision = 0
        Call ClearUserRef(.Lider)
        Dim i As Integer
        For i = 1 To 6
            Call ClearUserRef(.Miembros(i))
        Next i
    End With
End Sub

Public Sub DisbandManagedParty(ByVal leader As Integer)
    On Error GoTo Failed
    If PartyLeaderIndex(leader) <> leader Or leader = 0 Then Exit Sub
    Dim members(1 To 6) As t_UserReference, i As Integer, count As Integer, member As Integer
    count = UserList(leader).Grupo.CantidadMiembros
    For i = 1 To count
        members(i) = UserList(leader).Grupo.Miembros(i)
    Next i
    Dim partyId As Long: partyId = UserList(leader).Grupo.Id
    For i = 1 To count
        If IsValidUserRef(members(i)) Then
            member = members(i).ArrayIndex
            If UserList(member).Grupo.Id = partyId Then
                Call ClearManagedMembership(member)
                If IsRewardCredit(member) Then
                    Call WritePartyManagementState(member): Call RefreshCharStatus(member)
                    Call EvictLonePartyFromRestrictedMap(member)
                End If
            End If
        End If
    Next i
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyManagement.DisbandManagedParty", Erl)
End Sub

Public Sub HandlePartyManagement(ByVal userIndex As Integer, ByVal reader As Network.Reader)
    On Error GoTo Failed
    Dim request As Long, action As Byte, partyId As Long, revision As Long, target As t_UserReference, name As String
    request = reader.ReadInt32(): action = reader.ReadInt8()
    partyId = reader.ReadInt32(): revision = reader.ReadInt32()
    target.ArrayIndex = reader.ReadInt16(): target.VersionId = reader.ReadInt16(): name = reader.ReadString8()
    If reader.GetAvailable() <> 0 Or request <= 0 Or action > 5 Or Len(name) > 30 Then Exit Sub
    Call ExecutePartyManagement(userIndex, request, action, partyId, revision, target, name)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyManagement.HandlePartyManagement", Erl)
End Sub

Public Sub ExecutePartyManagement(ByVal userIndex As Integer, ByVal request As Long, ByVal action As Byte, ByVal partyId As Long, ByVal revision As Long, ByRef target As t_UserReference, ByVal name As String)
    On Error GoTo Failed
    If Not IsRewardCredit(userIndex) Or request <= 0 Then Exit Sub
    If userIndex > mCapacity Then
        mCapacity = UBound(UserList): ReDim Preserve mRequests(1 To mCapacity)
    End If
    Dim result As e_PartyResult, fingerprint As String, emptyCache As PartyRequestCache, leader As Integer, invited As Integer
    fingerprint = action & ":" & partyId & ":" & revision & ":" & target.ArrayIndex & ":" & target.VersionId & ":" & name
    With mRequests(userIndex)
        If .userId <> UserList(userIndex).Id Or .version <> UserList(userIndex).VersionId Then
            .userId = UserList(userIndex).Id: .version = UserList(userIndex).VersionId: .request = 0
        End If
        If request <= .request Then
            result = PartyStale
            If request = .request And fingerprint = .fingerprint Then result = .result
            Call WritePartyManagementState(userIndex, request, result): Exit Sub
        End If
        result = PartySuccess: leader = PartyLeaderIndex(userIndex)
        If action = 0 Then
            result = PartyState
        ElseIf action = 1 Then
            If UserList(userIndex).Grupo.EnGrupo Then
                result = PartyAlreadyMember
            ElseIf Not IsRewardUser(userIndex) Then
                result = PartyDead
            ElseIf EsGM(userIndex) Then
                result = PartyGameMaster
            Else
                Call GroupCreateSuccess(userIndex)
                If PartyLeaderIndex(userIndex) = 0 Then result = PartyUnavailable
            End If
        ElseIf leader = 0 Then
            result = PartyNotMember
        ElseIf partyId <> UserList(leader).Grupo.Id Or revision <> UserList(leader).Grupo.RosterRevision Then
            result = PartyStale
        ElseIf action = 4 Then
            If leader = userIndex Then
                result = PartyNotLeader
            Else
                Call RemoveManagedPartyMember(userIndex)
            End If
        ElseIf leader <> userIndex Then
            result = PartyNotLeader
        ElseIf action = 2 Then
            Dim resolved As t_UserReference
            name = Trim$(name): resolved = NameIndex(name)
            If IsValidUserRef(resolved) Then invited = resolved.ArrayIndex
            result = InviteManagedParty(userIndex, invited)
        ElseIf action = 3 Then
            If Not IsValidUserRef(target) Then
                result = PartyStale
            ElseIf target.ArrayIndex = userIndex Or PartyLeaderIndex(target.ArrayIndex) <> leader Then
                result = PartyStale
            Else
                Call RemoveManagedPartyMember(target.ArrayIndex)
            End If
        ElseIf action = 5 Then
            Call DisbandManagedParty(userIndex)
        Else
            result = PartyStale
        End If
        .request = request: .fingerprint = fingerprint: .result = result
    End With
    Call WritePartyManagementState(userIndex, request, result)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "modPartyManagement.ExecutePartyManagement", Erl)
End Sub
