Attribute VB_Name = "Unit_PartyManagement"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private recipients As Collection
Private fixtureVersion As Integer
Private rejectQuestion As Boolean
Public Function CapturePacket(ByVal userIndex As Integer, ByRef queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    Dim writer As Network.Writer, data() As Byte, payload As Variant, id As Long
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(data): Call writer.Clear()
    id = CLng(data(0)) + 256& * data(1)
    queued = Not (rejectQuestion And id = ServerPacketID.eShowPregunta)
    If queued Then
        payload = data: Call packets.Add(payload): Call recipients.Add(userIndex)
    End If
    CapturePacket = True
End Function
Private Sub Expect(ByVal value As Boolean, ByVal label As String)
    If Not value Then Err.Raise vbObjectError + 853, "Party management", label
End Sub
Private Sub Fixture()
    Call Unit_NpcBow.InitializeTestWorld()
    ReDim UserList(1 To 7)
    Set m_NameIndex = New Dictionary
    ' Leave one generation for cases that deliberately reconnect inside a fixture.
    fixtureVersion = fixtureVersion + 2
    Dim i As Integer, wrapper As clsUserRefWrapper
    For i = 1 To 7
        With UserList(i)
            .AccountID = 500 + i: .Id = 500 + i: .VersionId = fixtureVersion: .Name = "Party" & i
            .flags.UserLogged = True: .ConnectionDetails.ConnIDValida = True
            .flags.Privilegios = e_PlayerType.User
            .Stats.ELV = 10: .Stats.MinHp = 100: .Stats.MaxHp = 100
            .Stats.MinMAN = 40: .Stats.MaxMAN = 80
            .Faccion.Status = e_Facciones.Ciudadano: .clase = e_Class.Mage
            .Char.charindex = i: .pos.Map = 1: .pos.x = 50: .pos.y = 50
            .CurrentInventorySlots = 3
        End With
        Set wrapper = New clsUserRefWrapper: Call wrapper.SetFromIndex(i)
        Set m_NameIndex(UCase$(UserList(i).Name)) = wrapper
    Next i
    LastUser = 7: rejectQuestion = False
    Set packets = New Collection: Set recipients = New Collection: capturing = True
End Sub
Private Sub Request(ByVal userIndex As Integer, ByVal id As Long, ByVal action As Byte, Optional ByVal name As String = "", Optional ByVal targetIndex As Integer = 0, Optional ByVal revision As Long = -1)
    Dim target As t_UserReference, leader As Integer, party As Long
    leader = PartyLeaderIndex(userIndex)
    If leader > 0 Then
        party = UserList(leader).Grupo.Id
        If revision = -1 Then revision = UserList(leader).Grupo.RosterRevision
    Else
        revision = 0
    End If
    If targetIndex > 0 Then Call SetUserRef(target, targetIndex)
    Call ExecutePartyManagement(userIndex, id, action, party, revision, target, name)
End Sub
Private Function ResultFor(ByVal userIndex As Integer, ByVal requestId As Long) As Long
    Dim i As Long, data() As Byte, reader As New Network.Reader
    ResultFor = -1
    For i = packets.Count To 1 Step -1
        If recipients(i) = userIndex Then
            data = packets(i): Call reader.SetData(data)
            If reader.ReadInt16() = ServerPacketID.eUpdateGroupInfo Then
                If reader.ReadInt32() = requestId Then
                    ResultFor = reader.ReadInt8(): Exit Function
                End If
            End If
        End If
    Next i
End Function
Private Sub Add(ByVal invited As Integer)
    Call Expect(InviteManagedParty(1, invited) = PartySuccess, "invitation issued")
    Call Expect(AcceptManagedParty(invited) = PartySuccess, "invitation accepted")
End Sub
Private Sub Dispatch(ByVal userIndex As Integer, ByVal packetId As Long, Optional ByVal response As Boolean = False)
    Dim writer As Network.Writer, reader As New Network.Reader, data() As Byte
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.Clear(): Call writer.WriteInt16(packetId)
    If packetId = ClientPacketID.ePartyManagement Then
        Call writer.WriteInt32(1): Call writer.WriteInt8(1)
        Call writer.WriteInt32(0): Call writer.WriteInt32(0)
        Call writer.WriteInt16(0): Call writer.WriteInt16(0): Call writer.WriteString8("")
    ElseIf packetId = ClientPacketID.eResponderPregunta Then
        Call writer.WriteBool(response)
    ElseIf packetId = ClientPacketID.eHecharDeGrupo Then
        Call writer.WriteInt8(1)
    End If
    Call writer.GetData(data): Call writer.Clear(): Call reader.SetData(data)
    If packetId = ClientPacketID.ePartyManagement Then
        Call Expect(UBound(data) = 19, "management request has full 20-byte layout; bytes=" & (UBound(data) + 1))
    End If
    Mapping(1).PacketCount = 0
    Call Protocol.HandleIncomingData(1, reader, userIndex)
End Sub
Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Fixture()
    Dim id As Long, revision As Long, before As Long, i As Integer, target As t_UserReference
    Dim data() As Byte, reader As New Network.Reader, name As String, alive As Boolean
    Select Case index
    Case 1
        Call Request(1, 1, 1)
        Call Expect(ResultFor(1, 1) = PartySuccess, "create confirmed")
        Call Expect(PartyLeaderIndex(1) = 1 And UserList(1).Grupo.CantidadMiembros = 1, "one-member authoritative roster")
        Call Expect(UserList(1).Grupo.Id > 0 And IsValidUserRef(UserList(1).Grupo.Miembros(1)), "valid identity and self reference")
        data = packets(packets.Count): Call reader.SetData(data)
        Call Expect(reader.ReadInt16() = ServerPacketID.eUpdateGroupInfo And reader.ReadInt32() = 1 And reader.ReadInt8() = PartySuccess, "packet result header")
        Call Expect(reader.ReadInt16() = 1 And reader.ReadInt16() = fixtureVersion And Not reader.ReadBool(), "local identity and safety")
        Call Expect(reader.ReadInt32() = UserList(1).Grupo.Id And reader.ReadInt32() = 1, "party identity/revision")
        Call Expect(reader.ReadInt16() = 1 And reader.ReadInt16() = fixtureVersion And reader.ReadInt8() = 1, "leader and count")
        Call Expect(reader.ReadInt16() = 1 And reader.ReadInt16() = fixtureVersion And reader.ReadString8() = "Party1", "member identity/name")
        Call Expect(reader.ReadInt16() = 1 And reader.ReadInt16() = 0, "character/head")
        Call Expect(reader.ReadInt32() = 100 And reader.ReadInt32() = 100 And reader.ReadInt32() = 40 And reader.ReadInt32() = 80 And reader.GetAvailable() = 0, "complete resource bounds")
    Case 2
        Call Request(1, 1, 1): id = UserList(1).Grupo.Id
        UserList(1).Grupo.XpRotation.PartyId = id
        Call Request(1, 1, 1): Call Request(1, 2, 1)
        Call Expect(UserList(1).Grupo.Id = id And UserList(1).Grupo.CantidadMiembros = 1 And UserList(1).Grupo.XpRotation.PartyId = id, "duplicate create cannot reset party or rotation")
        Call Expect(ResultFor(1, 2) = PartyAlreadyMember, "new duplicate create controlled error")
    Case 3
        Call Request(1, 1, 1): id = UserList(1).Grupo.Id
        Call Request(1, 2, 2, "Party2")
        Call Expect(ResultFor(1, 2) = PartySuccess And UserList(2).flags.pregunta = 1 And UserList(2).flags.RespondiendoPregunta, "named invitation uses server question")
        Call Expect(AcceptManagedParty(2) = PartySuccess, "acceptance succeeds")
        Call Expect(UserList(1).Grupo.CantidadMiembros = 2 And UserList(1).Grupo.Id = id, "acceptance preserves one-person identity")
        Call Expect(UserList(2).Grupo.ProposalPartyId = 0, "proposal cleaned")
    Case 4
        Call Request(1, 1, 1): Call Request(1, 2, 2, "Nobody")
        Call Expect(ResultFor(1, 2) = PartyUnavailable And UserList(1).Grupo.CantidadMiembros = 1, "unknown name")
    Case 5 To 9
        Call Request(1, 1, 1)
        Select Case index
        Case 5: UserList(2).flags.SeguroParty = True
        Case 6: UserList(2).pos.Map = 2
        Case 7: UserList(2).pos.x = 59
        Case 8: UserList(2).Stats.ELV = 50
        Case 9: UserList(2).flags.RespondiendoPregunta = True: UserList(2).flags.pregunta = 2
        End Select
        Call Request(1, 2, 2, "Party2")
        Select Case index
        Case 5: Call Expect(ResultFor(1, 2) = PartyBlocked, "safety blocks name invite")
        Case 6, 7: Call Expect(ResultFor(1, 2) = PartyUnavailable, "name invite respects tile targeting map/distance")
        Case 8: Call Expect(ResultFor(1, 2) = PartyLevel, "existing level restriction")
        Case 9: Call Expect(ResultFor(1, 2) = PartyQuestionPending And UserList(2).flags.pregunta = 2, "do not overwrite other question")
        End Select
    Case 10
        Call Request(1, 1, 1): Call Add(2)
        Call Request(2, 1, 2, "Party3"): Call Expect(ResultFor(2, 1) = PartyNotLeader, "nonleader cannot invite")
        Call Request(2, 2, 3, "", 1): Call Expect(ResultFor(2, 2) = PartyNotLeader, "nonleader cannot kick")
        Call Request(2, 3, 5): Call Expect(ResultFor(2, 3) = PartyNotLeader, "nonleader cannot disband")
        Call Expect(UserList(1).Grupo.CantidadMiembros = 2, "unauthorized actions have no mutation")
    Case 11
        Call Request(1, 1, 1): Call Add(2)
        Call Request(2, 1, 4)
        Call Expect(UserList(1).Grupo.CantidadMiembros = 1 And PartyLeaderIndex(1) = 1 And Not UserList(2).Grupo.EnGrupo, "last member leave retains leader")
        Call Expect(ResultFor(2, 1) = PartySuccess, "leave confirmation")
    Case 12
        Call Request(1, 1, 1): Call Add(2)
        Call Request(1, 2, 3, "", 2)
        Call Expect(UserList(1).Grupo.CantidadMiembros = 1 And PartyLeaderIndex(1) = 1 And Not UserList(2).Grupo.EnGrupo, "last kick retains leader")
    Case 13
        Call Request(1, 1, 1): Call Add(2): Call Add(3)
        revision = UserList(1).Grupo.RosterRevision
        Call RemoveManagedPartyMember(2)
        Call Request(1, 2, 3, "", 3, revision)
        Call Expect(ResultFor(1, 2) = PartyStale And UserList(3).Grupo.EnGrupo, "stale slot-shift kick rejected")
    Case 14
        Call Request(1, 1, 1): Call Add(2)
        Call SetUserRef(target, 2): Call IncreaseVersionId(2)
        Call SetUserRef(UserList(1).Grupo.Miembros(2), 2)
        Call ExecutePartyManagement(1, 2, 3, UserList(1).Grupo.Id, UserList(1).Grupo.RosterRevision, target, "")
        Call Expect(ResultFor(1, 2) = PartyStale And UserList(2).Grupo.EnGrupo, "kick cannot hit reused session")
    Case 15
        Call Request(1, 1, 1): Call Expect(InviteManagedParty(1, 2) = PartySuccess, "issued")
        Call DisbandManagedParty(1): Call GroupCreateSuccess(1)
        Call Expect(AcceptManagedParty(2) = PartyExpired And UserList(1).Grupo.CantidadMiembros = 1, "old proposal cannot join recreated party")
    Case 16
        Call Request(1, 1, 1): Call Expect(InviteManagedParty(1, 2) = PartySuccess, "issued")
        Call IncreaseVersionId(1)
        Call Expect(AcceptManagedParty(2) = PartyExpired, "inviter reconnect invalidates proposal")
    Case 17
        Call Request(1, 1, 1)
        For i = 2 To 5: Call Add(i): Next i
        Call Expect(InviteManagedParty(1, 6) = PartySuccess And InviteManagedParty(1, 7) = PartySuccess, "two proposals for final slot")
        Call Expect(AcceptManagedParty(6) = PartySuccess And AcceptManagedParty(7) = PartyFull, "acceptances serialize capacity")
        Call Expect(UserList(1).Grupo.CantidadMiembros = 6 And Not UserList(7).Grupo.EnGrupo, "never exceed six")
    Case 18
        Call Request(1, 1, 1): Call Expect(InviteManagedParty(1, 2) = PartySuccess, "issued")
        UserList(2).pos.x = 59
        Call Expect(AcceptManagedParty(2) = PartyUnavailable, "acceptance revalidates movement eligibility")
    Case 19
        Call Request(1, 1, 1): Call Add(2)
        UserList(1).ConnectionDetails.ConnIDValida = False
        Call FinalizarGrupo(1)
        Call Expect(Not UserList(1).Grupo.EnGrupo And Not UserList(2).Grupo.EnGrupo, "leader disconnect disbands even without connection")
    Case 20
        Call Request(1, 1, 1): Call Add(2): Call Request(1, 2, 5)
        Call Expect(ResultFor(1, 2) = PartySuccess And Not UserList(1).Grupo.EnGrupo And Not UserList(2).Grupo.EnGrupo, "authoritative disband")
    Case 21
        Call Request(1, 1, 1)
        MapInfo(2).OnlyGroups = True
        Call Expect(Not PartyMapAdmissionForTest(1, 2), "one-person party cannot enter group map")
        Call Add(2): Call Expect(PartyMapAdmissionForTest(1, 2), "two valid members retain admission")
        Call RemoveManagedPartyMember(2): Call Expect(Not PartyMapAdmissionForTest(1, 2), "last leave removes group admission")
    Case 22
        Call Request(1, 1, 1)
        Dim npc As t_Npc, n As Integer
        npc.Numero = 10: npc.pos = UserList(1).pos: npc.npcType = Comun
        npc.GiveEXP = 100: npc.flags.ExpCount = 100: npc.Stats.MaxHp = 100: npc.Stats.MinHp = 100
        n = GetNextAvailableNpc(): NpcList(n) = npc
        ExpLevelUp(10) = MAXEXP
        Call CalcularDarExp(1, n, 20)
        npc = NpcList(n): Call AwardNpcDeathExperience(1, npc)
        Call Expect(UserList(1).Stats.Exp = 100, "lone leader follows solo damage/death behavior")
    Case 23
        Call Request(1, 1, 1): rejectQuestion = True
        Call Request(1, 2, 2, "Party2")
        Call Expect(ResultFor(1, 2) = PartyUnavailable And Not UserList(2).flags.RespondiendoPregunta, "failed question enqueue cannot report sent")
    Case 24
        Call Request(1, 1, 1): Call Expect(InviteManagedParty(1, 2) = PartySuccess, "issued")
        Call ClearPartyProposal(2)
        Call Expect(AcceptManagedParty(2) = PartyExpired And UserList(1).Grupo.CantidadMiembros = 1, "declined/cancelled proposal not accepted")
    Case 25
        Call Request(1, 1, 1): Call Add(2)
        Call Request(1, 2, 3, "", 1)
        Call Expect(ResultFor(1, 2) = PartyStale And UserList(1).Grupo.CantidadMiembros = 2, "cannot kick self")
    Case 26
        Call Request(1, 1, 1): Call Add(2)
        revision = UserList(1).Grupo.RosterRevision
        Call Request(1, 2, 3, "", 2): Call Request(1, 2, 3, "", 2, revision)
        Call Expect(UserList(1).Grupo.CantidadMiembros = 1, "duplicate destructive request cannot apply twice")
    Case 27
        Call Request(1, 1, 1): Call Request(1, 2, 5)
        Call IncreaseVersionId(1)
        Call Request(1, 1, 1)
        Call Expect(PartyLeaderIndex(1) = 1 And ResultFor(1, 1) = PartySuccess, "new session starts fresh request correlation")
    Case 28
        Call Dispatch(1, ClientPacketID.ePartyManagement)
        Call Expect(PartyLeaderIndex(1) = 1 And ResultFor(1, 1) = PartySuccess, "maintained packet dispatcher/reader creates party; leader=" & PartyLeaderIndex(1) & "; result=" & ResultFor(1, 1))
    Case 29
        Call Request(1, 1, 1): Call Add(2)
        Call Dispatch(2, ClientPacketID.eAbandonarGrupo)
        Call Expect(UserList(2).Grupo.EnGrupo And ResultFor(2, 0) = PartyStale, "uncorrelated leave cannot affect new membership")
        Call Dispatch(1, ClientPacketID.eHecharDeGrupo)
        Call Expect(UserList(2).Grupo.EnGrupo And ResultFor(1, 0) = PartyStale, "positional kick cannot hit shifted member")
    Case 30, 31
        Call Request(1, 1, 1): id = UserList(1).Grupo.Id
        Call Expect(InviteManagedParty(1, 2) = PartySuccess, "question issued")
        Call Dispatch(2, ClientPacketID.eResponderPregunta, index = 30)
        If index = 30 Then
            Call Expect(UserList(1).Grupo.CantidadMiembros = 2 And UserList(1).Grupo.Id = id, "production question acceptance preserves identity")
        Else
            Call Expect(UserList(1).Grupo.CantidadMiembros = 1 And Not UserList(2).Grupo.EnGrupo, "production decline retains lone leader")
            Call Expect(ResultFor(1, 0) = PartyDeclined And Not UserList(2).flags.RespondiendoPregunta, "decline clears proposal and informs inviter")
        End If
    End Select
    capturing = False
    Call UnitTesting.Progress("Party management " & index & " passed")
    Call UnitTesting.RunTest("Party management " & index, True)
    Exit Sub
Failed:
    Dim failure As String: failure = Err.Description
    capturing = False
    Call UnitTesting.Progress("Party management " & index & " failed: " & failure)
    Call UnitTesting.RunTestError("Party management " & index, failure)
End Sub
Public Function TestSuite() As Boolean
    Dim i As Integer, before As Long, oldExp As Long
    before = UnitTesting.GetFailedTests(): oldExp = ExpLevelUp(10)
    For i = 1 To 31: Call RunCase(i): Next i
    ExpLevelUp(10) = oldExp
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
