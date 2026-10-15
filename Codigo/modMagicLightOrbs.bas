Attribute VB_Name = "modMagicLightOrbs"
' Argentum 20 Game Server
' Copyright (C) 2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.
' See licence.txt for the full license.
Option Explicit

Private Const ORB_UPDATE_INTERVAL_MS As Long = 100
Private Const MAX_ORB_DURATION_SECONDS As Long = 3600
Public Const MAGIC_LIGHT_ORB_PARTICLE As Integer = 354

Private Type t_MagicLightOrb
    Active As Boolean
    Owner As t_UserReference
    CreatedAt As Long
    DurationMs As Long
End Type

Private UserOrbs() As t_MagicLightOrb
Private OrbsInitialized As Boolean
Private LastUpdate As Long
#If UNIT_TEST = 1 And DIRECT_PLAY = 0 Then
Private CaptureOrbPackets As Boolean
Private OrbUpserts As Long, OrbRemovals As Long, OrbTargetRequests As Long
Private OrbWireValid As Boolean
Private OrbPacketOrder As String
#End If

Public Function CanClassCastLightOrb(ByVal ClassId As e_Class) As Boolean
    CanClassCastLightOrb = (ClassId = Mage Or ClassId = Druid Or ClassId = Bard)
End Function

Private Function OrbHasTime(ByRef Orb As t_MagicLightOrb, ByVal NowRaw As Long) As Boolean
    OrbHasTime = Orb.Active And TicksElapsed(Orb.CreatedAt, NowRaw) < Orb.DurationMs
End Function

Public Function CastMagicLightOrb(ByVal UserIndex As Integer, ByVal SpellId As Integer) As Boolean
    On Error GoTo CastMagicLightOrb_Err
    If Not CanClassCastLightOrb(UserList(UserIndex).clase) Then Exit Function
    If Not IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then Exit Function
    If UserList(UserIndex).flags.TargetUser.ArrayIndex <> UserIndex Then Exit Function
    With Hechizos(SpellId)
        If .Duration <= 0 Or .Duration > MAX_ORB_DURATION_SECONDS Then Exit Function
        If .Particle <> MAGIC_LIGHT_ORB_PARTICLE Then Exit Function
    End With
    If Not OrbsInitialized Then
        ReDim UserOrbs(1 To UBound(UserList))
        OrbsInitialized = True
    End If
    ' Refresh the deadline without a remove/create flicker on clients.
    With UserOrbs(UserIndex)
        Call SetUserRef(.Owner, UserIndex)
        .CreatedAt = GetTickCountRaw()
        .DurationMs = CLng(Hechizos(SpellId).Duration) * 1000
        .Active = True
    End With
    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, MAGIC_LIGHT_ORB_PARTICLE, -1, False))
    CastMagicLightOrb = True
    Exit Function
CastMagicLightOrb_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.CastMagicLightOrb", Erl)
End Function

Public Sub RemoveMagicLightOrb(ByVal UserIndex As Integer)
    On Error GoTo RemoveMagicLightOrb_Err
    If Not OrbsInitialized Then Exit Sub
    If UserIndex < 1 Or UserIndex > UBound(UserOrbs) Then Exit Sub
    With UserOrbs(UserIndex)
        If Not .Active Then Exit Sub
        .Active = False
        ' Never remove the effects of a new occupant of a recycled user slot.
        If Not IsValidUserRef(.Owner) Then Exit Sub
    End With
    If UserList(UserIndex).Char.charindex > 0 Then
        ' Removal also reaches dead observers who saw the effect before dying.
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, MAGIC_LIGHT_ORB_PARTICLE, 0, True))
    End If
    Exit Sub
RemoveMagicLightOrb_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.RemoveMagicLightOrb", Erl)
End Sub

Public Sub UpdateMagicLightOrbs()
    On Error GoTo UpdateMagicLightOrbs_Err
    If Not OrbsInitialized Then Exit Sub
    Dim NowRaw As Long
    NowRaw = GetTickCountRaw()
    If TicksElapsed(LastUpdate, NowRaw) < ORB_UPDATE_INTERVAL_MS Then Exit Sub
    LastUpdate = NowRaw
    Dim owner As Integer
    For owner = 1 To UBound(UserOrbs)
        With UserOrbs(owner)
            If .Active Then
                If Not OrbHasTime(UserOrbs(owner), NowRaw) Or Not IsValidUserRef(.Owner) Or Not UserList(owner).ConnectionDetails.ConnIDValida Then
                    Call RemoveMagicLightOrb(owner)
                End If
            End If
        End With
    Next owner
    Exit Sub
UpdateMagicLightOrbs_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.UpdateMagicLightOrbs", Erl)
End Sub

' Only called after the visibility-filtered CharacterCreate for this owner.
' Normal map changes keep the original deadline; no floor state is retained.
Public Sub SendMagicLightOrbSnapshot(ByVal Recipient As Integer, ByVal OwnerIndex As Integer)
    On Error GoTo SendMagicLightOrbSnapshot_Err
    If Not OrbsInitialized Then Exit Sub
    If OwnerIndex < 1 Or OwnerIndex > UBound(UserOrbs) Then Exit Sub
    With UserOrbs(OwnerIndex)
        If Not OrbHasTime(UserOrbs(OwnerIndex), GetTickCountRaw()) Then Exit Sub
        If Not IsValidUserRef(.Owner) Then Exit Sub
    End With
    If UserList(Recipient).pos.Map <> UserList(OwnerIndex).pos.Map Then Exit Sub
    If UserList(OwnerIndex).flags.AdminInvisible = 1 Then
        If Not CanUserSeeAdminInvisible(Recipient, OwnerIndex) Then Exit Sub
    End If
    ' NotifyUser normally sends this after MakeUserChar. The orb needs it first,
    ' even when TCP splits the snapshot across client frames.
    If UserList(OwnerIndex).flags.invisible Or UserList(OwnerIndex).flags.Oculto Then
        Call WriteSetInvisible(Recipient, OwnerIndex, True)
    End If
    Call SendData(SendTarget.ToIndex, Recipient, PrepareMessageParticleFX(UserList(OwnerIndex).Char.charindex, MAGIC_LIGHT_ORB_PARTICLE, -1, False))
    Exit Sub
SendMagicLightOrbSnapshot_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.SendMagicLightOrbSnapshot", Erl)
End Sub

#If UNIT_TEST = 1 Then
#If DIRECT_PLAY = 0 Then
Public Function CaptureMagicLightOrbPacket() As Boolean
    If Not CaptureOrbPackets Then Exit Function
    Dim bytes() As Byte, writer As Network.Writer, packet As Integer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    packet = CInt(bytes(0)) + CInt(bytes(1)) * 256
    OrbPacketOrder = OrbPacketOrder & CStr(packet) & ","
    If packet = ServerPacketID.eParticleFX Then
        OrbWireValid = OrbWireValid And UBound(bytes) = 16 And bytes(4) = 98 And bytes(5) = 1
        If bytes(10) = 0 Then
            OrbUpserts = OrbUpserts + 1
            OrbWireValid = OrbWireValid And bytes(6) = 255 And bytes(7) = 255 And bytes(8) = 255 And bytes(9) = 255
        Else
            OrbRemovals = OrbRemovals + 1
        End If
    ElseIf packet = ServerPacketID.eWorkRequestTarget Then
        OrbTargetRequests = OrbTargetRequests + 1
    End If
    CaptureMagicLightOrbPacket = True
End Function

Private Function TestOrbLifecycle(Optional ByVal failSetup As Boolean = False) As String
    On Error GoTo TestOrbLifecycle_Err
    Dim owner As Integer, other As Integer, savedOwner As t_User, savedOther As t_User
    Dim savedSpell As t_Hechizo, emptyUser As t_User, emptySpell As t_Hechizo
    Dim originalDeadline As Long, sent As Long, savedSinMagia As Boolean
    Dim spellUpper As Long, spellLower As Long, extended As Boolean
    Dim usersSaved As Boolean, spellSaved As Boolean, mapSaved As Boolean, ownsOrbs As Boolean
    Dim failure As String
    spellLower = LBound(Hechizos): spellUpper = UBound(Hechizos)
    Call UnitTesting.Progress("orb fixture: loaded spell upper bound=" & CStr(spellUpper))
    owner = UBound(UserList): other = owner - 1
    savedOwner = UserList(owner): savedOther = UserList(other)
    usersSaved = True
    savedSinMagia = MapInfo(1).SinMagia: mapSaved = True
    ' Cross-repository deployment validation is explicit; ordinary unit runs use
    ' the fixture below, independently of which assets branch CI has installed.
    If LCase$(Trim$(Command$)) = "--validate-light-orb-resources" Then
        If Not failSetup Then
            If spellUpper < 352 Then
                Call UnitTesting.RunTestError("orb loaded resource self-cast contract", "Loaded spell 352 is missing")
            Else
                savedSpell = Hechizos(352)
                Call UnitTesting.RunTest("orb loaded resource self-cast contract", savedSpell.AutoLanzar = 1 And savedSpell.Target = uUsuarios And savedSpell.Tipo = uLightOrb And savedSpell.Duration = 60 And savedSpell.Cooldown = 2 And savedSpell.ManaRequerido = 20 And savedSpell.StaRequerido = 5 And savedSpell.Particle = MAGIC_LIGHT_ORB_PARTICLE And IsSet(savedSpell.SpellRequirementMask, e_SpellRequirementMask.eIsBindable))
            End If
        End If
    End If
    If spellUpper < 352 Then
        ReDim Preserve Hechizos(spellLower To 352)
        extended = True
    End If
    savedSpell = Hechizos(352): spellSaved = True
    If failSetup Then Err.Raise 5, , "injected orb fixture setup failure"
    If OrbsInitialized Then Err.Raise 5, , "Expected empty startup orb state"
    ownsOrbs = True
    CaptureOrbPackets = True: OrbWireValid = True
    OrbUpserts = 0: OrbRemovals = 0: OrbTargetRequests = 0
    UserList(owner) = emptyUser
    With UserList(owner)
        .VersionId = 12: .clase = Mage: .pos.Map = 1
        .pos.x = 50: .pos.y = 50: .Char.charindex = 10
        .ConnectionDetails.ConnIDValida = True
        .flags.Privilegios = e_PlayerType.User: .flags.NoPalabrasMagicas = 1
        .Stats.MinMAN = 100: .Stats.MinSta = 100: .Stats.MinHp = 100
        .Stats.UserSkills(Magia) = 100: .Stats.UserHechizos(1) = 352
        .Counters.StunEndTime = GetTickCountRaw()
    End With
    Hechizos(352) = emptySpell
    With Hechizos(352)
        .AutoLanzar = 1: .Target = uUsuarios: .Tipo = uLightOrb
        .Duration = 60: .Cooldown = 2: .Particle = MAGIC_LIGHT_ORB_PARTICLE
        .ManaRequerido = 20: .StaRequerido = 5
    End With
    MapInfo(1).SinMagia = False
    UserList(owner).clase = Warrior
    Call UseSpellSlot(owner, 1)
    Call UnitTesting.RunTest("orb invalid class rejected without cost", OrbUpserts = 0 And UserList(owner).Stats.MinMAN = 100)
    UserList(owner).clase = Mage: UserList(owner).Stats.MinMAN = 19
    Call UseSpellSlot(owner, 1)
    Call UnitTesting.RunTest("orb insufficient mana rejected", OrbUpserts = 0 And UserList(owner).Stats.MinSta = 100)
    UserList(owner).Stats.MinMAN = 100: UserList(owner).Stats.MinSta = 4
    Call UseSpellSlot(owner, 1)
    Call UnitTesting.RunTest("orb insufficient stamina rejected", OrbUpserts = 0 And UserList(owner).Stats.MinMAN = 100)
    UserList(owner).Stats.MinSta = 100
    Call UseSpellSlot(owner, 1)
    Call UnitTesting.RunTest("orb self cast needs no target request", OrbUpserts = 1 And OrbTargetRequests = 0)
    Call UnitTesting.RunTest("orb original resource cost", UserList(owner).Stats.MinMAN = 80 And UserList(owner).Stats.MinSta = 95)
    Call UseSpellSlot(owner, 1)
    Call UnitTesting.RunTest("orb cooldown rejects recast without cost", OrbUpserts = 1 And UserList(owner).Stats.MinMAN = 80)
    UserList(owner).Counters.UserHechizosInterval(1) = 0
    Call UseSpellSlot(owner, 1)
    Call UnitTesting.RunTest("orb refresh does not remove old presentation", OrbUpserts = 2 And OrbRemovals = 0)
    originalDeadline = UserOrbs(owner).CreatedAt
    UserList(owner).pos.Map = 2
    LastUpdate = 0
    Call UpdateMagicLightOrbs
    Call UnitTesting.RunTest("orb survives map transition with original deadline", UserOrbs(owner).Active And UserOrbs(owner).CreatedAt = originalDeadline)
    UserList(owner).pos.Map = 1: UserList(owner).flags.Muerto = 1
    LastUpdate = 0
    Call UpdateMagicLightOrbs
    Call UnitTesting.RunTest("orb preserves death lifetime", UserOrbs(owner).Active)
    UserList(other) = UserList(owner): UserList(other).flags.Muerto = 0: UserList(other).clase = Bard
    UserList(other).Char.charindex = 11: UserList(other).Counters.UserHechizosInterval(1) = 0
    Call UseSpellSlot(other, 1)
    Call UnitTesting.RunTest("orb independent concurrent casters", UserOrbs(owner).Active And UserOrbs(other).Active)
    sent = OrbUpserts
    Call SendMagicLightOrbSnapshot(other, owner)
    Call SendMagicLightOrbSnapshot(other, owner)
    Call UnitTesting.RunTest("orb repeat snapshots preserve deadline", OrbUpserts = sent + 2 And UserOrbs(owner).CreatedAt = originalDeadline)
    OrbPacketOrder = vbNullString
    Call MakeUserChar(False, other, owner, 1, 50, 50)
    Call UnitTesting.RunTest("orb snapshot follows character creation", OrbPacketOrder = CStr(ServerPacketID.eCharacterCreate) & "," & CStr(ServerPacketID.eParticleFX) & ",")
    UserList(owner).flags.invisible = 1
    OrbPacketOrder = vbNullString
    Call MakeUserChar(False, other, owner, 1, 50, 50)
    Call UnitTesting.RunTest("orb invisibility precedes light snapshot", OrbPacketOrder = CStr(ServerPacketID.eCharacterCreate) & "," & CStr(ServerPacketID.eSetInvisible) & "," & CStr(ServerPacketID.eParticleFX) & ",")
    UserList(owner).flags.invisible = 0
    UserList(owner).flags.AdminInvisible = 1
    sent = OrbUpserts
    Call SendMagicLightOrbSnapshot(other, owner)
    Call UnitTesting.RunTest("orb hidden admin not revealed", OrbUpserts = sent)
    UserList(owner).flags.AdminInvisible = 0
    UserOrbs(owner).DurationMs = 0: LastUpdate = 0
    Call UpdateMagicLightOrbs
    Call RemoveMagicLightOrb(owner)
    Call UnitTesting.RunTest("orb expiration removes exactly once", Not UserOrbs(owner).Active And OrbRemovals = 1 And UserOrbs(other).Active)
    UserList(other).ConnectionDetails.ConnIDValida = False: LastUpdate = 0
    Call UpdateMagicLightOrbs
    Call UnitTesting.RunTest("orb disconnect cleanup", Not UserOrbs(other).Active And OrbRemovals = 2)
    UserList(owner).flags.Muerto = 0: UserList(owner).Counters.UserHechizosInterval(1) = 0
    Call UseSpellSlot(owner, 1)
    UserList(owner).VersionId = 13: LastUpdate = 0
    Call UpdateMagicLightOrbs
    Call UnitTesting.RunTest("orb recycled slot never removes new owner's effect", Not UserOrbs(owner).Active And OrbRemovals = 2)
    Call UnitTesting.RunTest("orb character packet wire identity", OrbWireValid)
TestOrbLifecycle_Cleanup:
    ' Cleanup runs once, after unwinding the original handler. A failed restore
    ' must never Resume back into this block and endlessly record the same error.
    On Error Resume Next
    CaptureOrbPackets = False
    Err.Clear
    If usersSaved Then
        UserList(owner) = savedOwner: UserList(other) = savedOther
    End If
    If Err.Number <> 0 Then failure = failure & " | users cleanup: " & CStr(Err.Number) & " " & Err.Description
    Err.Clear
    If spellSaved Then Hechizos(352) = savedSpell
    If Err.Number <> 0 Then failure = failure & " | spell cleanup: " & CStr(Err.Number) & " " & Err.Description
    Err.Clear
    If extended Then ReDim Preserve Hechizos(spellLower To spellUpper)
    If Err.Number <> 0 Then failure = failure & " | spell bounds cleanup: " & CStr(Err.Number) & " " & Err.Description
    Err.Clear
    If mapSaved Then MapInfo(1).SinMagia = savedSinMagia
    If Err.Number <> 0 Then failure = failure & " | map cleanup: " & CStr(Err.Number) & " " & Err.Description
    Err.Clear
    If ownsOrbs Then
        Erase UserOrbs: OrbsInitialized = False: LastUpdate = 0
    End If
    If Err.Number <> 0 Then failure = failure & " | orb cleanup: " & CStr(Err.Number) & " " & Err.Description
    On Error GoTo 0
    TestOrbLifecycle = failure
    Exit Function
TestOrbLifecycle_Err:
    failure = "Err.Number=" & CStr(Err.Number) & "; " & Err.Description
    Resume TestOrbLifecycle_Cleanup
End Function
#End If
Public Function test_suite_light_orbs() As Boolean
    On Error GoTo test_suite_light_orbs_Err
    Dim orb As t_MagicLightOrb
    orb.Active = True
    orb.CreatedAt = 100
    orb.DurationMs = 60000
    Call UnitTesting.RunTest("orb eligible classes", CanClassCastLightOrb(Mage) And CanClassCastLightOrb(Druid) And CanClassCastLightOrb(Bard) And Not CanClassCastLightOrb(Warrior))
    Call UnitTesting.RunTest("orb late observer preserves lifetime", OrbHasTime(orb, 60099))
    Call UnitTesting.RunTest("orb expires at deadline", Not OrbHasTime(orb, 60100))
    orb.Active = False
    Call UnitTesting.RunTest("orb removed excluded from snapshot", Not OrbHasTime(orb, 200))
    orb.Active = True
    orb.CreatedAt = &HFFFFFFF0
    orb.DurationMs = 100
    Call UnitTesting.RunTest("orb clock wrap active", OrbHasTime(orb, &H10))
    Call UnitTesting.RunTest("orb clock wrap expiry", Not OrbHasTime(orb, &H54))
    Call UnitTesting.RunTest("orb dedicated character particle", MAGIC_LIGHT_ORB_PARTICLE = 354)
    #If DIRECT_PLAY = 0 Then
        Dim failure As String, originalUpper As Long
        originalUpper = UBound(Hechizos)
        failure = TestOrbLifecycle(True)
        Call UnitTesting.RunTest("orb setup failure restores fixture once", failure = "Err.Number=5; injected orb fixture setup failure" And UBound(Hechizos) = originalUpper And Not OrbsInitialized And Not CaptureOrbPackets)
        failure = TestOrbLifecycle()
        If LenB(failure) <> 0 Then Call UnitTesting.RunTestError("orb lifecycle", failure)
        Call UnitTesting.RunTest("orb lifecycle restores loaded spell bounds", UBound(Hechizos) = originalUpper)
    #End If
    test_suite_light_orbs = True
    Exit Function
test_suite_light_orbs_Err:
    Call UnitTesting.RunTestError("character light orb suite", Err.Description)
End Function
#End If
