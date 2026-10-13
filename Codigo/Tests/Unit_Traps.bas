Attribute VB_Name = "Unit_Traps"
Option Explicit
#If UNIT_TEST = 1 And DIRECT_PLAY = 0 Then
Private mUser As Integer, mNpc As Integer
Private mPackets As Collection
Private mCapture As Boolean
Private mSavedUser As t_User, mSavedNpc As t_Npc
Private mSavedTiles(11 To 14) As t_MapBlock
Private mSavedEffects(6 To 9) As t_EffectOverTime
Private mSavedObject As t_ObjData
Private mSavedConfig As ServerConfig
Private SoundCalls As Long, FxCalls As Long, CdClears As Long
Private Writer As Network.Writer

' Only the focused test-only startup path creates a synthetic world. The normal
' suite uses idle slots and restores every fixture it temporarily changes.
Public Sub InitializeTestWorld()
    Call InitializeCircularLogBuffer()
    ReDim UserList(1 To 2)
    ReDim MapInfo(1 To 1)
    ReDim MapData(1 To 100, 1 To 100, 1 To 1)
    ReDim EffectOverTime(1 To 9)
    ReDim ObjData(1 To 2)
    NumMaps = 1
    #If DIRECT_PLAY = 0 Then
        Call InitializeAuxiliaryBuffer()
    #End If
    Call InitializePools()
End Sub

Public Function CapturePacket() As Boolean
    If Not mCapture Then Exit Function
    Dim bytes() As Byte, value As Variant, packet As Integer
    Call Writer.GetData(bytes)
    Call Writer.Clear()
    packet = CInt(bytes(0)) + CInt(bytes(1)) * 256
    If packet = ServerPacketID.eUpdateTrap Then
        Call Expect(UBound(bytes) = 15, "fixed packet length")
        If bytes(2) = eTrapUpsert Then
            Call Expect(Not MapData(bytes(9), bytes(10), bytes(7)).Trap Is Nothing, "published trap exists")
            Call Expect(MapData(bytes(9), bytes(10), bytes(7)).Trap.RemainingMilliseconds = Read32(bytes, 12), "initialized before publication")
        End If
        value = bytes
        Call mPackets.Add(value)
    ElseIf packet = ServerPacketID.ePlayWave Then
        SoundCalls = SoundCalls + 1
    ElseIf packet = ServerPacketID.eCreateFX Then
        FxCalls = FxCalls + 1
    ElseIf packet = ServerPacketID.eSendSkillCdUpdate Then
        If Read32(bytes, 8) = 0 Then CdClears = CdClears + 1
    End If
    CapturePacket = True
End Function

Private Sub ResetFixture()
    Dim EmptyUser As t_User, EmptyNpc As t_Npc, EmptyTile As t_MapBlock
    Dim EmptyObject As t_ObjData, EmptyEffect As t_EffectOverTime, i As Integer
    Call ResetTrapTests()
    For i = 11 To 14
        If Not MapData(i, 10, 1).Trap Is Nothing Then Call MapData(i, 10, 1).Trap.Disable()
        MapData(i, 10, 1) = EmptyTile
    Next i
    UserList(mUser) = EmptyUser
    NpcList(mNpc) = EmptyNpc
    With UserList(mUser)
        .VersionId = 1
        .CurrentInventorySlots = 10
        .flags.UsingItemSlot = 1
        .flags.UsingItemObjIndex = 1
        .flags.Privilegios = e_PlayerType.User
        .invent.Object(1).ObjIndex = 1
        .invent.Object(1).amount = 20
        .pos.Map = 1: .pos.x = 10: .pos.y = 10
        .Stats.MinHp = 100: .Stats.MinSta = 100
    End With
    With NpcList(mNpc)
        .VersionId = 1
        .pos.Map = 1: .pos.x = 11: .pos.y = 10
        .Attackable = 1
        .flags.AttackableByEveryone = 1
        .Stats.MinHp = 1000: .Stats.MaxHp = 1000
    End With
    ObjData(1) = EmptyObject
    ObjData(1).OBJType = e_OBJType.otUsableOntarget
    ObjData(1).Subtipo = e_UssableOnTarget.eTrap
    ObjData(1).EfectoMagico = 6
    For i = 6 To 9
        EffectOverTime(i) = EmptyEffect
        EffectOverTime(i).Type = e_EffectOverTimeType.eTrap
        EffectOverTime(i).SubType = 1
        EffectOverTime(i).TickTime = 10000
        EffectOverTime(i).ClientEffectTypeId = 5
        EffectOverTime(i).OnHitFx = 8
        EffectOverTime(i).buffType = eInformativeBuff
    Next i
    EffectOverTime(7).SubType = 2: EffectOverTime(7).TickTime = 120000
    EffectOverTime(7).OnHitWav = 162: EffectOverTime(7).OnHitFx = 13
    EffectOverTime(7).TickPowerMin = -100: EffectOverTime(7).TickPowerMax = -100
    EffectOverTime(8).Type = eHealthModifier
    EffectOverTime(8).Ticks = 8: EffectOverTime(8).TickTime = 1500
    EffectOverTime(9).SubType = 3: EffectOverTime(9).TickTime = 120000
    EffectOverTime(9).ApplyEffectId = 8: EffectOverTime(9).OnHitFx = 0
    Set mPackets = New Collection
    SoundCalls = 0: FxCalls = 0: CdClears = 0
End Sub

Public Sub Expect(ByVal Condition As Boolean, ByVal Description As String)
    If Not Condition Then Err.Raise 5, "Trap test", Description
End Sub
Public Function Read32(ByVal bytes As Variant, ByVal Offset As Integer) As Long
    Read32 = CDbl(bytes(Offset)) + CDbl(bytes(Offset + 1)) * 256# + _
             CDbl(bytes(Offset + 2)) * 65536# + CDbl(bytes(Offset + 3)) * 16777216#
End Function
Private Function EventAt(ByVal Index As Integer) As Byte
    Dim bytes As Variant
    bytes = mPackets(Index)
    EventAt = bytes(2)
End Function
Private Sub ExpectPacket(ByVal Index As Integer, ByVal EventType As Byte, ByVal Id As Long, ByVal Lifetime As Long)
    Dim bytes As Variant
    bytes = mPackets(Index)
    Call Expect(bytes(0) = 192 And bytes(1) = 0, "packet 192")
    Call Expect(bytes(2) = EventType, "event ordering")
    Call Expect(Read32(bytes, 3) = Id, "stable unique id")
    Call Expect(Read32(bytes, 12) = Lifetime, "authoritative remaining time")
End Sub
Private Sub TestWire()
    Dim eventType As Integer, kind As Integer, i As Integer, bytes() As Byte, expected As Variant
    Call Expect(ServerPacketID.eUpdateTrap = 192, "ordinal in both build configurations")
    For eventType = 1 To 3
        For kind = 1 To 3
            Call PrepareTrapUpdate(eventType, &H1020304, &H1234, 99, 88, kind, &H5060708)
            Call Writer.GetData(bytes)
            Call Writer.Clear()
            expected = Array(192, 0, eventType, 4, 3, 2, 1, 52, 18, 99, 88, kind, 0, 0, 0, 0)
            If eventType = 1 Then
                expected(12) = 8: expected(13) = 7: expected(14) = 6: expected(15) = 5
            End If
            Call Expect(UBound(bytes) = 15, "exact 16 bytes, no owner fields")
            For i = 0 To 15
                Call Expect(bytes(i) = expected(i), "wire byte " & i)
            Next i
        Next kind
    Next eventType
    Call PrepareTrapUpdate(eTrapUpsert, 1, 1, 1, 1, 1, -1)
    Call Writer.GetData(bytes)
    Call Writer.Clear()
    Call Expect(Read32(bytes, 12) = 0, "negative lifetime clamp")
End Sub
Private Sub TestSnapshot()
    Dim trap As clsTrap, first As Variant, second As Variant
    Call Expect(CreateTrap(mUser, eUser, 1, 11, 10, 7), "create explosive")
    Set trap = MapData(11, 10, 1).Trap
    Call ExpectPacket(1, eTrapUpsert, 1, 120000)
    Call TickTrapTests(1250)
    Call trap.SendSnapshot(2)
    Call trap.SendSnapshot(2)
    Call ExpectPacket(2, eTrapUpsert, 1, 118750)
    Call ExpectPacket(3, eTrapUpsert, 1, 118750)
    first = mPackets(2): second = mPackets(3)
    Dim i As Integer
    For i = 0 To 15
        Call Expect(first(i) = second(i), "idempotent snapshot")
    Next i
    Call Expect(first(7) = 1 And first(9) = 11 And first(10) = 10 And first(11) = 2, "snapshot location and subtype")
    Call Expect(UserList(mUser).EffectOverTime.EffectCount = 1, "snapshot does not fabricate effects")
End Sub
Private Sub TestTrigger(ByVal EffectId As Integer)
    Dim trap As clsTrap
    Call Expect(CreateTrap(mUser, eUser, 1, 11, 10, EffectId), "place")
    Set trap = MapData(11, 10, 1).Trap
    Call ActivateTrap(mNpc, eNpc, 1, 11, 10)
    Call ActivateTrap(mNpc, eNpc, 1, 11, 10)
    Call Expect(Not trap.trigger(mNpc, eNpc), "duplicate direct trigger rejected")
    Call trap.SendSnapshot(2)
    Call Expect(mPackets.Count = 2, "trigger once, no stale upsert")
    Call ExpectPacket(2, eTrapTriggered, 1, 0)
    Call TickTrapTests(1)
    Call trap.IBaseEffectOverTime_OnRemove()
    Call Expect(mPackets.Count = 3, "remove once")
    Call ExpectPacket(3, eTrapRemoved, 1, 0)
    Call Expect(MapData(11, 10, 1).Trap Is Nothing, "tile released")
    Select Case EffectId
        Case 6
            Call Expect(NpcList(mNpc).flags.Paralizado = 1 And NpcList(mNpc).Contadores.Paralisis = 13, "real configured paralysis")
            Call Expect(SoundCalls = 1 And FxCalls = 2, "existing NPC and trap FX calls unchanged")
        Case 7
            Call Expect(NpcList(mNpc).Stats.MinHp = 900 And SoundCalls = 2 And FxCalls = 1, "exactly one real explosion")
        Case 9
            Call Expect(NpcList(mNpc).EffectOverTime.EffectCount = 1 And SoundCalls = 1 And FxCalls = 0, "exactly one real poison effect")
            Call Expect(NpcList(mNpc).EffectOverTime.EffectList(0).EotId = 8, "configured poison EOT")
    End Select
End Sub
Private Sub TestRejections()
    Call Expect(CreateTrap(mUser, eUser, 1, 11, 10, 6), "place")
    Dim trap As clsTrap
    Set trap = MapData(11, 10, 1).Trap
    Call Expect(Not trap.trigger(2, eUser), "PvE only")
    Call Expect(Not trap.trigger(0, eNpc), "invalid NPC")
    NpcList(mNpc).Attackable = 0
    Call ActivateTrap(mNpc, eNpc, 1, 11, 10)
    NpcList(mNpc).Attackable = 1
    NpcList(mNpc).flags.AfectaParalisis = 1
    Call Expect(Not trap.trigger(mNpc, eNpc), "immune NPC")
    Call Expect(mPackets.Count = 1 And SoundCalls = 0 And NpcList(mNpc).flags.Paralizado = 0, "no fake activation")
    Call Expect(trap.IsActive, "rejections do not consume trap")
End Sub
Private Sub TestExpiry()
    Call Expect(CreateTrap(mUser, eUser, 1, 11, 10, 6), "place")
    Call TickTrapTests(9999)
    Call Expect(MapData(11, 10, 1).Trap.RemainingMilliseconds = 1, "last live millisecond")
    Call TickTrapTests(1)
    Call TickTrapTests(1)
    Call Expect(mPackets.Count = 2, "expire once")
    Call ExpectPacket(2, eTrapRemoved, 1, 0)
    Call Expect(UserList(mUser).EffectOverTime.EffectCount = 0 And MapData(11, 10, 1).Trap Is Nothing, "EOT cleanup")
End Sub
Private Sub TestInvalidOwner()
    Call Expect(CreateTrap(mUser, eUser, 1, 11, 10, 6), "place")
    UserList(mUser).VersionId = 2
    Call MapData(11, 10, 1).Trap.SendSnapshot(2)
    Call ActivateTrap(mNpc, eNpc, 1, 11, 10)
    Call TickTrapTests(1)
    Call Expect(mPackets.Count = 2 And SoundCalls = 0, "invalid owner removed without activation")
    Call ExpectPacket(2, eTrapRemoved, 1, 0)
    Call Expect(CdClears = 0, "do not send old cooldown to recycled user slot")
End Sub
Private Sub TestDeath()
    Call Expect(CreateTrap(mUser, eUser, 1, 11, 10, 6), "place")
    UserList(mUser).flags.Muerto = 1
    Call ClearEffectList(UserList(mUser).EffectOverTime, e_EffectType.eAny, True)
    Call TickTrapTests(1)
    Call Expect(mPackets.Count = 2 And CdClears = 1, "death removes once")
    Call ExpectPacket(2, eTrapRemoved, 1, 0)
End Sub
Private Sub TestFourth()
    Dim i As Integer
    ' Equal-age traps must select the first active trap.
    For i = 11 To 14
        Call PlaceTrap(mUser, i, 10)
    Next i
    Call Expect(mPackets.Count = 5, "three upserts then removed/upsert")
    Call ExpectPacket(4, eTrapRemoved, 1, 0)
    Call ExpectPacket(5, eTrapUpsert, 4, 10000)
    Call Expect(MapData(11, 10, 1).Trap Is Nothing, "oldest released immediately")
    Call Expect(UserList(mUser).invent.Object(1).amount = 16 And UserList(mUser).CdTimes(0) <> 0, "real consumption/cooldown updated")
    Call TickTrapTests(1)
    Call Expect(UserList(mUser).EffectOverTime.EffectCount = 3 And mPackets.Count = 5, "no duplicate late removal")

    Call ResetFixture()
    Call PlaceTrap(mUser, 11, 10)
    Call MapData(11, 10, 1).Trap.Disable()
    ' An inactive entry at index zero must not win oldest-trap selection.
    For i = 11 To 14
        Call PlaceTrap(mUser, i, 10)
    Next i
    Call ExpectPacket(6, eTrapRemoved, 2, 0)
    Call ExpectPacket(7, eTrapUpsert, 5, 10000)
    Call TickTrapTests(1)
    Call Expect(UserList(mUser).EffectOverTime.EffectCount = 3 And mPackets.Count = 7, "inactive entries excluded from cap")

    Call ResetFixture()
    Call PlaceTrap(mUser, 11, 10)
    Call MapData(11, 10, 1).Trap.Disable()
    Call PlaceTrap(mUser, 11, 10)
    Call TickTrapTests(1)
    Call Expect(Not MapData(11, 10, 1).Trap Is Nothing, "late cleanup preserves replacement on same tile")
    Call Expect(MapData(11, 10, 1).Trap.IBaseEffectOverTime_UniqueId = 2, "replacement identity preserved")
End Sub
Private Sub TestInventory()
    Dim i As Integer
    For i = 1 To 8
        Call ResetFixture()
        Select Case i
            Case 1: UserList(mUser).flags.UsingItemSlot = -1
            Case 2: UserList(mUser).flags.UsingItemSlot = 11
            Case 3: UserList(mUser).invent.Object(1).ObjIndex = 0
            Case 4: UserList(mUser).invent.Object(1).ObjIndex = 30000
            Case 5: UserList(mUser).invent.Object(1).amount = 0
            Case 6: UserList(mUser).invent.Object(1).ObjIndex = 2
            Case 7: ObjData(1).OBJType = e_OBJType.otWeapon
            Case 8: UserList(mUser).flags.UsingItemObjIndex = 0
        End Select
        Dim amount As Integer
        amount = UserList(mUser).invent.Object(1).amount
        Call UserTargetableItem(mUser, 11, 10)
        Call Expect(mPackets.Count = 0 And UserList(mUser).invent.Object(1).amount = amount, "invalid inventory rejected")
        Call Expect(UserList(mUser).flags.UsingItemSlot = 0 And UserList(mUser).flags.UsingItemObjIndex = 0, "stale request cleared")
    Next i
    Call ResetFixture()
    Call UserTargetableItem(mUser, 11, 10)
    Call UserTargetableItem(mUser, 12, 10)
    Call Expect(mPackets.Count = 1 And UserList(mUser).invent.Object(1).amount = 19, "one-shot item target submission")
End Sub
Private Sub TestPlacementGuards()
    Dim i As Integer
    For i = 1 To 7
        Call ResetFixture()
        Select Case i
            Case 1: ObjData(1).Subtipo = e_UssableOnTarget.eArpon
            Case 2: ObjData(1).EfectoMagico = 8
            Case 3: EffectOverTime(6).SubType = 4
            Case 4
                ObjData(1).Cooldown = 30000
                UserList(mUser).CdTimes(0) = GetTickCountRaw()
            Case 5: UserList(mUser).flags.Muerto = 1
            Case 6: MapData(11, 10, 1).Blocked = 1
            Case 7: ObjData(1).EfectoMagico = 0
        End Select
        Dim cooldownBefore As Long
        cooldownBefore = UserList(mUser).CdTimes(0)
        Call PlaceTrap(mUser, 11, 10)
        Call Expect(mPackets.Count = 0 And UserList(mUser).invent.Object(1).amount = 20 And UserList(mUser).CdTimes(0) = cooldownBefore, "invalid placement no mutation")
    Next i
    Call ResetFixture()
    Call PlaceTrap(mUser, 15, 10)
    Call PlaceTrap(mUser, 0, 10)
    Call Expect(mPackets.Count = 0 And UserList(mUser).invent.Object(1).amount = 20, "range/bounds unchanged")
End Sub
Private Sub TestReuseAndClamp()
    Dim trap As New clsTrap
    Call trap.Setup(mUser, eUser, 7, 0, 1, 11, 10)
    Call trap.Setup(mUser, eUser, 8, 43, 1, 11, 10)
    Call trap.Setup(mUser, e_ReferenceType.eNone, 7, 43, 1, 11, 10)
    Call Expect(mPackets.Count = 0 And Not trap.IsActive, "invalid setup never published")
    Call trap.Setup(mUser, eUser, 7, 44, 1, 11, 10)
    Call trap.IBaseEffectOverTime_Update(-100)
    Call Expect(trap.RemainingMilliseconds = 120000, "ignore negative delta")
    Call trap.Disable()
    Call trap.Setup(mUser, eUser, 6, 45, 1, 12, 10)
    Call ExpectPacket(3, eTrapUpsert, 45, 10000)
    Call trap.IBaseEffectOverTime_Update(&H7FFFFFFF)
    Call Expect(trap.RemainingMilliseconds = 0, "long delta clamps without overflow")
    Call trap.IBaseEffectOverTime_OnRemove()
    Call ExpectPacket(4, eTrapRemoved, 45, 0)
End Sub
Private Sub TestIdWrap()
    Call Expect(TrapTestNextId(0) = 1, "positive initial id")
    Call Expect(TrapTestNextId(&H7FFFFFFE) = &H7FFFFFFF, "maximum signed id")
    Call Expect(TrapTestNextId(&H7FFFFFFF) = 1, "overflow skips zero")
End Sub
Private Sub RunCase(ByVal Number As Integer, ByVal Name As String)
    On Error GoTo FailedCase
    Call ResetFixture()
    Select Case Number
        Case 1: Call TestWire()
        Case 2: Call TestSnapshot()
        Case 3: Call TestTrigger(6)
        Case 4: Call TestTrigger(7)
        Case 5: Call TestTrigger(9)
        Case 6: Call TestRejections()
        Case 7: Call TestExpiry()
        Case 8: Call TestInvalidOwner()
        Case 9: Call TestDeath()
        Case 10: Call TestFourth()
        Case 11: Call TestInventory()
        Case 12: Call TestPlacementGuards()
        Case 13: Call TestReuseAndClamp()
        Case 14: Call TestIdWrap()
    End Select
    Call UnitTesting.RunTest("traps " & Name, True)
    Exit Sub
FailedCase:
    Call UnitTesting.RunTestError("traps " & Name, Err.Source & " " & Err.Description)
End Sub
Public Function test_suite_traps() As Boolean
    On Error GoTo FailedSuite
    Dim names As Variant, i As Integer, f As Integer
    Dim saved As Boolean, effectsSaved As Boolean, configPath As String
    Dim fs As New FileSystemObject
    For i = UBound(UserList) To LBound(UserList) Step -1
        If Not UserList(i).flags.UserLogged And Not UserList(i).ConnectionDetails.ConnIDValida Then
            mUser = i
            Exit For
        End If
    Next i
    For i = UBound(NpcList) To LBound(NpcList) Step -1
        If Not NpcList(i).flags.NPCActive Then
            mNpc = i
            Exit For
        End If
    Next i
    Call Expect(mUser > 0 And mNpc > 0, "idle fixture slots available")
    mSavedUser = UserList(mUser)
    mSavedNpc = NpcList(mNpc)
    mSavedObject = ObjData(1)
    Set mSavedConfig = SvrConfig
    For i = 11 To 14
        mSavedTiles(i) = MapData(i, 10, 1)
    Next i
    For i = 6 To 9
        mSavedEffects(i) = EffectOverTime(i)
    Next i
    saved = True
    ' Detach original map references before fixture cleanup touches these tiles.
    For i = 11 To 14
        Set MapData(i, 10, 1).Trap = Nothing
    Next i
    configPath = fs.BuildPath(fs.GetSpecialFolder(2), fs.GetTempName())
    f = FreeFile()
    Open configPath For Output As #f
    Print #f, "[CONFIGURACIONES]"
    Print #f, "TrapPlacementRange=4"
    Print #f, "TrapParaliceDuration=2"
    Close #f
    f = 0
    Set SvrConfig = New ServerConfig
    Call Expect(SvrConfig.LoadSettings(configPath) = 0, "load fixture configuration")
    Set Writer = Protocol_Writes.GetWriterBuffer()
    Call Writer.Clear()
    Set mPackets = New Collection
    mCapture = True
    Call BeginTrapTests()
    effectsSaved = True
    names = Array("Wire", "Snapshots", "Paralysis", "Explosion", "Poison", "Rejection/immunity", _
                  "Expiry", "Invalid owner", "Death", "Fourth trap", "Inventory", "Placement guards", "Pool reuse/time clamp", "Id wrap")
    For i = 0 To UBound(names)
        Call RunCase(i + 1, names(i))
    Next i
    test_suite_traps = True
Cleanup:
    On Error Resume Next
    If f <> 0 Then Close #f
    If effectsSaved Then
        Call ResetFixture()
        Call EndTrapTests()
    End If
    mCapture = False
    If saved Then
        UserList(mUser) = mSavedUser
        NpcList(mNpc) = mSavedNpc
        ObjData(1) = mSavedObject
        Set SvrConfig = mSavedConfig
        For i = 11 To 14
            MapData(i, 10, 1) = mSavedTiles(i)
            Set mSavedTiles(i).Trap = Nothing
        Next i
        For i = 6 To 9
            EffectOverTime(i) = mSavedEffects(i)
        Next i
    End If
    If Len(configPath) > 0 Then Call fs.DeleteFile(configPath)
    Set mPackets = Nothing
    Set Writer = Nothing
    Set mSavedConfig = Nothing
    mUser = 0: mNpc = 0
    Exit Function
FailedSuite:
    Call UnitTesting.RunTestError("traps fixture setup", Err.Description)
    Resume Cleanup
End Function
#End If
