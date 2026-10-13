Attribute VB_Name = "Unit_TrapLifecycle"
' Native regression suite, driven by tools/test-trap-lifecycle.ps1.
Option Explicit
Private Total As Integer, Failed As Integer, Results As String

Public Sub Expect(ByVal Condition As Boolean, ByVal Description As String)
    If Not Condition Then Err.Raise 5, "Trap test", Description
End Sub
Public Function Read32(ByVal bytes As Variant, ByVal Offset As Integer) As Long
    Read32 = CDbl(bytes(Offset)) + CDbl(bytes(Offset + 1)) * 256# + _
             CDbl(bytes(Offset + 2)) * 65536# + CDbl(bytes(Offset + 3)) * 16777216#
End Function
Private Function EventAt(ByVal Index As Integer) As Byte
    Dim bytes As Variant
    bytes = Packets(Index)
    EventAt = bytes(2)
End Function
Private Sub ExpectPacket(ByVal Index As Integer, ByVal EventType As Byte, ByVal Id As Long, ByVal Lifetime As Long)
    Dim bytes As Variant
    bytes = Packets(Index)
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
    Call Expect(CreateTrap(1, eUser, 1, 11, 10, 7), "create explosive")
    Set trap = MapData(11, 10, 1).Trap
    Call ExpectPacket(1, eTrapUpsert, 1, 120000)
    Call Tick(1250)
    Call trap.SendSnapshot(2)
    Call trap.SendSnapshot(2)
    Call ExpectPacket(2, eTrapUpsert, 1, 118750)
    Call ExpectPacket(3, eTrapUpsert, 1, 118750)
    first = Packets(2): second = Packets(3)
    Dim i As Integer
    For i = 0 To 15
        Call Expect(first(i) = second(i), "idempotent snapshot")
    Next i
    Call Expect(first(7) = 1 And first(9) = 11 And first(10) = 10 And first(11) = 2, "snapshot location and subtype")
    Call Expect(UserList(1).EffectOverTime.EffectCount = 1, "snapshot does not fabricate effects")
End Sub
Private Sub TestTrigger(ByVal EffectId As Integer)
    Dim trap As clsTrap
    Call Expect(CreateTrap(1, eUser, 1, 11, 10, EffectId), "place")
    Set trap = MapData(11, 10, 1).Trap
    Call ActivateTrap(1, eNpc, 1, 11, 10)
    Call ActivateTrap(1, eNpc, 1, 11, 10)
    Call Expect(Not trap.trigger(1, eNpc), "duplicate direct trigger rejected")
    Call trap.SendSnapshot(2)
    Call Expect(Packets.Count = 2, "trigger once, no stale upsert")
    Call ExpectPacket(2, eTrapTriggered, 1, 0)
    Call Tick(1)
    Call trap.IBaseEffectOverTime_OnRemove()
    Call Expect(Packets.Count = 3, "remove once")
    Call ExpectPacket(3, eTrapRemoved, 1, 0)
    Call Expect(MapData(11, 10, 1).Trap Is Nothing, "tile released")
    Select Case EffectId
        Case 6: Call Expect(ParalysisCalls = 1 And SoundCalls = 1 And FxCalls = 1, "one paralysis and unchanged effects")
        Case 7: Call Expect(DamageCalls = 1 And SoundCalls = 2 And FxCalls = 1, "one explosion and unchanged effects")
        Case 9: Call Expect(PoisonCalls = 1 And SoundCalls = 1 And FxCalls = 0, "one poison and unchanged effects")
    End Select
End Sub
Private Sub TestRejections()
    Call Expect(CreateTrap(1, eUser, 1, 11, 10, 6), "place")
    Dim trap As clsTrap
    Set trap = MapData(11, 10, 1).Trap
    Call Expect(Not trap.trigger(2, eUser), "PvE only")
    Call Expect(Not trap.trigger(0, eNpc), "invalid NPC")
    AllowAttack = False
    Call ActivateTrap(1, eNpc, 1, 11, 10)
    AllowAttack = True
    NpcList(1).flags.AfectaParalisis = 1
    Call Expect(Not trap.trigger(1, eNpc), "immune NPC")
    Call Expect(Packets.Count = 1 And SoundCalls = 0 And ParalysisCalls = 0, "no fake activation")
    Call Expect(trap.IsActive, "rejections do not consume trap")
End Sub
Private Sub TestExpiry()
    Call Expect(CreateTrap(1, eUser, 1, 11, 10, 6), "place")
    Call Tick(9999)
    Call Expect(MapData(11, 10, 1).Trap.RemainingMilliseconds = 1, "last live millisecond")
    Call Tick(1)
    Call Tick(1)
    Call Expect(Packets.Count = 2, "expire once")
    Call ExpectPacket(2, eTrapRemoved, 1, 0)
    Call Expect(UserList(1).EffectOverTime.EffectCount = 0 And Recycled = 1, "EOT cleanup")
End Sub
Private Sub TestInvalidOwner()
    Call Expect(CreateTrap(1, eUser, 1, 11, 10, 6), "place")
    UserList(1).VersionId = 2
    Call MapData(11, 10, 1).Trap.SendSnapshot(2)
    Call ActivateTrap(1, eNpc, 1, 11, 10)
    Call Tick(1)
    Call Expect(Packets.Count = 2 And SoundCalls = 0, "invalid owner removed without activation")
    Call ExpectPacket(2, eTrapRemoved, 1, 0)
    Call Expect(CdClears = 0, "do not send old cooldown to recycled user slot")
End Sub
Private Sub TestDeath()
    Call Expect(CreateTrap(1, eUser, 1, 11, 10, 6), "place")
    UserList(1).flags.Muerto = 1
    Call ClearEffectList(UserList(1).EffectOverTime, eAny, True)
    Call Tick(1)
    Call Expect(Packets.Count = 2 And CdClears = 1, "death removes once")
    Call ExpectPacket(2, eTrapRemoved, 1, 0)
End Sub
Private Sub TestFourth()
    Dim i As Integer
    ' Equal-age traps must select the first active trap.
    For i = 11 To 14
        Call PlaceTrap(1, i, 10)
    Next i
    Call Expect(Packets.Count = 5, "three upserts then removed/upsert")
    Call ExpectPacket(4, eTrapRemoved, 1, 0)
    Call ExpectPacket(5, eTrapUpsert, 4, 10000)
    Call Expect(MapData(11, 10, 1).Trap Is Nothing, "oldest released immediately")
    Call Expect(Consumed = 4 And CdCalls = 4, "consumption/cooldowns unchanged")
    Call Tick(1)
    Call Expect(UserList(1).EffectOverTime.EffectCount = 3 And Packets.Count = 5, "no duplicate late removal")

    Call ResetFixture()
    Call PlaceTrap(1, 11, 10)
    Call MapData(11, 10, 1).Trap.Disable()
    ' An inactive entry at index zero must not win oldest-trap selection.
    For i = 11 To 14
        Call PlaceTrap(1, i, 10)
    Next i
    Call ExpectPacket(6, eTrapRemoved, 2, 0)
    Call ExpectPacket(7, eTrapUpsert, 5, 10000)
    Call Tick(1)
    Call Expect(UserList(1).EffectOverTime.EffectCount = 3 And Packets.Count = 7, "inactive entries excluded from cap")

    Call ResetFixture()
    Call PlaceTrap(1, 11, 10)
    Call MapData(11, 10, 1).Trap.Disable()
    Call PlaceTrap(1, 11, 10)
    Call Tick(1)
    Call Expect(Not MapData(11, 10, 1).Trap Is Nothing, "late cleanup preserves replacement on same tile")
    Call Expect(MapData(11, 10, 1).Trap.IBaseEffectOverTime_UniqueId = 2, "replacement identity preserved")
End Sub
Private Sub TestInventory()
    Dim i As Integer
    For i = 1 To 8
        Call ResetFixture()
        Select Case i
            Case 1: UserList(1).flags.UsingItemSlot = -1
            Case 2: UserList(1).flags.UsingItemSlot = 11
            Case 3: UserList(1).invent.Object(1).ObjIndex = 0
            Case 4: UserList(1).invent.Object(1).ObjIndex = 30000
            Case 5: UserList(1).invent.Object(1).amount = 0
            Case 6: UserList(1).invent.Object(1).ObjIndex = 2
            Case 7: ObjData(1).OBJType = e_OBJType.otWeapon
            Case 8: UserList(1).flags.UsingItemObjIndex = 0
        End Select
        Call UserTargetableItem(1, 11, 10)
        Call Expect(Packets.Count = 0 And Consumed = 0, "invalid inventory rejected")
        Call Expect(UserList(1).flags.UsingItemSlot = 0 And UserList(1).flags.UsingItemObjIndex = 0, "stale request cleared")
    Next i
    Call ResetFixture()
    Call UserTargetableItem(1, 11, 10)
    Call UserTargetableItem(1, 12, 10)
    Call Expect(Packets.Count = 1 And Consumed = 1, "one-shot item target submission")
End Sub
Private Sub TestPlacementGuards()
    Dim i As Integer
    For i = 1 To 7
        Call ResetFixture()
        Select Case i
            Case 1: ObjData(1).Subtipo = e_UssableOnTarget.eArpon
            Case 2: ObjData(1).EfectoMagico = 8
            Case 3: EffectOverTime(6).SubType = 4
            Case 4: Cooldown = True
            Case 5: UserList(1).flags.Muerto = 1
            Case 6: MapData(11, 10, 1).Blocked = 1
            Case 7: ObjData(1).EfectoMagico = 30000
        End Select
        Call PlaceTrap(1, 11, 10)
        Call Expect(Packets.Count = 0 And Consumed = 0 And CdCalls = 0, "invalid placement no mutation")
    Next i
    Call ResetFixture()
    Call PlaceTrap(1, 15, 10)
    Call PlaceTrap(1, 0, 10)
    Call Expect(Packets.Count = 0 And Consumed = 0, "range/bounds unchanged")
End Sub
Private Sub TestReuseAndClamp()
    Dim trap As New clsTrap
    Call trap.Setup(1, eUser, 7, 0, 1, 11, 10)
    Call trap.Setup(1, eUser, 8, 43, 1, 11, 10)
    Call trap.Setup(1, eNone, 7, 43, 1, 11, 10)
    Call Expect(Packets.Count = 0 And Not trap.IsActive, "invalid setup never published")
    Call trap.Setup(1, eUser, 7, 44, 1, 11, 10)
    Call trap.IBaseEffectOverTime_Update(-100)
    Call Expect(trap.RemainingMilliseconds = 120000, "ignore negative delta")
    Call trap.Disable()
    Call trap.Setup(1, eUser, 6, 45, 1, 12, 10)
    Call ExpectPacket(3, eTrapUpsert, 45, 10000)
    Call trap.IBaseEffectOverTime_Update(&H7FFFFFFF)
    Call Expect(trap.RemainingMilliseconds = 0, "long delta clamps without overflow")
    Call trap.IBaseEffectOverTime_OnRemove()
    Call ExpectPacket(4, eTrapRemoved, 45, 0)
End Sub
Private Sub TestIdWrap()
    Call Expect(NextIdAfter(0) = 1, "positive initial id")
    Call Expect(NextIdAfter(&H7FFFFFFE) = &H7FFFFFFF, "maximum signed id")
    Call Expect(NextIdAfter(&H7FFFFFFF) = 1, "overflow skips zero")
End Sub
Private Sub RunCase(ByVal Number As Integer, ByVal Name As String)
    On Error GoTo FailedCase
    Total = Total + 1
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
    Results = Results & Name & ": PASS" & vbCrLf
    Exit Sub
FailedCase:
    Failed = Failed + 1
    Results = Results & Name & ": FAIL " & Err.Source & " " & Err.Description & vbCrLf
End Sub
Public Sub Main()
    Dim names As Variant, i As Integer, f As Integer
    names = Array("Wire", "Snapshots", "Paralysis", "Explosion", "Poison", "Rejection/immunity", _
                  "Expiry", "Invalid owner", "Death", "Fourth trap", "Inventory", "Placement guards", "Pool reuse/time clamp", "Id wrap")
    For i = 0 To UBound(names)
        Call RunCase(i + 1, names(i))
    Next i
    f = FreeFile
    Open App.Path & "\results.txt" For Output As #f
    Print #f, Results
    Print #f, "Total=" & Total & " Failed=" & Failed
    Close #f
End Sub
