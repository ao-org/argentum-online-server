Attribute VB_Name = "Unit_InventoryDestruction"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private testUser As Integer
Private Const TEST_NOW As Long = 1000

Public Sub InitializeTestWorld()
    Call InitializeCircularLogBuffer()
    ReDim UserList(1 To 2)
    ReDim ObjData(0 To 4000)
    ReDim MapInfo(1 To 1)
    ReDim MapData(1 To 100, 1 To 100, 1 To 1)
    NumMaps = 1
    Call InitializeAuxiliaryBuffer()
End Sub

Public Function CapturePacket(ByVal Route As Long, ByVal Recipient As Integer) As Boolean
    If Not capturing Then Exit Function
    Dim bytes() As Byte, writer As Network.writer, value As Variant
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    value = bytes
    Call packets.Add(value)
    CapturePacket = True
End Function

Private Sub Expect(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 180, "Inventory destruction test", detail
End Sub

Private Sub Fixture()
    Dim emptyUser As t_User, emptyObject As t_ObjData
    UserList(testUser) = emptyUser
    ObjData(1) = emptyObject
    ObjData(2) = emptyObject
    ObjData(1).OBJType = e_OBJType.otPotions
    ObjData(2).OBJType = e_OBJType.otArrows
    With UserList(testUser)
        .name = "destruction fixture"
        .Id = 101
        .VersionId = 7
        .flags.UserLogged = True
        .ConnectionDetails.ConnIDValida = True
        .CurrentInventorySlots = 20
        .pos.Map = 1: .pos.x = 50: .pos.y = 50
        .invent.NroItems = 1
        .invent.Object(1).ObjIndex = 1
        .invent.Object(1).amount = 10
        .Stats.GLD = 731
    End With
    Set packets = New Collection
End Sub

Private Function Prepare(ByRef reply As t_InventoryDestruction, Optional ByVal quantity As Long = 3, _
        Optional ByVal slot As Byte = 1, Optional ByVal requestId As Long = 123, _
        Optional ByVal tick As Long = TEST_NOW) As e_InventoryDestructionCode
    Dim expected As t_UserOBJ
    expected = UserList(testUser).invent.Object(1)
    If expected.ObjIndex > 0 Then expected.ElementalTags = expected.ElementalTags Or ObjData(expected.ObjIndex).ElementalTags
    Prepare = PrepareInventoryDestruction(testUser, requestId, slot, expected, quantity, tick, reply)
End Function

Private Function Confirm(ByRef prepared As t_InventoryDestruction, ByRef reply As t_InventoryDestruction, _
        Optional ByVal tick As Long = TEST_NOW + 1) As e_InventoryDestructionCode
    Dim token(0 To 15) As Byte, i As Integer
    For i = 0 To 15
        token(i) = prepared.Token(i)
    Next i
    Confirm = ConfirmInventoryDestruction(testUser, token, tick, reply)
End Function

Private Sub MoveSlot(ByVal source As Byte, ByVal destination As Byte)
    Dim payload(0 To 1) As Byte
    payload(0) = source: payload(1) = destination
    Call Protocol.DestructionTestMoveItem(testUser, payload)
End Sub

Private Sub TestCase(ByVal index As Integer)
    Dim prepared As t_InventoryDestruction, reply As t_InventoryDestruction
    Dim expected As t_UserOBJ, saved As t_UserOBJ, quantity As Long, i As Integer
    Dim bytes() As Byte, inputWriter As New Network.writer, inputReader As New Network.reader
    Select Case index
        Case 1, 2, 3
            quantity = Choose(index, 1, 3, 10)
            Call Expect(Prepare(prepared, quantity) = DestroyOkay, "prepare valid quantity")
            Call Expect(UserList(testUser).invent.Object(1).amount = 10, "prepare never removes")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "confirm succeeds")
            Call Expect(UserList(testUser).invent.Object(1).amount = 10 - quantity, "exact removal")
            Call Expect(UserList(testUser).Stats.GLD = 731, "no reward")
            Call Expect(UserList(testUser).flags.ModificoInventario, "normal persistence dirty flag")
            Call Expect(MapData(50, 50, 1).ObjInfo.ObjIndex = 0, "no ground object")
            If quantity = 10 Then
                Call Expect(UserList(testUser).invent.Object(1).ObjIndex = 0 And UserList(testUser).invent.NroItems = 0, "empty slot cleanup")
            End If
        Case 4
            ObjData(1).Newbie = 1: ObjData(1).Intirable = 1: ObjData(1).Instransferible = 1
            Call Expect(Prepare(prepared) = DestroyOkay, "drop restrictions do not restrict destruction")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "restricted item removed")
        Case 5
            For i = 0 To 1
                Call Expect(Prepare(prepared, 1, CByte(i * 255), 123 + i, TEST_NOW + i * 1000&) = DestroyInvalidRequest, "slot bounds")
            Next i
            UserList(testUser).invent.Object(1).ObjIndex = 0
            Call Expect(Prepare(prepared, 1, 1, 125, TEST_NOW + 2000) = DestroyInvalidRequest, "empty slot")
        Case 6
            For i = 0 To 3
                quantity = Choose(i + 1, 0, -1, 11, 2147483647)
                Call Expect(Prepare(prepared, quantity, 1, 123 + i, TEST_NOW + i * 1000&) = DestroyInvalidRequest, "validate Long before Integer narrowing")
            Next i
            Call Expect(UserList(testUser).invent.Object(1).amount = 10, "invalid amounts no removal")
        Case 7
            UserList(testUser).invent.Object(1).ObjIndex = 2
            UserList(testUser).invent.Object(1).Equipped = 1
            UserList(testUser).invent.EquippedMunitionObjIndex = 2
            UserList(testUser).invent.EquippedMunitionSlot = 1
            Call Expect(Prepare(prepared, 10) = DestroyOkay, "equipped prepare")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "final equipped remove")
            Call Expect(UserList(testUser).invent.EquippedMunitionObjIndex = 0 And UserList(testUser).invent.EquippedMunitionSlot = 0, "equipment pointers cleared")
            Call Expect(UserList(testUser).invent.Object(1).Equipped = 0, "unequipped")
        Case 8
            UserList(testUser).flags.Comerciando = True
            Call Expect(Prepare(prepared) = DestroyBusy, "trade rejected")
            Call Fixture()
            UserList(testUser).flags.Crafteando = 1
            Call Expect(Prepare(prepared) = DestroyBusy, "craft rejected")
            Call Fixture()
            UserList(testUser).Accion.AccionPendiente = True
            Call Expect(Prepare(prepared) = DestroyBusy, "pending action rejected")
        Case 9, 10
            If index = 9 Then
                ObjData(1).OBJType = otShips
                UserList(testUser).flags.Navegando = 1
                UserList(testUser).invent.EquippedShipSlot = 1
                Call Expect(Prepare(prepared) = DestroyActiveBoat, "active boat rejected")
                UserList(testUser).invent.EquippedShipSlot = 2
            Else
                ObjData(1).OBJType = otSaddles
                UserList(testUser).flags.Montado = 1
                UserList(testUser).invent.EquippedSaddleSlot = 1
                Call Expect(Prepare(prepared) = DestroyActiveMount, "active mount rejected")
                UserList(testUser).invent.EquippedSaddleSlot = 2
            End If
            Call Expect(Prepare(prepared, 3, 1, 124, TEST_NOW + 1000) = DestroyOkay, "spare stack allowed")
            Call Expect(Confirm(prepared, reply, TEST_NOW + 1001) = DestroyOkay, "spare removed")
        Case 11
            UserList(testUser).invent.Object(1).ObjIndex = OBJ_CAPTURA_BANDERA_1
            Call Expect(Prepare(prepared) = DestroyEventObjective, "capture flag protected")
        Case 12
            expected = UserList(testUser).invent.Object(1)
            UserList(testUser).invent.Object(1).ObjIndex = 2
            Call Expect(PrepareInventoryDestruction(testUser, 123, 1, expected, 3, TEST_NOW, reply) = DestroyStackChanged, "pre-prepare replacement")
        Case 13, 14, 15, 16
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare snapshot")
            Select Case index
                Case 13: UserList(testUser).invent.Object(1).ObjIndex = 2
                Case 14: UserList(testUser).invent.Object(1).amount = 11
                Case 15: UserList(testUser).invent.Object(1).ElementalTags = 1
                Case 16: UserList(testUser).invent.Object(1).Equipped = 1
            End Select
            Call Expect(Confirm(prepared, reply) = DestroyStackChanged, "snapshot mismatch rejected")
        Case 17
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare move")
            Call MoveSlot(1, 2)
            Call MoveSlot(2, 1)
            Call Expect(UserList(testUser).invent.Object(1).ObjIndex = 1 And UserList(testUser).invent.Object(1).amount = 10, "real move restored identical stack")
            Call Expect(Confirm(prepared, reply) = DestroyStackChanged, "identical replacement invalidates")
        Case 18
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare change")
            saved = UserList(testUser).invent.Object(1)
            Call ChangeUserInv(testUser, 1, saved)
            Call Expect(Confirm(prepared, reply) = DestroyStackChanged, "real ChangeUserInv invalidates equal bytes")
        Case 19
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare expiry")
            Call Expect(Confirm(prepared, reply, TEST_NOW + 30000) = DestroyExpired, "exact thirty second expiry")
            Call Expect(Confirm(prepared, reply, TEST_NOW + 30001) = DestroyExpired, "duplicate expired result")
        Case 20
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare duplicate")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "first confirm")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "duplicate result")
            Call Expect(UserList(testUser).invent.Object(1).amount = 7, "one removal only")
        Case 21
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare session")
            Call IncreaseVersionId(testUser)
            Call Expect(Confirm(prepared, reply) = DestroyInvalidToken And reply.RequestId = 0, "cross-generation rejected")
            Call Expect(Not UserList(testUser).InventoryDestruction.Prepared, "generation discards preparation")
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare before character replacement")
            UserList(testUser).Id = 102
            Call Expect(Confirm(prepared, reply) = DestroyInvalidToken, "cross-character rejected")
        Case 22
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare teardown")
            UserList(testUser).ConnectionDetails.ConnIDValida = False
            Call CloseSocketSL(testUser)
            Call Expect(Confirm(prepared, reply) = DestroyInvalidToken, "teardown token discarded")
        Case 23
            Call Expect(Not UserList(testUser).HooCapabilities.Negotiated, "fixture has no negotiated capabilities")
            Call Expect(UserList(testUser).HooCapabilities.CapabilityMask = 0, "fixture has no capability bits")
            Call Expect(Prepare(prepared) = DestroyOkay, "required protocol prepares without negotiation")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "required protocol confirms without negotiation")
            Call Expect(UserList(testUser).invent.Object(1).amount = 7, "unnegotiated confirmation removes exact quantity")
        Case 24
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare rate")
            Call Expect(Prepare(reply, 1, 1, 124) = DestroyRateLimited, "bounded preparation rate")
            Call Expect(Confirm(prepared, reply) = DestroyInvalidToken, "rejected replacement invalidates previous token")
        Case 25
            UserList(testUser).flags.Muerto = 1
            Call Expect(Prepare(prepared) = DestroyDead, "dead cannot prepare")
            Call Fixture()
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare alive")
            UserList(testUser).flags.Muerto = 1
            Call Expect(Confirm(prepared, reply) = DestroyDead, "death before confirm")
        Case 26
            ObjData(1).OBJType = otKeys
            Call Expect(Prepare(prepared) = DestroyOkay, "key is ordinary inventory")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "key removal has no database operation")
        Case 27
            ObjData(1).ElementalTags = 2
            UserList(testUser).invent.Object(1).ElementalTags = 1
            Call Expect(Prepare(prepared) = DestroyOkay, "wire union tags accepted")
            Call Expect(prepared.Item.ElementalTags = 1 And prepared.EffectiveElementalTags = 3, "raw and display snapshots separate")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "natural tags don't block destruction")
        Case 28
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare wire")
            Call WriteInventoryDestructionResult(testUser, 0, prepared)
            bytes = packets(packets.Count)
            Call Expect(UBound(bytes) = 39, "result exactly forty bytes")
            Call inputReader.SetData(bytes)
            Call Expect(inputReader.ReadInt16() = ServerPacketID.eHooInventoryDestructionResult, "append-only packet ID")
            Call Expect(inputReader.ReadInt32() = 123, "request ID order")
            For i = 0 To 15
                Call Expect(inputReader.ReadInt8() = prepared.Token(i), "token bytes")
            Next i
            Call Expect(inputReader.ReadInt8() = 0 And inputReader.ReadInt8() = 0, "stage/code")
            Call Expect(inputReader.ReadInt8() = 1 And inputReader.ReadInt16() = 1 And inputReader.ReadInt16() = 10, "slot/object/original amount")
            Call Expect(inputReader.ReadInt8() = 0 And inputReader.ReadInt32() = 0 And inputReader.ReadInt32() = 3 And inputReader.ReadInt16() = 30, "equipped/tags/quantity/expiry")
            #If PYMMO = 1 Then
                Call Expect(ServerPacketID.eHooInventoryDestructionResult = 210, "production result ID")
                Call Expect(ClientPacketID.eHooInventoryDestructionPrepare = 313 And ClientPacketID.eHooInventoryDestructionConfirm = 314, "production request IDs")
            #Else
                Call Expect(ServerPacketID.eHooInventoryDestructionResult = 211, "legacy enum result ordinal")
                Call Expect(ClientPacketID.eHooInventoryDestructionPrepare = 316 And ClientPacketID.eHooInventoryDestructionConfirm = 317, "conditional request ordinals")
            #End If
        Case 29
            Call inputWriter.WriteInt32(123)
            Call inputWriter.WriteInt8(1)
            Call inputWriter.WriteInt16(1)
            Call inputWriter.WriteInt16(10)
            Call inputWriter.WriteInt8(0)
            Call inputWriter.WriteInt32(0)
            Call inputWriter.WriteInt32(3)
            Call inputWriter.GetData(bytes)
            Call Expect(UBound(bytes) = 17, "prepare payload eighteen bytes")
            Call inputReader.SetData(bytes)
            Call HandleInventoryDestructionPrepare(testUser, inputReader)
            Call Expect(UserList(testUser).InventoryDestruction.Prepared, "real prepare parser")
            prepared = UserList(testUser).InventoryDestruction
            Call inputWriter.Clear()
            For i = 0 To 15
                Call inputWriter.WriteInt8(prepared.Token(i))
            Next i
            Call inputWriter.GetData(bytes)
            Call inputReader.SetData(bytes)
            Call HandleInventoryDestructionConfirm(testUser, inputReader)
            Call Expect(UserList(testUser).invent.Object(1).amount = 7, "real confirm parser")
        Case 30
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare malformed")
            ReDim bytes(0 To 14)
            Call inputReader.SetData(bytes)
            Call HandleInventoryDestructionConfirm(testUser, inputReader)
            ReDim bytes(0 To 16)
            Call inputReader.SetData(bytes)
            Call HandleInventoryDestructionConfirm(testUser, inputReader)
            ReDim bytes(0 To 18)
            Call inputReader.SetData(bytes)
            Call HandleInventoryDestructionPrepare(testUser, inputReader)
            Call Expect(UserList(testUser).InventoryDestruction.Prepared And UserList(testUser).invent.Object(1).amount = 10, "malformed lengths no partial mutation")
        Case 31
            Call Expect(Prepare(prepared, 3, 1, 123, &H7FFFFFF0) = DestroyOkay, "uptime sign-boundary prepare")
            Call Expect(Confirm(prepared, reply, &H80000010) = DestroyOkay, "wrap-safe expiry")
        Case 32
            UserList(testUser).invent.Object(1).ObjIndex = 2
            UserList(testUser).invent.Object(1).Equipped = 1
            UserList(testUser).invent.EquippedMunitionSlot = 1
            UserList(testUser).invent.EquippedMunitionObjIndex = 2
            Call Expect(Prepare(prepared, 3) = DestroyOkay, "partial equipped prepare")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "partial equipped removal")
            Call Expect(UserList(testUser).invent.Object(1).Equipped = 1 And UserList(testUser).invent.EquippedMunitionSlot = 1, "partial stack remains equipped")
        Case 33
            ObjData(1).OBJType = otAmulets
            ObjData(1).EfectoMagico = e_MagicItemEffect.eModifyAttributes
            ObjData(1).QueAtributo = 255
            ObjData(1).CuantoAumento = 1
            UserList(testUser).invent.Object(1).Equipped = 1
            Call Expect(Prepare(prepared, 10) = DestroyOkay, "prepare broken equipment definition")
            Call Expect(Confirm(prepared, reply) = DestroyInternalFailure, "swallowed unequip error is not success")
            Call Expect(UserList(testUser).invent.Object(1).amount = 10, "unequip error prevents removal")
            Call Expect(Confirm(prepared, reply) = DestroyInternalFailure, "failure cached without rerunning unequip")
        Case 34
            Call Expect(Prepare(prepared) = DestroyOkay, "first preparation")
            Call Expect(Prepare(reply, 2, 1, 124, TEST_NOW + 1000) = DestroyOkay, "replacement preparation")
            saved = UserList(testUser).invent.Object(1)
            Call Expect(Confirm(prepared, reply) = DestroyInvalidToken, "old token cannot consume replacement")
            Call Expect(UserList(testUser).InventoryDestruction.RequestId = 124 And UserList(testUser).InventoryDestruction.Prepared, "new preparation intact")
            prepared = UserList(testUser).InventoryDestruction
            Call Expect(Confirm(prepared, reply, TEST_NOW + 1001) = DestroyOkay, "replacement quantity applied")
            Call Expect(UserList(testUser).invent.Object(1).amount = 8, "only new exact quantity removed")
        Case 35
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare before trade")
            UserList(testUser).flags.Comerciando = True
            Call Expect(Confirm(prepared, reply) = DestroyBusy, "trade checked at confirmation")
            UserList(testUser).flags.Comerciando = False
            Call Expect(Confirm(prepared, reply) = DestroyBusy, "terminal rejection cannot resume")
        Case 36
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare before disconnect")
            UserList(testUser).ConnectionDetails.ConnIDValida = False
            Call Expect(Confirm(prepared, reply) = DestroyUnavailable, "disconnected character unavailable")
            Call Expect(UserList(testUser).invent.Object(1).amount = 10, "no disconnected removal")
        Case 37
            UserList(testUser).invent.Object(1).Equipped = 1
            Call Expect(Prepare(prepared) = DestroyOkay, "equipped wire prepare")
            Call WriteInventoryDestructionResult(testUser, 0, prepared)
            bytes = packets(packets.Count)
            Call Expect(bytes(29) = 1, "equipped byte is one, not VB6 Boolean minus one")
            prepared.Result = DestroyStackChanged
            Call WriteInventoryDestructionResult(testUser, 0, prepared)
            bytes = packets(packets.Count)
            Call Expect(bytes(38) = 0 And bytes(39) = 0, "error expiry zero")
        Case 38
            Call Expect(Prepare(prepared) = DestroyOkay, "prepare consume")
            saved = UserList(testUser).invent.Object(1)
            Call QuitarUserInvItem(testUser, 1, 1)
            UserList(testUser).invent.Object(1) = saved
            Call Expect(Confirm(prepared, reply) = DestroyStackChanged, "consumption invalidates even if same stack bytes restored")
        Case 39
            UserList(testUser).invent.EquippedShipSlot = 1
            UserList(testUser).invent.EquippedShipObjIndex = 1
            ObjData(1).OBJType = otShips
            Call Expect(Prepare(prepared, 10) = DestroyOkay, "inactive ship allowed")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "inactive ship removed")
            Call Expect(UserList(testUser).invent.EquippedShipSlot = 0 And UserList(testUser).invent.EquippedShipObjIndex = 0, "inactive transport references cleared")
        Case 40, 41
            UserList(testUser).flags.UsingItemSlot = CByte(index - 39)
            UserList(testUser).flags.UsingItemObjIndex = 1
            UserList(testUser).flags.UsingItemMap = 1
            Call Expect(Prepare(prepared) = DestroyOkay, "abandoned target cursor is not a transaction lock")
            Call Expect(Confirm(prepared, reply) = DestroyOkay, "explicit destruction still works")
            If index = 40 Then
                Call Expect(UserList(testUser).flags.UsingItemSlot = 0 And UserList(testUser).flags.UsingItemObjIndex = 0, "same-slot use cancelled before removal")
            Else
                Call Expect(UserList(testUser).flags.UsingItemSlot = 2 And UserList(testUser).flags.UsingItemObjIndex = 1, "unrelated targeting preserved")
            End If
    End Select
End Sub

Private Sub RunCase(ByVal index As Integer, ByVal name As String)
    On Error GoTo Failed
    Call Fixture()
    Call TestCase(index)
    Call UnitTesting.RunTest("inventory destruction " & name, True)
    Exit Sub
Failed:
    Call UnitTesting.RunTestError("inventory destruction " & name, CStr(Err.Number) & ": " & Err.Description)
End Sub

Public Function test_suite_inventory_destruction() As Boolean
    On Error GoTo Failed
    Dim oldUser As t_User, oldObject1 As t_ObjData, oldObject2 As t_ObjData
    Dim i As Long, before As Long, saved As Boolean, names As Variant
    before = UnitTesting.GetFailedTests()
    For i = UBound(UserList) To 1 Step -1
        If Not UserList(i).flags.UserLogged And Not UserList(i).ConnectionDetails.ConnIDValida Then
            testUser = i
            Exit For
        End If
    Next i
    Call Expect(testUser > 0, "idle fixture user")
    oldUser = UserList(testUser): oldObject1 = ObjData(1): oldObject2 = ObjData(2): saved = True
    capturing = True
    names = Array("single", "partial", "entire", "drop flags", "slots", "quantities", "equipment cleanup", _
        "transactions", "boat", "mount", "objective", "stale prepare", "replacement", "amount", "tags", "equipped state", _
        "move out and back", "ChangeUserInv", "expiry", "duplicate", "session identity", "teardown", "unnegotiated protocol", _
        "rate limit", "death", "key", "natural tags", "wire", "parsers", "malformed", "tick wrap", "partial equipment", _
        "unequip error", "replacement prepare", "trade after prepare", "disconnect", "equipped and error wire", _
        "consume restored stack", "inactive transport cleanup", "abandoned same-slot targeting", "unrelated targeting")
    For i = 0 To UBound(names)
        Call RunCase(i + 1, CStr(names(i)))
    Next i
    test_suite_inventory_destruction = UnitTesting.GetFailedTests() = before
Cleanup:
    capturing = False
    If saved Then
        UserList(testUser) = oldUser: ObjData(1) = oldObject1: ObjData(2) = oldObject2
    End If
    Exit Function
Failed:
    Call UnitTesting.RunTestError("inventory destruction suite", CStr(Err.Number) & ": " & Err.Description)
    Resume Cleanup
End Function
#End If
