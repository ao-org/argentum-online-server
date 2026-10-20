Attribute VB_Name = "modInventoryDestruction"
Option Explicit

Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef FirstByte As Byte) As Long
Private Const TOKEN_LIFETIME_MS As Long = 30000
Private Const PREPARE_INTERVAL_MS As Long = 500

' The state lives in t_User, never in persistence. Session resets discard both
' the prepared token and its bounded, last-confirmation replay result.
Public Sub ClearInventoryDestruction(ByVal UserIndex As Integer)
    Dim emptyState As t_InventoryDestruction
    UserList(UserIndex).InventoryDestruction = emptyState
End Sub

' This is a mutation latch, not a byte comparison: moving a stack out and back
' invalidates its token even if its final object/amount/tags are identical.
Public Sub InvalidateInventoryDestruction(ByVal UserIndex As Integer, Optional ByVal Slot As Integer = 0)
    With UserList(UserIndex).InventoryDestruction
        If Not .Prepared Then Exit Sub
        If Slot = 0 Or Slot = .Slot Then .Invalidated = True
    End With
End Sub

Private Function SameStack(ByRef leftItem As t_UserOBJ, ByRef rightItem As t_UserOBJ) As Boolean
    SameStack = leftItem.ObjIndex = rightItem.ObjIndex And _
                leftItem.amount = rightItem.amount And _
                leftItem.Equipped = rightItem.Equipped And _
                leftItem.ElementalTags = rightItem.ElementalTags
End Function

Private Function StateRejection(ByVal UserIndex As Integer, ByVal Slot As Byte) As e_InventoryDestructionCode
    On Error GoTo Failed
    StateRejection = DestroyUnavailable
    If UserIndex < 1 Or UserIndex > UBound(UserList) Then Exit Function
    With UserList(UserIndex)
        If Not .flags.UserLogged Or Not .ConnectionDetails.ConnIDValida Then Exit Function
        If .flags.YaGuardo Or .Counters.Salir > 0 Then Exit Function
        StateRejection = DestroyDead
        If .flags.Muerto <> 0 Then Exit Function
        StateRejection = DestroyBusy
        If .flags.Comerciando Or .flags.Subastando Or .flags.Crafteando <> 0 Then Exit Function
        If .ComUsu.DestUsu.ArrayIndex <> 0 Or .Accion.AccionPendiente Or .flags.PescandoEspecial Then Exit Function
        If .AutomatedAction.IsActive Then Exit Function
        StateRejection = DestroyInvalidRequest
        If Slot < 1 Or Slot > .CurrentInventorySlots Then Exit Function
        If Slot > UBound(.invent.Object) Then Exit Function
        Dim itemId As Integer
        itemId = .invent.Object(Slot).ObjIndex
        If itemId < 1 Or itemId > UBound(ObjData) Then Exit Function
        If .invent.Object(Slot).amount < 1 Or .invent.Object(Slot).Equipped > 1 Then Exit Function
        ' Gold and collection entries are not ordinary inventory destruction.
        If ObjData(itemId).OBJType = e_OBJType.otGoldCoin Then Exit Function
        If (.flags.Navegando <> 0 Or .flags.Nadando <> 0) And .invent.EquippedShipSlot = Slot Then
            StateRejection = DestroyActiveBoat
            Exit Function
        End If
        If .flags.Montado <> 0 And .invent.EquippedSaddleSlot = Slot Then
            StateRejection = DestroyActiveMount
            Exit Function
        End If
        StateRejection = DestroyEventObjective
        If itemId = OBJ_CAPTURA_BANDERA_1 Or itemId = OBJ_CAPTURA_BANDERA_2 Then Exit Function
        ' Inventory-owning scenarios (naval cargo) have pickup/drop bookkeeping,
        ' but no safe permanent-removal hook. Do not invoke drop to emulate it.
        Dim scenario As IBaseScenario
        Set scenario = CustomScenarios.GetMap(.pos.Map)
        If Not scenario Is Nothing Then
            If IsSet(scenario.GetExtraInterfaces(), eInventoryInterface) Then Exit Function
        End If
    End With
    StateRejection = DestroyOkay
    Exit Function
Failed:
    StateRejection = DestroyInternalFailure
    Call TraceError(Err.Number, Err.Description, "InventoryDestruction.StateRejection", Erl)
End Function

Public Function PrepareInventoryDestruction(ByVal UserIndex As Integer, ByVal RequestId As Long, _
        ByVal Slot As Byte, ByRef Expected As t_UserOBJ, ByVal Quantity As Long, _
        ByVal NowTick As Long, ByRef reply As t_InventoryDestruction) As e_InventoryDestructionCode
    On Error GoTo Failed
    Dim previous As t_InventoryDestruction, code As e_InventoryDestructionCode
    Dim emptyState As t_InventoryDestruction
    previous = UserList(UserIndex).InventoryDestruction
    reply = emptyState
    reply.RequestId = RequestId
    reply.Slot = Slot
    code = DestroyInvalidRequest
    ' Even a rejected replacement invalidates the old prepared token.
    Call ClearInventoryDestruction(UserIndex)
    With UserList(UserIndex).InventoryDestruction
        .HasPrepareTime = True
        .LastPrepareAt = NowTick
    End With
    If RequestId <= 0 Then GoTo Done
    If previous.HasPrepareTime Then
        If TicksElapsed(previous.LastPrepareAt, NowTick) < PREPARE_INTERVAL_MS Then
            code = DestroyRateLimited
            GoTo Done
        End If
    End If
    code = StateRejection(UserIndex, Slot)
    If code <> DestroyOkay Then GoTo Done
    If Quantity < 1 Or Quantity > CLng(UserList(UserIndex).invent.Object(Slot).amount) Then
        code = DestroyInvalidRequest
        GoTo Done
    End If
    Dim visibleItem As t_UserOBJ
    visibleItem = UserList(UserIndex).invent.Object(Slot)
    visibleItem.ElementalTags = visibleItem.ElementalTags Or ObjData(visibleItem.ObjIndex).ElementalTags
    If Not SameStack(Expected, visibleItem) Then
        code = DestroyStackChanged
        GoTo Done
    End If
    reply.Item = UserList(UserIndex).invent.Object(Slot)
    reply.EffectiveElementalTags = visibleItem.ElementalTags
    reply.Quantity = Quantity
    reply.OwnerVersion = UserList(UserIndex).VersionId
    reply.OwnerId = UserList(UserIndex).Id
    reply.PreparedAt = NowTick
    reply.LastPrepareAt = NowTick
    reply.HasPrepareTime = True
    If CoCreateGuid(reply.Token(0)) <> 0 Then
        code = DestroyInternalFailure
        GoTo Done
    End If
    reply.Prepared = True
    UserList(UserIndex).InventoryDestruction = reply
Done:
    reply.Result = code
    PrepareInventoryDestruction = code
    Exit Function
Failed:
    Call ClearInventoryDestruction(UserIndex)
    reply.Result = DestroyInternalFailure
    PrepareInventoryDestruction = DestroyInternalFailure
    Call TraceError(Err.Number, Err.Description, "InventoryDestruction.Prepare", Erl)
End Function

Public Function ConfirmInventoryDestruction(ByVal UserIndex As Integer, ByRef Token() As Byte, _
        ByVal NowTick As Long, ByRef reply As t_InventoryDestruction) As e_InventoryDestructionCode
    On Error GoTo Failed
    Dim emptyState As t_InventoryDestruction, i As Integer, matched As Boolean
    Dim code As e_InventoryDestructionCode, completed As Boolean
    reply = emptyState
    For i = 0 To 15
        reply.Token(i) = Token(i)
    Next i
    code = DestroyInvalidToken
    With UserList(UserIndex).InventoryDestruction
        If .RequestId <= 0 Then GoTo Done
        If .OwnerVersion <> UserList(UserIndex).VersionId Or .OwnerId <> UserList(UserIndex).Id Then GoTo Done
        For i = 0 To 15
            If .Token(i) <> Token(i) Then GoTo Done
        Next i
        matched = True
        reply = UserList(UserIndex).InventoryDestruction
        If .Completed Then
            code = .Result
            GoTo Done
        End If
        ' Consume before mutation, including ordinary rejections. There is never
        ' a second removal even if a send fails or the client repeats confirm.
        .Prepared = False
        .Completed = True
        .Result = DestroyInternalFailure
    End With
    code = DestroyExpired
    If TicksElapsed(reply.PreparedAt, NowTick) >= TOKEN_LIFETIME_MS Then GoTo Done
    code = DestroyStackChanged
    If reply.Invalidated Then GoTo Done
    code = StateRejection(UserIndex, reply.Slot)
    If code <> DestroyOkay Then GoTo Done
    code = DestroyStackChanged
    If Not SameStack(reply.Item, UserList(UserIndex).invent.Object(reply.Slot)) Then GoTo Done
    If (reply.Item.ElementalTags Or ObjData(reply.Item.ObjIndex).ElementalTags) <> reply.EffectiveElementalTags Then GoTo Done
    code = DestroyInternalFailure
    ' Existing unequip may intentionally refuse (e.g. special fishing) or catch
    ' an error. Both completion and final equipped state must be checked.
    If reply.Quantity = reply.Item.amount And reply.Item.Equipped = 1 Then
        Call Desequipar(UserIndex, reply.Slot, False, 0, completed)
        If Not completed Then GoTo Done
        If UserList(UserIndex).invent.Object(reply.Slot).Equipped <> 0 Then GoTo Done
    End If
    completed = False
    ' Target cursors do not reserve inventory and can outlive a client-side
    ' Escape. Cancel only the source-slot use before removal so a delayed
    ' WorkLeftClick cannot consume the remaining stack through an old request.
    If UserList(UserIndex).flags.UsingItemSlot = reply.Slot Then Call ClearPendingTargetableItem(UserIndex)
    Call QuitarUserInvItem(UserIndex, reply.Slot, CInt(reply.Quantity), completed)
    If Not completed Then GoTo Done
    If UserList(UserIndex).invent.Object(reply.Slot).amount <> reply.Item.amount - reply.Quantity Then GoTo Done
    If reply.Quantity = reply.Item.amount Then
        With UserList(UserIndex).invent
            If .EquippedShipSlot = reply.Slot Then
                .EquippedShipSlot = 0
                .EquippedShipObjIndex = 0
            End If
            If .EquippedSaddleSlot = reply.Slot Then
                .EquippedSaddleSlot = 0
                .EquippedSaddleObjIndex = 0
            End If
        End With
    End If
    code = DestroyOkay
    ' Store the outcome before logging/networking: those may fail independently.
    UserList(UserIndex).InventoryDestruction.Result = code
    Call UpdateUserInv(False, UserIndex, reply.Slot)
    Call LogInfoServidor("Inventory destruction characterId=" & CStr(UserList(UserIndex).Id) & _
        " character=" & UserList(UserIndex).name & " object=" & CStr(reply.Item.ObjIndex) & _
        " elementalTags=" & CStr(reply.Item.ElementalTags) & " quantity=" & CStr(reply.Quantity) & _
        " request=" & CStr(reply.RequestId))
Done:
    If matched Then UserList(UserIndex).InventoryDestruction.Result = code
    reply.Result = code
    ConfirmInventoryDestruction = code
    Exit Function
Failed:
    ' No exception may re-arm a consumed token or fabricate success.
    code = DestroyInternalFailure
    If matched Then
        UserList(UserIndex).InventoryDestruction.Prepared = False
        UserList(UserIndex).InventoryDestruction.Completed = True
        code = UserList(UserIndex).InventoryDestruction.Result
    End If
    reply.Result = code
    ConfirmInventoryDestruction = code
    Call TraceError(Err.Number, Err.Description, "InventoryDestruction.Confirm", Erl)
End Function

Public Sub HandleInventoryDestructionPrepare(ByVal UserIndex As Integer, ByVal message As Network.reader)
    On Error GoTo Failed
    ' Validate the entire fixed payload before reading or changing token state.
    If message.GetAvailable() <> 18 Then Exit Sub
    Dim requestId As Long, slot As Byte, expected As t_UserOBJ, quantity As Long
    Dim reply As t_InventoryDestruction
    requestId = message.ReadInt32()
    slot = message.ReadInt8()
    expected.ObjIndex = message.ReadInt16()
    expected.amount = message.ReadInt16()
    expected.Equipped = message.ReadInt8()
    expected.ElementalTags = message.ReadInt32()
    quantity = message.ReadInt32()
    If requestId <= 0 Then Exit Sub
    Call PrepareInventoryDestruction(UserIndex, requestId, slot, expected, quantity, GetTickCountRaw(), reply)
    Call WriteInventoryDestructionResult(UserIndex, 0, reply)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "InventoryDestruction.HandlePrepare", Erl)
End Sub

Public Sub HandleInventoryDestructionConfirm(ByVal UserIndex As Integer, ByVal message As Network.reader)
    On Error GoTo Failed
    If message.GetAvailable() <> 16 Then Exit Sub
    Dim token(0 To 15) As Byte, i As Integer, reply As t_InventoryDestruction
    For i = 0 To 15
        token(i) = message.ReadInt8()
    Next i
    Call ConfirmInventoryDestruction(UserIndex, token, GetTickCountRaw(), reply)
    Call WriteInventoryDestructionResult(UserIndex, 1, reply)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "InventoryDestruction.HandleConfirm", Erl)
End Sub

Public Sub WriteInventoryDestructionResult(ByVal UserIndex As Integer, ByVal Stage As Byte, ByRef reply As t_InventoryDestruction)
    On Error GoTo Failed
    Dim writer As Network.writer, i As Integer, expiry As Integer
    Set writer = Protocol_Writes.GetWriterBuffer()
    If Stage = 0 And reply.Result = DestroyOkay Then expiry = 30
    Call writer.WriteInt16(ServerPacketID.eHooInventoryDestructionResult)
    Call writer.WriteInt32(reply.RequestId)
    For i = 0 To 15
        Call writer.WriteInt8(reply.Token(i))
    Next i
    Call writer.WriteInt8(Stage)
    Call writer.WriteInt8(CByte(reply.Result))
    Call writer.WriteInt8(reply.Slot)
    Call writer.WriteInt16(reply.Item.ObjIndex)
    Call writer.WriteInt16(reply.Item.amount)
    Call writer.WriteInt8(reply.Item.Equipped)
    Call writer.WriteInt32(reply.EffectiveElementalTags)
    Call writer.WriteInt32(reply.Quantity)
    Call writer.WriteInt16(expiry)
    Call SendData(SendTarget.ToIndex, UserIndex)
    Exit Sub
Failed:
    If Not writer Is Nothing Then Call writer.Clear()
    Call TraceError(Err.Number, Err.Description, "InventoryDestruction.WriteResult", Erl)
End Sub
