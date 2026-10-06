Attribute VB_Name = "Unit_PlayerTrade"
Option Explicit
#If UNIT_TEST = 1 Then
Public Function test_suite_player_trade() As Boolean
    Call UnitTesting.RunTest("trade_multiple_items_and_gold", TestMultipleOffers())
    Call UnitTesting.RunTest("trade_split_preserves_tags", TestSplitTags())
    Call UnitTesting.RunTest("trade_full_board_is_atomic", TestFullBoard())
    Call UnitTesting.RunTest("trade_gold_checks_total_without_overflow", TestGoldLimit())
    Call UnitTesting.RunTest("trade_counts_duplicate_stacks", TestAggregateAmount())
    Call UnitTesting.RunTest("trade_rejection_clears_pending_offer", TestClearRequest())
    Call UnitTesting.RunTest("trade_invitation_requires_current_recipient", TestInvitationMatch())
    Call UnitTesting.RunTest("inventory_requires_long_quantity_across_stacks", TestLongInventoryAmounts())
    Call UnitTesting.RunTest("inventory_long_quantity_preserves_elemental_tags", TestLongInventoryTags())
    Call UnitTesting.RunTest("inventory_gold_requirement_accepts_long_range", TestLongInventoryGold())
    Call UnitTesting.RunTest("inventory_potion_limit_addition_uses_long", TestLongPotionLimit())
    test_suite_player_trade = True
End Function

Private Function TestMultipleOffers() As Boolean
    On Error GoTo TestMultipleOffers_Err
    Dim items(1 To 6) As t_Obj
    Dim item As t_Obj
    Dim gold As Long
    item.ObjIndex = 1: item.amount = 10000
    If Not AddSafeTradeOffer(items, gold, item, 50000, 10000) Then Exit Function
    item.ObjIndex = 2: item.amount = 1
    If Not AddSafeTradeOffer(items, gold, item, 50000, 10000) Then Exit Function
    item.ObjIndex = 3
    If Not AddSafeTradeOffer(items, gold, item, 50000, 10000) Then Exit Function
    item.ObjIndex = 4
    If Not AddSafeTradeOffer(items, gold, item, 50000, 10000) Then Exit Function
    item.ObjIndex = 0: item.amount = 20000
    If Not AddSafeTradeOffer(items, gold, item, 50000, 10000) Then Exit Function
    TestMultipleOffers = items(1).amount = 10000 And items(2).ObjIndex = 2 And items(3).ObjIndex = 3 And items(4).ObjIndex = 4 And gold = 20000
    Exit Function
TestMultipleOffers_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestMultipleOffers", Erl)
End Function

Private Function TestSplitTags() As Boolean
    On Error GoTo TestSplitTags_Err
    Dim items(1 To 6) As t_Obj
    Dim item As t_Obj
    Dim gold As Long
    item.ObjIndex = 1: item.amount = 9999: item.ElementalTags = 5
    If Not AddSafeTradeOffer(items, gold, item, 0, 10000) Then Exit Function
    item.amount = 2
    If Not AddSafeTradeOffer(items, gold, item, 0, 10000) Then Exit Function
    TestSplitTags = items(1).amount = 10000 And items(2).amount = 1 And items(2).ElementalTags = 5
    Exit Function
TestSplitTags_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestSplitTags", Erl)
End Function

Private Function TestFullBoard() As Boolean
    On Error GoTo TestFullBoard_Err
    Dim items(1 To 6) As t_Obj
    Dim item As t_Obj
    Dim gold As Long
    Dim i As Long
    For i = 1 To 6
        items(i).ObjIndex = i: items(i).amount = 10000
    Next i
    items(1).amount = 9999
    item.ObjIndex = 1: item.amount = 2
    If AddSafeTradeOffer(items, gold, item, 0, 10000) Then Exit Function
    If items(1).amount <> 9999 Then Exit Function
    item.amount = 1
    If Not AddSafeTradeOffer(items, gold, item, 0, 10000) Then Exit Function
    TestFullBoard = items(1).amount = 10000
    Exit Function
TestFullBoard_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestFullBoard", Erl)
End Function

Private Function TestGoldLimit() As Boolean
    On Error GoTo TestGoldLimit_Err
    Dim items(1 To 6) As t_Obj
    Dim item As t_Obj
    Dim gold As Long
    item.amount = 2147483647
    If Not AddSafeTradeOffer(items, gold, item, 2147483647, 10000) Then Exit Function
    item.amount = 1
    If AddSafeTradeOffer(items, gold, item, 2147483647, 10000) Then Exit Function
    TestGoldLimit = gold = 2147483647
    Exit Function
TestGoldLimit_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestGoldLimit", Erl)
End Function

Private Function TestAggregateAmount() As Boolean
    On Error GoTo TestAggregateAmount_Err
    Dim items(1 To 6) As t_Obj
    items(1).ObjIndex = 1: items(1).amount = 10000: items(1).ElementalTags = 1
    items(2).ObjIndex = 1: items(2).amount = 5000: items(2).ElementalTags = 1
    items(3).ObjIndex = 1: items(3).amount = 2000: items(3).ElementalTags = 2
    TestAggregateAmount = SafeTradeOfferedAmount(items, 1, 1) = 15000 And SafeTradeOfferedAmount(items, 1, 2) = 2000
    Exit Function
TestAggregateAmount_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestAggregateAmount", Erl)
End Function
Private Function TestClearRequest() As Boolean
    On Error GoTo TestClearRequest_Err
    Dim request As t_ComercioUsuario
    request.DestUsu.ArrayIndex = 2: request.DestUsu.VersionId = 20
    request.InvitationFrom.ArrayIndex = 1: request.InvitationFrom.VersionId = 10
    request.DestNick = "Other": request.Acepto = True: request.Oro = 1234
    request.itemsAenviar(6).ObjIndex = 1: request.itemsAenviar(6).amount = 10
    Call ClearSafeTradeRequest(request)
    TestClearRequest = request.DestUsu.ArrayIndex = 0 And request.InvitationFrom.ArrayIndex = 0 And Len(request.DestNick) = 0 And Not request.Acepto And request.Oro = 0 And request.itemsAenviar(6).ObjIndex = 0
    Exit Function
TestClearRequest_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestClearRequest", Erl)
End Function

Private Function TestInvitationMatch() As Boolean
    On Error GoTo TestInvitationMatch_Err
    Dim request As t_ComercioUsuario
    request.DestUsu.ArrayIndex = 2: request.DestUsu.VersionId = 20
    If Not SafeTradeInvitationMatches(request, 2, 20) Then Exit Function
    If SafeTradeInvitationMatches(request, 3, 20) Then Exit Function
    If SafeTradeInvitationMatches(request, 2, 21) Then Exit Function
    Call ClearSafeTradeRequest(request)
    TestInvitationMatch = Not SafeTradeInvitationMatches(request, 2, 20)
    Exit Function
TestInvitationMatch_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestInvitationMatch", Erl)
End Function

' Save and restore the fixture so these checks never alter an existing user.
Private Function CheckLongInventoryRequirement(ByVal objectIndex As Long, ByVal requiredAmount As Long, ByVal tags As Long, ByVal availableGold As Long) As Boolean
    On Error GoTo CheckLongInventoryRequirement_Err
    Dim savedInventory As t_Inventario
    Dim savedGold      As Long
    Dim savedSlots     As Byte
    Dim backupReady    As Boolean
    Dim emptyInventory As t_Inventario
    Dim i              As Long
    savedInventory = UserList(1).invent
    savedGold = UserList(1).Stats.GLD
    savedSlots = UserList(1).CurrentInventorySlots
    backupReady = True
    UserList(1).invent = emptyInventory
    UserList(1).Stats.GLD = availableGold
    UserList(1).CurrentInventorySlots = 8
    For i = 1 To 6
        UserList(1).invent.Object(i).ObjIndex = GOLD_OBJ_INDEX + 1
        UserList(1).invent.Object(i).amount = 10000
        UserList(1).invent.Object(i).ElementalTags = 1
    Next i
    ' Different tags and a different object must not count toward the total.
    UserList(1).invent.Object(7).ObjIndex = GOLD_OBJ_INDEX + 1
    UserList(1).invent.Object(7).amount = 10000
    UserList(1).invent.Object(7).ElementalTags = 2
    UserList(1).invent.Object(8).ObjIndex = GOLD_OBJ_INDEX + 2
    UserList(1).invent.Object(8).amount = 10000
    UserList(1).invent.Object(8).ElementalTags = 1
    CheckLongInventoryRequirement = TieneObjetos(objectIndex, requiredAmount, 1, tags)
    UserList(1).invent = savedInventory
    UserList(1).Stats.GLD = savedGold
    UserList(1).CurrentInventorySlots = savedSlots
    Exit Function
CheckLongInventoryRequirement_Err:
    If backupReady Then
        UserList(1).invent = savedInventory
        UserList(1).Stats.GLD = savedGold
        UserList(1).CurrentInventorySlots = savedSlots
    End If
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.CheckLongInventoryRequirement", Erl)
End Function

Private Function TestLongInventoryAmounts() As Boolean
    On Error GoTo TestLongInventoryAmounts_Err
    TestLongInventoryAmounts = CheckLongInventoryRequirement(GOLD_OBJ_INDEX + 1, 32767, 1, 0) And _
        CheckLongInventoryRequirement(GOLD_OBJ_INDEX + 1, 32768, 1, 0) And _
        CheckLongInventoryRequirement(GOLD_OBJ_INDEX + 1, 60000, 1, 0) And _
        Not CheckLongInventoryRequirement(GOLD_OBJ_INDEX + 1, 60001, 1, 0)
    Exit Function
TestLongInventoryAmounts_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestLongInventoryAmounts", Erl)
End Function

Private Function TestLongInventoryTags() As Boolean
    On Error GoTo TestLongInventoryTags_Err
    TestLongInventoryTags = CheckLongInventoryRequirement(GOLD_OBJ_INDEX + 1, 10000, 2, 0) And _
        Not CheckLongInventoryRequirement(GOLD_OBJ_INDEX + 1, 32768, 2, 0) And _
        Not CheckLongInventoryRequirement(GOLD_OBJ_INDEX + 1, 1, 0, 0)
    Exit Function
TestLongInventoryTags_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestLongInventoryTags", Erl)
End Function

Private Function TestLongInventoryGold() As Boolean
    On Error GoTo TestLongInventoryGold_Err
    TestLongInventoryGold = CheckLongInventoryRequirement(GOLD_OBJ_INDEX, 50000, 0, 50000) And _
        Not CheckLongInventoryRequirement(GOLD_OBJ_INDEX, 50001, 0, 50000) And _
        CheckLongInventoryRequirement(GOLD_OBJ_INDEX, 2147483647, 0, 2147483647)
    Exit Function
TestLongInventoryGold_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestLongInventoryGold", Erl)
End Function

Private Function TestLongPotionLimit() As Boolean
    On Error GoTo TestLongPotionLimit_Err
    Dim potionLimit As Integer
    potionLimit = 32767
    TestLongPotionLimit = CheckLongInventoryRequirement(GOLD_OBJ_INDEX + 1, CLng(potionLimit) + 1, 1, 0)
    Exit Function
TestLongPotionLimit_Err:
    Call TraceError(Err.Number, Err.Description, "Unit_PlayerTrade.TestLongPotionLimit", Erl)
End Function
#End If
