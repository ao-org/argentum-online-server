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
#End If
