Attribute VB_Name = "ModCollectibleCards"
Option Explicit

Private Const UPSERT_NEW_COLLECTIBLE_CARD As String = _
    "INSERT INTO account_collectible_cards (account_id, card_id, last_updated, quantity) " & _
    "VALUES (?, ?, ?, 1) " & _
    "ON CONFLICT(account_id, card_id) " & _
    "DO UPDATE SET quantity = quantity + 1, last_updated = excluded.last_updated;"

Public Sub AddCollectibleCardToUser(ByVal UserIndex As Integer, ByRef ObjCard As t_Obj)
    If ObjCard.ObjIndex = 0 Then Exit Sub
    With UserList(UserIndex)
        Call InsertCollectibleCardIntoDatabase(.AccountID, ObjCard)
    End With
End Sub

Public Sub InsertCollectibleCardIntoDatabase(ByVal Acount_Id As Integer, ByRef ObjCard As t_Obj)
    On Error GoTo InsertCollectibleCardIntoDatabase_Err
    Dim RS As ADODB.Recordset
    Set RS = Query(UPSERT_NEW_COLLECTIBLE_CARD, Acount_Id, ObjData(ObjCard.ObjIndex).CollectibleCardId, CStr(DateTime.Now))
    Exit Sub
InsertCollectibleCardIntoDatabase_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.InsertCollectibleCardIntoDatabase", Erl)
End Sub
