Attribute VB_Name = "ModCollectibleCards"
Option Explicit

Private Const INSER_NEW_COLLECTIBLE_CARD As String = "INSERT INTO account_collectible_cards (account_id,card_id,rarity,timestamp) VALUES (?,?,?,?);"

Public Sub AddCollectibleCardToUser(ByVal UserIndex As Integer, ByRef ObjCard As t_Obj)
    If ObjCard.ObjIndex = 0 Then Exit Sub
    With UserList(UserIndex)
    
        '.CollectibleCards should be array or map? maybe map would be better
        Dim i As Integer
        Dim Hit As Boolean
        Dim Index As Integer
        For i = 1 To UBound(.CollectibleCards)
            If .CollectibleCards(i).Id = ObjCard.ObjIndex Then
                If .CollectibleCards(i).Rarity = ObjData(ObjCard.ObjIndex).Rarity Then
                    Hit = True
                    Index = i
                    Exit For
                End If
            End If
        Next i
        
        If Hit Then
            .CollectibleCards(Index).Amount = .CollectibleCards(Index).Amount + 1
        Else
            ReDim Preserve .CollectibleCards(1 To (UBound(.CollectibleCards) + 1))
            .CollectibleCards(UBound(.CollectibleCards)).Amount = 1
            .CollectibleCards(UBound(.CollectibleCards)).Id = ObjCard.ObjIndex
            .CollectibleCards(UBound(.CollectibleCards)).Rarity = ObjData(ObjCard.ObjIndex).Rarity
        End If
        '////////////////////////////////////////////////////////////////////////////////////////////////////
        
        Call InsertCollectibleCardIntoDatabase(.AccountID, ObjCard)
    End With
End Sub

Public Sub InsertCollectibleCardIntoDatabase(ByVal Acount_Id As Integer, ByRef ObjCard As t_Obj)
    On Error GoTo InsertCollectibleCardIntoDatabase_Err
    Dim RS As ADODB.Recordset
    Set RS = Query(INSER_NEW_COLLECTIBLE_CARD, Acount_Id, ObjCard.ObjIndex, ObjData(ObjCard.ObjIndex).Rarity, CStr(DateTime.Now))
    Exit Sub
InsertCollectibleCardIntoDatabase_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.InsertCollectibleCardIntoDatabase", Erl)
End Sub
