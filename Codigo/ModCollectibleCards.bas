Attribute VB_Name = "ModCollectibleCards"
Option Explicit

'Cards And Npcs should be saved as the following
'If i have a scorpion npc and his card is the objIndex 32
'both the scorpion npc and the objindex32 should have matching properties
'CollectibleCardSlot
'CollectibleCardValue
'CollectibleCardSlot indicates in what fraction of a byte it's located, if the card
'visually says that the scorpion is the 234 then 234/8 -> 29
'that means that in memory the value will be stored in AccountCollectibleCardBitArray(29)
'and the CollectibleCardValue works as a MASK of bits
'Then the game compares the npc value with the value that the player has in his account whenever the bussiness logic sees fit (eg: combat, finance, etc)

Private Const MAX_COLLECTIBLE_CARDS_ARR = 128

Private Const UPSERT_NEW_COLLECTIBLE_CARDS As String = _
    "INSERT INTO account_collectible_cards (account_id, card_bit_array) " & _
    "VALUES (?, ?) " & _
    "ON CONFLICT(account_id) " & _
    "DO UPDATE SET card_bit_array = excluded.card_bit_array;"
    
Private Const GET_ACCOUNT_COLLECTIBLE_CARDS As String = "SELECT card_bit_array FROM account_collectible_cards WHERE account_id = ?"

Public Function SetupUserAccountAccountCollectibleCardBitArray(ByRef User As t_User)
    On Error GoTo GetUserCollectibleCards_Err
    Dim RS As ADODB.Recordset
    Dim HexString As String
    Dim i As Integer
    Dim ByteValue As String
    Dim TempValue As Long
    
    Set RS = Query(GET_ACCOUNT_COLLECTIBLE_CARDS, User.AccountID)
    
    If Not RS.EOF Then
        If Not IsNull(RS("card_bit_array").value) Then
            HexString = Trim$(RS("card_bit_array").value)
            
            ' Convert hex string back to byte array (256 chars = 128 bytes)
            For i = 1 To MAX_COLLECTIBLE_CARDS_ARR
                ByteValue = mid$(HexString, (i - 1) * 2 + 1, 2)
                
                If Len(ByteValue) = 2 Then
                    ' Use Val() to convert hex string other casts resulted in TypeMismatch Cbyte CInt Clng
                    TempValue = val("&H" & ByteValue)
                    User.AccountCollectibleCardBitArray(i) = TempValue
                End If
            Next i
            
            Call LogError("Successfully loaded " & MAX_COLLECTIBLE_CARDS_ARR & " bytes")
        End If
    End If
    Exit Function

GetUserCollectibleCards_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.SetupUserAccountAccountCollectibleCardBitArray", Erl)
End Function


Public Function SaveUserAccountCollectibleCards(ByVal UserIndex As Integer, ByVal QueryBreakDown As String)
    Dim RS As ADODB.Recordset
    Dim HexString As String
    Dim i As Integer
    With UserList(UserIndex)
        If .flags.DirtyCollectibleCardBitArray Then
            
            HexString = ""  ' Explicitly initialize
            
            For i = 1 To MAX_COLLECTIBLE_CARDS_ARR
                ' Convierte cada byte a 2 caracteres hexadecimales (ej. 3 -> "03")
                HexString = HexString & Right$("0" & Hex(.AccountCollectibleCardBitArray(i)), 2)
            Next i
            
            Set RS = Query(UPSERT_NEW_COLLECTIBLE_CARDS, .AccountID, HexString)

            .flags.DirtyCollectibleCardBitArray = False
        End If
    End With
End Function


Public Sub AddCollectibleCardToUser(ByVal UserIndex As Integer, ByRef ObjCard As t_Obj)
    If ObjCard.ObjIndex = 0 Then Exit Sub
    With UserList(UserIndex)
        .flags.DirtyCollectibleCardBitArray = True
        .AccountCollectibleCardBitArray(ObjData(ObjCard.ObjIndex).CollectibleCardSlot) = .AccountCollectibleCardBitArray(ObjData(ObjCard.ObjIndex).CollectibleCardSlot) Or ObjData(ObjCard.ObjIndex).CollectibleCardValue
    End With
End Sub

Public Function HasUserCollectedNpcCard(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .AccountCollectibleCardBitArray(NpcList(NpcIndex).CollectibleCardSlot) And NpcList(NpcIndex).CollectibleCardValue = NpcList(NpcIndex).CollectibleCardValue Then
            HasUserCollectedNpcCard = True
        End If
    End With
End Function

