Attribute VB_Name = "ModCollectibleCards"
Option Explicit

'Cards And Npcs should be saved as the following
'If i have a scorpion npc and his card is the objIndex 32
'both the scorpion npc and the objindex32 should have matching properties
'CollectibleCardIndex (1-1024)
'CollectibleCardIndex is the direct index into the AccountCollectibleCardQuantities array
'If the card visually says that the scorpion is #234, then the quantity is stored at
'AccountCollectibleCardQuantities(234)
'Then the game compares the npc value with the value that the player has in his account
'whenever the business logic sees fit (eg: combat, finance, etc)

Private Const MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE = 1024

Public Enum e_CardRarity
    Bronce = 1
    Silver = 2
    gold = 3
End Enum

Private Const UPSERT_NEW_COLLECTIBLE_CARDS As String = _
    "INSERT INTO account_collectible_cards (account_id, card_quantities) " & _
    "VALUES (?, ?) " & _
    "ON CONFLICT(account_id) " & _
    "DO UPDATE SET card_quantities = excluded.card_quantities;"
    
Private Const GET_ACCOUNT_COLLECTIBLE_CARDS As String = "SELECT card_quantities FROM account_collectible_cards WHERE account_id = ?"

Public Function SetupUserAccountAccountCollectibleCardBitArray(ByRef User As t_User)
    On Error GoTo GetUserCollectibleCards_Err
    Dim RS As ADODB.Recordset
    Dim Cmd As ADODB.Command
    Dim i As Integer
    Dim QuantityData() As Byte
    
    ' Create command to fetch the collectible card blob
    Set Cmd = New ADODB.Command
    With Cmd
        .ActiveConnection = Connection ' Your SQLite connection object
        .CommandText = GET_ACCOUNT_COLLECTIBLE_CARDS
        .CommandType = adCmdText
        
        ' Add parameter for account_id
        .Parameters.Append .CreateParameter("@AccountID", adInteger, adParamInput, , User.AccountID)
        
        ' Execute and get recordset
        Set RS = .Execute
    End With
    
    If Not RS Is Nothing Then
        If Not RS.EOF And Not RS.BOF Then
            ' Load quantity array (1024 bytes)
            If Not IsNull(RS!card_quantities) Then
                ' Get the quantity blob data
                QuantityData = RS!card_quantities
                
                ' Copy quantity data into the user's byte array
                For i = 1 To MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE
                    If i <= UBound(QuantityData) + 1 Then
                        User.AccountCollectibleCardQuantities(i) = QuantityData(i - 1) ' Arrays are 0-based from DB
                    Else
                        User.AccountCollectibleCardQuantities(i) = 0
                    End If
                Next i
            End If
        End If
        RS.Close
    End If
    
    Set RS = Nothing
    Set Cmd = Nothing
    Exit Function

GetUserCollectibleCards_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.SetupUserAccountAccountCollectibleCardBitArray", Erl)
    Set RS = Nothing
    Set Cmd = Nothing
End Function

Public Function SaveUserAccountCollectibleCards(ByVal UserIndex As Integer, ByVal QueryBreakDown As String)
    On Error GoTo SaveUserAccountCollectibleCards_Err
    Dim Cmd As ADODB.Command
    Dim QuantityData(0 To 1023) As Byte
    Dim i As Integer
    
    ' Copy user's quantity array to 0-based array for database (1024 bytes)
    For i = 1 To MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE
        QuantityData(i - 1) = UserList(UserIndex).AccountCollectibleCardQuantities(i)
    Next i
    
    ' Create command to upsert the blob
    Set Cmd = New ADODB.Command
    With Cmd
        .ActiveConnection = Connection ' Your SQLite connection object
        .CommandText = UPSERT_NEW_COLLECTIBLE_CARDS
        .CommandType = adCmdText
        
        ' Add account_id parameter
        .Parameters.Append .CreateParameter("@AccountID", adInteger, adParamInput, , UserList(UserIndex).AccountID)
        
        ' Add quantity blob parameter (1024 bytes)
        .Parameters.Append .CreateParameter("@QuantityData", adVarBinary, adParamInput, MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE, QuantityData)
        
        ' Execute the upsert
        .Execute
    End With
    
    Set Cmd = Nothing
    Exit Function

SaveUserAccountCollectibleCards_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.SaveUserAccountCollectibleCards", Erl)
    Set Cmd = Nothing
End Function

Public Sub AddCollectibleCardToUser(ByVal UserIndex As Integer, ByRef ObjCard As t_Obj)
    If ObjCard.ObjIndex = 0 Then Exit Sub
    
    Dim CardIndex As Integer
    CardIndex = ObjData(ObjCard.ObjIndex).CollectibleCardIndex
    
    ' Validate card index
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then Exit Sub
    
    With UserList(UserIndex)
        .flags.DirtyCollectibleCardCollection = True
        ' Increment the quantity for this specific card (cap at 255 per byte)
        Dim Inc As Integer
        If ObjData(ObjCard.ObjIndex).CollectibleCardRarity = e_CardRarity.Bronce Then
            Inc = CInt(SvrConfig.GetValue("CardBronceQuantity"))
        ElseIf ObjData(ObjCard.ObjIndex).CollectibleCardRarity = e_CardRarity.Silver Then
            Inc = CInt(SvrConfig.GetValue("CardSilverQuantity"))
        ElseIf ObjData(ObjCard.ObjIndex).CollectibleCardRarity = e_CardRarity.gold Then
            Inc = CInt(SvrConfig.GetValue("CardGoldQuantity"))
        Else
            Inc = 0
        End If

        If Inc > 0 Then
            Dim NewQty As Integer
            NewQty = .AccountCollectibleCardQuantities(CardIndex) + Inc
            If NewQty > 255 Then NewQty = 255
            .AccountCollectibleCardQuantities(CardIndex) = CByte(NewQty)
        End If
    End With
End Sub

Public Function HasUserCollectedNpcCard(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    Dim CardIndex As Integer
    CardIndex = NpcList(NpcIndex).CollectibleCardIndex
    
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        HasUserCollectedNpcCard = False
        Exit Function
    End If
    
    ' Check if quantity is greater than 0
    HasUserCollectedNpcCard = (UserList(UserIndex).AccountCollectibleCardQuantities(CardIndex) > 0)
End Function

Public Function HasUserCollectedCard(ByVal UserIndex As Integer, ByVal CardIndex As Integer) As Boolean
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        HasUserCollectedCard = False
        Exit Function
    End If
    
    HasUserCollectedCard = (UserList(UserIndex).AccountCollectibleCardQuantities(CardIndex) > 0)
End Function

Public Function GetUserCardQuantity(ByVal UserIndex As Integer, ByVal CardIndex As Integer) As Byte
    ' Returns the quantity of a specific card (0-255)
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetUserCardQuantity = 0
    Else
        GetUserCardQuantity = UserList(UserIndex).AccountCollectibleCardQuantities(CardIndex)
    End If
End Function

Public Function GetCardExpBonusForNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Single
    ' Returns the EXP multiplier based on the card level for THIS specific NPC
    ' Returns 1.0 if no bonus applies
    Dim CardIndex As Integer
    Dim CardQuantity As Byte
    
    ' Get the card index for this NPC
    CardIndex = NpcList(NpcIndex).CollectibleCardIndex
    
    ' Validate card index
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardExpBonusForNpc = 1#
        Exit Function
    End If
    
    ' Get the quantity of this specific card the user has
    CardQuantity = GetUserCardQuantity(UserIndex, CardIndex)
    
    ' Apply EXP bonus based on card level (quantity)
    ' Accumulative bonuses as the player levels up their card
    GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonusBase"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardExpBonusLvl1")) Then GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonus1"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardExpBonusLvl2")) Then GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonus2"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardExpBonusLvl3")) Then GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonus3"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardExpBonusLvl4")) Then GetCardExpBonusForNpc = CSng(SvrConfig.GetValue("CardExpBonus4"))
End Function

Public Function GetCardDamageBonusForNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Single
    ' Returns the damage multiplier based on the card level for THIS specific NPC
    ' Returns 1.0 if no bonus applies
    Dim CardIndex As Integer
    Dim CardQuantity As Byte
    
    ' Get the card index for this NPC
    CardIndex = NpcList(NpcIndex).CollectibleCardIndex
    
    ' Validate card index
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardDamageBonusForNpc = 1#
        Exit Function
    End If
    
    ' Get the quantity of this specific card the user has
    CardQuantity = GetUserCardQuantity(UserIndex, CardIndex)
    
    ' Apply damage bonus based on card level (quantity)
    ' Accumulative bonuses as the player levels up their card
    GetCardDamageBonusForNpc = CSng(SvrConfig.GetValue("CardDamageBonusBase"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageBonusLvl1")) Then GetCardDamageBonusForNpc = CSng(SvrConfig.GetValue("CardDamageBonus1"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageBonusLvl2")) Then GetCardDamageBonusForNpc = CSng(SvrConfig.GetValue("CardDamageBonus2"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageBonusLvl3")) Then GetCardDamageBonusForNpc = CSng(SvrConfig.GetValue("CardDamageBonus3"))
End Function

Public Function GetCardDamageReductionForNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Single
    ' Returns the damage reduction multiplier based on the card level for THIS specific NPC
    ' Returns 1.0 if no reduction applies (full damage)
    ' Returns values < 1.0 for damage reduction (e.g., 0.8 = 80% damage = 20% reduction)
    Dim CardIndex As Integer
    Dim CardQuantity As Byte
    
    ' Get the card index for this NPC
    CardIndex = NpcList(NpcIndex).CollectibleCardIndex
    
    ' Validate card index
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardDamageReductionForNpc = 1#
        Exit Function
    End If
    
    ' Get the quantity of this specific card the user has
    CardQuantity = GetUserCardQuantity(UserIndex, CardIndex)
    
    ' Apply damage reduction based on card level (quantity)
    ' Lower values = more reduction (we multiply the incoming damage by this)
    GetCardDamageReductionForNpc = CSng(SvrConfig.GetValue("CardDamageReductionBase"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageReductionLvl1")) Then GetCardDamageReductionForNpc = CSng(SvrConfig.GetValue("CardDamageReduction1"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageReductionLvl2")) Then GetCardDamageReductionForNpc = CSng(SvrConfig.GetValue("CardDamageReduction2"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDamageReductionLvl3")) Then GetCardDamageReductionForNpc = CSng(SvrConfig.GetValue("CardDamageReduction3"))
End Function


Public Function GetCardDropBonusForNpc(ByVal UserIndex As Integer, ByRef Npc As t_Npc) As Single
    Dim CardIndex As Integer
    Dim CardQuantity As Byte
    
    CardIndex = Npc.CollectibleCardIndex
    
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardDropBonusForNpc = 1#
        Exit Function
    End If
    
    CardQuantity = GetUserCardQuantity(UserIndex, CardIndex)
    
    GetCardDropBonusForNpc = CSng(SvrConfig.GetValue("CardDropBonusBase"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDropBonusLvl1")) Then GetCardDropBonusForNpc = CSng(SvrConfig.GetValue("CardDropBonus1"))
    If CLng(CardQuantity) >= CLng(SvrConfig.GetValue("CardDropBonusLvl2")) Then GetCardDropBonusForNpc = CSng(SvrConfig.GetValue("CardDropBonus2"))
End Function
