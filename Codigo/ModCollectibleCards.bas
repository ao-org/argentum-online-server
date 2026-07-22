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

Private Const MAX_COLLECTIBLE_CARDS_ARR_SIZE = 128
Private Const MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE = 1024

Private Const UPSERT_NEW_COLLECTIBLE_CARDS As String = _
    "INSERT INTO account_collectible_cards (account_id, card_bit_array, card_quantities) " & _
    "VALUES (?, ?, ?) " & _
    "ON CONFLICT(account_id) " & _
    "DO UPDATE SET card_bit_array = excluded.card_bit_array, card_quantities = excluded.card_quantities;"
    
Private Const GET_ACCOUNT_COLLECTIBLE_CARDS As String = "SELECT card_bit_array, card_quantities FROM account_collectible_cards WHERE account_id = ?"

Public Function SetupUserAccountAccountCollectibleCardBitArray(ByRef User As t_User)
    On Error GoTo GetUserCollectibleCards_Err
    Dim RS As ADODB.Recordset
    Dim Cmd As ADODB.Command
    Dim i As Integer
    Dim BlobData() As Byte
    Dim QuantityData() As Byte
    
    ' Create command to fetch the collectible card blob
    Set Cmd = New ADODB.Command
    With Cmd
        .ActiveConnection = Connection ' Your SQLite connection object
        .CommandText = GET_ACCOUNT_COLLECTIBLE_CARDS
        .CommandType = adCmdText
        
        ' Add parameter for account_id
        .Parameters.Append .CreateParameter("@AccountID", adInteger, adParamInput, , User.AccountID)
        
        ' Execute and get recordset -> .Execute is asynchronous
        Set RS = .Execute
    End With
    
    If Not RS Is Nothing Then
        If Not RS.EOF And Not RS.BOF Then
            ' Load bit array (128 bytes)
            If Not IsNull(RS!card_bit_array) Then
                ' Get the blob data
                BlobData = RS!card_bit_array
                
                ' Copy blob data into the user's byte array
                For i = 1 To MAX_COLLECTIBLE_CARDS_ARR_SIZE
                    If i <= UBound(BlobData) + 1 Then
                        User.AccountCollectibleCardBitArray(i) = BlobData(i - 1) ' Arrays are 0-based from DB
                    Else
                        User.AccountCollectibleCardBitArray(i) = 0
                    End If
                Next i
            End If
            
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
    Dim BlobData(0 To 127) As Byte
    Dim QuantityData(0 To 1023) As Byte
    Dim i As Integer
    
    ' Copy user's bit array to 0-based array for database (128 bytes)
    For i = 1 To MAX_COLLECTIBLE_CARDS_ARR_SIZE
        BlobData(i - 1) = UserList(UserIndex).AccountCollectibleCardBitArray(i)
    Next i
    
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
        
        ' Add blob parameter (128 bytes)
        .Parameters.Append .CreateParameter("@BlobData", adVarBinary, adParamInput, MAX_COLLECTIBLE_CARDS_ARR_SIZE, BlobData)
        
        ' Add quantity blob parameter (1024 bytes)
        .Parameters.Append .CreateParameter("@QuantityData", adVarBinary, adParamInput, MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE, QuantityData)
        
        ' Execute the upsert -> .Execute is asyncronous
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
    
    ' Calculate the actual card index (1-1024) from slot and bit position
    CardIndex = GetCardIndexFromSlotAndValue(ObjData(ObjCard.ObjIndex).CollectibleCardSlot, ObjData(ObjCard.ObjIndex).CollectibleCardValue)
    
    With UserList(UserIndex)
        .flags.DirtyCollectibleCardBitArray = True
        
        ' Set the bit flag
        .AccountCollectibleCardBitArray(ObjData(ObjCard.ObjIndex).CollectibleCardSlot) = .AccountCollectibleCardBitArray(ObjData(ObjCard.ObjIndex).CollectibleCardSlot) Or ObjData(ObjCard.ObjIndex).CollectibleCardValue
        
        ' Increment the quantity for this specific card (max 255 per byte)
        If CardIndex > 0 And CardIndex <= MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
            If .AccountCollectibleCardQuantities(CardIndex) < 255 Then
                .AccountCollectibleCardQuantities(CardIndex) = .AccountCollectibleCardQuantities(CardIndex) + 1
            End If
        End If
    End With
End Sub

Public Function HasUserCollectedNpcCard(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .AccountCollectibleCardBitArray(NpcList(NpcIndex).CollectibleCardSlot) And NpcList(NpcIndex).CollectibleCardValue = NpcList(NpcIndex).CollectibleCardValue Then
            HasUserCollectedNpcCard = True
        End If
    End With
End Function

Public Function GetUserCardQuantity(ByVal UserIndex As Integer, ByVal CardSlot As Integer, ByVal CardValue As Byte) As Byte
    ' Returns the quantity of a specific card (0-255)
    Dim CardIndex As Integer
    
    CardIndex = GetCardIndexFromSlotAndValue(CardSlot, CardValue)
    
    If CardIndex > 0 And CardIndex <= MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetUserCardQuantity = UserList(UserIndex).AccountCollectibleCardQuantities(CardIndex)
    Else
        GetUserCardQuantity = 0
    End If
End Function

' Converts slot (1-128) and bit mask value to card index (1-1024)
' CardValue is a bit mask (1, 2, 4, 8, 16, 32, 64, 128) representing which bit in the byte
Private Function GetCardIndexFromSlotAndValue(ByVal CardSlot As Integer, ByVal CardValue As Byte) As Integer
    Dim BitPosition As Integer
    
    ' Validate slot range
    If CardSlot < 1 Or CardSlot > MAX_COLLECTIBLE_CARDS_ARR_SIZE Then
        GetCardIndexFromSlotAndValue = 0
        Exit Function
    End If
    
    ' Determine bit position (0-7) from the mask value using bit shifting logic
    ' This is more efficient than Select Case for power-of-2 values
    Select Case CardValue
        Case 1: BitPosition = 0      ' 2^0
        Case 2: BitPosition = 1      ' 2^1
        Case 4: BitPosition = 2      ' 2^2
        Case 8: BitPosition = 3      ' 2^3
        Case 16: BitPosition = 4     ' 2^4
        Case 32: BitPosition = 5     ' 2^5
        Case 64: BitPosition = 6     ' 2^6
        Case 128: BitPosition = 7    ' 2^7
        Case Else
            ' Invalid bit mask value
            GetCardIndexFromSlotAndValue = 0
            Exit Function
    End Select
    
    ' Calculate card index: (Slot - 1) * 8 + BitPosition + 1
    ' Example: Slot 1, BitPosition 0 -> Card 1
    '          Slot 1, BitPosition 7 -> Card 8
    '          Slot 2, BitPosition 0 -> Card 9
    '          Slot 128, BitPosition 7 -> Card 1024
    GetCardIndexFromSlotAndValue = ((CardSlot - 1) * 8) + BitPosition + 1
End Function

' Converts card index (1-1024) back to slot and bit mask value
' Useful for reverse lookups and debugging
Public Sub GetSlotAndValueFromCardIndex(ByVal CardIndex As Integer, ByRef OutSlot As Integer, ByRef OutValue As Byte)
    ' Validate card index range
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        OutSlot = 0
        OutValue = 0
        Exit Sub
    End If
    
    ' Calculate slot: ((CardIndex - 1) \ 8) + 1
    ' Integer division to get which byte
    OutSlot = ((CardIndex - 1) \ 8) + 1
    
    ' Calculate bit position: (CardIndex - 1) Mod 8
    ' Remainder gives us which bit within the byte
    Dim BitPosition As Integer
    BitPosition = (CardIndex - 1) Mod 8
    
    ' Convert bit position to mask value (2^BitPosition)
    Select Case BitPosition
        Case 0: OutValue = 1      ' 2^0
        Case 1: OutValue = 2      ' 2^1
        Case 2: OutValue = 4      ' 2^2
        Case 3: OutValue = 8      ' 2^3
        Case 4: OutValue = 16     ' 2^4
        Case 5: OutValue = 32     ' 2^5
        Case 6: OutValue = 64     ' 2^6
        Case 7: OutValue = 128    ' 2^7
    End Select
End Sub

' Alternative: Get card index directly from visual card number
' If your cards are numbered 1-1024 visually, this is a direct mapping
Public Function GetCardIndexFromVisualNumber(ByVal VisualCardNumber As Integer) As Integer
    If VisualCardNumber < 1 Or VisualCardNumber > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardIndexFromVisualNumber = 0
    Else
        GetCardIndexFromVisualNumber = VisualCardNumber
    End If
End Function

' Helper: Get quantity of a specific card by its index (1-1024)
Public Function GetCardQuantityByIndex(ByVal UserIndex As Integer, ByVal CardIndex As Integer) As Byte
    If CardIndex < 1 Or CardIndex > MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE Then
        GetCardQuantityByIndex = 0
    Else
        GetCardQuantityByIndex = UserList(UserIndex).AccountCollectibleCardQuantities(CardIndex)
    End If
End Function

