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
    
    ' Initialize the array with zeros
    For i = 1 To MAX_COLLECTIBLE_CARDS_QUANTITY_SIZE
        User.AccountCollectibleCardQuantities(i) = 0
    Next i
    
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
        .flags.DirtyCollectibleCardBitArray = True
        
        ' Increment the quantity for this specific card (max 255 per byte)
        If .AccountCollectibleCardQuantities(CardIndex) < 255 Then
            .AccountCollectibleCardQuantities(CardIndex) = .AccountCollectibleCardQuantities(CardIndex) + 1
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
