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

Private Const UPSERT_NEW_COLLECTIBLE_CARDS As String = _
    "INSERT INTO account_collectible_cards (account_id, card_bit_array) " & _
    "VALUES (?, ?) " & _
    "ON CONFLICT(account_id) " & _
    "DO UPDATE SET card_bit_array = excluded.card_bit_array;"
    
Private Const GET_ACCOUNT_COLLECTIBLE_CARDS As String = "SELECT card_bit_array FROM account_collectible_cards WHERE account_id = ?"

Public Function SetupUserAccountAccountCollectibleCardBitArray(ByRef User As t_User)
    On Error GoTo GetUserCollectibleCards_Err
    Dim RS As ADODB.Recordset
    Dim Cmd As ADODB.Command
    Dim i As Integer
    Dim BlobData() As Byte
    
    ' Initialize the array with zeros
    For i = 1 To MAX_COLLECTIBLE_CARDS_ARR_SIZE
        User.AccountCollectibleCardBitArray(i) = 0
    Next i
    
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
    Dim i As Integer
    
    ' Copy user's byte array to 0-based array for database
    For i = 1 To MAX_COLLECTIBLE_CARDS_ARR_SIZE
        BlobData(i - 1) = UserList(UserIndex).AccountCollectibleCardBitArray(i)
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

