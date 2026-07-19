Attribute VB_Name = "ModCollectibleCards"
Option Explicit

Private Const UPSERT_NEW_COLLECTIBLE_CARDS As String = _
    "INSERT INTO account_collectible_cards (account_id, card_bit_array) " & _
    "VALUES (?, ?) " & _
    "ON CONFLICT(account_id) " & _
    "DO UPDATE SET card_bit_array = excluded.card_bit_array;"
    
Private Const GET_ACCOUNT_COLLECTIBLE_CARDS As String = "SELECT card_bit_array FROM account_collectible_cards WHERE account_id = ?"

Public Function SetupUserAccountAccountCollectibleCardBitArray(ByRef User As t_User)
    On Error GoTo GetUserCollectibleCards_Err
    Dim RS As ADODB.Recordset
    Dim mStream As ADODB.Stream
    Dim TempBuffer() As Byte
    Dim i As Integer
    
    Set RS = Query(GET_ACCOUNT_COLLECTIBLE_CARDS, User.AccountID)
    
    If Not RS.EOF Then
        ' 1. Configurar y cargar el Stream con el BLOB
        Set mStream = New ADODB.Stream
        mStream.Type = adTypeBinary
        mStream.Open
        mStream.Write RS("card_bit_array").value
        
        ' 2. Rebobinar el stream
        mStream.Position = 0
        
        ' 3. LEER TODO DE UNA SOLA VEZ (Recomendado por rendimiento)
        ' MAX_CARD_BIT_ARRAY debe ser la cantidad de BYTES (ej. 128 bytes para 1024 bits)
        TempBuffer = mStream.Read(128)
        
        ' 4. Copiar los datos al array de tu estructura
        ' Nota: Los arrays devueltos por Stream.Read siempre empiezan en el índice 0
        For i = 1 To 128
            ' Usamos (i - 1) porque TempBuffer es de base 0 (0 a MAX_CARD_BIT_ARRAY - 1)
            User.AccountCollectibleCardBitArray(i) = TempBuffer(i - 1)
        Next i
        mStream.Close
    End If
    Exit Function
GetUserCollectibleCards_Err:
    Call TraceError(Err.Number, Err.Description, "ModCollectibleCards.GetUserCollectibleCards", Erl)
End Function

Public Sub AddCollectibleCardToUser(ByVal UserIndex As Integer, ByRef ObjCard As t_Obj)
    If ObjCard.ObjIndex = 0 Then Exit Sub
    With UserList(UserIndex)
        .flags.DirtyCollectibleCardBitArray = True
        .AccountCollectibleCardBitArray(ObjData(ObjCard.ObjIndex).CollectibleCardSlot) = .AccountCollectibleCardBitArray(ObjData(ObjCard.ObjIndex).CollectibleCardSlot) & .AccountCollectibleCardBitArray(ObjData(ObjCard.ObjIndex).CollectibleCardValue)
    End With
End Sub
