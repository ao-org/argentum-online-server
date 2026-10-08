Attribute VB_Name = "ModCastle"
Option Explicit

Private Type t_CastleCoordinates
    inside As t_WorldPos
    outside As t_WorldPos
End Type

Private Type t_CastleInfo
    id As Long
    owner_account_id As Long
    owner_char_id As Long
    owner_char_name As String
    spawner_obj_id As Integer
    inside_key_obj_id As Integer
    foundation_date As Date
    is_active As Boolean
    castle_coordinates As t_CastleCoordinates
    castleWhiteList As Dictionary
    dirtyWhiteList As Boolean
    dirtyCastleData As Boolean
    name As String
End Type

Public Enum eCastleWhitelistOperation
    Add = 0
    Remove = 1
End Enum

Public CastleData() As t_CastleInfo
Private CastleModuleLoaded As Boolean


Private Const COUNT_ALL_CASTLES As String = "SELECT COUNT(*) FROM castle;"

Private Const UPDATE_EMPEROR_CASTLE As String = "UPDATE castle SET owner_account_id = ?, owner_character_id = ?, foundation_date = ?, is_active = ?, name = ? WHERE id = ?;"
Private Const UPDATE_OUTSIDE_CASTLE_LOCATION As String = "UPDATE castle_coordinates SET outside_map = ?, outside_x = ?, outside_y = ? WHERE castle_id = ?;"

Private Const INSERT_OR_IGNORE_NEW_CHAR_IN_CASTLE_WHITELIST As String = "INSERT OR IGNORE INTO castle_whitelist (character_name, castle_id) VALUES (?,?)"
Private Const DELETE_CHAR_IN_CASTLE_WHITELIST As String = "DELETE FROM castle_whitelist WHERE id = ?"



Private Const SELECT_ALL_CASTLE_WHITELISTS As String = "Select * FROM castle_whitelist"
Private Const SELECT_ALL_CASTLES As String = "SELECT * FROM castle;"
Private Const SELECT_ALL_CASTLE_COORDINATES = "SELECT * FROM castle_coordinates;"

Private Const SELECT_SPECIFIC_CASTLE_WHITELIST As String = "Select * FROM castle_whitelist WHERE castle_id = ?;"

Private Const CastleXNegativeOffset As Integer = 8
Private Const CastleYNegativeOffset As Integer = 8
Private Const CastleXPositiveOffset As Integer = 6
Private Const CastleYPositiveOffset As Integer = 2
Private Const CASTLE_REPOSITION_COOLDOWN_IN_DAYS As Integer = 7

Private Const CASTLE_MOCKUP_OBJ_INDEX = 6382

Private Const CASTLE_SIGN_POST_OBJ_INDEX = 6419
Public Const EMPEROR_RELIC_OBJ_INDEX_1 = 6362
Public Const EMPEROR_RELIC_OBJ_INDEX_20 = 6381

Private Const CASTLE_NAME_PREFIXES As String = "Dragon's¬Eagle's¬Shadow¬Iron¬Stone¬Crystal¬Dark¬Golden¬Silver¬Frost¬Storm¬Blood¬Ancient¬Royal¬Mystic¬Thunder¬Moon¬Star¬Raven's¬Wolf's"
Private Const CASTLE_NAME_SUFFIXES As String = "Keep¬Fortress¬Citadel¬Bastion¬Tower¬Hold¬Stronghold¬Castle¬Sanctuary¬Haven¬Peak¬Spire¬Gate¬Watch¬Rest¬Reach¬Guard¬Crest¬Crown¬Throne"

' Add this function to generate random castle names
Public Function GenerateRandomCastleName() As String
    On Error GoTo GenerateRandomCastleName_Err
    
    Dim Prefixes() As String
    Dim Suffixes() As String
    Dim RandomPrefix As String
    Dim RandomSuffix As String
    
    ' Split the name lists
    Prefixes = Split(CASTLE_NAME_PREFIXES, "¬")
    Suffixes = Split(CASTLE_NAME_SUFFIXES, "¬")
    
    ' Generate random indices
    Randomize timer
    RandomPrefix = Prefixes(Int((UBound(Prefixes) - LBound(Prefixes) + 1) * Rnd + LBound(Prefixes)))
    RandomSuffix = Suffixes(Int((UBound(Suffixes) - LBound(Suffixes) + 1) * Rnd + LBound(Suffixes)))
    
    ' Combine them
    GenerateRandomCastleName = RandomPrefix & " " & RandomSuffix
    
    Exit Function
    
GenerateRandomCastleName_Err:
    Call TraceError(Err.Number, Err.Description, "ModCastle.GenerateRandomCastleName", Erl)
    GenerateRandomCastleName = "Unnamed Castle"
End Function

Private Function IsCastleFootprintInMapBounds(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    IsCastleFootprintInMapBounds = False

    If Not InMapBounds(map, x - CastleXNegativeOffset, y - CastleYNegativeOffset) Then Exit Function
    If Not InMapBounds(map, x + CastleXPositiveOffset, y + CastleYPositiveOffset) Then Exit Function

    IsCastleFootprintInMapBounds = True
End Function

Private Sub ValidateCastleSchema()
    On Error GoTo ValidateCastleSchema_Err
    Dim RS As ADODB.Recordset, nullableOutsideColumns As Integer
    Dim failureNumber As Long, failureDescription As String
    Set RS = Query("SELECT date FROM migrations WHERE date = ?;", "20261007-01")
    If RS Is Nothing Then Call Err.Raise(5, , "Missing database migration history")
    If RS.EOF Then Call Err.Raise(5, , "Apply ScriptsDB/20261007-01-migrate castle identities.sql before loading castles")
    Call CloseCastleRecordset(RS)
    Set RS = Query("PRAGMA table_info(castle);")
    If RS Is Nothing Then Call Err.Raise(5, , "Cannot inspect castle schema")
    If RS.EOF Then Call Err.Raise(5, , "Missing castle table")
    Do While Not RS.EOF
        If LCase$(CStr(RS!name)) = "trigger" Then Call Err.Raise(5, , "Legacy castle.trigger remains after migration")
        Call RS.MoveNext()
    Loop
    Call CloseCastleRecordset(RS)
    Set RS = Query("PRAGMA table_info(castle_coordinates);")
    If RS Is Nothing Then Call Err.Raise(5, , "Cannot inspect castle coordinate schema")
    Do While Not RS.EOF
        Select Case LCase$(CStr(RS!name))
            Case "outside_map", "outside_x", "outside_y"
                If CLng(RS.Fields("notnull").value) <> 0 Then Call Err.Raise(5, , "Outside castle coordinates must allow NULL")
                nullableOutsideColumns = nullableOutsideColumns + 1
        End Select
        Call RS.MoveNext()
    Loop
    If nullableOutsideColumns <> 3 Then Call Err.Raise(5, , "Missing outside castle coordinate columns")
    Call CloseCastleRecordset(RS)
    Exit Sub
ValidateCastleSchema_Err:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Call CloseCastleRecordset(RS)
    Call Err.Raise(failureNumber, "ValidateCastleSchema", failureDescription)
End Sub

Public Sub LoadCastleModule()
    On Error GoTo LoadCastleModule_Err
    Dim failureNumber As Long, failureDescription As String
    CastleModuleLoaded = False
    Call ValidateCastleSchema()
    Call LoadCastleData
    Call LoadCastleCoordinates
    Call LoadCastleWhiteLists
    Call ResolveLegacyCastleEntrances()
    Dim i As Integer

    For i = LBound(CastleData) To UBound(CastleData)
        With CastleData(i)
            If .is_active And .castle_coordinates.outside.map > 0 Then
                If Not CanPublishCastle(i, .castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y) Then Call Err.Raise(5, "LoadCastleModule", "Invalid or conflicting castle placement: " & CStr(.id))
                Call CreateCastleInMap(.castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y, i)
            End If
        End With
    Next i
    CastleModuleLoaded = True

    Exit Sub
LoadCastleModule_Err:
    CastleModuleLoaded = False
    failureNumber = Err.Number
    failureDescription = Err.Description
    Call TraceError(failureNumber, failureDescription, "ModCastle.LoadCastleModule", Erl)
    Call Err.Raise(failureNumber, "ModCastle.LoadCastleModule", failureDescription)
End Sub

Public Sub LoadCastleWhiteLists()
    On Error GoTo LoadCastleWhitelists_Err
    Dim RS As ADODB.Recordset
    Dim failureNumber As Long, failureDescription As String
    Set RS = Query(SELECT_ALL_CASTLE_WHITELISTS)
    If RS Is Nothing Then Call Err.Raise(5, "LoadCastleWhiteLists", "Castle query failed")
    If RS.EOF Then
        Call CloseCastleRecordset(RS)
        Exit Sub
    End If
    
    Do While Not RS.EOF
        Dim CastleSlot As Integer
        CastleSlot = GetCastleSlotById(RS!castle_id)
        If CastleSlot < 1 Then Call Err.Raise(5, "LoadCastleWhiteLists", "Unknown castle ID")
        Dim CharacterName As String
        CharacterName = LCase$(CStr(RS!character_name))
        Call AddUserNameToWhiteListByCastleSlot(CastleSlot, CharacterName)
        Call RS.MoveNext()
    Loop
    Call RS.Close()
    Exit Sub
LoadCastleWhitelists_Err:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Call CloseCastleRecordset(RS)
    Call TraceError(failureNumber, failureDescription, "ModCastle.LoadCastleWhiteLists", Erl)
    Call Err.Raise(failureNumber, "ModCastle.LoadCastleWhiteLists", failureDescription)
End Sub


Public Function AddUserNameToWhiteListByCastleSlot(ByVal CastleSlot As Integer, ByVal CharacterName As String) As Boolean
    AddUserNameToWhiteListByCastleSlot = False
    CharacterName = LCase$(CharacterName)
    With CastleData(CastleSlot)
        'duplicated entry somewhere
        Debug.Assert Not .castleWhiteList.Exists(CharacterName)
        If .castleWhiteList.Exists(CharacterName) Then
            Call LogInfoServidor("Duplicated username: " & CharacterName & " in whitelist for castle " & .name)
            Exit Function
        End If
        Call .castleWhiteList.Add(CharacterName, True)
        Call LogInfoServidor("Username:" & CharacterName & " was added to the whitelist of castle: " & .name)
        AddUserNameToWhiteListByCastleSlot = True
    End With
End Function

Public Function RemoveUserNameToWhiteListByCastleSlot(ByVal CastleSlot As Integer, ByVal CharacterName As String) As Boolean
    RemoveUserNameToWhiteListByCastleSlot = False
    CharacterName = LCase$(CharacterName)
    With CastleData(CastleSlot)
        'duplicated entry somewhere
        Debug.Assert .castleWhiteList.Exists(CharacterName)
        If Not .castleWhiteList.Exists(CharacterName) Then
            Call LogInfoServidor("Tried to remove from whitelist " & CharacterName & " but was already removed ")
            Exit Function
        End If
        Call .castleWhiteList.Remove(CharacterName)
        Call LogInfoServidor("Username:" & CharacterName & " was removed to the whitelist of castle: " & .name)
        RemoveUserNameToWhiteListByCastleSlot = True
    End With
End Function



Public Function GetCastleSlotById(ByVal id As Long) As Integer
    GetCastleSlotById = -1
    Dim i As Integer
    
    For i = LBound(CastleData) To UBound(CastleData)
        If CastleData(i).id = id Then
            GetCastleSlotById = i
            Exit Function
        End If
    Next i
End Function


Public Sub LoadCastleData()
    On Error GoTo LoadCastleData_Err
    Dim RS As ADODB.Recordset
    Dim failureNumber As Long, failureDescription As String
    Set RS = Query(COUNT_ALL_CASTLES)
    If RS Is Nothing Then
        Call Err.Raise(5, "LoadCastleData", "Castle rows could not be initialized")
    End If
    ReDim CastleData(1 To RS.Fields(0).value)
    Call RS.Close()

    Dim i As Long
    i = 1
    Set RS = Query(SELECT_ALL_CASTLES)
    If RS Is Nothing Then Call Err.Raise(5, "LoadCastleData", "Castle query failed")
    If RS.EOF Then Call Err.Raise(5, "LoadCastleData", "Castle rows missing after count query")
    If RS.RecordCount <> UBound(CastleData) Then
        Call Err.Raise(5, "LoadCastleData", "Castle rows could not be initialized")
    End If

    Do While Not RS.EOF
    
        With CastleData(i)
            .id = (RS!id)
            
            
            If Not IsNull(RS!name) Then
                .name = (RS!name)
            End If
            If Not IsNull(RS!owner_account_id) Then
                .owner_account_id = (RS!owner_account_id)
            End If
    
            If Not IsNull(RS!owner_character_id) Then
                .owner_char_id = (RS!owner_character_id)
                .owner_char_name = LCase$(GetCharacterNameByUserId(RS!owner_character_id))
            End If
            
            If Not IsNull(RS!foundation_date) Then
                .foundation_date = (RS!foundation_date)
            End If
            
            Set .castleWhiteList = New Dictionary
            
            .spawner_obj_id = (RS!spawner_obj_id)
            .inside_key_obj_id = (RS!inside_key_obj_id)
            .is_active = (RS!is_active)
            i = i + 1
            RS.MoveNext
        
        End With
        
    Loop
    Call RS.Close()
    Exit Sub
LoadCastleData_Err:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Call CloseCastleRecordset(RS)
    Call TraceError(failureNumber, failureDescription, "ModCastle.LoadCastleData", Erl)
    Call Err.Raise(failureNumber, "ModCastle.LoadCastleData", failureDescription)
End Sub


Public Sub LoadCastleCoordinates()
    On Error GoTo LoadCastleCoordinates_Err
    Dim i As Integer
    Dim RS As ADODB.Recordset
    Dim failureNumber As Long, failureDescription As String
    Set RS = Query(SELECT_ALL_CASTLE_COORDINATES)
    If RS Is Nothing Then Call Err.Raise(5, "LoadCastleCoordinates", "Castle query failed")
    If RS.EOF Then
        Call CloseCastleRecordset(RS)
        Exit Sub
    End If
    Do While Not RS.EOF
        i = GetCastleSlotById(CLng(RS!castle_id))
        If i < 1 Then Call Err.Raise(5, "LoadCastleCoordinates", "Unknown castle ID")
        If Not IsNull(RS!outside_map) Then
            CastleData(i).castle_coordinates.outside.map = (RS!outside_map)
            CastleData(i).castle_coordinates.outside.x = (RS!outside_x)
            CastleData(i).castle_coordinates.outside.y = (RS!outside_y)
        End If
        CastleData(i).castle_coordinates.inside.map = (RS!inside_map)
        CastleData(i).castle_coordinates.inside.x = (RS!inside_x)
        CastleData(i).castle_coordinates.inside.y = (RS!inside_y)
        Call RS.MoveNext()
    Loop
    Call RS.Close()
Exit Sub
LoadCastleCoordinates_Err:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Call CloseCastleRecordset(RS)
    Call TraceError(failureNumber, failureDescription, "ModCastle.LoadCastleCoordinates", Erl)
    Call Err.Raise(failureNumber, "ModCastle.LoadCastleCoordinates", failureDescription)
End Sub


Public Function IsValidCastlePosition(ByVal UserIndex As Integer) As Boolean
    IsValidCastlePosition = False


    Dim CastleTopLeftCorner As t_WorldPos
    Dim CastleBottomRightCorner As t_WorldPos

    Dim UserTargetX As Integer
    Dim UserTargetY As Integer
    Dim UserTargetMap As Integer

    With UserList(UserIndex)

        If .flags.TargetX = 0 Or .flags.TargetY = 0 Or .flags.TargetMap = 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Function
        End If

        CastleTopLeftCorner.x = .flags.TargetX - CastleXNegativeOffset
        CastleTopLeftCorner.y = .flags.TargetY - CastleYNegativeOffset
        CastleTopLeftCorner.map = .flags.TargetMap

        CastleBottomRightCorner.x = .flags.TargetX + CastleXPositiveOffset
        CastleBottomRightCorner.y = .flags.TargetY + CastleYPositiveOffset
        CastleBottomRightCorner.map = .flags.TargetMap

        UserTargetX = .flags.TargetX
        UserTargetY = .flags.TargetY
        UserTargetMap = .flags.TargetMap

    End With

    If UserList(UserIndex).pos.map <> UserTargetMap Then
        Call LogError("Usuario " & UserList(UserIndex).name & "Interactuando con un mapa fuera de su rango, revisar")
        Exit Function
    End If

    If Not IsValidMapIndex(UserTargetMap) Then
        Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD)
        Exit Function
    End If

    If Not IsCastleFootprintInMapBounds(UserTargetMap, UserTargetX, UserTargetY) Then
        Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD)
        Exit Function
    End If

    If Not HasTileFlag(MapData(UserTargetX, UserTargetY, UserTargetMap).trigger, e_Trigger.CastleFoundationPosition) Then
        Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD)
        Exit Function
    End If

    If MapData(UserTargetX, UserTargetY, UserTargetMap).ObjInfo.ObjIndex = CASTLE_MOCKUP_OBJ_INDEX Then
        Call WriteLocaleMsg(UserIndex, MSG_CANT_FOUND_CASTLE_ON_TOP_OF_ANOTHER, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD)
        Exit Function
    End If

    If MapData(UserTargetX, UserTargetY, UserTargetMap).ObjInfo.ObjIndex <> CASTLE_SIGN_POST_OBJ_INDEX Then
        Call WriteLocaleMsg(UserIndex, MSG_CANOT_FOUND_WITHOUT_SIGN, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD)
        Exit Function
    End If

    Dim i As Integer
    Dim j As Integer
    For i = CastleTopLeftCorner.x To CastleBottomRightCorner.x
        For j = CastleTopLeftCorner.y To CastleBottomRightCorner.y

            If Not InMapBounds(UserTargetMap, i, j) Then
                Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD)
                Exit Function
            End If

        Next j
    Next i

    IsValidCastlePosition = True
End Function


Public Sub CreateCastleInMap(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal CastleIndex As Integer, Optional ByVal UserIndex As Integer = 0)
    If Not IsCastleFootprintInMapBounds(map, x, y) Then
        Call LogInfoServidor("CreateCastleInMap outside map bounds. map=" & CStr(map) & _
            " x=" & CStr(x) & _
            " y=" & CStr(y) & _
            " CastleIndex=" & CStr(CastleIndex) & _
            " UserIndex=" & CStr(UserIndex))
        Exit Sub
    End If

    If Not CanPublishCastle(CastleIndex, map, x, y) Then
        Call LogInfoServidor("Castle placement invalid or conflicting: " & CStr(CastleIndex))
        Exit Sub
    End If
    Call RemoveCastleEntrances(map, CastleData(CastleIndex).id)
    With CastleData(CastleIndex)

        'if not during server start...(player clicking the board)
         If UserIndex > 0 Then
            .castle_coordinates.outside.map = map
            .castle_coordinates.outside.x = x
            .castle_coordinates.outside.y = y
            .foundation_date = DateTime.Now
            .is_active = True
            .owner_account_id = UserList(UserIndex).AccountID
            .owner_char_id = UserList(UserIndex).id
            .owner_char_name = UserList(UserIndex).name
            .dirtyCastleData = True
            .name = GenerateRandomCastleName
        End If

        Dim CastleTopLeftCorner As t_WorldPos
        Dim CastleBottomRightCorner As t_WorldPos
        CastleTopLeftCorner.x = x - CastleXNegativeOffset
        CastleTopLeftCorner.y = y - CastleYNegativeOffset
        CastleTopLeftCorner.map = map

        CastleBottomRightCorner.x = x + CastleXPositiveOffset
        CastleBottomRightCorner.y = y + CastleYPositiveOffset
        CastleBottomRightCorner.map = map

        'erase preemptively all blocks, triggers, objects and npcs in the zone
        Dim i As Integer
        Dim j As Integer
        For i = CastleTopLeftCorner.x To CastleBottomRightCorner.x
            For j = CastleTopLeftCorner.y To CastleBottomRightCorner.y

            MapData(i, j, map).Blocked = 0

            If MapData(i, j, map).ObjInfo.ObjIndex > 0 Then
                Call EraseObj(MapData(i, j, map).ObjInfo.Amount, map, i, j)
            End If

            If MapData(i, j, map).NpcIndex > 0 Then
                Call QuitarNPC(MapData(i, j, map).NpcIndex, eAiResetNpc)
            End If

            Next j
        Next i


        'first layer from the bottom
        MapData(x - 3, y, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 4, y, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 5, y, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 6, y, map).Blocked = e_Block.ALL_SIDES
        MapData(x, y, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 1, y, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 2, y, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 3, y, map).Blocked = e_Block.ALL_SIDES
        MapData(x, y, map).trigger = MapData(x, y, map).trigger Or e_Trigger.CastleFoundationPosition
        Call RegisterCastleEntrance(map, x - 1, y, .id)
        Call RegisterCastleEntrance(map, x - 2, y, .id)
        Call RegisterCastleEntrance(map, x - 1, y + 1, .id)
        Call RegisterCastleEntrance(map, x - 2, y + 1, .id)


        'second layer form the bottom
        MapData(x, y - 1, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 3, y - 1, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 4, y - 1, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 5, y - 1, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 6, y - 1, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 1, y - 1, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 2, y - 1, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 3, y - 1, map).Blocked = e_Block.ALL_SIDES

        MapData(x - 1, y - 1, map).TileExit.map = .castle_coordinates.inside.map
        MapData(x - 1, y - 1, map).TileExit.x = .castle_coordinates.inside.x
        MapData(x - 1, y - 1, map).TileExit.y = .castle_coordinates.inside.y

        MapData(x - 2, y - 1, map).TileExit.map = .castle_coordinates.inside.map
        MapData(x - 2, y - 1, map).TileExit.x = .castle_coordinates.inside.x
        MapData(x - 2, y - 1, map).TileExit.y = .castle_coordinates.inside.y

        'third layer form the bottom
        MapData(x, y - 2, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 1, y - 2, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 2, y - 2, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 3, y - 2, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 4, y - 2, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 5, y - 2, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 6, y - 2, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 1, y - 2, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 2, y - 2, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 3, y - 2, map).Blocked = e_Block.ALL_SIDES

         'fourth layer form the bottom
        MapData(x, y - 3, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 1, y - 3, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 2, y - 3, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 3, y - 3, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 4, y - 3, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 5, y - 3, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 6, y - 3, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 1, y - 3, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 2, y - 3, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 3, y - 3, map).Blocked = e_Block.ALL_SIDES

         'fifth layer form the bottom
        MapData(x, y - 4, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 1, y - 4, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 2, y - 4, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 3, y - 4, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 4, y - 4, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 5, y - 4, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 6, y - 4, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 1, y - 4, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 2, y - 4, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 3, y - 4, map).Blocked = e_Block.ALL_SIDES

         'sixth layer form the bottom
        MapData(x, y - 5, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 1, y - 5, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 2, y - 5, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 3, y - 5, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 4, y - 5, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 5, y - 5, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 6, y - 5, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 1, y - 5, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 2, y - 5, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 3, y - 5, map).Blocked = e_Block.ALL_SIDES

         'seventh layer form the bottom
        MapData(x, y - 6, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 1, y - 6, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 2, y - 6, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 3, y - 6, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 4, y - 6, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 5, y - 6, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 6, y - 6, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 1, y - 6, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 2, y - 6, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 3, y - 6, map).Blocked = e_Block.ALL_SIDES

         'eighth layer form the bottom
        MapData(x, y - 7, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 1, y - 7, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 2, y - 7, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 3, y - 7, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 4, y - 7, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 5, y - 7, map).Blocked = e_Block.ALL_SIDES
        MapData(x - 6, y - 7, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 1, y - 7, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 2, y - 7, map).Blocked = e_Block.ALL_SIDES
        MapData(x + 3, y - 7, map).Blocked = e_Block.ALL_SIDES

        'create castle inside tile exits to the outside part
        If Not InMapBounds(.castle_coordinates.inside.map, .castle_coordinates.inside.x, .castle_coordinates.inside.y + 1) Then
            Call LogInfoServidor("CreateCastleInMap invalid inside exit 1. map=" & CStr(.castle_coordinates.inside.map) & _
                " x=" & CStr(.castle_coordinates.inside.x) & _
                " y=" & CStr(.castle_coordinates.inside.y + 1) & _
                " CastleIndex=" & CStr(CastleIndex))
            Exit Sub
        End If

        If Not InMapBounds(.castle_coordinates.inside.map, .castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1) Then
            Call LogInfoServidor("CreateCastleInMap invalid inside exit 2. map=" & CStr(.castle_coordinates.inside.map) & _
                " x=" & CStr(.castle_coordinates.inside.x + 1) & _
                " y=" & CStr(.castle_coordinates.inside.y + 1) & _
                " CastleIndex=" & CStr(CastleIndex))
            Exit Sub
        End If

        MapData(.castle_coordinates.inside.x, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.map = .castle_coordinates.outside.map
        MapData(.castle_coordinates.inside.x, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.x = .castle_coordinates.outside.x - 2
        MapData(.castle_coordinates.inside.x, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.y = .castle_coordinates.outside.y + 1

        MapData(.castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.map = .castle_coordinates.outside.map
        MapData(.castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.x = .castle_coordinates.outside.x - 1
        MapData(.castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.y = .castle_coordinates.outside.y + 1

        'erase castle sign
        If MapData(x, y, map).ObjInfo.Amount > 0 Then
            Call EraseObj(MapData(x, y, map).ObjInfo.Amount, map, x, y)
        End If
        'create castle visual mockup
        Dim CastleObj As t_Obj
        CastleObj.Amount = 1
        CastleObj.ObjIndex = CASTLE_MOCKUP_OBJ_INDEX
        CastleObj.CastleSlot = CastleIndex
        Call MakeObj(CastleObj, map, x, y)

    End With

End Sub


Public Sub DestroyCastleInMap(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal CastleIndex As Integer)
    If Not IsCastleFootprintInMapBounds(map, x, y) Then
        Call LogInfoServidor("DestroyCastleInMap outside map bounds. map=" & CStr(map) & _
            " x=" & CStr(x) & _
            " y=" & CStr(y) & _
            " CastleIndex=" & CStr(CastleIndex))
        Exit Sub
    End If

    Call RemoveCastleEntrances(map, CastleData(CastleIndex).id)
    If MapData(x, y, map).ObjInfo.Amount > 0 Then
        Call EraseObj(MapData(x, y, map).ObjInfo.Amount, map, x, y)
    End If

     'remove everything
    Dim CastleTopLeftCorner As t_WorldPos
    Dim CastleBottomRightCorner As t_WorldPos
    CastleTopLeftCorner.x = x - CastleXNegativeOffset
    CastleTopLeftCorner.y = y - CastleYNegativeOffset
    CastleTopLeftCorner.map = map

    CastleBottomRightCorner.x = x + CastleXPositiveOffset
    CastleBottomRightCorner.y = y + CastleYPositiveOffset
    CastleBottomRightCorner.map = map

    'erase preemptively all blocks, triggers, objects and npcs in the zone
    Dim i As Integer
    Dim j As Integer
    For i = CastleTopLeftCorner.x To CastleBottomRightCorner.x
        For j = CastleTopLeftCorner.y To CastleBottomRightCorner.y

        MapData(i, j, map).Blocked = 0

        If MapData(i, j, map).ObjInfo.ObjIndex > 0 Then
            Call EraseObj(MapData(i, j, map).ObjInfo.Amount, map, i, j)
        End If

        If MapData(i, j, map).NpcIndex > 0 Then
            Call QuitarNPC(MapData(i, j, map).NpcIndex, eAiResetNpc)
        End If

        Next j
    Next i

    MapData(x - 1, y - 1, map).TileExit.map = 0
    MapData(x - 1, y - 1, map).TileExit.x = 0
    MapData(x - 1, y - 1, map).TileExit.y = 0

    MapData(x - 2, y - 1, map).TileExit.map = 0
    MapData(x - 2, y - 1, map).TileExit.x = 0
    MapData(x - 2, y - 1, map).TileExit.y = 0

     'restore castle foundation trigger
    MapData(x, y, map).trigger = MapData(x, y, map).trigger Or e_Trigger.CastleFoundationPosition

     With CastleData(CastleIndex)
        If Not InMapBounds(.castle_coordinates.inside.map, .castle_coordinates.inside.x, .castle_coordinates.inside.y + 1) Then
            Call LogInfoServidor("DestroyCastleInMap invalid inside exit 1. map=" & CStr(.castle_coordinates.inside.map) & _
                " x=" & CStr(.castle_coordinates.inside.x) & _
                " y=" & CStr(.castle_coordinates.inside.y + 1) & _
                " CastleIndex=" & CStr(CastleIndex))
            Exit Sub
        End If

        If Not InMapBounds(.castle_coordinates.inside.map, .castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1) Then
            Call LogInfoServidor("DestroyCastleInMap invalid inside exit 2. map=" & CStr(.castle_coordinates.inside.map) & _
                " x=" & CStr(.castle_coordinates.inside.x + 1) & _
                " y=" & CStr(.castle_coordinates.inside.y + 1) & _
                " CastleIndex=" & CStr(CastleIndex))
            Exit Sub
        End If

        MapData(.castle_coordinates.inside.x, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.map = 0
        MapData(.castle_coordinates.inside.x, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.x = 0
        MapData(.castle_coordinates.inside.x, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.y = 0

        MapData(.castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.map = 0
        MapData(.castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.x = 0
        MapData(.castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1, .castle_coordinates.inside.map).TileExit.y = 0
    End With

    'create castle sign post
    Dim CastleSignObj As t_Obj
    CastleSignObj.Amount = 1
    CastleSignObj.ObjIndex = CASTLE_SIGN_POST_OBJ_INDEX
    Call MakeObj(CastleSignObj, map, x, y)
End Sub

Public Function IsEmperorCastleCreated(ByVal UserIndex As Integer, Optional ByVal castleId As Long = 0, Optional ByRef CastleIndex As Integer = -1) As Boolean
    IsEmperorCastleCreated = False
    Dim i As Integer
    For i = 1 To UBound(CastleData)
        With CastleData(i)
            If (castleId > 0 And .id = castleId) Or (castleId = 0 And .owner_account_id = UserList(UserIndex).AccountID) Then
                If Not .is_active Then Exit Function
                If Not IsCastleFootprintInMapBounds(.castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y) Then
                    Call LogInfoServidor("IsEmperorCastleCreated outside map bounds. map=" & CStr(.castle_coordinates.outside.map) & _
                        " x=" & CStr(.castle_coordinates.outside.x) & _
                        " y=" & CStr(.castle_coordinates.outside.y) & _
                        " CastleIndex=" & CStr(i))
                    Exit Function
                End If

                If (MapData(.castle_coordinates.outside.x, .castle_coordinates.outside.y, .castle_coordinates.outside.map).ObjInfo.ObjIndex = CASTLE_MOCKUP_OBJ_INDEX) Then
                    IsEmperorCastleCreated = True
                    CastleIndex = i
                End If
                Exit For
            End If
        End With
    Next i
End Function

Public Function HasCastleRelocationCooldownPassed(ByVal CastleIndex As Integer) As Boolean
HasCastleRelocationCooldownPassed = False
    Dim Acumulator As Long
    Acumulator = DateTime.Now - CastleData(CastleIndex).foundation_date
    If Acumulator >= CASTLE_REPOSITION_COOLDOWN_IN_DAYS Then
        HasCastleRelocationCooldownPassed = True
    End If
End Function

Public Sub CreateNewEmperorCastle(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo CreateEmperorCastle_Err
    Dim RS As ADODB.Recordset
    Dim castleSlot As Integer
    castleSlot = GetCastleSlotById(ObjData(ObjIndex).AssignedCastleIndex)
    If castleSlot < 1 Then Exit Sub
    If CastleData(castleSlot).owner_account_id <> 0 And CastleData(castleSlot).owner_account_id <> UserList(UserIndex).AccountID Then
        Call LogInfoServidor("Castle relocation denied: account does not own castle " & CStr(CastleData(castleSlot).id))
        Exit Sub
    End If
    With UserList(UserIndex)

        If Not CanPublishCastle(castleSlot, .flags.TargetMap, .flags.TargetX, .flags.TargetY) Then
            Call LogInfoServidor("Castle relocation rejected before removing old placement")
            Exit Sub
        End If
        If IsEmperorCastleCreated(UserIndex, CastleData(castleSlot).id) Then
            If Not HasCastleRelocationCooldownPassed(castleSlot) Then
                Call WriteLocaleMsg(UserIndex, MSG_CASTLE_RELOCATION_ON_COOLDOWN, e_TextChannel.TEXTCHANNEL_EVENT, e_FontTypeNames.FONTTYPE_New_Eventos)
                Exit Sub
            End If
            With CastleData(castleSlot)
                Call DestroyCastleInMap(.castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y, castleSlot)
                Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_BROADCAST_CASTLE_DESTROYED, castleSlot & "¬" & GetUserDisplayName(UserIndex), e_TextChannel.TEXTCHANNEL_GUILD, e_FontTypeNames.FONTTYPE_GUILD))
            End With
        End If

        Call CreateCastleInMap(.flags.TargetMap, .flags.TargetX, .flags.TargetY, castleSlot, UserIndex)
        
        Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_BROADCAST_CASTLE_LOCATION, .name & "¬" & GetUserDisplayName(UserIndex) & "¬" & .flags.TargetMap & "¬" & .flags.TargetX & "¬" & .flags.TargetY, e_TextChannel.TEXTCHANNEL_GUILD, e_FontTypeNames.FONTTYPE_GUILD))
        Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(e_SoundEffects.OldClanHorn, 50, 50))
        Call modSendData.SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(e_SoundEffects.NewCastleRPGVoice, 50, 50))
    End With
    Exit Sub
CreateEmperorCastle_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.CreateEmperorCastle", Erl)
End Sub

Function CheckCastleEntryWhiteList(ByVal UserIndex As Integer, ByVal castleId As Long) As Boolean
   CheckCastleEntryWhiteList = False

    Dim CastleIndex As Integer
    If Not IsEmperorCastleCreated(UserIndex, castleId, CastleIndex) Then
        Exit Function
    End If

    With CastleData(CastleIndex)

       'exception for castle owner
       If UserList(UserIndex).AccountID = .owner_account_id Then
            CheckCastleEntryWhiteList = True
            Exit Function
       End If
    
       If Not .castleWhiteList.Exists(LCase$(UserList(UserIndex).name)) Then
            Exit Function
       End If
    
       CheckCastleEntryWhiteList = True
   
   End With
   
End Function

Public Sub ModifyCastleEntryWhiteList(ByVal UserIndex As Integer, ByVal CharacterName As String, ByVal operation As eCastleWhitelistOperation)
    CharacterName = LCase$(CharacterName)
    With UserList(UserIndex)
        If Not ValidarNombre(CharacterName) Then
            Call LogInfoServidor("User: " & .name & " Tried to input an invalid charactername while modifying castle whitelist")
            Exit Sub
        End If
    
        If .Stats.tipoUsuario < e_TipoUsuario.tNoble Then
                Call WriteLocaleMsg(UserIndex, MSG_AT_LEAST_NOBLE_TO_FOUND_CASTLE, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD)
                Call LogInfoServidor("User with low patreon status trying to set a whitelist for a castle, name: " & .name)
                Exit Sub
        End If
        
        Dim CastleIndex As Integer
        If Not IsEmperorCastleCreated(UserIndex, 0, CastleIndex) Then
            Call LogError("Couldn't find the castle for user: " & .name)
            Debug.Assert False
            Exit Sub
        End If
    
        If CastleIndex <= -1 Then
            Call LogError("Couldn't find the castle for user: " & .name)
            Debug.Assert False
        End If
        
        Select Case operation
            Case eCastleWhitelistOperation.Add
                If AddUserNameToWhiteListByCastleSlot(CastleIndex, CharacterName) Then
                    CastleData(CastleIndex).dirtyWhiteList = True
                    Call WriteLocaleMsg(UserIndex, MSG_CHARNAME_ADDED_TO_WHITELIST, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD, CharacterName)
                End If
            Case eCastleWhitelistOperation.Remove
                If RemoveUserNameToWhiteListByCastleSlot(CastleIndex, CharacterName) Then
                    CastleData(CastleIndex).dirtyWhiteList = True
                    Call WriteLocaleMsg(UserIndex, MSG_CHARNAME_REMOVED_FROM_WHITELIST, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD, CharacterName)
                End If
            Case Else
                Call LogInfoServidor("User: " & .name & " used an invalid operation for modify castle white list")
        End Select

    End With
End Sub

Public Sub SaveCastleWhiteListToDb()
    On Error GoTo SaveCastleWhiteListToDb_Err
    Dim i As Integer, keyName As Variant
    Dim savedNames As Dictionary
    Dim RS As ADODB.Recordset, result As ADODB.Recordset
    For i = LBound(CastleData) To UBound(CastleData)
        With CastleData(i)
            If .dirtyWhiteList Then
                Set RS = Query(SELECT_SPECIFIC_CASTLE_WHITELIST, .id)
                If RS Is Nothing Then Exit Sub
                Set savedNames = New Dictionary
                Do While Not RS.EOF
                    savedNames.Item(LCase$(CStr(RS!character_name))) = True
                    If Not .castleWhiteList.Exists(LCase$(CStr(RS!character_name))) Then
                        Set result = Query(DELETE_CHAR_IN_CASTLE_WHITELIST, RS!id)
                        If result Is Nothing Then Exit Sub
                        Call CloseCastleRecordset(result)
                    End If
                    Call RS.MoveNext()
                Loop
                Call RS.Close()
                For Each keyName In .castleWhiteList.Keys
                    If Not savedNames.Exists(keyName) Then
                        Set result = Query(INSERT_OR_IGNORE_NEW_CHAR_IN_CASTLE_WHITELIST, keyName, .id)
                        If result Is Nothing Then Exit Sub
                        Call CloseCastleRecordset(result)
                    End If
                Next keyName
                .dirtyWhiteList = False
            End If
        End With
    Next i
    Exit Sub
SaveCastleWhiteListToDb_Err:
    Call CloseCastleRecordset(RS)
    Call CloseCastleRecordset(result)
    Call TraceError(Err.Number, Err.Description, "SaveCastleWhiteListToDb", Erl)
End Sub

Private Sub CloseCastleRecordset(ByRef recordset As ADODB.Recordset)
    If recordset Is Nothing Then Exit Sub
    If recordset.State = adStateOpen Then Call recordset.Close()
    Set recordset = Nothing
End Sub

Public Sub SaveCastleDataToDb()
    Dim i As Integer
    Dim RS As ADODB.Recordset
    For i = LBound(CastleData) To UBound(CastleData)
        With (CastleData(i))
            
            If .dirtyCastleData Then
                'update castle data in db
                Set RS = Query(UPDATE_EMPEROR_CASTLE, .owner_account_id, .owner_char_id, DateToSQLite(.foundation_date), Abs(CInt(.is_active)), .name, .id)
                If RS Is Nothing Then Exit Sub
                Call CloseCastleRecordset(RS)
                'update castle coordinates in db
                If .castle_coordinates.outside.map = 0 Then
                    Set RS = Query(UPDATE_OUTSIDE_CASTLE_LOCATION, Null, Null, Null, .id)
                Else
                    Set RS = Query(UPDATE_OUTSIDE_CASTLE_LOCATION, .castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y, .id)
                End If
                If RS Is Nothing Then Exit Sub
                Call CloseCastleRecordset(RS)
                .dirtyCastleData = False
                Call LogInfoServidor("Persisted new data for castle number: " & i & " name: " & .name)
            End If
            
        End With
    Next i
End Sub

Public Sub SaveCastlesToDb()
    Call SaveCastleDataToDb
    Call SaveCastleWhiteListToDb
End Sub

Public Sub SendCastleInfo(ByVal UserIndex As Integer, ByVal CastleSlot As Integer)
    If CastleSlot <= 0 Then
        Debug.Assert False
    End If

    With CastleData(CastleSlot)
        Call WriteLocaleMsg(UserIndex, MSG_CASTLE_FOUNDER_INFO, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFOBOLD, .name & "¬" & .foundation_date & "¬" & .owner_char_name)
    End With
End Sub







Private Function CanPublishCastle(ByVal castleIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If castleIndex < LBound(CastleData) Or castleIndex > UBound(CastleData) Then Exit Function
    If Not IsCastleFootprintInMapBounds(map, x, y) Then Exit Function
    With CastleData(castleIndex).castle_coordinates.inside
        If Not InMapBounds(.map, .x, .y + 1) Then Exit Function
        If Not InMapBounds(.map, .x + 1, .y + 1) Then Exit Function
    End With
    Dim dx As Integer, dy As Integer, existing As Long, key As Long
    For dx = -2 To -1
        For dy = 0 To 1
            key = TilePropertyKey(x + dx, y + dy)
            If Not MapInfo(map).StaticCastleEntrances Is Nothing Then
                If MapInfo(map).StaticCastleEntrances.Exists(key) Then Exit Function
            End If
            existing = CastleAtTile(map, x + dx, y + dy)
            If existing <> 0 And existing <> CastleData(castleIndex).id Then Exit Function
        Next dy
    Next dx
    CanPublishCastle = True
End Function

Public Sub ResolveLegacyCastleEntrances()
    On Error GoTo ResolveLegacyCastleEntrances_Err
    Dim failureNumber As Long, failureDescription As String
    Dim map As Integer, key As Variant, legacy As Long, castleId As Long
    Dim mapping As Dictionary, RS As ADODB.Recordset
    Set mapping = New Dictionary
    Set RS = Query("SELECT legacy_trigger, castle_id FROM castle_legacy_trigger_map;")
    If RS Is Nothing Then Call Err.Raise(5, "ResolveLegacyCastleEntrances", "Apply ScriptsDB/20261007-01-migrate castle identities.sql before starting this server")
    Do While Not RS.EOF
        If CLng(RS!legacy_trigger) < 21 Or CLng(RS!legacy_trigger) > 40 Or GetCastleSlotById(CLng(RS!castle_id)) < 1 Then Call Err.Raise(5, "ResolveLegacyCastleEntrances", "Invalid historical castle mapping")
        Call mapping.Add(CLng(RS!legacy_trigger), CLng(RS!castle_id))
        Call RS.MoveNext()
    Loop
    Call RS.Close()
    For map = 1 To NumMaps
        With MapInfo(map)
            If Not .LegacyCastleEntrances Is Nothing Then
                For Each key In .LegacyCastleEntrances.Keys
                    legacy = .LegacyCastleEntrances.Item(key)
                    If Not mapping.Exists(legacy) Then Call Err.Raise(5, "ResolveLegacyCastleEntrances", "Unknown legacy castle trigger")
                    castleId = mapping.Item(legacy)
                    Call RegisterCastleEntrance(map, CInt(CLng(key) And &HFFFF&), CInt(CLng(key) \ 65536), castleId, True)
                Next key
                Set .LegacyCastleEntrances = Nothing
            End If
            If Not .CastleEntrances Is Nothing Then
                For Each key In .CastleEntrances.Keys
                    If GetCastleSlotById(.CastleEntrances.Item(key)) < 1 Then Call Err.Raise(5, "ResolveLegacyCastleEntrances", "Unknown authored castle ID")
                Next key
            End If
        End With
    Next map
    Exit Sub
ResolveLegacyCastleEntrances_Err:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Call CloseCastleRecordset(RS)
    Call Err.Raise(failureNumber, "ResolveLegacyCastleEntrances", failureDescription)
End Sub

Public Sub RestoreCastlePlacementsOnMap(ByVal map As Integer)
    If Not CastleModuleLoaded Then Exit Sub
    If Not MapInfo(map).LegacyCastleEntrances Is Nothing Then Call ResolveLegacyCastleEntrances()
    Dim slot As Integer
    For slot = LBound(CastleData) To UBound(CastleData)
        With CastleData(slot)
            If .is_active And .castle_coordinates.outside.map > 0 Then
                If .castle_coordinates.outside.map = map Then
                    Call CreateCastleInMap(map, .castle_coordinates.outside.x, .castle_coordinates.outside.y, slot)
                ElseIf .castle_coordinates.inside.map = map Then
                    Call RestoreCastleInteriorExits(slot)
                End If
            End If
        End With
    Next slot
End Sub

Private Sub RestoreCastleInteriorExits(ByVal slot As Integer)
    Dim map As Integer, x As Integer, y As Integer
    With CastleData(slot).castle_coordinates
        map = .inside.map: x = .inside.x: y = .inside.y + 1
        If Not InMapBounds(map, x, y) Or Not InMapBounds(map, x + 1, y) Then Call Err.Raise(5, , "Invalid castle interior exits")
        MapData(x, y, map).TileExit = .outside
        MapData(x, y, map).TileExit.x = .outside.x - 2
        MapData(x, y, map).TileExit.y = .outside.y + 1
        MapData(x + 1, y, map).TileExit = .outside
        MapData(x + 1, y, map).TileExit.x = .outside.x - 1
        MapData(x + 1, y, map).TileExit.y = .outside.y + 1
    End With
End Sub
