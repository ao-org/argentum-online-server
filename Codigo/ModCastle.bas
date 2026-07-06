Attribute VB_Name = "ModCastle"
Option Explicit

Private Type t_CastleCoordinates
    inside As t_WorldPos
    outside As t_WorldPos
End Type

Private Type t_CastleInfo
    id As Long
    trigger As Integer
    owner_account_id As Integer
    owner_char_id As Integer
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

Public CastleData() As t_CastleInfo


Private Const COUNT_ALL_CASTLES As String = "SELECT COUNT(*) FROM castle;"

Private Const UPDATE_EMPEROR_CASTLE As String = "UPDATE castle SET owner_account_id = ?, owner_character_id = ?, foundation_date = ?, is_active = ?, name = ? WHERE id = ?;"
Private Const UPDATE_OUTSIDE_CASTLE_LOCATION As String = "UPDATE castle_coordinates SET outside_map = ?, outside_x = ?, outside_y = ? WHERE id = ?;"

Private Const UPDATE_OR_INSERT_NEW_CASTLE_WHITELIST As String = "INSERT INTO"

Private Const SELECT_ALL_CASTLE_WHITELISTS As String = "Select * FROM castle_whitelist"
Private Const SELECT_ALL_CASTLES As String = "SELECT * FROM castle;"
Private Const SELECT_ALL_CASTLE_COORDINATES = "SELECT * FROM castle_coordinates;"

Private Const CastleXNegativeOffset As Integer = 8
Private Const CastleYNegativeOffset As Integer = 8
Private Const CastleXPositiveOffset As Integer = 6
Private Const CastleYPositiveOffset As Integer = 2

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

Public Sub LoadCastleModule()
    On Error GoTo LoadCastleModule_Err

    Call LoadCastleData
    Call LoadCastleCoordinates
    Call LoadCastleWhiteLists
    Dim i As Integer

    For i = LBound(CastleData) To UBound(CastleData)
        With CastleData(i)
            If .is_active Then
                Call CreateCastleInMap(.castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y, i)
            End If
        End With
    Next i

    Exit Sub
LoadCastleModule_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.LoadCastleModule", Erl)
End Sub

Public Sub LoadCastleWhiteLists()
    On Error GoTo LoadCastleWhitelists_Err
    Dim RS As ADODB.Recordset
    Set RS = Query(SELECT_ALL_CASTLE_WHITELISTS)
    If RS Is Nothing Or RS.RecordCount = 0 Then Exit Sub
    
    Do While Not RS.EOF
        Dim CastleSlot As Integer
        CastleSlot = GetCastleSlotById(RS!castle_id)
        Debug.Assert CastleSlot > 0
        Dim CharacterName As String
        CharacterName = RS!character_name
        Call AddUserNameToWhiteListByCastleSlot(CastleSlot, CharacterName)
        RS.MoveNext
    Loop

    Exit Sub
LoadCastleWhitelists_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.LoadCastleWhitelists", Erl)
End Sub


Public Function AddUserNameToWhiteListByCastleSlot(ByVal CastleSlot As Integer, ByVal CharacterName As String) As Boolean
    AddUserNameToWhiteListByCastleSlot = False
    With CastleData(CastleSlot)
        'duplicated entry somewhere
        Debug.Assert Not .castleWhiteList.Exists(CharacterName)
        If .castleWhiteList.Exists(CharacterName) Then
            Call LogInfoServidor("Duplicated username: " & CharacterName & " in whitelist for castle " & .name)
            Exit Function
        End If
        Call .castleWhiteList.Add(CharacterName, .trigger)
        Call LogInfoServidor("Username:" & CharacterName & " was added to the whitelist of castle: " & .name)
        AddUserNameToWhiteListByCastleSlot = True
    End With
End Function

Public Function GetCastleSlotById(ByVal id As Integer) As Integer
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
    Set RS = Query(COUNT_ALL_CASTLES)
    If RS Is Nothing Then
        Debug.Assert False
        Exit Sub
    End If
    ReDim CastleData(1 To RS.Fields(0).value)

    Dim i As Long
    i = 1
    Set RS = Query(SELECT_ALL_CASTLES)
    If RS Is Nothing Or RS.RecordCount = 0 Then Exit Sub
    If RS.RecordCount <> UBound(CastleData) Then
        Debug.Assert False
        Exit Sub
    End If

    Do While Not RS.EOF
    
        With CastleData(i)
            .id = (RS!id)
            .trigger = (RS!trigger)
            
            
            If Not IsNull(RS!name) Then
                .name = (RS!name)
            End If
            If Not IsNull(RS!owner_account_id) Then
                .owner_account_id = (RS!owner_account_id)
            End If
    
            If Not IsNull(RS!owner_character_id) Then
                .owner_char_id = (RS!owner_character_id)
                .owner_char_name = GetCharacterNameByUserId(RS!owner_character_id)
            End If
            
            If Not IsNull(RS!foundation_date) Then
                .foundation_date = (RS!foundation_date)
            End If
            
            Set .castleWhiteList = New Dictionary
            
            .spawner_obj_id = (RS!spawner_obj_id)
            .inside_key_obj_id = (RS!inside_key_obj_id)
            .is_active = (RS!is_active)
            If .owner_account_id <> 0 Then
                Call .castleWhiteList.Add(.owner_account_id, .trigger) 'add castle owner to the whitelist
            End If
            i = i + 1
            RS.MoveNext
        
        End With
        
    Loop
    Exit Sub
LoadCastleData_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.LoadCastleData", Erl)
End Sub


Public Sub LoadCastleCoordinates()
    On Error GoTo LoadCastleCoordinates_Err
    Dim i As Integer
    Dim RS As ADODB.Recordset
    Set RS = Query(SELECT_ALL_CASTLE_COORDINATES)
    If RS Is Nothing Or RS.RecordCount = 0 Then Exit Sub
    i = 1
    Do While Not RS.EOF
        If i <> (RS!castle_id) Then
            Debug.Assert False
        End If
        CastleData(i).castle_coordinates.outside.map = (RS!outside_map)
        CastleData(i).castle_coordinates.outside.x = (RS!outside_x)
        CastleData(i).castle_coordinates.outside.y = (RS!outside_y)
        CastleData(i).castle_coordinates.inside.map = (RS!inside_map)
        CastleData(i).castle_coordinates.inside.x = (RS!inside_x)
        CastleData(i).castle_coordinates.inside.y = (RS!inside_y)
        i = i + 1
        RS.MoveNext
    Loop
Exit Sub
LoadCastleCoordinates_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.LoadCastleCoordinates", Erl)
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
            Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, FONTTYPE_INFOBOLD)
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
        Call LogError("Usuario " & UserList(UserIndex).Name & "Interactuando con un mapa fuera de su rango, revisar")
        Exit Function
    End If

    If Not IsValidMapIndex(UserTargetMap) Then
        Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, FONTTYPE_INFOBOLD)
        Exit Function
    End If

    If Not IsCastleFootprintInMapBounds(UserTargetMap, UserTargetX, UserTargetY) Then
        Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, FONTTYPE_INFOBOLD)
        Exit Function
    End If

    If MapData(UserTargetMap, UserTargetX, UserTargetY).trigger <> e_Trigger.CASTLE_FOUNDATION_POSITION Then
        Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, FONTTYPE_INFOBOLD)
        Exit Function
    End If

    If MapData(UserTargetMap, UserTargetX, UserTargetY).ObjInfo.ObjIndex = CASTLE_MOCKUP_OBJ_INDEX Then
        Call WriteLocaleMsg(UserIndex, MSG_CANT_FOUND_CASTLE_ON_TOP_OF_ANOTHER, FONTTYPE_INFOBOLD)
        Exit Function
    End If

    If MapData(UserTargetMap, UserTargetX, UserTargetY).ObjInfo.ObjIndex <> CASTLE_SIGN_POST_OBJ_INDEX Then
        Call WriteLocaleMsg(UserIndex, MSG_CANOT_FOUND_WITHOUT_SIGN, FONTTYPE_INFOBOLD)
        Exit Function
    End If

    Dim i As Integer
    Dim j As Integer
    For i = CastleTopLeftCorner.x To CastleBottomRightCorner.x
        For j = CastleTopLeftCorner.y To CastleBottomRightCorner.y

            If Not InMapBounds(UserTargetMap, i, j) Then
                Call WriteLocaleMsg(UserIndex, MSG_INVALID_CASTLE_POSITION, FONTTYPE_INFOBOLD)
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

    With CastleData(CastleIndex)

        'if not during server start...(player clicking the board)
         If UserIndex > 0 Then
            .castle_coordinates.outside.map = map
            .castle_coordinates.outside.x = x
            .castle_coordinates.outside.y = y
            .foundation_date = DateTime.Now
            .is_active = True
            .owner_account_id = UserList(UserIndex).AccountID
            .owner_char_id = UserList(UserIndex).Id
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

            MapData(map, i, j).Blocked = 0
            MapData(map, i, j).trigger = e_Trigger.nada

            If MapData(map, i, j).ObjInfo.ObjIndex > 0 Then
                Call EraseObj(MapData(map, i, j).ObjInfo.Amount, map, i, j)
            End If

            If MapData(map, i, j).NpcIndex > 0 Then
                Call QuitarNPC(MapData(map, i, j).NpcIndex, eAiResetNpc)
            End If

            Next j
        Next i


        'first layer from the bottom
        MapData(map, x - 3, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 4, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 5, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 6, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 1, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 2, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 3, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x, y).trigger = e_Trigger.CASTLE_FOUNDATION_POSITION
        MapData(map, x - 1, y).trigger = .trigger
        MapData(map, x - 2, y).trigger = .trigger
        MapData(map, x - 1, y + 1).trigger = .trigger
        MapData(map, x - 2, y + 1).trigger = .trigger


        'second layer form the bottom
        MapData(map, x, y - 1).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 3, y - 1).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 4, y - 1).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 5, y - 1).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 6, y - 1).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 1, y - 1).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 2, y - 1).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 3, y - 1).Blocked = e_Block.ALL_SIDES

        MapData(map, x - 1, y - 1).TileExit.map = .castle_coordinates.inside.map
        MapData(map, x - 1, y - 1).TileExit.x = .castle_coordinates.inside.x
        MapData(map, x - 1, y - 1).TileExit.y = .castle_coordinates.inside.y

        MapData(map, x - 2, y - 1).TileExit.map = .castle_coordinates.inside.map
        MapData(map, x - 2, y - 1).TileExit.x = .castle_coordinates.inside.x
        MapData(map, x - 2, y - 1).TileExit.y = .castle_coordinates.inside.y

        'third layer form the bottom
        MapData(map, x, y - 2).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 1, y - 2).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 2, y - 2).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 3, y - 2).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 4, y - 2).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 5, y - 2).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 6, y - 2).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 1, y - 2).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 2, y - 2).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 3, y - 2).Blocked = e_Block.ALL_SIDES

         'fourth layer form the bottom
        MapData(map, x, y - 3).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 1, y - 3).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 2, y - 3).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 3, y - 3).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 4, y - 3).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 5, y - 3).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 6, y - 3).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 1, y - 3).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 2, y - 3).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 3, y - 3).Blocked = e_Block.ALL_SIDES

         'fifth layer form the bottom
        MapData(map, x, y - 4).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 1, y - 4).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 2, y - 4).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 3, y - 4).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 4, y - 4).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 5, y - 4).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 6, y - 4).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 1, y - 4).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 2, y - 4).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 3, y - 4).Blocked = e_Block.ALL_SIDES

         'sixth layer form the bottom
        MapData(map, x, y - 5).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 1, y - 5).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 2, y - 5).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 3, y - 5).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 4, y - 5).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 5, y - 5).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 6, y - 5).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 1, y - 5).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 2, y - 5).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 3, y - 5).Blocked = e_Block.ALL_SIDES

         'seventh layer form the bottom
        MapData(map, x, y - 6).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 1, y - 6).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 2, y - 6).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 3, y - 6).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 4, y - 6).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 5, y - 6).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 6, y - 6).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 1, y - 6).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 2, y - 6).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 3, y - 6).Blocked = e_Block.ALL_SIDES

         'eighth layer form the bottom
        MapData(map, x, y - 7).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 1, y - 7).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 2, y - 7).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 3, y - 7).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 4, y - 7).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 5, y - 7).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 6, y - 7).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 1, y - 7).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 2, y - 7).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 3, y - 7).Blocked = e_Block.ALL_SIDES

        'create castle inside tile exits to the outside part
        #If DEBUGGING = 0 Then
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

        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x, .castle_coordinates.inside.y + 1).TileExit.map = .castle_coordinates.outside.map
        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x, .castle_coordinates.inside.y + 1).TileExit.x = .castle_coordinates.outside.x - 2
        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x, .castle_coordinates.inside.y + 1).TileExit.y = .castle_coordinates.outside.y + 1

        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1).TileExit.map = .castle_coordinates.outside.map
        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1).TileExit.x = .castle_coordinates.outside.x - 1
        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1).TileExit.y = .castle_coordinates.outside.y + 1
        #End If

        'erase castle sign
        If MapData(map, x, y).ObjInfo.Amount > 0 Then
            Call EraseObj(MapData(map, x, y).ObjInfo.Amount, map, x, y)
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

    If MapData(map, x, y).ObjInfo.Amount > 0 Then
        Call EraseObj(MapData(map, x, y).ObjInfo.Amount, map, x, y)
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

        MapData(map, i, j).Blocked = 0
        MapData(map, i, j).trigger = e_Trigger.nada

        If MapData(map, i, j).ObjInfo.ObjIndex > 0 Then
            Call EraseObj(MapData(map, i, j).ObjInfo.Amount, map, i, j)
        End If

        If MapData(map, i, j).NpcIndex > 0 Then
            Call QuitarNPC(MapData(map, i, j).NpcIndex, eAiResetNpc)
        End If

        Next j
    Next i

    MapData(map, x - 1, y - 1).TileExit.map = 0
    MapData(map, x - 1, y - 1).TileExit.x = 0
    MapData(map, x - 1, y - 1).TileExit.y = 0

    MapData(map, x - 2, y - 1).TileExit.map = 0
    MapData(map, x - 2, y - 1).TileExit.x = 0
    MapData(map, x - 2, y - 1).TileExit.y = 0

     'restore castle foundation trigger
    MapData(map, x, y).trigger = e_Trigger.CASTLE_FOUNDATION_POSITION

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

        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x, .castle_coordinates.inside.y + 1).TileExit.map = 0
        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x, .castle_coordinates.inside.y + 1).TileExit.x = 0
        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x, .castle_coordinates.inside.y + 1).TileExit.y = 0

        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1).TileExit.map = 0
        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1).TileExit.x = 0
        MapData(.castle_coordinates.inside.map, .castle_coordinates.inside.x + 1, .castle_coordinates.inside.y + 1).TileExit.y = 0
    End With

    'create castle sign post
    Dim CastleSignObj As t_Obj
    CastleSignObj.Amount = 1
    CastleSignObj.ObjIndex = CASTLE_SIGN_POST_OBJ_INDEX
    Call MakeObj(CastleSignObj, map, x, y)
End Sub

Public Function IsEmperorCastleCreated(ByVal UserIndex As Integer, Optional ByRef CastleIndex As Integer = -1) As Boolean
    IsEmperorCastleCreated = False
    Dim i As Integer
    For i = 1 To UBound(CastleData)
        With CastleData(i)
            If .owner_account_id = UserList(UserIndex).AccountID Then
                If Not IsCastleFootprintInMapBounds(.castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y) Then
                    Call LogInfoServidor("IsEmperorCastleCreated outside map bounds. map=" & CStr(.castle_coordinates.outside.map) & _
                        " x=" & CStr(.castle_coordinates.outside.x) & _
                        " y=" & CStr(.castle_coordinates.outside.y) & _
                        " CastleIndex=" & CStr(i))
                    Exit Function
                End If

                If (MapData(.castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y).ObjInfo.ObjIndex = CASTLE_MOCKUP_OBJ_INDEX) Then
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
    If Acumulator >= 7 Then
        HasCastleRelocationCooldownPassed = True
    End If
End Function

Public Sub CreateNewEmperorCastle(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo CreateEmperorCastle_Err
    Dim RS As ADODB.Recordset
    With UserList(UserIndex)

        If IsEmperorCastleCreated(UserIndex) Then
            If Not HasCastleRelocationCooldownPassed(ObjData(ObjIndex).AssignedCastleIndex) Then
                Call WriteLocaleMsg(UserIndex, MSG_CASTLE_RELOCATION_ON_COOLDOWN, FONTTYPE_INFOBOLD)
                Exit Sub
            End If
            With CastleData(ObjData(ObjIndex).AssignedCastleIndex)
                Call DestroyCastleInMap(.castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y, ObjData(ObjIndex).AssignedCastleIndex)
                Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_BROADCAST_CASTLE_DESTROYED, ObjData(ObjIndex).AssignedCastleIndex & "¬" & GetUserDisplayName(UserIndex), e_FontTypeNames.FONTTYPE_GUILD))
            End With
        End If

        Call CreateCastleInMap(.flags.TargetMap, .flags.TargetX, .flags.TargetY, ObjData(ObjIndex).AssignedCastleIndex, UserIndex)
        
        Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_BROADCAST_CASTLE_LOCATION, .name & "¬" & GetUserDisplayName(UserIndex) & "¬" & .flags.TargetMap & "¬" & .flags.TargetX & "¬" & .flags.TargetY, e_FontTypeNames.FONTTYPE_GUILD))
        Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(e_SoundEffects.OldClanHorn, 50, 50))
        Call modSendData.SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(e_SoundEffects.NewCastleRPGVoice, 50, 50))
    End With
    Exit Sub
CreateEmperorCastle_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.CreateEmperorCastle", Erl)
End Sub

Function CheckCastleEntryWhiteList(ByVal UserIndex As Integer, ByVal trigger As Integer) As Boolean
   CheckCastleEntryWhiteList = False

    Dim CastleIndex As Integer
    If Not IsEmperorCastleCreated(UserIndex, CastleIndex) Then
        Exit Function
    End If

    With CastleData(CastleIndex)

       'exception for castle owner
       If UserList(UserIndex).AccountID = .owner_account_id Then
            CheckCastleEntryWhiteList = True
            Exit Function
       End If
    
       If Not .castleWhiteList.Exists(UserList(UserIndex).name) Then
            Exit Function
       End If
    
       If .castleWhiteList.Item(UserList(UserIndex).name) = trigger Then
            CheckCastleEntryWhiteList = True
       End If
   
   End With
   
End Function

Public Sub AddToCastleEntryWhiteList(ByVal UserIndex As Integer, ByVal CharacterName As String)
    With UserList(UserIndex)
        If .Stats.tipoUsuario < e_TipoUsuario.tNoble Then
            Call WriteLocaleMsg(UserIndex, MSG_AT_LEAST_NOBLE_TO_FOUND_CASTLE, FONTTYPE_INFOBOLD)
            Call LogInfoServidor("User with low patreon status trying to set a whitelist for a castle, name: " & .name)
            Exit Sub
        End If
        
        Dim CastleIndex As Integer
        If Not IsEmperorCastleCreated(UserIndex, CastleIndex) Then
            Call LogError("Couldn't find the castle for user: " & .name)
            Debug.Assert False
            Exit Sub
        End If
    
        If CastleIndex <= -1 Then
            Call LogError("Couldn't find the castle for user: " & .name)
            Debug.Assert False
        End If
        
        If AddUserNameToWhiteListByCastleSlot(CastleIndex, CharacterName) Then
            CastleData(CastleIndex).dirtyWhiteList = True
        End If
        
    End With
End Sub

Public Sub SaveCastleWhiteListToDb()
    Dim i As Integer
    Dim RS As ADODB.Recordset
    For i = LBound(CastleData) To UBound(CastleData)
        With (CastleData(i))
            If .dirtyWhiteList Then
                
            End If
        End With
    Next i
End Sub

Public Sub SaveCastleDataToDb()
    Dim i As Integer
    Dim RS As ADODB.Recordset
    For i = LBound(CastleData) To UBound(CastleData)
        With (CastleData(i))
            
            If .dirtyCastleData Then
                'update castle data in db
                Set RS = Query(UPDATE_EMPEROR_CASTLE, .owner_account_id, .id, DateToSQLite(.foundation_date), 1, .name, i)
                'update castle coordinates in db
                Set RS = Query(UPDATE_OUTSIDE_CASTLE_LOCATION, .castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y, i)
                Call LogInfoServidor("Persisted castle number: " & i & " name: " & .name)
            End If
            
        End With
    Next i
End Sub

Public Sub SaveCastlesToDb()
    Call SaveCastleDataToDb
End Sub

Public Sub SendCastleInfo(ByVal UserIndex As Integer, ByVal CastleSlot As Integer)
    If CastleSlot <= 0 Then
        Debug.Assert False
    End If

    With CastleData(CastleSlot)
        Call WriteLocaleMsg(UserIndex, MSG_CASTLE_FOUNDER_INFO, FONTTYPE_INFOBOLD, .name & "¬" & .foundation_date & "¬" & .owner_char_name)
    End With
End Sub






