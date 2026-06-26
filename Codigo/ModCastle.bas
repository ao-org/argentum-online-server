Attribute VB_Name = "ModCastle"
Option Explicit

Private Type t_CastleCoordinates
    inside As t_WorldPos
    outside As t_WorldPos
End Type

Private Type t_CastleInfo
    trigger As Integer
    owner_account_id As Integer
    owner_char_id As Integer
    spawner_obj_id As Integer
    portal_obj_id As Integer
    foundation_date As Date
    is_active As Boolean
    castle_coordinates As t_CastleCoordinates
End Type
Public CastleData() As t_CastleInfo
Public CastleWhiteList As Dictionary

Private Const COUNT_ALL_CASTLES As String = "SELECT COUNT(*) FROM castle;"
Private Const CHECK_EMPEROR_CASTLE As String = "SELECT * FROM castle WHERE owner_account_id = ?;"
Private Const SELECT_ALL_CASTLES As String = "SELECT * FROM castle;"
Private Const ADD_NEW_EMPEROR_CASTLE As String = "INSERT INTO castle (owner_account_id,owner_character_id, foundation_date, is_active,outside_map,outside_x,outside_y) VALUES (?,?,?,?,?,?,?);"
Private Const SELECT_ALL_CASTLE_WHITELISTS As String = "Select * FROM castle_whitelist"
Private Const CASTLE_OBJ = 6382

Public Function IsEmperorCastleCreated(ByVal UserIndex As Integer) As Boolean
    IsEmperorCastleCreated = False
    Dim RS As ADODB.Recordset
    Set RS = Query(CHECK_EMPEROR_CASTLE, UserList(UserIndex).AccountID)
    If RS Is Nothing Then
        Exit Function
    End If
    
    If Not IsNull(RS!foundation_date) Then
        IsEmperorCastleCreated = True
    End If
        
End Function

Public Sub CreateNewEmperorCastle(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo CreateEmperorCastle_Err
    If IsEmperorCastleCreated(UserIndex) Then Exit Sub
    Dim RS As ADODB.Recordset
    With UserList(UserIndex)
        Set RS = Query(ADD_NEW_EMPEROR_CASTLE, .AccountID, .Name, DateToSQLite(DateTime.Now), 1, .flags.TargetMap, .flags.TargetX, .flags.TargetY)
        Call CreateCastleInMap(.flags.TargetMap, .flags.TargetX, .flags.TargetY, ObjData(ObjIndex).AssignedCastleIndex)
    End With
    Exit Sub
CreateEmperorCastle_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.CreateEmperorCastle", Erl)
End Sub

Public Sub LoadCastleModule()
    On Error GoTo LoadCastleModule_Err
    Set CastleWhiteList = New Dictionary
    Call LoadCastleData
    Call LoadCastleWhitelists
    Dim i As Integer
    
    For i = LBound(CastleData) To UBound(CastleData)
        With CastleData(i)
            If InMapBounds(.castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y) And MapaValido(.castle_coordinates.outside.map) Then
                Call CreateCastleInMap(.castle_coordinates.outside.map, .castle_coordinates.outside.x, .castle_coordinates.outside.y, i)
            End If
        End With
    Next i
    
    Exit Sub
LoadCastleModule_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.LoadCastleModule", Erl)
End Sub

Function CheckCastleEntryWhiteList(ByVal UserIndex As Integer, ByVal trigger As Integer) As Boolean
   CheckCastleEntryWhiteList = False
   
   'exception for castle owner
   If CastleWhiteList.Item(UserList(UserIndex).AccountID) = trigger Then
        CheckCastleEntryWhiteList = True
        Exit Function
   End If
   
   If Not CastleWhiteList.Exists(UserList(UserIndex).Name) Then
        Exit Function
   End If
   
   If CastleWhiteList.Item(UserList(UserIndex).Name) = trigger Then
        CheckCastleEntryWhiteList = True
   End If
End Function


Public Sub LoadCastleWhitelists()
    On Error GoTo LoadCastleWhitelists_Err
    Dim RS As ADODB.Recordset
    Set RS = Query(SELECT_ALL_CASTLE_WHITELISTS)
    If RS Is Nothing Or RS.RecordCount = 0 Then Exit Sub

    Do While Not RS.EOF
        Call CastleWhiteList.Add((RS!character_name), CastleData(RS!Castle_Id).trigger)
        RS.MoveNext
    Loop
    Exit Sub
LoadCastleWhitelists_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.LoadCastleWhitelists", Erl)
End Sub

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
    
            CastleData(i).trigger = (RS!trigger)
            
            If Not IsNull(RS!owner_account_id) Then
                CastleData(i).owner_account_id = (RS!owner_account_id)
            End If
            
            If Not IsNull(RS!owner_character_id) Then
                CastleData(i).owner_char_id = (RS!owner_character_id)
            End If
            
            CastleData(i).spawner_obj_id = (RS!spawner_obj_id)
            CastleData(i).portal_obj_id = (RS!portal_obj_id)
            
            CastleData(i).castle_coordinates.outside.map = (RS!outside_map)
            CastleData(i).castle_coordinates.outside.x = (RS!outside_x)
            CastleData(i).castle_coordinates.outside.y = (RS!outside_y)
            
            CastleData(i).castle_coordinates.inside.map = (RS!inside_map)
            CastleData(i).castle_coordinates.inside.x = (RS!inside_x)
            CastleData(i).castle_coordinates.inside.y = (RS!inside_y)
            
            
            CastleData(i).is_active = (RS!is_active)
            
            
            If CastleData(i).owner_account_id <> 0 Then
                Call CastleWhiteList.Add(CastleData(i).owner_account_id, CastleData(i).trigger) 'add castle owner to the whitelist
            End If
        
            i = i + 1
            RS.MoveNext
    Loop
    Exit Sub
LoadCastleData_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.LoadCastleData", Erl)
End Sub


Public Function IsValidCastlePosition(ByVal UserIndex As Integer) As Boolean
    IsValidCastlePosition = False

    Dim CastleTopLeftCorner As t_WorldPos
    Dim CastleBottomRightCorner As t_WorldPos
    
    Dim UserTargetX As Integer
    Dim UserTargetY As Integer
    Dim UserTargetMap As Integer
    
    With UserList(UserIndex)
        CastleTopLeftCorner.x = .flags.TargetX - 6
        CastleTopLeftCorner.y = .flags.TargetY - 7
        CastleTopLeftCorner.map = .flags.TargetMap
        
        CastleBottomRightCorner.x = .flags.TargetX + 3
        CastleBottomRightCorner.y = .flags.TargetY
        CastleBottomRightCorner.map = .flags.TargetMap
        
        UserTargetX = .flags.TargetX
        UserTargetY = .flags.TargetY
        UserTargetMap = .flags.TargetMap
    End With
    
    
    If MapData(UserTargetMap, UserTargetX, UserTargetY).trigger <> e_Trigger.CASTLE_FOUNDATION_POSITION Then
        'TODO not valid clastle foundation position errormsg
        Exit Function
    End If
    
    If UserList(UserIndex).pos.map <> UserTargetMap Then
        'TODO user not in map errormsg
        Exit Function
    End If
    
    If Not IsValidMapIndex(UserTargetMap) Then
        'TODO map is not valid errormsg
        Exit Function
    End If
    
    Dim i As Integer
    Dim j As Integer
    For i = CastleTopLeftCorner.x To CastleBottomRightCorner.x
        For j = CastleTopLeftCorner.y To CastleBottomRightCorner.y
        
            If Not InMapBounds(UserTargetMap, i, j) Then
                'TODO castle wont be in map bounds errormsg
                Exit Function
            End If
            
        Next j
    Next i
    IsValidCastlePosition = True
End Function


Public Sub CreateCastleInMap(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal CastleIndex As Integer)
    Dim CastleObj As t_Obj
    CastleObj.Amount = 1
    CastleObj.ObjIndex = CASTLE_OBJ
    Call MakeObj(CastleObj, map, x, y)
    
    
    With CastleData(CastleIndex)
        'first layer from the bottom
        MapData(map, x - 3, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 4, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 5, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x - 6, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 1, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 2, y).Blocked = e_Block.ALL_SIDES
        MapData(map, x + 3, y).Blocked = e_Block.ALL_SIDES
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
    
    End With
    
End Sub

