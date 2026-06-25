Attribute VB_Name = "ModCastle"
Option Explicit

Private Type t_CastleInfo
    trigger As Integer
    owner_account_id As Integer
    owner_char_id As Integer
    obj_id As Integer
    foundation_date As Date
    is_active As Boolean
    map As Integer
    x As Integer
    y As Integer
End Type

Private Type t_WhiteListEntry
    character_name As String
    castle_id As Integer
End Type


Public CastleData() As t_CastleInfo
Public CastleWhiteList As Dictionary

Private Const COUNT_ALL_ACTIVE_CASTLES As String = "SELECT COUNT(*) FROM castle;"
Private Const CHECK_EMPEROR_CASTLE As String = "Select 1 FROM castle WHERE owner_account_id = ?;"
Private Const SELECT_ALL_CASTLES As String = "SELECT * FROM castle WHERE id = ?;"
Private Const ADD_NEW_EMPEROR_CASTLE As String = "INSERT INTO castle (owner_account_id,owner_character_id, foundation_date, is_active) VALUES (?,?,?,?);"
Private Const SELECT_ALL_CASTLE_WHITELISTS As String = "Select * FROM castle_whitelist"

Public Function IsEmperorCastleCreated(ByVal UserIndex As Integer) As Boolean

    IsEmperorCastleCreated = False
    Dim RS As ADODB.Recordset
    Set RS = Query(CHECK_EMPEROR_CASTLE)
    If RS Is Nothing Or RS.RecordCount = 0 Then Exit Function
    IsEmperorCastleCreated = True
    
End Function

Public Sub CreateEmperorCastle(ByVal UserIndex As Integer)
    On Error GoTo CreateEmperorCastle_Err
    If IsEmperorCastleCreated(UserIndex) Then Exit Sub
    Dim RS As ADODB.Recordset
    Set RS = Query(ADD_NEW_EMPEROR_CASTLE, UserList(UserIndex).AccountID, UserList(UserIndex).Name, SQLiteToDate(DateTime.Now), 1)
    Exit Sub
    
CreateEmperorCastle_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.CreateEmperorCastle", Erl)
End Sub

Public Sub LoadCastleModule()
    On Error GoTo LoadCastleModule_Err
    Set CastleWhiteList = New Dictionary
    Dim RS As ADODB.Recordset
    Set RS = Query(COUNT_ALL_ACTIVE_CASTLES)
    If RS Is Nothing Or RS.RecordCount = 0 Then
        Debug.Assert False
        Exit Sub
    End If
    ReDim CastleData(1 To RS.Fields(0).value)
    

    
    Call LoadCastleWhitelists
    
    
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
    Dim i As Long
    
    For i = 0 To RS.RecordCount
        
        
    Next i
    


LoadCastleWhitelists_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.LoadCastleWhitelists", Erl)
End Sub

Public Sub LoadCastleData()
    On Error GoTo LoadCastleData_Err
    Dim RS As ADODB.Recordset
    
    
    Dim i As Long
    Dim y As Long
    For i = LBound(CastleData) To UBound(CastleData)
        Set RS = Query(SELECT_ALL_CASTLES, i)
        If RS Is Nothing Or RS.RecordCount = 0 Then Exit Sub
            CastleData(i).trigger = (RS!trigger)
            CastleData(i).owner_account_id = (RS!owner_account_id)
            CastleData(i).owner_char_id = (RS!owner_char_id)
            CastleData(i).is_active = (RS!is_active)
            CastleData(i).obj_id = (RS!obj_id)
            CastleData(i).map = (RS!map)
            CastleData(i).x = (RS!x)
            CastleData(i).y = (RS!y)
            Call CastleWhiteList.Add(CastleData(i).owner_account_id, CastleData(i).trigger) 'add castle owner to the whitelist
    Next i
    
LoadCastleData_Err:
Call TraceError(Err.Number, Err.Description, "ModCastle.LoadCastleWhitelists", Erl)
End Sub
