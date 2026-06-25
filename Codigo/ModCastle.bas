Attribute VB_Name = "ModCastle"
Option Explicit

Private Type t_CastleInfo
    trigger As Integer
    owner_account_id As Integer
    white_list As String ' ";" separated string of names
    is_active As Boolean
End Type


Public CastleData() As t_CastleInfo
Public CastleWhiteList As Dictionary

Private Const SELECT_ALL_CASTLE As String = "SELECT * FROM castle WHERE is_active = 1;"
Private Const CHECK_EMPEROR_CASTLE As String = "Select 1 FROM castle WHERE owner_account_id = ?;"
Private Const ADD_NEW_EMPEROR_CASTLE As String = "INSERT INTO castle (owner_account_id,owner_character_id, foundation_date, is_active) VALUES (?,?,?,?);"

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
    Set RS = Query(SELECT_ALL_CASTLE)
    If RS Is Nothing Or RS.RecordCount = 0 Then Exit Sub
    Dim i As Integer
    Dim y As Integer
    ReDim CastleData(1 To RS.RecordCount)
    Dim str() As String
    For i = LBound(CastleData) To UBound(CastleData)
        CastleData(i).is_active = (RS!is_active)
        CastleData(i).owner_account_id = (RS!owner_account_id)
        CastleData(i).trigger = (RS!trigger)
        
        If Not RS!white_list = Null Then
            CastleData(i).white_list = (RS!white_list)
            str = Split(CastleData(i).white_list, ";")
            If UBound(str) > 0 Then
                For y = 0 To UBound(str)
                    Call CastleWhiteList.Add(str(y), CastleData(i).trigger) 'add each memeber in the list to the whitelist
                Next y
            End If
        End If
        
        Call CastleWhiteList.Add(CastleData(i).owner_account_id, CastleData(i).trigger) 'add castle owner to the whitelist
    
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
