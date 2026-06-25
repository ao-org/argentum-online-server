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
Private Const ADD_NEW_EMPEROR_CASTLE As String = "INSERT INTO castle (trigger, owner_account_id, foundation_date, is_active) VALUES (?,?,?,?);"

Public Function IsEmperorCastleCreated(ByVal UserIndex As Integer) As Boolean
    IsEmperorCastleCreated = False
    Dim RS As ADODB.Recordset
    Set RS = Query(CHECK_EMPEROR_CASTLE)
    If RS Is Nothing Or RS.RecordCount = 0 Then Exit Function
    IsEmperorCastleCreated = True
End Function

Public Sub CreateEmperorCastle(ByVal UserIndex As Integer)
    If IsEmperorCastleCreated(UserIndex) Then Exit Sub
    Dim RS As ADODB.Recordset
    
    

End Sub





Public Sub LoadCastleModule()
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
        CastleData(i).white_list = (RS!white_list)
        
        str = Split(CastleData(i).white_list, ";")
        If UBound(str) > 0 Then
            For y = 0 To UBound(str)
                Call CastleWhiteList.Add(str(y), CastleData(i).trigger)
            Next y
        End If
    Next i
    
    
    
End Sub

Function CheckCastleEntryWhiteList(ByVal UserIndex As Integer, ByVal trigger As Integer) As Boolean
   CheckCastleEntryWhiteList = False
   If Not CastleWhiteList.Exists(UserList(UserIndex).Name) Then
        Exit Function
   End If
   Dim valid_trigger As Integer
   valid_trigger = CastleWhiteList(UserList(UserIndex).Name)
   If valid_trigger = trigger Then
        CheckCastleEntryWhiteList = True
   End If
End Function
