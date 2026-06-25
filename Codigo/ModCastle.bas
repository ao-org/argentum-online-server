Attribute VB_Name = "ModCastle"
Option Explicit

Private Type CastleInfo
    trigger As Integer
    owner_account_id As Integer
    white_list As String ' ";" separated string of names
    is_active As Boolean
End Type


Public CastleData() As CastleInfo
Public CastleWhiteList() As Dictionary

Private Const SELECT_ALL_CASTLE As String = "SELECT * FROM castle WHERE is_active = 1;"

Public Sub LoadCastleModule()
    Dim RS As ADODB.Recordset
    Set RS = Query(SELECT_ALL_CASTLE)
    If RS Is Nothing Then Exit Sub
    Dim i As Integer
    Dim y As Integer
    ReDim CastleData(1 To RS.RecordCount)
    ReDim CastleWhiteList(1 To RS.RecordCount)
    Dim str() As String
    For i = LBound(CastleData) To UBound(CastleData)
        CastleData(i).is_active = (RS!is_active)
        CastleData(i).owner_acc_id = (RS!owner_account_id)
        CastleData(i).trigger = (RS!trigger)
        CastleData(i).white_list = (RS!white_list)
        
        str = Split(CastleData(i).white_list, ";")
        If UBound(str) > 0 Then
            For y = 0 To UBound(str)
                Call CastleWhiteList(i).Add(str(y), 1)
            Next y
        End If
        
        
    Next i
    
    
    
End Sub

Function CheckCastleEntryWhiteList(ByVal UserIndex As Integer) As Boolean


End Function
