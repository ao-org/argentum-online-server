Attribute VB_Name = "modHousingDatabase"
Option Explicit

' Ordinary housing only. One authoritative world process; no character-save work.
Public Enum e_HousingCode
    HousingOkay = 0
    HousingBusy = 1
    HousingUnavailable = 2
    HousingIneligible = 3
    HousingStale = 4
    HousingOccupied = 5
    HousingMinimumPeriod = 6
    HousingInconsistent = 7
    HousingInvalid = 8
    HousingUnresolved = 9
End Enum
Public Enum e_HousingOperation
    HousingCatalogue = 0
    HousingClaim = 1
    HousingRelease = 2
    HousingExchange = 3
End Enum
Public Type t_HousingRequest
    RequestId As Long
    Operation As Byte
    ExpectedKey As Integer
    TargetKey As Integer
    Cursor As Integer
    Search As String
End Type
Public Type t_HousingHouse
    Key As Integer
    Description As String
End Type
Public Type t_HousingReply
    RequestId As Long
    Operation As Byte
    Code As e_HousingCode
    Current As t_HousingHouse
    AssignedUTC As String
    EligibleUTC As String
    ServerUTC As String
    CanAcquire As Boolean
    CanRelease As Boolean
    NextCursor As Integer
    Count As Byte
    Houses(1 To 4) As t_HousingHouse
End Type
Private housingConnection As ADODB.Connection
Private connectionSettings As String
Private Const ORDINARY As String = " NOT EXISTS (SELECT 1 FROM castle c WHERE c.inside_key_obj_id=h.key_obj) "
#If UNIT_TEST = 1 Then
Public HousingTestLastError As String
Public HousingTestFailAfterRelease As Boolean
Public HousingTestFailBeforeCommit As Boolean
Public HousingTestLoseCommitReply As Boolean
#End If

Public Function HousingConnectionString() As String
    If Len(Database_Source) <> 0 Then
        HousingConnectionString = "DATA SOURCE=" & Database_Source & ";"
    Else
        HousingConnectionString = "DRIVER={SQLite3 ODBC Driver};DATABASE=" & App.Path & "/" & DatabaseFileName
    End If
End Function

Public Function HousingDbOpen(ByVal connectionString As String) As Boolean
    On Error GoTo Failed
    Call HousingDbClose()
    connectionSettings = connectionString
    Set housingConnection = New ADODB.Connection
    ' Cache bounded result rows: SQLite ODBC forward-only text fields cannot be read twice.
    housingConnection.CursorLocation = adUseClient
    housingConnection.ConnectionString = connectionString & ";Timeout=1;"
    housingConnection.CommandTimeout = 1
    Call housingConnection.Open()
    ' The installed x86 SQLite ODBC busy handler has its own retry quantum.
    ' Timeout=0 means effectively unbounded waiting on this driver, not fail-fast.
    Call housingConnection.Execute("PRAGMA busy_timeout=1")
    Dim rs As ADODB.Recordset
    Set rs = HousingQuery("SELECT h.key_obj,h.account_id,h.description,h.assigned_at FROM house_key h WHERE 0=1 AND " & ORDINARY)
    Call rs.Close()
    HousingDbOpen = True
    Exit Function
Failed:
    Dim failure As String
    failure = Err.Description
    Call LogDatabaseError("Housing connection unavailable: " & CStr(Err.Number))
    Call HousingDbClose()
#If UNIT_TEST = 1 Then
    Err.Raise vbObjectError + 732, "Housing database", failure
#End If
End Function

Public Sub HousingDbClose()
    On Error Resume Next
    If Not housingConnection Is Nothing Then
        If housingConnection.State <> adStateClosed Then Call housingConnection.Close()
    End If
    Set housingConnection = Nothing
End Sub

Public Function HousingQuery(ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    Dim cmd As New ADODB.Command, arg As Variant, parameterIndex As Long
    Set cmd.ActiveConnection = housingConnection
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 1
    cmd.CommandText = sql
    cmd.Prepared = True
    For Each arg In args
        parameterIndex = parameterIndex + 1
        If VarType(arg) = vbString Then
            Call cmd.Parameters.Append(cmd.CreateParameter("p" & CStr(parameterIndex), adVarWChar, adParamInput, Len(arg) + 1, CStr(arg)))
        Else
            Call cmd.Parameters.Append(cmd.CreateParameter("p" & CStr(parameterIndex), adInteger, adParamInput, 4, CLng(arg)))
        End If
    Next
    Set HousingQuery = cmd.Execute()
End Function

Private Sub ExecuteOne(ByVal sql As String, ByVal first As Long, ByVal second As Long)
    Dim rs As ADODB.Recordset
    Set rs = HousingQuery(sql, first, second)
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then Call rs.Close()
    End If
    Set rs = HousingQuery("SELECT changes()")
    If CLng(rs.Fields(0).Value) <> 1 Then Err.Raise vbObjectError + 720, "Housing", "Unexpected affected row count"
    Call rs.Close()
End Sub

Private Function ReadCurrent(ByVal accountId As Long, ByRef reply As t_HousingReply, ByRef entitled As Boolean) As e_HousingCode
    Dim rs As ADODB.Recordset, tier As Long
    ReadCurrent = HousingUnavailable
    Set rs = HousingQuery("SELECT is_active_patron, CURRENT_TIMESTAMP AS server_utc FROM account WHERE id=?", accountId)
    If rs.EOF Then
        Call rs.Close()
        Exit Function
    End If
    If Not IsNull(rs.Fields(0).Value) Then tier = CLng(rs.Fields(0).Value)
    entitled = (tier = 6057394 Or tier = 6057395 Or tier = 28929328 Or tier = 28929365)
    reply.ServerUTC = CStr(rs.Fields(1).Value)
    Call rs.Close()
    Set rs = HousingQuery("SELECT h.key_obj,h.description,CAST(h.assigned_at AS TEXT) AS assigned_utc," & _
        "datetime(h.assigned_at,'+72 hours') AS eligible_utc,CASE WHEN julianday('now')>=julianday(h.assigned_at,'+72 hours') THEN 1 ELSE 0 END AS may_release " & _
        "FROM house_key h WHERE h.account_id=? AND " & ORDINARY & " ORDER BY h.key_obj LIMIT 2", accountId)
    reply.Current.Key = 0: reply.Current.Description = "": reply.AssignedUTC = "": reply.EligibleUTC = ""
    reply.CanRelease = False: reply.CanAcquire = entitled
    If Not rs.EOF Then
        If IsNull(rs.Fields(0).Value) Or IsNull(rs.Fields(1).Value) Then GoTo Invalid
        If CLng(rs.Fields(0).Value) < 1 Or CLng(rs.Fields(0).Value) > 32767 Then GoTo Invalid
        reply.Current.Key = CInt(rs.Fields(0).Value)
        reply.Current.Description = CStr(rs.Fields(1).Value)
        If Len(reply.Current.Description) = 0 Or Len(reply.Current.Description) > 512 Or InStr(reply.Current.Description, vbNullChar) > 0 Then GoTo Invalid
        ' Invalid historical timestamps remain unavailable; never repair them here.
        If IsNull(rs.Fields(2).Value) Or IsNull(rs.Fields(3).Value) Then GoTo Invalid
        reply.AssignedUTC = CStr(rs.Fields(2).Value)
        reply.EligibleUTC = CStr(rs.Fields(3).Value)
        If Len(reply.AssignedUTC) < 19 Or Len(reply.AssignedUTC) > 32 Or Len(reply.EligibleUTC) <> 19 Or InStr(reply.AssignedUTC, vbNullChar) > 0 Then GoTo Invalid
        reply.CanRelease = (CLng(rs.Fields(4).Value) = 1) And HousingMapped(reply.Current.Key)
        reply.CanAcquire = entitled And reply.CanRelease
        Call rs.MoveNext()
        If Not rs.EOF Then GoTo Invalid
    End If
    Call rs.Close()
    ReadCurrent = HousingOkay
    Exit Function
Invalid:
    Call rs.Close()
    reply.Current.Key = 0: reply.Current.Description = "": reply.AssignedUTC = "": reply.EligibleUTC = ""
    reply.CanAcquire = False: reply.CanRelease = False
    ReadCurrent = HousingInconsistent
End Function

Private Sub ReadCatalogue(ByRef request As t_HousingRequest, ByRef reply As t_HousingReply)
    Dim rs As ADODB.Recordset, pattern As String, key As Long, scanned As Long
    pattern = "%" & Replace(Replace(Replace(request.Search, "!", "!!"), "%", "!%"), "_", "!_") & "%"
    Set rs = HousingQuery("SELECT h.key_obj,h.description FROM house_key h WHERE h.account_id IS NULL AND h.key_obj>? " & _
        "AND h.key_obj BETWEEN 1 AND 32767 AND h.description LIKE ? ESCAPE '!' AND " & ORDINARY & " ORDER BY h.key_obj LIMIT 128", request.Cursor, pattern)
    ' Bound scanned rows as well as returned rows; continue sparse catalogues by key.
    Do While Not rs.EOF
        key = CLng(rs.Fields(0).Value)
        scanned = scanned + 1
        If HousingMapped(CInt(key)) Then
            If reply.Count = 4 Then
                reply.NextCursor = reply.Houses(4).Key
                Exit Do
            End If
            If IsNull(rs.Fields(1).Value) Then Err.Raise vbObjectError + 721, "Housing", "Invalid description"
            If Len(CStr(rs.Fields(1).Value)) = 0 Or Len(CStr(rs.Fields(1).Value)) > 512 Or InStr(CStr(rs.Fields(1).Value), vbNullChar) > 0 Then Err.Raise vbObjectError + 721, "Housing", "Invalid description"
            reply.Count = reply.Count + 1
            reply.Houses(reply.Count).Key = CInt(key)
            reply.Houses(reply.Count).Description = CStr(rs.Fields(1).Value)
        End If
        Call rs.MoveNext()
    Loop
    If scanned = 128 And reply.NextCursor = 0 Then reply.NextCursor = CInt(key)
    Call rs.Close()
End Sub

Public Function HousingDbOperation(ByVal accountId As Long, ByRef request As t_HousingRequest, ByRef reply As t_HousingReply) As e_HousingCode
    On Error GoTo Failed
    Dim inTransaction As Boolean, commitAttempted As Boolean, entitled As Boolean
    Dim blankReply As t_HousingReply
    reply = blankReply
    reply.RequestId = request.RequestId: reply.Operation = request.Operation
    Dim outcome As e_HousingCode, oldKey As Integer, newKey As Integer, rs As ADODB.Recordset
    outcome = HousingUnavailable
    If accountId <= 0 Then GoTo Done
    If housingConnection Is Nothing Then
        If request.Operation <> HousingCatalogue Or Len(connectionSettings) = 0 Then GoTo Done
        If Not HousingDbOpen(connectionSettings) Then GoTo Done
    End If
    If request.Operation > HousingExchange Or request.Cursor < 0 Or Len(request.Search) > 96 Then
        outcome = HousingInvalid
        GoTo Done
    End If
    Call housingConnection.Execute("BEGIN IMMEDIATE")
    inTransaction = True
    outcome = ReadCurrent(accountId, reply, entitled)
    If outcome <> HousingOkay Then GoTo Rollback
    oldKey = reply.Current.Key: newKey = oldKey
    If request.Operation = HousingCatalogue Then
        If Not HousingPrepareKeys(accountId, oldKey) Then
            outcome = HousingUnavailable
            GoTo Rollback
        End If
        Call ReadCatalogue(request, reply)
        Call housingConnection.Execute("ROLLBACK")
        inTransaction = False
        Call HousingApplyKeys(accountId, oldKey)
        GoTo Done
    End If
    If request.ExpectedKey <> oldKey Then
        outcome = HousingStale
        GoTo Rollback
    End If
    If request.Operation = HousingRelease Or request.Operation = HousingExchange Then
        If oldKey = 0 Then
            outcome = HousingStale
            GoTo Rollback
        End If
        If Not HousingMapped(oldKey) Then
            outcome = HousingUnavailable
            GoTo Rollback
        End If
        If Not reply.CanRelease Then
            outcome = HousingMinimumPeriod
            GoTo Rollback
        End If
        If Not HousingCanVacate(oldKey) Then
            outcome = HousingOccupied
            GoTo Rollback
        End If
        newKey = 0
    ElseIf request.Operation <> HousingClaim Or oldKey <> 0 Then
        outcome = HousingStale
        GoTo Rollback
    End If
    If request.Operation <> HousingRelease Then
        If Not entitled Then
            outcome = HousingIneligible
            GoTo Rollback
        End If
        newKey = request.TargetKey
        If newKey <= 0 Or newKey = oldKey Or Not HousingMapped(newKey) Then
            outcome = HousingUnavailable
            GoTo Rollback
        End If
        Set rs = HousingQuery("SELECT h.key_obj FROM house_key h WHERE h.key_obj=? AND h.account_id IS NULL AND " & ORDINARY, newKey)
        If rs.EOF Then
            Call rs.Close()
            outcome = HousingStale
            GoTo Rollback
        End If
        Call rs.Close()
    End If
    If Not HousingPrepareKeys(accountId, newKey) Then
        outcome = HousingUnavailable
        GoTo Rollback
    End If
    If oldKey <> 0 Then Call ExecuteOne("UPDATE house_key SET account_id=NULL WHERE key_obj=? AND account_id=?", oldKey, accountId)
#If UNIT_TEST = 1 Then
    If HousingTestFailAfterRelease Then Err.Raise vbObjectError + 722, "Housing", "Injected statement failure"
#End If
    If newKey <> 0 Then Call ExecuteOne("UPDATE house_key SET account_id=?,assigned_at=CURRENT_TIMESTAMP WHERE key_obj=? AND account_id IS NULL", accountId, newKey)
    outcome = ReadCurrent(accountId, reply, entitled)
    If outcome <> HousingOkay Then GoTo Rollback
#If UNIT_TEST = 1 Then
    If HousingTestFailBeforeCommit Then Err.Raise vbObjectError + 723, "Housing", "Injected commit failure"
#End If
    commitAttempted = True
    Call housingConnection.Execute("COMMIT")
    inTransaction = False
#If UNIT_TEST = 1 Then
    If HousingTestLoseCommitReply Then Err.Raise vbObjectError + 724, "Housing", "Injected lost commit reply"
#End If
    ' No yielding between commit and authorization/door reconciliation.
    Call HousingApplyKeys(accountId, newKey)
    If oldKey <> 0 Then Call HousingCloseDoors(oldKey)
    outcome = HousingOkay
    GoTo Done
Rollback:
    Call housingConnection.Execute("ROLLBACK")
    inTransaction = False
Done:
    reply.Code = outcome
    HousingDbOperation = outcome
    Exit Function
Failed:
    Dim number As Long, detail As String, rollbackFailed As Boolean
    number = Err.Number: detail = Err.Description
#If UNIT_TEST = 1 Then
    HousingTestLastError = detail
#End If
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then Call rs.Close()
    End If
    If inTransaction Then
        Err.Clear
        Call housingConnection.Execute("ROLLBACK")
        rollbackFailed = (Err.Number <> 0)
    End If
    On Error GoTo Unresolved
    If rollbackFailed Then
        ' Closing the dedicated connection resolves any still-open transaction.
        ' No later request may accidentally run in its unknown transaction state.
        Call HousingDbClose()
        If Not commitAttempted Then GoTo Unresolved
        If Not HousingDbOpen(connectionSettings) Then GoTo Unresolved
    End If
    If commitAttempted Then
        ' Read actual committed ownership, not coordinates or a retry of the mutation.
        ' The serialized world has not admitted another gameplay action in this call.
        outcome = ReadCurrent(accountId, reply, entitled)
        If outcome <> HousingOkay Then GoTo Unresolved
        If Not HousingPrepareKeys(accountId, reply.Current.Key) Then GoTo Unresolved
        Call HousingApplyKeys(accountId, reply.Current.Key)
        If oldKey <> 0 And reply.Current.Key <> oldKey Then Call HousingCloseDoors(oldKey)
        outcome = HousingUnresolved ' Client must refresh status, never replay automatically.
    Else
        outcome = HousingUnavailable
        If InStr(1, detail, "locked", vbTextCompare) > 0 Or InStr(1, detail, "busy", vbTextCompare) > 0 Then outcome = HousingBusy
    End If
    Call LogDatabaseError("Housing operation failed: " & CStr(number) & "; outcome=" & CStr(outcome))
    GoTo Done
Unresolved:
    Call HousingBlockAccount(accountId)
    outcome = HousingUnresolved
    GoTo Done
End Function
