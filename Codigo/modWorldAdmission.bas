Attribute VB_Name = "modWorldAdmission"
Option Explicit

Private Const CLAIM_SQL As String = "UPDATE account SET world_claim=?,world_instance=?,world_lock_path=? WHERE world_claim='' AND id IN (SELECT a.id FROM tokens t JOIN account a ON a.email=t.username WHERE t.encrypted_token=? AND t.decrypted_token=?)"
Private Const RELEASE_SQL As String = "UPDATE account SET world_claim='',world_instance='',world_lock_path='' WHERE world_claim=? AND world_instance=?"

' Housing and world admission share short SQLite write transactions, never a
' transaction spanning gameplay. All runtime coordination uses ADO async work.
Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef value As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal path As String, ByVal access As Long, ByVal sharing As Long, ByVal security As Long, ByVal disposition As Long, ByVal flags As Long, ByVal template As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal handle As Long) As Long
Private mEnabled As Boolean
Private mStartupFailed As Boolean
Private mInstance As String
Private mLockPath As String
Private mProcessLock As Long
Private mJobs As New Dictionary
Private mConnections As New Dictionary
Private mSaves As New Dictionary
Private mSaveScope As String
Private mObservers() As clsDatabaseCompletion
Private mDatabase As ADODB.Connection
Private mObserver As clsDatabaseCompletion
Private mCommand As ADODB.Command
Private mJob As clsWorldAdmission
Private mDone As Boolean
Private mFailed As Boolean
Private mRows As Long
Private mResult As ADODB.Recordset
Private mSession As ADODB.Recordset

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Sub CheckListenerSafety()
    ' GUI restart must not bypass a failed coordinated startup.
    If mStartupFailed Then Err.Raise 5, , "World admission initialization failed; restart after correcting the database error"
End Sub

Private Function NewIdentity() As String
    Dim bytes(0 To 15) As Byte
    Dim i As Long
    If CoCreateGuid(bytes(0)) <> 0 Then Err.Raise 5, , "Unable to allocate world claim identity"
    For i = 0 To 15
        NewIdentity = NewIdentity & Right$("0" & Hex$(bytes(i)), 2)
    Next i
End Function

Public Sub Initialize()
    On Error GoTo Failed
    Dim rows As ADODB.Recordset
    Dim path As String, oldInstance As String, oldLockPath As String
    Dim handle As Long, i As Long
    Dim failure As String
    Set rows = Connection.Execute("SELECT name FROM sqlite_master WHERE type='table' AND name='housing_coordination'")
    If rows.EOF Then
        Call rows.Close
        Exit Sub
    End If
    Call rows.Close
#If PYMMO = 0 Then
    Err.Raise 5, , "Coordinated admission requires PYMMO=1"
#End If
#If LOGIN_STRESS_TEST = 1 Then
    Err.Raise 5, , "Coordinated admission does not permit stress authentication bypass"
#End If
    Set rows = Connection.Execute("PRAGMA database_list")
    Do Until rows.EOF
        If CStr(rows.Fields(1).value) = "main" Then path = CStr(rows.Fields(2).value)
        Call rows.MoveNext
    Loop
    Call rows.Close
    If Len(path) = 0 Or Left$(path, 2) = "\\" Then Err.Raise 5, , "World coordination requires a local SQLite file"
    mInstance = NewIdentity()
    mLockPath = path & ".world-" & mInstance & ".lock"
    mProcessLock = CreateFile(mLockPath, &HC0000000, 0, 0, 4, &H80, 0)
    If mProcessLock = -1 Then Err.Raise 5, , "Cannot acquire world process lock"
    Set mDatabase = New ADODB.Connection
    mDatabase.CursorLocation = adUseClient
    Call mDatabase.Open("DRIVER={SQLite3 ODBC Driver};DATABASE=" & path)
    ' Startup refuses contention rather than waiting on a lock. Runtime work
    ' is asynchronous and does not use the blocking Query helper.
    Call mDatabase.Execute("PRAGMA busy_timeout=0")
    Call mDatabase.Execute("BEGIN IMMEDIATE")
    Set rows = mDatabase.Execute("SELECT DISTINCT world_instance,world_lock_path FROM account WHERE world_claim<>''")
    Do Until rows.EOF
        oldInstance = CStr(rows.Fields(0).value)
        oldLockPath = CStr(rows.Fields(1).value)
        If Len(oldInstance) <> 32 Or oldInstance Like "*[!0-9A-F]*" Then Err.Raise 5, , "Invalid world instance identity"
        If Left$(oldLockPath, 2) = "\\" Or Right$(oldLockPath, 44) <> ".world-" & oldInstance & ".lock" Then Err.Raise 5, , "Invalid world instance lock path"
        ' Use the original holder's path, not this process's database alias.
        handle = CreateFile(oldLockPath, &HC0000000, 0, 0, 4, &H80, 0)
        If handle <> -1 Then
            ' Exclusive OS ownership proves the old process is gone. No TTL.
            Call StartupCommand("UPDATE account SET world_claim='',world_instance='',world_lock_path='' WHERE world_instance=?", oldInstance)
            Call CloseHandle(handle)
        End If
        Call rows.MoveNext
    Loop
    Call rows.Close
    Call mDatabase.Execute("UPDATE housing_coordination SET world_version=3 WHERE id=1")
    Call mDatabase.Execute("COMMIT")
    Set mObserver = New clsDatabaseCompletion
    Call mObserver.Attach(mDatabase, True)
    ReDim mObservers(1 To MAX_ASYNC)
    For i = 1 To MAX_ASYNC
        Set mObservers(i) = New clsDatabaseCompletion
        Call mObservers(i).Attach(Connection_async(i), False)
    Next i
    mEnabled = True
    Exit Sub
Failed:
    ' Never listen without admission protection after housing schema is present.
    failure = Err.Description
    mStartupFailed = True
    On Error Resume Next
    If Not mDatabase Is Nothing Then
        Call mDatabase.Execute("ROLLBACK")
        Call mDatabase.Close
    End If
    Set mDatabase = Nothing
    If mProcessLock <> 0 And mProcessLock <> -1 Then Call CloseHandle(mProcessLock)
    mProcessLock = 0
    Call LogError("World admission initialization failed: " & failure)
    On Error GoTo 0
    Err.Raise 5, "modWorldAdmission.Initialize", "World admission initialization failed; no listener may start"
End Sub

Public Function Queue(ByVal connectionId As Long, ByVal packetId As Long, ByVal message As Network.reader) As Boolean
    On Error GoTo Failed
    Dim job As clsWorldAdmission
    Dim bytes() As Byte, encrypted() As Byte
    Queue = True
    If mConnections.Exists(CStr(connectionId)) Or mJobs.Count >= MaxUsers + 128 Then GoTo Failed
    Set job = New clsWorldAdmission
    job.ClaimId = NewIdentity()
    job.ConnectionId = connectionId
    job.PacketId = packetId
    Call message.GetData(bytes)
    Set job.Message = New Network.reader
    Call job.Message.SetData(bytes)
    Call job.Message.Skip(2)
    job.EncryptedToken = job.Message.ReadString8
    If Len(job.EncryptedToken) <> 88 Then GoTo Failed
    Call AO20CryptoSysWrapper.Str2ByteArr(job.EncryptedToken, encrypted)
    job.DecryptedToken = AO20CryptoSysWrapper.DECRYPT(PrivateKey, cnvStringFromHexStr(cnvToHex(encrypted)))
    If Len(job.DecryptedToken) <> 64 Or Not IsBase64(job.DecryptedToken) Then GoTo Failed
    Call job.Message.SetData(bytes)
    Call mJobs.Add(job.ClaimId, job)
    Call mConnections.Add(CStr(connectionId), job.ClaimId)
    Exit Function
Failed:
    Call KickConnection(connectionId, "world_admission_rejected", "modWorldAdmission.Queue")
End Function

Private Function BuildCommand(ByVal sql As String, ByVal values As Variant) As ADODB.Command
    Dim value As Variant
    Dim command As New ADODB.Command
    Set command.ActiveConnection = mDatabase
    command.CommandType = adCmdText
    command.CommandText = sql
    For Each value In values
        Call command.Parameters.Append(command.CreateParameter(, adVarChar, adParamInput, Len(CStr(value)), CStr(value)))
    Next value
    Set BuildCommand = command
End Function

Private Sub StartupCommand(ByVal sql As String, ParamArray values() As Variant)
    Dim command As ADODB.Command
    Set command = BuildCommand(sql, values)
    Call command.Execute()
End Sub

Private Sub StartCommand(ByVal sql As String, ParamArray values() As Variant)
    Set mCommand = BuildCommand(sql, values)
    mDone = False
    mFailed = False
    mRows = 0
    Set mResult = Nothing
    Call mCommand.Execute(, , adAsyncExecute)
End Sub

Public Sub CoordinatorCompleted(ByVal rows As Long, ByVal failure As ADODB.Error, ByVal result As ADODB.Recordset)
    mRows = rows
    mFailed = Not failure Is Nothing
    Set mResult = result
    mDone = True
End Sub

Private Function SavesFinished(ByVal claimId As String) As Boolean
    Dim key As Variant, entry As Variant
    For Each key In mSaves.Keys
        entry = mSaves(key)
        If entry(0) = claimId Then Exit Function
    Next key
    SavesFinished = True
End Function

Private Function IsCurrent(ByVal job As clsWorldAdmission) As Boolean
    If mConnections.Exists(CStr(job.ConnectionId)) Then IsCurrent = (mConnections(CStr(job.ConnectionId)) = job.ClaimId)
End Function

Public Sub Poll()
    On Error GoTo Failed
    If Not mEnabled Then Exit Sub
    Dim key As Variant, job As clsWorldAdmission
    If Not mJob Is Nothing Then
        If Not mDone Then Exit Sub
        Set job = mJob
        Set mJob = Nothing
        Set mCommand = Nothing
        If mFailed Then
            ' Failed/uncertain writes are never interpreted as a released claim.
            job.Retired = True
            job.Stage = 3
            If job.UserIndex = 0 And IsCurrent(job) Then Call KickConnection(job.ConnectionId, "world_admission_database_error")
        ElseIf job.Stage = 1 Then
            If mRows <> 1 Then
                job.Retired = True
                job.Stage = 3
                If IsCurrent(job) Then Call KickConnection(job.ConnectionId, "world_account_busy_or_invalid_session")
            Else
                job.Stage = 2
            End If
        ElseIf job.Stage = 2 Then
            If Not job.Retired And IsCurrent(job) Then
                If Not mResult Is Nothing Then
                    If Not mResult.EOF Then
                        Set mSession = mResult
                        job.Stage = 4
                        mSaveScope = job.ClaimId
                        #If PYMMO = 1 Then
                            Call Protocol.ResumeWorldAdmission(job.ConnectionId, job.PacketId, job.Message)
                        #End If
                        Call EndSave()
                        Set mSession = Nothing
                    End If
                End If
            End If
            If job.UserIndex = 0 Then
                job.Retired = True
                job.Stage = 3
            End If
        ElseIf job.Stage = 3 Then
            If IsCurrent(job) Then Call mConnections.Remove(CStr(job.ConnectionId))
            Call mJobs.Remove(job.ClaimId)
        End If
        If Not mResult Is Nothing Then
            If mResult.State = adStateOpen Then Call mResult.Close
            Set mResult = Nothing
        End If
        Exit Sub
    End If
    For Each key In mJobs.Keys
        Set job = mJobs(key)
        If job.Retired Then
            If Not job.SaveFailed And SavesFinished(job.ClaimId) Then
                job.Stage = 3
                Set mJob = job
                Call StartCommand(RELEASE_SQL, job.ClaimId, mInstance)
                Exit Sub
            End If
        ElseIf job.Stage = 0 Then
            job.Stage = 1
            Set mJob = job
            ' One autocommit statement validates the still-current token AND
            ' claims the account under the same SQLite writer exclusion as pay.
            Call StartCommand(CLAIM_SQL, job.ClaimId, mInstance, mLockPath, job.EncryptedToken, job.DecryptedToken)
            Exit Sub
        ElseIf job.Stage = 2 Then
            Set mJob = job
            Call StartCommand("SELECT t.* FROM tokens t JOIN account a ON a.email=t.username WHERE a.world_claim=? AND a.world_instance=? AND t.encrypted_token=? AND t.decrypted_token=?", job.ClaimId, mInstance, job.EncryptedToken, job.DecryptedToken)
            Exit Sub
        End If
    Next key
    Exit Sub
Failed:
    Call LogError("World admission asynchronous operation failed: " & Err.Description)
    Call SaveFailed()
    Call EndSave()
    If Not mJob Is Nothing Then
        mJob.Retired = True
        mJob.Stage = 3
    End If
    Set mJob = Nothing
End Sub

#If UNIT_TEST = 1 Then
Private Sub AwaitTestCompletion()
    Dim started As Long
    started = GetTickCountRaw()
    Do While Not mDone And TicksElapsed(started, GetTickCountRaw()) < 5000
        DoEvents
    Loop
    If Not mDone Then Err.Raise 5, , "ADO completion event was not delivered"
End Sub

Public Function test_suite_world_admission() As Boolean
    On Error GoTo Failed
    Dim originalEnabled As Boolean, failedBefore As Long
    Dim path As String, first As ADODB.Connection, saveConnection As ADODB.Connection
    Dim saveObserver As clsDatabaseCompletion, saveCommand As ADODB.Command
    Dim started As Long, elapsed As Double, ticks As Long
    Dim rows As ADODB.Recordset, job As clsWorldAdmission, otherJob As clsWorldAdmission
    Dim originalInstance As String
    Dim liveLock As Long, otherLock As Long
    failedBefore = UnitTesting.GetFailedTests()
    originalEnabled = mEnabled
    originalInstance = mInstance
    If mEnabled Or mJobs.Count <> 0 Then Err.Raise 5, , "Admission tests require an isolated non-coordinated test server"
    path = App.Path & "\admission-test-" & NewIdentity() & ".db"
    liveLock = CreateFile(path & ".lock", &HC0000000, 0, 0, 4, &H80, 0)
    otherLock = CreateFile(path & ".lock", &HC0000000, 0, 0, 4, &H80, 0)
    Call UnitTesting.RunTest("admission live instance lock cannot be reclaimed", liveLock <> -1 And otherLock = -1)
    If otherLock <> -1 Then Call CloseHandle(otherLock)
    Call CloseHandle(liveLock)
    liveLock = 0
    otherLock = CreateFile(path & ".lock", &HC0000000, 0, 0, 4, &H80, 0)
    Call UnitTesting.RunTest("admission released instance lock can be reclaimed", otherLock <> -1)
    If otherLock <> -1 Then Call CloseHandle(otherLock)
    Set first = New ADODB.Connection
    Call first.Open("DRIVER={SQLite3 ODBC Driver};DATABASE=" & path)
    Call first.Execute("CREATE TABLE account(id INTEGER PRIMARY KEY,email TEXT,world_claim TEXT NOT NULL DEFAULT '',world_instance TEXT NOT NULL DEFAULT '',world_lock_path TEXT NOT NULL DEFAULT '')")
    Call first.Execute("CREATE TABLE tokens(id INTEGER PRIMARY KEY,username TEXT,encrypted_token TEXT,decrypted_token TEXT)")
    Call first.Execute("INSERT INTO account(id,email) VALUES(1,'owner')")
    Call first.Execute("INSERT INTO tokens VALUES(1,'owner','encrypted','decrypted')")
    Set mDatabase = New ADODB.Connection
    mDatabase.CursorLocation = adUseClient
    Call mDatabase.Open("DRIVER={SQLite3 ODBC Driver};DATABASE=" & path)
    Call mDatabase.Execute("PRAGMA busy_timeout=0")
    Set mObserver = New clsDatabaseCompletion
    Call mObserver.Attach(mDatabase, True)
    mEnabled = True
    Call first.Execute("BEGIN IMMEDIATE")
    started = GetTickCountRaw()
    Call StartCommand(CLAIM_SQL, "first", "world", path & ".lock", "encrypted", "decrypted")
    elapsed = TicksElapsed(started, GetTickCountRaw())
    Call UnitTesting.RunTest("admission dispatch never waits for database lock", elapsed < 250)
    Do While TicksElapsed(started, GetTickCountRaw()) < 250
        ticks = ticks + 1
        DoEvents
    Loop
    Call UnitTesting.RunTest("admission keeps loop responsive during contention", ticks > 1)
    Call first.Execute("ROLLBACK")
    Call AwaitTestCompletion()
    ' Contention may finish with BUSY or succeed after release, depending on ODBC.
    Call StartCommand(RELEASE_SQL, "first", "world")
    Call AwaitTestCompletion()
    Call StartCommand(CLAIM_SQL, "first", "world", path & ".lock", "encrypted", "decrypted")
    Call AwaitTestCompletion()
    Call UnitTesting.RunTest("admission validates token and claims account atomically", Not mFailed And mRows = 1)
    Call StartCommand(CLAIM_SQL, "second", "world", path & ".lock", "encrypted", "decrypted")
    Call AwaitTestCompletion()
    Call UnitTesting.RunTest("admission rejects duplicate account claim", Not mFailed And mRows = 0)
    Call StartCommand(RELEASE_SQL, "stale", "world")
    Call AwaitTestCompletion()
    Call UnitTesting.RunTest("admission old cleanup cannot release current claim", Not mFailed And mRows = 0)
    Call StartCommand(RELEASE_SQL, "first", "other-process")
    Call AwaitTestCompletion()
    Call UnitTesting.RunTest("admission cleanup is process specific", Not mFailed And mRows = 0)
    Call StartCommand(RELEASE_SQL, "first", "world")
    Call AwaitTestCompletion()
    Call UnitTesting.RunTest("admission matching cleanup releases claim", Not mFailed And mRows = 1)
    Set rows = first.Execute("SELECT id,username,encrypted_token,decrypted_token FROM tokens")
    Call UnitTesting.RunTest("admission preserves original session token", rows.Fields(0).value = 1 And rows.Fields(1).value = "owner" And rows.Fields(2).value = "encrypted" And rows.Fields(3).value = "decrypted")
    Call rows.Close()
    Call first.Execute("DELETE FROM tokens")
    Call StartCommand(CLAIM_SQL, "old-token", "world", path & ".lock", "encrypted", "decrypted")
    Call AwaitTestCompletion()
    Call UnitTesting.RunTest("admission revoked issued token cannot claim", Not mFailed And mRows = 0)

    Set saveConnection = New ADODB.Connection
    Call saveConnection.Open("DRIVER={SQLite3 ODBC Driver};DATABASE=" & path)
    Set saveObserver = New clsDatabaseCompletion
    Call saveObserver.Attach(saveConnection, False)
    Set saveCommand = New ADODB.Command
    Set saveCommand.ActiveConnection = saveConnection
    saveCommand.CommandType = adCmdText
    saveCommand.CommandText = "UPDATE account SET email='saved' WHERE id=1"
    Set job = New clsWorldAdmission
    job.ClaimId = "save-owner"
    job.Stage = 4
    Call mJobs.Add(job.ClaimId, job)
    mSaveScope = job.ClaimId
    Call TrackSave(saveCommand)
    Call UnitTesting.RunTest("admission cannot release before own save completion", Not SavesFinished(job.ClaimId))
    Call UnitTesting.RunTest("admission unrelated account does not wait for save", SavesFinished("other-owner"))
    Call saveCommand.Execute(, , adAsyncExecute)
    started = GetTickCountRaw()
    Do While Not SavesFinished(job.ClaimId) And TicksElapsed(started, GetTickCountRaw()) < 5000
        DoEvents
    Loop
    Call UnitTesting.RunTest("admission observes actual asynchronous save completion", SavesFinished(job.ClaimId))
    Call TrackSave(saveCommand)
    Call SaveDispatchFailed(saveCommand, False)
    Call UnitTesting.RunTest("admission definite pre-dispatch failure retires tracking", SavesFinished(job.ClaimId) And Not job.SaveFailed)
    Call TrackSave(saveCommand)
    Call SaveDispatchFailed(saveCommand, True)
    Call UnitTesting.RunTest("admission uncertain dispatch retains only owner", job.SaveFailed And SavesFinished("other-owner"))
    Call UnitTesting.RunTest("admission failed dispatch leaves no global barrier", mSaves.Count = 0)
    ' Exercise the production retirement path, not just the readiness predicate.
    mInstance = "test-process"
    Call first.Execute("INSERT INTO account(id,email,world_claim,world_instance,world_lock_path) VALUES(2,'other','other-owner','test-process','test.lock')")
    Set otherJob = New clsWorldAdmission
    otherJob.ClaimId = "other-owner"
    otherJob.Retired = True
    Call mJobs.Add(otherJob.ClaimId, otherJob)
    Call Poll()
    Call AwaitTestCompletion()
    Call Poll()
    Set rows = first.Execute("SELECT COUNT(*) FROM account WHERE id=2 AND world_claim<>''")
    Call UnitTesting.RunTest("admission dispatch failure cannot strand unrelated account", CLng(rows.Fields(0).Value) = 0 And job.SaveFailed)
    Call rows.Close()
    Call UnitTesting.RunTest("admission unrelated retired job is released", Not mJobs.Exists("other-owner"))
    Call SaveDispatchFailed(saveCommand, True)
    Call SaveCompleted(saveCommand, Nothing)
    Call SaveCompleted(saveCommand, Nothing)
    Call UnitTesting.RunTest("admission late duplicate completion preserves uncertainty", job.SaveFailed And mSaves.Count = 0)
    Call mJobs.Remove(job.ClaimId)
    Set job = New clsWorldAdmission
    job.ClaimId = "replacement-owner"
    Call mJobs.Add(job.ClaimId, job)
    Call SaveCompleted(saveCommand, Nothing)
    Call UnitTesting.RunTest("admission old completion cannot mark new claim", Not job.SaveFailed And SavesFinished(job.ClaimId))
    Call EndSave()
    Call TrackSave(saveCommand)
    Call UnitTesting.RunTest("admission unrelated SQL is not tracked", mSaves.Count = 0)
    Set job = New clsWorldAdmission
    job.ClaimId = "new-session"
    job.ConnectionId = 9
    Call mConnections.Add("9", "old-session")
    Call UnitTesting.RunTest("admission stale callback ignores recycled connection", Not IsCurrent(job))
    mConnections("9") = "new-session"
    Call UnitTesting.RunTest("admission callback matches exact claim identity", IsCurrent(job))
Cleanup:
    On Error Resume Next
    Set rows = Nothing
    Set mSession = Nothing
    Set mResult = Nothing
    Set mCommand = Nothing
    Set saveCommand = Nothing
    Set saveObserver = Nothing
    If Not saveConnection Is Nothing Then Call saveConnection.Close
    Set saveConnection = Nothing
    Set mObserver = Nothing
    If Not mDatabase Is Nothing Then Call mDatabase.Close
    Set mDatabase = Nothing
    If Not first Is Nothing Then Call first.Close
    Set first = Nothing
    Call mSaves.RemoveAll
    Call mConnections.RemoveAll
    Call mJobs.RemoveAll
    Call EndSave()
    mEnabled = originalEnabled
    mInstance = originalInstance
    Kill path
    Kill path & ".lock"
    test_suite_world_admission = (UnitTesting.GetFailedTests() = failedBefore)
    Exit Function
Failed:
    Call UnitTesting.RunTestError("world admission", Err.Description)
    Resume Cleanup
End Function
#End If

Public Function SessionForConnection(ByVal connectionId As Long) As ADODB.Recordset
    If mConnections.Exists(CStr(connectionId)) Then Set SessionForConnection = mSession
End Function

Public Sub AttachUser(ByVal connectionId As Long, ByVal userIndex As Integer)
    If Not mEnabled Then Exit Sub
    Dim job As clsWorldAdmission
    If Not mConnections.Exists(CStr(connectionId)) Then Err.Raise 5, , "Missing world admission claim"
    Set job = mJobs(mConnections(CStr(connectionId)))
    job.UserIndex = userIndex
    job.UserVersion = UserList(userIndex).VersionId
End Sub

Public Sub ConnectionClosed(ByVal connectionId As Long)
    If Not mEnabled Then Exit Sub
    Dim job As clsWorldAdmission
    If Not mConnections.Exists(CStr(connectionId)) Then Exit Sub
    Set job = mJobs(mConnections(CStr(connectionId)))
    Call mConnections.Remove(CStr(connectionId))
    If job.UserIndex = 0 Then
        job.Retired = True
    End If
End Sub

Private Function FindUser(ByVal userIndex As Integer) As clsWorldAdmission
    Dim key As Variant, job As clsWorldAdmission
    For Each key In mJobs.Keys
        Set job = mJobs(key)
        If job.UserIndex = userIndex And job.UserVersion = UserList(userIndex).VersionId And Not job.Retired Then
            Set FindUser = job
            Exit Function
        End If
    Next key
End Function

Public Sub UserRetired(ByVal userIndex As Integer)
    If Not mEnabled Then Exit Sub
    Dim job As clsWorldAdmission
    Set job = FindUser(userIndex)
    If job Is Nothing Then Exit Sub
    job.Retired = True
End Sub

Public Sub BeginSave(ByVal userIndex As Integer)
    If Not mEnabled Then Exit Sub
    Dim job As clsWorldAdmission
    Set job = FindUser(userIndex)
    mSaveScope = vbNullString
    If Not job Is Nothing Then mSaveScope = job.ClaimId
End Sub

Public Sub EndSave()
    mSaveScope = vbNullString
End Sub

Public Sub SaveFailed()
    If Len(mSaveScope) = 0 Then Exit Sub
    Dim job As clsWorldAdmission
    Set job = mJobs(mSaveScope)
    job.SaveFailed = True
End Sub

Public Sub TrackSave(ByVal command As ADODB.Command)
    If Not mEnabled Or Len(mSaveScope) = 0 Then Exit Sub
    Dim key As String
    key = CStr(ObjPtr(command))
    If mSaves.Exists(key) Then Err.Raise 5, , "Database command reused before completion"
    ' Keep the command alive and immutable until its execution is retired.
    Call mSaves.Add(key, Array(mSaveScope, command))
End Sub

Public Sub SaveDispatchFailed(ByVal command As ADODB.Command, ByVal executionEntered As Boolean)
    If Not mEnabled Or command Is Nothing Then Exit Sub
    Dim key As String, entry As Variant, job As clsWorldAdmission
    key = CStr(ObjPtr(command))
    If Not mSaves.Exists(key) Then Exit Sub
    entry = mSaves(key)
    ' Once Execute was entered, a thrown COM error does not prove that the
    ' provider stopped. Retain uncertainty on this exact claim, never globally.
    If executionEntered And mJobs.Exists(entry(0)) Then
        Set job = mJobs(entry(0))
        job.SaveFailed = True
    End If
    Call mSaves.Remove(key)
End Sub

Public Sub SaveCompleted(ByVal command As ADODB.Command, ByVal failure As ADODB.Error)
    If Not mEnabled Or command Is Nothing Then Exit Sub
    Dim key As String, entry As Variant, job As clsWorldAdmission
    key = CStr(ObjPtr(command))
    If Not mSaves.Exists(key) Then Exit Sub
    entry = mSaves(key)
    If Not failure Is Nothing Then
        If Len(entry(0)) > 0 Then
            If mJobs.Exists(entry(0)) Then
                Set job = mJobs(entry(0))
                job.SaveFailed = True
            End If
        End If
        Call LogError("Asynchronous save failed; its world claim will not be released: " & failure.Description)
    End If
    Call mSaves.Remove(key)
End Sub
