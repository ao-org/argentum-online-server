Attribute VB_Name = "Unit_CharacterSaves"
Option Explicit
#If UNIT_TEST = 1 Then
Private db As ADODB.Connection
Private path As String
Private connectionString As String
Private passed As Boolean
Private snapshots() As t_User
Private originalDB As ADODB.Connection

Public Function test_suite_character_saves() As Boolean
    On Error GoTo Failed
    ReDim snapshots(1 To 6)
    Dim oldMax As Integer, originalConnection As String
    snapshots(1) = UserList(1)
    oldMax = MaxUsers
    originalConnection = "DRIVER={SQLite3 ODBC Driver};DATABASE=" & App.Path & "/Database.db"
    If Len(Database_Source) Then originalConnection = "DATA SOURCE=" & Database_Source & ";"
    path = App.Path & "\character-save-unit.db"
    If Len(Dir$(path)) Then Kill path
    connectionString = "DRIVER={SQLite3 ODBC Driver};DATABASE=" & path
    Set db = New ADODB.Connection
    Call db.Open(connectionString)
    Call db.Execute("CREATE TABLE saved(id INTEGER PRIMARY KEY, value INTEGER NOT NULL CHECK(value >= 0))")
    Call db.Execute("INSERT INTO saved VALUES(901,0),(902,0)")
    Call db.Execute("CREATE TABLE audit(value TEXT)")
    passed = True
    Call TestOrdering()
    Call TestFailure()
    Call TestAcknowledgement()
    Call TestCapacityAndShutdown()
    Call TestCallbacks()
    Call TestProductionSnapshot()
Done:
    On Error Resume Next
    If Not originalDB Is Nothing Then Set Connection = originalDB
    Set originalDB = Nothing
    Call InitializeCharacterSaves(originalConnection)
    UserList(1) = snapshots(1)
    MaxUsers = oldMax
    If Not db Is Nothing Then Call db.Close()
    Set db = Nothing
    test_suite_character_saves = passed
    Exit Function
Failed:
    passed = False
    Call UnitTesting.RunTestError("Character save native regression", CStr(Err.Number) & " " & Err.Description)
    Resume Done
End Function

Private Sub Check(ByVal condition As Boolean, ByVal name As String)
    Call UnitTesting.RunTest("saves " & name, condition)
    If Not condition Then passed = False
End Sub

Private Function Reserve(ByVal id As Long) As t_User
    Dim reason As String, User As t_User
    User.Id = id
    If Not ReserveCharacterSaveSession(User, reason) Then Call Err.Raise(5, , reason)
    Reserve = User
End Function

Private Function NewJob(ByRef User As t_User, ByVal finalSave As Boolean, ByVal sql As String) As clsCharacterSave
    Dim job As New clsCharacterSave
    Call job.Initialize(0, User, finalSave)
    Call job.Add(sql)
    Call job.Seal(User)
    Set NewJob = job
End Function

Private Sub AwaitJob(ByVal job As clsCharacterSave)
    Dim started As Long
    started = GetTickCountRaw()
    Do While job.Result = CharacterSavePending Or job.Result = CharacterSaveRunning
        Call TickCharacterSaves()
        DoEvents
        If TicksElapsed(started, GetTickCountRaw()) > 10000 Then Call Err.Raise(5, , "Native save callback timeout")
    Loop
    Call TickCharacterSaves()
End Sub

Private Function Value(ByVal sql As String) As Long
    Dim rs As ADODB.Recordset
    Set rs = db.Execute(sql)
    Value = CLng(rs.Fields(0).value)
    Call rs.Close()
End Function

Private Sub TestOrdering()
    Call InitializeCharacterSaves(connectionString)
    Dim reason As String
    Dim auto As New clsCharacterSave, final As clsCharacterSave
    snapshots(2) = Reserve(901)
    Call auto.Initialize(0, snapshots(2), False)
    Call auto.Add("WITH RECURSIVE n(x) AS (VALUES(1) UNION ALL SELECT x+1 FROM n WHERE x<1000000) UPDATE saved SET value=(SELECT max(x)/1000000 FROM n) WHERE id=901")
    Call auto.Add("INSERT INTO audit VALUES('auto')")
    Call auto.Seal(snapshots(2))
    Set final = NewJob(snapshots(2), True, "UPDATE saved SET value=2 WHERE id=901")
    Call EnqueueCharacterSave(auto)
    Call EnqueueCharacterSave(final)
    snapshots(2).Id = 902
    Call Check(Not ReserveCharacterSaveSession(snapshots(2), reason), "replacement login cannot retarget an existing reservation")
    Call ReleaseCharacterSaveSession(snapshots(2))
    snapshots(4).Id = 901
    Call Check(Not ReserveCharacterSaveSession(snapshots(4), reason), "immediate reconnect refused")
    Call TickCharacterSaves()
    Call Check(CharacterHasUnfinishedSave(901) And final.Result = CharacterSavePending, "logout waits for delayed autosave")
    Call AwaitJob(final)
    Call Check(auto.Result = CharacterSaveSucceeded And final.Result = CharacterSaveSucceeded And Value("SELECT value FROM saved WHERE id=901") = 2, "final state wins")
    Call Check(Value("SELECT count(*) FROM audit") = 1, "autosave executes once")
    Call Check(ReserveCharacterSaveSession(snapshots(4), reason), "completion permits reconnect")
    Call ReleaseCharacterSaveSession(snapshots(4))
End Sub

Private Sub TestFailure()
    Call InitializeCharacterSaves(connectionString)
    Dim reason As String
    Dim job As New clsCharacterSave, final As clsCharacterSave, independent As clsCharacterSave
    snapshots(2) = Reserve(901)
    snapshots(3) = Reserve(902)
    Call job.Initialize(0, snapshots(2), False)
    Call job.Add("INSERT INTO audit VALUES('before-failure')")
    Call job.Add("INSERT INTO missing_save_table VALUES(1)")
    Call job.Add("UPDATE saved SET value=3 WHERE id=901")
    Call job.Seal(snapshots(2))
    Set final = NewJob(snapshots(2), True, "UPDATE saved SET value=4 WHERE id=901")
    Set independent = NewJob(snapshots(3), True, "UPDATE saved SET value=7 WHERE id=902")
    Call EnqueueCharacterSave(job)
    Call EnqueueCharacterSave(final)
    Call EnqueueCharacterSave(independent)
    Call ReleaseCharacterSaveSession(snapshots(2))
    Call ReleaseCharacterSaveSession(snapshots(3))
    Call AwaitJob(job)
    Call AwaitJob(independent)
    snapshots(4).Id = 901
    Call Check(job.Result = CharacterSaveFailed And job.NextStatement = 2 And final.Result = CharacterSavePending, "write failure is not complete; final held")
    Call Check(Not ReserveCharacterSaveSession(snapshots(4), reason), "failed save blocks stale reload")
    Call Check(independent.Result = CharacterSaveSucceeded And Value("SELECT value FROM saved WHERE id=902") = 7, "another character completes independently")
    Call Check(Value("SELECT value FROM saved WHERE id=901") = 2, "later statements do not run after error")
    ' Explicit repair and operator resume do not replay the already committed audit insert.
    Call db.Execute("CREATE TABLE missing_save_table(value INTEGER)")
    Call Check(RetryFailedCharacterSave(901), "operator resumes verified failed statement")
    Call AwaitJob(final)
    Call Check(Value("SELECT count(*) FROM audit WHERE value='before-failure'") = 1 And Value("SELECT value FROM saved WHERE id=901") = 4, "recovery preserves completed-write order")
    snapshots(2) = Reserve(901)
    Set job = NewJob(snapshots(2), True, "UPDATE saved SET value=-1 WHERE id=901")
    Call EnqueueCharacterSave(job)
    Call AwaitJob(job)
    Call Check(job.Result = CharacterSaveFailed And Value("SELECT value FROM saved WHERE id=901") = 4, "autocommit constraint failure cannot succeed")
End Sub

Private Sub TestAcknowledgement()
    Call InitializeCharacterSaves(connectionString)
    Dim job As New clsCharacterSave
    snapshots(2) = Reserve(901)
    snapshots(2).Stats.UserSkills(1) = 10
    snapshots(2).Persist.LastSkills(1) = 10
    snapshots(2).Persist.LastInventory(1).amount = 2
    UserList(1) = snapshots(2)
    snapshots(5) = snapshots(2)
    Call job.Initialize(1, snapshots(2), False)
    Call job.Add("UPDATE saved SET value=10 WHERE id=901")
    Call job.Seal(snapshots(5))
    Call EnqueueCharacterSave(job)
    UserList(1).Stats.UserSkills(1) = 11
    UserList(1).Stats.SkillDirty(1) = True
    UserList(1).invent.Object(1).amount = 3
    Call AwaitJob(job)
    Call Check(UserList(1).Stats.SkillDirty(1) And UserList(1).Persist.LastSkills(1) = 10, "older save cannot clear newer skill change")
    Call Check(UserList(1).invent.Object(1).amount = 3 And UserList(1).Persist.LastInventory(1).amount = 2, "inventory acknowledgement retains newer state")
    Set job = New clsCharacterSave
    Call job.Initialize(1, snapshots(2), False)
    Call job.Add("UPDATE saved SET value=11 WHERE id=901")
    Call job.Seal(snapshots(5))
    Call EnqueueCharacterSave(job)
    UserList(1).SaveSession = snapshots(2).SaveSession + 1
    UserList(1).Persist.LastSkills(1) = 25
    Call AwaitJob(job)
    Call Check(UserList(1).Persist.LastSkills(1) = 25, "reused player slot ignores stale completion")
    Call job.Committed()
    Call Check(UserList(1).Persist.LastSkills(1) = 25, "duplicate completion harmless")
    Call ReleaseCharacterSaveSession(snapshots(2))
End Sub

Private Sub TestCapacityAndShutdown()
    Call InitializeCharacterSaves(connectionString)
    Dim reason As String, oldMax As Integer
    Dim auto As clsCharacterSave, final As clsCharacterSave
    oldMax = MaxUsers
    MaxUsers = 1
    snapshots(2) = Reserve(901)
    snapshots(3).Id = 902
    Call Check(Not ReserveCharacterSaveSession(snapshots(3), reason), "capacity rejects admission before final slot is needed")
    Set auto = NewJob(snapshots(2), False, "UPDATE saved SET value=12 WHERE id=901")
    Call EnqueueCharacterSave(auto)
    Call Check(Not CanCaptureCharacterSave(snapshots(2), False) And CanCaptureCharacterSave(snapshots(2), True), "autosave deferred; final capacity reserved")
    Set final = NewJob(snapshots(2), True, "UPDATE saved SET value=13 WHERE id=901")
    Call EnqueueCharacterSave(final)
    Call Check(Not CanCaptureCharacterSave(snapshots(2), False) And Not CanCaptureCharacterSave(snapshots(2), True), "no save can follow final snapshot")
    Call StopCharacterSaveAdmission()
    Call ReleaseCharacterSaveSession(snapshots(2))
    Call Check(CharacterSavesUnresolved(), "shutdown retains pending work")
    Call AwaitJob(final)
    Call Check(Not CharacterSavesUnresolved(), "shutdown drain observes completed writes")
    Call Check(Not ReserveCharacterSaveSession(snapshots(3), reason), "shutdown continues refusing admission")
    MaxUsers = oldMax
End Sub

Private Sub TestProductionSnapshot()
    ' Clone only this UNIT_TEST fixture. Production files/schemas are never changed.
    Dim clone As String, id As Long, accountId As Long, reason As String
    Dim creation As clsCharacterSave, auto As clsCharacterSave, deferred As clsCharacterSave, final As clsCharacterSave
    clone = App.Path & "\character-save-snapshot-unit.db"
    If Len(Dir$(clone)) Then Kill clone
    Call Connection.Execute("VACUUM INTO '" & Replace(clone, "'", "''") & "'")
    Call db.Close()
    Call db.Open("DRIVER={SQLite3 ODBC Driver};DATABASE=" & clone)
    Call InitializeCharacterSaves("DRIVER={SQLite3 ODBC Driver};DATABASE=" & clone)
    Set originalDB = Connection
    Set Connection = db
    Call db.Execute("INSERT INTO account(email,password,salt,validate_code) VALUES('save-test@invalid','test','test','test')")
    accountId = Value("SELECT last_insert_rowid()")
    UserList(1) = snapshots(6)
    With UserList(1)
        .AccountID = accountId
        .name = "save_completion_fixture"
        .clase = 1
        .raza = 1
        .genero = 1
        .Stats.ELV = 1
        .Stats.MaxHp = 10
        .Stats.MinHp = 10
        .pos.Map = 1
        .pos.x = 50
        .pos.y = 50
        .Char.body = 1
        .Char.head = 1
    End With
    Call SaveNewCharacterDB(1)
    id = UserList(1).Id
    Set creation = CharacterSaveTestJob(id)
    Call Check(CharacterHasUnfinishedSave(id), "creation child rows participate in ordering")
    Call AwaitJob(creation)
    Call Check(creation.Result = CharacterSaveSucceeded, "native creation writes finish")
    Call InitUserPersistSnapshot(1)
    With UserList(1)
        .flags.UserLogged = True
        .Stats.GLD = 100
        .invent.Object(1).ObjIndex = 1
        .invent.Object(1).amount = 10
        .BancoInvent.Object(1).ObjIndex = 1
        .BancoInvent.Object(1).amount = 5
    End With
    Call SaveChangesInUser(1)
    Set auto = CharacterSaveTestJob(id)
    With UserList(1)
        .Stats.GLD = 200
        .invent.Object(1).amount = 9
        .BancoInvent.Object(1).ObjIndex = 0
        .BancoInvent.Object(1).amount = 0
        .Stats.UserSkills(1) = 25
        .Stats.SkillDirty(1) = True
        .QuestStats.Quests(1).Dirty = True
        .SaveResetEpoch = 1
        .SavePublicationPrice = 12345
    End With
    Call SaveCharacterDB(1)
    Call AwaitJob(auto)
    Set deferred = CharacterSaveTestJob(id)
    Call Check(Not (auto Is deferred) And CharacterHasUnfinishedSave(id), "reset deferred behind autosave captures follow-up")
    UserList(1).SaveResetEpoch = 2
    Call SaveCharacterDB(1, True)
    Set final = CharacterSaveTestJob(id, True)
    Call ReleaseCharacterSaveSession(UserList(1))
    UserList(1) = snapshots(6)
    UserList(1).Id = id
    UserList(1).AccountID = accountId
    Call Check(Not LoadCharacterFromDB(1, reason) And Len(reason) > 0, "real loader refuses pending previous session")
    Call AwaitJob(final)
    Call Check(auto.Result = CharacterSaveSucceeded And deferred.Result = CharacterSaveSucceeded And final.Result = CharacterSaveSucceeded, "full production snapshot completes")
    Call Check(Value("SELECT gold FROM user WHERE id=" & id) = 200, "final main character values persisted")
    Call Check(Value("SELECT amount FROM inventory_item WHERE user_id=" & id & " AND number=1") = 9, "final inventory persisted")
    Call Check(Value("SELECT amount FROM bank_item WHERE user_id=" & id & " AND number=1") = 0, "final bank revert beats unacknowledged autosave")
    Call Check(Value("SELECT value FROM skillpoint WHERE user_id=" & id & " AND number=1") = 25, "final dirty skills persisted")
    Call Check(Value("SELECT count(*) FROM user WHERE id=" & id & " AND last_logout>0") = 1, "logout timestamp included before completion")
    Call Check(Value("SELECT is_reset FROM user WHERE id=" & id) = 1, "reset side write included in completion")
    Call Check(Value("SELECT price_in_mao FROM user WHERE id=" & id & " AND is_locked_in_mao=1") = 12345, "publication side write included in completion")
    Call Check(Not CharacterHasUnfinishedSave(id), "completed final releases loader guard")
    Set Connection = originalDB
    Set originalDB = Nothing
End Sub

Private Sub TestCallbacks()
    Call InitializeCharacterSaves(connectionString)
    snapshots(2) = Reserve(901)
    Dim worker As New clsCharacterSaveWorker, job As New clsCharacterSave
    Dim oldCommand As ADODB.Command, unrelated As New ADODB.Command
    Call worker.OpenConnection(connectionString)
    worker.TestSynchronous = True
    Call job.Initialize(0, snapshots(2), False)
    Call job.Add("UPDATE saved SET value=14 WHERE id=901")
    Call job.Add("UPDATE saved SET value=15 WHERE id=901")
    Call job.Seal(snapshots(2))
    Set oldCommand = job.Statement(1)
    Set unrelated = job.Statement(1)
    Call Check(Not (oldCommand Is unrelated), "each dispatch has a fresh command identity")
    oldCommand.CommandText = "UPDATE saved SET value=999 WHERE id=901"
    Call Check(unrelated.CommandText = "UPDATE saved SET value=14 WHERE id=901", "dispatch cannot mutate sealed snapshot")
    Call job.MarkRunning()
    Call worker.Start(job)
    Call worker.Tick()
    Set oldCommand = worker.TestCommand
    Call Check(worker.TestCallbackDuringSubmit And job.Result = CharacterSaveRunning, "callback during submission recorded before advancement")
    Call worker.TestLateCompletion(unrelated)
    Call worker.Tick()
    Call Check(job.NextStatement = 2, "one statement completion advances exactly once")
    Call worker.TestLateCompletion(oldCommand)
    Call worker.Tick()
    Call Check(job.Result = CharacterSaveSucceeded And Value("SELECT value FROM saved WHERE id=901") = 15, "old callback cannot complete a replacement command")
    Call worker.TestLateCompletion(oldCommand)
    Call Check(worker.Idle, "duplicate late callback after completion ignored")
    Set job = NewJob(snapshots(2), True, "UPDATE saved SET value=16 WHERE id=901")
    Call job.MarkRunning()
    worker.TestDispatchError = True
    Call worker.Start(job)
    Call worker.Tick()
    Call Check(job.Result = CharacterSaveFailed And Not job.Retryable And Not worker.Idle, "uncertain dispatch quarantines exact connection and snapshot")
    Call worker.TestLateCompletion(oldCommand)
    Call Check(Value("SELECT value FROM saved WHERE id=901") = 15, "uncertain failure never reexecutes on late callback")
    Call ReleaseCharacterSaveSession(snapshots(2))
End Sub
#End If
