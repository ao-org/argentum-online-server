Attribute VB_Name = "Unit_Housing"
Option Explicit
#If UNIT_TEST = 1 Then
Private fixtureConnection As ADODB.Connection
Private fixturePath As String
Private capturing As Boolean
Private clearPackets As Long
Private lastPacket() As Byte

Public Function CaptureHousingPacket() As Boolean
    If Not capturing Then Exit Function
    Dim packetWriter As Network.writer
    Set packetWriter = Protocol_Writes.GetWriterBuffer()
    Call packetWriter.GetData(lastPacket)
    Call packetWriter.Clear()
    clearPackets = clearPackets + 1
    CaptureHousingPacket = True
End Function

Private Sub Expect(ByVal value As Boolean, ByVal reason As String)
    If Not value Then Err.Raise vbObjectError + 730, "Housing test", reason & ": " & HousingTestLastError
End Sub

Private Sub Sql(ByVal text As String)
    Call fixtureConnection.Execute(text)
End Sub

Private Function Scalar(ByVal text As String) As Variant
    Dim rs As ADODB.Recordset
    Set rs = fixtureConnection.Execute(text)
    Scalar = rs.Fields(0).Value
    Call rs.Close()
End Function

Private Sub Fixture()
    Dim fs As New Scripting.FileSystemObject, k As Long, o As Long
    fixturePath = fs.BuildPath(fs.GetSpecialFolder(2), fs.GetTempName()) & ".sqlite"
    Set fixtureConnection = New ADODB.Connection
    fixtureConnection.ConnectionString = "DRIVER={SQLite3 ODBC Driver};Database=" & fixturePath & ";Timeout=1;"
    Call fixtureConnection.Open()
    Call Sql("PRAGMA busy_timeout=1")
    Call Sql("CREATE TABLE account(id INTEGER PRIMARY KEY,is_active_patron INTEGER)")
    Call Sql("CREATE TABLE house_key(key_obj INTEGER PRIMARY KEY,account_id INTEGER,description TEXT,assigned_at TEXT)")
    Call Sql("CREATE TABLE castle(inside_key_obj_id INTEGER)")
    Call Sql("INSERT INTO account VALUES(1,6057394),(2,6057395),(3,0)")
    Call Sql("INSERT INTO house_key VALUES(10,NULL,'House A',NULL),(20,NULL,'House B',NULL),(30,1,'Castle',NULL)")
    Call Sql("INSERT INTO castle VALUES(30)")
    ReDim UserList(1 To 3)
    ReDim ObjData(0 To 40)
    ReDim MapInfo(1 To 1)
    ReDim MapData(1 To 100, 1 To 100, 1 To 1)
    NumMaps = 1: MapInfo(1).MapResource = 1
    MinXBorder = 1: MaxXBorder = 100: MinYBorder = 1: MaxYBorder = 100
    For k = 10 To 20 Step 10
        ObjData(k).OBJType = otKeys: ObjData(k).clave = k
        For o = k + 1 To k + 3
            With ObjData(o)
                .OBJType = otDoors: .clave = k: .Subtipo = 0
                .IndexAbierta = k + 1: .IndexCerrada = k + 2: .IndexCerradaLlave = k + 3
                If o <> k + 1 Then .Cerrada = 1
                If o = k + 3 Then .Llave = 1
            End With
        Next
        MapData(k, 20, 1).ObjInfo.ObjIndex = k + 3
    Next
    For k = 1 To 3
        UserList(k).Id = 100 + k: UserList(k).AccountID = 1
        UserList(k).VersionId = k: UserList(k).pos.Map = 1
        UserList(k).pos.x = 50: UserList(k).pos.y = 50
        UserList(k).Keys(1) = 30
        Call HousingResetSession(CInt(k))
    Next
    Call Expect(HousingInitialize("DRIVER={SQLite3 ODBC Driver};Database=" & fixturePath), "initialize synthetic catalogue/doors")
    HousingTestFailAfterRelease = False: HousingTestFailBeforeCommit = False: HousingTestLoseCommitReply = False
    clearPackets = 0: HousingTestLastError = ""
End Sub

Private Sub Cleanup()
    On Error Resume Next
    Call HousingDbClose()
    If Not fixtureConnection Is Nothing Then Call fixtureConnection.Close()
    Set fixtureConnection = Nothing
    ' Only paths created by this fixture, never a configured world database.
    If Len(fixturePath) > 0 Then Kill fixturePath
    If Len(fixturePath) > 0 Then
        Kill fixturePath & ".ready"
        Kill fixturePath & ".release"
        Kill fixturePath & ".done"
    End If
End Sub

Private Function Attempt(ByVal op As e_HousingOperation, ByVal expected As Integer, ByVal target As Integer, ByRef reply As t_HousingReply, Optional ByVal account As Long = 1) As e_HousingCode
    Dim request As t_HousingRequest, blankReply As t_HousingReply
    reply = blankReply
    request.Operation = op: request.ExpectedKey = expected: request.TargetKey = target
    Attempt = HousingDbOperation(account, request, reply)
End Function

Private Sub MatureOwner()
    Call Sql("UPDATE house_key SET account_id=1,assigned_at=datetime('now','-73 hours') WHERE key_obj=10")
    Call Expect(HousingRefreshLogin(1), "refresh mature owner")
End Sub

Private Sub TestCase(ByVal n As Long)
    Dim reply As t_HousingReply, oldTimestamp As String, request As t_HousingRequest, tick As Long, result As e_HousingCode, i As Long
    Select Case n
    Case 1
        Call Expect(Attempt(HousingClaim, 0, 10, reply) = HousingOkay, "claim")
        oldTimestamp = CStr(Scalar("SELECT assigned_at FROM house_key WHERE key_obj=10"))
        Call Expect(Len(oldTimestamp) = 19 And reply.EligibleUTC <> "", "database UTC period")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingMinimumPeriod, "release locked")
        Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingMinimumPeriod, "exchange locked")
        Call Expect(Attempt(HousingCatalogue, 0, 0, reply) = HousingOkay, "status")
        Call Expect(HousingRefreshLogin(1), "reconnect refresh")
        Call Expect(Scalar("SELECT assigned_at FROM house_key WHERE key_obj=10") = oldTimestamp, "reads/failures preserve timestamp")
    Case 2
        Call MatureOwner()
        Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingOkay, "atomic exchange")
        Call Expect(IsNull(Scalar("SELECT account_id FROM house_key WHERE key_obj=10")), "A available")
        Call Expect(Scalar("SELECT account_id FROM house_key WHERE key_obj=20") = 1, "B owned")
        Call Expect(Attempt(HousingRelease, 20, 0, reply) = HousingMinimumPeriod, "new ownership period")
    Case 3, 4
        Call MatureOwner()
        HousingTestFailAfterRelease = (n = 3): HousingTestFailBeforeCommit = (n = 4)
        Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingUnavailable, "failure")
        Call Expect(Scalar("SELECT account_id FROM house_key WHERE key_obj=10") = 1, "rollback A")
        Call Expect(IsNull(Scalar("SELECT account_id FROM house_key WHERE key_obj=20")), "rollback B")
        Call Expect(UserList(1).Keys(2) = 10, "rollback keys")
    Case 5
        Call MatureOwner()
        UserList(2).Keys(3) = 10: UserList(2).Keys(4) = 10
        UserList(1).flags.UserLogged = True
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingOkay, "release")
        Call Expect(UserList(1).Keys(1) = 30 And UserList(2).Keys(1) = 30, "castle retained")
        Call Expect(UserList(2).Keys(2) = 0 And UserList(2).Keys(3) = 0 And UserList(2).Keys(4) = 0, "all duplicate keys cleared including retained user")
        Call Expect(clearPackets >= MAXKEYS, "blankReply slots transmitted")
        Call Expect(Attempt(HousingClaim, 0, 20, reply) = HousingOkay, "claim after lawful release has no account cooldown")
    Case 6
        Call MatureOwner()
        UserList(3).AccountID = 2: UserList(3).pos.x = 10: UserList(3).pos.y = 21
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingOccupied, "door threshold")
    Case 7
        Call MatureOwner()
        UserList(1).pos.x = 10: UserList(1).pos.y = 19
        UserList(1).AccountID = 2
        Call Expect(Not HousingDoorAccess(1, 13), "stale key denied")
        Call Expect(HousingDoorAccess(2, 13), "current owner allowed")
        Call Expect(HousingDoorAccess(1, 31), "unrelated key unaffected")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingOkay, "interior guest does not block release")
        Call Expect(UserList(1).pos.x = 10 And UserList(1).pos.y = 19, "no interior relocation")
    Case 8
        Call Expect(Attempt(HousingClaim, 0, 10, reply, 3) = HousingIneligible, "acquisition entitlement")
        Call MatureOwner()
        Call Sql("UPDATE account SET is_active_patron=0 WHERE id=1")
        Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingIneligible, "expired exchange")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingOkay, "expired release allowed")
    Case 9
        Call MatureOwner()
        Call Sql("UPDATE house_key SET assigned_at=datetime('now','-72 hours') WHERE key_obj=10")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingOkay, "at exact 72-hour boundary")
        Call Expect(CLng(Scalar("SELECT julianday('2026-01-04 00:00:00')>=julianday('2026-01-01 00:00:00','+72 hours')")) = 1, "exact comparison inclusive")
        Call Expect(CLng(Scalar("SELECT julianday('2026-01-03 23:59:59')>=julianday('2026-01-01 00:00:00','+72 hours')")) = 0, "one second early")
    Case 10
        Call MatureOwner()
        Call Sql("UPDATE house_key SET assigned_at='invalid' WHERE key_obj=10")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingInconsistent, "invalid timestamp is not repaired")
        Call Sql("UPDATE house_key SET assigned_at=0 WHERE key_obj=10")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingInconsistent, "numeric date is not an acquisition timestamp")
        Call Expect(reply.AssignedUTC = "" And reply.Current.Description = "", "malformed stored fields never enter the reply")
    Case 11
        Call MatureOwner()
        Call Sql("UPDATE house_key SET account_id=1,assigned_at=CURRENT_TIMESTAMP WHERE key_obj=20")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingInconsistent, "multiple assignments rejected")
    Case 12
        Call MatureOwner()
        Call Sql("BEGIN IMMEDIATE")
        tick = GetTickCountRaw()
        Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingBusy, "independent writer lock")
        Call Expect(CDbl(GetTickCountRaw()) - CDbl(tick) < 1000, "native driver contention bounded below one second")
        Call Sql("ROLLBACK")
        Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingOkay, "connection reusable")
    Case 13
        UserList(1).flags.UserLogged = True
        request.Operation = HousingClaim: request.TargetKey = 10: request.RequestId = 1
        Call Expect(HousingProcessRequest(1, request, 1000, reply) = HousingOkay, "request claim")
        oldTimestamp = CStr(Scalar("SELECT assigned_at FROM house_key WHERE key_obj=10"))
        Call Expect(HousingProcessRequest(1, request, 2000, reply) = HousingOkay, "duplicate receipt")
        request.TargetKey = 20
        Call Expect(HousingProcessRequest(1, request, 3000, reply) = HousingInvalid, "same ID different payload")
        Call Expect(Scalar("SELECT assigned_at FROM house_key WHERE key_obj=10") = oldTimestamp, "duplicate no timestamp refresh")
        Call HousingResetSession(1)
        UserList(1).AccountID = 2: UserList(1).VersionId = 9
        Call Expect(HousingProcessRequest(1, request, 4000, reply) = HousingOkay, "slot reuse new account")
    Case 14
        Call MatureOwner()
        HousingTestLoseCommitReply = True
        Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingUnresolved, "lost completion not rollback")
        Call Expect(Scalar("SELECT account_id FROM house_key WHERE key_obj=20") = 1, "committed visible")
        Call Expect(UserList(1).Keys(2) = 20 And Not HousingDoorAccess(1, 13), "authoritative reconciliation")
        HousingTestLoseCommitReply = False
        Call Expect(Attempt(HousingCatalogue, 0, 0, reply) = HousingOkay, "status resolves")
    Case 15
        Call MatureOwner()
        Call Expect(Attempt(HousingExchange, 20, 10, reply) = HousingStale, "expected ownership")
        Call Expect(Attempt(HousingExchange, 10, 30, reply) = HousingUnavailable, "castle not ordinary")
    Case 16
        For i = 1 To MAXKEYS
            UserList(2).Keys(i) = 30
        Next
        Call Expect(Attempt(HousingClaim, 0, 10, reply) = HousingUnavailable, "retained character key capacity")
        Call Expect(IsNull(Scalar("SELECT account_id FROM house_key WHERE key_obj=10")), "capacity failure writes nothing")
    Case 17
        Call MatureOwner()
        MapData(10, 20, 1).ObjInfo.ObjIndex = 11
        MapData(30, 20, 1).ObjInfo.ObjIndex = 11
        Call Expect(HousingInitialize("DRIVER={SQLite3 ODBC Driver};Database=" & fixturePath), "discover multiple entrances without external configuration")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingOkay, "release open house")
        Call Expect(MapData(10, 20, 1).ObjInfo.ObjIndex = 13, "closed locked variant")
        Call Expect(MapData(30, 20, 1).ObjInfo.ObjIndex = 13, "second discovered entrance also locked")
        Call Expect((MapData(10, 20, 1).Blocked And e_Block.NORTH) <> 0, "door collision")
        Call Expect((MapData(9, 21, 1).Blocked And e_Block.SOUTH) <> 0, "paired threshold collision")
    Case 18
        Call MatureOwner()
        oldTimestamp = CStr(Scalar("SELECT assigned_at FROM house_key WHERE key_obj=10"))
        Call Sql("UPDATE house_key SET account_id=2,assigned_at=CURRENT_TIMESTAMP WHERE key_obj=20")
        Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingStale, "competing target owner")
        Call Expect(Scalar("SELECT account_id FROM house_key WHERE key_obj=10") = 1, "A retained")
        Call Expect(Scalar("SELECT assigned_at FROM house_key WHERE key_obj=10") = oldTimestamp, "A period unchanged")
    Case 19
        request.Operation = HousingCatalogue: request.Search = "House"
        Call Expect(HousingDbOperation(1, request, reply) = HousingOkay And reply.Count = 2, "parameterized catalogue")
        request.Search = "%' OR 1=1 --": reply.Count = 0
        Call Expect(HousingDbOperation(1, request, reply) = HousingOkay And reply.Count = 0, "literal wildcard and SQL syntax")
    Case 20
        UserList(1).flags.UserLogged = True
        request.Operation = HousingClaim: request.TargetKey = 10: request.RequestId = 1
        HousingTestLoseCommitReply = True
        Call Expect(HousingProcessRequest(1, request, 1000, reply) = HousingUnresolved, "uncertain request")
        HousingTestLoseCommitReply = False
        request.RequestId = 2: request.ExpectedKey = 10: request.Operation = HousingRelease
        Call Expect(HousingProcessRequest(1, request, 2000, reply) = HousingUnresolved, "status required")
        Call Sql("BEGIN IMMEDIATE")
        request.RequestId = 3: request.Operation = HousingCatalogue
        Call Expect(HousingProcessRequest(1, request, 3000, reply) = HousingBusy, "status cannot read while locked")
        Call Sql("ROLLBACK")
        request.RequestId = 4: request.Operation = HousingRelease
        Call Expect(HousingProcessRequest(1, request, 4000, reply) = HousingUnresolved, "failed status does not clear uncertainty")
        request.RequestId = 5: request.Operation = HousingCatalogue
        Call Expect(HousingProcessRequest(1, request, 5000, reply) = HousingOkay, "status resolves receipt")
        request.RequestId = 6: request.Operation = HousingRelease
        Call Expect(HousingProcessRequest(1, request, 6000, reply) = HousingMinimumPeriod, "no second write after uncertain claim")
    Case 21
        Call Expect(Attempt(HousingClaim, 0, 10, reply) = HousingOkay, "claim before restart")
        oldTimestamp = CStr(Scalar("SELECT assigned_at FROM house_key WHERE key_obj=10"))
        Call HousingDbClose()
        Call Expect(HousingInitialize("DRIVER={SQLite3 ODBC Driver};Database=" & fixturePath), "new connection")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingMinimumPeriod, "restart retains period")
        Call Expect(Scalar("SELECT assigned_at FROM house_key WHERE key_obj=10") = oldTimestamp, "restart timestamp stable")
    Case 23, 24
        If n = 23 Then Call MatureOwner()
        Call Sql("CREATE TABLE housing_test_fixture(marker TEXT)")
        Call Sql("INSERT INTO housing_test_fixture VALUES('disposable-native-housing')")
        Call Sql("ALTER TABLE account ADD COLUMN email TEXT")
        Call Sql("UPDATE account SET email='ONE' WHERE id=1")
        Call Sql("UPDATE account SET email='TWO' WHERE id=2")
        Call Sql("CREATE TABLE tokens(decrypted_token TEXT,username TEXT)")
        Call Sql("INSERT INTO tokens VALUES('" & String$(64, "1") & "','ONE'),('" & String$(64, "2") & "','TWO')")
        Dim peerAccount As Long, target As Long, command As String, process As Object, shell As Object
        peerAccount = 2: target = 20
        If n = 24 Then peerAccount = 1: target = 10
        command = Chr$(34) & Environ$("HOUSING_TEST_PYTHON") & Chr$(34) & " " & Chr$(34) & Environ$("HOUSING_TEST_PYMMO_PEER") & Chr$(34) & " " & Chr$(34) & fixturePath & Chr$(34) & " " & CStr(peerAccount) & " " & CStr(target) & " " & Chr$(34) & fixturePath & ".gate" & Chr$(34)
        Set shell = CreateObject("WScript.Shell")
        Set process = shell.Exec(command)
        Call WaitGate(fixturePath & ".ready")
        If n = 23 Then
            Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingBusy, "VB6 exchange against real PyMMO transaction")
        Else
            Call Expect(Attempt(HousingClaim, 0, 20, reply) = HousingBusy, "same account claim against real PyMMO")
        End If
        Dim gateFile As Integer
        gateFile = FreeFile
        Open fixturePath & ".release" For Output As #gateFile
        Print #gateFile, "release"
        Close #gateFile
        Call WaitGate(fixturePath & ".done")
        If n = 23 Then
            Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingStale, "PyMMO owns B; A retained")
            Call Expect(Scalar("SELECT account_id FROM house_key WHERE key_obj=10") = 1, "atomic loser retains A")
        Else
            Call Expect(Attempt(HousingClaim, 0, 20, reply) = HousingStale, "same account already owns A")
            Call Expect(Scalar("SELECT count(*) FROM house_key WHERE account_id=1 AND key_obj<>30") = 1, "one ordinary house across servers")
        End If
        Call Expect(Scalar("SELECT account_id FROM house_key WHERE key_obj=" & CStr(target)) = peerAccount, "real PyMMO commit visible to ADO")
    Case 25
        Dim inputWriter As New Network.writer, inputReader As New Network.reader, bytes() As Byte
        UserList(1).flags.UserLogged = True
        Call inputWriter.WriteInt16(ClientPacketID.eHousingRequest)
        Call inputWriter.WriteInt32(123)
        Call inputWriter.WriteInt8(0)
        Call inputWriter.WriteInt16(0)
        Call inputWriter.WriteInt16(0)
        Call inputWriter.WriteInt16(0)
        Call inputWriter.WriteString8("House")
        Call inputWriter.GetData(bytes)
        Call inputReader.SetData(bytes)
        Call Expect(inputReader.ReadInt16() = ClientPacketID.eHousingRequest, "request ID")
        Call HandleHousingRequest(1, inputReader)
        Call inputReader.SetData(lastPacket)
        Call Expect(inputReader.ReadInt16() = ServerPacketID.eHousingResult, "result ID")
        Call Expect(inputReader.ReadInt32() = 123, "request correlation")
        If Len(Environ$("HOUSING_TEST_WIRE_OUTPUT")) > 0 Then
            Dim wireFile As Integer
            wireFile = FreeFile
            Open Environ$("HOUSING_TEST_WIRE_OUTPUT") For Binary Access Write As #wireFile
            Put #wireFile, , lastPacket
            Close #wireFile
        End If
    Case 26
        Call MatureOwner()
        Call Expect(Attempt(HousingExchange, 10, 20, reply) = HousingOkay, "exchange before stale-key routes")
        With UserList(1)
            .pos.x = 10: .pos.y = 19
            .clase = 1: .raza = 1: .genero = 1: .Stats.ELV = 1
            .Keys(2) = 10 ' A stale physical/keyring copy cannot restore permission.
            .flags.TargetObj = 13: .flags.TargetObjMap = 1
            .flags.TargetObjX = 10: .flags.TargetObjY = 20
            .invent.Object(1).ObjIndex = 10: .invent.Object(1).amount = 1
        End With
        Call Expect(ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_UnlockAndOpen, 10, 20) = eHooHouseDoorActionResult_NoAccess, "HOO old key denied")
        IntervaloUserPuedeUsarU = 0: IntervaloUserPuedeUsarClic = 0
        i = clearPackets
        Call UsarLlave(1, 2)
        Call Expect(clearPackets = i + 1, "keyring emitted rejection")
        Dim keyReader As New Network.reader
        Call keyReader.SetData(lastPacket)
        Call Expect(keyReader.ReadInt16() = ServerPacketID.eLocaleMsg, "keyring rejection packet")
        Call Expect(keyReader.ReadInt16() = MSG_NO_LLAVE_SIRVE, "keyring ownership rejection")
        Call Expect(CanUseObject(1, 10) = 0, "physical key otherwise usable")
        i = clearPackets
        Call UseInvItem(1, 1, 0)
        Call Expect(clearPackets = i + 1, "inventory emitted rejection")
        Call keyReader.SetData(lastPacket)
        Call Expect(keyReader.ReadInt16() = ServerPacketID.eLocaleMsg, "inventory rejection packet")
        Call Expect(keyReader.ReadInt16() = MSG_NO_LLAVE_SIRVE, "inventory ownership rejection")
        Call Expect(MapData(10, 20, 1).ObjInfo.ObjIndex = 13, "all routes leave old door locked")
    Case 22
        Call MatureOwner()
        Call Sql("UPDATE house_key SET assigned_at=datetime('now','-71 hours') WHERE key_obj=10")
        Call Expect(Attempt(HousingRelease, 10, 0, reply) = HousingMinimumPeriod, "before boundary")
        Call Expect(UserList(2).Keys(2) = 10, "another character cannot bypass period")
    End Select
End Sub

Public Sub RunHousingTests()
    Call InitializeCircularLogBuffer()
    Call InitializeAuxiliaryBuffer()
    capturing = True
    Dim i As Long
    For i = 1 To 22
        Call RunOne(i)
    Next
    If Len(Environ$("HOUSING_TEST_PYMMO_PEER")) > 0 And Len(Environ$("HOUSING_TEST_PYTHON")) > 0 Then
        Call RunOne(23)
        Call RunOne(24)
    End If
    Call RunOne(25)
    Call RunOne(26)
    capturing = False
End Sub

Private Sub WaitGate(ByVal path As String)
    Dim started As Long
    started = GetTickCountRaw()
    Do While Len(Dir$(path)) = 0
        If CDbl(GetTickCountRaw()) - CDbl(started) > 15000 Then Err.Raise 5, "Housing test", "Peer gate timeout"
        Call General.Sleep(5)
    Loop
End Sub

Private Sub RunOne(ByVal n As Long)
    On Error GoTo Failed
    Call Fixture()
    Call TestCase(n)
    Call UnitTesting.RunTest("housing case " & CStr(n), True)
    Call Cleanup()
    Exit Sub
Failed:
    Call UnitTesting.RunTestError("housing case " & CStr(n), CStr(Err.Number) & ": " & Err.Description)
    Call Cleanup()
End Sub
#End If
