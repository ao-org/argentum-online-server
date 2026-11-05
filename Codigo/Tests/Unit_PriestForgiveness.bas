Attribute VB_Name = "Unit_PriestForgiveness"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private requestBytes() As Byte

Public Function CapturePacket(ByVal Route As SendTarget, ByVal Recipient As Integer, _
        ByRef Queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    Dim bytes() As Byte, writer As Network.Writer, value As Variant
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    value = bytes
    Call packets.Add(value)
    Queued = True
    CapturePacket = True
End Function

Public Function TestSuite() As Boolean
    TestSuite = True
    Dim oldMaxUsers As Integer
    oldMaxUsers = MaxUsers
    MaxUsers = 4
    Call Unit_NpcBow.InitializeTestWorld(1)
    Dim price As Long
    Call UnitTesting.RunTest("pardon price without citizen kills", _
        CalculateHooForgivenessPrice(0, price) And _
        price = CLng(CDbl(SvrConfig.GetValue("CostoPerdonPorCiudadano")) / 2))
    Call UnitTesting.RunTest("pardon price includes citizen kills and multiplier", _
        CalculateHooForgivenessPrice(2, price) And _
        price = 2 * SvrConfig.GetValue("GoldMult") * SvrConfig.GetValue("CostoPerdonPorCiudadano"))
    Call UnitTesting.RunTest("pardon price rejects Long overflow", _
        Not CalculateHooForgivenessPriceForValues(2, 2, 2147483647#, price))
    Call UnitTesting.RunTest("pardon price rejects negative citizen-kill state", _
        Not CalculateHooForgivenessPriceForValues(-1, 2, 1000, price))
    Call UnitTesting.RunTest("forgiveness offer expires across tick wrap", _
        HooForgivenessOfferExpired(2147470000, -2147460000) And _
        Not HooForgivenessOfferExpired(100, 30100))
    Call UnitTesting.RunTest("forgiveness requires generic HOO protocol v2", _
        TestProtocolGate())
    Call UnitTesting.RunTest("eligible criminal and priest can request pardon", _
        TestEligibility(eHooForgivenessOffer))
    Call UnitTesting.RunTest("citizen cannot request pardon", _
        TestFactionRejection(e_Facciones.Ciudadano, eHooForgivenessAlreadyCitizen))
    Call UnitTesting.RunTest("chaos faction cannot request pardon", _
        TestFactionRejection(e_Facciones.Caos, eHooForgivenessFaction))
    Call UnitTesting.RunTest("concilio faction cannot request pardon", _
        TestFactionRejection(e_Facciones.concilio, eHooForgivenessFaction))
    Call UnitTesting.RunTest("dead character cannot request pardon", TestDeadRejection())
    Call UnitTesting.RunTest("distant character cannot request pardon", TestDistanceRejection())
    Call UnitTesting.RunTest("invalid priest cannot grant pardon", TestInvalidPriest())
    Call UnitTesting.RunTest("newbie priest restriction follows player level", TestNewbiePriest())
    Call UnitTesting.RunTest("criminal and chaotic guilds cannot request pardon", TestGuildRejection())
    Call UnitTesting.RunTest("map trigger blocks pardon before payment", TestRestrictedLocation())
    Call UnitTesting.RunTest("offer confirmation charges the exact quote once", TestConfirmOnce())
    Call UnitTesting.RunTest("changed price requires a new confirmation", TestChangedPrice())
    Call UnitTesting.RunTest("replaced priest invalidates the offer", TestReplacedPriest())
    Call UnitTesting.RunTest("range change invalidates the offer without charging", TestRangeChange())
    MaxUsers = oldMaxUsers
End Function

Private Sub SetRequest(ByVal RequestId As Long, ByVal action As Byte, _
        ByVal CharacterIndex As Integer, ByVal OfferToken As Long)
    Dim writer As New Network.Writer, bytes() As Byte, i As Long
    Call writer.WriteInt32(RequestId)
    Call writer.WriteInt8(action)
    Call writer.WriteInt16(CharacterIndex)
    Call writer.WriteInt32(OfferToken)
    Call writer.GetData(bytes)
    ReDim requestBytes(0 To UBound(bytes))
    For i = 0 To UBound(bytes)
        requestBytes(i) = bytes(i)
    Next i
    Set reader = New Network.reader
    Call reader.SetData(requestBytes)
End Sub

Private Function LastResult(ByRef RequestId As Long, ByRef Result As e_HooForgivenessResult, _
        ByRef Price As Long, ByRef OfferToken As Long) As Boolean
    Dim i As Long, bytes() As Byte, packetReader As Network.reader
    For i = packets.Count To 1 Step -1
        bytes = packets(i)
        Set packetReader = New Network.reader
        Call packetReader.SetData(bytes)
        If packetReader.ReadInt16() = ServerPacketID.eHooPriestForgivenessResult Then
            RequestId = packetReader.ReadInt32()
            Result = packetReader.ReadInt8()
            Price = packetReader.ReadInt32()
            OfferToken = packetReader.ReadInt32()
            LastResult = True
            Exit Function
        End If
    Next i
End Function

Private Sub SetupRequestFixture()
    Call Unit_NpcBow.InitializeTestWorld(1)
    MaxUsers = 4
    Set packets = New Collection
    capturing = True
    Call SetupEligibilityFixture()
    With UserList(1)
        .Stats.GLD = 100000
        .Faccion.ciudadanosMatados = 0
        .HooCapabilities.Negotiated = True
        .HooCapabilities.ProtocolVersion = HOO_PROTOCOL_V2_VERSION
    End With
End Sub

Private Function TestConfirmOnce() As Boolean
    On Error GoTo TestDone
    Call SetupRequestFixture()
    Dim requestId As Long, result As e_HooForgivenessResult, price As Long
    Dim token As Long, startingGold As Long
    startingGold = UserList(1).Stats.GLD
    Call SetRequest(41, 0, NpcList(1).Char.charindex, 0)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    If requestId <> 41 Or result <> eHooForgivenessOffer Or token <= 0 Then GoTo TestDone
    Dim offerToken As Long
    offerToken = token
    Call SetRequest(41, 1, NpcList(1).Char.charindex, offerToken)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    If result <> eHooForgivenessSuccess Or UserList(1).Stats.GLD <> startingGold - price Or _
            UserList(1).Faccion.Status <> e_Facciones.Ciudadano Then GoTo TestDone
    Dim afterSuccess As Long
    afterSuccess = UserList(1).Stats.GLD
    Call SetRequest(41, 1, NpcList(1).Char.charindex, offerToken)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    TestConfirmOnce = result = eHooForgivenessInvalidOffer And _
        UserList(1).Stats.GLD = afterSuccess
TestDone:
    If Err.Number <> 0 Then Call UnitTesting.Progress("priest confirm-once error " & CStr(Err.Number) & ": " & Err.Description)
    capturing = False
End Function

Private Function TestChangedPrice() As Boolean
    On Error GoTo TestDone
    Call SetupRequestFixture()
    Dim requestId As Long, result As e_HooForgivenessResult, price As Long
    Dim token As Long, startingGold As Long, oldToken As Long
    startingGold = UserList(1).Stats.GLD
    Call SetRequest(42, 0, NpcList(1).Char.charindex, 0)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    oldToken = token
    UserList(1).Faccion.ciudadanosMatados = 1
    Call SetRequest(42, 1, NpcList(1).Char.charindex, oldToken)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    If result <> eHooForgivenessPriceChanged Or token = oldToken Or _
            UserList(1).Stats.GLD <> startingGold Then GoTo TestDone
    Call SetRequest(42, 1, NpcList(1).Char.charindex, token)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    TestChangedPrice = result = eHooForgivenessSuccess And _
        UserList(1).Stats.GLD = startingGold - price
TestDone:
    If Err.Number <> 0 Then Call UnitTesting.Progress("priest changed-price error " & CStr(Err.Number) & ": " & Err.Description)
    capturing = False
End Function

Private Function TestReplacedPriest() As Boolean
    On Error GoTo TestDone
    Call SetupRequestFixture()
    Dim requestId As Long, result As e_HooForgivenessResult, price As Long, token As Long
    Call SetRequest(43, 0, NpcList(1).Char.charindex, 0)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    NpcList(1).VersionId = NpcList(1).VersionId + 1
    Call SetRequest(43, 1, NpcList(1).Char.charindex, token)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    TestReplacedPriest = result = eHooForgivenessInvalidPriest And _
        UserList(1).Stats.GLD = 100000 And UserList(1).Faccion.Status = e_Facciones.Criminal
TestDone:
    If Err.Number <> 0 Then Call UnitTesting.Progress("priest replaced-priest error " & CStr(Err.Number) & ": " & Err.Description)
    capturing = False
End Function

Private Function TestRangeChange() As Boolean
    On Error GoTo TestDone
    Call SetupRequestFixture()
    Dim requestId As Long, result As e_HooForgivenessResult, price As Long, token As Long
    Call SetRequest(44, 0, NpcList(1).Char.charindex, 0)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    UserList(1).pos.x = 55
    Call SetRequest(44, 1, NpcList(1).Char.charindex, token)
    Call HandleHooPriestForgiveness(1)
    If Not LastResult(requestId, result, price, token) Then GoTo TestDone
    TestRangeChange = result = eHooForgivenessTooFar And _
        UserList(1).Stats.GLD = 100000 And UserList(1).Faccion.Status = e_Facciones.Criminal
TestDone:
    If Err.Number <> 0 Then Call UnitTesting.Progress("priest range-change error " & CStr(Err.Number) & ": " & Err.Description)
    capturing = False
End Function

Private Sub SetupEligibilityFixture()
    Dim emptyUser As t_User, emptyNpc As t_Npc
    UserList(1) = emptyUser
    NpcList(1) = emptyNpc
    With UserList(1)
        .Id = 10: .VersionId = 7: .InUse = True
        .flags.UserLogged = True: .Faccion.Status = e_Facciones.Criminal
        .pos.Map = 1: .pos.x = 50: .pos.y = 50
        .Stats.ELV = LimiteNewbie + 1
    End With
    With NpcList(1)
        .flags.NPCActive = True: .VersionId = 8: .Char.charindex = 2
        .npcType = e_NPCType.Revividor: .Numero = 11
        .pos.Map = 1: .pos.x = 51: .pos.y = 50
    End With
    Call SetNpcRef(UserList(1).flags.TargetNPC, 1)
End Sub

Private Function TestEligibility(ByVal expected As e_HooForgivenessResult) As Boolean
    Call SetupEligibilityFixture()
    Dim price As Long
    TestEligibility = GetHooForgivenessEligibility(1, _
        UserList(1).flags.TargetNPC, price) = expected And price > 0
End Function

Private Function TestFactionRejection(ByVal faction As e_Facciones, _
        ByVal expected As e_HooForgivenessResult) As Boolean
    Call SetupEligibilityFixture()
    UserList(1).Faccion.Status = faction
    Dim price As Long
    TestFactionRejection = GetHooForgivenessEligibility(1, _
        UserList(1).flags.TargetNPC, price) = expected
End Function

Private Function TestDeadRejection() As Boolean
    Call SetupEligibilityFixture()
    UserList(1).flags.Muerto = 1
    Dim price As Long
    TestDeadRejection = GetHooForgivenessEligibility(1, _
        UserList(1).flags.TargetNPC, price) = eHooForgivenessDead
End Function

Private Function TestDistanceRejection() As Boolean
    Call SetupEligibilityFixture()
    UserList(1).pos.x = 55
    Dim price As Long
    TestDistanceRejection = GetHooForgivenessEligibility(1, _
        UserList(1).flags.TargetNPC, price) = eHooForgivenessTooFar
End Function

Private Function TestInvalidPriest() As Boolean
    Call SetupEligibilityFixture()
    NpcList(1).npcType = Comun
    Dim price As Long
    TestInvalidPriest = GetHooForgivenessEligibility(1, _
        UserList(1).flags.TargetNPC, price) = eHooForgivenessInvalidPriest
End Function

Private Function TestNewbiePriest() As Boolean
    Call SetupEligibilityFixture()
    NpcList(1).npcType = e_NPCType.ResucitadorNewbie
    Dim price As Long
    If GetHooForgivenessEligibility(1, UserList(1).flags.TargetNPC, price) <> _
            eHooForgivenessInvalidPriest Then Exit Function
    UserList(1).Stats.ELV = 1
    TestNewbiePriest = GetHooForgivenessEligibility(1, _
        UserList(1).flags.TargetNPC, price) = eHooForgivenessOffer
End Function

Private Function TestGuildRejection() As Boolean
    On Error GoTo TestDone
    Call SetupEligibilityFixture()
    Dim originalCount As Integer, originalGuild As clsClan, testGuild As clsClan
    originalCount = CANTIDADDECLANES
    CANTIDADDECLANES = 1
    Set testGuild = New clsClan
    Call testGuild.SetTestAlignment(e_ALINEACION_GUILD.ALINEACION_CRIMINAL)
    Set originalGuild = modGuilds.SwapTestGuild(1, testGuild)
    UserList(1).GuildIndex = 1
    Dim price As Long
    If GetHooForgivenessEligibility(1, UserList(1).flags.TargetNPC, price) <> eHooForgivenessGuild Then GoTo TestDone
    Call testGuild.SetTestAlignment(e_ALINEACION_GUILD.ALINEACION_CAOTICA)
    TestGuildRejection = GetHooForgivenessEligibility(1, _
        UserList(1).flags.TargetNPC, price) = eHooForgivenessGuild
TestDone:
    If Not originalGuild Is Nothing Then Set originalGuild = modGuilds.SwapTestGuild(1, originalGuild)
    CANTIDADDECLANES = originalCount
End Function

Private Function TestRestrictedLocation() As Boolean
    Call SetupEligibilityFixture()
    MapData(50, 50, 1).trigger = 6
    Dim price As Long
    TestRestrictedLocation = GetHooForgivenessEligibility(1, _
        UserList(1).flags.TargetNPC, price) = eHooForgivenessRestrictedLocation
End Function

Private Function TestProtocolGate() As Boolean
    Dim wasNegotiated As Boolean, oldVersion As Byte, oldMask As Long
    wasNegotiated = UserList(1).HooCapabilities.Negotiated
    oldVersion = UserList(1).HooCapabilities.ProtocolVersion
    oldMask = UserList(1).HooCapabilities.CapabilityMask
    UserList(1).HooCapabilities.Negotiated = True
    UserList(1).HooCapabilities.ProtocolVersion = HOO_PROTOCOL_V2_VERSION
    TestProtocolGate = UserSupportsPriestForgiveness(1)
    UserList(1).HooCapabilities.ProtocolVersion = HOO_CAP_PROTOCOL_VERSION
    If UserSupportsPriestForgiveness(1) Then TestProtocolGate = False
    UserList(1).HooCapabilities.Negotiated = wasNegotiated
    UserList(1).HooCapabilities.ProtocolVersion = oldVersion
    UserList(1).HooCapabilities.CapabilityMask = oldMask
End Function

#End If
