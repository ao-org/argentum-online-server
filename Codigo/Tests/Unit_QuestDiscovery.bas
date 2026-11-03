Attribute VB_Name = "Unit_QuestDiscovery"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private attempts As Long
Private rejectAt As Long
Private fixtureVersion As Integer
Public Function CapturePacket(ByRef queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    Dim bytes() As Byte, writer As Network.Writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    queued = True
    If UBound(bytes) >= 1 Then
        If CLng(bytes(0)) + 256& * bytes(1) = ServerPacketID.eQuestDiscovery Then
            attempts = attempts + 1
            queued = attempts <> rejectAt
            If queued Then
                Dim payload As Variant: payload = bytes
                Call packets.Add(payload)
            End If
        End If
    End If
    CapturePacket = True
End Function
Private Sub Expect(ByVal value As Boolean, ByVal label As String)
    If Not value Then Err.Raise vbObjectError + 849, "Quest discovery", label
End Sub
Private Sub Fixture()
    fixtureVersion = fixtureVersion + 1
    Call Unit_NpcBow.InitializeTestWorld()
    Call ResetTurnInTestCounters()
    ReDim QuestList(0 To 300)
    ReDim NpcInfoCache(1 To 20)
    ReDim GlobalQuestInfo(1 To 1)
    With UserList(1)
        .Id = 140: .VersionId = fixtureVersion: .flags.UserLogged = True
        .Stats.ELV = 10: .clase = e_Class.Mage: .Faccion.Status = e_Facciones.Ciudadano
        .QuestStats.NumQuestsDone = 0
        Dim slot As Integer
        For slot = 1 To MAXUSERQUESTS
            .QuestStats.Quests(slot).QuestIndex = 0
        Next slot
    End With
    QuestList(1).nombre = "Eligible quest"
    With NpcInfoCache(10)
        .Exists = True: .NumQuest = 1
        ReDim .QuestNumber(1 To 1): .QuestNumber(1) = 1
    End With
    Call ResetQuestDiscoveryIndex()
    DiscoveryEvaluationCount = 0
    Set packets = New Collection
    attempts = 0: rejectAt = 0: capturing = True
End Sub
Private Sub AuthoredProvidersFixture()
    ' Eligibility/relationships copied from assets Dat/NPCs.dat and Quests.DAT;
    ' no subtitle/name matching. Item requirements affect turn-in, not acceptance.
    ReDim NpcInfoCache(1 To 432)
    ReDim QuestList(0 To 427)
    NpcInfoCache(432).Exists = True: NpcInfoCache(432).NumQuest = 2
    ReDim NpcInfoCache(432).QuestNumber(1 To 2)
    NpcInfoCache(432).QuestNumber(1) = 231: NpcInfoCache(432).QuestNumber(2) = 230
    QuestList(231).nombre = "La Aguja y la Garra": QuestList(231).RequiredLevel = 1
    QuestList(231).RequiredOBJs = 1: ReDim QuestList(231).RequiredOBJ(1 To 1)
    QuestList(231).RequiredOBJ(1).ObjIndex = 538: QuestList(231).RequiredOBJ(1).amount = 1
    QuestList(230).nombre = "El Rugido de Bruin": QuestList(230).RequiredLevel = 1: QuestList(230).RequiredQuest = 231
    NpcInfoCache(415).Exists = True: NpcInfoCache(415).NumQuest = 1
    ReDim NpcInfoCache(415).QuestNumber(1 To 1): NpcInfoCache(415).QuestNumber(1) = 38
    QuestList(38).nombre = "Comienza la Aventura": QuestList(38).RequiredLevel = 13
    NpcInfoCache(408).Exists = True: NpcInfoCache(408).NumQuest = 1
    ReDim NpcInfoCache(408).QuestNumber(1 To 1): NpcInfoCache(408).QuestNumber(1) = 427
    QuestList(427).nombre = "Un pequeno problema con las ratas": QuestList(427).RequiredLevel = 1: QuestList(427).LimitLevel = 8
    UserList(1).Stats.ELV = 1
    Call ResetQuestDiscoveryIndex()
End Sub
Private Function LastSnapshotOffers(ByVal npc As Integer, ByVal quest As Integer) As Boolean
    Dim reader As New Network.Reader, bytes() As Byte, count As Integer, i As Integer, n As Integer, q As Integer
    bytes = packets(packets.Count): Call reader.SetData(bytes)
    Call Expect(reader.ReadInt16() = ServerPacketID.eQuestDiscovery, "discovery packet")
    Call reader.ReadInt32(): Call reader.ReadInt16()
    Call Expect(reader.ReadInt8() = 1, "atomic final snapshot")
    count = reader.ReadInt16()
    For i = 1 To count
        n = reader.ReadInt16(): q = reader.ReadInt16(): Call reader.ReadString8()
        If n = npc And q = quest Then LastSnapshotOffers = True
    Next i
    Call Expect(reader.GetAvailable() = 0, "exact discovery snapshot boundary")
End Function
Private Sub CheckQuestNpcSpawnHealth(ByVal caseIndex As Integer)
    On Error GoTo Failed
    Dim previousReader As clsIniManager, definitionPath As String, handle As Integer
    Dim opened As Boolean, npcIndex As Integer, origin As t_WorldPos
    Dim expectedCurrent As Long, expectedMaximum As Long, expectedType As e_NPCType
    Dim bytes() As Byte, reader As New Network.Reader, writer As Network.Writer
    Set previousReader = LeerNPCs
    definitionPath = Environ$("TEMP") & "\ao-quest-npc-health-test.ini"
    handle = FreeFile
    Open definitionPath For Output As #handle
    opened = True
    ' Exact NPC432 section at assets 78d1253e6e662a53a78e58389bd4514cd58d0c77.
    Print #handle, "[NPC432]"
    Print #handle, "Name=Alaric von Nix <Quest de Pieles>"
    Print #handle, "Desc=¡Ah.. un nuevo rostro entre los copos de esta ciudad! Bienvenido a mi taller, viajero."
    Print #handle, "NpcType=17"
    Print #handle, "Head=0"
    Print #handle, "Body=6174"
    Print #handle, "BodyIdle=6174"
    Print #handle, "Ataque1=6174"
    Print #handle, "Heading=3"
    Print #handle, "NpcFaceGrh=32286"
    Print #handle, "Movement=1"
    Print #handle, "ShowName=1"
    Print #handle, "NumQuest=2"
    Print #handle, "QuestNumber1=231"
    Print #handle, "QuestNumber2=230"
    Print #handle, "SoundOpen=964"
    Print #handle, "CollectibleCardIndex=312"
    Print #handle, "DropCount=1"
    Print #handle, "Drop1=1173-10000-1-1"
    Print #handle, "en_Name=Alaric von Nix <Quest>"
    Print #handle, "en_Desc=Ah... a new face among the flakes of this city! Welcome to my workshop, traveler."
    Print #handle, "pt_Name=Alaric von Nix <Missão>"
    Print #handle, "pt_Desc=Ah... um novo rosto entre os flocos desta cidade! Bem-vindo à minha oficina, viajante."
    Print #handle, "fr_Name=Alaric von Nix <Quête>"
    Print #handle, "fr_Desc=Ah... un nouveau visage parmi les flocons de cette ville! Bienvenue dans mon atelier, voyageur."
    Print #handle, "it_Name=Alaric von Nix <Missione>"
    Print #handle, "it_Desc=Ah... un nuovo volto tra i fiocchi di questa città! Benvenuto nel mio laboratorio, viaggiatore."
    Print #handle, "NoWiki=1"
    expectedCurrent = 1: expectedMaximum = 1: expectedType = e_NPCType.Quest
    Select Case caseIndex
    Case 17
        Print #handle, "MinHP=5": Print #handle, "MaxHP=9"
        expectedCurrent = 5: expectedMaximum = 9
    Case 18
        Print #handle, "MinHP=0": Print #handle, "MaxHP=0"
        expectedCurrent = 0: expectedMaximum = 0
    Case 19
        Print #handle, "Attackable=1"
        expectedCurrent = 0: expectedMaximum = 0
    Case 20
        ' A separate ordinary definition, also without HP, must not get this default.
        Print #handle, "[NPC433]": Print #handle, "NpcType=0": Print #handle, "Body=6174"
        expectedCurrent = 0: expectedMaximum = 0: expectedType = e_NPCType.Comun
    Case 21
        Print #handle, "Hostile=1"
        expectedCurrent = 0: expectedMaximum = 0
    Case 22
        Print #handle, "MaxHP=4"
        expectedCurrent = 0: expectedMaximum = 4
    End Select
    Close #handle
    opened = False
    Set LeerNPCs = New clsIniManager
    Call LeerNPCs.Initialize(definitionPath, True)
    Call BuildNpcInfoCache()
    Call Expect(NpcInfoCacheInitialized And NpcInfoCache(432).Exists, "production template loader")
    Call Expect(NpcInfoCache(432).StatsMinHp = IIf(caseIndex = 17, 5, 0), "cache preserves authored current HP")
    origin.Map = 1: origin.x = 50: origin.y = 50
    If caseIndex = 20 Then
        npcIndex = CrearNPC(433, 1, origin, 0, True)
    Else
        npcIndex = CrearNPC(432, 1, origin, 0, True)
    End If
    Call Expect(npcIndex > 0, "production CrearNPC allocates and places fresh spawn")
    Call Expect(NpcList(npcIndex).Char.charindex > 0 And MapData(50, 50, 1).NpcIndex = npcIndex, "spawn placement and character identity")
    Call Expect(NpcList(npcIndex).Stats.MinHp = expectedCurrent And NpcList(npcIndex).Stats.MaxHp = expectedMaximum, "fresh runtime health")
    Call Expect(NpcList(npcIndex).npcType = expectedType, "NPC classification unchanged")
    If caseIndex = 16 Then
        Call Expect(NpcList(npcIndex).Attackable = 0 And NpcList(npcIndex).Hostile = 0, "Alaric remains noncombat")
        Call Expect(NpcList(npcIndex).NumQuest = 2 And NpcList(npcIndex).QuestNumber(1) = 231 And NpcList(npcIndex).QuestNumber(2) = 230, "Alaric provider identity retained")
    End If
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.Clear()
    Call PrepareMessageNpcUpdateHP(npcIndex)
    Call writer.GetData(bytes): Call writer.Clear(): Call reader.SetData(bytes)
    Call Expect(reader.ReadInt16() = ServerPacketID.eCharUpdateHP And reader.ReadInt16() = NpcList(npcIndex).Char.charindex, "production health update identity")
    Call Expect(reader.ReadInt32() = expectedCurrent And reader.ReadInt32() = expectedMaximum, "production health update values")
    Call Expect(reader.ReadInt32() = 0 And reader.GetAvailable() = 0, "exact health packet boundary")
    If caseIndex = 16 Then
        Call Expect(ModifyHealth(npcIndex, -1), "subsequent health mutation can reach zero")
        Call writer.Clear(): Call PrepareMessageNpcUpdateHP(npcIndex)
        Call writer.GetData(bytes): Call writer.Clear(): Call reader.SetData(bytes)
        Call reader.ReadInt16(): Call reader.ReadInt16()
        Call Expect(reader.ReadInt32() = 0 And reader.ReadInt32() = 1, "dead 0/1 never normalized by update or serialization")
        Call Expect(NpcList(npcIndex).Stats.MinHp = 0 And NpcList(npcIndex).Stats.MaxHp = 1, "health serialization does not revive actor")
    End If
    Set LeerNPCs = previousReader: NpcInfoCacheInitialized = False
    Kill definitionPath
    Exit Sub
Failed:
    Dim failureNumber As Long, failureText As String
    failureNumber = Err.Number: failureText = Err.Description
    If opened Then Close #handle
    Set LeerNPCs = previousReader: NpcInfoCacheInitialized = False
    If Len(definitionPath) > 0 Then
        If FileExist(definitionPath) Then Kill definitionPath
    End If
    Err.Raise failureNumber, "CheckQuestNpcSpawnHealth", failureText
End Sub
Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Fixture()
    Dim bytes() As Byte, reader As New Network.Reader, firstSequence As Long, q As Integer, count As Integer
    Dim checks As Long, before As Long, flags As Byte, page As Integer
    Select Case index
    Case 1
        Call Expect(IsQuestDiscoverable(1, 1), "valid offer")
        Call Expect(CanUserAcceptQuest(1, 0, 1, QuestList(1)), "acceptance shares predicate")
        Call Expect(PublishQuestDiscovery(1, True), "initial snapshot queued")
        bytes = packets(1): Call reader.SetData(bytes)
        Call Expect(reader.ReadInt16() = ServerPacketID.eQuestDiscovery, "packet id")
        firstSequence = reader.ReadInt32()
        Call Expect(firstSequence > 0 And reader.ReadInt16() = 0, "transaction identity")
        Call Expect(reader.ReadInt8() = 1 And reader.ReadInt16() = 1, "complete one-offer replacement")
        Call Expect(reader.ReadInt16() = 10 And reader.ReadInt16() = 1, "template and quest identities")
        Call Expect(reader.ReadString8() = "Eligible quest" And reader.GetAvailable() = 0, "exact title and packet length")
    Case 2
        QuestList(1).Trabajador = True
        Call Expect(Not IsQuestDiscoverable(1, 1), "worker restriction")
        UserList(1).clase = e_Class.Trabajador
        Call Expect(IsQuestDiscoverable(1, 1), "worker permitted")
        QuestList(1).Trabajador = False
        QuestList(1).RequiredClassesCount = 1
        ReDim QuestList(1).RequiredClass(1 To 1)
        QuestList(1).RequiredClass(1) = e_Class.Mage
        Call Expect(Not IsQuestDiscoverable(1, 1), "class restriction")
        UserList(1).clase = e_Class.Mage
        Call Expect(IsQuestDiscoverable(1, 1), "required class permitted")
    Case 3
        QuestList(1).RequiredLevel = 11
        Call Expect(Not IsQuestDiscoverable(1, 1), "minimum level")
        QuestList(1).RequiredLevel = 0: QuestList(1).LimitLevel = 9
        Call Expect(Not IsQuestDiscoverable(1, 1), "maximum level")
        QuestList(1).LimitLevel = 0: QuestList(1).RequiredSkill.SkillType = 1
        QuestList(1).RequiredSkill.RequiredValue = 20: UserList(1).Stats.UserSkills(1) = 0
        Call Expect(Not IsQuestDiscoverable(1, 1), "required skill")
        UserList(1).Stats.UserSkills(1) = 20
        Call Expect(IsQuestDiscoverable(1, 1), "skill attained")
    Case 4
        QuestList(1).RequiredQuest = 2
        Call Expect(Not IsQuestDiscoverable(1, 1), "prerequisite hidden")
        Call AddDoneQuest(1, 2)
        Call Expect(IsQuestDiscoverable(1, 1), "prerequisite complete")
        Call AddDoneQuest(1, 1)
        Call Expect(Not IsQuestDiscoverable(1, 1), "non-repeatable complete")
        QuestList(1).Repetible = 1
        Call Expect(IsQuestDiscoverable(1, 1), "repeatable reacceptance")
        UserList(1).QuestStats.Quests(1).QuestIndex = 1
        Call Expect(Not IsQuestDiscoverable(1, 1), "active repeatable excluded")
    Case 5
        QuestList(1).PermittedFactions = e_QuestPermittedFactions.RoyalArmy
        Call Expect(Not IsQuestDiscoverable(1, 1), "faction restriction")
        UserList(1).Faccion.Status = e_Facciones.Armada
        Call Expect(IsQuestDiscoverable(1, 1), "faction permitted")
        For q = 1 To MAXUSERQUESTS
            UserList(1).QuestStats.Quests(q).QuestIndex = 2
        Next q
        Call Expect(Not IsQuestDiscoverable(1, 1), "no free quest slot")
        UserList(1).QuestStats.Quests(1).QuestIndex = 0
        Call Expect(IsQuestDiscoverable(1, 1), "free slot restored")
    Case 6
        QuestList(1).GlobalQuestIndex = 1: QuestList(1).GlobalQuestThresholdNeeded = 10
        Call ResetQuestDiscoveryIndex()
        Call Expect(Not IsQuestDiscoverable(1, 1), "global threshold")
        Call Expect(PublishQuestDiscovery(1), "empty initial snapshot")
        before = DiscoveryEvaluationCount
        GlobalQuestInfo(1).GatheringGlobalCounter = 1
        Call RefreshGlobalQuestTurnIns(1): Call ProcessQuestTurnInUpdates()
        Call Expect(DiscoveryEvaluationCount = before And packets.Count = 1, "below-threshold contributions do not rescan offers")
        GlobalQuestInfo(1).GatheringGlobalCounter = 10
        Call RefreshGlobalQuestTurnIns(1): Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 2, "unaccepted global offer refreshed")
        Call RefreshGlobalQuestTurnIns(1): Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 2, "unchanged global signature suppressed")
    Case 7
        Call SendQuestTurnInSnapshot(1): Call SendQuestTurnInSnapshot(1): Call ProcessQuestTurnInUpdates()
        Call Expect(DiscoveryEvaluationCount = 1 And packets.Count = 1, "dirty mutations coalesced")
        before = DiscoveryEvaluationCount
        Call SendQuestTurnInSnapshot(1): Call ProcessQuestTurnInUpdates()
        Call Expect(DiscoveryEvaluationCount = before And packets.Count = 1, "unrelated mutation skips eligibility and packet")
        UserList(1).Stats.ELV = 11
        Call SendQuestTurnInSnapshot(1): Call ProcessQuestTurnInUpdates()
        Call Expect(DiscoveryEvaluationCount = before + 1 And packets.Count = 1, "changed input with identical availability sends nothing")
        Call AddDoneQuest(1, 1): Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 2, "completion invalidates availability")
        bytes = packets(2): Call reader.SetData(bytes)
        Call reader.ReadInt16(): Call reader.ReadInt32(): Call reader.ReadInt16(): Call reader.ReadInt8()
        Call Expect(reader.ReadInt16() = 0, "empty replacement clears old offer")
    Case 8
        rejectAt = 1
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates(1000)
        Call Expect(packets.Count = 0, "failed enqueue not cached")
        Call ProcessQuestTurnInUpdates(1100): Call Expect(attempts = 1, "retry bounded")
        Call ProcessQuestTurnInUpdates(1250): Call Expect(packets.Count = 1, "retry succeeds")
        Call SendQuestTurnInSnapshot(1): Call ProcessQuestTurnInUpdates(1500)
        Call Expect(packets.Count = 1, "successful cache suppresses duplicate")
    Case 9
        rejectAt = 1: Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates(1000)
        UserList(1).VersionId = 2: UserList(1).Id = 141
        Call ProcessQuestTurnInUpdates(1300): Call Expect(packets.Count = 0, "slot reuse cannot inherit dirty work")
        Call SendQuestTurnInSnapshot(1, True): Call ProcessQuestTurnInUpdates(1600)
        Call Expect(packets.Count = 1, "new session forced snapshot")
    Case 10
        NpcInfoCache(10).NumQuest = 255: ReDim NpcInfoCache(10).QuestNumber(1 To 255)
        For q = 1 To 255
            NpcInfoCache(10).QuestNumber(q) = q: QuestList(q).nombre = "Quest " & q
        Next q
        Call ResetQuestDiscoveryIndex()
        rejectAt = 2
        Call Expect(Not PublishQuestDiscovery(1, True), "partial page failure")
        Call Expect(packets.Count = 1, "first page accepted only")
        rejectAt = 0
        Call Expect(PublishQuestDiscovery(1, True), "retry resumes transaction")
        Call Expect(packets.Count = 2, "partial-send resumes without repeating accepted page")
        For page = 1 To 2
            bytes = packets(page): Call reader.SetData(bytes)
            Call reader.ReadInt16(): Call reader.ReadInt32()
            Call Expect(reader.ReadInt16() = page - 1, "ordered pages")
            flags = reader.ReadInt8(): count = reader.ReadInt16()
            If page = 1 Then Call Expect(flags = 0 And count = 128, "full nonterminal page")
            If page = 2 Then Call Expect(flags = 1 And count = 127, "terminal remainder")
            For q = 1 To count
                Call reader.ReadInt16(): Call reader.ReadInt16(): Call reader.ReadString8()
            Next q
            Call Expect(reader.GetAvailable() = 0, "exact page layout")
        Next page
    Case 11
        QuestList(1).GlobalQuestIndex = 1
        ReDim GlobalQuestInfo(1).IsBossAlive(1 To 1)
        GlobalQuestInfo(1).IsActive = False
        Call Expect(Not IsQuestDiscoverable(1, 1), "inactive global event")
        GlobalQuestInfo(1).IsActive = True
        Call Expect(IsQuestDiscoverable(1, 1), "active event without boss")
        GlobalQuestInfo(1).IsBossAlive(1) = True
        Call Expect(Not IsQuestDiscoverable(1, 1), "live boss gate")
        Call ResetQuestDiscoveryIndex()
        Call Expect(PublishQuestDiscovery(1), "blocked initial event snapshot")
        GlobalQuestInfo(1).IsBossAlive(1) = False
        Call RefreshGlobalQuestTurnIns(1): Call ProcessQuestTurnInUpdates()
        Call Expect(packets.Count = 2, "boss-state change refreshes discovery")
    Case 12
        Call AuthoredProvidersFixture()
        Call Expect(QuestAcceptanceFailure(1, 0, 231, QuestList(231)) = QuestAcceptAllowed, "Alaric 231 is genuinely acceptable in controlled level-one fixture")
        Call Expect(QuestAcceptanceFailure(1, 0, 230, QuestList(230)) = QuestAcceptPrerequisite, "Alaric 230 requires completed 231")
        Call SendQuestObjectiveSnapshot(1, True): Call ProcessQuestTurnInUpdates()
        Call Expect(LastSnapshotOffers(432, 231) And Not LastSnapshotOffers(432, 230), "zero accepted quests still publish Alaric eligible offer without dialog")
        Call Expect(LastSnapshotOffers(408, 427) And Not LastSnapshotOffers(415, 38), "Romina level-one eligible; Heather requires 13")
    Case 13
        Call AuthoredProvidersFixture(): UserList(1).Stats.ELV = 13
        Call SendQuestObjectiveSnapshot(1, True): Call ProcessQuestTurnInUpdates()
        Call Expect(LastSnapshotOffers(415, 38) And Not LastSnapshotOffers(408, 427), "Heather eligibility and Romina maximum level are preserved")
    Case 14
        Call AuthoredProvidersFixture()
        Call SendQuestObjectiveSnapshot(1, True): Call ProcessQuestTurnInUpdates()
        NpcList(1).flags.NPCActive = True: NpcList(1).Numero = 432
        NpcList(1).pos = UserList(1).pos: NpcList(1).NumQuest = 2
        ReDim NpcList(1).QuestNumber(1 To 2)
        NpcList(1).QuestNumber(1) = 231: NpcList(1).QuestNumber(2) = 230
        Call SetNpcRef(UserList(1).flags.TargetNPC, 1)
        Dim acceptBytes(0 To 0) As Byte: acceptBytes(0) = 1
        Set Protocol.reader = New Network.Reader: Call Protocol.reader.SetData(acceptBytes)
        Call Protocol.HandleQuestAccept(1): Call ProcessQuestTurnInUpdates()
        Call Expect(TieneQuest(1, 231) > 0, "production acceptance retains actual quest ID")
        Call Expect(Not LastSnapshotOffers(432, 231), "confirmed acceptance removes available marker")
    Case 15
        Call AuthoredProvidersFixture(): rejectAt = 1
        Call SendQuestObjectiveSnapshot(1, True): Call ProcessQuestTurnInUpdates(1000)
        Call Expect(packets.Count = 0, "failed login enqueue is not cached as sent")
        rejectAt = 0: Call ProcessQuestTurnInUpdates(1250)
        Call Expect(LastSnapshotOffers(432, 231), "bounded retry publishes initial Alaric snapshot")
        Call ResetQuestDiscoveryUser(1): Call IncreaseVersionId(1)
        Call SendQuestObjectiveSnapshot(1, True): Call ProcessQuestTurnInUpdates(1500)
        Call Expect(LastSnapshotOffers(432, 231), "reconnect republishes eligible offer")
    Case 16 To 22
        Call CheckQuestNpcSpawnHealth(index)
    End Select
    capturing = False
    Call UnitTesting.Progress("Quest discovery " & index & " passed")
    Call UnitTesting.RunTest("Quest discovery " & index, True)
    Exit Sub
Failed:
    Dim failure As String
    failure = Err.Description
    capturing = False
    Call UnitTesting.Progress("Quest discovery " & index & " failed: " & failure)
    Call UnitTesting.RunTestError("Quest discovery " & index, failure)
End Sub
Public Function TestSuite() As Boolean
    Dim before As Long, i As Integer
    before = UnitTesting.GetFailedTests()
    For i = 1 To 22
        Call RunCase(i)
    Next i
    TestSuite = UnitTesting.GetFailedTests() = before
End Function
#End If
