Attribute VB_Name = "MapSchemaHarness"
' Copyright (C) 2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Public Const MAP_LAYER_COUNT As Long = 5
Public Const NumMaps As Long = 2
Public Const XMinMapSize As Long = 1
Public Const XMaxMapSize As Long = 100
Public Const YMinMapSize As Long = 1
Public Const YMaxMapSize As Long = 100
Public Const KNOWN_TILE_FLAGS As Long = 511
Public Const KNOWN_ZONE_FLAGS As Long = 2097151
Public Const FLAG_AGUA As Long = 32
Public Const FLAG_ARBOL As Long = 64
Public Type Position
    map As Integer
    x As Byte
    y As Byte
End Type
Public Type t_WorldPos
    map As Integer
    x As Byte
    y As Byte
End Type
Public Type ObjectInfo
    ObjIndex As Integer
    amount As Long
    data As Long
End Type
Public Type LightInfo
    Color As Long
    Rango As Byte
End Type
Public Type t_MapBlock
    Blocked As Long
    Graphic(1 To 5) As Long
    trigger As Long
    ParticulaIndex As Long
    Luz As LightInfo
    ObjInfo As ObjectInfo
    NpcIndex As Integer
    TileExit As Position
End Type
Public Type MapProperties
    MapResource As Integer
    ZoneFlags As Long
    CastleEntrances As Dictionary
    StaticCastleEntrances As Dictionary
    LegacyCastleEntrances As Dictionary
    RoofSeams As Dictionary
    TileOverrides As Dictionary
End Type
Public Type UserFlags
    UserLogged As Boolean
End Type
Public Type UserInfo
    pos As Position
    flags As UserFlags
    isGm As Boolean
    capable As Boolean
    AccountID As Long
    name As String
End Type
Public Type ObjectDefinition
    OBJType As Long
    Subtipo As Long
    VidaUtil As Long
End Type
Public Type NpcDefinition
    pos As Position
    Orig As Position
    name As String
    Numero As Integer
End Type
Public Enum e_OBJType
    otOreDeposit = 1
    otTrees = 2
    otTeleport = 3
End Enum
Public Enum e_TeleportSubType
    eTransportNetwork = 1
End Enum
Public MapData(1 To 100, 1 To 100, 1 To 2) As t_MapBlock
Public MapInfo(1 To 2) As MapProperties
Public UserList(1 To 3) As UserInfo
Public LastUser As Integer
Public Writer As Network.Writer
Public Const HOO_CAP_TILE_PROPERTIES_V1 As Long = 16
Public Const ToIndex As Long = 1
Public Enum ServerPacketID
    eHooTileProperties = 206
End Enum
Public ObjData(0 To 32767) As ObjectDefinition
Public NpcList(0 To 32767) As NpcDefinition
Public DatPath As String
Public legacyConfigEnabled As Boolean
Public legacyNoDrop As Boolean
Public DBError As String
Private simulatedScriptFailure As Boolean
Private rollbackCount As Long
Private failures As Long, checks As Long
Private sent As Long
Private spawnedNpcs As Long

Public Sub LogInfoServidor(ByVal message As String)
End Sub

Private Sub TestCastleIdentity()
    ReDim CastleData(1 To 3)
    CastleData(1).id = 7: CastleData(2).id = 1: CastleData(3).id = 70001
    Set CastleData(1).castleWhiteList = New Dictionary
    Set CastleData(2).castleWhiteList = New Dictionary
    Set CastleData(3).castleWhiteList = New Dictionary
    Call Check(GetCastleSlotById(1) = 2 And GetCastleSlotById(70001) = 3 And GetCastleSlotById(99) = -1, "shuffled noncontiguous castle IDs")
    With CastleData(1)
        .is_active = True: .owner_account_id = 9
        .castle_coordinates.outside.map = 1: .castle_coordinates.outside.x = 50: .castle_coordinates.outside.y = 50
    End With
    With CastleData(3)
        .is_active = True: .owner_account_id = 4147
        .castle_coordinates.outside.map = 2: .castle_coordinates.outside.x = 50: .castle_coordinates.outside.y = 50
        .castle_coordinates.inside.map = 2: .castle_coordinates.inside.x = 50: .castle_coordinates.inside.y = 72
    End With
    MapData(50, 50, 1).ObjInfo.ObjIndex = 6382
    MapData(50, 50, 2).ObjInfo.ObjIndex = 6382
    UserList(1).AccountID = 9: UserList(1).name = "visitor"
    Call Check(CheckCastleEntryWhiteList(1, 7), "owner can enter selected castle")
    Call Check(Not CheckCastleEntryWhiteList(1, 70001), "owning another castle does not bypass target whitelist")
    Call CastleData(3).castleWhiteList.Add("visitor", True)
    Call Check(CheckCastleEntryWhiteList(1, 70001), "selected castle whitelist membership")
    CastleData(3).is_active = False
    Call Check(Not CheckCastleEntryWhiteList(1, 70001), "inactive castle denies access")
    CastleData(3).is_active = True
    Call Check(Not CanPublishCastle(2, 1, 50, 50), "unconfigured interior prevents publication")
    Call Check(CanPublishCastle(3, 1, 50, 50), "valid footprint can publish")
    Call RegisterCastleEntrance(1, 48, 50, 123, True)
    Call Check(Not CanPublishCastle(3, 1, 50, 50), "static/dynamic conflict rejected before relocation")
    Call Check(Not CanPublishCastle(3, 1, 2, 2), "out-of-bounds footprint rejected")
End Sub

Public Function EsGM(ByVal userIndex As Integer) As Boolean
    EsGM = UserList(userIndex).isGm
End Function
Public Function InMapBounds(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    InMapBounds = map >= 1 And map <= 2 And x >= 1 And x <= 100 And y >= 1 And y <= 100
End Function
Public Function UserSupportsHooCapability(ByVal userIndex As Integer, ByVal capability As Long) As Boolean
    UserSupportsHooCapability = UserList(userIndex).capable
End Function
Public Sub CapturePacket(ByVal recipient As Integer)
    Dim bytes() As Byte, reader As Network.reader
    Call Writer.GetData(bytes)
    Call Check(UBound(bytes) = 11, "tile update is exactly12 bytes")
    Set reader = New Network.reader
    Call reader.SetData(bytes)
    Call Check(reader.ReadInt16() = 206, "tile update packet206")
    Call Check(reader.ReadInt16() = 1 And reader.ReadInt16() = 1 And reader.ReadInt16() = 1, "tile update position")
    Call Check(reader.ReadInt32() = 511, "tile update Int32 flagmask")
    Call Writer.Clear()
    sent = sent + 1
End Sub

Private Sub TestZoneFlags()
    Dim bit As Long, bitIndex As Integer
    For bitIndex = 0 To 20
        bit = 2 ^ bitIndex
        MapInfo(1).ZoneFlags = KNOWN_ZONE_FLAGS
        Call SetMapZoneFlag(1, bit, False)
        Call Check(Not HasMapZoneFlag(1, bit) And MapInfo(1).ZoneFlags = (KNOWN_ZONE_FLAGS Xor bit), "clearing one map flag preserves every other flag")
        Call SetMapZoneFlag(1, bit, True)
        Call Check(HasMapZoneFlag(1, bit) And MapInfo(1).ZoneFlags = KNOWN_ZONE_FLAGS, "setting one map flag preserves every other flag")
    Next bitIndex
    legacyConfigEnabled = True
    Call Check(TestLegacyFlags("511", 1) = (KNOWN_ZONE_FLAGS And Not (e_ZoneFlags.Prison Or e_ZoneFlags.SafeFight)), "all legacy flags, defaults and Map.dat flags")
    legacyConfigEnabled = False
    Call Check(TestLegacyFlags("0", 0) = (e_ZoneFlags.DropItems Or e_ZoneFlags.FriendlyFire), "absent legacy flags preserve runtime defaults")
    legacyNoDrop = True
    Call Check(TestLegacyFlags("0", 0) = e_ZoneFlags.FriendlyFire, "legacy no-drop list overrides drop default")
    legacyNoDrop = False
    Call Check(TestLegacyFlags("nEwBiE", 0) = (e_ZoneFlags.NewbieOnly Or e_ZoneFlags.DropItems Or e_ZoneFlags.FriendlyFire), "legacy textual newbie restriction")
    Call Check(TestLegacyFlags("none", 0) = (e_ZoneFlags.DropItems Or e_ZoneFlags.FriendlyFire), "legacy unknown text has no restriction bits")
    MapData(1, 1, 1).Blocked = 0: MapData(1, 1, 1).trigger = 0
    Call Check(Not IsWaterTile(1, 1, 1), "dry tile")
    MapData(1, 1, 1).trigger = e_Trigger.SwimSuitPath
    Call Check(IsWaterTile(1, 1, 1), "live swimming flag changes authoritative water predicate")
    MapData(1, 1, 1).trigger = 0
    Call Check(Not IsWaterTile(1, 1, 1), "clearing swimming flag restores dry tile")
    MapData(1, 1, 1).Blocked = FLAG_AGUA
    MapData(1, 1, 1).trigger = e_Trigger.SwimSuitPath
    MapData(1, 1, 1).trigger = 0
    Call Check(IsWaterTile(1, 1, 1), "clearing swimming flag preserves authored water")
End Sub

Private Sub TestRuntimeProperties()
    Dim copy As Dictionary, zone As Long
    Set Writer = New Network.Writer
    LastUser = 3
    UserList(1).pos.map = 1: UserList(1).pos.x = 1: UserList(1).pos.y = 1
    UserList(2).pos.map = 1: UserList(2).pos.x = 2: UserList(2).pos.y = 1
    UserList(3).pos.map = 2: UserList(3).pos.x = 1: UserList(3).pos.y = 1
    UserList(1).flags.UserLogged = True: UserList(2).flags.UserLogged = True: UserList(3).flags.UserLogged = True
    MapData(1, 1, 1).trigger = 9: MapData(2, 1, 1).trigger = 8
    Call Check(BothInPvPArena(1, 2), "different complete masks still same arena")
    Call Check(Not CrossesPvPArenaBoundary(1, 2), "arena no boundary")
    MapData(2, 1, 1).trigger = 1
    Call Check(CrossesPvPArenaBoundary(1, 2) And CrossesPvPArenaBoundary(2, 1), "arena boundary symmetric")
    MapData(1, 1, 1).trigger = 1
    Call Check(Not BothInPvPArena(1, 2) And Not CrossesPvPArenaBoundary(1, 2), "no arena uses normal rules")
    MapInfo(1).ZoneFlags = 1
    Call Check(IsPrisonMap(1), "prison independent of tile flags")
    Call Check(Not SetTileTriggerFlags(1, 511), "non-GM cannot edit")
    UserList(1).isGm = True
    Call Check(Not SetTileTriggerFlags(1, -1) And Not SetTileTriggerFlags(1, 512), "unknown flags reject")
    UserList(1).capable = True: UserList(3).capable = True
    Call Check(SetTileTriggerFlags(1, 511), "GM can replace full mask")
    Call Check(sent = 1, "broadcast only negotiated same-map user")
    UserList(2).pos.x = 1: UserList(2).capable = True
    Call ReplayTileProperties(2, 1)
    Call Check(sent = 2, "late join gets accepted edits")
    Set copy = CopyPropertyDictionary(MapInfo(1).TileOverrides)
    copy.Item(TilePropertyKey(1, 1)) = 0
    Call Check(MapInfo(1).TileOverrides.Item(TilePropertyKey(1, 1)) = 511, "instance overrides copied without alias")
    Set MapInfo(1).CastleEntrances = Nothing: Set MapInfo(1).StaticCastleEntrances = Nothing
    MapData(47, 69, 1).trigger = 9
    Call RegisterCastleEntrance(1, 47, 69, 70001)
    Call RegisterCastleEntrance(1, 48, 69, 70001)
    Call RegisterCastleEntrance(1, 47, 70, 70001)
    Call RegisterCastleEntrance(1, 48, 70, 70001)
    Call Check(CastleAtTile(1, 47, 69) = 70001, "castle uses stableLongID")
    Call RemoveCastleEntrances(1, 70001)
    Call Check(CastleAtTile(1, 47, 69) = 0 And MapData(47, 69, 1).trigger = 9, "destroy entrance preserves unrelated tileflags")
End Sub

Public Function FileExist(ByVal path As String, ByVal attributes As Long) As Boolean
    On Error GoTo Missing
    FileExist = (GetAttr(path) And vbDirectory) = 0
Missing:
End Function
Public Sub TraceError(ByVal number As Long, ByVal description As String, ByVal source As String, Optional ByVal line As Long = 0)
End Sub
Public Function HayAgua(ByVal map As Long, ByVal x As Integer, ByVal y As Integer) As Boolean
End Function
Public Function EsArbol(ByVal graphic As Long) As Boolean
End Function
Public Function OpenNPC(ByVal definition As Integer) As Integer
    spawnedNpcs = spawnedNpcs + 1
    OpenNPC = definition
    NpcList(definition).Numero = definition
    NpcList(definition).name = "test"
End Function
Public Sub MakeNPCChar(ByVal a As Boolean, ByVal b As Long, ByVal npc As Integer, ByVal map As Long, ByVal x As Integer, ByVal y As Integer)
End Sub
Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    checks = checks + 1
    If Not condition Then
        failures = failures + 1
        Print #1, "FAIL: " & label
    End If
End Sub
Private Function LoadFixture(ByVal path As String) As Boolean
    On Error GoTo Rejected
    Dim x As Long, y As Long, emptyTile As t_MapBlock, emptyMap As MapProperties
    For x = 1 To 100
        For y = 1 To 100
            MapData(x, y, 1) = emptyTile
        Next y
    Next x
    MapInfo(1) = emptyMap
    spawnedNpcs = 0
    MapData(100, 100, 1).trigger = 257
    Call CargarMapaFormatoCSM(1, path)
    LoadFixture = True
    Exit Function
Rejected:
    Call Check(spawnedNpcs = 0, "invalid map spawns no NPCs")
    Call Check(MapData(100, 100, 1).trigger = 257, "invalid map leaves existing tiles unchanged")
    Print #1, "Rejected: " & path & ": " & Err.Description
End Function
Private Function RejectLegacy(ByVal value As Integer) As Boolean
    On Error GoTo Rejected
    Dim zone As Long, flags As Long
    flags = ConvertLegacyTrigger(value, zone)
    Exit Function
Rejected:
    RejectLegacy = True
End Function
Public Sub Main()
    On Error GoTo Failed
    Dim args() As String, root As String, maps As String, entry As String, zone As Long, value As Integer, count As Long
    Dim commandLine As String
    commandLine = Trim$(Command$)
    If Left$(commandLine, 1) = Chr$(34) And Right$(commandLine, 1) = Chr$(34) Then commandLine = Mid$(commandLine, 2, Len(commandLine) - 2)
    args = Split(commandLine, "|")
    If args(0) = "split-sql" Then
        Call ExportMigrationStatements(args(1), args(2))
        Exit Sub
    End If
    root = args(0)
    maps = args(1)
    Open App.path & "\results.txt" For Output As #1
    Call Check(LoadFixture(root & "\all_sections.csm"), "shared all-sections fixture")
    Call Check(MapInfo(1).ZoneFlags = KNOWN_ZONE_FLAGS, "all map-wide flags")
    Call Check(MapData(2, 2, 1).trigger = 511, "combined flag width")
    Call Check((MapData(1, 1, 1).Blocked And 32) <> 0, "authored water survives graphics")
    Call Check(MapInfo(1).CastleEntrances.Item(TilePropertyKey(4, 4)) = 70001, "castle ID width")
    Call Check(MapInfo(1).RoofSeams.Count = 2, "roof seams")
    If FileExist(App.path & "\roundtrip.csm", vbNormal) Then Kill App.path & "\roundtrip.csm"
    Call SaveMapCsm3(1, App.path & "\roundtrip.csm")
    Call Check(LoadFixture(App.path & "\roundtrip.csm"), "native CSM3 writer roundtrip")
    Call SetMapZoneFlag(1, e_ZoneFlags.Safe, False)
    If FileExist(App.path & "\updated-flags.csm", vbNormal) Then Kill App.path & "\updated-flags.csm"
    Call SaveMapCsm3(1, App.path & "\updated-flags.csm")
    Call Check(LoadFixture(App.path & "\updated-flags.csm"), "native writer saves updated map flags")
    Call Check(MapInfo(1).ZoneFlags = (KNOWN_ZONE_FLAGS Xor e_ZoneFlags.Safe), "modern flags stay authoritative on reload")
    entry = Dir$(root & "\*.csm")
    Do While Len(entry) > 0
        If entry <> "all_sections.csm" Then Call Check(Not LoadFixture(root & "\" & entry), "reject " & entry)
        entry = Dir$()
    Loop
    Call Check(ConvertLegacyTrigger(16, zone) = 0, "old16 clears")
    Call Check(ConvertLegacyTrigger(2, zone) = 2 And ConvertLegacyTrigger(3, zone) = 2, "NPC restrictions merge")
    Call Check(ConvertLegacyTrigger(8, zone) = 32 And ConvertLegacyTrigger(11, zone) = 32 And ConvertLegacyTrigger(18, zone) = 32, "swimming paths merge")
    Call Check(ConvertLegacyTrigger(19, zone) = 0 And zone = 1, "old19 promotes entire map")
    For value = 60 To 73
        Call Check(ConvertLegacyTrigger(value, zone) = 1, "roof60-73")
    Next value
    For value = 90 To 99
        Call Check(ConvertLegacyTrigger(value, zone) = 1, "roof90-99")
    Next value
    Call Check(RejectLegacy(9) And RejectLegacy(15) And RejectLegacy(74), "unknown legacy values reject")
    Call Check(HasTileFlag(9, e_Trigger.PvPArena) And HasTileFlag(9, e_Trigger.UnderRoof), "combined arena and roof predicates")
    entry = Dir$(maps & "\*.csm")
    Do While Len(entry) > 0
        Call Check(LoadFixture(maps & "\" & entry), entry)
        If LCase$(entry) = "mapa66.csm" Then Call Check(IsPrisonMap(1), "production prison map66")
        count = count + 1
        entry = Dir$()
    Loop
    Call Check(count = 773, "complete production map inventory")
    Call TestSqlSplitter()
    Call TestMigrationRollback()
    Call TestZoneFlags()
    Call TestRuntimeProperties()
    Call TestCastleIdentity()
    Print #1, CStr(checks) & " checks, " & CStr(failures) & " failures, " & CStr(count) & " maps"
    Close #1
    Exit Sub
Failed:
    Print #1, "FATAL: " & CStr(Err.Number) & " " & Err.Description
    Close #1
End Sub

Public Function GetVar(ByVal path As String, ByVal section As String, ByVal key As String) As String
    GetVar = IIf(legacyConfigEnabled, "1", "0")
End Function

Public Function EsMapaNoDrop(ByVal map As Long) As Boolean
    EsMapaNoDrop = legacyNoDrop
End Function

Public Function FileText(ByVal path As String) As String
    FileText = "BEGIN;SELECT 1;COMMIT;"
End Function

Public Function Query(ByVal sql As String) As ADODB.Recordset
    If sql = "ROLLBACK;" Then
        rollbackCount = rollbackCount + 1
        DBError = "Simulated rollback error must not replace the original error"
    ElseIf simulatedScriptFailure Then
        DBError = "Original SQL failure"
    Else
        Set Query = New ADODB.Recordset
    End If
End Function

Private Sub TestMigrationRollback()
    simulatedScriptFailure = True
    Call Check(Not RunScriptInFile("isolated stub"), "failed dated script reports failure")
    Call Check(rollbackCount = 1, "failed dated script rolls back transaction")
    Call Check(DBError = "Original SQL failure", "rollback preserves original database error")
    simulatedScriptFailure = False
    Call Check(RunScriptInFile("isolated stub"), "successful dated script reports success")
    Call Check(rollbackCount = 1, "successful dated script does not roll back")
End Sub

Private Function RejectSql(ByVal sql As String) As Boolean
    On Error GoTo Rejected
    Dim statements As Collection
    Set statements = SplitMigrationSql(sql)
    Exit Function
Rejected:
    RejectSql = True
End Function

Private Sub TestSqlSplitter()
    Dim statements As Collection, q As String
    q = Chr$(34)
    Set statements = SplitMigrationSql("-- ignore;" & vbCrLf & "BEGIN; SELECT 'a;''b', " & q & "c;" & q & q & "d" & q & ", [e;f], `g;``h`; /* skip; */ COMMIT; -- trailing")
    Call Check(statements.Count = 3, "SQL splitter preserves quoted semicolons and ignores comments")
    Call Check(InStr(statements(2), "'a;''b'") > 0 And InStr(statements(2), "[e;f]") > 0, "SQL splitter preserves escaped quotes and bracket identifiers")
    Set statements = SplitMigrationSql("SELECT/* gap */1; SELECT '--;/*text*/'; SELECT 2")
    Call Check(statements.Count = 3 And InStr(statements(1), "SELECT 1") > 0, "SQL comments separate tokens; final statement need not end with semicolon")
    Set statements = SplitMigrationSql("; ; --only" & vbLf & "/* ; */")
    Call Check(statements.Count = 0, "SQL empty/comment-only statements ignored")
    Call Check(RejectSql("SELECT 'unfinished") And RejectSql("/* unfinished") And RejectSql("SELECT [unfinished"), "SQL unfinished quotes/comments reject before execution")
    Call Check(RejectSql("CREATE TRIGGER t AFTER INSERT ON x BEGIN SELECT 1; END;"), "unsupported compound trigger fails explicitly")
End Sub

Private Sub ExportMigrationStatements(ByVal inputPath As String, ByVal outputPrefix As String)
    On Error GoTo ExportFailed
    Dim fh As Integer, sql As String, statements As Collection, index As Long
    fh = FreeFile()
    Open inputPath For Binary As #fh
    sql = Space$(LOF(fh))
    Get #fh, , sql
    Close #fh
    fh = 0
    Set statements = SplitMigrationSql(sql)
    For index = 1 To statements.Count
        fh = FreeFile()
        Open outputPrefix & "." & CStr(index) & ".sql" For Output As #fh
        Print #fh, CStr(statements(index))
        Close #fh
        fh = 0
    Next index
    fh = FreeFile()
    Open outputPrefix & ".count" For Output As #fh
    Print #fh, CStr(statements.Count)
    Close #fh
    fh = 0
    Exit Sub
ExportFailed:
    Dim failure As String
    failure = Err.Description
    If fh > 0 Then Close #fh
    fh = FreeFile()
    Open outputPrefix & ".error" For Output As #fh
    Print #fh, failure
    Close #fh
    fh = 0
End Sub
