Attribute VB_Name = "Unit_ManagedConfig"
Option Explicit
#If UNIT_TEST = 1 Then
Private mChecks As Long
Private mFailures As Long
Private mResults As String

Private Sub Check(ByVal name As String, ByVal passed As Boolean)
    mChecks = mChecks + 1
    If Not passed Then mFailures = mFailures + 1
    mResults = mResults & IIf(passed, "PASS ", "FAIL ") & name & vbCrLf
End Sub

Private Sub PutFixture(ByVal path As String, ByVal text As String)
    Dim handle As Integer
    handle = FreeFile()
    Open path For Output As #handle
    Print #handle, text
    Close #handle
End Sub

Private Function InvalidOption(ByVal arguments As String) As Boolean
    On Error GoTo Expected
    Dim present As Boolean, value As String
    value = StartupOption(arguments, "/config-dir", present)
    Exit Function
Expected:
    InvalidOption = True
End Function

Private Function InvalidDirectory(ByVal arguments As String) As Boolean
    On Error GoTo Expected
    Call InitializeManagedConfig(arguments)
    Exit Function
Expected:
    InvalidDirectory = InStr(1, Err.Description, "config", vbTextCompare) > 0
End Function

Private Function Unreadable(ByVal config As e_ManagedConfig) As Boolean
    On Error GoTo Expected
    Dim reader As clsIniManager
    Set reader = ReadManagedConfig(config)
    Exit Function
Expected:
    Unreadable = InStr(1, Err.Description, ManagedConfigName(config), vbTextCompare) > 0
End Function

Private Function InvalidWrite() As Boolean
    On Error GoTo Expected
    Call WriteManagedConfigValue(mcServer, "INIT", "ServerSoloGMs", "1")
    Exit Function
Expected:
    InvalidWrite = True
End Function

Private Function InvalidIntervalSave() As Boolean
    On Error GoTo Expected
    Call SaveIntervalos()
    Exit Function
Expected:
    InvalidIntervalSave = InStr(1, Err.Description, "intervalos.ini", vbTextCompare) > 0
End Function

' A dedicated UNIT_TEST entry point: no database, network listener or production resources.
Public Sub RunManagedConfigTests()
    On Error GoTo Failed
    Dim fs As New FileSystemObject
    Dim root As String, external As String, legacy As String, resources As String
    Dim commandLine As String, present As Boolean, value As String
    Dim rootCreated As Boolean
    Dim config As e_ManagedConfig, path As String, saved As String
    Dim handle As Integer, locked As Boolean, reader As clsIniManager
    Dim serverText As String, mainText As String, featureText As String, balanceText As String, eventsText As String
    Dim intervalsText As String
    Dim i As Long, row As String, beforeError As Long
    root = fs.BuildPath(fs.GetSpecialFolder(2), "AO managed config " & fs.GetTempName())
    Call fs.CreateFolder(root)
    rootCreated = True
    external = root & "\External Config"
    legacy = root & "\Legacy"
    resources = root & "\Resources\"
    Call fs.CreateFolder(external)
    Call fs.CreateFolder(legacy)
    Call fs.CreateFolder(resources)
    Call ChDrive(root)
    Call ChDir(root)
    IniPath = legacy & "\"
    DatPath = resources
    Call InitializeCircularLogBuffer()
    Call CompleteConfigStartup()

    value = StartupOption("/other=S1 /config-dir=" & Chr$(34) & external & "\" & Chr$(34) & " --other=value", "/config-dir", present)
    Call Check("quoted directory and trailing slash", present And value = external & "\")
    value = StartupOption("/config-dir=" & Chr$(34) & external & Chr$(34) & " /other=S2", "/other", present)
    Call Check("unrelated option follows directory", present And value = "S2")
    value = StartupOption("/other=S1 /CONFIG-DIR " & Chr$(34) & external & Chr$(34), "/config-dir", present)
    Call Check("case-insensitive separated value", present And value = external)
    value = StartupOption("--unrelated /other=S1", "/config-dir", present)
    Call Check("absent option", Not present And value = vbNullString)
    value = StartupOption("/config-dir=" & Chr$(34) & external & "\/other=not-an-option" & Chr$(34), "/other", present)
    Call Check("no option matching inside path", Not present)
    Call Check("empty value rejected", InvalidOption("/config-dir="))
    Call Check("empty quotes rejected", InvalidOption("/config-dir=" & Chr$(34) & Chr$(34)))
    Call Check("missing value rejected", InvalidOption("/config-dir /other=S1"))
    Call Check("unclosed quote rejected", InvalidOption("/config-dir=" & Chr$(34) & external))
    Call Check("duplicate option rejected", InvalidOption("/config-dir=a /config-dir=b"))
    Call Check("missing directory rejected", InvalidDirectory("/config-dir=" & Chr$(34) & root & "\absent" & Chr$(34)))

    Call InitializeManagedConfig("/other=S1")
    For config = mcServer To mcIntervals
        If config = mcBalance Or config = mcEvents Then path = resources Else path = IniPath
        Call Check("legacy path " & ManagedConfigName(config), ManagedConfigPath(config) = path & ManagedConfigName(config))
    Next config
    intervalsText = "[INTERVALOS]" & vbCrLf & "SanaIntervaloSinDescansar=12" & vbCrLf & "IntervaloSed=250" & vbCrLf & "IntervaloHambre=500" & vbCrLf & "TimeoutPrimerPaquete=125" & vbCrLf & "IntervaloNpcAI=175" & vbCrLf & "IntervaloWS=3" & vbCrLf & "IntervalUnderworldSpawn=1000"
    Call PutFixture(IniPath & "intervalos.ini", intervalsText)
    Call PutFixture(root & "\intervalos.ini", Replace$(intervalsText, "SanaIntervaloSinDescansar=12", "SanaIntervaloSinDescansar=99"))
    Call LoadIntervalos()
    Call Check("legacy intervals use IniPath not current directory", SanaIntervaloSinDescansar = 12)
    Call LoadFeatureToggles()
    Call Check("initial missing legacy toggles remain optional", Not IsFeatureEnabled("unit-config"))
    ' A file in the current directory must not override the configured legacy base.
    Call PutFixture(root & "\feature_toggle.ini", "[INIT]" & vbCrLf & "TOGGLECOUNT=0")
    featureText = "[INIT]" & vbCrLf & "TOGGLECOUNT=1" & vbCrLf & "[TOGGLE1]" & vbCrLf & "name=unit-config" & vbCrLf & "value=1"
    Call PutFixture(IniPath & "feature_toggle.ini", featureText)
    Call LoadFeatureToggles()
    Call Check("legacy toggles independent of current directory", IsFeatureEnabled("unit-config"))

    serverText = "[INIT]" & vbCrLf & "MaxUsers=1" & vbCrLf & "StartPort=7654" & vbCrLf & "Hide=0" & vbCrLf & "ResPos=1-2-3" & vbCrLf & "Admines=1" & vbCrLf & "[Admines]" & vbCrLf & "Admin1=TESTGM|test@example.invalid" & vbCrLf & "[DIRECTORIOS]" & vbCrLf & "DatPath=" & resources
    mainText = "[CONFIGURACIONES]" & vbCrLf & "ExpMult=2" & vbCrLf & "[EVENTOS]" & vbCrLf & "0=1-5-2"
    balanceText = "[EXTRA]" & vbCrLf & "RecoveryMana=10" & vbCrLf & "[ElementalMatrixForNpcs]"
    For i = 1 To MAX_ELEMENT_TAGS
        row = row & "1 "
    Next i
    For i = 1 To MAX_ELEMENT_TAGS
        balanceText = balanceText & vbCrLf & "Row" & i & "=" & Trim$(row)
    Next i
    eventsText = "[Invasiones]" & vbCrLf & "Cantidad=1" & vbCrLf & "Invasion1=Test" & vbCrLf & "[Test]" & vbCrLf & "MaxNPCs=1" & vbCrLf & "MaxVidaMuralla=100" & vbCrLf & "SpawnBoxes=1" & vbCrLf & "SpawnBox1=1-1-1-2-2-NORTE-1"
    Call PutFixture(external & "\Server.ini", serverText)
    Call PutFixture(external & "\Configuracion.ini", mainText)
    Call PutFixture(external & "\feature_toggle.ini", featureText)
    Call PutFixture(external & "\Balance.dat", balanceText)
    Call PutFixture(external & "\Eventos.dat", eventsText)
    Call PutFixture(external & "\intervalos.ini", Replace$(intervalsText, "SanaIntervaloSinDescansar=12", "SanaIntervaloSinDescansar=24"))
    Call PutFixture(resources & "Ciudades.dat", "[INIT]")
    Call PutFixture(resources & "Retos.dat", "[Salas]" & vbCrLf & "Cantidad=0")
    Call PutFixture(resources & "MapasEspeciales.dat", "[MapasEventos]" & vbCrLf & "Cantidad=0")
    commandLine = "/other=S1 /config-dir=" & Chr$(34) & external & Chr$(34)
    Call InitializeManagedConfig(commandLine)
    Call Check("external selected", ExternalConfigEnabled())
    Call InitializeManagedConfig("/config-dir=" & Chr$(34) & "External Config\" & Chr$(34))
    Call Check("relative directory frozen as absolute", ManagedConfigPath(mcServer) = external & "\Server.ini")
    Call InitializeManagedConfig(commandLine)
    Call Check("global bases untouched", IniPath = legacy & "\" And DatPath = resources)
    For config = mcServer To mcIntervals
        path = external & "\" & ManagedConfigName(config)
        Call Check("external path " & ManagedConfigName(config), ManagedConfigPath(config) = path)
        saved = path & ".saved"
        Call fs.MoveFile(path, saved)
        Call Check("missing preflight " & ManagedConfigName(config), InvalidDirectory(commandLine))
        Call Check("missing read no fallback " & ManagedConfigName(config), Unreadable(config))
        Call Check("failed selection keeps prior directory", ManagedConfigPath(config) = path)
        Call Check("read did not recreate missing file", Not fs.FileExists(path))
        Call fs.MoveFile(saved, path)
        handle = FreeFile()
        Open path For Binary Access Read Write Lock Read Write As #handle
        locked = True
        Call Check("locked preflight " & ManagedConfigName(config), InvalidDirectory(commandLine))
        Call Check("locked read " & ManagedConfigName(config), Unreadable(config))
        Close #handle
        locked = False
    Next config
    Call LoadIntervalos()
    Call Check("external intervals override legacy file", SanaIntervaloSinDescansar = 24)
    Call Check("interval units unchanged", IntervaloSed = 10 And IntervaloHambre = 20 And TimeoutPrimerPaquete = 125 And FrmInterv.txtTimeoutPrimerPaquete.Text = "5")
    Call Check("AI and world-save intervals unchanged", IntervaloNPCAI = 175 And MinutosWs = 3 And IntervaloNpcOwner = 10000 And IntervalUnderworldSpawn = 1000)
    Call PutFixture(external & "\intervalos.ini", Replace$(Replace$(intervalsText, "SanaIntervaloSinDescansar=12", "SanaIntervaloSinDescansar=36"), "IntervalUnderworldSpawn=1000", "IntervalUnderworldSpawn=2147483648"))
    beforeError = frmMain.List1.ListCount
    Call LoadIntervalos()
    Call Check("late interval overflow retains earlier globals and UI and reports", SanaIntervaloSinDescansar = 24 And IntervalUnderworldSpawn = 1000 And FrmInterv.txtSanaIntervaloSinDescansar.Text = "24" And frmMain.List1.ListCount > beforeError)
    Call PutFixture(external & "\intervalos.ini", Replace$(intervalsText, "SanaIntervaloSinDescansar=12", "SanaIntervaloSinDescansar=32768"))
    Call LoadIntervalos()
    Call Check("interval Integer bounds unchanged", SanaIntervaloSinDescansar = 24)
    Call PutFixture(external & "\intervalos.ini", vbNullString)
    Call LoadIntervalos()
    Call Check("empty interval reload retains state", SanaIntervaloSinDescansar = 24 And IntervaloNPCAI = 175)
    Call fs.MoveFile(external & "\intervalos.ini", external & "\intervals.saved")
    beforeError = frmMain.List1.ListCount
    Call LoadIntervalos()
    Call Check("missing interval reload reports without legacy fallback", SanaIntervaloSinDescansar = 24 And frmMain.List1.ListCount > beforeError)
    Call Check("missing interval save does not create file", InvalidIntervalSave() And Not fs.FileExists(external & "\intervalos.ini"))
    Call fs.MoveFile(external & "\intervals.saved", external & "\intervalos.ini")
    Call PutFixture(external & "\intervalos.ini", intervalsText)
    handle = FreeFile()
    Open external & "\intervalos.ini" For Binary Access Read Write Lock Read Write As #handle
    locked = True
    beforeError = frmMain.List1.ListCount
    Call LoadIntervalos()
    Call Check("locked interval reload retains state and reports", SanaIntervaloSinDescansar = 24 And frmMain.List1.ListCount > beforeError)
    Close #handle
    locked = False
    Call PutFixture(external & "\intervalos.ini", Replace$(Replace$(intervalsText, "SanaIntervaloSinDescansar=12", "SanaIntervaloSinDescansar=36"), "IntervalUnderworldSpawn=1000", "IntervalUnderworldSpawn=2000"))
    Call LoadIntervalos()
    Call Check("corrected interval reload publishes", SanaIntervaloSinDescansar = 36 And IntervalUnderworldSpawn = 2000 And FrmInterv.txtSanaIntervaloSinDescansar.Text = "36")
    Call SaveIntervalos()
    Set reader = ReadManagedConfig(mcIntervals)
    Call Check("interval save uses external file and original units", reader.GetValue("INTERVALOS", "SanaIntervaloSinDescansar") = "36" And reader.GetValue("INTERVALOS", "IntervaloSed") = "250" And reader.GetValue("INTERVALOS", "IntervaloHambre") = "500")
    Call Check("interval save leaves legacy file unchanged", GetVar(IniPath & "intervalos.ini", "INTERVALOS", "SanaIntervaloSinDescansar") = "12" And Not fs.FileExists(IniPath & "intervalo.ini"))
    Call LoadIntervalos()
    Call Check("interval save reload roundtrip", IntervaloSed = 10 And IntervaloHambre = 20 And SanaIntervaloSinDescansar = 36 And TimeoutPrimerPaquete = 125 And IntervaloNPCAI = 175)
    Call SetAttr(external & "\intervalos.ini", vbReadOnly)
    Call Check("read-only interval save reports failure", InvalidIntervalSave())
    Call SetAttr(external & "\intervalos.ini", vbNormal)
    Call PutFixture(external & "\intervalos.ini", "[INTERVALOS]")
    Call LoadIntervalos()
    Call Check("missing interval keys retain classic defaults", IntervaloNPCAI = DEFAULT_NPC_AI_INTERVAL_MS And MinutosWs = 10 And IntervaloNpcOwner = 10000 And SanaIntervaloSinDescansar = 0)
    Call PutFixture(external & "\intervalos.ini", "[INTERVALOS]" & vbCrLf & "IntervaloNpcAI=0" & vbCrLf & "IntervaloWS=0")
    Call LoadIntervalos()
    Call Check("nonpositive AI and world-save defaults unchanged", IntervaloNPCAI = DEFAULT_NPC_AI_INTERVAL_MS And MinutosWs = 10)
    ' Real loaders: success, malformed conversion after earlier fields, then corrected reload.
    MaxUsers = 1
    Call LoadSini()
    Call Check("Server.ini external values", Puerto = 7654 And ResPos.x = 2 And DatPath = resources)
    Call PutFixture(external & "\Server.ini", Replace$(Replace$(serverText, "StartPort=7654", "StartPort=8765"), "Hide=0", "Hide=999"))
    beforeError = frmMain.List1.ListCount
    Call LoadSini()
    Call Check("Server.ini failed reload retains earlier values and reports", Puerto = 7654 And frmMain.List1.ListCount > beforeError)
    Call PutFixture(external & "\Server.ini", serverText)
    Call loadAdministrativeUsers()
    Call Check("external administrative users", AdministratorAccounts.Exists("TESTGM"))
    Call PutFixture(external & "\Server.ini", Replace$(serverText, "Admines=1", "Admines=99999"))
    Call loadAdministrativeUsers()
    Call Check("admin failed reload retains dictionary", AdministratorAccounts.Exists("TESTGM"))
    Call PutFixture(external & "\Server.ini", serverText)
    Call LoadFeatureToggles()
    Call PutFixture(external & "\feature_toggle.ini", Replace$(featureText, "TOGGLECOUNT=1", "TOGGLECOUNT=99999"))
    Call LoadFeatureToggles()
    Call Check("features failed reload retains dictionary", IsFeatureEnabled("unit-config"))
    Call PutFixture(external & "\feature_toggle.ini", Replace$(featureText, "value=1", "value=0"))
    Call LoadFeatureToggles()
    Call Check("features successful reload publishes", Not IsFeatureEnabled("unit-config"))
    Call fs.MoveFile(external & "\feature_toggle.ini", external & "\feature.saved")
    Call LoadFeatureToggles()
    Call Check("missing feature reload does not fall back to legacy true", Not IsFeatureEnabled("unit-config"))
    Call fs.MoveFile(external & "\feature.saved", external & "\feature_toggle.ini")
    Call LoadBalance()
    Call Check("balance initial load", RecoveryMana = 10)
    Call PutFixture(external & "\Balance.dat", Replace$(balanceText, "RecoveryMana=10", "RecoveryMana=20" & vbCrLf & "HomeTimer=99999"))
    Call LoadBalance()
    Call Check("balance late overflow leaves earlier scalar intact", RecoveryMana = 10)
    Call PutFixture(external & "\Balance.dat", Replace$(balanceText, "RecoveryMana=10", "RecoveryMana=20"))
    Call LoadBalance()
    Call Check("balance corrected reload publishes", RecoveryMana = 20)
    Call LoadMainConfigFile()
    Call Check("main and event initial load", SvrConfig.GetValue("ExpMult") = 2 And Evento(0).Tipo = 1 And Invasiones(1).MaxVidaMuralla = 100)
    Call PutFixture(external & "\Configuracion.ini", Replace$(mainText, "0=1-5-2", "0=" & Chr$(34) & "1-5-2" & Chr$(34)))
    Call CargarEventos()
    Call Check("scheduled events retain Win32 quoted-value semantics", Evento(0).Tipo = 1 And Evento(0).Duracion = 5)
    Call PutFixture(external & "\Configuracion.ini", Replace$(mainText, "ExpMult=2", "ExpMult=2" & vbCrLf & "UnassistedSpellsAllowed=" & Chr$(34) & "1,5,52" & Chr$(34)))
    Call Check("spell allowlist retains Win32 quoted-value semantics", ReadManagedProfileValue(mcMain, "CONFIGURACIONES", "UnassistedSpellsAllowed") = "1,5,52")
    Call PutFixture(external & "\Configuracion.ini", Replace$(mainText, "ExpMult=2", "ExpMult=99999"))
    Call LoadMainConfigFile()
    Call Check("main conversion error keeps previous object", SvrConfig.GetValue("ExpMult") = 2)
    Call PutFixture(external & "\Configuracion.ini", Replace$(mainText, "0=1-5-2", "0=1-999-2"))
    Call CargarEventos()
    Call Check("scheduled event failed reload retains values", Evento(0).Duracion = 5)
    Call PutFixture(external & "\Configuracion.ini", Replace$(mainText, "ExpMult=2", "ExpMult=3"))
    Call PutFixture(external & "\Eventos.dat", Replace$(eventsText, "SpawnBoxes=1", "SpawnBoxes=99999"))
    Call LoadMainConfigFile()
    Call Check("invasion failure does not publish new main config", SvrConfig.GetValue("ExpMult") = 2 And Invasiones(1).MaxVidaMuralla = 100)
    Call CargarInfoEventos()
    Call Check("invasion standalone failed reload retains state", Invasiones(1).MaxVidaMuralla = 100)
    Call PutFixture(external & "\Eventos.dat", Replace$(eventsText, "MaxVidaMuralla=100", "MaxVidaMuralla=200"))
    Call LoadMainConfigFile()
    Call Check("corrected main and invasion reload", SvrConfig.GetValue("ExpMult") = 3 And Invasiones(1).MaxVidaMuralla = 200)
    Call WriteManagedConfigValue(mcServer, "INIT", "ServerSoloGMs", "1")
    Set reader = ReadManagedConfig(mcServer)
    Call Check("managed write reaches external file", reader.GetValue("INIT", "ServerSoloGMs") = "1" And Not fs.FileExists(IniPath & "Server.ini"))
    Call SetAttr(external & "\Server.ini", vbReadOnly Or vbHidden)
    Set reader = ReadManagedConfig(mcServer)
    Call Check("read-only hidden config remains readable", reader.GetValue("INIT", "ServerSoloGMs") = "1")
    Call Check("read-only managed write reports failure", InvalidWrite())
    Call SetAttr(external & "\Server.ini", vbNormal)
    Call fs.MoveFile(external & "\Server.ini", external & "\server.saved")
    Call Check("failed write cannot recreate missing file", InvalidWrite() And Not fs.FileExists(external & "\Server.ini"))
    Call fs.MoveFile(external & "\server.saved", external & "\Server.ini")
    Call PutFixture(external & "\Balance.dat", vbNullString)
    Call Check("empty configuration rejected", Unreadable(mcBalance))
    Call PutFixture(external & "\Balance.dat", "not an INI")
    Call Check("invalid configuration rejected", Unreadable(mcBalance))
    Call PutFixture(external & "\Configuracion.ini", vbNullString)
    beforeError = frmMain.List1.ListCount
    Call CargarEventos()
    Call Check("empty scheduled-event reload retains state and reports", Evento(0).Duracion = 5 And frmMain.List1.ListCount > beforeError)
    GoTo Finished
Failed:
    mFailures = mFailures + 1
    mResults = mResults & "ERROR " & Err.Number & ": " & Err.Description & vbCrLf
Finished:
    On Error Resume Next
    If locked Then Close #handle
    Call ChDir(App.Path)
    If rootCreated Then Call fs.DeleteFolder(root, True)
    Call PutFixture(App.Path & "\managed_config_test_results.txt", mResults & "Checks=" & mChecks & " Failed=" & mFailures)
End Sub
#End If
