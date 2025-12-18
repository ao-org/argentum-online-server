Attribute VB_Name = "ModGlobalQuests"
Option Explicit

Public Type t_GlobalQuestData
    GatheringThreshold As Long
    GatheringGlobalCounter As Long
    GatheringGlobalInstallments As Long
    GatheringInitialInstallments As Long
    IsBossAlive As Boolean
    BossIndex As Integer
    BossSpawnPosition As t_WorldPos
    FinishOnThresholdReach As Boolean
    Name As String
    StartDate As Date
    EndDate As Date
    ObjectIndex As Integer
    IsActive As Boolean
End Type

Private m_GlobalQuestEndAttempt As Long
Public GlobalQuestInfo()        As t_GlobalQuestData

Public Sub ContributeToGlobalQuestCounter(ByVal Amount As Long, ByVal GlobalQuestIndex As Integer)
    With GlobalQuestInfo(GlobalQuestIndex)
        .GatheringGlobalCounter = .GatheringGlobalCounter + Amount
        If .GatheringGlobalCounter >= .GatheringGlobalInstallments Then
            .GatheringGlobalInstallments = .GatheringGlobalInstallments + .GatheringInitialInstallments
            If Not .IsBossAlive Then
                Call SpawnNpc(.BossIndex, .BossSpawnPosition, False, False, True, 0)
                .IsBossAlive = True
            End If
        End If
    End With
End Sub

Public Sub InsertContributionIntoDatabase(ByVal UserIndex As Integer, ByVal Amount, ByVal GlobalQuestIndex As Integer)
    On Error GoTo InsertContributionIntoDatabase_Err
    Dim RS          As ADODB.Recordset
    Dim QueryString As String
    QueryString = "INSERT INTO global_quest_user_contribution (event_id,user_id,timestamp,amount) VALUES (?,?,?,?);"
    Set RS = Query(QueryString, GlobalQuestIndex, UserList(UserIndex).Id, CStr(DateTime.Now), Amount)
    Exit Sub
InsertContributionIntoDatabase_Err:
    Call TraceError(Err.Number, Err.Description, "ModGlobalQuests.InsertContributionIntoDatabase", Erl)
End Sub

Public Sub UpdateGlobalQuestActiveStateIntoDatabase(ByVal Status As Boolean, ByVal GlobalQuestIndex As Integer)
    On Error GoTo UpdateGlobalQuestActiveStateIntoDatabase_Err
    Dim RS          As ADODB.Recordset
    Dim QueryString As String
    QueryString = "UPDATE global_quest_desc SET is_active = ? WHERE event_id = ?;"
    Set RS = Query(QueryString, Status, GlobalQuestIndex)
    Exit Sub
UpdateGlobalQuestActiveStateIntoDatabase_Err:
    Call TraceError(Err.Number, Err.Description, "ModGlobalQuests.UpdateGlobalQuestActiveStateIntoDatabase", Erl)
End Sub

Public Sub LoadGlobalQuests()
    On Error GoTo LoadGlobalQuests_Err
    Dim MaxGlobalQuests
    If Not FileExist(DatPath & "GlobalQuests.dat", vbArchive) Then
        Debug.Assert False
        Exit Sub
    End If
    Dim GlobalQuest As Integer
    Dim IniFile     As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "GlobalQuests.dat")
    MaxGlobalQuests = val(IniFile.GetValue("INIT", "NumGlobalQuest"))
    ReDim Preserve GlobalQuestInfo(1 To MaxGlobalQuests) As t_GlobalQuestData
    Dim i As Integer
    For i = 1 To MaxGlobalQuests
        With GlobalQuestInfo(i)
            .GatheringThreshold = CLng(val(IniFile.GetValue("GlobalQuest" & i, "GatheringThreshold")))
            .GatheringInitialInstallments = CLng(val(IniFile.GetValue("GlobalQuest" & i, "GatheringInitialInstallments")))
            .GatheringGlobalInstallments = CLng(val(IniFile.GetValue("GlobalQuest" & i, "GatheringInitialInstallments")))
            .BossSpawnPosition.Map = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossSpawnPositionMap")))
            .BossSpawnPosition.x = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossSpawnPositionX")))
            .BossSpawnPosition.y = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossSpawnPositionY")))
            .BossIndex = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossIndex")))
            .FinishOnThresholdReach = val(IniFile.GetValue("GlobalQuest" & i, "FinishOnThresholdReach"))
            .Name = IniFile.GetValue("GlobalQuest" & i, "Name")
            .StartDate = CDate(IniFile.GetValue("GlobalQuest" & i, "StartDate"))
            .EndDate = CDate(IniFile.GetValue("GlobalQuest" & i, "EndDate"))
            .ObjectIndex = val(IniFile.GetValue("GlobalQuest" & i, "ObjectIndex"))
            .IsActive = False
            Dim RS As ADODB.Recordset
            Set RS = Query("SELECT * FROM global_quest_desc WHERE event_id = ?;", i)
            If RS Is Nothing Then Exit Sub
            'if global quest doesnt exist create it
            Dim QueryString As String
            If RS.RecordCount = 0 Then
                QueryString = "INSERT INTO global_quest_desc (event_id, name, obj_id, threshold, start_date, end_date, is_active) VALUES (?,?, ?, ?, ?, ?, ?);"
                Set RS = Query(QueryString, i, .Name, .ObjectIndex, .GatheringThreshold, DateToSQLite(.StartDate), DateToSQLite(.EndDate), False)
                'if exists load everything and reconstruct the current total user contribution
            Else
                .Name = RS!Name
                .ObjectIndex = RS!obj_id
                .GatheringThreshold = RS!threshold
                .StartDate = SQLiteToDate(RS!start_date)
                .EndDate = SQLiteToDate(RS!end_date)
                .IsActive = RS!is_active
                QueryString = "SELECT SUM(amount) AS total_amount FROM global_quest_user_contribution WHERE event_id = ?;"
                Set RS = Query(QueryString, i)
                If Not IsNull(RS!total_amount) Then
                    .GatheringGlobalCounter = RS!total_amount
                End If
            End If
        End With
    Next i
    Set IniFile = Nothing
    Exit Sub
LoadGlobalQuests_Err:
    Call TraceError(Err.Number, Err.Description, "ModGlobalQuests.LoadGlobalQuests", Erl)
End Sub

Public Function FinishGlobalQuestCheck(ByVal UserIndex As Integer, ByVal GlobalQuestIndex As Integer, ByVal GlobalQuestThresholdNeeded As Long) As Boolean
    'boss alive mechanics shoudln't interfer with unique prizes
    If GlobalQuestThresholdNeeded > 0 Then
        If GlobalQuestInfo(GlobalQuestIndex).GatheringGlobalCounter < GlobalQuestThresholdNeeded Then
            Call WriteLocaleMsg(UserIndex, 2123, FONTTYPE_WARNING, GlobalQuestInfo(GlobalQuestIndex).GatheringGlobalCounter & "¬" & GlobalQuestInfo(GlobalQuestIndex).GatheringThreshold & "¬" & GlobalQuestThresholdNeeded)
            Exit Function
        End If
        'global quest unique prizes should be redeemable even if the event is finished
        GoTo SkipEventIsActive
    Else
        If GlobalQuestInfo(GlobalQuestIndex).IsBossAlive Then
            Call WriteLocaleMsg(UserIndex, 2121, FONTTYPE_WARNING)
            Exit Function
        End If
    End If
    If Not GlobalQuestInfo(GlobalQuestIndex).IsActive Then
        Call WriteLocaleMsg(UserIndex, 2124, FONTTYPE_WARNING)
        Exit Function
    End If
SkipEventIsActive:
    FinishGlobalQuestCheck = True
End Function

Public Sub FinishGlobalQuest(ByVal UserIndex As Integer, ByVal ContributionAmount As Integer, ByVal GlobalQuestIndex As Integer, ByVal GlobalQuestThresholdNeeded As Long)
    'gathering threshold locked quests cannot also contribute to the global event
    If GlobalQuestIndex > 0 And GlobalQuestThresholdNeeded = 0 Then
        Call ContributeToGlobalQuestCounter(ContributionAmount, GlobalQuestIndex)
        Call InsertContributionIntoDatabase(UserIndex, ContributionAmount, GlobalQuestIndex)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(2122, UserList(UserIndex).Name & "¬" & ContributionAmount & "¬" & GlobalQuestInfo(GlobalQuestIndex).ObjectIndex & "¬" & GlobalQuestInfo(GlobalQuestIndex).GatheringGlobalCounter & "¬" & GlobalQuestInfo(GlobalQuestIndex).GatheringThreshold, e_FontTypeNames.FONTTYPE_INFOIAO))
        If GlobalQuestInfo(GlobalQuestIndex).FinishOnThresholdReach Then
            If GlobalQuestInfo(GlobalQuestIndex).GatheringGlobalCounter >= GlobalQuestInfo(GlobalQuestIndex).GatheringThreshold Then
                GlobalQuestInfo(GlobalQuestIndex).IsActive = False
                Call UpdateGlobalQuestActiveStateIntoDatabase(False, GlobalQuestIndex)
            End If
        End If
    End If
End Sub

Public Sub MaybeChangeGlobalQuestsState()
    On Error GoTo MaybeChangeGlobalQuestsState_Err
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If m_GlobalQuestEndAttempt = 0 Then
        m_GlobalQuestEndAttempt = nowRaw
        Exit Sub
    End If
    If TicksElapsed(m_GlobalQuestEndAttempt, nowRaw) < IntervalChangeGlobalQuestsState Then Exit Sub
    m_GlobalQuestEndAttempt = nowRaw
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    Dim i As Integer
    For i = 1 To UBound(GlobalQuestInfo)
        If GlobalQuestInfo(i).IsActive And HasGlobalQuestEnded(GlobalQuestInfo(i)) Then
            Call FinalizeGlobalQuest(i)
        ElseIf Not GlobalQuestInfo(i).IsActive And HasGlobalQuestStarted(GlobalQuestInfo(i)) Then
            Call StartGlobalQuest(i)
        End If
    Next i
    Exit Sub
MaybeChangeGlobalQuestsState_Err:
    Call TraceError(Err.Number, Err.Description, "ModGlobalQuests.MaybeChangeGlobalQuestsState", Erl)
End Sub

Public Function HasGlobalQuestEnded(ByRef GlobalQuestData As t_GlobalQuestData) As Boolean
    HasGlobalQuestEnded = GlobalQuestData.EndDate - DateTime.Now < 0
End Function

Public Function HasGlobalQuestStarted(ByRef GlobalQuestData As t_GlobalQuestData) As Boolean
    HasGlobalQuestStarted = DateTime.Now - GlobalQuestData.StartDate > 0
End Function

Public Sub FinalizeGlobalQuest(ByVal GlobalQuestIndex As Integer)
    With GlobalQuestInfo(GlobalQuestIndex)
        .IsActive = False
        Call UpdateGlobalQuestActiveStateIntoDatabase(False, GlobalQuestIndex)
        'TBD change map indexes and de-spawn corresponding npcs
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(2127, .Name & "¬" & .GatheringGlobalCounter & "¬" & .GatheringThreshold, e_FontTypeNames.FONTTYPE_INFOIAO))
    End With
End Sub

Public Sub StartGlobalQuest(ByVal GlobalQuestIndex As Integer)
    With GlobalQuestInfo(GlobalQuestIndex)
        .IsActive = True
        Call UpdateGlobalQuestActiveStateIntoDatabase(True, GlobalQuestIndex)
        'TBD change map indexes and spawn corresponding npcs
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(2128, .Name & "¬" & .GatheringThreshold & "¬" & .ObjectIndex, e_FontTypeNames.FONTTYPE_INFOIAO))
    End With
End Sub

Function SQLiteToDate(strDate As String) As Date
    ' For format "2025-12-18 14:30:00"
    SQLiteToDate = CDate(strDate)
End Function

Function DateToSQLite(dt As Date) As String
    DateToSQLite = Format$(dt, "yyyy-mm-dd hh:nn:ss")
End Function
