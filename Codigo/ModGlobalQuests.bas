Attribute VB_Name = "ModGlobalQuests"
Option Explicit

Public Type t_GlobalQuestData
    GatheringThreshold As Long
    GatheringGlobalCounter As Long
    GatheringGlobalInstallments As Long
    GatheringInitialInstallments As Long
    IsBossAlive As Boolean
    BossIndex As Integer
    BossSpawnMap            As Integer
    BossSpawnPositionTopLeft As t_Position
    BossSpawnPositionBottomRight As t_Position
    FinishOnThresholdReach As Boolean
    Name As String
    StartDate As Date
    EndDate As Date
    ObjectIndex As Integer
    IsActive As Boolean
End Type

Private m_GlobalQuestEndAttempt                       As Long
Public GlobalQuestInfo()                              As t_GlobalQuestData
Private Const INSERT_GLOBAL_QUEST_USER_CONTRIBUTION   As String = "INSERT INTO global_quest_user_contribution (event_id,user_id,timestamp,amount) VALUES (?, ?, ?, ?);"
Private Const UPDATE_GLOBAL_QUEST_DESC                As String = "UPDATE global_quest_desc SET is_active = ? WHERE event_id = ?;"
Private Const MODIFY_GLOBAL_QUEST_DESC                As String = "UPDATE global_quest_desc SET name = ?, obj_id = ?, threshold = ?, start_date = ?, end_date = ? WHERE event_id = ?;"
Private Const INSERT_NEW_GLOBAL_QUEST_DESC            As String = "INSERT INTO global_quest_desc (event_id, name, obj_id, threshold, start_date, end_date, is_active) VALUES (?,?, ?, ?, ?, ?, ?);"
Private Const SELECT_ALL_GLOBAL_QUEST                 As String = "SELECT * FROM global_quest_desc WHERE event_id = ?;"
Private Const SUM_TOTAL_AMOUNT_FROM_USER_CONTRIBUTION As String = "SELECT SUM(amount) AS total_amount FROM global_quest_user_contribution WHERE event_id = ?;"

Public Sub ContributeToGlobalQuestCounter(ByVal Amount As Long, ByVal GlobalQuestIndex As Integer)
    With GlobalQuestInfo(GlobalQuestIndex)
        .GatheringGlobalCounter = .GatheringGlobalCounter + Amount
        If .GatheringGlobalCounter >= .GatheringGlobalInstallments Then
            .GatheringGlobalInstallments = .GatheringGlobalInstallments + .GatheringInitialInstallments
            If Not .IsBossAlive Then
                Dim RandomizedSpawnPosition As t_WorldPos
                RandomizedSpawnPosition.Map = .BossSpawnMap
                RandomizedSpawnPosition.x = RandomNumber(.BossSpawnPositionTopLeft.x, .BossSpawnPositionBottomRight.x)
                RandomizedSpawnPosition.y = RandomNumber(.BossSpawnPositionTopLeft.y, .BossSpawnPositionBottomRight.y)
                Call SpawnNpc(.BossIndex, RandomizedSpawnPosition, False, False, True, 0)
                .IsBossAlive = True
            End If
        End If
    End With
End Sub

Public Sub InsertContributionIntoDatabase(ByVal UserIndex As Integer, ByVal Amount, ByVal GlobalQuestIndex As Integer)
    On Error GoTo InsertContributionIntoDatabase_Err
    Dim RS As ADODB.Recordset
    Set RS = Query(INSERT_GLOBAL_QUEST_USER_CONTRIBUTION, GlobalQuestIndex, UserList(UserIndex).Id, CStr(DateTime.Now), Amount)
    Exit Sub
InsertContributionIntoDatabase_Err:
    Call TraceError(Err.Number, Err.Description, "ModGlobalQuests.InsertContributionIntoDatabase", Erl)
End Sub

Public Sub UpdateGlobalQuestActiveStateIntoDatabase(ByVal Status As Boolean, ByVal GlobalQuestIndex As Integer)
    On Error GoTo UpdateGlobalQuestActiveStateIntoDatabase_Err
    Dim RS As ADODB.Recordset
    Set RS = Query(UPDATE_GLOBAL_QUEST_DESC, Status, GlobalQuestIndex)
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
            .BossSpawnMap = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossSpawnMap")))
            .BossSpawnPositionBottomRight.x = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossSpawnPositionBottomRightX")))
            .BossSpawnPositionBottomRight.y = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossSpawnPositionBottomRightY")))
            .BossSpawnPositionTopLeft.x = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossSpawnPositionTopLeftX")))
            .BossSpawnPositionTopLeft.y = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossSpawnPositionTopLeftY")))
            .BossIndex = CInt(val(IniFile.GetValue("GlobalQuest" & i, "BossIndex")))
            .FinishOnThresholdReach = val(IniFile.GetValue("GlobalQuest" & i, "FinishOnThresholdReach"))
            .Name = IniFile.GetValue("GlobalQuest" & i, "Name")
            .StartDate = CDate(IniFile.GetValue("GlobalQuest" & i, "StartDate"))
            .EndDate = CDate(IniFile.GetValue("GlobalQuest" & i, "EndDate"))
            .ObjectIndex = val(IniFile.GetValue("GlobalQuest" & i, "ObjectIndex"))
            .IsActive = False
            Dim RS As ADODB.Recordset
            Set RS = Query(SELECT_ALL_GLOBAL_QUEST, i)
            If RS Is Nothing Then Exit Sub
            'if global quest doesnt exist create it
            If RS.RecordCount = 0 Then
                Set RS = Query(INSERT_NEW_GLOBAL_QUEST_DESC, i, .Name, .ObjectIndex, .GatheringThreshold, DateToSQLite(.StartDate), DateToSQLite(.EndDate), False)
                'if exists load everything and reconstruct the current total user contribution
            Else
                .Name = RS!Name
                .ObjectIndex = RS!obj_id
                .GatheringThreshold = RS!threshold
                .StartDate = SQLiteToDate(RS!start_date)
                .EndDate = SQLiteToDate(RS!end_date)
                .IsActive = RS!is_active
                Set RS = Query(SUM_TOTAL_AMOUNT_FROM_USER_CONTRIBUTION, i)
                If Not IsNull(RS!total_amount) Then
                    .GatheringGlobalCounter = RS!total_amount
                    .GatheringGlobalInstallments = RS!total_amount + .GatheringInitialInstallments
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
        'if the end date is programmed to be in the future
        If GlobalQuestInfo(i).IsActive And HasGlobalQuestEnded(GlobalQuestInfo(i)) And Not GlobalQuestInfo(i).FinishOnThresholdReach Then
            Call FinalizeGlobalQuest(i)
        ElseIf Not GlobalQuestInfo(i).IsActive And HasGlobalQuestStarted(GlobalQuestInfo(i)) And Not HasGlobalQuestEnded(GlobalQuestInfo(i)) Then
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

Public Function IsGlobalQuestInTheFuture(ByRef GlobalQuestData As t_GlobalQuestData) As Boolean
    IsGlobalQuestInTheFuture = GlobalQuestData.EndDate - DateTime.Now > 0
End Function

Public Sub FinalizeGlobalQuest(ByVal GlobalQuestIndex As Integer)
    If GlobalQuestIndex < LBound(GlobalQuestInfo) Or GlobalQuestIndex > UBound(GlobalQuestInfo) Then Exit Sub
    Debug.Assert Not HasGlobalQuestEnded(GlobalQuestInfo(GlobalQuestIndex))
    Debug.Assert HasGlobalQuestStarted(GlobalQuestInfo(GlobalQuestIndex))
    If (Not GlobalQuestInfo(GlobalQuestIndex).IsActive) Then
        LogError "Calling FinalizeGlobalQueston a quest that has already finished"
        Exit Sub
    End If
    With GlobalQuestInfo(GlobalQuestIndex)
        .IsActive = False
        Call UpdateGlobalQuestActiveStateIntoDatabase(False, GlobalQuestIndex)
        'TBD change map indexes and de-spawn corresponding npcs
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(2127, .Name & "¬" & .GatheringGlobalCounter & "¬" & .GatheringThreshold, e_FontTypeNames.FONTTYPE_INFOIAO))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(e_SoundEffects.BAOLegionHorn, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

Public Sub StartGlobalQuest(ByVal GlobalQuestIndex As Integer)
    Debug.Assert Not HasGlobalQuestEnded(GlobalQuestInfo(GlobalQuestIndex))
    If GlobalQuestInfo(GlobalQuestIndex).IsActive And HasGlobalQuestStarted(GlobalQuestInfo(GlobalQuestIndex)) Then
        LogError "Calling StartGlobalQuest on a quest that has already started : " & GlobalQuestInfo(GlobalQuestIndex).Name
        Exit Sub
    End If
    With GlobalQuestInfo(GlobalQuestIndex)
        LogError "Starting GlobalQuest " & GlobalQuestIndex & " " & .Name
        .IsActive = True
        Call UpdateGlobalQuestActiveStateIntoDatabase(True, GlobalQuestIndex)
        'TBD change map indexes and spawn corresponding npcs
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(2128, .Name & "¬" & .GatheringThreshold & "¬" & .ObjectIndex, e_FontTypeNames.FONTTYPE_INFOIAO))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(e_SoundEffects.BAOLegionHorn, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

Function SQLiteToDate(strDate As String) As Date
    ' For format "2025-12-18 14:30:00"
    SQLiteToDate = CDate(strDate)
End Function

Function DateToSQLite(dt As Date) As String
    DateToSQLite = Format$(dt, "yyyy-mm-dd hh:nn:ss")
End Function

Public Sub HandleModifyGlobalQuest(ByVal UserIndex As Integer)
    '/modglobalquest GLOBALQUESTINDEX STARTDATE ENDDATE NAME OBJINDEX GATHERINGTHRESHOLD
    Dim GlobalQuestIndex      As Integer
    Dim newStartDate          As Date
    Dim newEndDate            As Date
    Dim newName               As String
    Dim newObjIndex           As Integer
    Dim newGatheringThreshold As Long
    If Not EsGM(UserIndex) Then Exit Sub
    If (UserList(UserIndex).flags.Privilegios And e_PlayerType.Admin) = 0 Then
        Exit Sub
    End If
    GlobalQuestIndex = reader.ReadInt16()
    newStartDate = reader.ReadString16()
    newEndDate = reader.ReadString16()
    newName = reader.ReadString16()
    newObjIndex = reader.ReadInt16()
    newGatheringThreshold = reader.ReadInt32()
    If GlobalQuestIndex > UBound(GlobalQuestInfo) Or GlobalQuestIndex < LBound(GlobalQuestInfo) Then
        Call LogError("Invalid index for given global quest, it's out of bounds")
        Exit Sub
    End If
    GlobalQuestInfo(GlobalQuestIndex).StartDate = newStartDate
    GlobalQuestInfo(GlobalQuestIndex).EndDate = newEndDate
    GlobalQuestInfo(GlobalQuestIndex).Name = newName
    GlobalQuestInfo(GlobalQuestIndex).ObjectIndex = newObjIndex
    GlobalQuestInfo(GlobalQuestIndex).GatheringThreshold = newGatheringThreshold
    Dim RS As ADODB.Recordset
    Set RS = Query(MODIFY_GLOBAL_QUEST_DESC, newName, newObjIndex, newGatheringThreshold, DateToSQLite(newStartDate), DateToSQLite(newEndDate), GlobalQuestIndex)
End Sub
