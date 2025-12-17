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
    StartDate As String
    EndDate As String
    ObjectIndex As Integer
End Type

Public GlobalQuestInfo() As t_GlobalQuestData

Public Sub ContributeToGlobalQuestGlobalCounter(ByVal Amount As Long, ByVal GlobalQuestIndex)
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
            .StartDate = IniFile.GetValue("GlobalQuest" & i, "StartDate")
            .EndDate = IniFile.GetValue("GlobalQuest" & i, "EndDate")
            .ObjectIndex = val(IniFile.GetValue("GlobalQuest" & i, "ObjectIndex"))
            Dim RS As ADODB.Recordset
            Set RS = Query("SELECT * FROM global_quest_desc WHERE event_id = ?;", i)
            If RS Is Nothing Then Exit Sub
            'if global quest doesnt exist create it
            Dim QueryString As String
            If RS.RecordCount = 0 Then
                QueryString = "INSERT INTO global_quest_desc (event_id, name, obj_id, threshold, start_date, end_date) VALUES (?,?, ?, ?, ?, ?);"
                Set RS = Query(QueryString, i, .Name, .ObjectIndex, .GatheringThreshold, .StartDate, .EndDate)
                'if exists load everything and reconstruct the current total user contribution
            Else
                .Name = RS!Name
                .ObjectIndex = RS!obj_id
                .GatheringThreshold = RS!threshold
                .StartDate = RS!start_date
                .EndDate = RS!end_date
                QueryString = "SELECT SUM(amount) AS total_amount FROM global_quest_user_contribution WHERE event_id = ?;"
                Set RS = Query(QueryString, i)
                If RS!total_amount <> Null Then
                    .GatheringGlobalCounter = RS!total_amount
                End If
            End If
        End With
    Next i
    Set IniFile = Nothing
    Exit Sub
LoadGlobalQuests_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadGlobalQuests", Erl)
End Sub
