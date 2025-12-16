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

Public GlobalQuestInfo As t_GlobalQuestData

Public Sub ContributeToGlobalQuestGlobalCounter(ByVal Amount As Long)
    With GlobalQuestInfo
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
    With GlobalQuestInfo
        If Not FileExist(DatPath & "GlobalQuests.dat", vbArchive) Then
            Exit Sub
        End If
        Dim IniFile As clsIniManager
        Set IniFile = New clsIniManager
        Call IniFile.Initialize(DatPath & "GlobalQuests.dat")
        .GatheringThreshold = CLng(val(IniFile.GetValue("GlobalQuest", "GatheringThreshold")))
        .GatheringInitialInstallments = CLng(val(IniFile.GetValue("GlobalQuest", "GatheringInitialInstallments")))
        .GatheringGlobalInstallments = CLng(val(IniFile.GetValue("GlobalQuest", "GatheringInitialInstallments")))
        .BossSpawnPosition.Map = CInt(val(IniFile.GetValue("GlobalQuest", "BossSpawnPositionMap")))
        .BossSpawnPosition.x = CInt(val(IniFile.GetValue("GlobalQuest", "BossSpawnPositionX")))
        .BossSpawnPosition.y = CInt(val(IniFile.GetValue("GlobalQuest", "BossSpawnPositionY")))
        .BossIndex = CInt(val(IniFile.GetValue("GlobalQuest", "BossIndex")))
        .FinishOnThresholdReach = 0
        Set IniFile = Nothing
    End With
    Exit Sub
LoadGlobalQuests_Err:
    Call TraceError(Err.Number, Err.Description, "ES.LoadGlobalQuests", Erl)
End Sub
