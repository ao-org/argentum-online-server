Attribute VB_Name = "ModGlobalQuests"
Option Explicit

Public GlobalQuestGatheringThreshold As Long
Public GlobalQuestGatheringGlobalCounter As Long
Public GlobalQuestGatheringGlobalInstallments As Long

Public GlobalQuestGatheringInitialInstallments As Long


Public GlobalQuestIsBossAlive As Boolean
Public GlobalQuestBossIndex As Integer
Public GlobalQuestBossSpawnPosition As t_WorldPos

Public Sub ContributeToGlobalQuestGlobalCounter(ByVal Amount As Long)
    GlobalQuestGatheringGlobalCounter = GlobalQuestGatheringGlobalCounter + Amount
    If GlobalQuestGatheringGlobalCounter >= GlobalQuestGatheringGlobalInstallments Then
        GlobalQuestGatheringGlobalInstallments = GlobalQuestGatheringGlobalInstallments + GlobalQuestGatheringInitialInstallments
        If Not GlobalQuestIsBossAlive Then
            Call SpawnNpc(GlobalQuestBossIndex, GlobalQuestBossSpawnPosition, False, False, True, 0)
            GlobalQuestIsBossAlive = True
        End If
    End If
End Sub
