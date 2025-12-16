Attribute VB_Name = "ModSeasonalEvents"
Option Explicit

Public SeasonalEventGatheringGlobalThreshold As Long
Public SeasonalEventGatheringGlobalCounter As Long
Public SeasonalEventGatheringInstallment As Long

Public SeasonalEventIsBossAlive As Boolean
Public SeasonalEventBossIndex As Integer
Public SeasonalEventBossSpawnPosition As t_WorldPos

Public Sub ContributeToSeasonalEventGlobalCounter(ByVal Amount As Long)
    SeasonalEventGatheringGlobalCounter = SeasonalEventGatheringGlobalCounter + Amount
    If SeasonalEventGatheringGlobalCounter >= SeasonalEventGatheringInstallment And SeasonalEventGatheringGlobalCounter Mod SeasonalEventGatheringInstallment <= 1 Then
        If Not SeasonalEventIsBossAlive Then
            Call SpawnNpc(SeasonalEventBossIndex, SeasonalEventBossSpawnPosition, False, False, True, 0)
            SeasonalEventIsBossAlive = True
        End If
    End If
End Sub
