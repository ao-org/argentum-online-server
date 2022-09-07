Attribute VB_Name = "CustomScenarios"
Option Explicit

Public Enum e_DamageSourceType
    e_phisical
    e_magic
    e_pet
End Enum

Private CustomScenarioList As New Dictionary
Private ScenarioUpdateList() As IBaseScenario

Public Function GetMap(ByVal mapIndex As Integer) As IBaseScenario
On Error GoTo GetMap_Err:
    Set GetMap = Nothing
    If CustomScenarioList.Exists(mapIndex) Then
        Set GetMap = CustomScenarioList.Item(mapIndex)
    End If
    Exit Function
GetMap_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.GetMap " & mapIndex, Erl)
End Function

Public Sub SetCustomScenario(ByVal mapIndex As Integer, ByRef scenario As IBaseScenario)
    Call CustomScenarioList.Add(mapIndex, scenario)
End Sub

Public Sub RemoveCustomScenario(ByVal mapIndex As Integer)
    Call CustomScenarioList.Remove(mapIndex)
End Sub

Public Function AddUpdateScenario(ByRef scenario As IBaseScenario) As Integer
On Error GoTo AddUpdateScenario_Err:
       Dim Pos As Integer
100    If IsArrayInitialized(ScenarioUpdateList) Then
102       Pos = UBound(ScenarioUpdateList)
       Else
           Pos = 0
       End If
104    ReDim Preserve ScenarioUpdateList(Pos + 1) As IBaseScenario
106    Set ScenarioUpdateList(Pos) = scenario
       AddUpdateScenario = Pos
       Exit Function
AddUpdateScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.AddUpdateScenario", Erl)
End Function

Public Sub RemoveUpdateScenario(ByRef Index As Integer)
    Set ScenarioUpdateList(Index) = Nothing
End Sub

Public Sub UpdateAll()
On Error GoTo UpdateAll_Err:
    If Not IsArrayInitialized(ScenarioUpdateList) Then
        Exit Sub
    End If
    Dim i As Integer
    For i = 0 To UBound(ScenarioUpdateList)
        If Not ScenarioUpdateList(i) Is Nothing Then
            Call ScenarioUpdateList(i).Update
        End If
    Next
    Exit Sub
UpdateAll_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UpdateAll", Erl)
End Sub

Public Sub Reset(ByVal MapNumber As Integer)
On Error GoTo Reset_Err:
    Dim scenario As IBaseScenario
    Set scenario = GetMap(MapNumber)
    If scenario Is Nothing Then
        Exit Sub
    End If
    Call scenario.Reset
    Exit Sub
Reset_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.Reset " & mapNumber, Erl)
End Sub

Public Sub PlayerKillNpc(ByVal MapNumber As Integer, ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
On Error GoTo PlayerKillNpc_Err:
    Dim scenario As IBaseScenario
    Set scenario = GetMap(MapNumber)
    If scenario Is Nothing Then
        Exit Sub
    End If
    Call scenario.PlayerKillNpc(npcIndex, userIndex, sourceType, sourceIndex)
    Exit Sub
PlayerKillNpc_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PlayerKillNpc map:" & mapNumber & " npc: " & npcIndex & " user: " & userIndex, Erl)
End Sub

Public Sub NPcKillPlayer(ByVal MapNumber As Integer, ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
On Error GoTo NPcKillPlayer_Err:
    Dim scenario As IBaseScenario
    Set scenario = GetMap(MapNumber)
    If scenario Is Nothing Then
        Exit Sub
    End If
    Call scenario.NPcKillPlayer(npcIndex, userIndex, sourceType, sourceIndex)
    Exit Sub
NPcKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.NPcKillPlayer", Erl)
End Sub

Public Sub PlayerKillPlayer(ByVal MapNumber As Integer, ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
On Error GoTo PlayerKillPlayer_Err:
    Dim scenario As IBaseScenario
    Set scenario = GetMap(MapNumber)
    If scenario Is Nothing Then
        Exit Sub
    End If
    Call scenario.PlayerKillPlayer(killerIndex, deadIndex, sourceType, sourceIndex)
    Exit Sub
PlayerKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PlayerKillPlayer", Erl)
End Sub

Public Sub PrepareNewEvent(ByVal eventType As e_EventType)
On Error GoTo PrepareNewEvent_Err:
    Select Case EventType
        Case e_EventType.NpcHunt
            Set GenericGlobalLobby.scenario = New ScenarioHunt
    End Select
    Exit Sub
PrepareNewEvent_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PrepareNewEvent", Erl)
End Sub

