Attribute VB_Name = "CustomScenarios"
Option Explicit

Public Enum e_DamageSourceType
    e_phisical
    e_magic
    e_pet
End Enum

Private CustomScenarioList As New Dictionary
Private ScenarioUpdateList() As IBaseScenario

Public Sub InitializeCustomMaps()
    Set CustomScenarioList = New Dictionary
    Dim ScenarioUpdateList(0) As IBaseScenario
End Sub

Public Function GetMap(ByVal mapIndex As Integer) As IBaseScenario
    Set GetMap = Nothing
    If CustomScenarioList.Exists(mapIndex) Then
        Set GetMap = CustomScenarioList.Item(mapIndex)
    End If
End Function

Public Sub SetCustomScenario(ByVal mapIndex As Integer, ByRef scenario As IBaseScenario)
    Call CustomScenarioList.Add(mapIndex, scenario)
End Sub

Public Function AddUpdateScenario(ByRef scenario As IBaseScenario) As Integer
    Dim pos As Integer
    pos = UBound(ScenarioUpdateList)
    ReDim Preserve ScenarioUpdateList(pos + 1) As IBaseScenario
    Set ScenarioUpdateList(pos) = scenario
    AddUpdateScenario = pos
End Function

Public Sub UpdateAll()
    Dim o As Variant
    Dim key As String
    Dim map As IBaseScenario
    
    For Each o In CustomScenarioList.Keys
        Set map = CustomScenarioList(o)
        Call map.Update
    Next
End Sub

Public Sub Reset(ByVal MapNumber As Integer)
    Dim scenario As IBaseScenario
    Set scenario = GetMap(MapNumber)
    If scenario Is Nothing Then
        Exit Sub
    End If
    Call scenario.Reset
End Sub

Public Sub PlayerKillNpc(ByVal MapNumber As Integer, ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    Dim scenario As IBaseScenario
    Set scenario = GetMap(MapNumber)
    If scenario Is Nothing Then
        Exit Sub
    End If
    Call scenario.PlayerKillNpc(npcIndex, userIndex, sourceType, sourceIndex)
End Sub

Public Sub NPcKillPlayer(ByVal MapNumber As Integer, ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    Dim scenario As IBaseScenario
    Set scenario = GetMap(MapNumber)
    If scenario Is Nothing Then
        Exit Sub
    End If
    Call scenario.NPcKillPlayer(npcIndex, userIndex, sourceType, sourceIndex)
End Sub

Public Sub PlayerKillPlayer(ByVal MapNumber As Integer, ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    Dim scenario As IBaseScenario
    Set scenario = GetMap(MapNumber)
    If scenario Is Nothing Then
        Exit Sub
    End If
    Call scenario.PlayerKillPlayer(killerIndex, deadIndex, sourceType, sourceIndex)
End Sub

Public Sub PrepareNewEvent(ByVal eventType As e_EventType)
    
End Sub
