Attribute VB_Name = "CustomScenarios"
Option Explicit

Public Enum e_DamageSourceType
    e_phisical
    e_magic
    e_pet
    e_dot
    e_trap
End Enum

Public Type e_Rank
    PlayerIndex As Integer
    Score As Integer
End Type

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
    Debug.Assert Index < UBound(ScenarioUpdateList)
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
        Case e_EventType.DeathMatch
            Set GenericGlobalLobby.scenario = New ScenarioDeathMatch
    End Select
    Exit Sub
PrepareNewEvent_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PrepareNewEvent", Erl)
End Sub

Public Sub ClearMap(ByVal mapNumber As Integer)
On Error GoTo ClearMap_Err:
    Dim x As Long
    Dim y As Long
        
106 For y = 1 To 99
108     For x = 1 To 99
110         If MapData(mapNumber, x, y).ObjInfo.objIndex > 0 Then
114             If ItemNoEsDeMapa(MapData(mapNumber, x, y).ObjInfo.objIndex) Then
116                 Call EraseObj(MAX_INVENTORY_OBJS, mapNumber, x, y)
                End If
            End If
118     Next x
120 Next y
    Exit Sub
ClearMap_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.ClearMap", Erl)
    Resume Next
End Sub

Public Function IsEventActive() As Boolean
    If CurrentActiveEventType = CaptureTheFlag Then
        IsEventActive = Not InstanciaCaptura Is Nothing
    Else
        IsEventActive = GenericGlobalLobby.State > e_LobbyState.UnInitilized And GenericGlobalLobby.State < Completed
    End If
End Function

Public Sub UserDisconnected(ByVal mapNumber As Integer, ByVal userIndex As Integer)
    Call RegisterDisconnectedUser(GenericGlobalLobby, userIndex)
End Sub

Public Sub UserConnected(ByVal userIndex)
    Call RegisterReconnectedUser(GenericGlobalLobby, userIndex)
End Sub
