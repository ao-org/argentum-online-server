Attribute VB_Name = "CustomScenarios"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Public Enum e_DamageSourceType
    e_phisical
    e_magic
    e_pet
    e_dot
    e_trap
End Enum

Public Enum e_ScenarioInterfaces
    eInventoryInterface = 1
End Enum

Public Type e_Rank
    PlayerIndex As Integer
    Score As Integer
End Type

Private CustomScenarioList   As New Dictionary
Private ScenarioUpdateList() As IBaseScenario
Private AvailableUpdateSlots As t_IndexHeap
Private ActiveUpdateSlots    As t_IndexHeap
Const InitialUpdateSize = 20

Private Sub InitializeUpdateStacks()
    ReDim ScenarioUpdateList(InitialUpdateSize) As IBaseScenario
    ReDim AvailableUpdateSlots.IndexInfo(InitialUpdateSize)
    ReDim ActiveUpdateSlots.IndexInfo(InitialUpdateSize)
    Dim i As Integer
    For i = 1 To InitialUpdateSize
        AvailableUpdateSlots.IndexInfo(i) = InitialUpdateSize - (i - 1)
    Next i
    AvailableUpdateSlots.currentIndex = InitialUpdateSize
    ActiveUpdateSlots.currentIndex = 0
End Sub

Private Sub IncreaseArraySize(ByVal ExtraSlots As Integer)
    Dim NewSize As Integer
    NewSize = UBound(ScenarioUpdateList) + ExtraSlots
    ReDim Preserve ScenarioUpdateList(NewSize) As IBaseScenario
    ReDim Preserve AvailableUpdateSlots.IndexInfo(NewSize)
    ReDim Preserve ActiveUpdateSlots.IndexInfo(NewSize)
    Dim i As Integer
    For i = 1 To ExtraSlots
        AvailableUpdateSlots.IndexInfo(i) = NewSize - (i - 1)
    Next i
    AvailableUpdateSlots.currentIndex = ExtraSlots
End Sub

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

Public Sub SetCustomScenario(ByVal mapIndex As Integer, ByRef Scenario As IBaseScenario)
    Call CustomScenarioList.Add(mapIndex, Scenario)
End Sub

Public Sub RemoveCustomScenario(ByVal mapIndex As Integer)
    Call CustomScenarioList.Remove(mapIndex)
End Sub

Public Function AddUpdateScenario(ByRef Scenario As IBaseScenario) As Integer
    On Error GoTo AddUpdateScenario_Err:
    Dim pos As Integer
    If AvailableUpdateSlots.currentIndex = 0 And ActiveUpdateSlots.currentIndex = 0 Then
        Call InitializeUpdateStacks
    End If
    If AvailableUpdateSlots.currentIndex = 0 Then
        Call IncreaseArraySize(InitialUpdateSize)
    End If
    pos = AvailableUpdateSlots.IndexInfo(AvailableUpdateSlots.currentIndex)
    AvailableUpdateSlots.currentIndex = AvailableUpdateSlots.currentIndex - 1
    ActiveUpdateSlots.IndexInfo(ActiveUpdateSlots.currentIndex) = pos
    ActiveUpdateSlots.currentIndex = ActiveUpdateSlots.currentIndex + 1
    Set ScenarioUpdateList(pos) = Scenario
    AddUpdateScenario = pos
    Exit Function
AddUpdateScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.AddUpdateScenario", Erl)
End Function

Public Sub RemoveUpdateScenario(ByRef Index As Integer)
    Debug.Assert Index < UBound(ScenarioUpdateList)
    Set ScenarioUpdateList(Index) = Nothing
    Dim i As Integer
    For i = 0 To ActiveUpdateSlots.currentIndex - 1
        If ActiveUpdateSlots.IndexInfo(i) = Index Then
            ActiveUpdateSlots.IndexInfo(i) = ActiveUpdateSlots.IndexInfo(ActiveUpdateSlots.currentIndex - 1)
            ActiveUpdateSlots.currentIndex = ActiveUpdateSlots.currentIndex - 1
            AvailableUpdateSlots.currentIndex = AvailableUpdateSlots.currentIndex + 1
            AvailableUpdateSlots.IndexInfo(AvailableUpdateSlots.currentIndex) = Index
        End If
    Next
End Sub

Public Sub UpdateAll()
    On Error GoTo UpdateAll_Err:
    Dim i As Integer
    For i = 0 To ActiveUpdateSlots.currentIndex - 1
        If Not ScenarioUpdateList(ActiveUpdateSlots.IndexInfo(i)) Is Nothing Then
            Call ScenarioUpdateList(ActiveUpdateSlots.IndexInfo(i)).Update
        End If
    Next
    Exit Sub
UpdateAll_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UpdateAll", Erl)
End Sub

Public Sub Reset(ByVal MapNumber As Integer)
    On Error GoTo Reset_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.Reset
    Exit Sub
Reset_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.Reset " & MapNumber, Erl)
End Sub

Public Sub PlayerKillNpc(ByVal MapNumber As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    On Error GoTo PlayerKillNpc_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.PlayerKillNpc(NpcIndex, UserIndex, SourceType, SourceIndex)
    Exit Sub
PlayerKillNpc_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PlayerKillNpc map:" & MapNumber & " npc: " & NpcIndex & " user: " & UserIndex, Erl)
End Sub

Public Sub NPcKillPlayer(ByVal MapNumber As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    On Error GoTo NPcKillPlayer_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.NPcKillPlayer(NpcIndex, UserIndex, SourceType, SourceIndex)
    Exit Sub
NPcKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.NPcKillPlayer", Erl)
End Sub

Public Sub PlayerKillPlayer(ByVal MapNumber As Integer, _
                            ByVal killerIndex As Integer, _
                            ByVal deadIndex As Integer, _
                            ByVal SourceType As e_DamageSourceType, _
                            ByVal SourceIndex As Integer)
    On Error GoTo PlayerKillPlayer_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.PlayerKillPlayer(killerIndex, deadIndex, SourceType, SourceIndex)
    Exit Sub
PlayerKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PlayerKillPlayer", Erl)
End Sub

Public Sub UserDie(ByVal UserIndex As Integer)
    On Error GoTo UserDie_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(UserList(UserIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.UserDie(UserIndex)
    Exit Sub
UserDie_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserDie", Erl)
End Sub

Public Sub NpcDie(ByVal NpcIndex As Integer)
    On Error GoTo NpcDie_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(NpcList(NpcIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.NpcDie(NpcIndex)
    Exit Sub
NpcDie_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.NpcDie", Erl)
End Sub

'inform that user want to pickup an item, return false if user can't pick up item at pos
Public Function UserCanPickUpItem(ByVal UserIndex As Integer) As Boolean
    On Error GoTo UserPickUpItem_Err:
    UserCanPickUpItem = True
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(UserList(UserIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Function
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Function
    End If
    Set InventoryInterface = Scenario
    UserCanPickUpItem = InventoryInterface.UserCanPickUpItem(UserIndex)
    Exit Function
UserPickUpItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserCanPickUpItem", Erl)
End Function

Public Sub UserDidPickupItem(ByVal UserIndex As Integer, ByVal ItemId As Integer)
    On Error GoTo UserDidPickupItem_Err:
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(UserList(UserIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Sub
    End If
    Set InventoryInterface = Scenario
    Call InventoryInterface.UserDidPickupItem(UserIndex, ItemId)
    Exit Sub
UserDidPickupItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserDidPickupItem", Erl)
End Sub

'inform that user want to drop an item
Public Sub UserDropItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    On Error GoTo UserDropItem_Err:
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Sub
    End If
    Set InventoryInterface = Scenario
    Call InventoryInterface.UserDropItem(UserIndex, Slot, Map, TileX, TileY)
    Exit Sub
UserDropItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserDropItem", Erl)
End Sub

Public Function UserCanDropItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer) As Boolean
    On Error GoTo UserCanDropItem_Err:
    UserCanDropItem = True
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(Map)
    If Scenario Is Nothing Then
        Exit Function
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Function
    End If
    Set InventoryInterface = Scenario
    UserCanDropItem = InventoryInterface.UserCanDropItemAt(UserIndex, Slot, Map, TileX, TileY)
    Exit Function
UserCanDropItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserCanDropItem", Erl)
End Function

Public Sub PrepareNewEvent(ByVal eventType As e_EventType, ByVal LobbyIndex As Integer)
    On Error GoTo PrepareNewEvent_Err:
    Debug.Assert LobbyIndex < UBound(LobbyList)
    Select Case eventType
        Case e_EventType.NpcHunt
            Set LobbyList(LobbyIndex).Scenario = New ScenarioHunt
        Case e_EventType.DeathMatch
            Set LobbyList(LobbyIndex).Scenario = New ScenarioDeathMatch
        Case e_EventType.NavalBattle
            Set LobbyList(LobbyIndex).Scenario = New ScenarioNavalBoarding
    End Select
    If Not LobbyList(LobbyIndex).Scenario Is Nothing Then
        LobbyList(LobbyIndex).Scenario.SetLobbyIndex (LobbyIndex)
    End If
    Exit Sub
PrepareNewEvent_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PrepareNewEvent", Erl)
End Sub

Public Sub ClearMap(ByVal MapNumber As Integer)
    On Error GoTo ClearMap_Err:
    Dim x As Long
    Dim y As Long
    For y = 1 To 99
        For x = 1 To 99
            If MapData(MapNumber, x, y).ObjInfo.ObjIndex > 0 Then
                If ItemNoEsDeMapa(MapData(MapNumber, x, y).ObjInfo.ObjIndex) Then
                    Call EraseObj(GetMaxInvOBJ(), MapNumber, x, y)
                End If
            End If
        Next x
    Next y
    Exit Sub
ClearMap_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.ClearMap", Erl)
    Resume Next
End Sub

Public Function IsEventActive() As Boolean
    If CurrentActiveEventType = CaptureTheFlag Then
        IsEventActive = Not InstanciaCaptura Is Nothing
    Else
        If GlobalLobbyIndex >= 0 Then
            IsEventActive = LobbyList(GlobalLobbyIndex).State > e_LobbyState.UnInitilized And LobbyList(GlobalLobbyIndex).State < InProgress
        Else
            IsEventActive = False
        End If
    End If
End Function

Public Sub UserDisconnected(ByVal MapNumber As Integer, ByVal UserIndex As Integer)
    Call RegisterDisconnectedUser(UserIndex)
End Sub

Public Sub UserConnected(ByVal UserIndex)
    Call RegisterReconnectedUser(UserIndex)
End Sub

Public Sub GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(NpcList(NpcIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.GetNextWaypointForNpc(NpcIndex, PosX, PosY)
    Exit Sub
End Sub
