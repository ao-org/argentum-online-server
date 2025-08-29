VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBaseScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Sub SetLobbyIndex(ByVal Value As Integer)
    On Error Goto SetLobbyIndex_Err
    Exit Sub
SetLobbyIndex_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.SetLobbyIndex", Erl)
End Sub

Public Function GetScenarioName() As String
    On Error Goto GetScenarioName_Err
    Exit Function
GetScenarioName_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.GetScenarioName", Erl)
End Function
Public Sub BroadcastOpenScenario()
    On Error Goto BroadcastOpenScenario_Err
    Exit Sub
BroadcastOpenScenario_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.BroadcastOpenScenario", Erl)
End Sub

Public Sub SendRules(ByVal UserIndex As Integer)
    On Error Goto SendRules_Err
    Exit Sub
SendRules_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.SendRules", Erl)
End Sub

Public Function GetExtraInterfaces() As Long
    On Error Goto GetExtraInterfaces_Err
    Exit Function
GetExtraInterfaces_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.GetExtraInterfaces", Erl)
End Function

Public Sub Reset()
    On Error Goto Reset_Err
    'empty interface
    Exit Sub
Reset_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.Reset", Erl)
End Sub

Public Function RequiresSpawn() As Boolean
    On Error Goto RequiresSpawn_Err
    RequiresSpawn = False
    Exit Function
RequiresSpawn_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.RequiresSpawn", Erl)
End Function

Public Function ValidateUser(ByVal userIndex As Integer) As Integer
    On Error Goto ValidateUser_Err
    'empty interface
    Exit Function
ValidateUser_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.ValidateUser", Erl)
End Function

Public Sub WillSummonPlayer(ByVal UserIndex As Integer)
    On Error Goto WillSummonPlayer_Err
    'empty interface
    Exit Sub
WillSummonPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.WillSummonPlayer", Erl)
End Sub

Public Sub UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef newState As e_LobbyState)
    On Error Goto UpdateLobbyState_Err
'empty interface
    Exit Sub
UpdateLobbyState_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.UpdateLobbyState", Erl)
End Sub

Public Sub PlayerKillNpc(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    On Error Goto PlayerKillNpc_Err
    'empty interface
    Exit Sub
PlayerKillNpc_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.PlayerKillNpc", Erl)
End Sub

Public Sub NPcKillPlayer(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    On Error Goto NPcKillPlayer_Err
    'empty interface
    Exit Sub
NPcKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.NPcKillPlayer", Erl)
End Sub

Public Sub PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    On Error Goto PlayerKillPlayer_Err
    'empty interface
    Exit Sub
PlayerKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.PlayerKillPlayer", Erl)
End Sub

Public Sub NpcDie(ByVal NpcIndex)
    On Error Goto NpcDie_Err

    Exit Sub
NpcDie_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.NpcDie", Erl)
End Sub

Public Sub UserDie(ByVal UserIndex)
    On Error Goto UserDie_Err

    Exit Sub
UserDie_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.UserDie", Erl)
End Sub

Public Sub Update()
    On Error Goto Update_Err
    'empty interface
    Exit Sub
Update_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.Update", Erl)
End Sub

Public Sub OnUserDisconnected(ByVal userIndex As Integer)
    On Error Goto OnUserDisconnected_Err
    
    Exit Sub
OnUserDisconnected_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.OnUserDisconnected", Erl)
End Sub

Public Sub OnUserReconnect(ByVal userIndex As Integer)
    On Error Goto OnUserReconnect_Err

    Exit Sub
OnUserReconnect_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.OnUserReconnect", Erl)
End Sub

Public Sub GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
    On Error Goto GetNextWaypointForNpc_Err

    Exit Sub
GetNextWaypointForNpc_Err:
    Call TraceError(Err.Number, Err.Description, "IBaseScenario.GetNextWaypointForNpc", Erl)
End Sub

