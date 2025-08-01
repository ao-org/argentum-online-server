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
End Sub

Public Function GetScenarioName() As String
End Function
Public Sub BroadcastOpenScenario()
End Sub

Public Sub SendRules(ByVal UserIndex As Integer)
End Sub

Public Function GetExtraInterfaces() As Long
End Function

Public Sub Reset()
    'empty interface
End Sub

Public Function RequiresSpawn() As Boolean
    RequiresSpawn = False
End Function

Public Function ValidateUser(ByVal userIndex As Integer) As Integer
    'empty interface
End Function

Public Sub WillSummonPlayer(ByVal UserIndex As Integer)
    'empty interface
End Sub

Public Sub UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef newState As e_LobbyState)
'empty interface
End Sub

Public Sub PlayerKillNpc(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    'empty interface
End Sub

Public Sub NPcKillPlayer(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    'empty interface
End Sub

Public Sub PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    'empty interface
End Sub

Public Sub NpcDie(ByVal NpcIndex)

End Sub

Public Sub UserDie(ByVal UserIndex)

End Sub

Public Sub Update()
    'empty interface
End Sub

Public Sub OnUserDisconnected(ByVal userIndex As Integer)
    
End Sub

Public Sub OnUserReconnect(ByVal userIndex As Integer)

End Sub

Public Sub GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)

End Sub

