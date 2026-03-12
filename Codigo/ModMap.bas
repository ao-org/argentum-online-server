Attribute VB_Name = "ModMap"
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

Public Function CanAddTrapAt(ByVal mapIndex As Integer, ByVal PosX As Integer, ByVal PosY As Integer) As Boolean
    If Not MapData(mapIndex, PosX, PosY).Trap Is Nothing Then
        Exit Function
    End If
    If MapData(mapIndex, PosX, PosY).Blocked Then Exit Function
    If MapData(mapIndex, PosX, PosY).NpcIndex > 0 Then Exit Function
    If MapData(mapIndex, PosX, PosY).UserIndex > 0 Then Exit Function
    If MapData(mapIndex, PosX, PosY).ObjInfo.ObjIndex > 0 Then Exit Function
    CanAddTrapAt = True
End Function

Public Sub ActivateTrap(ByVal TargetIndex, ByVal TargetType As e_ReferenceType, ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
    If MapData(Map, PosX, PosY).Trap Is Nothing Then
        Exit Sub
    End If
    If Not MapData(Map, PosX, PosY).Trap.CanAffectTarget(TargetIndex, TargetType) Then
        Exit Sub
    End If
    Call MapData(Map, PosX, PosY).Trap.trigger(TargetIndex, TargetType)
End Sub
