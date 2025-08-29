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

Public Function CanAddTrapAt(ByVal mapIndex As Integer, ByVal posX As Integer, ByVal posY As Integer) As Boolean
    On Error Goto CanAddTrapAt_Err
    If Not MapData(mapIndex, posX, posY).Trap Is Nothing Then
        Exit Function
    End If
    If MapData(mapIndex, posX, posY).Blocked Then Exit Function
    If MapData(mapIndex, posX, posY).npcIndex > 0 Then Exit Function
    If MapData(mapIndex, posX, posY).UserIndex > 0 Then Exit Function
    If MapData(mapIndex, posX, posY).ObjInfo.objIndex > 0 Then Exit Function
    CanAddTrapAt = True
    Exit Function
CanAddTrapAt_Err:
    Call TraceError(Err.Number, Err.Description, "ModMap.CanAddTrapAt", Erl)
End Function

Public Sub ActivateTrap(ByVal TargetIndex, ByVal TargetType As e_ReferenceType, ByVal map As Integer, ByVal posX As Integer, ByVal posY As Integer)
    On Error Goto ActivateTrap_Err
    If MapData(map, posX, posY).Trap Is Nothing Then
     Exit Sub
    End If
    If Not MapData(map, posX, posY).Trap.CanAffectTarget(TargetIndex, TargetType) Then
        Exit Sub
    End If
    Call MapData(map, posX, posY).Trap.trigger(TargetIndex, TargetType)
    Exit Sub
ActivateTrap_Err:
    Call TraceError(Err.Number, Err.Description, "ModMap.ActivateTrap", Erl)
End Sub
