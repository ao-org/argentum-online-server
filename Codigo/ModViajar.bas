Attribute VB_Name = "ModViajar"
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
'
Public Sub IniciarTransporte(ByVal UserIndex As Integer)
        On Error GoTo IniciarTransporte_Err
        Dim destinos As Byte
100     destinos = NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).NumDestinos
        Exit Sub
IniciarTransporte_Err:
102     Call TraceError(Err.Number, Err.Description, "ModViajar.IniciarTransporte", Erl)
End Sub
