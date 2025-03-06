Attribute VB_Name = "ModCuentas"
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

Public Function GetUserGuildIndex(ByVal username As String) As Integer

        On Error GoTo GetUserGuildIndex_Err

100     If InStrB(username, "\") <> 0 Then
102         username = Replace(username, "\", vbNullString)

        End If

104     If InStrB(username, "/") <> 0 Then
106         username = Replace(username, "/", vbNullString)

        End If

108     If InStrB(username, ".") <> 0 Then
110         username = Replace(username, ".", vbNullString)

        End If

116     GetUserGuildIndex = GetUserGuildIndexDatabase(username)
        Exit Function
GetUserGuildIndex_Err:
118     Call TraceError(Err.Number, Err.Description, "ModCuentas.GetUserGuildIndex", Erl)

End Function

Public Function ObtenerCriminal(ByVal name As String) As Byte

        On Error GoTo ErrorHandler

        Dim Criminal As Byte
102     Criminal = GetUserStatusDatabase(name)

106     If EsRolesMaster(name) Then
108         Criminal = 3
110     ElseIf EsConsejero(name) Then
112         Criminal = 4
114     ElseIf EsSemiDios(name) Then
116         Criminal = 5
118     ElseIf EsDios(name) Then
120         Criminal = 6
122     ElseIf EsAdmin(name) Then
124         Criminal = 7

        End If

126     ObtenerCriminal = Criminal
        Exit Function
ErrorHandler:
128     ObtenerCriminal = 1

End Function
