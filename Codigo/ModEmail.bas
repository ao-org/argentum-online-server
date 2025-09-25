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
    If InStrB(username, "\") <> 0 Then
        username = Replace(username, "\", vbNullString)
    End If
    If InStrB(username, "/") <> 0 Then
        username = Replace(username, "/", vbNullString)
    End If
    If InStrB(username, ".") <> 0 Then
        username = Replace(username, ".", vbNullString)
    End If
    GetUserGuildIndex = GetUserGuildIndexDatabase(username)
    Exit Function
GetUserGuildIndex_Err:
    Call TraceError(Err.Number, Err.Description, "ModCuentas.GetUserGuildIndex", Erl)
End Function

Public Function ObtenerCriminal(ByVal name As String) As Byte
    On Error GoTo ErrorHandler
    Dim Criminal As Byte
    Criminal = GetUserStatusDatabase(name)
    If EsRolesMaster(name) Then
        Criminal = 3
    ElseIf EsConsejero(name) Then
        Criminal = 4
    ElseIf EsSemiDios(name) Then
        Criminal = 5
    ElseIf EsDios(name) Then
        Criminal = 6
    ElseIf EsAdmin(name) Then
        Criminal = 7
    End If
    ObtenerCriminal = Criminal
    Exit Function
ErrorHandler:
    ObtenerCriminal = 1
End Function
