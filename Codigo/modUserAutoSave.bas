Attribute VB_Name = "modUserAutoSave"
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
'    Copyright (C) 2002 Mrquez Pablo Ignacio
Option Explicit

Private m_LastAutoSaveAttempt As Long

Public Sub ResetUserAutoSaveTimer()
    m_LastAutoSaveAttempt = GetTickCountRaw()
End Sub

Public Sub MaybeRunUserAutoSave()
    On Error GoTo Handler

    If Not IsFeatureEnabled("auto_save_chars") Then Exit Sub

    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()

    If m_LastAutoSaveAttempt = 0 Then
        m_LastAutoSaveAttempt = nowRaw
        Exit Sub
    End If

    If TicksElapsed(m_LastAutoSaveAttempt, nowRaw) < IntervaloTimerGuardarUsuarios Then Exit Sub

    m_LastAutoSaveAttempt = nowRaw

    Dim UserIndex        As Integer
    Dim UserGuardados    As Integer
    Dim PerformanceTimer As Long

    Call PerformanceTestStart(PerformanceTimer)

    For UserIndex = 1 To LastUser
        With UserList(UserIndex)
            If .flags.UserLogged Then
                nowRaw = GetTickCountRaw()
                If TicksElapsed(.Counters.LastSave, nowRaw) > IntervaloGuardarUsuarios Then
                    Call SaveUser(UserIndex)
                    UserGuardados = UserGuardados + 1
                    If UserGuardados > NumUsers Then Exit For
                    If TicksElapsed(PerformanceTimer, GetTickCountRaw()) > 100 Then Exit For
                End If
            End If
        End With
    Next

    Call PerformTimeLimitCheck(PerformanceTimer, "modUserAutoSave.MaybeRunUserAutoSave", 100)
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "modUserAutoSave.MaybeRunUserAutoSave")
End Sub
