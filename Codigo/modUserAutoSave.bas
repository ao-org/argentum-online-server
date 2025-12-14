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

Private Const AutoSaveLoopTimeLimitMs As Long = 200

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

    Dim UserIndex                         As Integer
    Dim UserGuardados                     As Integer
    Dim PerformanceTimer                  As Long
    Dim elapsedSinceLastSave              As Long
    Dim totalElapsedTime                  As Long
    Dim unsavedUsersDueToShortInterval    As Integer
    Dim unsavedUsersDueToTimeLimit        As Integer
    Dim totalLoggedUsers                  As Integer
    Dim processedLoggedUsers              As Integer

    For UserIndex = 1 To LastUser
        If UserList(UserIndex).flags.UserLogged Then
            totalLoggedUsers = totalLoggedUsers + 1
        End If
    Next

    Call PerformanceTestStart(PerformanceTimer)

    For UserIndex = 1 To LastUser
        With UserList(UserIndex)
            If .flags.UserLogged Then
                processedLoggedUsers = processedLoggedUsers + 1
                nowRaw = GetTickCountRaw()
                elapsedSinceLastSave = TicksElapsed(.Counters.LastSave, nowRaw)

                If elapsedSinceLastSave > IntervaloGuardarUsuarios Then
                    Call SaveChangesInUser(UserIndex)
                    .Counters.LastSave = nowRaw
                    UserGuardados = UserGuardados + 1
                    If UserGuardados > NumUsers Then Exit For
                    If TicksElapsed(PerformanceTimer, GetTickCountRaw()) > AutoSaveLoopTimeLimitMs Then
                        Exit For
                    End If
                Else
                    unsavedUsersDueToShortInterval = unsavedUsersDueToShortInterval + 1
                End If
            End If
        End With
    Next

    totalElapsedTime = CLng(TicksElapsed(PerformanceTimer, GetTickCountRaw()))
    unsavedUsersDueToTimeLimit = totalLoggedUsers - processedLoggedUsers
    If unsavedUsersDueToTimeLimit < 0 Then unsavedUsersDueToTimeLimit = 0

    Call LogPerformance("Auto-save summary - total users: " & LastUser & _
                        " | logged: " & totalLoggedUsers & _
                        " | saved: " & UserGuardados & _
                        " | not saved (interval not elapsed): " & unsavedUsersDueToShortInterval & _
                        " | not saved (time limit reached): " & unsavedUsersDueToTimeLimit & _
                        " | elapsed ms: " & totalElapsedTime)

    Call PerformTimeLimitCheck(PerformanceTimer, "modUserAutoSave.MaybeRunUserAutoSave", AutoSaveLoopTimeLimitMs)
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "modUserAutoSave.MaybeRunUserAutoSave")
End Sub
