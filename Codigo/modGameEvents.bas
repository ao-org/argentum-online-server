Attribute VB_Name = "modGameEvents"
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

Private Const GAME_EVENTS_INTERVAL_MS As Long = 40
Private Const USER_LOOP_TIME_LIMIT_MS As Long = 400
Private Const CUSTOM_SCENARIOS_TIME_LIMIT_MS As Long = 100

Private m_LastGameEventsTick As Long

Public Sub ResetGameEventsTimer()
    m_LastGameEventsTick = GetTickCountRaw()
End Sub

Public Sub MaybeRunGameEvents()
    On Error GoTo Handler

    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()

    If m_LastGameEventsTick = 0 Then
        m_LastGameEventsTick = nowRaw
        Exit Sub
    End If

    If TicksElapsed(m_LastGameEventsTick, nowRaw) < GAME_EVENTS_INTERVAL_MS Then Exit Sub

    m_LastGameEventsTick = nowRaw

    Dim iUserIndex       As Long
    Dim PerformanceTimer As Long

    Call PerformanceTestStart(PerformanceTimer)

    For iUserIndex = 1 To LastUser
        With UserList(iUserIndex)
            If .flags.UserLogged Then
                Call DoTileEvents(iUserIndex, .pos.Map, .pos.x, .pos.y)
                If .flags.Muerto = 0 Then
                    'Efectos en mapas
                    If (.flags.Privilegios And e_PlayerType.User) <> 0 Then
                        Call EfectoLava(iUserIndex)
                        Call EfectoFrio(iUserIndex)
                        If .flags.Envenenado <> 0 Then Call EfectoVeneno(iUserIndex)
                        If .flags.Incinerado <> 0 Then Call EfectoIncineramiento(iUserIndex)
                    End If
                    If .flags.Meditando Then Call DoMeditar(iUserIndex)
                    If .flags.Mimetizado <> 0 Then Call EfectoMimetismo(iUserIndex)
                    If .flags.AdminInvisible <> 1 Then
                        If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                    End If
                    If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                    Call EfectoStamina(iUserIndex)
                End If 'Muerto
            End If 'UserLogged
        End With
    Next iUserIndex

    Call PerformTimeLimitCheck(PerformanceTimer, "MaybeRunGameEvents User loop", USER_LOOP_TIME_LIMIT_MS)
    Call CustomScenarios.UpdateAll
    Call PerformTimeLimitCheck(PerformanceTimer, "MaybeRunGameEvents customScenarios", CUSTOM_SCENARIOS_TIME_LIMIT_MS)
    Call MaybeRunShipTravel

    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description & vbNewLine & "UserIndex:" & iUserIndex, "modGameEvents.MaybeRunGameEvents")
End Sub
