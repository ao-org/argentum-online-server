Attribute VB_Name = "WorldTime"
' Argentum 20 Game Server
'
'    Copyright (C) 2026 Noland Studios LTD
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
' ============================
'  WorldTime.bas  (VB6)
'  Wrap-safe world time helpers (uses modTicksMasked)
'  - Uses GetTickCountRaw() + TicksElapsed() [mod 2^32]
'  - Always returns non-negative [0 .. DayLen-1]
' ============================
Option Explicit
Public Const DEFAULT_NPC_DAY_START_HOUR As Long = 6
Public Const DEFAULT_NPC_NIGHT_START_HOUR As Long = 20
' ---- Module state ----
Private WT_DayLenMs As Long     ' ms per in-game day (>=1)
Private WT_BaseTick As Long     ' synthetic base tick (raw tick) so elapsed = TicksElapsed(Base, NowRaw)
Private WT_Inited   As Boolean

' ============================
' Public API
' ============================
' Init / reset.
' Server: pass startElapsedMs:=0 at day start (or a saved offset if resuming).
' Client: usually not needed; clients call WorldTime_HandleHora from the packet.
Public Sub WorldTime_Init(ByVal dayLenMs As Long, Optional ByVal startElapsedMs As Long = 0)
    If dayLenMs <= 0 Then dayLenMs = 1
    WT_DayLenMs = dayLenMs
    WT_BaseTick = AddMod32(WorldTime_NowRaw(), -(startElapsedMs And &H7FFFFFFF)) ' store raw base; differences via TicksElapsed()
    WT_Inited = True
End Sub

' Client/Server: apply the server's HORA payload.
' elapsedFromServerMs = ms into the current day (server view).
' dayLenMs           = ms per in-game day.
Public Sub WorldTime_HandleHora(ByVal elapsedFromServerMs As Long, ByVal dayLenMs As Long)
    If dayLenMs <= 0 Then dayLenMs = 1
    WT_DayLenMs = dayLenMs
    Dim elapsedNorm As Long
    elapsedNorm = PosMod(CDbl(elapsedFromServerMs), WT_DayLenMs)
    WT_BaseTick = AddMod32(WorldTime_NowRaw(), -elapsedNorm)   ' base is raw; only compare via TicksElapsed()
    WT_Inited = True
End Sub

' Current in-game time in milliseconds within the day [0 .. DayLen-1]
Public Function WorldTime_Ms() As Long
    If Not WT_Inited Then WorldTime_Init 1, 0
    WorldTime_Ms = PosMod(TicksElapsed(WT_BaseTick, WorldTime_NowRaw()), WT_DayLenMs)
End Function

' Convenience: seconds within the day
Public Function WorldTime_Sec() As Long
    WorldTime_Sec = WorldTime_Ms() \ 1000
End Function

' Current in-game minute within the 24-hour fantasy day [0 .. 1439].
Public Function WorldTime_MinuteOfDay() As Long
    Dim dayLen As Long
    dayLen = WorldTime_DayLenMs()
    If dayLen <= 0 Then dayLen = 1
    WorldTime_MinuteOfDay = CLng(Fix((WorldTime_Ms() * 1440#) / dayLen))
End Function

' Evaluates a fantasy minute against configurable day/night boundaries.
' Day starts inclusively at dayStartHour; night starts inclusively at nightStartHour.
Public Function WorldTime_IsDaytimeAtMinute(ByVal minuteOfDay As Long, _
                                             ByVal dayStartHour As Long, _
                                             ByVal nightStartHour As Long) As Boolean
    If dayStartHour < 0 Or dayStartHour > 23 Or _
       nightStartHour < 0 Or nightStartHour > 23 Or _
       dayStartHour = nightStartHour Then
        dayStartHour = DEFAULT_NPC_DAY_START_HOUR
        nightStartHour = DEFAULT_NPC_NIGHT_START_HOUR
    End If

    minuteOfDay = PosMod(CDbl(minuteOfDay), 1440)

    Dim dayStartMinute As Long
    Dim nightStartMinute As Long
    dayStartMinute = dayStartHour * 60
    nightStartMinute = nightStartHour * 60

    If dayStartMinute < nightStartMinute Then
        WorldTime_IsDaytimeAtMinute = (minuteOfDay >= dayStartMinute And minuteOfDay < nightStartMinute)
    Else
        WorldTime_IsDaytimeAtMinute = (minuteOfDay >= dayStartMinute Or minuteOfDay < nightStartMinute)
    End If
End Function

Public Function WorldTime_IsDaytime(ByVal dayStartHour As Long, ByVal nightStartHour As Long) As Boolean
    WorldTime_IsDaytime = WorldTime_IsDaytimeAtMinute(WorldTime_MinuteOfDay(), dayStartHour, nightStartHour)
End Function

' Get/Set day length
Public Function WorldTime_DayLenMs() As Long
    If WT_DayLenMs <= 0 Then WT_DayLenMs = 1
    WorldTime_DayLenMs = WT_DayLenMs
End Function

Public Sub WorldTime_SetDayLenMs(ByVal dayLenMs As Long)
    If dayLenMs <= 0 Then dayLenMs = 1
    WT_DayLenMs = dayLenMs
End Sub

' Optional: quick re-anchor using a fresh server elapsed (no dayLen change)
Public Sub WorldTime_Resync(ByVal serverElapsedMs As Long)
    If WT_DayLenMs <= 0 Then WT_DayLenMs = 1
    Dim elapsedNorm As Long
    elapsedNorm = PosMod(CDbl(serverElapsedMs), WT_DayLenMs)
    WT_BaseTick = AddMod32(WorldTime_NowRaw(), -elapsedNorm)
    WT_Inited = True
End Sub

' SERVER helper: compute payload (elapsed, dayLen) for the HORA packet.
' Returns normalized elapsed in [0 .. DayLen-1] based on current WT_BaseTick.
Public Sub WorldTime_PrepareHora(ByRef outElapsedMs As Long, ByRef outDayLenMs As Long)
    If Not WT_Inited Then WorldTime_Init 1, 0
    outDayLenMs = WT_DayLenMs
    outElapsedMs = PosMod(TicksElapsed(WT_BaseTick, WorldTime_NowRaw()), WT_DayLenMs)
End Sub

' ============================
' Internal
' ============================
' Always use RAW ticks for this module (pairs with TicksElapsed 2^32)
Private Function WorldTime_NowRaw() As Long
    WorldTime_NowRaw = GetTickCountRaw()
End Function
