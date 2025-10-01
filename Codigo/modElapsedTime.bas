Attribute VB_Name = "modElapsedTime"
' Argentum 20 Game Server
'
'    Copyright (C) 2025 Noland Studios LTD
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
Option Explicit


' ============================================
' Why NOT to use masked GetTickCount() for timing
' ============================================
'
' Windows timeGetTime() returns a 32-bit unsigned millisecond counter:
'   Range: 0 .. 4,294,967,295 (2^32 - 1)
'   Wrap period: ~49.7 days
'
' Our legacy wrapper applied a mask:
'   GetTickCount = timeGetTime() And &H7FFFFFFF
'
' That forces the high (sign) bit to 0, so the value is always >= 0.
' But it also cuts the usable range in half:
'   Range: 0 .. 2,147,483,647 (2^31 - 1)
'   Wrap period: ~24.9 days
'
' The problem: code that does naive subtraction across a wrap breaks.
'
' Example:
'   start = 2,147,482,647   ' (2^31 - 1000), just before wrap
'   now   = 500             ' just after wrap
'
'   naive = now - start     ' 500 - 2,147,482,647 = -2,147,482,147
'   If naive > delay Then   ' always False (negative), even though 500 ms passed
'
' This means timers, respawns, cooldowns, etc. can "stall" for ~24.9 days
' after a wrap if they rely on (now - start > delay).
'
' Correct approach:
'   - Use the full unmasked tick: GetTickCountRaw() (0..2^32-1)
'   - Compute elapsed with wrap-safe math: TicksElapsed(start, now)
'   - Or compare deadlines with unsigned logic: TickAfter(now, deadline)
'
' This ensures timers keep working correctly across the natural wrap of the
' Windows tick counter.
'
'
'
' ==============================================================
' Migration Guide: replacing masked GetTickCount() code
' ==============================================================
'
' Problem:
'   Legacy code often uses masked ticks with naive subtraction:
'
'       start = GetTickCount()
'       If GetTickCount() - start > delay Then ...
'
'   This breaks across the 2^31 wrap (~24.9 days). The subtraction
'   goes negative and the timer stalls.
'
' Correct usage with wrap-safe helpers:
'
' 1) Elapsed style
'    -------------------------
'       start = GetTickCountRaw()
'       If TicksElapsed(start, GetTickCountRaw()) > delay Then
'           ' timer fired
'       End If
'
'    - Always non-negative
'    - Best when you only need "has delay passed?"
'
' 2) Deadline style (recommended for repeating timers)
'    -------------------------
'       start    = GetTickCountRaw()
'       deadline = AddMod32(start, delay)
'       If TickAfter(GetTickCountRaw(), deadline) Then
'           ' timer fired
'       End If
'
'       ' For periodic tasks:
'       If TickAfter(GetTickCountRaw(), nextRun) Then
'           nextRun = AddMod32(nextRun, interval)
'           RunTask
'       End If
'
'    - Stores the absolute time when the timer should fire
'    - Safer for repeating events (no drift)
'
' 3) Examples of legacy code that must be migrated
'    -------------------------
'       If GetTickCount() - RespawnTime > PlayerRespawnTime Then ...
'       If GetTickCount() > cooldownDeadline Then ...
'       Do While GetTickCount() < wakeAt: DoEvents: Loop
'
'    Replace with:
'       If TicksElapsed(RespawnTime, GetTickCountRaw()) > PlayerRespawnTime Then ...
'       If TickAfter(GetTickCountRaw(), cooldownDeadline) Then ...
'       Do While Not TickAfter(GetTickCountRaw(), wakeAt): DoEvents: Loop
'
' Helpers provided by this module:
'   - GetTickCountRaw()  : unmasked 2^32 tick counter
'   - TicksElapsed()     : wrap-safe elapsed (ms)
'   - TickAfter()        : wrap-safe "is A after B?" compare
'   - AddMod32()         : add with wrap (for deadlines/intervals)
'   - PosMod()           : normalize to positive range
'
' ==============================================================
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Const TICKS32 As Double = 4294967296#

' Legacy (keep for now, used by old code paths)
' --------------------------------------------
' NOTE: Do not rely on masked GetTickCount().
' Masking (&H7FFFFFFF) halves the range (2^31),
' wraps every ~24.9 days, and breaks naive math:
'   (now - start > delay) fails across wrap.
'
' Use GetTickCountRaw() + TicksElapsed() /
' TickAfter() for wrap-safe timing instead.
' --------------------------------------------
Public Function GetTickCount() As Long
    GetTickCount = timeGetTime() And &H7FFFFFFF
End Function

' New raw version (preferred)
Public Function GetTickCountRaw() As Long
    GetTickCountRaw = timeGetTime()
    'to simulare neg stamps do: Or &H80000000
End Function

Public Function TicksElapsed(ByVal startTick As Long, ByVal currentTick As Long) As Double
    If currentTick >= startTick Then
        TicksElapsed = CDbl(currentTick - startTick)
    Else
        TicksElapsed = (TICKS32 - CDbl(startTick)) + CDbl(currentTick)
    End If
End Function

Public Function TickAfter(ByVal a As Long, ByVal b As Long) As Boolean
    TickAfter = (a - b) >= 0
End Function

Public Function PosMod(ByVal a As Double, ByVal m As Long) As Long
    If m <= 0 Then PosMod = 0: Exit Function
    Dim r As Double
    r = a - m * Fix(a / m)
    If r >= m Then r = r - m
    If r < 0 Then r = r + m
    PosMod = CLng(r)
End Function

' Add two tick values modulo 2^32 (wrap-safe)
Public Function AddMod32(ByVal a As Long, ByVal b As Long) As Long
    Dim s As Double
    s = CDbl(a And &HFFFFFFFF) + CDbl(b And &HFFFFFFFF)
    ' reduce modulo 2^32
    s = s - TICKS32 * Fix(s / TICKS32)
    AddMod32 = CLng(s)
End Function

' ==============================================================
' DeadlinePassed
' --------------------------------------------------------------
' Wrap-safe check for whether "now" has passed a given deadline.
'
' Why not just use TickAfter(now, deadline)?
' ------------------------------------------
' The standard trick:
'     TickAfter = (now - deadline) >= 0
' works correctly on a modulo-2^32 tick ring as long as the
' unsigned distance between "now" and "deadline" is < 2^31.
'
' But in our stun/cooldown code, we sometimes store StunEndTime=0
' to mean "no stun / unset".
'
' Problem:
'   If deadline=0 and now has the high bit set
'   (e.g. nowRaw = &H86F0E019 = -2031327897 signed),
'   then (now - 0) is negative in signed math.
'   TickAfter() returns False ? interpreted as "deadline not passed"
'   ? player/NPC is treated as still stunned.
'
' In reality, a deadline of 0 should always mean "already expired".
'
' Fix:
'   Special-case deadline=0 as always passed.
'   Otherwise fall back to the wrap-safe TickAfter compare.
'
' Usage:
'   If DeadlinePassed(GetTickCountRaw(), counters.StunEndTime) Then ...
'
' ==============================================================
Public Function DeadlinePassed(ByVal nowRaw As Long, ByVal deadline As Long) As Boolean
    If deadline = 0 Then
        DeadlinePassed = True        ' treat 0 as "no deadline"
    Else
        DeadlinePassed = (nowRaw - deadline) >= 0   ' wrap-safe TickAfter
    End If
End Function



