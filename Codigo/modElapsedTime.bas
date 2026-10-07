Attribute VB_Name = "modElapsedTime"
' Argentum 20 Game Server
'
'    Copyright (C) 2025-2026 Noland Studios LTD
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
' Windows timeGetTime() returns a modulo-2^32 millisecond counter.
' VB6 stores that bit pattern in a signed Long:
'   &H7FFFFFFF =  2147483647
'   &H80000000 = -2147483648
'   &HFFFFFFFF =          -1
'
' Crossing &H7FFFFFFF -> &H80000000 only changes the signed representation.
' The real timer wrap is &HFFFFFFFF -> &H00000000 after ~49.7 days.
'
' Our older wrapper applied a mask:
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
'   start = 2,147,482,647   ' (2^31 - 1000), just before masked wrap
'   now   = 500             ' just after masked wrap
'
'   naive = now - start     ' 500 - 2,147,482,647 = -2,147,482,147
'   If naive > delay Then   ' always False (negative), even though 500 ms passed
'
' Correct approach:
'   - Use the full unmasked tick: GetTickCountRaw()
'   - Convert signed Long bit patterns to unsigned Double before arithmetic
'   - Compute elapsed with wrap-safe math: TicksElapsed(start, now)
'   - Or compare deadlines with unsigned logic: TickAfter(now, deadline)
'
' Double exactly represents every 32-bit integer, so it is safe as a temporary
' unsigned-32 arithmetic representation.
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
'   This breaks across the 2^31 masked wrap (~24.9 days). The subtraction
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
'   - GetTickCountRaw()  : raw modulo-2^32 tick counter bit pattern
'   - TicksElapsed()     : wrap-safe elapsed (ms)
'   - TickAfter()        : wrap-safe "is A at-or-after B?" compare
'   - AddMod32()         : add with wrap (for deadlines/intervals)
'   - PosMod()           : normalize to positive range
'
' ==============================================================
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Const TICKS32 As Double = 4294967296#
Private Const HALF_TICKS32 As Double = 2147483648#

' New raw version (preferred)
Public Function GetTickCountRaw() As Long
    GetTickCountRaw = timeGetTime()
    'to simulare neg stamps do: Or &H80000000
End Function

Public Function TicksElapsed(ByVal startTick As Long, ByVal currentTick As Long) As Double
    Dim elapsed As Double

    elapsed = Unsigned32FromLong(currentTick) - Unsigned32FromLong(startTick)
    If elapsed < 0# Then elapsed = elapsed + TICKS32

    TicksElapsed = elapsed
End Function

Public Function TickAfter(ByVal a As Long, ByVal b As Long) As Boolean
    ' Modular timestamp ordering is only unambiguous for intervals under 2^31 ms.
    TickAfter = (TicksElapsed(b, a) < HALF_TICKS32)
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
    Dim sum As Double

    sum = NormalizeUnsigned32(Unsigned32FromLong(a) + Unsigned32FromLong(b))
    AddMod32 = LongFromUnsigned32(sum)
End Function

' ==============================================================
' DeadlinePassed
' --------------------------------------------------------------
' Wrap-safe check for whether "now" has passed a given deadline.
'
' Why not just use TickAfter(now, deadline)?
' ------------------------------------------
' The comparison itself is handled by TickAfter(). DeadlinePassed() only adds
' the server's legacy sentinel rule: deadline=0 means "no active future
' deadline", so it is always treated as already passed.
'
' Without that sentinel case, a stored zero deadline would become ambiguous
' once real raw ticks legitimately pass through the high signed Long bit.
'
' Usage:
'   If DeadlinePassed(GetTickCountRaw(), counters.StunEndTime) Then ...
'
' ==============================================================
Public Function DeadlinePassed(ByVal nowRaw As Long, ByVal deadline As Long) As Boolean
    If deadline = 0 Then
        DeadlinePassed = True        ' treat 0 as "no deadline"
    Else
        DeadlinePassed = TickAfter(nowRaw, deadline)
    End If
End Function

Private Function Unsigned32FromLong(ByVal value As Long) As Double
    If value < 0 Then
        Unsigned32FromLong = CDbl(value) + TICKS32
    Else
        Unsigned32FromLong = CDbl(value)
    End If
End Function

Private Function LongFromUnsigned32(ByVal value As Double) As Long
    If value >= HALF_TICKS32 Then
        LongFromUnsigned32 = CLng(value - TICKS32)
    Else
        LongFromUnsigned32 = CLng(value)
    End If
End Function

Private Function NormalizeUnsigned32(ByVal value As Double) As Double
    value = value - TICKS32 * Fix(value / TICKS32)
    If value < 0# Then value = value + TICKS32

    NormalizeUnsigned32 = value
End Function
