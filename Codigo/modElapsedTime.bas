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
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Const TICKS32 As Double = 4294967296#

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

Public Function DeadlinePassed(ByVal nowRaw As Long, ByVal deadline As Long) As Boolean
    If deadline = 0 Then
        DeadlinePassed = True        ' treat 0 as "no deadline"
    Else
        DeadlinePassed = (nowRaw - deadline) >= 0   ' wrap-safe TickAfter
    End If
End Function
