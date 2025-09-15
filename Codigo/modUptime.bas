Attribute VB_Name = "modUptime"
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
'
'
'
Option Explicit

' When the process starts, we save the current time.
Public g_ServerStart As Date

Public Sub Uptime_Init()
    g_ServerStart = Now
End Sub

' Total seconds since start (safe for many years).
Public Function GetServerUptimeSeconds() As Long
    GetServerUptimeSeconds = DateDiff("s", g_ServerStart, Now)
End Function

' Human-friendly formatting: 3d 12:34:56
Public Function FormatUptime() As String
    Dim total As Long, d As Long, h As Long, m As Long, s As Long
    
    total = GetServerUptimeSeconds()
    d = total \ 86400: total = total Mod 86400
    h = total \ 3600: total = total Mod 3600
    m = total \ 60:   s = total Mod 60
    
    FormatUptime = CStr(d) & "d " & Right$("0" & h, 2) & ":" & Right$("0" & m, 2) & ":" & Right$("0" & s, 2)
End Function

