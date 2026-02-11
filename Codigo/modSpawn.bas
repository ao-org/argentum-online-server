Attribute VB_Name = "modSpawn"
' Argentum 20 Game Server
'
'    Copyright (C) 225 Noland Studios LTD
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

#If LOGIN_STRESS_TEST = 1 Then
Public Const SPAWN_SEARCH_MAX_RADIUS As Long = 65
#Else
Public Const SPAWN_SEARCH_MAX_RADIUS As Long = 12
#End If

Private Function MapMinX() As Long: MapMinX = LBound(MapData, 2): End Function
Private Function MapMaxX() As Long: MapMaxX = UBound(MapData, 2): End Function
Private Function MapMinY() As Long: MapMinY = LBound(MapData, 3): End Function
Private Function MapMaxY() As Long: MapMaxY = UBound(MapData, 3): End Function

Private Function InBounds(ByVal M As Integer, ByVal x As Long, ByVal y As Long) As Boolean
    If M < LBound(MapData, 1) Or M > UBound(MapData, 1) Then Exit Function
    If x < MapMinX Or x > MapMaxX Then Exit Function
    If y < MapMinY Or y > MapMaxY Then Exit Function
    InBounds = True
End Function

Private Function IsFreeSpawnTile(ByVal M As Integer, ByVal x As Long, ByVal y As Long, ByVal esAgua As Boolean) As Boolean
    ' Hard guard first: never touch MapData if OOB
    If Not InBounds(M, x, y) Then Exit Function

    ' Keep same legality criteria as old code (agua/tierra)
    Dim ok As Boolean
    If esAgua Then
        ok = LegalPos(m, x, y, True, True, False, False)
    Else
        ok = LegalPos(m, x, y, False, True, False, False)
    End If
    If Not ok Then Exit Function

    ' Finally check for occupancy
    If MapData(M, x, y).UserIndex = 0 And MapData(M, x, y).NpcIndex = 0 Then
        IsFreeSpawnTile = True
    End If
End Function

Public Function FindNearestFreeTile( _
    ByVal M As Integer, _
    ByVal CX As Long, _
    ByVal CY As Long, _
    ByVal esAgua As Boolean, _
    ByVal maxRadius As Long, _
    ByRef outX As Long, _
    ByRef outY As Long _
) As Boolean
    Dim xmin As Long, xmax As Long, ymin As Long, ymax As Long
    xmin = MapMinX: xmax = MapMaxX
    ymin = MapMinY: ymax = MapMaxY

    ' Try center first
    If IsFreeSpawnTile(M, CX, CY, esAgua) Then
        outX = CX: outY = CY
        FindNearestFreeTile = True
        Exit Function
    End If

    ' Clamp the radius so we don't walk past map edges
    Dim r As Long, rMax As Long
    rMax = maxRadius
    If rMax > (xmax - xmin) Then rMax = (xmax - xmin)
    If rMax > (ymax - ymin) Then rMax = (ymax - ymin)

    Dim x As Long, y As Long
    Dim xl As Long, xr As Long, yt As Long, yb As Long

    For r = 1 To rMax
        ' Ring bounds, clamped to map
        xl = CX - r: If xl < xmin Then xl = xmin
        xr = CX + r: If xr > xmax Then xr = xmax
        yt = CY - r: If yt < ymin Then yt = ymin
        yb = CY + r: If yb > ymax Then yb = ymax

        ' Top edge (y=yt)
        y = yt
        For x = xl To xr
            If IsFreeSpawnTile(M, x, y, esAgua) Then outX = x: outY = y: FindNearestFreeTile = True: Exit Function
        Next x

        ' Bottom edge (y=yb) – skip if same row as top
        If yb <> yt Then
            y = yb
            For x = xl To xr
                If IsFreeSpawnTile(M, x, y, esAgua) Then outX = x: outY = y: FindNearestFreeTile = True: Exit Function
            Next x
        End If

        ' Left edge (x=xl), excluding corners already checked
        If xl <> xr Then
            x = xl
            For y = yt + 1 To yb - 1
                If IsFreeSpawnTile(M, x, y, esAgua) Then outX = x: outY = y: FindNearestFreeTile = True: Exit Function
            Next y

            ' Right edge (x=xr)
            x = xr
            For y = yt + 1 To yb - 1
                If IsFreeSpawnTile(M, x, y, esAgua) Then outX = x: outY = y: FindNearestFreeTile = True: Exit Function
            Next y
        End If
    Next r
End Function


