Attribute VB_Name = "modTileProperties"
' Copyright (C) 2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit

Public Const KNOWN_TILE_FLAGS As Long = 511
Public Const KNOWN_ZONE_FLAGS As Long = 2097151

Public Function HasTileFlag(ByVal flags As Long, ByVal flag As e_Trigger) As Boolean
    HasTileFlag = (flags And flag) <> 0
End Function

Public Function HasZoneFlag(ByVal flags As Long, ByVal flag As e_ZoneFlags) As Boolean
    HasZoneFlag = (flags And flag) <> 0
End Function

Public Function HasMapZoneFlag(ByVal map As Long, ByVal flag As e_ZoneFlags) As Boolean
    HasMapZoneFlag = HasZoneFlag(MapInfo(map).ZoneFlags, flag)
End Function

Public Sub SetMapZoneFlag(ByVal map As Long, ByVal flag As e_ZoneFlags, ByVal enabled As Boolean)
    If enabled Then
        MapInfo(map).ZoneFlags = MapInfo(map).ZoneFlags Or flag
    Else
        MapInfo(map).ZoneFlags = MapInfo(map).ZoneFlags And Not flag
    End If
End Sub

Public Function IsPrisonMap(ByVal map As Integer) As Boolean
    IsPrisonMap = HasMapZoneFlag(map, e_ZoneFlags.Prison)
End Function

Public Function IsInPvPArena(ByVal userIndex As Integer) As Boolean
    With UserList(userIndex).pos
        IsInPvPArena = HasTileFlag(MapData(.x, .y, .map).trigger, e_Trigger.PvPArena)
    End With
End Function

Public Function BothInPvPArena(ByVal source As Integer, ByVal target As Integer) As Boolean
    BothInPvPArena = IsInPvPArena(source) And IsInPvPArena(target)
End Function

Public Function CrossesPvPArenaBoundary(ByVal source As Integer, ByVal target As Integer) As Boolean
    CrossesPvPArenaBoundary = IsInPvPArena(source) Xor IsInPvPArena(target)
End Function

Public Function IsWaterTile(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    With MapData(x, y, map)
        IsWaterTile = (.Blocked And FLAG_AGUA) <> 0 Or HasTileFlag(.trigger, e_Trigger.SwimSuitPath)
    End With
End Function

Public Function IsSwimmingSuit(ByVal objIndex As Integer) As Boolean
    IsSwimmingSuit = objIndex = iObjTraje Or objIndex = iObjTrajeAltoNw Or objIndex = iObjTrajeBajoNw
End Function

Public Function HasAdjacentSwimSuitPath(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    Dim dx As Integer, dy As Integer, direction As Integer
    For direction = 0 To 3
        dx = 0: dy = 0
        Select Case direction
            Case 0: dx = -1
            Case 1: dx = 1
            Case 2: dy = -1
            Case 3: dy = 1
        End Select
        If InMapBounds(map, x + dx, y + dy) Then
            If HasTileFlag(MapData(x + dx, y + dy, map).trigger, e_Trigger.SwimSuitPath) Then
                HasAdjacentSwimSuitPath = True
                Exit Function
            End If
        End If
    Next direction
End Function

' Historical values are interpreted only at the legacy file boundary.
Public Function ConvertLegacyTrigger(ByVal legacy As Integer, ByRef zoneFlags As Long) As Long
    Select Case legacy
        Case 0, 4, 12, 13, 16, 17, 20, 200, 201
            ConvertLegacyTrigger = e_Trigger.None
        Case 1, 60 To 73, 90 To 99
            ConvertLegacyTrigger = e_Trigger.UnderRoof
        Case 2, 3
            ConvertLegacyTrigger = e_Trigger.AntiNpcRespawn
        Case 5
            ConvertLegacyTrigger = e_Trigger.PathUnblocker
        Case 6
            ConvertLegacyTrigger = e_Trigger.PvPArena
        Case 7
            ConvertLegacyTrigger = e_Trigger.AutoResurrection
        Case 8, 11, 18
            ConvertLegacyTrigger = e_Trigger.SwimSuitPath
        Case 10
            ConvertLegacyTrigger = e_Trigger.NoFishing
        Case 14
            ConvertLegacyTrigger = e_Trigger.GhostOnlyTranslator
        Case 19
            zoneFlags = zoneFlags Or e_ZoneFlags.Prison
        Case 21 To 40
            ' Caller records a pending reference resolved from the database audit mapping.
        Case 41
            ConvertLegacyTrigger = e_Trigger.CastleFoundationPosition
        Case Else
            Call Err.Raise(5, "ConvertLegacyTrigger", "Unknown legacy trigger: " & CStr(legacy))
    End Select
End Function

Public Function IsMapDataCoordinate(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    ' Stored map geometry includes the outer margin excluded by gameplay InMapBounds.
    IsMapDataCoordinate = map > 0 And map <= NumMaps And x >= XMinMapSize And x <= XMaxMapSize And y >= YMinMapSize And y <= YMaxMapSize
End Function

Public Function TilePropertyKey(ByVal x As Integer, ByVal y As Integer) As Long
    TilePropertyKey = CLng(y) * 65536 + x
End Function

Public Function CastleAtTile(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Long
    If MapInfo(map).CastleEntrances Is Nothing Then Exit Function
    Dim key As Long
    key = TilePropertyKey(x, y)
    If MapInfo(map).CastleEntrances.Exists(key) Then CastleAtTile = MapInfo(map).CastleEntrances.Item(key)
End Function

Public Sub RegisterCastleEntrance(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal castleId As Long, Optional ByVal authored As Boolean = False)
    If Not IsMapDataCoordinate(map, x, y) Or castleId <= 0 Then Call Err.Raise(5, "RegisterCastleEntrance", "Invalid castle entrance")
    With MapInfo(map)
        If .CastleEntrances Is Nothing Then Set .CastleEntrances = New Dictionary
        If .CastleEntrances.Exists(TilePropertyKey(x, y)) Then Call Err.Raise(5, "RegisterCastleEntrance", "Conflicting castle entrance")
        Call .CastleEntrances.Add(TilePropertyKey(x, y), castleId)
        If authored Then
            If .StaticCastleEntrances Is Nothing Then Set .StaticCastleEntrances = New Dictionary
            Call .StaticCastleEntrances.Add(TilePropertyKey(x, y), castleId)
        End If
    End With
End Sub

Public Sub RemoveCastleEntrances(ByVal map As Integer, ByVal castleId As Long)
    Dim key As Variant
    With MapInfo(map)
        If .CastleEntrances Is Nothing Then Exit Sub
        For Each key In .CastleEntrances.Keys
            If .CastleEntrances.Item(key) = castleId Then
                If .StaticCastleEntrances Is Nothing Then
                    Call .CastleEntrances.Remove(key)
                ElseIf Not .StaticCastleEntrances.Exists(key) Then
                    Call .CastleEntrances.Remove(key)
                End If
            End If
        Next key
    End With
End Sub

Public Function CopyPropertyDictionary(ByVal source As Dictionary) As Dictionary
    If source Is Nothing Then Exit Function
    Dim result As Dictionary, key As Variant
    Set result = New Dictionary
    For Each key In source.Keys
        Call result.Add(key, source.Item(key))
    Next key
    Set CopyPropertyDictionary = result
End Function

Public Function SetTileTriggerFlags(ByVal userIndex As Integer, ByVal flags As Long) As Boolean
    If Not EsGM(userIndex) Then Exit Function
    If flags < 0 Or (flags And Not KNOWN_TILE_FLAGS) <> 0 Then Exit Function
    With UserList(userIndex).pos
        If Not InMapBounds(.map, .x, .y) Then Exit Function
        MapData(.x, .y, .map).trigger = flags
        If MapInfo(.map).TileOverrides Is Nothing Then Set MapInfo(.map).TileOverrides = New Dictionary
        MapInfo(.map).TileOverrides.Item(TilePropertyKey(.x, .y)) = flags
        Dim recipient As Integer
        For recipient = 1 To LastUser
            If UserList(recipient).flags.UserLogged And UserList(recipient).pos.map = .map Then
                Call WriteHooTileProperties(recipient, .map, .x, .y, flags)
            End If
        Next recipient
    End With
    SetTileTriggerFlags = True
End Function

Public Sub ReplayTileProperties(ByVal userIndex As Integer, ByVal map As Integer)
    If MapInfo(map).TileOverrides Is Nothing Then Exit Sub
    Dim key As Variant
    For Each key In MapInfo(map).TileOverrides.Keys
        Call WriteHooTileProperties(userIndex, map, CInt(CLng(key) And &HFFFF&), CInt(CLng(key) \ 65536), MapInfo(map).TileOverrides.Item(key))
    Next key
End Sub
