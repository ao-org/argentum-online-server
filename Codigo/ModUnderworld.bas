Attribute VB_Name = "ModUnderworld"
Option Explicit


Public UnderworldLastSpawnTimestamp As Long
Public UnderworldMapPool() As t_WorldPos
Public OverworldPortalPool() As t_WorldPos
Public UnderworldMinSpawnThreshold As Byte
Public UnderworldMaxSpawnThreshold As Byte
Private m_UnderworldLastSpawnTimestamp As Long
Private Const UNDERWORLD_BROADCAST_MSG_ID As Integer = 2174
Private Const UNDERWORLD_PORTAL_OBJ_IDX As Integer = 6355
Private ALREADY_OPENED_PORTALS As Boolean

Public Sub MaybeSpawnUnderworldPortals()
    On Error GoTo MaybeSpawnUnderworldPortals_Err
    If Not IsFeatureEnabled("underworld") Then Exit Sub
    If UBound(UnderworldMapPool) < 1 Then Exit Sub
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If TicksElapsed(m_UnderworldLastSpawnTimestamp, nowRaw) < IntervalUnderworldSpawn Then Exit Sub
    m_UnderworldLastSpawnTimestamp = nowRaw
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    Dim currentHour As Integer
    currentHour = Hour(Now)
    Dim i As Integer
    Dim SourcePosition As t_WorldPos
    Dim DestinationPosition As t_WorldPos
    If IsUnderworldOpen Then
        If Not ALREADY_OPENED_PORTALS Then
            For i = 1 To UBound(UnderworldMapPool)
                SourcePosition.Map = OverworldPortalPool(i).Map
                SourcePosition.x = OverworldPortalPool(i).x
                SourcePosition.y = OverworldPortalPool(i).y
                DestinationPosition.Map = UnderworldMapPool(i).Map
                DestinationPosition.x = UnderworldMapPool(i).x
                DestinationPosition.y = UnderworldMapPool(i).y
                Call CreateUnderworldTp(SourcePosition, DestinationPosition)
            Next i
                Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(UNDERWORLD_BROADCAST_MSG_ID, vbNullString, e_FontTypeNames.FONTTYPE_CITIZEN))
                Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(e_SoundEffects.Flames, 50, 50))
                ALREADY_OPENED_PORTALS = True
        End If
    Else
        If ALREADY_OPENED_PORTALS Then
            For i = 1 To UBound(UnderworldMapPool)
                SourcePosition.Map = OverworldPortalPool(i).Map
                SourcePosition.x = OverworldPortalPool(i).x
                SourcePosition.y = OverworldPortalPool(i).y
                DestinationPosition.Map = UnderworldMapPool(i).Map
                DestinationPosition.x = UnderworldMapPool(i).x
                DestinationPosition.y = UnderworldMapPool(i).y
                Call DestroyUnderworldTp(SourcePosition, DestinationPosition)
            Next i
            Call KickUsersFromUnderworld
        End If
        ALREADY_OPENED_PORTALS = False
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "ModUnderworld.MaybeSpawnUnderworldPortals", 100)
    Exit Sub
MaybeSpawnUnderworldPortals_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.MaybeSpawnUnderworldPortals", Erl)
End Sub

Public Sub KickUsersFromUnderworld()
    Dim i As Integer
    Dim x As Byte
    Dim y As Byte
    For i = 1 To UBound(UnderworldMapPool)
        For x = MinXBorder To MaxXBorder
            For y = MinYBorder To MaxYBorder
                If MapData(UnderworldMapPool(i).map, x, y).UserIndex > 0 Then
                    With UserList(MapData(UnderworldMapPool(i).map, x, y).UserIndex)
                        Call WarpUserChar(MapData(UnderworldMapPool(i).map, x, y).UserIndex, Ciudades(.Hogar).map, Ciudades(.Hogar).x, Ciudades(.Hogar).y, True)
                    End With
                End If
            Next y
        Next x
    Next i
End Sub

Public Sub DestroyUnderworldTp(ByRef Source As t_WorldPos, ByRef Dest As t_WorldPos)
    If Not MapaValido(Source.Map) Or Not InMapBounds(Source.Map, Source.x, Source.y) Then Exit Sub
    If MapData(Source.Map, Source.x, Source.y).ObjInfo.ObjIndex <> UNDERWORLD_PORTAL_OBJ_IDX Then Exit Sub
    Call EraseObj(MapData(Source.Map, Source.x, Source.y).ObjInfo.ObjIndex, Source.Map, Source.x, Source.y)
    MapData(Source.Map, Source.x, Source.y - 4).TileExit.Map = 0
    MapData(Source.Map, Source.x, Source.y - 4).TileExit.x = 0
    MapData(Source.Map, Source.x, Source.y - 4).TileExit.y = 0
End Sub

Public Sub CreateUnderworldTp(ByRef Source As t_WorldPos, ByRef Dest As t_WorldPos)
    If Not MapaValido(Source.Map) Or Not InMapBounds(Source.Map, Source.x, Source.y) Then Exit Sub
    If Not MapaValido(Dest.Map) Or Not InMapBounds(Dest.Map, Dest.x, Dest.y) Then Exit Sub
        If MapData(Source.Map, Source.x, Source.y).ObjInfo.ObjIndex = UNDERWORLD_PORTAL_OBJ_IDX Then Exit Sub
        If MapData(Source.Map, Source.x, Source.y).ObjInfo.ObjIndex > 0 Then
            Call EraseObj(MapData(Source.Map, Source.x, Source.y).ObjInfo.ObjIndex, Source.Map, Source.x, Source.y)
        End If
        Dim Objeto As t_Obj
        Objeto.Amount = 1
        Objeto.ObjIndex = UNDERWORLD_PORTAL_OBJ_IDX
        Call MakeObj(Objeto, Source.Map, Source.x, Source.y)
        MapData(Source.Map, Source.x, Source.y - 4).TileExit.Map = Dest.Map
        MapData(Source.Map, Source.x, Source.y - 4).TileExit.x = Dest.x
        MapData(Source.Map, Source.x, Source.y - 4).TileExit.y = Dest.y
End Sub

Public Function IsUserIndexInsideTheUnderworld(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim i As Integer
        For i = 1 To UBound(UnderworldMapPool)
            If UserList(UserIndex).pos.Map = UnderworldMapPool(i).Map Then
                IsUserIndexInsideTheUnderworld = True
                Exit Function
            End If
        Next i
    End With
End Function

Public Function IsUnderworldOpen() As Boolean
    If currentHour >= UnderworldMinSpawnThreshold And currentHour < UnderworldMaxSpawnThreshold Then
        IsUnderworldOpen = True
    Else
        IsUnderworldOpen = False
    End If
End Function

Public Sub LoadUnderworldModule()
    m_UnderworldLastSpawnTimestamp = GetTickCountRaw()
    If Not FileExist(DatPath & "UnderworldMapPool.dat", vbArchive) Then
        Debug.Assert False
        Call LogError("Missing file UnderworldMapPool.Dat")
        Exit Sub
    End If
    Dim IniFile     As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "UnderworldMapPool.dat")
    Dim MaxUnderworldMaps
    MaxUnderworldMaps = val(IniFile.GetValue("INIT", "UnderworldMapPool"))
    UnderworldMaxSpawnThreshold = val(IniFile.GetValue("INIT", "UnderworldMaxSpawnThreshold"))
    UnderworldMinSpawnThreshold = val(IniFile.GetValue("INIT", "UnderworldMinSpawnThreshold"))
    If MaxUnderworldMaps <= 0 Then
        Debug.Assert False
        MaxUnderworldMaps = 0
        Exit Sub
    End If
    If UnderworldMaxSpawnThreshold < 0 Or UnderworldMaxSpawnThreshold > 23 Then
        Debug.Assert False
        UnderworldMaxSpawnThreshold = -1
        Exit Sub
    End If
    If UnderworldMinSpawnThreshold < 0 Or UnderworldMinSpawnThreshold > 23 Then
        Debug.Assert False
        UnderworldMaxSpawnThreshold = -1
        Exit Sub
    End If
    ReDim Preserve UnderworldMapPool(1 To MaxUnderworldMaps)
    ReDim Preserve OverworldPortalPool(1 To MaxUnderworldMaps)
    Dim i As Integer
    For i = 1 To MaxUnderworldMaps
        UnderworldMapPool(i).Map = CInt(val(IniFile.GetValue("Portal" & i, "DestinationMap")))
        UnderworldMapPool(i).x = CInt(val(IniFile.GetValue("Portal" & i, "DestinationX")))
        UnderworldMapPool(i).y = CInt(val(IniFile.GetValue("Portal" & i, "DestinationY")))
        OverworldPortalPool(i).Map = CInt(val(IniFile.GetValue("Portal" & i, "SourceMap")))
        OverworldPortalPool(i).x = CInt(val(IniFile.GetValue("Portal" & i, "SourceX")))
        OverworldPortalPool(i).y = CInt(val(IniFile.GetValue("Portal" & i, "SourceY")))
    Next i
End Sub
