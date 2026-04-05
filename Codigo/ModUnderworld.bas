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
    If currentHour >= UnderworldMinSpawnThreshold And currentHour < UnderworldMaxSpawnThreshold Then
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
        End If
        ALREADY_OPENED_PORTALS = False
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "ModUnderworld.MaybeSpawnUnderworldPortals", 100)
    Exit Sub
MaybeSpawnUnderworldPortals_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.MaybeSpawnUnderworldPortals", Erl)
End Sub

Public Sub DestroyUnderworldTp(ByRef Source As t_WorldPos, ByRef Dest As t_WorldPos)
    If Not MapaValido(Source.Map) Or Not InMapBounds(Source.Map, Source.x, Source.y) Then Exit Sub
    If MapData(Source.Map, Source.x, Source.y).ObjInfo.ObjIndex <> UNDERWORLD_PORTAL_OBJ_IDX Then Exit Sub
    Call EraseObj(MapData(Source.Map, Source.x, Source.y).ObjInfo.ObjIndex, Source.Map, Source.x, Source.y)
    Dim i As Integer
    Dim j As Integer
    For i = Source.y - 6 To Source.y
        For j = Source.x - 4 To Source.x + 4
            If InMapBounds(Source.Map, j, i) Then
                MapData(Source.Map, j, i).TileExit.Map = 0
                MapData(Source.Map, j, i).TileExit.x = 0
                MapData(Source.Map, j, i).TileExit.y = 0
            End If
        Next j
    Next i
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
        Dim i As Integer
        Dim j As Integer
        For i = Source.y - 6 To Source.y
            For j = Source.x - 4 To Source.x + 4
                MapData(Source.Map, j, i).TileExit.Map = Dest.Map
                MapData(Source.Map, j, i).TileExit.x = Dest.x
                MapData(Source.Map, j, i).TileExit.y = Dest.y
            Next j
        Next i
End Sub

Public Function IsPlayerInsideTheUnderworld(ByVal UserIndex As Integer)
    With UserList(UserIndex)
    End With
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
