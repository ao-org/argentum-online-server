Attribute VB_Name = "ModUnderworld"
Option Explicit

Public UnderworldMapPool() As t_WorldPos
Public OverworldPortalPool() As t_WorldPos
Public IsUnderworldInitialized As Boolean
Private m_UnderworldLastSpawnTimestamp As Long
Private Const UNDERWORLD_BROADCAST_MSG_ID As Integer = 2174
Private Const UNDERWORLD_PORTAL_OBJ_IDX As Integer = 6357
Private ALREADY_OPENED_PORTALS As Boolean
Private Const UNDERWORLD_CENTER_MAP_NUMBER As Integer = 127
Private Const DAY_MIN_OUT_OF_BOUNDS As Integer = 24
Private Const DAY_MAX_OUT_OF_BOUNDS As Integer = -1
Private Const DAY_START As Byte = 0
Private Const DAY_END As Byte = 23

Private Type t_TimeWindow
    Lower As Byte
    Upper As Byte
End Type

Public UnderworldTimeWindows() As t_TimeWindow
Public UnderworldWindowCount As Integer

Public Sub MaybeSpawnUnderworldPortals()
    On Error GoTo MaybeSpawnUnderworldPortals_Err
    If Not IsFeatureEnabled("underworld") Then Exit Sub
    If Not IsUnderworldInitialized Then Exit Sub
    If UBound(UnderworldMapPool) < 1 Then Exit Sub
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If TicksElapsed(m_UnderworldLastSpawnTimestamp, nowRaw) < IntervalUnderworldSpawn Then Exit Sub
    m_UnderworldLastSpawnTimestamp = nowRaw
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
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
    Dim LoopC As Long
    Dim tempIndex As Integer
    Dim HomeCityId As e_City
    
    For i = 1 To UBound(UnderworldMapPool)
        If Not MapaValido(UnderworldMapPool(i).Map) Then
            Debug.Assert False 'invalid map shouldn't happen
            Exit Sub
        End If
        For LoopC = ConnGroups(UnderworldMapPool(i).map).CountEntrys To 1 Step -1
            tempIndex = ConnGroups(UnderworldMapPool(i).Map).UserEntrys(LoopC)
            With UserList(tempIndex)
                HomeCityId = .Hogar
                If .ConnectionDetails.ConnIDValida Then
                    Call WarpUserChar(tempIndex, Cities(HomeCityId).Map, Cities(HomeCityId).x, Cities(HomeCityId).y, True)
                End If
            End With
        Next LoopC
    Next i
    
    If Not MapaValido(UNDERWORLD_CENTER_MAP_NUMBER) Then
        Debug.Assert False 'invalid map shouldn't happen
        Exit Sub
    End If
    
    For LoopC = ConnGroups(UNDERWORLD_CENTER_MAP_NUMBER).CountEntrys To 1 Step -1
        tempIndex = ConnGroups(UNDERWORLD_CENTER_MAP_NUMBER).UserEntrys(LoopC)
        With UserList(tempIndex)
            If .ConnectionDetails.ConnIDValida Then
                HomeCityId = .Hogar
                Call WarpUserChar(tempIndex, Cities(HomeCityId).Map, Cities(HomeCityId).x, Cities(HomeCityId).y, True)
            End If
        End With
    Next LoopC
End Sub

Public Sub DestroyUnderworldTp(ByRef Source As t_WorldPos, ByRef Dest As t_WorldPos)
    If Not MapaValido(Source.Map) Or Not InMapBounds(Source.Map, Source.x, Source.y) Or Not InMapBounds(Source.Map, Source.x, Source.y - 4) Then
        Debug.Assert False
        Exit Sub
    End If
    If MapData(Source.Map, Source.x, Source.y).ObjInfo.ObjIndex <> UNDERWORLD_PORTAL_OBJ_IDX Then Exit Sub
    Call EraseObj(MapData(Source.Map, Source.x, Source.y).ObjInfo.ObjIndex, Source.Map, Source.x, Source.y)
    MapData(Source.Map, Source.x, Source.y - 4).TileExit.Map = 0
    MapData(Source.Map, Source.x, Source.y - 4).TileExit.x = 0
    MapData(Source.Map, Source.x, Source.y - 4).TileExit.y = 0
End Sub

Public Sub CreateUnderworldTp(ByRef Source As t_WorldPos, ByRef Dest As t_WorldPos)
    If Not MapaValido(Source.Map) Or Not InMapBounds(Source.Map, Source.x, Source.y) Or Not InMapBounds(Source.Map, Source.x, Source.y - 4) Then
        Debug.Assert False
        Exit Sub
    End If
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
            If .pos.map = UnderworldMapPool(i).map Then
                IsUserIndexInsideTheUnderworld = True
                Exit Function
            End If
        Next i
            If .pos.map = UNDERWORLD_CENTER_MAP_NUMBER Then
                IsUserIndexInsideTheUnderworld = True
                Exit Function
            End If
    End With
End Function

Public Function IsUnderworldOpen() As Boolean
    Dim currentHour As Integer
    Dim i As Integer
    currentHour = Hour(Now)

    For i = 1 To UnderworldWindowCount
        With UnderworldTimeWindows(i)
            If .Lower > .Upper Then
                ' Wraps midnight
                If currentHour >= .Lower Or currentHour < .Upper Then
                    IsUnderworldOpen = True
                    Exit Function
                End If
            ElseIf .Lower < .Upper Then
                ' Normal range
                If currentHour >= .Lower And currentHour < .Upper Then
                    IsUnderworldOpen = True
                    Exit Function
                End If
            End If
            ' .Lower = .Upper => franja inválida/deshabilitada, se ignora
        End With
    Next i
    IsUnderworldOpen = False
End Function

Public Sub LoadUnderworldModule()
    IsUnderworldInitialized = True
    m_UnderworldLastSpawnTimestamp = GetTickCountRaw()
    If Not FileExist(DatPath & "UnderworldMapPool.dat", vbArchive) Then
        Debug.Assert False
        IsUnderworldInitialized = False
        Call LogError("Missing file UnderworldMapPool.Dat")
        Exit Sub
    End If

    Dim IniFile As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "UnderworldMapPool.dat")

    Dim MaxUnderworldMaps As Integer
    MaxUnderworldMaps = val(IniFile.GetValue("INIT", "UnderworldMapPool"))
    If MaxUnderworldMaps <= 0 Then
        Debug.Assert False
        MaxUnderworldMaps = 0
        IsUnderworldInitialized = False
        Exit Sub
    End If

    UnderworldWindowCount = val(IniFile.GetValue("INIT", "UnderworldWindowCount"))
    If UnderworldWindowCount <= 0 Then
        Debug.Assert False
        Call LogError("UnderworldMapPool.dat: UnderworldWindowCount inválido")
        IsUnderworldInitialized = False
        Exit Sub
    End If

    ReDim UnderworldTimeWindows(1 To UnderworldWindowCount)
    Dim i As Integer
    Dim SectionName As String
    For i = 1 To UnderworldWindowCount
        SectionName = "Window" & i
        UnderworldTimeWindows(i).Lower = val(IniFile.GetValue(SectionName, "LowerLimitOfTime"))
        UnderworldTimeWindows(i).Upper = val(IniFile.GetValue(SectionName, "UpperLimitOfTime"))

        If UnderworldTimeWindows(i).Lower < DAY_START Or UnderworldTimeWindows(i).Lower > DAY_END _
        Or UnderworldTimeWindows(i).Upper < DAY_START Or UnderworldTimeWindows(i).Upper > DAY_END Then
            Debug.Assert False
            Call LogError("UnderworldMapPool.dat: " & SectionName & " fuera de rango")
            IsUnderworldInitialized = False
            Exit Sub
        End If
    Next i

    ReDim Preserve UnderworldMapPool(1 To MaxUnderworldMaps)
    ReDim Preserve OverworldPortalPool(1 To MaxUnderworldMaps)
    For i = 1 To MaxUnderworldMaps
        UnderworldMapPool(i).Map = CInt(val(IniFile.GetValue("Portal" & i, "DestinationMap")))
        UnderworldMapPool(i).x = CInt(val(IniFile.GetValue("Portal" & i, "DestinationX")))
        UnderworldMapPool(i).y = CInt(val(IniFile.GetValue("Portal" & i, "DestinationY")))
        OverworldPortalPool(i).Map = CInt(val(IniFile.GetValue("Portal" & i, "SourceMap")))
        OverworldPortalPool(i).x = CInt(val(IniFile.GetValue("Portal" & i, "SourceX")))
        OverworldPortalPool(i).y = CInt(val(IniFile.GetValue("Portal" & i, "SourceY")))
    Next i
End Sub
