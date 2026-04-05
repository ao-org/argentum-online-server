Attribute VB_Name = "ModUnderworld"
Option Explicit


Public UnderworldLastSpawnTimestamp As Long
Public UnderworldMapPool() As t_WorldPos
Public OverworldPortalPool() As t_WorldPos
Public UnderworldMinSpawnThreshold As Byte
Public UnderworldMaxSpawnThreshold As Byte
Private m_UnderworldLastSpawnTimestamp As Long
Private Const UNDERWORLD_BROADCAST_MSG_ID As Integer = 1
Private Const UNDERWORLD_PORTAL_OBJ_IDX As Integer = 6355

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
    
    If currentHour >= UnderworldMinSpawnThreshold And currentHour < UnderworldMaxSpawnThreshold Then
        Dim i As Integer
        For i = 1 To UBound(UnderworldMapPool)
            Dim SourcePosition As t_WorldPos
            Dim DestinationPosition As t_WorldPos
            SourcePosition.Map = UnderworldMapPool(i).Map
            SourcePosition.x = UnderworldMapPool(i).x
            SourcePosition.y = UnderworldMapPool(i).y
            DestinationPosition.Map = OverworldPortalPool(i).Map
            DestinationPosition.x = OverworldPortalPool(i).x
            DestinationPosition.y = OverworldPortalPool(i).y
           Call CreateUnderworldTp(SourcePosition, DestinationPosition)
        Next i
        Call modSendData.SendData(ToAll, 0, PrepareMessageLocaleMsg(UNDERWORLD_BROADCAST_MSG_ID, vbNullString, e_FontTypeNames.FONTTYPE_CITIZEN))
    
    Else
            'Call DestroyTPs?
        Exit Sub
    End If
    
    
    
    Call PerformTimeLimitCheck(PerformanceTimer, "ModUnderworld.MaybeSpawnUnderworldPortals", 100)
    Exit Sub
MaybeSpawnUnderworldPortals_Err:
    Call TraceError(Err.Number, Err.Description, "ModUnderworld.MaybeSpawnUnderworldPortals", Erl)
End Sub


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


Public Sub DestroyUnderworldTps(ByRef Source As t_WorldPos, ByRef Dest As t_WorldPos)

End Sub

Public Sub CreateUnderworldTp(ByRef Source As t_WorldPos, ByRef Dest As t_WorldPos)
    If Not MapaValido(Source.Map) Or Not InMapBounds(Source.Map, Source.x, Source.y) Then Exit Sub
    If Not MapaValido(Dest.Map) Or Not InMapBounds(Dest.Map, Dest.x, Dest.y) Then Exit Sub
    With MapData(Source.Map, Source.x, Source.y)
        If .ObjInfo.ObjIndex = UNDERWORLD_PORTAL_OBJ_IDX Then Exit Sub
        If .ObjInfo.ObjIndex > 0 Then
            Call EraseObj(.ObjInfo.ObjIndex, Source.Map, Source.x, Source.y)
        End If
        Dim Objeto As t_Obj
        Objeto.Amount = 1
        Objeto.ObjIndex = UNDERWORLD_PORTAL_OBJ_IDX
        Call MakeObj(Objeto, Source.Map, Source.x, Source.y)
        .TileExit.Map = Dest.Map
        .TileExit.x = Dest.x
        .TileExit.y = Dest.y
    End With
    
        
End Sub

Public Function IsPlayerInsideTheUnderworld(ByVal UserIndex As Integer)
    With UserList(UserIndex)
    End With
End Function


