Attribute VB_Name = "ModUnderworld"
Option Explicit


Public UnderworldLastSpawnTimestamp As Long
Public UnderworldMapPool() As t_WorldPos
Public UnderworldMinSpawnThreshold As Byte
Public UnderworldMaxSpawnThreshold As Byte
Private m_UnderworldLastSpawnTimestamp As Long
Private Const UNDERWORLD_BROADCAST_MSG_ID As Integer = 1

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
            'Call SpawnTp?
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
    ReDim Preserve UnderworldMapPool(1 To MaxUnderworldMaps)
    Dim i As Integer
    For i = 1 To MaxUnderworldMaps
        UnderworldMapPool(i).Map = CInt(val(IniFile.GetValue("Portal" & i, "Map")))
        UnderworldMapPool(i).x = CInt(val(IniFile.GetValue("Portal" & i, "x")))
        UnderworldMapPool(i).y = CInt(val(IniFile.GetValue("Portal" & i, "y")))
    Next i
End Sub
