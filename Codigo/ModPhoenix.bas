Attribute VB_Name = "ModPhoenix"
Option Explicit

Public PhoenixLastSpawnTimestamp As Long
Public PhoenixMapPool() As Integer
Private m_LastPhoenixSpawnAttempt As Long
Public IsPhoenixAlive As Boolean
Public Const PHOENIX_NPC_INDEX = 1373
Private Const PHOENIX_BROADCAST_MSG_ID = 2159
Private PhoenixSpawnPosition As t_WorldPos

Public Sub MaybeSpawnFenix()
    On Error GoTo MaybeSpawnFenix_Err
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If TicksElapsed(m_LastPhoenixSpawnAttempt, nowRaw) < IntervalPhoenixSpawn Then Exit Sub
    m_LastPhoenixSpawnAttempt = nowRaw
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    If Not IsPhoenixAlive Then
        PhoenixSpawnPosition.Map = PhoenixMapPool(RandomNumber(LBound(PhoenixMapPool), UBound(PhoenixMapPool)))
        Call SpawnNpc(PHOENIX_NPC_INDEX, PhoenixSpawnPosition, 1, False, False, 0, 3)
        Call modSendData.SendData(ToAll, 0, PrepareMessageLocaleMsg(PHOENIX_BROADCAST_MSG_ID, vbNullString, e_FontTypeNames.FONTTYPE_CITIZEN))
        IsPhoenixAlive = True
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "ModPhoenix.MaybeSpawnFenix", 100)
    Exit Sub
MaybeSpawnFenix_Err:
    Call TraceError(Err.Number, Err.Description, "ModPhoenix.MaybeSpawnFenix", Erl)
End Sub


Public Sub LoadPhoenixModule()
    m_LastPhoenixSpawnAttempt = GetTickCountRaw()
    If Not FileExist(DatPath & "PhoenixMapPool.dat", vbArchive) Then
        Debug.Assert False
        Exit Sub
    End If
    Dim IniFile     As clsIniManager
    Set IniFile = New clsIniManager
    Call IniFile.Initialize(DatPath & "PhoenixMapPool.dat")
    Dim MaxPhoenixMaps
    MaxPhoenixMaps = val(IniFile.GetValue("INIT", "MaxPhoenixMaps"))
    ReDim Preserve PhoenixMapPool(1 To MaxPhoenixMaps)
    Dim i As Integer
    For i = 1 To MaxPhoenixMaps
        PhoenixMapPool(i) = CLng(val(IniFile.GetValue("Maps", "Map" & i)))
    Next i
    PhoenixSpawnPosition.x = 50
    PhoenixSpawnPosition.y = 50
End Sub
