VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScenarioKillNpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const RequiredWeaponId = 233
Const MaxActiveNpc = 30
Const IniPAth = "/Scenarios/Snakes.ini"

Implements IBaseScenario

Private Type e_Rank
    userIndex As Integer
    score As Integer
End Type

Dim ValidNpc As New Dictionary
Dim InvalidNpc As New Dictionary
Dim SpawnCreatureCount As Integer
Dim ElapsedTime As Long
Dim LastFrameTime As Long
Dim LastScoreUpdate As Long
Dim ActiveNpc As Integer
Dim PlayersScore As Dictionary
Dim MapNumber As Integer
Dim TopFive(5) As e_Rank
Dim Lobby As t_Lobby

Private Sub RegisterNpc(ByVal tag, ByVal sectionSize, ByRef registry As Dictionary, ByRef file As clsIniManager)
    Dim i As Integer
    Dim npcType As Integer
    Dim npcCount As Integer
    For i = 1 To sectionSize
        npcType = val(file.GetValue(tag & i, "NPCID"))
        npcCount = val(file.GetValue(tag & i, "count"))
        Call registry.Add(npcType, npcCount)
    Next i
End Sub

Private Sub Class_Initialize()
        Dim validNpcCount, invalidNpcCount As Integer
        Dim ScenearioIni As clsIniManager
104     Set ScenearioIni = New clsIniManager
        Debug.Assert FileExist(IniPAth & IniPAth)
        
106     Call ScenearioIni.Initialize(IniPAth & IniPAth)
108     validNpcCount = val(ScenearioIni.GetValue("INIT", "TargetNpc"))
        invalidNpcCount = val(ScenearioIni.GetValue("INIT", "OtherNpc"))
        Call RegisterNpc("TARGET", validNpcCount, ValidNpc, ScenearioIni)
        Call RegisterNpc("OTHER", invalidNpcCount, InvalidNpc, ScenearioIni)
End Sub

Public Sub IBaseScenario_Reset()

End Sub

Private Sub IBaseScenario_PlayerKillNpc(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    Dim npcType As Integer
    npcType = NpcList(npcIndex).npcType
    If sourceType = e_phisical And sourceIndex = RequiredWeaponId Then
        If ValidNpc.Exists(npcType) Then
            PlayersScore.Item(userIndex) = PlayersScore.Item(userIndex) + 1
            Call UpdateRanking(PlayersScore.Item(userIndex), userIndex)
        End If
    Else
        Call WriteLocaleMsg(userIndex, 402, FONTTYPE_INFO, ObjData(RequiredWeaponId).name)
    End If
    Call SpawnNpc(npcType)
End Sub

Private Sub IBaseScenario_NPcKillPlayer(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)

End Sub

Private Sub IBaseScenario_PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)

End Sub

Private Sub PublishScoreboard()
    Dim i As Integer
    For i = 0 To UBound(TopFive)
        If TopFive(i).score <= 0 Then
            Exit Sub
        End If
        'send log to everyone in the map
    Next i
End Sub

Private Sub IBaseScenario_Update()
    If Lobby.State <> e_LobbyState.InProgress Then
        Exit Sub
    End If
    Dim frameTime As Long
    ElapsedTime = ElapsedTime + GetTickCount() - LastFrameTime
    If ElapsedTime - LastScoreUpdate > 30000 Then
        LastScoreUpdate = ElapsedTime
        Call PublishScoreboard
    End If
End Sub

Private Sub SpawnNpc(ByVal npcType As Integer)
    Dim spawnPos As t_WorldPos
    Dim npcId As Integer
    spawnPos.map = MapNumber
    spawnPos.X = RandomNumber(1, 99)
    spawnPos.y = RandomNumber(1, 99)
    npcId = CrearNPC(npcType, MapNumber, spawnPos)
    NpcList(npcId).flags.Respawn = 0
End Sub

Public Sub UpdateRanking(ByVal value As Integer, ByVal userIndex As Integer)
    Dim insertPos As Integer
    insertPos = InsertionPos(value)
    If (insertPos <= UBound(TopFive)) Then
        ShiftElement (insertPos)
        TopFive(insertPos).userIndex = userIndex
        TopFive(insertPos).score = value
    End If
End Sub

Private Function InsertionPos(ByVal score As Integer) As Integer
    If TopFive(5).score >= score Then
        InsertionPos = UBound(TopFive) + 1
        Exit Function
    End If
    
    Dim currentPosition As Integer
    For currentPosition = UBound(TopFive) To 1 Step -1
        If TopFive(currentPosition).score >= score Then
            Exit For
        End If
    Next currentPosition
    InsertionPos = currentPosition + 1
End Function

Sub ShiftElement(position As Integer)
    Dim i As Integer
    For i = UBound(TopFive) To position + 1 Step -1
        TopFive(i) = TopFive(i - 1)
    Next i
End Sub
