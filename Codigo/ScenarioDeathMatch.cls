VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScenarioDeathMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const configPath = "../init/Scenarios/DeathMatch.ini"

Implements IBaseScenario

Dim MapNumber As Integer

Public Sub IBaseScenario_Reset()
End Sub

Public Function IBaseScenario_RequiresSpawn() As Boolean
    IBaseScenario_RequiresSpawn = False
End Function

Public Function IBaseScenario_ValidateUser(ByVal UserIndex As Integer) As Integer
End Function

Public Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    Dim spawnPos As t_WorldPos
    spawnPos.map = MapNumber
    spawnPos.X = RandomNumber(12, 85)
    spawnPos.y = RandomNumber(12, 85)
    GenericGlobalLobby.SummonCoordinates = spawnPos
End Sub

Public Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef newState As e_LobbyState)
    If oldState = newState Then
        Exit Sub
    End If
End Sub

Public Sub IBaseScenario_PlayerKillNpc(ByVal npcIndex As Integer, ByVal UserIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
End Sub

Public Sub IBaseScenario_NPcKillPlayer(ByVal npcIndex As Integer, ByVal UserIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
End Sub

Public Sub IBaseScenario_PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
End Sub

Public Sub IBaseScenario_Update()
End Sub

