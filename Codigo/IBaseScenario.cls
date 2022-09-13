VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBaseScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Reset()
    'empty interface
End Sub

Public Function RequiresSpawn() As Boolean
    RequiresSpawn = False
End Function

Public Function ValidateUser(ByVal userIndex As Integer) As Integer
    'empty interface
End Function

Public Sub WillSummonPlayer(ByVal UserIndex As Integer)
    'empty interface
End Sub

Public Sub UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef newState As e_LobbyState)
'empty interface
End Sub

Public Sub PlayerKillNpc(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    'empty interface
End Sub

Public Sub NPcKillPlayer(ByVal npcIndex As Integer, ByVal userIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    'empty interface
End Sub

Public Sub PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal sourceType As e_DamageSourceType, ByVal sourceIndex As Integer)
    'empty interface
End Sub

Public Sub Update()
    'empty interface
End Sub

Public Sub OnUserDisconnected(ByVal userIndex As Integer)
    
End Sub

Public Sub OnUserReconnect(ByVal userIndex As Integer)

End Sub




