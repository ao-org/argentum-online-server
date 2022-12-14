VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRespawnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RespawnTime As Long
Private EquipedSlots(8) As Integer
Public EquipedElementCount As Integer
Public Strength As Byte
Public Agility As Byte
Public PowerUpTime As Integer
Public HasPowerUp As Boolean
Public UserSlotIteration As Long

Public Sub Reset()
    EquipedElementCount = 0
End Sub

Public Sub AddEquipedElement(ByVal Slot As Integer)
    EquipedSlots(EquipedElementCount) = Slot
    EquipedElementCount = EquipedElementCount + 1
End Sub

Public Function GetActiveSlot(ByVal index As Integer) As Integer
    Debug.Assert index < EquipedElementCount
    GetActiveSlot = EquipedSlots(index)
End Function


