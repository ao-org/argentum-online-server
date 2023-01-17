VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpdateHpOverTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime


Dim DotHp As Integer
Dim TargetTicks As Integer
Dim TickCount As Integer
Dim TickTime As Integer
Dim TimeSinceLastTick As Integer

Private DotInfo As t_BaseDotInfo

Public Property Get IBaseEffectOverTime_NextEffect() As IBaseEffectOverTime
    DotInfo.NexEffectt
End Property

Public Property Let IBaseEffectOverTime_NextEffect(ByRef NextElement As IBaseEffectOverTime)
    Set DotInfo.NexEffectt = NextElement
End Property

Public Sub IBaseEffectOverTime_SetTarget(ByVal index As Integer, ByVal RefType As e_ReferenceType)
    Call SetRef(DotInfo.TargetRef, index, RefType)
End Sub

Public Property Get IBaseEffectOverTime_UniqueId() As Integer
    IBaseEffectOverTime_UniqueId = DotInfo.UniqueId
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)

End Sub

Public Function IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = False
End Function

