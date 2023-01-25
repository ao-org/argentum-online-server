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
Dim TickTime As Long
Dim TimeSinceLastTick As Long
Dim TickEffect As Integer

Private DotInfo As t_BaseDotInfo
Private source As t_AnyReference

Public Sub Setup(ByVal sourceIndex As Integer, ByVal sourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, ByVal ID As Integer)
    TimeSinceLastTick = 0
    TickCount = 0
    Call SetRef(DotInfo.TargetRef, TargetIndex, TargetType)
    Call SetRef(source, sourceIndex, sourceType)
    DotInfo.UniqueId = ID
    DotInfo.RemoveEffect = False
    DotInfo.EotId = EffectId
    With EffectOverTime(EffectId)
        DotHp = .TickPower
        TargetTicks = .Ticks
        TickTime = .TickTime
        DotInfo.EotType = .Type
        TickEffect = .TickFX
    End With
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = DotInfo.EotType
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Integer
    IBaseEffectOverTime_UniqueId = DotInfo.UniqueId
End Property

Public Property Get IBaseEffectOverTime_CasterIsValid() As Boolean
    IBaseEffectOverTime_CasterIsValid = IsValidRef(source)
End Property

Public Property Get IBaseEffectOverTime_CasterArrayIndex() As Integer
    IBaseEffectOverTime_CasterArrayIndex = source.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_CasterRefType() As e_ReferenceType
    IBaseEffectOverTime_CasterRefType = source.RefType
End Property

Public Property Get IBaseEffectOverTime_CasterUserId() As Long
    IBaseEffectOverTime_CasterUserId = source.UserId
End Property

Public Property Get IBaseEffectOverTime_TargetIsValid() As Boolean
    IBaseEffectOverTime_TargetIsValid = IsValidRef(DotInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_TargetArrayIndex() As Integer
    IBaseEffectOverTime_TargetArrayIndex = DotInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_TargetRefType() As e_ReferenceType
    IBaseEffectOverTime_TargetRefType = DotInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_TargetUserId() As Long
    IBaseEffectOverTime_TargetUserId = DotInfo.TargetRef.UserId
End Property

Public Property Get IBaseEffectOverTime_EffectType() As e_EffectType
    IBaseEffectOverTime_EffectType = EffectOverTime(DotInfo.EotId).BuffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    If TickCount >= TargetTicks Or Not IsValidRef(DotInfo.TargetRef) Then
        DotInfo.RemoveEffect = True
        Exit Sub
    End If
    TimeSinceLastTick = TimeSinceLastTick + DeltaTime
    If TimeSinceLastTick >= TickTime Then
        TimeSinceLastTick = 0
        TickCount = TickCount + 1
        Call PerformAction
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Sub IBaseEffectOverTime_Reset()
    TickCount = 0
End Sub

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub PerformAction()
    If DotInfo.TargetRef.RefType = e_ReferenceType.eUser Then
        If IsValidRef(source) Then
            Call UserMod.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, source.ArrayIndex, DotHp, e_dot)
        Else
            Call UserMod.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, 0, DotHp, e_dot)
        End If
        If TickEffect > 0 And IsVisible(UserList(DotInfo.TargetRef.ArrayIndex)) = 0 Then Call SendData(SendTarget.ToPCAliveArea, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, TickEffect, 0, UserList(DotInfo.TargetRef.ArrayIndex).pos.X, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
    ElseIf DotInfo.TargetRef.RefType = e_ReferenceType.eNpc Then
        If IsValidRef(source) Then
            Call NPCs.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, source.ArrayIndex, DotHp, e_dot)
        Else
            Call NPCs.DoDamageOrHeal(DotInfo.TargetRef.ArrayIndex, 0, DotHp, e_dot)
        End If
        If TickEffect > 0 Then Call SendData(SendTarget.ToNPCAliveArea, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(NpcList(DotInfo.TargetRef.ArrayIndex).Char.charindex, TickEffect, 0, NpcList(DotInfo.TargetRef.ArrayIndex).pos.X, NpcList(DotInfo.TargetRef.ArrayIndex).pos.y))
    End If
End Sub


