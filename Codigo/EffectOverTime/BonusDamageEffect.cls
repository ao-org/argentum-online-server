VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BonusDamageEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime

Dim TickTime As Long
Dim TimeSinceLastTick As Long
Dim TickCount As Integer
Dim ApplyEffectId As Integer
Private EffectInfo As t_BaseDotInfo
Private Source As t_AnyReference

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, ByVal ID As Long)
    TimeSinceLastTick = 0
    Call SetBaseDot(EffectInfo, TargetIndex, TargetType, ID, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    Call SetEotEffect
End Sub

Private Sub SetEotEffect()
    With EffectOverTime(EffectInfo.EotId)
        TickTime = .TickTime
        TickCount = .Ticks
        If EffectInfo.TargetRef.RefType = eUser Then
            If .ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(EffectInfo.TargetRef.ArrayIndex, .ClientEffectTypeId, EffectInfo.UniqueId, .TickTime, .buffType, TickCount)
        End If
    End With
    Call EffectsOverTime.ApplyEotModifier(EffectInfo.TargetRef, EffectOverTime(EffectInfo.EotId))
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eBonusDamage
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Integer
    IBaseEffectOverTime_UniqueId = EffectInfo.UniqueId
End Property

Public Property Get IBaseEffectOverTime_CasterIsValid() As Boolean
    IBaseEffectOverTime_CasterIsValid = IsValidRef(EffectInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_CasterArrayIndex() As Integer
    IBaseEffectOverTime_CasterArrayIndex = EffectInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_CasterRefType() As e_ReferenceType
    IBaseEffectOverTime_CasterRefType = EffectInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_CasterUserId() As Long
    IBaseEffectOverTime_CasterUserId = EffectInfo.TargetRef.userID
End Property

Public Property Get IBaseEffectOverTime_TargetIsValid() As Boolean
    IBaseEffectOverTime_TargetIsValid = IsValidRef(EffectInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_TargetArrayIndex() As Integer
    IBaseEffectOverTime_TargetArrayIndex = EffectInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_TargetRefType() As e_ReferenceType
    IBaseEffectOverTime_TargetRefType = EffectInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_TargetUserId() As Long
    IBaseEffectOverTime_TargetUserId = EffectInfo.TargetRef.userID
End Property

Public Property Get IBaseEffectOverTime_EffectType() As e_EffectType
    IBaseEffectOverTime_EffectType = EffectOverTime(EffectInfo.EotId).buffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    TimeSinceLastTick = TimeSinceLastTick + deltaTime
    If TimeSinceLastTick >= TickTime Or Not IsValidRef(EffectInfo.TargetRef) Then
        EffectInfo.RemoveEffect = True
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = EffectInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal Value As Boolean)
    EffectInfo.RemoveEffect = Value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    TimeSinceLastTick = 0
    If EffectInfo.TargetRef.RefType = eUser Then
        If EffectInfo.EotId = NewEffectId Then
            If EffectOverTime(EffectInfo.EotId).ClientEffectTypeId > 0 Then
                Call WriteSendSkillCdUpdate(EffectInfo.TargetRef.ArrayIndex, EffectOverTime(EffectInfo.EotId).ClientEffectTypeId, EffectInfo.UniqueId, _
                                            TickTime, EffectOverTime(EffectInfo.EotId).buffType)
            End If
        Else
            If EffectOverTime(EffectInfo.EotId).ClientEffectTypeId > 0 Then
                Call WriteSendSkillCdUpdate(EffectInfo.TargetRef.ArrayIndex, EffectOverTime(EffectInfo.EotId).ClientEffectTypeId, _
                                            EffectInfo.UniqueId, 0, EffectOverTime(EffectInfo.EotId).buffType)
            End If
        End If
    End If
    EffectInfo.EotId = NewEffectId
    Call EffectsOverTime.ApplyEotModifier(EffectInfo.TargetRef, EffectOverTime(EffectInfo.EotId))
    Call SetEotEffect
    IBaseEffectOverTime_Reset = True
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = EffectInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If EffectInfo.Removed Then Exit Sub
    EffectInfo.Removed = True
    EffectInfo.RemoveEffect = True
    Call EffectsOverTime.RemoveEotModifier(EffectInfo.TargetRef, EffectOverTime(EffectInfo.EotId))
    If EffectInfo.TargetRef.RefType = eUser Then
        If EffectOverTime(EffectInfo.EotId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(EffectInfo.TargetRef.ArrayIndex, EffectInfo.EotId, EffectInfo.UniqueId, 0, EffectOverTime(EffectInfo.EotId).buffType)
        End If
    End If
End Sub

Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
End Sub

Public Sub IBaseEffectOverTime_TartgetWillAtack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TartgetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType) As Long
    TickCount = TickCount - 1
    If TickCount <= 0 Then
        Call IBaseEffectOverTime_OnRemove
    End If
    IBaseEffectOverTime_ApplyDamageReduction = Damage
End Function

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
End Sub

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = e_EffectCallbackMask.eTargetApplyDamageReduction
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub
