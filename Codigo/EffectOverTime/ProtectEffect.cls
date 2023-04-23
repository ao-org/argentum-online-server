VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProtectEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime

Dim TimeSinceLastTick As Long
Dim DamageProtection As Long
Dim TickTime As Long

Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, ByVal ID As Long)
    TimeSinceLastTick = 0
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, ID, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    With EffectOverTime(EffectId)
        DamageProtection = RandomNumber(.TickPowerMin, .TickPowerMax)
        TickTime = .TickTime
        If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(TargetIndex, .ClientEffectTypeId, ID, .TickTime, .buffType)
    End With
    Call AddShieldToReference(DotInfo.TargetRef, DamageProtection)
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eProtection
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Integer
    IBaseEffectOverTime_UniqueId = DotInfo.UniqueId
End Property

Public Property Get IBaseEffectOverTime_CasterIsValid() As Boolean
    IBaseEffectOverTime_CasterIsValid = IsValidRef(Source)
End Property

Public Property Get IBaseEffectOverTime_CasterArrayIndex() As Integer
    IBaseEffectOverTime_CasterArrayIndex = Source.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_CasterRefType() As e_ReferenceType
    IBaseEffectOverTime_CasterRefType = Source.RefType
End Property

Public Property Get IBaseEffectOverTime_CasterUserId() As Long
    IBaseEffectOverTime_CasterUserId = Source.userID
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
    IBaseEffectOverTime_TargetUserId = DotInfo.TargetRef.userID
End Property

Public Property Get IBaseEffectOverTime_EffectType() As e_EffectType
    IBaseEffectOverTime_EffectType = EffectOverTime(DotInfo.EotId).buffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    TimeSinceLastTick = TimeSinceLastTick + deltaTime
    If TimeSinceLastTick >= TickTime Or Not IsValidRef(DotInfo.TargetRef) Then
        DotInfo.RemoveEffect = True
        Exit Sub
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal Value As Boolean)
    DotInfo.RemoveEffect = Value
End Property

Public Sub IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer)
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, TickTime - TimeSinceLastTick, EffectOverTime(DotInfo.EotId).buffType)
End Sub

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    Call AddShieldToReference(DotInfo.TargetRef, -DamageProtection)
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, EffectOverTime(DotInfo.EotId).buffType)
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
    If EffectOverTime(DotInfo.EotId).OnHitFx > 0 Then
        If IsVisible(UserList(DotInfo.TargetRef.ArrayIndex)) Then
            Call SendData(SendTarget.ToPCAliveArea, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, EffectOverTime(DotInfo.EotId).OnHitFx, 0, UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
        Else
            Call SendData(SendTarget.ToIndex, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, EffectOverTime(DotInfo.EotId).OnHitFx, 0, UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
        End If
    End If
    Dim Reduction As Long
    Reduction = DamageProtection + Damage
    If Reduction > 0 Then
        DamageProtection = Reduction
        Call AddShieldToReference(DotInfo.TargetRef, Damage)
        IBaseEffectOverTime_ApplyDamageReduction = 0
    Else
        Call AddShieldToReference(DotInfo.TargetRef, -DamageProtection)
        DamageProtection = 0
        DotInfo.RemoveEffect = True
        IBaseEffectOverTime_ApplyDamageReduction = Reduction
    End If
End Function

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
End Sub

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = e_EffectCallbackMask.eTargetApplyDamageReduction
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property


