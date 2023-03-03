VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EffectProvoked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime

Dim TickTime As Long
Dim TimeSinceLastTick As Long
Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference
Private PhysicalDamageModfier As Single
Private MagicDamageModifier As Single
Private ReductionActive As Boolean


Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, ByVal ID As Integer)
    Call SetRef(DotInfo.TargetRef, TargetIndex, TargetType)
    Call SetRef(Source, SourceIndex, SourceType)
    DotInfo.UniqueId = ID
    DotInfo.RemoveEffect = False
    DotInfo.EotId = EffectId
    TimeSinceLastTick = 0
    ReductionActive = False
    With EffectOverTime(EffectId)
        TickTime = .TickTime
        PhysicalDamageModfier = .PhysicalDamageDone
        MagicDamageModifier = .MagicDamageDone
        If TargetType = eUser Then
            If .ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(TargetIndex, .ClientEffectTypeId, ID, .TickTime, .BuffType)
        End If
    End With
    If TargetType = eNpc Then
        If SourceType = eUser Then
            Call SetUserRef(NpcList(TargetIndex).TargetUser, SourceIndex)
            NpcList(TargetIndex).flags.AttackedBy = UserList(SourceIndex).name
        Else
            Call SetNpcRef(NpcList(TargetIndex).TargetNPC, SourceIndex)
        End If
        Call SetMask(NpcList(TargetIndex).flags.StatusMask, eTaunted)
        
    End If
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eProvoked
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
    IBaseEffectOverTime_EffectType = EffectOverTime(DotInfo.EotId).BuffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    TimeSinceLastTick = TimeSinceLastTick + deltaTime
    If TimeSinceLastTick >= TickTime Or Not IsValidRef(DotInfo.TargetRef) Or Not IsValidRef(Source) Then
        DotInfo.RemoveEffect = True
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Sub IBaseEffectOverTime_Reset()
End Sub

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.TargetRef.RefType = eNpc Then
        Call UnsetMask(NpcList(DotInfo.TargetRef.ArrayIndex).flags.StatusMask, eTaunted)
    End If
End Sub

Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
End Sub

Public Sub IBaseEffectOverTime_TartgetWillAtack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    Call ApplyReduction(TargetUserId)
End Sub

Public Sub IBaseEffectOverTime_TartgetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    Call RestoreDamage(TargetUserId)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    Call RestoreDamage(TargetUserId)
End Sub

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Private Sub ApplyReduction(ByVal TargetUserId As Integer)
    If TargetUserId <> Source.ArrayIndex And DotInfo.TargetRef.RefType = eUser Then
        Call IncreaseSingle(UserList(DotInfo.TargetRef.ArrayIndex).Modifiers.MagicDamageBonus, MagicDamageModifier)
        Call IncreaseSingle(UserList(DotInfo.TargetRef.ArrayIndex).Modifiers.PhysicalDamageBonus, PhysicalDamageModfier)
        ReductionActive = True
    End If
End Sub

Private Sub RestoreDamage(ByVal TargetUserId As Integer)
    If ReductionActive And DotInfo.TargetRef.RefType = eUser Then
        Call IncreaseSingle(UserList(DotInfo.TargetRef.ArrayIndex).Modifiers.MagicDamageBonus, -MagicDamageModifier)
        Call IncreaseSingle(UserList(DotInfo.TargetRef.ArrayIndex).Modifiers.PhysicalDamageBonus, -PhysicalDamageModfier)
        ReductionActive = False
    End If
End Sub

