VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnequipItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime

Private mDotInfo As t_BaseDotInfo
Private mElapsedTime As Long
Private mTimeLimit As Long
Private mUnequipedslotMask As Long
Private mDebuffIcon As Integer

Public Sub Setup(ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, ByVal EffectId As Integer, ByVal ID As Long, ByVal ItemSlotMask As Long)
    Debug.Assert TargetType = eUser
    mElapsedTime = 0
    Call SetBaseDot(mDotInfo, TargetIndex, TargetType, ID, EffectId)
    mUnequipedslotMask = ItemSlotMask
    Call SetMask(UserList(TargetIndex).flags.DisabledSlot, mUnequipedslotMask)
    Select Case ItemSlotMask
        Case e_InventorySlotMask.eWeapon
            mDebuffIcon = 18
        Case e_InventorySlotMask.eShiled
            mDebuffIcon = 19
        Case e_InventorySlotMask.eHelm
            mDebuffIcon = 20
    End Select
    With EffectOverTime(EffectId)
        mTimeLimit = .TickTime
        If mDotInfo.TargetRef.RefType = eUser And EffectOverTime(mDotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(TargetIndex, .ClientEffectTypeId, ID, .TickTime, .buffType)
    End With
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eUnequip
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Integer
    IBaseEffectOverTime_UniqueId = mDotInfo.UniqueId
End Property

Public Property Get IBaseEffectOverTime_CasterIsValid() As Boolean
    IBaseEffectOverTime_CasterIsValid = IsValidRef(mDotInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_CasterArrayIndex() As Integer
    IBaseEffectOverTime_CasterArrayIndex = mDotInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_CasterRefType() As e_ReferenceType
    IBaseEffectOverTime_CasterRefType = mDotInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_CasterUserId() As Long
    IBaseEffectOverTime_CasterUserId = mDotInfo.TargetRef.userID
End Property

Public Property Get IBaseEffectOverTime_TargetIsValid() As Boolean
    IBaseEffectOverTime_TargetIsValid = IsValidRef(mDotInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_TargetArrayIndex() As Integer
    IBaseEffectOverTime_TargetArrayIndex = mDotInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_TargetRefType() As e_ReferenceType
    IBaseEffectOverTime_TargetRefType = mDotInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_TargetUserId() As Long
    IBaseEffectOverTime_TargetUserId = mDotInfo.TargetRef.userID
End Property

Public Property Get IBaseEffectOverTime_EffectType() As e_EffectType
    IBaseEffectOverTime_EffectType = EffectOverTime(mDotInfo.EotId).buffType
End Property

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
End Sub

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = 0
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    mElapsedTime = mElapsedTime + deltaTime
    If mElapsedTime >= mTimeLimit Or Not IsValidRef(mDotInfo.TargetRef) Then
        mDotInfo.RemoveEffect = True
        Exit Sub
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = mDotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal Value As Boolean)
    mDotInfo.RemoveEffect = Value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    IBaseEffectOverTime_Reset = True
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = mDotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If mDotInfo.Removed Then Exit Sub
    mDotInfo.Removed = True
    mDotInfo.RemoveEffect = True
    Call UnsetMask(UserList(mDotInfo.TargetRef.ArrayIndex).flags.DisabledSlot, mUnequipedslotMask)
    If mDotInfo.TargetRef.RefType = eUser And EffectOverTime(mDotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(mDotInfo.TargetRef.ArrayIndex, EffectOverTime(mDotInfo.EotId).ClientEffectTypeId, _
                                    mDotInfo.UniqueId, 0, EffectOverTime(mDotInfo.EotId).buffType)
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
    IBaseEffectOverTime_ApplyDamageReduction = Damage
End Function

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub




