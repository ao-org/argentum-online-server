VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultipleAttacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Implements IBaseEffectOverTime
Private mid                 As Long
Private mDotInfo            As t_BaseDotInfo
Private mElapsedTime        As Long
Private mTimeLimit          As Long
Private mTicks              As Long
Private mPhysicalDamageDone As Single
Private mBlowCount          As Integer
Private mPendingBlow        As Integer
Private mPendingBlowTime    As Integer
Private mPendingBlowPos     As t_WorldPos

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal EffectId As Integer, ByVal Id As Long)
    mElapsedTime = 0
    Call SetBaseDot(mDotInfo, SourceIndex, SourceType, Id, EffectId)
    With EffectOverTime(EffectId)
        mTimeLimit = .TickTime
        mPhysicalDamageDone = .PhysicalDamageDone
        mBlowCount = .SubType
        mTicks = .Ticks
        If mDotInfo.TargetRef.RefType = eUser And EffectOverTime(mDotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(SourceIndex, .ClientEffectTypeId, Id, _
                .TickTime, .TickTime, .buffType, mTicks)
    End With
    Call UpdateIncreaseModifier(mDotInfo.TargetRef, PhysiccalBonus, mPhysicalDamageDone)
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eMultipleAttacks
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Long
    IBaseEffectOverTime_UniqueId = mDotInfo.UniqueId
End Property

Public Property Get IBaseEffectOverTime_SharedTypeId() As Integer
    IBaseEffectOverTime_SharedTypeId = EffectOverTime(mDotInfo.EotId).SharedTypeId
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
    IBaseEffectOverTime_CasterUserId = mDotInfo.TargetRef.UserId
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
    IBaseEffectOverTime_TargetUserId = mDotInfo.TargetRef.UserId
End Property

Public Property Get IBaseEffectOverTime_EffectType() As e_EffectType
    IBaseEffectOverTime_EffectType = EffectOverTime(mDotInfo.EotId).buffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal DeltaTime As Long)
    mElapsedTime = mElapsedTime + DeltaTime
    If mElapsedTime >= mTimeLimit Or Not IsValidRef(mDotInfo.TargetRef) Then
        mDotInfo.RemoveEffect = True
        Exit Sub
    End If
    If mPendingBlow > 0 Then
        mPendingBlowTime = mPendingBlowTime + DeltaTime
        If mPendingBlowTime > 100 Then
            Call UserAttackPosition(mDotInfo.TargetRef.ArrayIndex, mPendingBlowPos, True)
            mPendingBlow = mPendingBlow - 1
            If mPendingBlow = 0 And mTicks = 0 Then
                mDotInfo.RemoveEffect = True
            End If
        End If
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = mDotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    mDotInfo.RemoveEffect = value
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
    If IsValidRef(mDotInfo.TargetRef) Then
        Call UpdateIncreaseModifier(mDotInfo.TargetRef, PhysiccalBonus, -mPhysicalDamageDone)
    End If
    If mDotInfo.TargetRef.RefType = eUser And EffectOverTime(mDotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(mDotInfo.TargetRef.ArrayIndex, EffectOverTime(mDotInfo.EotId).ClientEffectTypeId, mDotInfo.UniqueId, 0, 0, EffectOverTime( _
                mDotInfo.EotId).buffType)
    End If
End Sub

Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
End Sub

Public Sub IBaseEffectOverTime_TargetWillAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Function IBaseEffectOverTime_ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
    IBaseEffectOverTime_ChangeTarget = False
End Function

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, _
                                                         ByVal SourceUserId As Integer, _
                                                         ByVal SourceType As e_ReferenceType, _
                                                         ByVal AttackType As e_DamageSourceType) As Long
    IBaseEffectOverTime_ApplyDamageReduction = Damage
End Function

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
    If UserList(mDotInfo.TargetRef.ArrayIndex).invent.EquippedWeaponObjIndex > 0 Then
        If ObjData(UserList(mDotInfo.TargetRef.ArrayIndex).invent.EquippedWeaponObjIndex).WeaponType <> eKnuckle Then
            mDotInfo.RemoveEffect = True
            Exit Sub
        End If
    End If
    Dim i As Integer
    mPendingBlowPos.Map = Map
    mPendingBlowPos.x = PosX
    mPendingBlowPos.y = PosY
    mPendingBlow = mBlowCount
    mPendingBlowTime = 0
    mTicks = mTicks - 1
    If mDotInfo.TargetRef.RefType = eUser And EffectOverTime(mDotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(mDotInfo.TargetRef.ArrayIndex, EffectOverTime(mDotInfo.EotId).ClientEffectTypeId, mDotInfo.UniqueId, mTimeLimit - mElapsedTime, mTimeLimit, _
                EffectOverTime(mDotInfo.EotId).buffType, mTicks)
    End If
End Sub

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = e_EffectCallbackMask.eTargetWillAttackPosition
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub
