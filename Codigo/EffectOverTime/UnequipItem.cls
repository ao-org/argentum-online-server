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

Private mDotInfo As t_BaseDotInfo
Private mElapsedTime As Long
Private mTimeLimit As Long
Private mUnequipedslotMask As Long
Private mDebuffIcon As Integer

Public Sub Setup(ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, ByVal EffectId As Integer, ByVal ID As Long, ByVal ItemSlotMask As Long)
    On Error Goto Setup_Err
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
        If mDotInfo.TargetRef.RefType = eUser And mDebuffIcon > 0 Then
            Call WriteSendSkillCdUpdate(TargetIndex, mDebuffIcon, id, .TickTime, .TickTime, .buffType)
        End If
    End With
    Exit Sub
Setup_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.Setup", Erl)
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eUnequip
End Property

Public Property Get IBaseEffectOverTime_SharedTypeId() As Integer
    IBaseEffectOverTime_SharedTypeId = EffectOverTime(mDotInfo.EotId).SharedTypeId
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Long
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
    On Error Goto IBaseEffectOverTime_TargetWillAttackPosition_Err
    Exit Sub
IBaseEffectOverTime_TargetWillAttackPosition_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_TargetWillAttackPosition", Erl)
End Sub

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = 0
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    On Error Goto IBaseEffectOverTime_Update_Err
    mElapsedTime = mElapsedTime + deltaTime
    If mElapsedTime >= mTimeLimit Or Not IsValidRef(mDotInfo.TargetRef) Then
        mDotInfo.RemoveEffect = True
        Exit Sub
    End If
    Exit Sub
IBaseEffectOverTime_Update_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_Update", Erl)
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = mDotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal Value As Boolean)
    mDotInfo.RemoveEffect = Value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    On Error Goto IBaseEffectOverTime_Reset_Err
    IBaseEffectOverTime_Reset = True
    Exit Function
IBaseEffectOverTime_Reset_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_Reset", Erl)
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = mDotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    On Error Goto IBaseEffectOverTime_OnRemove_Err
    If mDotInfo.Removed Then Exit Sub
    mDotInfo.Removed = True
    mDotInfo.RemoveEffect = True
    Call UnsetMask(UserList(mDotInfo.TargetRef.ArrayIndex).flags.DisabledSlot, mUnequipedslotMask)
    If mDotInfo.TargetRef.RefType = eUser And mDebuffIcon > 0 Then
        Call WriteSendSkillCdUpdate(mDotInfo.TargetRef.ArrayIndex, mDebuffIcon, _
                                    mDotInfo.UniqueId, 0, 0, EffectOverTime(mDotInfo.EotId).buffType)
    End If
    Exit Sub
IBaseEffectOverTime_OnRemove_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_OnRemove", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
    On Error Goto IBaseEffectOverTime_TargetUseMagic_Err
    Exit Sub
IBaseEffectOverTime_TargetUseMagic_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_TargetUseMagic", Erl)
End Sub

Public Sub IBaseEffectOverTime_TartgetWillAtack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TartgetWillAtack_Err
    Exit Sub
IBaseEffectOverTime_TartgetWillAtack_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_TartgetWillAtack", Erl)
End Sub

Public Function IBaseEffectOverTime_ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
    On Error Goto IBaseEffectOverTime_ChangeTarget_Err
    IBaseEffectOverTime_ChangeTarget = False
    Exit Function
IBaseEffectOverTime_ChangeTarget_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_ChangeTarget", Erl)
End Function

Public Sub IBaseEffectOverTime_TartgetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TartgetDidHit_Err
    Exit Sub
IBaseEffectOverTime_TartgetDidHit_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_TartgetDidHit", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TargetFailedAttack_Err
    Exit Sub
IBaseEffectOverTime_TargetFailedAttack_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_TargetFailedAttack", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TargetWasDamaged_Err
    Exit Sub
IBaseEffectOverTime_TargetWasDamaged_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_TargetWasDamaged", Erl)
End Sub

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType) As Long
    On Error Goto IBaseEffectOverTime_ApplyDamageReduction_Err
    IBaseEffectOverTime_ApplyDamageReduction = Damage
    Exit Function
IBaseEffectOverTime_ApplyDamageReduction_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_ApplyDamageReduction", Erl)
End Function

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
    On Error Goto IBaseEffectOverTime_TargetChangeTerrain_Err
    Exit Sub
IBaseEffectOverTime_TargetChangeTerrain_Err:
    Call TraceError(Err.Number, Err.Description, "UnequipItem.IBaseEffectOverTime_TargetChangeTerrain", Erl)
End Sub




