VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatModifier"
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

Dim TickTime As Long
Dim TimeSinceLastTick As Long
Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference
Dim Modifier As Single
Private RequireTransform As Integer

Public Sub Setup(ByVal sourceIndex As Integer, ByVal sourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, ByVal id As Long)
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, ID, EffectId)
    Call SetRef(Source, sourceIndex, sourceType)
    Debug.Print "setting new id: " & ID & " to: " & DotInfo.UniqueId & " on target: " & TargetIndex
    TimeSinceLastTick = 0
    With EffectOverTime(EffectId)
        TickTime = .TickTime
        RequireTransform = .RequireTransform
        If IsSet(.EffectModifiers, e_ModifierTypes.MagicBonus) And SourceType = eUser Then
            If UserList(SourceIndex).invent.WeaponEqpObjIndex > 0 Then
                Modifier = ObjData(UserList(SourceIndex).invent.WeaponEqpObjIndex).MagicDamageBonus / 100
            End If
        End If
        If DotInfo.TargetRef.RefType = eUser And .ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(TargetIndex, .ClientEffectTypeId, ID, .TickTime, .TickTime, .buffType)
    End With
    Call EffectsOverTime.ApplyEotModifier(DotInfo.TargetRef, EffectOverTime(EffectId), Modifier)
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eApplyModifiers
End Property

Public Property Get IBaseEffectOverTime_SharedTypeId() As Integer
    IBaseEffectOverTime_SharedTypeId = EffectOverTime(DotInfo.EotId).SharedTypeId
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Long
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
    If TimeSinceLastTick >= TickTime Or Not IsValidRef(DotInfo.TargetRef) Then
        DotInfo.RemoveEffect = True
        Exit Sub
    End If
    If DotInfo.TargetRef.RefType = eUser Then
        If TestRequiredEquipedItem(UserList(DotInfo.TargetRef.ArrayIndex).invent, EffectOverTime(DotInfo.EotId).SpellRequirementMask, EffectOverTime(DotInfo.EotId).RequireWeaponType) <> e_SpellRequirementMask.eNone Then
            DotInfo.RemoveEffect = True
            Exit Sub
        End If
        If RequireTransform > 0 Then
            If RequireTransform <> UserList(DotInfo.TargetRef.ArrayIndex).flags.ActiveTransform Then
                DotInfo.RemoveEffect = True
            End If
        End If
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    TimeSinceLastTick = 0
    IBaseEffectOverTime_Reset = True
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, TickTime, TickTime, EffectOverTime(DotInfo.EotId).buffType)
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    Call EffectsOverTime.RemoveEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId), Modifier)
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, 0, EffectOverTime(DotInfo.EotId).buffType)
End Sub

Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
End Sub

Public Sub IBaseEffectOverTime_TartgetWillAtack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TartgetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
End Sub

Public Function IBaseEffectOverTime_ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
    If DotInfo.TargetRef.RefType = eUser Then
        If EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, DotInfo.EotId, DotInfo.UniqueId, 0, 0, EffectOverTime(DotInfo.EotId).buffType)
    End If
    Call EffectsOverTime.RemoveEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId), Modifier)
    Call SetRef(DotInfo.TargetRef, NewTargetIndex, NewTargetType)
    If DotInfo.TargetRef.RefType = eUser Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, TickTime - TimeSinceLastTick, TickTime, EffectOverTime(DotInfo.EotId).buffType)
    End If
    Call EffectsOverTime.ApplyEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId), Modifier)
    IBaseEffectOverTime_ChangeTarget = True
End Function

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = e_EffectCallbackMask.eTargetWasDamaged
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    If EffectOverTime(DotInfo.EotId).OnHitFx > 0 Then
        If IsVisible(UserList(DotInfo.TargetRef.ArrayIndex)) Then
            Call SendData(SendTarget.ToPCAliveArea, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, EffectOverTime(DotInfo.EotId).OnHitFx, 0, UserList(DotInfo.TargetRef.ArrayIndex).pos.X, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
        Else
            Call SendData(SendTarget.ToIndex, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, EffectOverTime(DotInfo.EotId).OnHitFx, 0, UserList(DotInfo.TargetRef.ArrayIndex).pos.X, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
        End If
    End If
End Sub

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType) As Long
    IBaseEffectOverTime_ApplyDamageReduction = Damage
End Function

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub
