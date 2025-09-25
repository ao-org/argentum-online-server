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
Dim TimeSinceLastTick As Long
Dim DamageProtection  As Long
Dim TickTime          As Long
Private DotInfo       As t_BaseDotInfo
Private Source        As t_AnyReference

Public Sub Setup(ByVal SourceIndex As Integer, _
                 ByVal SourceType As e_ReferenceType, _
                 ByVal TargetIndex As Integer, _
                 ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, _
                 ByVal Id As Long)
    TimeSinceLastTick = 0
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, Id, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    With EffectOverTime(EffectId)
        DamageProtection = RandomNumber(.TickPowerMin, .TickPowerMax)
        TickTime = .TickTime
        If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(TargetIndex, .ClientEffectTypeId, Id, .TickTime, .TickTime, .buffType)
        End If
    End With
    Call AddShieldToReference(DotInfo.TargetRef, DamageProtection)
    Call EffectsOverTime.ApplyEotModifier(DotInfo.TargetRef, EffectOverTime(EffectId), 0)
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eProtection
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
    IBaseEffectOverTime_CasterUserId = Source.UserId
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
    IBaseEffectOverTime_EffectType = EffectOverTime(DotInfo.EotId).buffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal DeltaTime As Long)
    TimeSinceLastTick = TimeSinceLastTick + DeltaTime
    If TimeSinceLastTick >= TickTime Or Not IsValidRef(DotInfo.TargetRef) Then
        DotInfo.RemoveEffect = True
        Exit Sub
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    IBaseEffectOverTime_Reset = True
    TimeSinceLastTick = 0
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, TickTime, TickTime, EffectOverTime( _
                DotInfo.EotId).buffType)
    End If
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    Call EffectsOverTime.RemoveEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId), 0)
    Call AddShieldToReference(DotInfo.TargetRef, -DamageProtection)
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, 0, EffectOverTime(DotInfo.EotId).buffType)
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

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, _
                                                         ByVal SourceUserId As Integer, _
                                                         ByVal SourceType As e_ReferenceType, _
                                                         ByVal AttackType As e_DamageSourceType) As Long
    If EffectOverTime(DotInfo.EotId).OnHitFx > 0 Then
        If IsVisible(UserList(DotInfo.TargetRef.ArrayIndex)) Then
            Call SendData(SendTarget.ToPCAliveArea, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, EffectOverTime( _
                    DotInfo.EotId).OnHitFx, 0, UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
        Else
            Call SendData(SendTarget.ToIndex, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, EffectOverTime( _
                    DotInfo.EotId).OnHitFx, 0, UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
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

Public Function IBaseEffectOverTime_ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
    If DotInfo.TargetRef.RefType = eUser Then
        If EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, DotInfo.EotId, DotInfo.UniqueId, 0, 0, EffectOverTime(DotInfo.EotId).buffType)
        End If
    End If
    Call EffectsOverTime.RemoveEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId), 0)
    Call AddShieldToReference(DotInfo.TargetRef, -DamageProtection)
    Call SetRef(DotInfo.TargetRef, NewTargetIndex, NewTargetType)
    If DotInfo.TargetRef.RefType = eUser Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, TickTime - TimeSinceLastTick, TickTime, _
                EffectOverTime(DotInfo.EotId).buffType)
    End If
    Call AddShieldToReference(DotInfo.TargetRef, DamageProtection)
    Call EffectsOverTime.ApplyEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId), 0)
    IBaseEffectOverTime_ChangeTarget = True
End Function

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = e_EffectCallbackMask.eTargetApplyDamageReduction
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub
