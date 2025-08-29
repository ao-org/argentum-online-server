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
Private PhysicalDamageModfier As Single
Private MagicDamageModifier As Single
Private ReductionActive As Boolean


Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
    On Error Goto Setup_Err
                 ByVal EffectId As Integer, ByVal id As Long)
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, ID, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    TimeSinceLastTick = 0
    ReductionActive = False
    With EffectOverTime(EffectId)
        TickTime = .TickTime
        PhysicalDamageModfier = .PhysicalDamageDone
        MagicDamageModifier = .MagicDamageDone
        If TargetType = eUser Then
            If .ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(TargetIndex, .ClientEffectTypeId, ID, .TickTime, .TickTime, .buffType)
        End If
    End With
    If TargetType = eNpc Then
        If SourceType = eUser Then
            Call SetUserRef(NpcList(TargetIndex).TargetUser, SourceIndex)
            NpcList(TargetIndex).flags.AttackedBy = UserList(SourceIndex).name
            NpcList(TargetIndex).flags.AttackedTime = GlobalFrameTime
        Else
            Call SetNpcRef(NpcList(TargetIndex).TargetNPC, SourceIndex)
        End If
        Call SetMask(NpcList(TargetIndex).flags.StatusMask, eTaunted)
        
    End If
    Exit Sub
Setup_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.Setup", Erl)
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eProvoked
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
    On Error Goto IBaseEffectOverTime_Update_Err
    TimeSinceLastTick = TimeSinceLastTick + deltaTime
    If TimeSinceLastTick >= TickTime Or Not IsValidRef(DotInfo.TargetRef) Or Not IsValidRef(Source) Then
        DotInfo.RemoveEffect = True
    End If
    Exit Sub
IBaseEffectOverTime_Update_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_Update", Erl)
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    On Error Goto IBaseEffectOverTime_Reset_Err
    IBaseEffectOverTime_Reset = True
    Exit Function
IBaseEffectOverTime_Reset_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_Reset", Erl)
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    On Error Goto IBaseEffectOverTime_OnRemove_Err
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    If DotInfo.TargetRef.RefType = eNpc Then
        Call UnsetMask(NpcList(DotInfo.TargetRef.ArrayIndex).flags.StatusMask, eTaunted)
    End If
    Exit Sub
IBaseEffectOverTime_OnRemove_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_OnRemove", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
    On Error Goto IBaseEffectOverTime_TargetUseMagic_Err
    Exit Sub
IBaseEffectOverTime_TargetUseMagic_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_TargetUseMagic", Erl)
End Sub

Public Sub IBaseEffectOverTime_TartgetWillAtack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TartgetWillAtack_Err
    Call ApplyReduction(TargetUserId)
    Exit Sub
IBaseEffectOverTime_TartgetWillAtack_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_TartgetWillAtack", Erl)
End Sub

Public Sub IBaseEffectOverTime_TartgetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TartgetDidHit_Err
    Call RestoreDamage(TargetUserId)
    Exit Sub
IBaseEffectOverTime_TartgetDidHit_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_TartgetDidHit", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TargetFailedAttack_Err
    Call RestoreDamage(TargetUserId)
    Exit Sub
IBaseEffectOverTime_TargetFailedAttack_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_TargetFailedAttack", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TargetWasDamaged_Err
    Exit Sub
IBaseEffectOverTime_TargetWasDamaged_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_TargetWasDamaged", Erl)
End Sub

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType) As Long
    On Error Goto IBaseEffectOverTime_ApplyDamageReduction_Err
    IBaseEffectOverTime_ApplyDamageReduction = Damage
    Exit Function
IBaseEffectOverTime_ApplyDamageReduction_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_ApplyDamageReduction", Erl)
End Function

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
    On Error Goto IBaseEffectOverTime_TargetWillAttackPosition_Err
    Exit Sub
IBaseEffectOverTime_TargetWillAttackPosition_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_TargetWillAttackPosition", Erl)
End Sub

Public Function IBaseEffectOverTime_ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
    On Error Goto IBaseEffectOverTime_ChangeTarget_Err
    IBaseEffectOverTime_ChangeTarget = False
    Exit Function
IBaseEffectOverTime_ChangeTarget_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_ChangeTarget", Erl)
End Function

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = e_EffectCallbackMask.eTartgetWillAtack Or e_EffectCallbackMask.eTartgetDidHit Or e_EffectCallbackMask.eTargetFailedAttack
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Private Sub ApplyReduction(ByVal TargetUserId As Integer)
    On Error Goto ApplyReduction_Err
    If TargetUserId <> Source.ArrayIndex And DotInfo.TargetRef.RefType = eUser Then
        Call IncreaseSingle(UserList(DotInfo.TargetRef.ArrayIndex).Modifiers.MagicDamageBonus, MagicDamageModifier)
        Call IncreaseSingle(UserList(DotInfo.TargetRef.ArrayIndex).Modifiers.PhysicalDamageBonus, PhysicalDamageModfier)
        ReductionActive = True
    End If
    Exit Sub
ApplyReduction_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.ApplyReduction", Erl)
End Sub

Private Sub RestoreDamage(ByVal TargetUserId As Integer)
    On Error Goto RestoreDamage_Err
    If ReductionActive And DotInfo.TargetRef.RefType = eUser Then
        Call IncreaseSingle(UserList(DotInfo.TargetRef.ArrayIndex).Modifiers.MagicDamageBonus, -MagicDamageModifier)
        Call IncreaseSingle(UserList(DotInfo.TargetRef.ArrayIndex).Modifiers.PhysicalDamageBonus, -PhysicalDamageModfier)
        ReductionActive = False
    End If
    Exit Sub
RestoreDamage_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.RestoreDamage", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
    On Error Goto IBaseEffectOverTime_TargetChangeTerrain_Err
    Exit Sub
IBaseEffectOverTime_TargetChangeTerrain_Err:
    Call TraceError(Err.Number, Err.Description, "EffectProvoked.IBaseEffectOverTime_TargetChangeTerrain", Erl)
End Sub
