VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransformEffect"
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


Dim DotHpMin As Integer
Dim DotHpMax As Integer
Dim TargetTicks As Integer
Dim TickCount As Integer
Dim TickManaConsumption As Integer
Dim TickStaminaConsumption As Integer
Dim TickTime As Long
Dim TimeSinceLastTick As Long
Dim TickEffect As Integer
Dim Modifiers As Integer
Dim EffectBonus As Single
Dim npcId As Integer
Dim StatusMask As Long

Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
    On Error Goto Setup_Err
                 ByVal EffectId As Integer, ByVal ID As Long)
    TimeSinceLastTick = 0
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, ID, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    Call SetEffect(EffectId)
    Exit Sub
Setup_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.Setup", Erl)
End Sub

Public Sub SetEffect(ByVal EffectId As Integer)
    On Error Goto SetEffect_Err
    DotInfo.EotId = EffectId
    TickCount = 0
    With EffectOverTime(EffectId)
        TickTime = .TickTime
        TickEffect = .TickFX
        TargetTicks = .Ticks
        TickManaConsumption = .TickManaConsumption
        TickStaminaConsumption = .TickStaminaConsumption
        npcId = .npcId
        StatusMask = .ApplyStatusMask
        If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, .ClientEffectTypeId, DotInfo.UniqueId, .TickTime * TargetTicks, .TickTime * TargetTicks, .buffType)
        End If
    End With
    Call TransformChacter
    Exit Sub
SetEffect_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.SetEffect", Erl)
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eTransform
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
    IBaseEffectOverTime_EffectType = EffectOverTime(DotInfo.EotId).buffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    On Error Goto IBaseEffectOverTime_Update_Err
    If TickCount >= TargetTicks Or Not IsValidRef(DotInfo.TargetRef) Then
        DotInfo.RemoveEffect = True
        Exit Sub
    End If
    TimeSinceLastTick = TimeSinceLastTick + deltaTime
    If TimeSinceLastTick >= TickTime Then
        TimeSinceLastTick = 0
        TickCount = TickCount + 1
        If DotInfo.TargetRef.RefType = eUser Then
            If UserMod.ModifyMana(DotInfo.TargetRef.ArrayIndex, -TickManaConsumption, True) Then
                DotInfo.RemoveEffect = True
            End If
        End If
        If DotInfo.TargetRef.RefType = eUser Then
            If UserMod.ModifyStamina(DotInfo.TargetRef.ArrayIndex, -TickStaminaConsumption, True) Then
                DotInfo.RemoveEffect = True
            End If
        End If
    End If
    Exit Sub
IBaseEffectOverTime_Update_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_Update", Erl)
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal Value As Boolean)
    DotInfo.RemoveEffect = Value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    On Error Goto IBaseEffectOverTime_Reset_Err
    If (NewEffectId = DotInfo.EotId) Then
        IBaseEffectOverTime_Reset = False
        DotInfo.RemoveEffect = True
        Exit Function
    End If
    TimeSinceLastTick = 0
    TickCount = 0
    IBaseEffectOverTime_Reset = True
    Call ReturnCharToNormal
    Call SetEffect(NewEffectId)
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, _
                                    DotInfo.UniqueId, (TickTime * TargetTicks) - TimeSinceLastTick, TickTime * TargetTicks, EffectOverTime(DotInfo.EotId).buffType)
    End If
    Exit Function
IBaseEffectOverTime_Reset_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_Reset", Erl)
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    On Error Goto IBaseEffectOverTime_OnRemove_Err
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    Call ReturnCharToNormal
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, _
                                    0, 0, EffectOverTime(DotInfo.EotId).buffType)
    End If
    Exit Sub
IBaseEffectOverTime_OnRemove_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_OnRemove", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
    On Error Goto IBaseEffectOverTime_TargetUseMagic_Err
    Exit Sub
IBaseEffectOverTime_TargetUseMagic_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_TargetUseMagic", Erl)
End Sub

Public Sub IBaseEffectOverTime_TartgetWillAtack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TartgetWillAtack_Err
    Exit Sub
IBaseEffectOverTime_TartgetWillAtack_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_TartgetWillAtack", Erl)
End Sub

Public Sub IBaseEffectOverTime_TartgetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TartgetDidHit_Err
    Exit Sub
IBaseEffectOverTime_TartgetDidHit_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_TartgetDidHit", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TargetFailedAttack_Err
    Exit Sub
IBaseEffectOverTime_TargetFailedAttack_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_TargetFailedAttack", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TargetWasDamaged_Err
    Exit Sub
IBaseEffectOverTime_TargetWasDamaged_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_TargetWasDamaged", Erl)
End Sub

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType) As Long
    On Error Goto IBaseEffectOverTime_ApplyDamageReduction_Err
    IBaseEffectOverTime_ApplyDamageReduction = Damage
    Exit Function
IBaseEffectOverTime_ApplyDamageReduction_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_ApplyDamageReduction", Erl)
End Function

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
    On Error Goto IBaseEffectOverTime_TargetWillAttackPosition_Err
    Exit Sub
IBaseEffectOverTime_TargetWillAttackPosition_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_TargetWillAttackPosition", Erl)
End Sub

Public Function IBaseEffectOverTime_ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
    On Error Goto IBaseEffectOverTime_ChangeTarget_Err
    IBaseEffectOverTime_ChangeTarget = False
    Exit Function
IBaseEffectOverTime_ChangeTarget_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_ChangeTarget", Erl)
End Function

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = eTargetChangeTerrain
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
    On Error Goto IBaseEffectOverTime_TargetChangeTerrain_Err
    If DotInfo.TargetRef.RefType <> eUser Then Exit Sub
    With UserList(DotInfo.TargetRef.ArrayIndex)
        If .flags.Navegando > 0 Or .flags.Montado > 0 Or .flags.Nadando > 0 Then
            DotInfo.RemoveEffect = True
        End If
    End With
    Exit Sub
IBaseEffectOverTime_TargetChangeTerrain_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.IBaseEffectOverTime_TargetChangeTerrain", Erl)
End Sub

Private Sub TransformChacter()
    On Error Goto TransformChacter_Err
    Call EffectsOverTime.ApplyEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId))
    If (DotInfo.TargetRef.RefType = eUser) Then
        With UserList(DotInfo.TargetRef.ArrayIndex)
            Dim NpcIndex As String
            NpcIndex = "NPC" & npcId
136         .Char.body = val(LeerNPCs.GetValue(NpcIndex, "Body"))
138         .Char.head = val(LeerNPCs.GetValue(NpcIndex, "Head"))
142         .Char.BodyIdle = val(LeerNPCs.GetValue(NpcIndex, "BodyIdle"))
143         .Char.Ataque1 = val(LeerNPCs.GetValue(NpcIndex, "Ataque1"))
156         .Char.WeaponAnim = val(LeerNPCs.GetValue(NpcIndex, "Arma"))
158         .Char.ShieldAnim = val(LeerNPCs.GetValue(NpcIndex, "Escudo"))
160         .Char.CascoAnim = val(LeerNPCs.GetValue(NpcIndex, "Casco"))
161         .Char.CartAnim = val(LeerNPCs.GetValue(NpcIndex, "Cart"))
            .Char.CastAnimation = val(LeerNPCs.GetValue(NpcIndex, "CastAnimation"))
            .flags.ActiveTransform = NpcId
            Dim CantidadAnimaciones As Integer
144         CantidadAnimaciones = val(LeerNPCs.GetValue(NpcIndex, "Animaciones"))
128         Call ChangeUserChar(DotInfo.TargetRef.ArrayIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
            Call SetMask(.flags.StatusMask, StatusMask)
130         Call RefreshCharStatus(DotInfo.TargetRef.ArrayIndex)
        End With
    End If
    Exit Sub
TransformChacter_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.TransformChacter", Erl)
End Sub

Private Sub ReturnCharToNormal()
    On Error Goto ReturnCharToNormal_Err
    Call EffectsOverTime.RemoveEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId))
    If DotInfo.TargetRef.RefType = eUser Then
        UserList(DotInfo.TargetRef.ArrayIndex).flags.ActiveTransform = 0
        Call UnsetMask(UserList(DotInfo.TargetRef.ArrayIndex).flags.StatusMask, StatusMask)
        Call UpdateCharWithEquipedItems(DotInfo.TargetRef.ArrayIndex)
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, 0, EffectOverTime(DotInfo.EotId).buffType)
    End If
    Exit Sub
ReturnCharToNormal_Err:
    Call TraceError(Err.Number, Err.Description, "TransformEffect.ReturnCharToNormal", Erl)
End Sub


