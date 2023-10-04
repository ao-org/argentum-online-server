Attribute VB_Name = "ModReferenceUtils"
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

Public Function GetPosition(ByRef Reference As t_AnyReference) As t_WorldPos
    If Not IsValidRef(Reference) Then
        Exit Function
    End If
    If Reference.RefType = eNpc Then
        GetPosition = NpcList(Reference.ArrayIndex).pos
    ElseIf Reference.RefType = eUser Then
        GetPosition = UserList(Reference.ArrayIndex).pos
    End If
End Function

Public Sub SetTranslationState(ByRef Reference As t_AnyReference, ByVal NewState As Boolean)
    If Not IsValidRef(Reference) Then
        Exit Sub
    End If
    If Reference.RefType = eNpc Then
        NpcList(Reference.ArrayIndex).flags.TranslationActive = NewState
    ElseIf Reference.RefType = eUser Then
        UserList(Reference.ArrayIndex).flags.TranslationActive = NewState
    End If
End Sub

Public Function UserCanAttack(ByVal UserIndex As Integer, ByVal UserVersionId, ByRef Reference As t_AnyReference) As e_AttackInteractionResult
    If Reference.RefType = eUser Then
        UserCanAttack = UserMod.CanAttackUser(UserIndex, UserVersionId, Reference.ArrayIndex, Reference.VersionId)
    Else
        UserCanAttack = UserCanAttackNpc(UserIndex, Reference.ArrayIndex).CanAttack
    End If
End Function

Public Function NpcCanAttack(ByVal NpcIndex As Integer, ByRef Reference As t_AnyReference) As e_AttackInteractionResult
    If Reference.RefType = eUser Then
        NpcCanAttack = NPCs.CanAttackUser(NpcIndex, Reference.ArrayIndex)
    Else
        NpcCanAttack = NPCs.CanAttackNpc(NpcIndex, Reference.ArrayIndex)
    End If
End Function

Public Sub UpdateIncreaseModifier(ByRef Reference As t_AnyReference, ByVal Modifier As e_ModifierTypes, ByVal Value As Single)
    If Reference.RefType = eUser Then
        Select Case Modifier
            Case e_ModifierTypes.MagicBonus
                Call IncreaseSingle(UserList(Reference.ArrayIndex).Modifiers.MagicDamageBonus, Value)
            Case e_ModifierTypes.MagicReduction
                Call IncreaseSingle(UserList(Reference.ArrayIndex).Modifiers.MagicDamageReduction, Value)
            Case e_ModifierTypes.MovementSpeed
                Call IncreaseSingle(UserList(Reference.ArrayIndex).Modifiers.MovementSpeed, Value)
                Call ActualizarVelocidadDeUsuario(Reference.ArrayIndex)
            Case e_ModifierTypes.PhysicalReduction
                Call IncreaseSingle(UserList(Reference.ArrayIndex).Modifiers.PhysicalDamageReduction, Value)
            Case e_ModifierTypes.PhysiccalBonus
                Call IncreaseSingle(UserList(Reference.ArrayIndex).Modifiers.PhysicalDamageBonus, Value)
            Case e_ModifierTypes.HitBonus
                Call IncreaseInteger(UserList(Reference.ArrayIndex).Modifiers.HitBonus, Value)
            Case e_ModifierTypes.EvasionBonus
                Call IncreaseInteger(UserList(Reference.ArrayIndex).Modifiers.EvasionBonus, Value)
            Case e_ModifierTypes.SelfHealingBonus
                Call IncreaseSingle(UserList(Reference.ArrayIndex).Modifiers.SelfHealingBonus, Value)
            Case e_ModifierTypes.MagicHealingBonus
                Call IncreaseSingle(UserList(Reference.ArrayIndex).Modifiers.MagicHealingBonus, Value)
            Case e_ModifierTypes.PhysicalLinearBonus
                Call IncreaseInteger(UserList(Reference.ArrayIndex).Modifiers.PhysicalDamageLinearBonus, Value)
            Case e_ModifierTypes.DefenseBonus
                Call IncreaseInteger(UserList(Reference.ArrayIndex).Modifiers.DefenseBonus, Value)
        End Select
    Else
        Select Case Modifier
            Case e_ModifierTypes.MagicBonus
                Call IncreaseSingle(NpcList(Reference.ArrayIndex).Modifiers.MagicDamageBonus, Value)
            Case e_ModifierTypes.MagicReduction
                Call IncreaseSingle(NpcList(Reference.ArrayIndex).Modifiers.MagicDamageReduction, Value)
            Case e_ModifierTypes.MovementSpeed
                Call IncreaseSingle(NpcList(Reference.ArrayIndex).Modifiers.MovementSpeed, Value)
                Call UpdateNpcSpeed(Reference.ArrayIndex)
            Case e_ModifierTypes.PhysicalReduction
                Call IncreaseSingle(NpcList(Reference.ArrayIndex).Modifiers.PhysicalDamageReduction, Value)
            Case e_ModifierTypes.PhysiccalBonus
                Call IncreaseSingle(NpcList(Reference.ArrayIndex).Modifiers.PhysicalDamageBonus, Value)
            Case e_ModifierTypes.HitBonus
                Call IncreaseInteger(NpcList(Reference.ArrayIndex).Modifiers.HitBonus, Value)
            Case e_ModifierTypes.EvasionBonus
                Call IncreaseInteger(NpcList(Reference.ArrayIndex).Modifiers.EvasionBonus, Value)
            Case e_ModifierTypes.SelfHealingBonus
                Call IncreaseSingle(NpcList(Reference.ArrayIndex).Modifiers.SelfHealingBonus, Value)
            Case e_ModifierTypes.MagicHealingBonus
                Call IncreaseSingle(NpcList(Reference.ArrayIndex).Modifiers.MagicHealingBonus, Value)
            Case e_ModifierTypes.PhysicalLinearBonus
                Call IncreaseInteger(NpcList(Reference.ArrayIndex).Modifiers.PhysicalDamageLinearBonus, Value)
            Case e_ModifierTypes.DefenseBonus
                Call IncreaseInteger(NpcList(Reference.ArrayIndex).Modifiers.DefenseBonus, Value)
        End Select
    End If
End Sub

Public Function DoDamageToTarget(ByVal UserIndex As Integer, ByRef TargetRef As t_AnyReference, ByVal Damage As Integer, _
                                 ByVal DamageType As e_DamageSourceType, ByVal ObjIndex As Integer) As e_DamageResult
    If Not IsValidRef(TargetRef) Then
        Exit Function
    End If
    If TargetRef.RefType = eNpc Then
        DoDamageToTarget = UserDamageToNpc(UserIndex, TargetRef.ArrayIndex, Damage, DamageType, ObjIndex)
    ElseIf TargetRef.RefType = eUser Then
        DoDamageToTarget = UserDoDamageToUser(UserIndex, TargetRef.ArrayIndex, Damage, DamageType, ObjIndex)
    End If
End Function

Public Function NpcDoDamageToTarget(ByVal NpcIndex As Integer, ByRef TargetRef As t_AnyReference, ByVal Damage As Integer, _
                                 ByVal DamageType As e_DamageSourceType, ByVal ObjIndex As Integer) As e_DamageResult
    If Not IsValidRef(TargetRef) Then
        Exit Function
    End If
    If TargetRef.RefType = eNpc Then
        NpcDoDamageToTarget = NpcDamageToNpc(NpcIndex, TargetRef.ArrayIndex, Damage)
    ElseIf TargetRef.RefType = eUser Then
        NpcDoDamageToTarget = NpcDoDamageToUser(NpcIndex, TargetRef.ArrayIndex, Damage, DamageType, ObjIndex)
    End If
End Function

Public Function RefDoDamageToTarget(ByRef SourceRef As t_AnyReference, ByRef TargetRef As t_AnyReference, ByVal Damage As Integer, _
                                 ByVal DamageType As e_DamageSourceType, ByVal ObjIndex As Integer) As e_DamageResult
    If Not IsValidRef(SourceRef) Then
        Exit Function
    End If
    If SourceRef.RefType = eNpc Then
        RefDoDamageToTarget = NpcDoDamageToTarget(SourceRef.ArrayIndex, TargetRef, Damage, DamageType, ObjIndex)
    ElseIf SourceRef.RefType = eUser Then
        RefDoDamageToTarget = DoDamageToTarget(SourceRef.ArrayIndex, TargetRef, Damage, DamageType, ObjIndex)
    End If
End Function

Public Function AddShieldToReference(ByRef SourceRef As t_AnyReference, ByVal ShieldSize As Long)
    If SourceRef.RefType = eUser Then
        Call IncreaseLong(UserList(SourceRef.ArrayIndex).Stats.Shield, ShieldSize)
        WriteUpdateHP (SourceRef.ArrayIndex)
    Else
        Call IncreaseLong(NpcList(SourceRef.ArrayIndex).Stats.Shield, ShieldSize)
        Call SendData(SendTarget.ToNPCAliveArea, SourceRef.ArrayIndex, PrepareMessageNpcUpdateHP(SourceRef.ArrayIndex))
    End If
End Function

Public Function GetName(ByRef SourceRef As t_AnyReference) As String
    If SourceRef.RefType = eUser Then
        GetName = UserList(SourceRef.ArrayIndex).name
    Else
        GetName = NpcList(SourceRef.ArrayIndex).name
    End If
End Function

Public Sub SetStatusMask(ByRef TargetRef As t_AnyReference, ByVal Mask As Long)
    If TargetRef.RefType = eUser Then
        Call SetMask(UserList(TargetRef.ArrayIndex).flags.StatusMask, Mask)
    Else
        Call SetMask(NpcList(TargetRef.ArrayIndex).flags.StatusMask, Mask)
    End If
End Sub

Public Sub UnsetStatusMask(ByRef TargetRef As t_AnyReference, ByVal Mask As Long)
    If TargetRef.RefType = eUser Then
        Call UnsetMask(UserList(TargetRef.ArrayIndex).flags.StatusMask, Mask)
    Else
        Call UnsetMask(NpcList(TargetRef.ArrayIndex).flags.StatusMask, Mask)
    End If
End Sub

Public Function IsDead(ByRef TargetRef As t_AnyReference)
    If TargetRef.RefType = eUser Then
        IsDead = UserList(TargetRef.ArrayIndex).flags.Muerto = 1
    Else
        IsDead = NpcList(TargetRef.ArrayIndex).Stats.MinHp = 0
    End If
End Function
