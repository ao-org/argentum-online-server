Attribute VB_Name = "ModReferenceUtils"
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
        UserCanAttack = UserCanAttackNpc(UserIndex, Reference.ArrayIndex)
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
