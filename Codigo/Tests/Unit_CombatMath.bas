Attribute VB_Name = "Unit_CombatMath"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_combatmath() As Boolean
    Call UnitTesting.RunTest("test_minimoint_smaller", test_minimoint_smaller())
    Call UnitTesting.RunTest("test_maximoint_larger", test_maximoint_larger())
    Call UnitTesting.RunTest("test_minimoint_equal", test_minimoint_equal())
    Call UnitTesting.RunTest("test_maximoint_equal", test_maximoint_equal())
    Call UnitTesting.RunTest("test_minimoint_negative", test_minimoint_negative())
    Call UnitTesting.RunTest("test_maximoint_negative", test_maximoint_negative())
    Call UnitTesting.RunTest("test_minmax_int_property", test_minmax_int_property())
    Call UnitTesting.RunTest("EOT PvE physical bonuses on users", test_eot_pve_physical_bonus(eUser))
    Call UnitTesting.RunTest("EOT PvE physical bonuses on NPCs", test_eot_pve_physical_bonus(eNpc))
    test_suite_combatmath = True
End Function

' Verify MinimoInt returns the smaller of two different values.
Private Function test_minimoint_smaller() As Boolean
    On Error GoTo Err_Handler
    test_minimoint_smaller = True

    If MinimoInt(3, 7) <> 3 Then test_minimoint_smaller = False: Exit Function
    If MinimoInt(10, 2) <> 2 Then test_minimoint_smaller = False: Exit Function

    Exit Function
Err_Handler:
    test_minimoint_smaller = False
End Function

' Verify MaximoInt returns the larger of two different values.
Private Function test_maximoint_larger() As Boolean
    On Error GoTo Err_Handler
    test_maximoint_larger = True

    If MaximoInt(3, 7) <> 7 Then test_maximoint_larger = False: Exit Function
    If MaximoInt(10, 2) <> 10 Then test_maximoint_larger = False: Exit Function

    Exit Function
Err_Handler:
    test_maximoint_larger = False
End Function

' Verify MinimoInt with equal values returns that same value.
Private Function test_minimoint_equal() As Boolean
    On Error GoTo Err_Handler
    test_minimoint_equal = True

    If MinimoInt(5, 5) <> 5 Then test_minimoint_equal = False: Exit Function

    Exit Function
Err_Handler:
    test_minimoint_equal = False
End Function

' Verify MaximoInt with equal values returns that same value.
Private Function test_maximoint_equal() As Boolean
    On Error GoTo Err_Handler
    test_maximoint_equal = True

    If MaximoInt(5, 5) <> 5 Then test_maximoint_equal = False: Exit Function

    Exit Function
Err_Handler:
    test_maximoint_equal = False
End Function

' Verify MinimoInt with negative values returns the correct minimum.
Private Function test_minimoint_negative() As Boolean
    On Error GoTo Err_Handler
    test_minimoint_negative = True

    If MinimoInt(-10, -3) <> -10 Then test_minimoint_negative = False: Exit Function
    If MinimoInt(-5, 5) <> -5 Then test_minimoint_negative = False: Exit Function

    Exit Function
Err_Handler:
    test_minimoint_negative = False
End Function

' Verify MaximoInt with negative values returns the correct maximum.
Private Function test_maximoint_negative() As Boolean
    On Error GoTo Err_Handler
    test_maximoint_negative = True

    If MaximoInt(-10, -3) <> -3 Then test_maximoint_negative = False: Exit Function
    If MaximoInt(-5, 5) <> 5 Then test_maximoint_negative = False: Exit Function

    Exit Function
Err_Handler:
    test_maximoint_negative = False
End Function

' Property 1: MinimoInt and MaximoInt correctness
' For any two Integer values a and b, MinimoInt(a,b) <= a, MinimoInt(a,b) <= b,
' MaximoInt(a,b) >= a, MaximoInt(a,b) >= b, and one of {a, b} equals the min and max.
' Uses 200 randomized trials to approximate universal quantification.
Private Function test_minmax_int_property() As Boolean
    On Error GoTo Err_Handler
    test_minmax_int_property = True
    
    Dim i As Long
    Dim a As Integer
    Dim b As Integer
    Dim minResult As Integer
    Dim maxResult As Integer
    
    For i = 1 To 200
        ' Generate random Integer values across full Integer range (-32768 to 32767)
        a = CInt(Int(Rnd * 65536) - 32768)
        b = CInt(Int(Rnd * 65536) - 32768)
        
        minResult = MinimoInt(a, b)
        maxResult = MaximoInt(a, b)
        
        ' MinimoInt must be <= both inputs
        If minResult > a Then test_minmax_int_property = False: Exit Function
        If minResult > b Then test_minmax_int_property = False: Exit Function
        
        ' MaximoInt must be >= both inputs
        If maxResult < a Then test_minmax_int_property = False: Exit Function
        If maxResult < b Then test_minmax_int_property = False: Exit Function
        
        ' Min must equal one of {a, b}
        If minResult <> a And minResult <> b Then test_minmax_int_property = False: Exit Function
        
        ' Max must equal one of {a, b}
        If maxResult <> a And maxResult <> b Then test_minmax_int_property = False: Exit Function
    Next i
    Exit Function
Err_Handler:
    test_minmax_int_property = False
End Function

' Exercise the same apply/remove path used by expiry, replacement and dispels.
Private Function test_eot_pve_physical_bonus(ByVal ownerType As e_ReferenceType) As Boolean
    On Error GoTo Err_Handler
    Dim target As t_AnyReference
    Dim savedModifiers As t_ActiveModifiers
    Dim emptyModifiers As t_ActiveModifiers
    Dim normalEffect As t_EffectOverTime
    Dim pveEffect As t_EffectOverTime
    Dim bonus As Variant
    Dim stateSaved As Boolean
    Dim savedSpeed As Single
    Dim savedMoveInterval As Long

    Call SetRef(target, 1, ownerType)
    If ownerType = eUser Then
        savedModifiers = UserList(1).Modifiers
        savedSpeed = UserList(1).Char.speeding
        UserList(1).Modifiers = emptyModifiers
    Else
        savedModifiers = NpcList(1).Modifiers
        savedSpeed = NpcList(1).Char.speeding
        savedMoveInterval = NpcList(1).IntervaloMovimiento
        NpcList(1).Modifiers = emptyModifiers
    End If
    stateSaved = True

    normalEffect.PhysicalLinearBonus = 20
    normalEffect.PhysicalDamageDone = 0.25
    Call EffectsOverTime.ApplyEotModifier(target, normalEffect)
    If Not check_eot_physical_bonus(ownerType, 20, 20, 1.25, 1.25) Then Err.Raise vbObjectError + 1, , "Unexpected physical bonus"

    pveEffect.PhysicalBonusPveOnly = True
    pveEffect.PhysicalDamageDone = 0.5
    For Each bonus In Array(15, 25, 40)
        pveEffect.PhysicalLinearBonus = CInt(bonus)
        Call EffectsOverTime.ApplyEotModifier(target, pveEffect)
        If Not check_eot_physical_bonus(ownerType, 20, 20 + CInt(bonus), 1.25, 1.75) Then Err.Raise vbObjectError + 1, , "Unexpected physical bonus"
        Call EffectsOverTime.RemoveEotModifier(target, pveEffect)
        If Not check_eot_physical_bonus(ownerType, 20, 20, 1.25, 1.25) Then Err.Raise vbObjectError + 1, , "Unexpected physical bonus"
    Next bonus

    ' Scaling and stacking stay independent of unrestricted bonuses.
    pveEffect.PhysicalLinearBonus = 15
    Call EffectsOverTime.ApplyEotModifier(target, pveEffect, 0.5)
    Call EffectsOverTime.ApplyEotModifier(target, pveEffect)
    If Not check_eot_physical_bonus(ownerType, 20, 57, 1.25, 2.5) Then Err.Raise vbObjectError + 1, , "Unexpected physical bonus"
    Call EffectsOverTime.RemoveEotModifier(target, normalEffect)
    If Not check_eot_physical_bonus(ownerType, 0, 37, 1, 2.25) Then Err.Raise vbObjectError + 1, , "Unexpected physical bonus"
    Call EffectsOverTime.RemoveEotModifier(target, pveEffect, 0.5)
    Call EffectsOverTime.RemoveEotModifier(target, pveEffect)
    If Not check_eot_physical_bonus(ownerType, 0, 0, 1, 1) Then Err.Raise vbObjectError + 1, , "Unexpected physical bonus"

    ' Reused entity slots must not retain restricted bonuses.
    Call EffectsOverTime.ApplyEotModifier(target, pveEffect)
    If ownerType = eUser Then
        Call ClearModifiers(UserList(1).Modifiers)
    Else
        Call ClearModifiers(NpcList(1).Modifiers)
    End If
    If Not check_eot_physical_bonus(ownerType, 0, 0, 1, 1) Then Err.Raise vbObjectError + 1, , "Unexpected physical bonus"
    test_eot_pve_physical_bonus = True

Cleanup:
    If stateSaved Then
        If ownerType = eUser Then
            UserList(1).Modifiers = savedModifiers
            UserList(1).Char.speeding = savedSpeed
        Else
            NpcList(1).Modifiers = savedModifiers
            NpcList(1).Char.speeding = savedSpeed
            NpcList(1).IntervaloMovimiento = savedMoveInterval
        End If
    End If
    Exit Function
Err_Handler:
    test_eot_pve_physical_bonus = False
    Resume Cleanup
End Function

Private Function check_eot_physical_bonus(ByVal ownerType As e_ReferenceType, ByVal userFlat As Integer, ByVal npcFlat As Integer, _
                                         ByVal userMultiplier As Single, ByVal npcMultiplier As Single) As Boolean
    If ownerType = eUser Then
        If UserMod.GetLinearDamageBonus(1, eUser) <> userFlat Then Exit Function
        If UserMod.GetLinearDamageBonus(1, eNpc) <> npcFlat Then Exit Function
        If UserMod.GetLinearDamageBonus(1, e_ReferenceType.eNone) <> userFlat Then Exit Function
        If UserMod.GetPhysicalDamageModifier(UserList(1), eUser) <> userMultiplier Then Exit Function
        If UserMod.GetPhysicalDamageModifier(UserList(1), eNpc) <> npcMultiplier Then Exit Function
    Else
        If NPCs.GetLinearDamageBonus(1, eUser) <> userFlat Then Exit Function
        If NPCs.GetLinearDamageBonus(1, eNpc) <> npcFlat Then Exit Function
        If NPCs.GetLinearDamageBonus(1, e_ReferenceType.eNone) <> userFlat Then Exit Function
        If NPCs.GetPhysicalDamageModifier(NpcList(1), eUser) <> userMultiplier Then Exit Function
        If NPCs.GetPhysicalDamageModifier(NpcList(1), eNpc) <> npcMultiplier Then Exit Function
    End If
    check_eot_physical_bonus = True
End Function

#End If
