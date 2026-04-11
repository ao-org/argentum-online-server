Attribute VB_Name = "Unit_Factions"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Factions Test Suite
' Tests the faction interaction rules from GameLogic.bas:
' - FactionCanAttackFaction: determines if one faction can attack another
' - FactionCanHelpFaction: determines if one faction can help another
' - ClampChance: constrains a percentage value to [0..100]
' - ByteArr2String: converts a byte array to a VB string
'
' Faction values (e_Facciones):
'   Criminal=0, Ciudadano=1, Caos=2, Armada=3, concilio=4, consejo=5
'
' Rule summary:
'   Ciudadano/Armada/consejo are allies (cannot attack each other)
'   Caos/concilio are allies (cannot attack each other)
'   Cross-group attacks are allowed
'   Helping a criminal is forbidden for Ciudadano/Armada/consejo
' ==========================================================================
Public Function test_suite_factions() As Boolean
    Call UnitTesting.RunTest("test_same_faction_no_attack", test_same_faction_no_attack())
    Call UnitTesting.RunTest("test_allied_faction_no_attack", test_allied_faction_no_attack())
    Call UnitTesting.RunTest("test_cross_faction_can_attack", test_cross_faction_can_attack())
    Call UnitTesting.RunTest("test_help_same_faction", test_help_same_faction())
    Call UnitTesting.RunTest("test_help_opposing_faction", test_help_opposing_faction())
    Call UnitTesting.RunTest("test_help_criminal_blocked", test_help_criminal_blocked())
    Call UnitTesting.RunTest("test_clamp_chance_normal", test_clamp_chance_normal())
    Call UnitTesting.RunTest("test_clamp_chance_edges", test_clamp_chance_edges())
    Call UnitTesting.RunTest("test_byte_arr_to_string", test_byte_arr_to_string())
    test_suite_factions = True
End Function

' Verifies that a faction cannot attack itself.
' Ciudadano vs Ciudadano = False, Caos vs Caos = False.
Private Function test_same_faction_no_attack() As Boolean
    On Error GoTo Err_Handler
    test_same_faction_no_attack = True
    ' Ciudadano attacking Ciudadano: same faction, should be blocked (returns False)
    If FactionCanAttackFaction(e_Facciones.Ciudadano, e_Facciones.Ciudadano) Then
        test_same_faction_no_attack = False: Exit Function
    End If
    ' Caos attacking Caos: same faction, should be blocked (returns False)
    If FactionCanAttackFaction(e_Facciones.Caos, e_Facciones.Caos) Then
        test_same_faction_no_attack = False: Exit Function
    End If
    Exit Function
Err_Handler:
    test_same_faction_no_attack = False
End Function

' Verifies that allied factions cannot attack each other.
' Ciudadano/Armada/consejo are allies; Caos/concilio are allies.
Private Function test_allied_faction_no_attack() As Boolean
    On Error GoTo Err_Handler
    test_allied_faction_no_attack = True
    ' Ciudadano and Armada are in the same alliance group, attack should be blocked
    If FactionCanAttackFaction(e_Facciones.Ciudadano, e_Facciones.Armada) Then
        test_allied_faction_no_attack = False: Exit Function
    End If
    ' Armada and consejo are also allies, attack should be blocked
    If FactionCanAttackFaction(e_Facciones.Armada, e_Facciones.consejo) Then
        test_allied_faction_no_attack = False: Exit Function
    End If
    ' Caos and concilio are allies on the evil side, attack should be blocked
    If FactionCanAttackFaction(e_Facciones.Caos, e_Facciones.concilio) Then
        test_allied_faction_no_attack = False: Exit Function
    End If
    Exit Function
Err_Handler:
    test_allied_faction_no_attack = False
End Function

' Verifies that opposing factions CAN attack each other.
' Ciudadano can attack Caos, Armada can attack concilio, etc.
Private Function test_cross_faction_can_attack() As Boolean
    On Error GoTo Err_Handler
    test_cross_faction_can_attack = True
    ' Ciudadano vs Caos: opposing factions, attack should be allowed (returns True)
    If Not FactionCanAttackFaction(e_Facciones.Ciudadano, e_Facciones.Caos) Then
        test_cross_faction_can_attack = False: Exit Function
    End If
    ' Caos vs Armada: opposing factions, attack should be allowed
    If Not FactionCanAttackFaction(e_Facciones.Caos, e_Facciones.Armada) Then
        test_cross_faction_can_attack = False: Exit Function
    End If
    ' Armada vs Criminal: Criminal is not in any alliance, anyone can attack them
    If Not FactionCanAttackFaction(e_Facciones.Armada, e_Facciones.Criminal) Then
        test_cross_faction_can_attack = False: Exit Function
    End If
    Exit Function
Err_Handler:
    test_cross_faction_can_attack = False
End Function


' Verifies FactionCanHelpFaction() returns eInteractionOk when helping
' a member of the same alliance group.
Private Function test_help_same_faction() As Boolean
    On Error GoTo Err_Handler
    test_help_same_faction = True
    ' Ciudadano helping Ciudadano: same faction, should be OK
    If FactionCanHelpFaction(e_Facciones.Ciudadano, e_Facciones.Ciudadano) <> eInteractionOk Then
        test_help_same_faction = False: Exit Function
    End If
    ' Ciudadano helping Armada: same alliance group, should be OK
    If FactionCanHelpFaction(e_Facciones.Ciudadano, e_Facciones.Armada) <> eInteractionOk Then
        test_help_same_faction = False: Exit Function
    End If
    ' Caos helping concilio: same evil alliance, should be OK
    If FactionCanHelpFaction(e_Facciones.Caos, e_Facciones.concilio) <> eInteractionOk Then
        test_help_same_faction = False: Exit Function
    End If
    ' Caos helping Caos: same faction, should be OK
    If FactionCanHelpFaction(e_Facciones.Caos, e_Facciones.Caos) <> eInteractionOk Then
        test_help_same_faction = False: Exit Function
    End If
    Exit Function
Err_Handler:
    test_help_same_faction = False
End Function

' Verifies FactionCanHelpFaction() returns eOposingFaction when trying
' to help a member of the opposing alliance group.
Private Function test_help_opposing_faction() As Boolean
    On Error GoTo Err_Handler
    test_help_opposing_faction = True
    ' Ciudadano trying to help Caos: opposing factions, should be blocked
    If FactionCanHelpFaction(e_Facciones.Ciudadano, e_Facciones.Caos) <> eOposingFaction Then
        test_help_opposing_faction = False: Exit Function
    End If
    ' Ciudadano trying to help concilio: opposing factions, should be blocked
    If FactionCanHelpFaction(e_Facciones.Ciudadano, e_Facciones.concilio) <> eOposingFaction Then
        test_help_opposing_faction = False: Exit Function
    End If
    ' Caos trying to help Armada: opposing factions, should be blocked
    If FactionCanHelpFaction(e_Facciones.Caos, e_Facciones.Armada) <> eOposingFaction Then
        test_help_opposing_faction = False: Exit Function
    End If
    ' Caos trying to help Ciudadano: opposing factions, should be blocked
    If FactionCanHelpFaction(e_Facciones.Caos, e_Facciones.Ciudadano) <> eOposingFaction Then
        test_help_opposing_faction = False: Exit Function
    End If
    Exit Function
Err_Handler:
    test_help_opposing_faction = False
End Function

' Verifies FactionCanHelpFaction() returns eCantHelpCriminal when a
' Ciudadano/Armada/consejo member tries to help a Criminal.
' Criminals are outcasts that the "good" factions refuse to assist.
Private Function test_help_criminal_blocked() As Boolean
    On Error GoTo Err_Handler
    test_help_criminal_blocked = True
    ' Ciudadano trying to help Criminal: explicitly forbidden
    If FactionCanHelpFaction(e_Facciones.Ciudadano, e_Facciones.Criminal) <> eCantHelpCriminal Then
        test_help_criminal_blocked = False: Exit Function
    End If
    ' Armada trying to help Criminal: also forbidden
    If FactionCanHelpFaction(e_Facciones.Armada, e_Facciones.Criminal) <> eCantHelpCriminal Then
        test_help_criminal_blocked = False: Exit Function
    End If
    ' consejo trying to help Criminal: also forbidden
    If FactionCanHelpFaction(e_Facciones.consejo, e_Facciones.Criminal) <> eCantHelpCriminal Then
        test_help_criminal_blocked = False: Exit Function
    End If
    Exit Function
Err_Handler:
    test_help_criminal_blocked = False
End Function

' Verifies ClampChance() constrains a value to the [0..100] integer range.
' Values within range are truncated to integer (no rounding). Values outside
' are clamped to the nearest boundary.
Private Function test_clamp_chance_normal() As Boolean
    On Error GoTo Err_Handler
    test_clamp_chance_normal = True
    ' 50.0 is within range, truncated to integer 50
    If ClampChance(50!) <> 50 Then test_clamp_chance_normal = False: Exit Function
    ' 0.0 is the lower boundary
    If ClampChance(0!) <> 0 Then test_clamp_chance_normal = False: Exit Function
    ' 100.0 is the upper boundary
    If ClampChance(100!) <> 100 Then test_clamp_chance_normal = False: Exit Function
    ' 75.9 is truncated (Fix) to 75, not rounded to 76
    If ClampChance(75.9!) <> 75 Then test_clamp_chance_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_clamp_chance_normal = False
End Function

' Verifies ClampChance() clamps out-of-range values to 0 or 100.
Private Function test_clamp_chance_edges() As Boolean
    On Error GoTo Err_Handler
    test_clamp_chance_edges = True
    ' Negative value gets clamped to 0
    If ClampChance(-10!) <> 0 Then test_clamp_chance_edges = False: Exit Function
    ' Value above 100 gets clamped to 100
    If ClampChance(150!) <> 100 Then test_clamp_chance_edges = False: Exit Function
    ' Large negative gets clamped to 0
    If ClampChance(-999!) <> 0 Then test_clamp_chance_edges = False: Exit Function
    Exit Function
Err_Handler:
    test_clamp_chance_edges = False
End Function

' Verifies ByteArr2String() converts a byte array into a VB string
' by mapping each byte to its corresponding ASCII character.
Private Function test_byte_arr_to_string() As Boolean
    On Error GoTo Err_Handler
    test_byte_arr_to_string = True
    ' Build a byte array for "ABC" (codes 65, 66, 67)
    Dim arr(0 To 2) As Byte
    arr(0) = 65  ' A
    arr(1) = 66  ' B
    arr(2) = 67  ' C
    ' The function should concatenate them into "ABC"
    If ByteArr2String(arr) <> "ABC" Then test_byte_arr_to_string = False: Exit Function
    ' Single byte array for "X" (code 88)
    Dim arr2(0 To 0) As Byte
    arr2(0) = 88  ' X
    If ByteArr2String(arr2) <> "X" Then test_byte_arr_to_string = False: Exit Function
    Exit Function
Err_Handler:
    test_byte_arr_to_string = False
End Function

#End If