Attribute VB_Name = "Unit_Bitmask"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Bitmask Test Suite
' Tests bitwise operations: set/unset/reset bits, shift left/right,
' and Integer-width (16-bit) mask variants.
' ==========================================================================
Public Function test_suite_bitmask() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("test_set_mask", test_set_mask())
    Call UnitTesting.RunTest("test_is_set", test_is_set())
    Call UnitTesting.RunTest("test_unset_mask", test_unset_mask())
    Call UnitTesting.RunTest("test_reset_mask", test_reset_mask())
    Call UnitTesting.RunTest("test_shift_left_right", test_shift_left_right())
    Call UnitTesting.RunTest("test_int_mask_ops", test_int_mask_ops())
    
    Debug.Print "Bitmask suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_bitmask = True
End Function

' Verifies SetMask() turns on individual bits in a Long bitmask.
' Setting bit 1 on 0 gives 1, bit 4 gives 4, both together gives 5 (1 OR 4).
Private Function test_set_mask() As Boolean
    On Error GoTo test_set_mask_Err
    test_set_mask = True
    
    Dim mask As Long
    mask = 0
    ' Set bit 1 (value 1) on an empty mask -> mask becomes 1
    Call SetMask(mask, 1)
    If mask <> 1 Then test_set_mask = False: Exit Function
    
    ' Set bit 4 (value 4) on a fresh mask -> mask becomes 4
    mask = 0
    Call SetMask(mask, 4)
    If mask <> 4 Then test_set_mask = False: Exit Function
    
    ' Set both bit 1 and bit 4 -> mask becomes 5 (binary 101)
    mask = 0
    Call SetMask(mask, 1)
    Call SetMask(mask, 4)
    If mask <> 5 Then test_set_mask = False: Exit Function
    Exit Function
test_set_mask_Err:
    test_set_mask = False
End Function

' Verifies IsSet() correctly detects which bits are on/off.
' After setting bit 1: bit 1 is set, bits 2 and 4 are not.
' After also setting bit 4: bits 1 and 4 are set, bit 2 still is not.
Private Function test_is_set() As Boolean
    On Error GoTo test_is_set_Err
    test_is_set = True
    
    Dim mask As Long
    mask = 0
    ' After setting bit 1, only bit 1 should be detected
    Call SetMask(mask, 1)
    If Not IsSet(mask, 1) Then test_is_set = False: Exit Function
    ' Bit 2 was never set, should return False
    If IsSet(mask, 2) Then test_is_set = False: Exit Function
    ' Bit 4 was never set, should return False
    If IsSet(mask, 4) Then test_is_set = False: Exit Function
    
    ' Now also set bit 4; both 1 and 4 should be detected, but not 2
    Call SetMask(mask, 4)
    If Not IsSet(mask, 1) Then test_is_set = False: Exit Function
    If Not IsSet(mask, 4) Then test_is_set = False: Exit Function
    If IsSet(mask, 2) Then test_is_set = False: Exit Function
    Exit Function
test_is_set_Err:
    test_is_set = False
End Function

' Verifies UnsetMask() clears a single bit without affecting others.
' With bits 1, 4, 8 set: unsetting bit 4 leaves only 1 and 8.
Private Function test_unset_mask() As Boolean
    On Error GoTo test_unset_mask_Err
    test_unset_mask = True
    
    Dim mask As Long
    mask = 0
    ' Set bits 1, 4, and 8 (mask = 13 = binary 1101)
    Call SetMask(mask, 1)
    Call SetMask(mask, 4)
    Call SetMask(mask, 8)
    
    ' Clear only bit 4; bits 1 and 8 should remain untouched
    Call UnsetMask(mask, 4)
    If IsSet(mask, 4) Then test_unset_mask = False: Exit Function
    If Not IsSet(mask, 1) Then test_unset_mask = False: Exit Function
    If Not IsSet(mask, 8) Then test_unset_mask = False: Exit Function
    Exit Function
test_unset_mask_Err:
    test_unset_mask = False
End Function

' Verifies ResetMask() clears all bits back to 0 regardless of current state.
Private Function test_reset_mask() As Boolean
    On Error GoTo test_reset_mask_Err
    test_reset_mask = True
    
    Dim mask As Long
    mask = 0
    ' Set several bits so the mask is non-zero
    Call SetMask(mask, 1)
    Call SetMask(mask, 4)
    Call SetMask(mask, 16)
    
    ' ResetMask should clear everything back to 0
    Call ResetMask(mask)
    If mask <> 0 Then test_reset_mask = False: Exit Function
    Exit Function
test_reset_mask_Err:
    test_reset_mask = False
End Function

' Verifies ShiftLeft() and ShiftRight() perform bitwise shifts correctly.
' ShiftLeft(1, n) should produce 2^n. ShiftRight reverses it.
' Out-of-range shifts (>=32) return 0 as a safety measure.
Private Function test_shift_left_right() As Boolean
    On Error GoTo test_shift_left_right_Err
    test_shift_left_right = True
    
    ' ShiftLeft: shifting 1 left by N positions gives 2^N
    If ShiftLeft(1, 0) <> 1 Then test_shift_left_right = False: Exit Function
    If ShiftLeft(1, 1) <> 2 Then test_shift_left_right = False: Exit Function
    If ShiftLeft(1, 2) <> 4 Then test_shift_left_right = False: Exit Function
    If ShiftLeft(1, 3) <> 8 Then test_shift_left_right = False: Exit Function
    If ShiftLeft(1, 10) <> 1024 Then test_shift_left_right = False: Exit Function
    
    ' ShiftRight: reverses ShiftLeft (1024 >> 10 = 1, 8 >> 3 = 1, etc.)
    If ShiftRight(1024, 10) <> 1 Then test_shift_left_right = False: Exit Function
    If ShiftRight(8, 3) <> 1 Then test_shift_left_right = False: Exit Function
    If ShiftRight(4, 1) <> 2 Then test_shift_left_right = False: Exit Function
    If ShiftRight(1, 0) <> 1 Then test_shift_left_right = False: Exit Function
    
    ' Shifting by 32 or more is out of range for a 32-bit Long; returns 0
    If ShiftLeft(1, 32) <> 0 Then test_shift_left_right = False: Exit Function
    If ShiftRight(1, 32) <> 0 Then test_shift_left_right = False: Exit Function
    Exit Function
test_shift_left_right_Err:
    test_shift_left_right = False
End Function

' Verifies the Integer-width (16-bit) bitmask operations:
' SetIntMask, IsIntSet, UnsetIntMask, and ResetIntMask.
' Same logic as Long variants but operating on Integer type.
Private Function test_int_mask_ops() As Boolean
    On Error GoTo test_int_mask_ops_Err
    test_int_mask_ops = True
    
    Dim mask As Integer
    mask = 0
    
    ' SetIntMask: set bit 1 on a 16-bit Integer mask
    Call SetIntMask(mask, 1)
    If mask <> 1 Then test_int_mask_ops = False: Exit Function
    
    ' IsIntSet: bit 1 is set, bit 2 is not
    If Not IsIntSet(mask, 1) Then test_int_mask_ops = False: Exit Function
    If IsIntSet(mask, 2) Then test_int_mask_ops = False: Exit Function
    
    ' Set another bit (4); both 1 and 4 should now be set
    Call SetIntMask(mask, 4)
    If Not IsIntSet(mask, 4) Then test_int_mask_ops = False: Exit Function
    If Not IsIntSet(mask, 1) Then test_int_mask_ops = False: Exit Function
    
    ' UnsetIntMask: clear bit 1, bit 4 should remain
    Call UnsetIntMask(mask, 1)
    If IsIntSet(mask, 1) Then test_int_mask_ops = False: Exit Function
    If Not IsIntSet(mask, 4) Then test_int_mask_ops = False: Exit Function
    
    ' ResetIntMask: clear all bits back to 0
    Call ResetIntMask(mask)
    If mask <> 0 Then test_int_mask_ops = False: Exit Function
    Exit Function
test_int_mask_ops_Err:
    test_int_mask_ops = False
End Function

#End If
