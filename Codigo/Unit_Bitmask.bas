Attribute VB_Name = "Unit_Bitmask"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_bitmask() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("test_set_mask", test_set_mask())
    Call UnitTesting.RunTest("test_is_set", test_is_set())
    Call UnitTesting.RunTest("test_unset_mask", test_unset_mask())
    Call UnitTesting.RunTest("test_reset_mask", test_reset_mask())
    Call UnitTesting.RunTest("test_shift_left_right", test_shift_left_right())
    
    Debug.Print "Bitmask suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_bitmask = True
End Function

Private Function test_set_mask() As Boolean
    On Error GoTo test_set_mask_Err
    test_set_mask = True
    
    Dim mask As Long
    mask = 0
    Call SetMask(mask, 1)
    If mask <> 1 Then test_set_mask = False: Exit Function
    
    mask = 0
    Call SetMask(mask, 4)
    If mask <> 4 Then test_set_mask = False: Exit Function
    
    ' Set multiple bits
    mask = 0
    Call SetMask(mask, 1)
    Call SetMask(mask, 4)
    If mask <> 5 Then test_set_mask = False: Exit Function
    Exit Function
test_set_mask_Err:
    test_set_mask = False
End Function

Private Function test_is_set() As Boolean
    On Error GoTo test_is_set_Err
    test_is_set = True
    
    Dim mask As Long
    mask = 0
    Call SetMask(mask, 1)
    If Not IsSet(mask, 1) Then test_is_set = False: Exit Function
    If IsSet(mask, 2) Then test_is_set = False: Exit Function
    If IsSet(mask, 4) Then test_is_set = False: Exit Function
    
    Call SetMask(mask, 4)
    If Not IsSet(mask, 1) Then test_is_set = False: Exit Function
    If Not IsSet(mask, 4) Then test_is_set = False: Exit Function
    If IsSet(mask, 2) Then test_is_set = False: Exit Function
    Exit Function
test_is_set_Err:
    test_is_set = False
End Function

Private Function test_unset_mask() As Boolean
    On Error GoTo test_unset_mask_Err
    test_unset_mask = True
    
    Dim mask As Long
    mask = 0
    Call SetMask(mask, 1)
    Call SetMask(mask, 4)
    Call SetMask(mask, 8)
    
    ' Unset bit 4, others should remain
    Call UnsetMask(mask, 4)
    If IsSet(mask, 4) Then test_unset_mask = False: Exit Function
    If Not IsSet(mask, 1) Then test_unset_mask = False: Exit Function
    If Not IsSet(mask, 8) Then test_unset_mask = False: Exit Function
    Exit Function
test_unset_mask_Err:
    test_unset_mask = False
End Function

Private Function test_reset_mask() As Boolean
    On Error GoTo test_reset_mask_Err
    test_reset_mask = True
    
    Dim mask As Long
    mask = 0
    Call SetMask(mask, 1)
    Call SetMask(mask, 4)
    Call SetMask(mask, 16)
    
    Call ResetMask(mask)
    If mask <> 0 Then test_reset_mask = False: Exit Function
    Exit Function
test_reset_mask_Err:
    test_reset_mask = False
End Function

Private Function test_shift_left_right() As Boolean
    On Error GoTo test_shift_left_right_Err
    test_shift_left_right = True
    
    ' ShiftLeft basic cases
    If ShiftLeft(1, 0) <> 1 Then test_shift_left_right = False: Exit Function
    If ShiftLeft(1, 1) <> 2 Then test_shift_left_right = False: Exit Function
    If ShiftLeft(1, 2) <> 4 Then test_shift_left_right = False: Exit Function
    If ShiftLeft(1, 3) <> 8 Then test_shift_left_right = False: Exit Function
    If ShiftLeft(1, 10) <> 1024 Then test_shift_left_right = False: Exit Function
    
    ' ShiftRight basic cases
    If ShiftRight(1024, 10) <> 1 Then test_shift_left_right = False: Exit Function
    If ShiftRight(8, 3) <> 1 Then test_shift_left_right = False: Exit Function
    If ShiftRight(4, 1) <> 2 Then test_shift_left_right = False: Exit Function
    If ShiftRight(1, 0) <> 1 Then test_shift_left_right = False: Exit Function
    
    ' Out of range returns 0
    If ShiftLeft(1, 32) <> 0 Then test_shift_left_right = False: Exit Function
    If ShiftRight(1, 32) <> 0 Then test_shift_left_right = False: Exit Function
    Exit Function
test_shift_left_right_Err:
    test_shift_left_right = False
End Function

#End If
