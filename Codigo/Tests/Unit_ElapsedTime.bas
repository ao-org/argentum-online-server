Attribute VB_Name = "Unit_ElapsedTime"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Elapsed Time Test Suite
' Tests tick-based timing utilities: elapsed calculation, tick comparison,
' positive modulo, modular addition, and deadline checking.
' ==========================================================================
Public Function test_suite_elapsed_time() As Boolean
    Call UnitTesting.RunTest("test_ticks_elapsed_simple", test_ticks_elapsed_simple())
    Call UnitTesting.RunTest("test_ticks_elapsed_signed_boundary", test_ticks_elapsed_signed_boundary())
    Call UnitTesting.RunTest("test_ticks_elapsed_true_wrap", test_ticks_elapsed_true_wrap())
    Call UnitTesting.RunTest("test_ticks_elapsed_direct_subtraction_overflow", test_ticks_elapsed_direct_subtraction_overflow())
    Call UnitTesting.RunTest("test_tick_after_basic", test_tick_after_basic())
    Call UnitTesting.RunTest("test_tick_after_signed_boundary", test_tick_after_signed_boundary())
    Call UnitTesting.RunTest("test_tick_after_true_wrap", test_tick_after_true_wrap())
    Call UnitTesting.RunTest("test_pos_mod_positive", test_pos_mod_positive())
    Call UnitTesting.RunTest("test_pos_mod_negative", test_pos_mod_negative())
    Call UnitTesting.RunTest("test_pos_mod_zero_modulus", test_pos_mod_zero_modulus())
    Call UnitTesting.RunTest("test_add_mod32_simple", test_add_mod32_simple())
    Call UnitTesting.RunTest("test_add_mod32_signed_boundary", test_add_mod32_signed_boundary())
    Call UnitTesting.RunTest("test_add_mod32_true_wrap", test_add_mod32_true_wrap())
    Call UnitTesting.RunTest("test_add_mod32_modular_subtraction", test_add_mod32_modular_subtraction())
    Call UnitTesting.RunTest("test_deadline_passed_zero", test_deadline_passed_zero())
    Call UnitTesting.RunTest("test_deadline_passed_normal", test_deadline_passed_normal())
    Call UnitTesting.RunTest("test_deadline_passed_signed_boundary", test_deadline_passed_signed_boundary())
    Call UnitTesting.RunTest("test_deadline_passed_true_wrap", test_deadline_passed_true_wrap())
    test_suite_elapsed_time = True
End Function

' Verifies TicksElapsed() returns the correct difference for normal cases
' where current tick >= start tick. Same tick = 0 elapsed.
Private Function test_ticks_elapsed_simple() As Boolean
    On Error GoTo Err_Handler
    test_ticks_elapsed_simple = True
    If TicksElapsed(100, 200) <> 100 Then test_ticks_elapsed_simple = False: Exit Function
    If TicksElapsed(0, 1000) <> 1000 Then test_ticks_elapsed_simple = False: Exit Function
    If TicksElapsed(500, 500) <> 0 Then test_ticks_elapsed_simple = False: Exit Function
    Exit Function
Err_Handler:
    test_ticks_elapsed_simple = False
End Function

' Crossing &H7FFFFFFF -> &H80000000 is the signed Long representation boundary,
' not the real modulo-2^32 timer wrap.
Private Function test_ticks_elapsed_signed_boundary() As Boolean
    On Error GoTo Err_Handler
    test_ticks_elapsed_signed_boundary = True
    If TicksElapsed(&H7FFFFFF0, &H80000010) <> 32 Then test_ticks_elapsed_signed_boundary = False: Exit Function
    Exit Function
Err_Handler:
    test_ticks_elapsed_signed_boundary = False
End Function

' Crossing &HFFFFFFFF -> &H00000000 is the real modulo-2^32 timer wrap.
Private Function test_ticks_elapsed_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_ticks_elapsed_true_wrap = True
    If TicksElapsed(&HFFFFFFF0, &H10) <> 32 Then test_ticks_elapsed_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_ticks_elapsed_true_wrap = False
End Function

' Verifies cases where direct signed Long subtraction would overflow before CDbl().
Private Function test_ticks_elapsed_direct_subtraction_overflow() As Boolean
    On Error GoTo Err_Handler
    test_ticks_elapsed_direct_subtraction_overflow = True
    If TicksElapsed(&H80000010, &H7FFFFFF0) <> 4294967264# Then test_ticks_elapsed_direct_subtraction_overflow = False: Exit Function
    If TicksElapsed(&H7FFFFFF0, &H80000010) <> 32 Then test_ticks_elapsed_direct_subtraction_overflow = False: Exit Function
    Exit Function
Err_Handler:
    test_ticks_elapsed_direct_subtraction_overflow = False
End Function

' Verifies TickAfter(a, b): True when a is at-or-after b, False when a is before b.
Private Function test_tick_after_basic() As Boolean
    On Error GoTo Err_Handler
    test_tick_after_basic = True
    If Not TickAfter(100, 50) Then test_tick_after_basic = False: Exit Function
    If Not TickAfter(100, 100) Then test_tick_after_basic = False: Exit Function
    If TickAfter(50, 100) Then test_tick_after_basic = False: Exit Function
    If Not TickAfter(0, 0) Then test_tick_after_basic = False: Exit Function
    Exit Function
Err_Handler:
    test_tick_after_basic = False
End Function

Private Function test_tick_after_signed_boundary() As Boolean
    On Error GoTo Err_Handler
    test_tick_after_signed_boundary = True
    If Not TickAfter(&H80000010, &H7FFFFFF0) Then test_tick_after_signed_boundary = False: Exit Function
    If TickAfter(&H7FFFFFF0, &H80000010) Then test_tick_after_signed_boundary = False: Exit Function
    Exit Function
Err_Handler:
    test_tick_after_signed_boundary = False
End Function

Private Function test_tick_after_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_tick_after_true_wrap = True
    If Not TickAfter(&H10, &HFFFFFFF0) Then test_tick_after_true_wrap = False: Exit Function
    If TickAfter(&HFFFFFFF0, &H10) Then test_tick_after_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_tick_after_true_wrap = False
End Function

' Verifies PosMod() with positive inputs returns the standard remainder.
Private Function test_pos_mod_positive() As Boolean
    On Error GoTo Err_Handler
    test_pos_mod_positive = True
    If PosMod(10, 3) <> 1 Then test_pos_mod_positive = False: Exit Function
    If PosMod(9, 3) <> 0 Then test_pos_mod_positive = False: Exit Function
    If PosMod(0, 5) <> 0 Then test_pos_mod_positive = False: Exit Function
    If PosMod(7, 7) <> 0 Then test_pos_mod_positive = False: Exit Function
    Exit Function
Err_Handler:
    test_pos_mod_positive = False
End Function

' Verifies PosMod() with negative inputs wraps to a positive result.
Private Function test_pos_mod_negative() As Boolean
    On Error GoTo Err_Handler
    test_pos_mod_negative = True
    If PosMod(-1, 3) <> 2 Then test_pos_mod_negative = False: Exit Function
    If PosMod(-3, 3) <> 0 Then test_pos_mod_negative = False: Exit Function
    If PosMod(-7, 5) <> 3 Then test_pos_mod_negative = False: Exit Function
    Exit Function
Err_Handler:
    test_pos_mod_negative = False
End Function

' Verifies PosMod() returns 0 as a safe fallback when the modulus is 0 or negative.
Private Function test_pos_mod_zero_modulus() As Boolean
    On Error GoTo Err_Handler
    test_pos_mod_zero_modulus = True
    If PosMod(10, 0) <> 0 Then test_pos_mod_zero_modulus = False: Exit Function
    If PosMod(10, -1) <> 0 Then test_pos_mod_zero_modulus = False: Exit Function
    Exit Function
Err_Handler:
    test_pos_mod_zero_modulus = False
End Function

' Verifies AddMod32() basic addition: 10+20=30, 0+0=0, identity 100+0=100.
Private Function test_add_mod32_simple() As Boolean
    On Error GoTo Err_Handler
    test_add_mod32_simple = True
    If AddMod32(10, 20) <> 30 Then test_add_mod32_simple = False: Exit Function
    If AddMod32(0, 0) <> 0 Then test_add_mod32_simple = False: Exit Function
    If AddMod32(100, 0) <> 100 Then test_add_mod32_simple = False: Exit Function
    Exit Function
Err_Handler:
    test_add_mod32_simple = False
End Function

Private Function test_add_mod32_signed_boundary() As Boolean
    On Error GoTo Err_Handler
    test_add_mod32_signed_boundary = True
    If AddMod32(&H7FFFFFF0, &H20) <> &H80000010 Then test_add_mod32_signed_boundary = False: Exit Function
    Exit Function
Err_Handler:
    test_add_mod32_signed_boundary = False
End Function

Private Function test_add_mod32_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_add_mod32_true_wrap = True
    If AddMod32(&HFFFFFFF0, &H20) <> &H10 Then test_add_mod32_true_wrap = False: Exit Function
    If AddMod32(&HFFFFFFFF, 1) <> 0 Then test_add_mod32_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_add_mod32_true_wrap = False
End Function

' Verifies callers can create an earlier raw timestamp without signed Long
' subtraction overflowing at either the sign boundary or the true wrap.
Private Function test_add_mod32_modular_subtraction() As Boolean
    On Error GoTo Err_Handler
    test_add_mod32_modular_subtraction = True
    If AddMod32(&H80000010, -32) <> &H7FFFFFF0 Then test_add_mod32_modular_subtraction = False: Exit Function
    If AddMod32(&H10, -32) <> &HFFFFFFF0 Then test_add_mod32_modular_subtraction = False: Exit Function
    Exit Function
Err_Handler:
    test_add_mod32_modular_subtraction = False
End Function

' Verifies DeadlinePassed() treats deadline=0 as "always passed" regardless
' of the current tick value.
Private Function test_deadline_passed_zero() As Boolean
    On Error GoTo Err_Handler
    test_deadline_passed_zero = True
    If Not DeadlinePassed(0, 0) Then test_deadline_passed_zero = False: Exit Function
    If Not DeadlinePassed(100, 0) Then test_deadline_passed_zero = False: Exit Function
    If Not DeadlinePassed(-1, 0) Then test_deadline_passed_zero = False: Exit Function
    Exit Function
Err_Handler:
    test_deadline_passed_zero = False
End Function

' Verifies DeadlinePassed() for normal cases: now at-or-after deadline means passed.
Private Function test_deadline_passed_normal() As Boolean
    On Error GoTo Err_Handler
    test_deadline_passed_normal = True
    If Not DeadlinePassed(100, 50) Then test_deadline_passed_normal = False: Exit Function
    If Not DeadlinePassed(100, 100) Then test_deadline_passed_normal = False: Exit Function
    If DeadlinePassed(50, 100) Then test_deadline_passed_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_deadline_passed_normal = False
End Function

Private Function test_deadline_passed_signed_boundary() As Boolean
    On Error GoTo Err_Handler
    test_deadline_passed_signed_boundary = True
    If Not DeadlinePassed(&H80000010, &H7FFFFFF0) Then test_deadline_passed_signed_boundary = False: Exit Function
    If DeadlinePassed(&H7FFFFFF0, &H80000010) Then test_deadline_passed_signed_boundary = False: Exit Function
    Exit Function
Err_Handler:
    test_deadline_passed_signed_boundary = False
End Function

Private Function test_deadline_passed_true_wrap() As Boolean
    On Error GoTo Err_Handler
    test_deadline_passed_true_wrap = True
    If Not DeadlinePassed(&H10, &HFFFFFFF0) Then test_deadline_passed_true_wrap = False: Exit Function
    If DeadlinePassed(&HFFFFFFF0, &H10) Then test_deadline_passed_true_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_deadline_passed_true_wrap = False
End Function

#End If
