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
    Call UnitTesting.RunTest("test_ticks_elapsed_wrap", test_ticks_elapsed_wrap())
    Call UnitTesting.RunTest("test_tick_after_basic", test_tick_after_basic())
    Call UnitTesting.RunTest("test_pos_mod_positive", test_pos_mod_positive())
    Call UnitTesting.RunTest("test_pos_mod_negative", test_pos_mod_negative())
    Call UnitTesting.RunTest("test_pos_mod_zero_modulus", test_pos_mod_zero_modulus())
    Call UnitTesting.RunTest("test_add_mod32_simple", test_add_mod32_simple())
    Call UnitTesting.RunTest("test_deadline_passed_zero", test_deadline_passed_zero())
    Call UnitTesting.RunTest("test_deadline_passed_normal", test_deadline_passed_normal())
    test_suite_elapsed_time = True
End Function

' Verifies TicksElapsed() returns the correct difference for normal cases
' where current tick >= start tick. Same tick = 0 elapsed.
Private Function test_ticks_elapsed_simple() As Boolean
    On Error GoTo Err_Handler
    test_ticks_elapsed_simple = True
    ' Started at tick 100, now at tick 200: 200 - 100 = 100 elapsed
    If TicksElapsed(100, 200) <> 100 Then test_ticks_elapsed_simple = False: Exit Function
    ' Started at tick 0, now at tick 1000: 1000 elapsed
    If TicksElapsed(0, 1000) <> 1000 Then test_ticks_elapsed_simple = False: Exit Function
    ' Same start and current tick: 0 elapsed
    If TicksElapsed(500, 500) <> 0 Then test_ticks_elapsed_simple = False: Exit Function
    Exit Function
Err_Handler:
    test_ticks_elapsed_simple = False
End Function

' Verifies TicksElapsed() handles tick counter wrap-around correctly.
' When current < start (counter overflowed), elapsed should still be positive.
Private Function test_ticks_elapsed_wrap() As Boolean
    On Error GoTo Err_Handler
    test_ticks_elapsed_wrap = True
    ' Simulate wrap-around: start was near Long.MaxValue, current is 100
    ' This happens when GetTickCount overflows past 2^31. Elapsed must still be positive.
    Dim elapsed As Double
    elapsed = TicksElapsed(&H7FFFFFFF, 100)
    If elapsed <= 0 Then test_ticks_elapsed_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_ticks_elapsed_wrap = False
End Function

' Verifies TickAfter(a, b): True when a >= b, False when a < b.
' Also checks the edge case where both are 0.
Private Function test_tick_after_basic() As Boolean
    On Error GoTo Err_Handler
    test_tick_after_basic = True
    ' 100 is after 50 (100 >= 50)
    If Not TickAfter(100, 50) Then test_tick_after_basic = False: Exit Function
    ' 100 is at-or-after 100 (equal counts as "after")
    If Not TickAfter(100, 100) Then test_tick_after_basic = False: Exit Function
    ' 50 is NOT after 100 (50 < 100)
    If TickAfter(50, 100) Then test_tick_after_basic = False: Exit Function
    ' 0 is at-or-after 0 (edge case: both zero)
    If Not TickAfter(0, 0) Then test_tick_after_basic = False: Exit Function
    Exit Function
Err_Handler:
    test_tick_after_basic = False
End Function

' Verifies PosMod() with positive inputs returns the standard remainder.
' 10 mod 3 = 1, 9 mod 3 = 0, 0 mod 5 = 0, 7 mod 7 = 0.
Private Function test_pos_mod_positive() As Boolean
    On Error GoTo Err_Handler
    test_pos_mod_positive = True
    ' 10 / 3 = remainder 1
    If PosMod(10, 3) <> 1 Then test_pos_mod_positive = False: Exit Function
    ' 9 / 3 = remainder 0 (evenly divisible)
    If PosMod(9, 3) <> 0 Then test_pos_mod_positive = False: Exit Function
    ' 0 mod anything = 0
    If PosMod(0, 5) <> 0 Then test_pos_mod_positive = False: Exit Function
    ' 7 mod 7 = 0 (evenly divisible)
    If PosMod(7, 7) <> 0 Then test_pos_mod_positive = False: Exit Function
    Exit Function
Err_Handler:
    test_pos_mod_positive = False
End Function

' Verifies PosMod() with negative inputs wraps to a positive result.
' -1 mod 3 = 2, -3 mod 3 = 0, -7 mod 5 = 3.
Private Function test_pos_mod_negative() As Boolean
    On Error GoTo Err_Handler
    test_pos_mod_negative = True
    ' -1 mod 3: VB Mod gives -1, but PosMod wraps to 2 (positive result)
    If PosMod(-1, 3) <> 2 Then test_pos_mod_negative = False: Exit Function
    ' -3 mod 3 = 0 (evenly divisible, even when negative)
    If PosMod(-3, 3) <> 0 Then test_pos_mod_negative = False: Exit Function
    ' -7 mod 5: VB Mod gives -2, PosMod wraps to 3
    If PosMod(-7, 5) <> 3 Then test_pos_mod_negative = False: Exit Function
    Exit Function
Err_Handler:
    test_pos_mod_negative = False
End Function

' Verifies PosMod() returns 0 as a safe fallback when the modulus is 0 or negative.
Private Function test_pos_mod_zero_modulus() As Boolean
    On Error GoTo Err_Handler
    test_pos_mod_zero_modulus = True
    ' Modulus of 0 is undefined; PosMod returns 0 as a safe fallback
    If PosMod(10, 0) <> 0 Then test_pos_mod_zero_modulus = False: Exit Function
    ' Negative modulus is also invalid; returns 0
    If PosMod(10, -1) <> 0 Then test_pos_mod_zero_modulus = False: Exit Function
    Exit Function
Err_Handler:
    test_pos_mod_zero_modulus = False
End Function

' Verifies AddMod32() basic addition: 10+20=30, 0+0=0, identity 100+0=100.
Private Function test_add_mod32_simple() As Boolean
    On Error GoTo Err_Handler
    test_add_mod32_simple = True
    ' Simple addition: 10 + 20 = 30
    If AddMod32(10, 20) <> 30 Then test_add_mod32_simple = False: Exit Function
    ' Zero + zero = zero
    If AddMod32(0, 0) <> 0 Then test_add_mod32_simple = False: Exit Function
    ' Adding zero is identity: 100 + 0 = 100
    If AddMod32(100, 0) <> 100 Then test_add_mod32_simple = False: Exit Function
    Exit Function
Err_Handler:
    test_add_mod32_simple = False
End Function

' Verifies DeadlinePassed() treats deadline=0 as "always passed" regardless
' of the current tick value (even negative).
Private Function test_deadline_passed_zero() As Boolean
    On Error GoTo Err_Handler
    test_deadline_passed_zero = True
    ' A deadline of 0 is treated as "no deadline" / "already passed"
    ' regardless of what the current tick is
    If Not DeadlinePassed(0, 0) Then test_deadline_passed_zero = False: Exit Function
    If Not DeadlinePassed(100, 0) Then test_deadline_passed_zero = False: Exit Function
    ' Even a negative tick still counts as passed when deadline is 0
    If Not DeadlinePassed(-1, 0) Then test_deadline_passed_zero = False: Exit Function
    Exit Function
Err_Handler:
    test_deadline_passed_zero = False
End Function

' Verifies DeadlinePassed() for normal cases: now >= deadline means passed,
' now < deadline means not passed.
Private Function test_deadline_passed_normal() As Boolean
    On Error GoTo Err_Handler
    test_deadline_passed_normal = True
    ' now=100, deadline=50: we're past the deadline
    If Not DeadlinePassed(100, 50) Then test_deadline_passed_normal = False: Exit Function
    ' now=100, deadline=100: exactly at the deadline counts as passed
    If Not DeadlinePassed(100, 100) Then test_deadline_passed_normal = False: Exit Function
    ' now=50, deadline=100: we haven't reached the deadline yet
    If DeadlinePassed(50, 100) Then test_deadline_passed_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_deadline_passed_normal = False
End Function

#End If
