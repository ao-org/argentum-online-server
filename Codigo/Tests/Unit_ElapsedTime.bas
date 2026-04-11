Attribute VB_Name = "Unit_ElapsedTime"
Option Explicit
#If UNIT_TEST = 1 Then

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

Private Function test_ticks_elapsed_simple() As Boolean
    On Error GoTo Err_Handler
    test_ticks_elapsed_simple = True
    ' Normal case: current > start
    If TicksElapsed(100, 200) <> 100 Then test_ticks_elapsed_simple = False: Exit Function
    If TicksElapsed(0, 1000) <> 1000 Then test_ticks_elapsed_simple = False: Exit Function
    If TicksElapsed(500, 500) <> 0 Then test_ticks_elapsed_simple = False: Exit Function
    Exit Function
Err_Handler:
    test_ticks_elapsed_simple = False
End Function

Private Function test_ticks_elapsed_wrap() As Boolean
    On Error GoTo Err_Handler
    test_ticks_elapsed_wrap = True
    ' Wrap case: current < start (tick counter wrapped around)
    Dim elapsed As Double
    elapsed = TicksElapsed(&H7FFFFFFF, 100)
    If elapsed <= 0 Then test_ticks_elapsed_wrap = False: Exit Function
    Exit Function
Err_Handler:
    test_ticks_elapsed_wrap = False
End Function

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

Private Function test_pos_mod_zero_modulus() As Boolean
    On Error GoTo Err_Handler
    test_pos_mod_zero_modulus = True
    If PosMod(10, 0) <> 0 Then test_pos_mod_zero_modulus = False: Exit Function
    If PosMod(10, -1) <> 0 Then test_pos_mod_zero_modulus = False: Exit Function
    Exit Function
Err_Handler:
    test_pos_mod_zero_modulus = False
End Function

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

Private Function test_deadline_passed_zero() As Boolean
    On Error GoTo Err_Handler
    test_deadline_passed_zero = True
    ' deadline=0 always means "passed"
    If Not DeadlinePassed(0, 0) Then test_deadline_passed_zero = False: Exit Function
    If Not DeadlinePassed(100, 0) Then test_deadline_passed_zero = False: Exit Function
    If Not DeadlinePassed(-1, 0) Then test_deadline_passed_zero = False: Exit Function
    Exit Function
Err_Handler:
    test_deadline_passed_zero = False
End Function

Private Function test_deadline_passed_normal() As Boolean
    On Error GoTo Err_Handler
    test_deadline_passed_normal = True
    ' now >= deadline = passed
    If Not DeadlinePassed(100, 50) Then test_deadline_passed_normal = False: Exit Function
    If Not DeadlinePassed(100, 100) Then test_deadline_passed_normal = False: Exit Function
    ' now < deadline = not passed
    If DeadlinePassed(50, 100) Then test_deadline_passed_normal = False: Exit Function
    Exit Function
Err_Handler:
    test_deadline_passed_normal = False
End Function

#End If
