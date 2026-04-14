Attribute VB_Name = "Unit_Math"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Math Test Suite
' Tests core math utilities: percentage calculation, distance, random numbers,
' min/max, counter overflow, world-position distance, trigonometry, vectors.
' ==========================================================================
Public Function test_suite_math() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("test_percentage", test_percentage())
    Call UnitTesting.RunTest("test_distance_zero", test_distance_zero())
    Call UnitTesting.RunTest("test_distance_positive", test_distance_positive())
    Call UnitTesting.RunTest("test_random_number_bounds", test_random_number_bounds())
    Call UnitTesting.RunTest("test_max_min", test_max_min())
    Call UnitTesting.RunTest("test_increment_counter_overflow", test_increment_counter_overflow())
    Call UnitTesting.RunTest("test_increment_counter_negative", test_increment_counter_negative())
    Call UnitTesting.RunTest("test_distancia_same_map", test_distancia_same_map())
    Call UnitTesting.RunTest("test_distancia_diff_map", test_distancia_diff_map())
    Call UnitTesting.RunTest("test_to_radians", test_to_radians())
    Call UnitTesting.RunTest("test_get_direction", test_get_direction())
    Call UnitTesting.RunTest("test_tick_after", test_tick_after())
    Call UnitTesting.RunTest("test_pos_mod", test_pos_mod())
    Call UnitTesting.RunTest("test_add_mod32", test_add_mod32())
    Call UnitTesting.RunTest("test_rotate_vector", test_rotate_vector())
    Call UnitTesting.RunTest("test_get_normal", test_get_normal())
    
    Debug.Print "Math suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_math = True
End Function

' Verifies Porcentaje() correctly computes (value * percent / 100).
' Checks known values (1% of 100 = 1, etc.) and iterates 1..100 and 1..1000
' to confirm linear scaling holds for all integer percentages.
Private Function test_percentage() As Boolean
    On Error GoTo test_percentage_Err
    test_percentage = True
    
    If Porcentaje(100#, 1#) <> 1 Then test_percentage = False: Exit Function
    If Porcentaje(100#, 2#) <> 2 Then test_percentage = False: Exit Function
    If Porcentaje(100#, 5#) <> 5 Then test_percentage = False: Exit Function
    If Porcentaje(100#, 10#) <> 10 Then test_percentage = False: Exit Function
    If Porcentaje(100#, 100#) <> 100 Then test_percentage = False: Exit Function
    
    Dim i As Integer
    For i = 1 To 100
        If Porcentaje(100#, CDbl(i)) <> CDbl(i) Then
            test_percentage = False: Exit Function
        End If
    Next i
    For i = 1 To 1000
        If Porcentaje(1000#, CDbl(i)) <> CDbl(i) * 10 Then
            test_percentage = False: Exit Function
        End If
    Next i
    Exit Function
test_percentage_Err:
    test_percentage = False
End Function

' Verifies that the distance between the origin and itself is zero.
Private Function test_distance_zero() As Boolean
    On Error GoTo test_distance_zero_Err
    test_distance_zero = (Distance(0, 0, 0, 0) = 0)
    Exit Function
test_distance_zero_Err:
    test_distance_zero = False
End Function

' Verifies that Distance() always returns a positive value for non-zero inputs,
' including cases with negative coordinates.
Private Function test_distance_positive() As Boolean
    On Error GoTo test_distance_positive_Err
    test_distance_positive = True
    
    Dim i As Integer
    For i = 1 To 100
        If Distance(i, 0, 0, 0) <= 0 Then
            test_distance_positive = False: Exit Function
        End If
    Next i
    For i = 1 To 100
        If Distance(i, 0, -i, 0) <= 0 Then
            test_distance_positive = False: Exit Function
        End If
    Next i
    Exit Function
test_distance_positive_Err:
    test_distance_positive = False
End Function

' Verifies RandomNumber() always returns a value within [min, max].
' Tests degenerate cases (min=max) and iterates 1000 times each for
' positive and negative ranges to confirm bounds are never violated.
Private Function test_random_number_bounds() As Boolean
    On Error GoTo test_random_number_bounds_Err
    test_random_number_bounds = True
    
    If RandomNumber(0, 0) <> 0 Then test_random_number_bounds = False: Exit Function
    If RandomNumber(-1, -1) <> -1 Then test_random_number_bounds = False: Exit Function
    If RandomNumber(1, 1) <> 1 Then test_random_number_bounds = False: Exit Function
    
    Dim i As Integer
    Dim n As Long
    For i = 1 To 1000
        n = RandomNumber(0, CLng(i))
        If n < 0 Or n > i Then
            test_random_number_bounds = False: Exit Function
        End If
    Next i
    For i = 1 To 1000
        n = RandomNumber(CLng(-i), 0)
        If n < -i Or n > 0 Then
            test_random_number_bounds = False: Exit Function
        End If
    Next i
    Exit Function
test_random_number_bounds_Err:
    test_random_number_bounds = False
End Function

' Verifies Max() and Min() return the correct value for normal, equal,
' and mixed-sign inputs.
Private Function test_max_min() As Boolean
    On Error GoTo test_max_min_Err
    test_max_min = True
    
    If max(1#, 2#) <> 2# Then test_max_min = False: Exit Function
    If max(2#, 1#) <> 2# Then test_max_min = False: Exit Function
    If max(5#, 5#) <> 5# Then test_max_min = False: Exit Function
    If max(-1#, 1#) <> 1# Then test_max_min = False: Exit Function
    
    If Min(1#, 2#) <> 1# Then test_max_min = False: Exit Function
    If Min(2#, 1#) <> 1# Then test_max_min = False: Exit Function
    If Min(5#, 5#) <> 5# Then test_max_min = False: Exit Function
    If Min(-1#, 1#) <> -1# Then test_max_min = False: Exit Function
    Exit Function
test_max_min_Err:
    test_max_min = False
End Function

' Verifies that IncrementLongCounter resets to 0 when the counter is at
' Long.MaxValue (0x7FFFFFFF), preventing overflow errors.
Private Function test_increment_counter_overflow() As Boolean
    On Error GoTo test_increment_counter_overflow_Err
    test_increment_counter_overflow = True
    
    Dim counter As Long
    counter = &H7FFFFFFF
    Call IncrementLongCounter(counter, "")
    If counter <> 0 Then test_increment_counter_overflow = False: Exit Function
    Exit Function
test_increment_counter_overflow_Err:
    test_increment_counter_overflow = False
End Function

' Verifies that IncrementLongCounter resets negative counters to 0
' instead of incrementing them, treating negative values as invalid state.
Private Function test_increment_counter_negative() As Boolean
    On Error GoTo test_increment_counter_negative_Err
    test_increment_counter_negative = True
    
    Dim counter As Long
    counter = -5
    Call IncrementLongCounter(counter, "")
    If counter <> 0 Then test_increment_counter_negative = False: Exit Function
    
    counter = -1
    Call IncrementLongCounter(counter, "")
    If counter <> 0 Then test_increment_counter_negative = False: Exit Function
    Exit Function
test_increment_counter_negative_Err:
    test_increment_counter_negative = False
End Function

' Verifies Distancia() (world-position distance) on the same map.
' Uses Manhattan distance: |x1-x2| + |y1-y2| + |map1-map2|*100.
' Same position should be 0; (50,50)->(53,54) on map 1 should be 7.
Private Function test_distancia_same_map() As Boolean
    On Error GoTo test_distancia_same_map_Err
    test_distancia_same_map = True
    
    Dim p1 As t_WorldPos, p2 As t_WorldPos
    ' Same position = 0
    p1.Map = 1: p1.x = 50: p1.y = 50
    p2.Map = 1: p2.x = 50: p2.y = 50
    If Distancia(p1, p2) <> 0 Then test_distancia_same_map = False: Exit Function
    
    ' Manhattan distance on same map
    p2.x = 53: p2.y = 54
    ' |50-53| + |50-54| + |1-1|*100 = 3 + 4 + 0 = 7
    If Distancia(p1, p2) <> 7 Then test_distancia_same_map = False: Exit Function
    Exit Function
test_distancia_same_map_Err:
    test_distancia_same_map = False
End Function

' Verifies Distancia() adds a 100-unit penalty per map difference.
' Same coords on maps 1 vs 3 = |1-3|*100 = 200.
' With offset coords: 2 + 5 + 200 = 207.
Private Function test_distancia_diff_map() As Boolean
    On Error GoTo test_distancia_diff_map_Err
    test_distancia_diff_map = True
    
    Dim p1 As t_WorldPos, p2 As t_WorldPos
    p1.Map = 1: p1.x = 50: p1.y = 50
    p2.Map = 3: p2.x = 50: p2.y = 50
    ' Same pos, different maps: |1-3|*100 = 200
    If Distancia(p1, p2) <> 200 Then test_distancia_diff_map = False: Exit Function
    
    p2.x = 52: p2.y = 55
    ' |50-52| + |50-55| + |1-3|*100 = 2 + 5 + 200 = 207
    If Distancia(p1, p2) <> 207 Then test_distancia_diff_map = False: Exit Function
    Exit Function
test_distancia_diff_map_Err:
    test_distancia_diff_map = False
End Function

' Verifies ToRadians() converts degrees to radians correctly.
' 0° -> 0, 180° -> PI, 90° -> PI/2 (within floating-point tolerance).
Private Function test_to_radians() As Boolean
    On Error GoTo test_to_radians_Err
    test_to_radians = True
    
    ' 0 degrees = 0 radians
    If Abs(ToRadians(0#)) > 0.0001 Then test_to_radians = False: Exit Function
    ' 180 degrees ~ PI
    If Abs(ToRadians(180#) - 3.14159265) > 0.001 Then test_to_radians = False: Exit Function
    ' 90 degrees ~ PI/2
    If Abs(ToRadians(90#) - 1.5707963) > 0.001 Then test_to_radians = False: Exit Function
    Exit Function
test_to_radians_Err:
    test_to_radians = False
End Function

' Verifies GetDirection() returns the vector from p1 to p2.
' (10,20)->(15,25) should give (5,5); reversed should give (-5,-5).
Private Function test_get_direction() As Boolean
    On Error GoTo test_get_direction_Err
    test_get_direction = True
    
    Dim p1 As t_WorldPos, p2 As t_WorldPos
    Dim v As t_Vector
    p1.Map = 1: p1.x = 10: p1.y = 20
    p2.Map = 1: p2.x = 15: p2.y = 25
    v = GetDirection(p1, p2)
    If v.x <> 5 Then test_get_direction = False: Exit Function
    If v.y <> 5 Then test_get_direction = False: Exit Function
    
    ' Negative direction
    v = GetDirection(p2, p1)
    If v.x <> -5 Then test_get_direction = False: Exit Function
    If v.y <> -5 Then test_get_direction = False: Exit Function
    Exit Function
test_get_direction_Err:
    test_get_direction = False
End Function

' Verifies TickAfter(a, b) returns True when a >= b (tick a is at or after b).
' 100 >= 50 -> True, 100 >= 100 -> True, 50 >= 100 -> False, 0 >= 0 -> True.
Private Function test_tick_after() As Boolean
    On Error GoTo test_tick_after_Err
    test_tick_after = True
    
    ' a >= b means a is "after" b
    If Not TickAfter(100, 50) Then test_tick_after = False: Exit Function
    If Not TickAfter(100, 100) Then test_tick_after = False: Exit Function
    If TickAfter(50, 100) Then test_tick_after = False: Exit Function
    
    ' Edge: 0 is after 0
    If Not TickAfter(0, 0) Then test_tick_after = False: Exit Function
    Exit Function
test_tick_after_Err:
    test_tick_after = False
End Function

' Verifies PosMod() always returns a non-negative remainder.
' Negative inputs wrap correctly (-1 mod 3 = 2).
' Zero or negative modulus returns 0 as a safe fallback.
Private Function test_pos_mod() As Boolean
    On Error GoTo test_pos_mod_Err
    test_pos_mod = True
    
    ' Basic modulo
    If PosMod(10, 3) <> 1 Then test_pos_mod = False: Exit Function
    If PosMod(9, 3) <> 0 Then test_pos_mod = False: Exit Function
    If PosMod(0, 5) <> 0 Then test_pos_mod = False: Exit Function
    
    ' Negative input should still return positive result
    If PosMod(-1, 3) <> 2 Then test_pos_mod = False: Exit Function
    
    ' m <= 0 returns 0
    If PosMod(10, 0) <> 0 Then test_pos_mod = False: Exit Function
    If PosMod(10, -1) <> 0 Then test_pos_mod = False: Exit Function
    Exit Function
test_pos_mod_Err:
    test_pos_mod = False
End Function

' Verifies AddMod32() performs addition with 32-bit modular arithmetic.
' Basic cases: 10+20=30, 0+0=0, identity (100+0=100).
Private Function test_add_mod32() As Boolean
    On Error GoTo test_add_mod32_Err
    test_add_mod32 = True
    
    ' Simple addition
    If AddMod32(10, 20) <> 30 Then test_add_mod32 = False: Exit Function
    If AddMod32(0, 0) <> 0 Then test_add_mod32 = False: Exit Function
    
    ' Adding 0 is identity
    If AddMod32(100, 0) <> 100 Then test_add_mod32 = False: Exit Function
    Exit Function
test_add_mod32_Err:
    test_add_mod32 = False
End Function

' Verifies RotateVector() rotates a 2D vector by the given angle in radians.
' (1,0) rotated 0° stays (1,0); rotated 90° becomes (0,1); rotated 180° becomes (-1,0).
Private Function test_rotate_vector() As Boolean
    On Error GoTo test_rotate_vector_Err
    test_rotate_vector = True
    
    Dim v As t_Vector
    Dim r As t_Vector
    
    ' Rotate (1, 0) by 0 radians = (1, 0)
    v.x = 1: v.y = 0
    r = RotateVector(v, 0#)
    If Abs(r.x - 1#) > 0.001 Then test_rotate_vector = False: Exit Function
    If Abs(r.y - 0#) > 0.001 Then test_rotate_vector = False: Exit Function
    
    ' Rotate (1, 0) by PI/2 ~ (0, 1)
    r = RotateVector(v, ToRadians(90#))
    If Abs(r.x - 0#) > 0.01 Then test_rotate_vector = False: Exit Function
    If Abs(r.y - 1#) > 0.01 Then test_rotate_vector = False: Exit Function
    
    ' Rotate (1, 0) by PI ~ (-1, 0)
    r = RotateVector(v, ToRadians(180#))
    If Abs(r.x - (-1#)) > 0.01 Then test_rotate_vector = False: Exit Function
    If Abs(r.y - 0#) > 0.01 Then test_rotate_vector = False: Exit Function
    Exit Function
test_rotate_vector_Err:
    test_rotate_vector = False
End Function

' Verifies GetNormal() returns a unit-length vector in the same direction.
' (3,4) has length 5, so normalized = (0.6, 0.8). (1,0) is already unit length.
Private Function test_get_normal() As Boolean
    On Error GoTo test_get_normal_Err
    test_get_normal = True
    
    Dim v As t_Vector
    Dim n As t_Vector
    
    ' Normalize (3, 4) -> (0.6, 0.8), length = 5
    v.x = 3: v.y = 4
    n = GetNormal(v)
    If Abs(n.x - 0.6) > 0.001 Then test_get_normal = False: Exit Function
    If Abs(n.y - 0.8) > 0.001 Then test_get_normal = False: Exit Function
    
    ' Normalize (1, 0) -> (1, 0)
    v.x = 1: v.y = 0
    n = GetNormal(v)
    If Abs(n.x - 1#) > 0.001 Then test_get_normal = False: Exit Function
    If Abs(n.y - 0#) > 0.001 Then test_get_normal = False: Exit Function
    Exit Function
test_get_normal_Err:
    test_get_normal = False
End Function

#End If
