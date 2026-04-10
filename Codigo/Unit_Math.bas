Attribute VB_Name = "Unit_Math"
Option Explicit
#If UNIT_TEST = 1 Then

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
    
    Debug.Print "Math suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_math = True
End Function

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

Private Function test_distance_zero() As Boolean
    On Error GoTo test_distance_zero_Err
    test_distance_zero = (Distance(0, 0, 0, 0) = 0)
    Exit Function
test_distance_zero_Err:
    test_distance_zero = False
End Function

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

Private Function test_increment_counter_overflow() As Boolean
    On Error GoTo test_increment_counter_overflow_Err
    test_increment_counter_overflow = True
    
    Dim counter As Long
    counter = &H7FFFFFFF
    Call IncrementLongCounter(counter, "test")
    If counter <> 0 Then test_increment_counter_overflow = False: Exit Function
    Exit Function
test_increment_counter_overflow_Err:
    test_increment_counter_overflow = False
End Function

Private Function test_increment_counter_negative() As Boolean
    On Error GoTo test_increment_counter_negative_Err
    test_increment_counter_negative = True
    
    Dim counter As Long
    counter = -5
    Call IncrementLongCounter(counter, "test")
    If counter <> 0 Then test_increment_counter_negative = False: Exit Function
    
    counter = -1
    Call IncrementLongCounter(counter, "test")
    If counter <> 0 Then test_increment_counter_negative = False: Exit Function
    Exit Function
test_increment_counter_negative_Err:
    test_increment_counter_negative = False
End Function

#End If
