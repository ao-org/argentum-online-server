Attribute VB_Name = "Unit_BitShift"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Bit Shift Test Suite
' Tests ShiftRight and ShiftLeft from Matematicas.bas.
' Requirements: 5.1, 5.2, 5.3, 5.4
' ==========================================================================
Public Function test_suite_bit_shift() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("bitshift_right_known", test_shift_right_known())
    Call UnitTesting.RunTest("bitshift_left_known", test_shift_left_known())
    Call UnitTesting.RunTest("bitshift_out_of_range", test_shift_out_of_range())
    Call UnitTesting.RunTest("bitshift_sign_bit", test_shift_sign_bit())
    Call UnitTesting.RunTest("bitshift_prop_left_right_roundtrip", test_prop_shift_roundtrip())
    
    Debug.Print "BitShift suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_bit_shift = True
End Function

' Verifies ShiftRight() produces correct results for known inputs.
' ShiftRight(16, 2) = 4 and ShiftRight(255, 4) = 15.
' Validates: Requirements 5.1
Private Function test_shift_right_known() As Boolean
    On Error GoTo Fail
    
    If ShiftRight(16, 2) <> 4 Then
        test_shift_right_known = False
        Exit Function
    End If
    
    If ShiftRight(255, 4) <> 15 Then
        test_shift_right_known = False
        Exit Function
    End If
    
    test_shift_right_known = True
    Exit Function
Fail:
    test_shift_right_known = False
End Function

' Verifies ShiftLeft() produces correct results for known inputs.
' ShiftLeft(1, 4) = 16 and ShiftLeft(1, 0) = 1.
' Validates: Requirements 5.2
Private Function test_shift_left_known() As Boolean
    On Error GoTo Fail
    
    If ShiftLeft(1, 4) <> 16 Then
        test_shift_left_known = False
        Exit Function
    End If
    
    If ShiftLeft(1, 0) <> 1 Then
        test_shift_left_known = False
        Exit Function
    End If
    
    test_shift_left_known = True
    Exit Function
Fail:
    test_shift_left_known = False
End Function

' Verifies ShiftRight and ShiftLeft return 0 for out-of-range bit counts (> 31).
' Validates: Requirements 5.3
Private Function test_shift_out_of_range() As Boolean
    On Error GoTo Fail
    
    If ShiftRight(255, 32) <> 0 Then
        test_shift_out_of_range = False
        Exit Function
    End If
    
    If ShiftLeft(255, 32) <> 0 Then
        test_shift_out_of_range = False
        Exit Function
    End If
    
    test_shift_out_of_range = True
    Exit Function
Fail:
    test_shift_out_of_range = False
End Function

' Verifies ShiftLeft(1, 31) returns &H80000000 (the sign bit).
' Validates: Requirements 5.4
Private Function test_shift_sign_bit() As Boolean
    On Error GoTo Fail
    
    test_shift_sign_bit = (ShiftLeft(1, 31) = &H80000000)
    Exit Function
Fail:
    test_shift_sign_bit = False
End Function

' Feature: full-coverage-unit-tests, Property 6: Bit shift left/right round-trip
' For any non-negative Long value n and valid BitCount b (0 to 30),
' where n fits in (31 - b) bits, ShiftRight(ShiftLeft(n, b), b) equals n.
' **Validates: Requirements 5.5**
Private Function test_prop_shift_roundtrip() As Boolean
    On Error GoTo Fail
    Dim iterations As Long: iterations = 0
    Dim i As Long
    For i = 1 To 120
        Dim b As Byte: b = CByte(i Mod 31)  ' 0 to 30
        Dim maxBits As Long: maxBits = 31 - CLng(b)
        Dim n As Long
        If maxBits > 0 Then
            n = (CLng(i) * 7) Mod (2 ^ maxBits)  ' fits in maxBits bits
        Else
            n = 0
        End If
        If ShiftRight(ShiftLeft(n, b), b) <> n Then
            test_prop_shift_roundtrip = False
            Exit Function
        End If
        iterations = iterations + 1
    Next i
    test_prop_shift_roundtrip = (iterations >= 100)
    Exit Function
Fail:
    test_prop_shift_roundtrip = False
End Function

#End If
