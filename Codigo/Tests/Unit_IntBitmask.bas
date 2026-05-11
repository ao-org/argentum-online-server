Attribute VB_Name = "Unit_IntBitmask"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Integer Bitmask Test Suite
' Tests 16-bit Integer bitmask operations: SetIntMask, IsIntSet,
' UnsetIntMask, and ResetIntMask from Matematicas.bas.
' ==========================================================================
Public Function test_suite_int_bitmask() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("intmask_set", test_set_int_mask())
    Call UnitTesting.RunTest("intmask_is_set", test_is_int_set())
    Call UnitTesting.RunTest("intmask_unset", test_unset_int_mask())
    Call UnitTesting.RunTest("intmask_reset", test_reset_int_mask())
    Call UnitTesting.RunTest("intmask_prop_set_unset_roundtrip", test_prop_int_set_unset_roundtrip())
    
    Debug.Print "IntBitmask suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_int_bitmask = True
End Function

' Verifies SetIntMask() turns on a bit in a 16-bit Integer mask.
' Setting value 8 on a zero mask should produce mask = 8.
Private Function test_set_int_mask() As Boolean
    On Error GoTo Fail
    
    Dim Mask As Integer
    Mask = 0
    
    Call SetIntMask(Mask, 8)
    
    test_set_int_mask = (Mask = 8)
    Exit Function
Fail:
    test_set_int_mask = False
End Function

' Verifies IsIntSet() correctly detects which bits are on/off.
' After setting bit 4, IsIntSet should return True for 4 and False for 2.
Private Function test_is_int_set() As Boolean
    On Error GoTo Fail
    
    Dim Mask As Integer
    Mask = 0
    
    Call SetIntMask(Mask, 4)
    
    If Not IsIntSet(Mask, 4) Then
        test_is_int_set = False
        Exit Function
    End If
    
    If IsIntSet(Mask, 2) Then
        test_is_int_set = False
        Exit Function
    End If
    
    test_is_int_set = True
    Exit Function
Fail:
    test_is_int_set = False
End Function

' Verifies UnsetIntMask() clears a previously set bit while leaving others intact.
' With bits 4 and 8 set, unsetting 4 should leave only 8.
Private Function test_unset_int_mask() As Boolean
    On Error GoTo Fail
    
    Dim Mask As Integer
    Mask = 0
    
    Call SetIntMask(Mask, 4)
    Call SetIntMask(Mask, 8)
    
    Call UnsetIntMask(Mask, 4)
    
    If IsIntSet(Mask, 4) Then
        test_unset_int_mask = False
        Exit Function
    End If
    
    If Not IsIntSet(Mask, 8) Then
        test_unset_int_mask = False
        Exit Function
    End If
    
    test_unset_int_mask = True
    Exit Function
Fail:
    test_unset_int_mask = False
End Function

' Verifies ResetIntMask() sets the mask back to 0 regardless of current state.
Private Function test_reset_int_mask() As Boolean
    On Error GoTo Fail
    
    Dim Mask As Integer
    Mask = 0
    
    Call SetIntMask(Mask, 1)
    Call SetIntMask(Mask, 4)
    Call SetIntMask(Mask, 16)
    
    Call ResetIntMask(Mask)
    
    test_reset_int_mask = (Mask = 0)
    Exit Function
Fail:
    test_reset_int_mask = False
End Function

' Feature: full-coverage-unit-tests, Property 5: Integer bitmask set/unset round-trip
' For any single-bit power-of-2 Integer value v (bit positions 0-14), starting from
' a zero mask, after SetIntMask then UnsetIntMask, IsIntSet returns False.
' Validates: Requirements 4.5
Private Function test_prop_int_set_unset_roundtrip() As Boolean
    On Error GoTo Fail
    Dim i As Long
    Dim bitPos As Long
    Dim v As Integer
    Dim Mask As Integer
    For i = 1 To 120
        bitPos = (i - 1) Mod 15  ' 0 to 14
        v = CInt(2 ^ bitPos)  ' single-bit power-of-2 Integer
        Mask = 0
        Call SetIntMask(Mask, v)
        Call UnsetIntMask(Mask, v)
        If IsIntSet(Mask, v) Then
            test_prop_int_set_unset_roundtrip = False
            Exit Function
        End If
    Next i
    test_prop_int_set_unset_roundtrip = True
    Exit Function
Fail:
    test_prop_int_set_unset_roundtrip = False
End Function

#End If
