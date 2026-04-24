Attribute VB_Name = "Unit_CombatMath"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_combatmath() As Boolean
    Call UnitTesting.RunTest("test_minimoint_smaller", test_minimoint_smaller())
    Call UnitTesting.RunTest("test_maximoint_larger", test_maximoint_larger())
    Call UnitTesting.RunTest("test_minimoint_equal", test_minimoint_equal())
    Call UnitTesting.RunTest("test_maximoint_equal", test_maximoint_equal())
    Call UnitTesting.RunTest("test_minimoint_negative", test_minimoint_negative())
    Call UnitTesting.RunTest("test_maximoint_negative", test_maximoint_negative())
    Call UnitTesting.RunTest("test_minmax_int_property", test_minmax_int_property())
    test_suite_combatmath = True
End Function

' Verify MinimoInt returns the smaller of two different values.
Private Function test_minimoint_smaller() As Boolean
    On Error GoTo Err_Handler
    test_minimoint_smaller = True

    If MinimoInt(3, 7) <> 3 Then test_minimoint_smaller = False: Exit Function
    If MinimoInt(10, 2) <> 2 Then test_minimoint_smaller = False: Exit Function

    Exit Function
Err_Handler:
    test_minimoint_smaller = False
End Function

' Verify MaximoInt returns the larger of two different values.
Private Function test_maximoint_larger() As Boolean
    On Error GoTo Err_Handler
    test_maximoint_larger = True

    If MaximoInt(3, 7) <> 7 Then test_maximoint_larger = False: Exit Function
    If MaximoInt(10, 2) <> 10 Then test_maximoint_larger = False: Exit Function

    Exit Function
Err_Handler:
    test_maximoint_larger = False
End Function

' Verify MinimoInt with equal values returns that same value.
Private Function test_minimoint_equal() As Boolean
    On Error GoTo Err_Handler
    test_minimoint_equal = True

    If MinimoInt(5, 5) <> 5 Then test_minimoint_equal = False: Exit Function

    Exit Function
Err_Handler:
    test_minimoint_equal = False
End Function

' Verify MaximoInt with equal values returns that same value.
Private Function test_maximoint_equal() As Boolean
    On Error GoTo Err_Handler
    test_maximoint_equal = True

    If MaximoInt(5, 5) <> 5 Then test_maximoint_equal = False: Exit Function

    Exit Function
Err_Handler:
    test_maximoint_equal = False
End Function

' Verify MinimoInt with negative values returns the correct minimum.
Private Function test_minimoint_negative() As Boolean
    On Error GoTo Err_Handler
    test_minimoint_negative = True

    If MinimoInt(-10, -3) <> -10 Then test_minimoint_negative = False: Exit Function
    If MinimoInt(-5, 5) <> -5 Then test_minimoint_negative = False: Exit Function

    Exit Function
Err_Handler:
    test_minimoint_negative = False
End Function

' Verify MaximoInt with negative values returns the correct maximum.
Private Function test_maximoint_negative() As Boolean
    On Error GoTo Err_Handler
    test_maximoint_negative = True

    If MaximoInt(-10, -3) <> -3 Then test_maximoint_negative = False: Exit Function
    If MaximoInt(-5, 5) <> 5 Then test_maximoint_negative = False: Exit Function

    Exit Function
Err_Handler:
    test_maximoint_negative = False
End Function

' Property 1: MinimoInt and MaximoInt correctness
' For any two Integer values a and b, MinimoInt(a,b) <= a, MinimoInt(a,b) <= b,
' MaximoInt(a,b) >= a, MaximoInt(a,b) >= b, and one of {a, b} equals the min and max.
' Uses 200 randomized trials to approximate universal quantification.
Private Function test_minmax_int_property() As Boolean
    On Error GoTo Err_Handler
    test_minmax_int_property = True
    
    Dim i As Long
    Dim a As Integer
    Dim b As Integer
    Dim minResult As Integer
    Dim maxResult As Integer
    
    For i = 1 To 200
        ' Generate random Integer values across full Integer range (-32768 to 32767)
        a = CInt(Int(Rnd * 65536) - 32768)
        b = CInt(Int(Rnd * 65536) - 32768)
        
        minResult = MinimoInt(a, b)
        maxResult = MaximoInt(a, b)
        
        ' MinimoInt must be <= both inputs
        If minResult > a Then test_minmax_int_property = False: Exit Function
        If minResult > b Then test_minmax_int_property = False: Exit Function
        
        ' MaximoInt must be >= both inputs
        If maxResult < a Then test_minmax_int_property = False: Exit Function
        If maxResult < b Then test_minmax_int_property = False: Exit Function
        
        ' Min must equal one of {a, b}
        If minResult <> a And minResult <> b Then test_minmax_int_property = False: Exit Function
        
        ' Max must equal one of {a, b}
        If maxResult <> a And maxResult <> b Then test_minmax_int_property = False: Exit Function
    Next i
    Exit Function
Err_Handler:
    test_minmax_int_property = False
End Function

#End If
