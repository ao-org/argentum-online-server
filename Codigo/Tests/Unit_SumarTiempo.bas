Attribute VB_Name = "Unit_SumarTiempo"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' SumarTiempo Test Suite
' Tests the SumarTiempo function from modTime.bas which converts a number
' of seconds into "hh:mm:ss" formatted string.
' Requirements: 9.1, 9.2, 9.3, 9.4
' ==========================================================================
Public Function test_suite_sumar_tiempo() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("sumartiempo_zero", test_sumar_zero())
    Call UnitTesting.RunTest("sumartiempo_61", test_sumar_61())
    Call UnitTesting.RunTest("sumartiempo_3600", test_sumar_3600())
    Call UnitTesting.RunTest("sumartiempo_3661", test_sumar_3661())
    
    Debug.Print "SumarTiempo suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_sumar_tiempo = True
End Function

' Verifies SumarTiempo(0) returns "00:00:00".
' Validates: Requirements 9.1
Private Function test_sumar_zero() As Boolean
    On Error GoTo Fail
    
    test_sumar_zero = (SumarTiempo(0) = "00:00:00")
    Exit Function
Fail:
    test_sumar_zero = False
End Function

' Verifies SumarTiempo(61) returns "00:01:01".
' Validates: Requirements 9.2
Private Function test_sumar_61() As Boolean
    On Error GoTo Fail
    
    test_sumar_61 = (SumarTiempo(61) = "00:01:01")
    Exit Function
Fail:
    test_sumar_61 = False
End Function

' Verifies SumarTiempo(3600) returns "01:00:00".
' Validates: Requirements 9.3
Private Function test_sumar_3600() As Boolean
    On Error GoTo Fail
    
    test_sumar_3600 = (SumarTiempo(3600) = "01:00:00")
    Exit Function
Fail:
    test_sumar_3600 = False
End Function

' Verifies SumarTiempo(3661) returns "01:01:01".
' Validates: Requirements 9.4
Private Function test_sumar_3661() As Boolean
    On Error GoTo Fail
    
    test_sumar_3661 = (SumarTiempo(3661) = "01:01:01")
    Exit Function
Fail:
    test_sumar_3661 = False
End Function

#End If
