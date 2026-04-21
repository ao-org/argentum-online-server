Attribute VB_Name = "Unit_TimeFormat"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Time Format Test Suite
' Tests GetTimeString Spanish formatting (singular/plural, " y " joiner,
' zero-omission) and SumarTiempo "hh:mm:ss" output.
' ==========================================================================
Public Function test_suite_time_format() As Boolean
    ' Example-based tests (Req 8.1 - 8.6)
    Call UnitTesting.RunTest("test_both_mins_and_secs", test_both_mins_and_secs())
    Call UnitTesting.RunTest("test_zero_minutes_omits_minuto", test_zero_minutes_omits_minuto())
    Call UnitTesting.RunTest("test_zero_seconds_omits_segundo", test_zero_seconds_omits_segundo())
    Call UnitTesting.RunTest("test_singular_one_minute", test_singular_one_minute())
    Call UnitTesting.RunTest("test_singular_one_second", test_singular_one_second())
    Call UnitTesting.RunTest("test_sumartiempo_format", test_sumartiempo_format())
    
    ' Property test (Property 10)
    Call UnitTesting.RunTest("test_prop_gettimestring_format", test_prop_gettimestring_format())
    
    test_suite_time_format = True
End Function

' Req 8.1: Both minutes and seconds > 0 joined by " y ".
Private Function test_both_mins_and_secs() As Boolean
    On Error GoTo Err_Handler
    Dim result As String
    result = GetTimeString(3, 15)
    
    ' Must contain both parts and the joiner
    If InStr(result, "3 minutos") = 0 Then test_both_mins_and_secs = False: Exit Function
    If InStr(result, "15 segundos") = 0 Then test_both_mins_and_secs = False: Exit Function
    If InStr(result, " y ") = 0 Then test_both_mins_and_secs = False: Exit Function
    
    test_both_mins_and_secs = True
    Exit Function
Err_Handler:
    test_both_mins_and_secs = False
End Function

' Req 8.2: Zero minutes omits "minuto".
Private Function test_zero_minutes_omits_minuto() As Boolean
    On Error GoTo Err_Handler
    Dim result As String
    result = GetTimeString(0, 20)
    
    If InStr(result, "20 segundos") = 0 Then test_zero_minutes_omits_minuto = False: Exit Function
    If InStr(result, "minuto") > 0 Then test_zero_minutes_omits_minuto = False: Exit Function
    
    test_zero_minutes_omits_minuto = True
    Exit Function
Err_Handler:
    test_zero_minutes_omits_minuto = False
End Function

' Req 8.3: Zero seconds omits "segundo".
Private Function test_zero_seconds_omits_segundo() As Boolean
    On Error GoTo Err_Handler
    Dim result As String
    result = GetTimeString(4, 0)
    
    If InStr(result, "4 minutos") = 0 Then test_zero_seconds_omits_segundo = False: Exit Function
    If InStr(result, "segundo") > 0 Then test_zero_seconds_omits_segundo = False: Exit Function
    
    test_zero_seconds_omits_segundo = True
    Exit Function
Err_Handler:
    test_zero_seconds_omits_segundo = False
End Function

' Req 8.4: Singular form for 1 minute.
Private Function test_singular_one_minute() As Boolean
    On Error GoTo Err_Handler
    Dim result As String
    result = GetTimeString(1, 10)
    
    ' Must use singular "minuto", not "minutos"
    If InStr(result, "1 minuto") = 0 Then test_singular_one_minute = False: Exit Function
    If InStr(result, "minutos") > 0 Then test_singular_one_minute = False: Exit Function
    
    test_singular_one_minute = True
    Exit Function
Err_Handler:
    test_singular_one_minute = False
End Function

' Req 8.5: Singular form for 1 second.
Private Function test_singular_one_second() As Boolean
    On Error GoTo Err_Handler
    Dim result As String
    result = GetTimeString(2, 1)
    
    ' Must use singular "segundo", not "segundos"
    If InStr(result, "1 segundo") = 0 Then test_singular_one_second = False: Exit Function
    If InStr(result, "segundos") > 0 Then test_singular_one_second = False: Exit Function
    
    test_singular_one_second = True
    Exit Function
Err_Handler:
    test_singular_one_second = False
End Function

' Req 8.6: SumarTiempo returns "hh:mm:ss" format.
Private Function test_sumartiempo_format() As Boolean
    On Error GoTo Err_Handler
    Dim result As String
    result = SumarTiempo(3661)  ' 1 hour, 1 minute, 1 second
    
    ' Should be "01:01:01"
    test_sumartiempo_format = (result = "01:01:01")
    Exit Function
Err_Handler:
    test_sumartiempo_format = False
End Function

' Feature: unit-test-coverage, Property 10: GetTimeString format correctness
Private Function test_prop_gettimestring_format() As Boolean
    On Error GoTo Err_Handler
    
    Dim m As Integer
    Dim s As Integer
    Dim result As String
    Dim iterations As Long
    iterations = 0
    
    ' Loop over a grid of (minutes, seconds) pairs covering all three cases.
    ' minutes: 0..12, seconds: 0..12 -> 169 pairs (>100)
    For m = 0 To 12
        For s = 0 To 12
            result = GetTimeString(m, s)
            
            ' Case 1: both > 0 -> " y " must be present
            If m > 0 And s > 0 Then
                If InStr(result, " y ") = 0 Then
                    test_prop_gettimestring_format = False
                    Exit Function
                End If
            End If
            
            ' Case 2: minutes = 0 -> no "minuto" anywhere
            If m = 0 Then
                If InStr(result, "minuto") > 0 Then
                    test_prop_gettimestring_format = False
                    Exit Function
                End If
            End If
            
            ' Case 3: seconds = 0 -> no "segundo" anywhere
            If s = 0 Then
                If InStr(result, "segundo") > 0 Then
                    test_prop_gettimestring_format = False
                    Exit Function
                End If
            End If
            
            iterations = iterations + 1
        Next s
    Next m
    
    ' Verify we ran at least 100 iterations
    test_prop_gettimestring_format = (iterations >= 100)
    Exit Function
Err_Handler:
    test_prop_gettimestring_format = False
End Function

#End If
