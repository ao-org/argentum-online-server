Attribute VB_Name = "Unit_Timer"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Timer Test Suite
' Tests the game timer system: initialization, fire/no-fire thresholds,
' occurrence counting, and human-readable time string formatting.
' ==========================================================================
Public Function test_suite_timer() As Boolean
    Call UnitTesting.RunTest("test_set_timer", test_set_timer())
    Call UnitTesting.RunTest("test_update_time_not_fired", test_update_time_not_fired())
    Call UnitTesting.RunTest("test_update_time_fired", test_update_time_fired())
    Call UnitTesting.RunTest("test_update_time_occurrences", test_update_time_occurrences())
    Call UnitTesting.RunTest("test_get_time_string_mins_and_secs", test_get_time_string_mins_and_secs())
    Call UnitTesting.RunTest("test_get_time_string_only_seconds", test_get_time_string_only_seconds())
    Call UnitTesting.RunTest("test_get_time_string_only_minutes", test_get_time_string_only_minutes())
    test_suite_timer = True
End Function

' Verifies SetTimer() initializes a timer with zero elapsed time,
' the given interval, and zero occurrences.
Private Function test_set_timer() As Boolean
    On Error GoTo Err_Handler
    test_set_timer = True
    Dim t As t_Timer
    ' Create a timer with a 1000ms interval
    Call SetTimer(t, 1000)
    ' ElapsedTime should start at 0 (no time has passed yet)
    If t.ElapsedTime <> 0 Then test_set_timer = False: Exit Function
    ' Interval should be set to what we passed in
    If t.Interval <> 1000 Then test_set_timer = False: Exit Function
    ' No fires yet, so Occurrences should be 0
    If t.Occurrences <> 0 Then test_set_timer = False: Exit Function
    Exit Function
Err_Handler:
    test_set_timer = False
End Function

' Verifies UpdateTime() does NOT fire when elapsed time (500ms) is less
' than the interval (1000ms).
Private Function test_update_time_not_fired() As Boolean
    On Error GoTo Err_Handler
    test_update_time_not_fired = True
    Dim t As t_Timer
    Call SetTimer(t, 1000)
    ' Only 500ms have passed, which is less than the 1000ms interval
    ' UpdateTime should return False (timer hasn't fired)
    If UpdateTime(t, 500) Then test_update_time_not_fired = False: Exit Function
    Exit Function
Err_Handler:
    test_update_time_not_fired = False
End Function

' Verifies UpdateTime() fires (returns True) when elapsed time (1001ms)
' exceeds the interval (1000ms).
Private Function test_update_time_fired() As Boolean
    On Error GoTo Err_Handler
    test_update_time_fired = True
    Dim t As t_Timer
    Call SetTimer(t, 1000)
    ' 1001ms have passed, which exceeds the 1000ms interval
    ' UpdateTime should return True (timer fired)
    If Not UpdateTime(t, 1001) Then test_update_time_fired = False: Exit Function
    Exit Function
Err_Handler:
    test_update_time_fired = False
End Function

' Verifies that each time UpdateTime() fires, the Occurrences counter
' increments by 1 (100ms interval, fired 3 times -> Occurrences = 3).
Private Function test_update_time_occurrences() As Boolean
    On Error GoTo Err_Handler
    test_update_time_occurrences = True
    Dim t As t_Timer
    ' Set a fast 100ms interval timer
    Call SetTimer(t, 100)
    ' First fire: 101ms elapsed, exceeds 100ms interval -> fires, Occurrences = 1
    Call UpdateTime(t, 101)
    If t.Occurrences <> 1 Then test_update_time_occurrences = False: Exit Function
    ' Second fire: another 101ms -> Occurrences = 2
    Call UpdateTime(t, 101)
    If t.Occurrences <> 2 Then test_update_time_occurrences = False: Exit Function
    ' Third fire: another 101ms -> Occurrences = 3
    Call UpdateTime(t, 101)
    If t.Occurrences <> 3 Then test_update_time_occurrences = False: Exit Function
    Exit Function
Err_Handler:
    test_update_time_occurrences = False
End Function

' Verifies GetTimeString() formats both minutes and seconds together.
' 2 min 30 sec should produce a string containing "2 minutos", "30 segundos",
' and the conjunction " y " between them.
Private Function test_get_time_string_mins_and_secs() As Boolean
    On Error GoTo Err_Handler
    test_get_time_string_mins_and_secs = True
    Dim result As String
    ' Format 2 minutes and 30 seconds
    result = GetTimeString(2, 30)
    ' Should contain the minutes part in Spanish
    If InStr(result, "2 minutos") = 0 Then test_get_time_string_mins_and_secs = False: Exit Function
    ' Should contain the seconds part in Spanish
    If InStr(result, "30 segundos") = 0 Then test_get_time_string_mins_and_secs = False: Exit Function
    ' Should contain " y " (and) joining the two parts
    If InStr(result, " y ") = 0 Then test_get_time_string_mins_and_secs = False: Exit Function
    Exit Function
Err_Handler:
    test_get_time_string_mins_and_secs = False
End Function

' Verifies GetTimeString() with 0 minutes shows only seconds,
' with no mention of "minuto" at all.
Private Function test_get_time_string_only_seconds() As Boolean
    On Error GoTo Err_Handler
    test_get_time_string_only_seconds = True
    Dim result As String
    ' Format 0 minutes and 45 seconds (seconds only)
    result = GetTimeString(0, 45)
    ' Should show the seconds
    If InStr(result, "45 segundos") = 0 Then test_get_time_string_only_seconds = False: Exit Function
    ' Should NOT mention minutes at all when minutes = 0
    If InStr(result, "minuto") > 0 Then test_get_time_string_only_seconds = False: Exit Function
    Exit Function
Err_Handler:
    test_get_time_string_only_seconds = False
End Function

' Verifies GetTimeString() with 0 seconds shows only minutes,
' with no mention of "segundo" at all.
Private Function test_get_time_string_only_minutes() As Boolean
    On Error GoTo Err_Handler
    test_get_time_string_only_minutes = True
    Dim result As String
    ' Format 5 minutes and 0 seconds (minutes only)
    result = GetTimeString(5, 0)
    ' Should show the minutes
    If InStr(result, "5 minutos") = 0 Then test_get_time_string_only_minutes = False: Exit Function
    ' Should NOT mention seconds at all when seconds = 0
    If InStr(result, "segundo") > 0 Then test_get_time_string_only_minutes = False: Exit Function
    Exit Function
Err_Handler:
    test_get_time_string_only_minutes = False
End Function

#End If
