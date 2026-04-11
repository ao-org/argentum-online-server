Attribute VB_Name = "Unit_Timer"
Option Explicit
#If UNIT_TEST = 1 Then

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

Private Function test_set_timer() As Boolean
    On Error GoTo Err_Handler
    test_set_timer = True
    Dim t As t_Timer
    Call SetTimer(t, 1000)
    If t.ElapsedTime <> 0 Then test_set_timer = False: Exit Function
    If t.Interval <> 1000 Then test_set_timer = False: Exit Function
    If t.Occurrences <> 0 Then test_set_timer = False: Exit Function
    Exit Function
Err_Handler:
    test_set_timer = False
End Function

Private Function test_update_time_not_fired() As Boolean
    On Error GoTo Err_Handler
    test_update_time_not_fired = True
    Dim t As t_Timer
    Call SetTimer(t, 1000)
    ' Add 500ms - should not fire
    If UpdateTime(t, 500) Then test_update_time_not_fired = False: Exit Function
    Exit Function
Err_Handler:
    test_update_time_not_fired = False
End Function

Private Function test_update_time_fired() As Boolean
    On Error GoTo Err_Handler
    test_update_time_fired = True
    Dim t As t_Timer
    Call SetTimer(t, 1000)
    ' Add 1001ms - should fire
    If Not UpdateTime(t, 1001) Then test_update_time_fired = False: Exit Function
    Exit Function
Err_Handler:
    test_update_time_fired = False
End Function

Private Function test_update_time_occurrences() As Boolean
    On Error GoTo Err_Handler
    test_update_time_occurrences = True
    Dim t As t_Timer
    Call SetTimer(t, 100)
    ' Fire 3 times
    Call UpdateTime(t, 101)
    If t.Occurrences <> 1 Then test_update_time_occurrences = False: Exit Function
    Call UpdateTime(t, 101)
    If t.Occurrences <> 2 Then test_update_time_occurrences = False: Exit Function
    Call UpdateTime(t, 101)
    If t.Occurrences <> 3 Then test_update_time_occurrences = False: Exit Function
    Exit Function
Err_Handler:
    test_update_time_occurrences = False
End Function

Private Function test_get_time_string_mins_and_secs() As Boolean
    On Error GoTo Err_Handler
    test_get_time_string_mins_and_secs = True
    Dim result As String
    result = GetTimeString(2, 30)
    If InStr(result, "2 minutos") = 0 Then test_get_time_string_mins_and_secs = False: Exit Function
    If InStr(result, "30 segundos") = 0 Then test_get_time_string_mins_and_secs = False: Exit Function
    If InStr(result, " y ") = 0 Then test_get_time_string_mins_and_secs = False: Exit Function
    Exit Function
Err_Handler:
    test_get_time_string_mins_and_secs = False
End Function

Private Function test_get_time_string_only_seconds() As Boolean
    On Error GoTo Err_Handler
    test_get_time_string_only_seconds = True
    Dim result As String
    result = GetTimeString(0, 45)
    If InStr(result, "45 segundos") = 0 Then test_get_time_string_only_seconds = False: Exit Function
    If InStr(result, "minuto") > 0 Then test_get_time_string_only_seconds = False: Exit Function
    Exit Function
Err_Handler:
    test_get_time_string_only_seconds = False
End Function

Private Function test_get_time_string_only_minutes() As Boolean
    On Error GoTo Err_Handler
    test_get_time_string_only_minutes = True
    Dim result As String
    result = GetTimeString(5, 0)
    If InStr(result, "5 minutos") = 0 Then test_get_time_string_only_minutes = False: Exit Function
    If InStr(result, "segundo") > 0 Then test_get_time_string_only_minutes = False: Exit Function
    Exit Function
Err_Handler:
    test_get_time_string_only_minutes = False
End Function

#End If
