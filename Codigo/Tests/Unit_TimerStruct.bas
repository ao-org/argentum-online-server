Attribute VB_Name = "Unit_TimerStruct"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Timer Struct Test Suite
' Tests SetTimer, UpdateTime, and GetTimeString from modTime.bas.
' Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6, 6.7, 6.8
' ==========================================================================
Public Function test_suite_timer_struct() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("timer_init", test_timer_init())
    Call UnitTesting.RunTest("timer_fire", test_timer_fire())
    Call UnitTesting.RunTest("timer_no_fire", test_timer_no_fire())
    Call UnitTesting.RunTest("timer_getstr_mins_secs", test_getstr_mins_secs())
    Call UnitTesting.RunTest("timer_getstr_1min", test_getstr_1min())
    Call UnitTesting.RunTest("timer_getstr_1sec", test_getstr_1sec())
    Call UnitTesting.RunTest("timer_getstr_secs_only", test_getstr_secs_only())
    Call UnitTesting.RunTest("timer_getstr_singular_both", test_getstr_singular_both())
    
    Debug.Print "TimerStruct suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_timer_struct = True
End Function

' Verifies SetTimer initializes a t_Timer with ElapsedTime=0, Interval=1000, Occurrences=0.
' Validates: Requirements 6.1
Private Function test_timer_init() As Boolean
    On Error GoTo Fail
    
    Dim t As t_Timer
    Call modTime.SetTimer(t, 1000)
    
    If t.ElapsedTime <> 0 Then
        test_timer_init = False
        Exit Function
    End If
    
    If t.Interval <> 1000 Then
        test_timer_init = False
        Exit Function
    End If
    
    If t.Occurrences <> 0 Then
        test_timer_init = False
        Exit Function
    End If
    
    test_timer_init = True
    Exit Function
Fail:
    test_timer_init = False
End Function

' Verifies UpdateTime returns True and increments Occurrences when delta exceeds interval.
' Validates: Requirements 6.2
Private Function test_timer_fire() As Boolean
    On Error GoTo Fail
    
    Dim t As t_Timer
    Call modTime.SetTimer(t, 1000)
    
    Dim fired As Boolean
    fired = UpdateTime(t, 1500)
    
    If Not fired Then
        test_timer_fire = False
        Exit Function
    End If
    
    If t.Occurrences <> 1 Then
        test_timer_fire = False
        Exit Function
    End If
    
    test_timer_fire = True
    Exit Function
Fail:
    test_timer_fire = False
End Function

' Verifies UpdateTime returns False and Occurrences stays 0 when delta is smaller than interval.
' Validates: Requirements 6.3
Private Function test_timer_no_fire() As Boolean
    On Error GoTo Fail
    
    Dim t As t_Timer
    Call modTime.SetTimer(t, 1000)
    
    Dim fired As Boolean
    fired = UpdateTime(t, 500)
    
    If fired Then
        test_timer_no_fire = False
        Exit Function
    End If
    
    If t.Occurrences <> 0 Then
        test_timer_no_fire = False
        Exit Function
    End If
    
    test_timer_no_fire = True
    Exit Function
Fail:
    test_timer_no_fire = False
End Function

' Verifies GetTimeString(2, 30) returns "2 minutos y 30 segundos".
' Validates: Requirements 6.4
Private Function test_getstr_mins_secs() As Boolean
    On Error GoTo Fail
    
    test_getstr_mins_secs = (GetTimeString(2, 30) = "2 minutos y 30 segundos")
    Exit Function
Fail:
    test_getstr_mins_secs = False
End Function

' Verifies GetTimeString(1, 0) returns "1 minuto".
' Validates: Requirements 6.5
Private Function test_getstr_1min() As Boolean
    On Error GoTo Fail
    
    test_getstr_1min = (GetTimeString(1, 0) = "1 minuto")
    Exit Function
Fail:
    test_getstr_1min = False
End Function

' Verifies GetTimeString(0, 1) returns "1 segundo".
' Validates: Requirements 6.6
Private Function test_getstr_1sec() As Boolean
    On Error GoTo Fail
    
    test_getstr_1sec = (GetTimeString(0, 1) = "1 segundo")
    Exit Function
Fail:
    test_getstr_1sec = False
End Function

' Verifies GetTimeString(0, 45) returns "45 segundos".
' Validates: Requirements 6.7
Private Function test_getstr_secs_only() As Boolean
    On Error GoTo Fail
    
    test_getstr_secs_only = (GetTimeString(0, 45) = "45 segundos")
    Exit Function
Fail:
    test_getstr_secs_only = False
End Function

' Verifies GetTimeString(1, 1) returns "1 minuto y 1 segundo".
' Validates: Requirements 6.8
Private Function test_getstr_singular_both() As Boolean
    On Error GoTo Fail
    
    test_getstr_singular_both = (GetTimeString(1, 1) = "1 minuto y 1 segundo")
    Exit Function
Fail:
    test_getstr_singular_both = False
End Function

#End If
