Attribute VB_Name = "Unit_Uptime"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Uptime Test Suite
' Tests FormatUptime from modUptime.bas — verifies that seconds are correctly
' formatted as "Xd HH:MM:SS".  Each test manipulates g_ServerStart via
' DateAdd to simulate a specific uptime, then restores the original value.
' ==========================================================================
Public Function test_suite_uptime() As Boolean
    ' Example-based tests (Req 9.1 - 9.5)
    Call UnitTesting.RunTest("test_uptime_zero_seconds", test_uptime_zero_seconds())
    Call UnitTesting.RunTest("test_uptime_3661_seconds", test_uptime_3661_seconds())
    Call UnitTesting.RunTest("test_uptime_90061_seconds", test_uptime_90061_seconds())
    Call UnitTesting.RunTest("test_uptime_86399_seconds", test_uptime_86399_seconds())

    test_suite_uptime = True
End Function

' Req 9.1: 0 seconds -> "0d 00:00:00"
Private Function test_uptime_zero_seconds() As Boolean
    On Error GoTo Fail
    Dim origStart As Date
    origStart = g_ServerStart

    g_ServerStart = Now

    Dim result As String
    result = FormatUptime()

    g_ServerStart = origStart
    test_uptime_zero_seconds = (result = "0d 00:00:00")
    Exit Function
Fail:
    g_ServerStart = origStart
    test_uptime_zero_seconds = False
End Function

' Req 9.2: 3661 seconds (1h 1m 1s) -> "0d 01:01:01"
Private Function test_uptime_3661_seconds() As Boolean
    On Error GoTo Fail
    Dim origStart As Date
    origStart = g_ServerStart

    g_ServerStart = DateAdd("s", -3661, Now)

    Dim result As String
    result = FormatUptime()

    g_ServerStart = origStart
    test_uptime_3661_seconds = (result = "0d 01:01:01")
    Exit Function
Fail:
    g_ServerStart = origStart
    test_uptime_3661_seconds = False
End Function

' Req 9.3: 90061 seconds (1d 1h 1m 1s) -> "1d 01:01:01"
Private Function test_uptime_90061_seconds() As Boolean
    On Error GoTo Fail
    Dim origStart As Date
    origStart = g_ServerStart

    g_ServerStart = DateAdd("s", -90061, Now)

    Dim result As String
    result = FormatUptime()

    g_ServerStart = origStart
    test_uptime_90061_seconds = (result = "1d 01:01:01")
    Exit Function
Fail:
    g_ServerStart = origStart
    test_uptime_90061_seconds = False
End Function

' Req 9.4: 86399 seconds (23h 59m 59s) -> "0d 23:59:59"
Private Function test_uptime_86399_seconds() As Boolean
    On Error GoTo Fail
    Dim origStart As Date
    origStart = g_ServerStart

    g_ServerStart = DateAdd("s", -86399, Now)

    Dim result As String
    result = FormatUptime()

    g_ServerStart = origStart
    test_uptime_86399_seconds = (result = "0d 23:59:59")
    Exit Function
Fail:
    g_ServerStart = origStart
    test_uptime_86399_seconds = False
End Function

#End If
