Attribute VB_Name = "Unit_WorldTime"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_world_time() As Boolean
    Call UnitTesting.RunTest("test_init_positive_daylenms", test_init_positive_daylenms())
    Call UnitTesting.RunTest("test_init_zero_clamps_to_one", test_init_zero_clamps_to_one())
    Call UnitTesting.RunTest("test_init_negative_clamps_to_one", test_init_negative_clamps_to_one())
    Call UnitTesting.RunTest("test_handlehora_range", test_handlehora_range())
    Call UnitTesting.RunTest("test_preparehora_output_values", test_preparehora_output_values())
    Call UnitTesting.RunTest("test_prop_ms_range_invariant", test_prop_ms_range_invariant())
    Call UnitTesting.RunTest("test_prop_sec_consistency", test_prop_sec_consistency())
    Call UnitTesting.RunTest("test_prop_daylenms_roundtrip", test_prop_daylenms_roundtrip())
    
    test_suite_world_time = True
End Function

' Req 7.1: Init with positive dayLenMs returns WorldTime_Ms in [0, dayLenMs - 1].
Private Function test_init_positive_daylenms() As Boolean
    On Error GoTo Err_Handler
    
    Dim dayLen As Long
    dayLen = 60000  ' 60 seconds
    
    Call WorldTime_Init(dayLen, 0)
    
    Dim ms As Long
    ms = WorldTime_Ms()
    
    test_init_positive_daylenms = (ms >= 0 And ms < dayLen)
    Exit Function
Err_Handler:
    test_init_positive_daylenms = False
End Function

' Req 7.2: Init with zero dayLenMs clamps to 1.
Private Function test_init_zero_clamps_to_one() As Boolean
    On Error GoTo Err_Handler
    
    Call WorldTime_Init(0, 0)
    
    ' After clamping to 1, DayLenMs should be 1
    Dim dayLen As Long
    dayLen = WorldTime_DayLenMs()
    
    ' Ms must be in [0, 0] i.e. exactly 0
    Dim ms As Long
    ms = WorldTime_Ms()
    
    test_init_zero_clamps_to_one = (dayLen = 1 And ms = 0)
    Exit Function
Err_Handler:
    test_init_zero_clamps_to_one = False
End Function

' Req 7.2: Init with negative dayLenMs clamps to 1.
Private Function test_init_negative_clamps_to_one() As Boolean
    On Error GoTo Err_Handler
    
    Call WorldTime_Init(-500, 0)
    
    Dim dayLen As Long
    dayLen = WorldTime_DayLenMs()
    
    Dim ms As Long
    ms = WorldTime_Ms()
    
    test_init_negative_clamps_to_one = (dayLen = 1 And ms = 0)
    Exit Function
Err_Handler:
    test_init_negative_clamps_to_one = False
End Function

' Req 7.6: HandleHora returns WorldTime_Ms in [0, dayLenMs - 1].
Private Function test_handlehora_range() As Boolean
    On Error GoTo Err_Handler
    
    Dim dayLen As Long
    dayLen = 120000  ' 2 minutes
    
    ' Simulate receiving a server hora packet with elapsed = 30000 ms
    Call WorldTime_HandleHora(30000, dayLen)
    
    Dim ms As Long
    ms = WorldTime_Ms()
    
    test_handlehora_range = (ms >= 0 And ms < dayLen)
    Exit Function
Err_Handler:
    test_handlehora_range = False
End Function

' Req 7.7: PrepareHora returns outElapsedMs in [0, dayLenMs - 1] and outDayLenMs = configured day length.
Private Function test_preparehora_output_values() As Boolean
    On Error GoTo Err_Handler
    
    Dim dayLen As Long
    dayLen = 90000  ' 90 seconds
    
    Call WorldTime_Init(dayLen, 5000)
    
    Dim outElapsed As Long
    Dim outDayLen As Long
    Call WorldTime_PrepareHora(outElapsed, outDayLen)
    
    test_preparehora_output_values = (outElapsed >= 0 And outElapsed < dayLen And outDayLen = dayLen)
    Exit Function
Err_Handler:
    test_preparehora_output_values = False
End Function

' Feature: unit-test-coverage, Property 7: WorldTime millisecond range invariant
Private Function test_prop_ms_range_invariant() As Boolean
    On Error GoTo Err_Handler
    
    Dim i As Long
    Dim dayLen As Long
    Dim ms As Long
    Dim outElapsed As Long
    Dim outDayLen As Long
    
    ' Loop over 120 positive dayLenMs values
    For i = 1 To 120
        dayLen = i * 500  ' values: 500, 1000, 1500, ... 60000
        
        ' Test via Init
        Call WorldTime_Init(dayLen, CLng(i * 137))  ' varying start offsets
        ms = WorldTime_Ms()
        If ms < 0 Or ms >= dayLen Then
            test_prop_ms_range_invariant = False
            Exit Function
        End If
        
        ' Test via HandleHora
        Call WorldTime_HandleHora(CLng(i * 251), dayLen)
        ms = WorldTime_Ms()
        If ms < 0 Or ms >= dayLen Then
            test_prop_ms_range_invariant = False
            Exit Function
        End If
        
        ' Test via PrepareHora
        Call WorldTime_Init(dayLen, CLng(i * 73))
        Call WorldTime_PrepareHora(outElapsed, outDayLen)
        If outElapsed < 0 Or outElapsed >= dayLen Then
            test_prop_ms_range_invariant = False
            Exit Function
        End If
    Next i
    
    test_prop_ms_range_invariant = True
    Exit Function
Err_Handler:
    test_prop_ms_range_invariant = False
End Function

' Feature: unit-test-coverage, Property 8: WorldTime second consistency invariant
Private Function test_prop_sec_consistency() As Boolean
    On Error GoTo Err_Handler
    
    Dim i As Long
    Dim dayLen As Long
    Dim ms As Long
    Dim sec As Long
    
    ' Loop over 120 WorldTime states with varying dayLen and offsets
    For i = 1 To 120
        dayLen = 1000 + i * 500  ' values: 1500, 2000, ... 61500
        
        Call WorldTime_Init(dayLen, CLng(i * 311))
        
        ms = WorldTime_Ms()
        sec = WorldTime_Sec()
        
        If sec <> (ms \ 1000) Then
            test_prop_sec_consistency = False
            Exit Function
        End If
    Next i
    
    test_prop_sec_consistency = True
    Exit Function
Err_Handler:
    test_prop_sec_consistency = False
End Function

' Feature: unit-test-coverage, Property 9: WorldTime DayLenMs get/set round-trip
Private Function test_prop_daylenms_roundtrip() As Boolean
    On Error GoTo Err_Handler
    
    Dim i As Long
    Dim testVal As Long
    Dim result As Long
    
    ' Loop over 120 positive Long values
    For i = 1 To 120
        testVal = i * 1000  ' values: 1000, 2000, ... 120000
        
        Call WorldTime_SetDayLenMs(testVal)
        result = WorldTime_DayLenMs()
        
        If result <> testVal Then
            test_prop_daylenms_roundtrip = False
            Exit Function
        End If
    Next i
    
    ' Also test some larger values
    Dim largeVals(4) As Long
    largeVals(0) = 1
    largeVals(1) = 86400000  ' 24 hours in ms
    largeVals(2) = 3600000   ' 1 hour in ms
    largeVals(3) = 2147483647 ' max Long
    largeVals(4) = 42
    
    Dim j As Long
    For j = 0 To 4
        Call WorldTime_SetDayLenMs(largeVals(j))
        result = WorldTime_DayLenMs()
        
        If result <> largeVals(j) Then
            test_prop_daylenms_roundtrip = False
            Exit Function
        End If
    Next j
    
    test_prop_daylenms_roundtrip = True
    Exit Function
Err_Handler:
    test_prop_daylenms_roundtrip = False
End Function

#End If
