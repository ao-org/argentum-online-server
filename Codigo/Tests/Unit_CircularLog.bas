Attribute VB_Name = "Unit_CircularLog"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Circular Log Buffer Test Suite
' Tests InitializeCircularLogBuffer, AddLogToCircularBuffer, and
' GetLastMessages from Logging.bas.
' Requirements: 7.1, 7.2, 7.3
' ==========================================================================
Public Function test_suite_circular_log() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("circlog_init", test_circlog_init())
    Call UnitTesting.RunTest("circlog_partial_fill", test_circlog_partial_fill())
    Call UnitTesting.RunTest("circlog_wraparound", test_circlog_wraparound())
    Call UnitTesting.RunTest("circlog_prop_capacity_invariant", test_prop_capacity_invariant())
    
    Debug.Print "CircularLog suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_circular_log = True
End Function

' Verifies InitializeCircularLogBuffer with size 5 sets size=5 and currentIndex=0.
' Validates: Requirements 7.1
Private Function test_circlog_init() As Boolean
    On Error GoTo Fail
    
    Call Logging.InitializeCircularLogBuffer(5)
    
    If CircularLogBuffer.size <> 5 Then
        test_circlog_init = False
        Exit Function
    End If
    
    If CircularLogBuffer.currentIndex <> 0 Then
        test_circlog_init = False
        Exit Function
    End If
    
    test_circlog_init = True
    Exit Function
Fail:
    test_circlog_init = False
End Function

' Verifies that adding fewer messages than buffer size returns those messages
' and empty strings for unused slots via GetLastMessages.
' Validates: Requirements 7.2
Private Function test_circlog_partial_fill() As Boolean
    On Error GoTo Fail
    
    Call Logging.InitializeCircularLogBuffer(5)
    
    Call Logging.AddLogToCircularBuffer("Msg1")
    Call Logging.AddLogToCircularBuffer("Msg2")
    Call Logging.AddLogToCircularBuffer("Msg3")
    
    Dim msgs() As String
    msgs = Logging.GetLastMessages()
    
    ' After adding 3 messages to a size-5 buffer, GetLastMessages returns
    ' indices 1-5 in circular order: two empty slots then the three messages.
    ' i=1: (3+1) Mod 5 = 4 -> ""
    ' i=2: (3+2) Mod 5 = 0 -> ""
    ' i=3: (3+3) Mod 5 = 1 -> "Msg1"
    ' i=4: (3+4) Mod 5 = 2 -> "Msg2"
    ' i=5: (3+5) Mod 5 = 3 -> "Msg3"
    If msgs(1) <> "" Then
        test_circlog_partial_fill = False
        Exit Function
    End If
    
    If msgs(2) <> "" Then
        test_circlog_partial_fill = False
        Exit Function
    End If
    
    If msgs(3) <> "Msg1" Then
        test_circlog_partial_fill = False
        Exit Function
    End If
    
    If msgs(4) <> "Msg2" Then
        test_circlog_partial_fill = False
        Exit Function
    End If
    
    If msgs(5) <> "Msg3" Then
        test_circlog_partial_fill = False
        Exit Function
    End If
    
    test_circlog_partial_fill = True
    Exit Function
Fail:
    test_circlog_partial_fill = False
End Function

' Verifies that adding more messages than buffer size overwrites oldest messages
' and GetLastMessages returns only the most recent messages.
' Validates: Requirements 7.3
Private Function test_circlog_wraparound() As Boolean
    On Error GoTo Fail
    
    Call Logging.InitializeCircularLogBuffer(5)
    
    ' Add 7 messages to a size-5 buffer to force wrap-around
    Call Logging.AddLogToCircularBuffer("Msg1")
    Call Logging.AddLogToCircularBuffer("Msg2")
    Call Logging.AddLogToCircularBuffer("Msg3")
    Call Logging.AddLogToCircularBuffer("Msg4")
    Call Logging.AddLogToCircularBuffer("Msg5")
    Call Logging.AddLogToCircularBuffer("Msg6")
    Call Logging.AddLogToCircularBuffer("Msg7")
    
    Dim msgs() As String
    msgs = Logging.GetLastMessages()
    
    ' After 7 adds, currentIndex=2. GetLastMessages reads in circular order:
    ' i=1: (2+1) Mod 5 = 3 -> "Msg3"
    ' i=2: (2+2) Mod 5 = 4 -> "Msg4"
    ' i=3: (2+3) Mod 5 = 0 -> "Msg5"
    ' i=4: (2+4) Mod 5 = 1 -> "Msg6"
    ' i=5: (2+5) Mod 5 = 2 -> "Msg7"
    If msgs(1) <> "Msg3" Then
        test_circlog_wraparound = False
        Exit Function
    End If
    
    If msgs(2) <> "Msg4" Then
        test_circlog_wraparound = False
        Exit Function
    End If
    
    If msgs(3) <> "Msg5" Then
        test_circlog_wraparound = False
        Exit Function
    End If
    
    If msgs(4) <> "Msg6" Then
        test_circlog_wraparound = False
        Exit Function
    End If
    
    If msgs(5) <> "Msg7" Then
        test_circlog_wraparound = False
        Exit Function
    End If
    
    test_circlog_wraparound = True
    Exit Function
Fail:
    test_circlog_wraparound = False
End Function

' Feature: full-coverage-unit-tests, Property 7
' Property 7: Circular buffer capacity invariant
' For any sequence of N messages added to a buffer of size S where N > S,
' GetLastMessages shall contain exactly the last S messages in order.
' **Validates: Requirements 7.4**
Private Function test_prop_capacity_invariant() As Boolean
    On Error GoTo Fail
    Dim iterations As Long: iterations = 0
    Dim i As Long
    Dim S As Integer: S = 5
    Dim N As Long
    Dim j As Long
    Dim msgs() As String
    Dim expectedMsg As String
    Dim k As Long
    
    For i = 1 To 120
        ' Re-initialize buffer each iteration to reset module-level state
        Call Logging.InitializeCircularLogBuffer(S)
        
        ' N is always > S: ranges from S+2 to S+11
        N = S + (i Mod 10) + 1
        
        ' Add N messages
        For j = 1 To N
            Call Logging.AddLogToCircularBuffer("Iter" & i & "_Msg" & j)
        Next j
        
        ' Get last messages
        msgs = Logging.GetLastMessages()
        
        ' Verify the returned array contains exactly the last S messages in order
        ' The last S messages are from (N - S + 1) to N
        For k = 1 To S
            expectedMsg = "Iter" & i & "_Msg" & (N - S + k)
            If msgs(k) <> expectedMsg Then
                test_prop_capacity_invariant = False
                Exit Function
            End If
        Next k
        
        iterations = iterations + 1
    Next i
    
    test_prop_capacity_invariant = (iterations >= 100)
    Exit Function
Fail:
    test_prop_capacity_invariant = False
End Function

#End If
