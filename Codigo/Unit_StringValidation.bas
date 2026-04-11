Attribute VB_Name = "Unit_StringValidation"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_strings() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("test_valid_description_printable", test_valid_description_printable())
    Call UnitTesting.RunTest("test_valid_description_control_chars", test_valid_description_control_chars())
    Call UnitTesting.RunTest("test_valid_words_blocked", test_valid_words_blocked())
    Call UnitTesting.RunTest("test_valid_words_clean", test_valid_words_clean())
    
    Debug.Print "StringValidation suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_strings = True
End Function

Private Function test_valid_description_printable() As Boolean
    On Error GoTo test_valid_description_printable_Err
    test_valid_description_printable = True
    
    ' All printable ASCII (32-125) should pass
    If Not ValidDescription("hello world") Then
        test_valid_description_printable = False: Exit Function
    End If
    If Not ValidDescription("test 123 abc") Then
        test_valid_description_printable = False: Exit Function
    End If
    If Not ValidDescription("a b c d e f") Then
        test_valid_description_printable = False: Exit Function
    End If
    ' Space (32) is valid
    If Not ValidDescription(" ") Then
        test_valid_description_printable = False: Exit Function
    End If
    ' Build a string of all printable chars 32-125
    Dim s As String
    Dim i As Integer
    s = ""
    For i = 32 To 125
        s = s & Chr$(i)
    Next i
    If Not ValidDescription(s) Then
        test_valid_description_printable = False: Exit Function
    End If
    Exit Function
test_valid_description_printable_Err:
    test_valid_description_printable = False
End Function

Private Function test_valid_description_control_chars() As Boolean
    On Error GoTo test_valid_description_control_chars_Err
    test_valid_description_control_chars = True
    
    ' Control char (code 1) should fail
    If ValidDescription(Chr$(1) & "hello") Then
        test_valid_description_control_chars = False: Exit Function
    End If
    ' Tab (code 9) should fail
    If ValidDescription("hello" & Chr$(9) & "world") Then
        test_valid_description_control_chars = False: Exit Function
    End If
    ' Char 126 (tilde ~) should fail per the implementation (car >= 126)
    If ValidDescription(Chr$(126)) Then
        test_valid_description_control_chars = False: Exit Function
    End If
    ' Char 127 (DEL) should fail
    If ValidDescription(Chr$(127)) Then
        test_valid_description_control_chars = False: Exit Function
    End If
    ' Null char should fail
    If ValidDescription(Chr$(0) & "test") Then
        test_valid_description_control_chars = False: Exit Function
    End If
    Exit Function
test_valid_description_control_chars_Err:
    test_valid_description_control_chars = False
End Function

Private Function test_valid_words_blocked() As Boolean
    On Error GoTo test_valid_words_blocked_Err
    test_valid_words_blocked = True
    
    ' Ensure BlockedWordsDescription is loaded
    ' If array is not initialized, skip this test (return True)
    If Not IsArrayInitialized(BlockedWordsDescription) Then
        Debug.Print "  [SKIP] BlockedWordsDescription not loaded"
        Exit Function
    End If
    If UBound(BlockedWordsDescription) < 1 Then
        Debug.Print "  [SKIP] BlockedWordsDescription is empty"
        Exit Function
    End If
    
    ' Test with the first blocked word
    Dim blockedWord As String
    blockedWord = BlockedWordsDescription(LBound(BlockedWordsDescription))
    
    If ValidWordsDescription(blockedWord) Then
        test_valid_words_blocked = False: Exit Function
    End If
    ' Blocked word embedded in a sentence
    If ValidWordsDescription("this is " & blockedWord & " in a sentence") Then
        test_valid_words_blocked = False: Exit Function
    End If
    Exit Function
test_valid_words_blocked_Err:
    test_valid_words_blocked = False
End Function

Private Function test_valid_words_clean() As Boolean
    On Error GoTo test_valid_words_clean_Err
    test_valid_words_clean = True
    
    ' Ensure BlockedWordsDescription is loaded
    If Not IsArrayInitialized(BlockedWordsDescription) Then
        Debug.Print "  [SKIP] BlockedWordsDescription not loaded"
        Exit Function
    End If
    
    ' Clean strings should pass
    If Not ValidWordsDescription("hello world") Then
        test_valid_words_clean = False: Exit Function
    End If
    If Not ValidWordsDescription("this is a normal description") Then
        test_valid_words_clean = False: Exit Function
    End If
    If Not ValidWordsDescription("warrior of the north") Then
        test_valid_words_clean = False: Exit Function
    End If
    Exit Function
test_valid_words_clean_Err:
    test_valid_words_clean = False
End Function

#End If
