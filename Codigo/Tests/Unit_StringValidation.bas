Attribute VB_Name = "Unit_StringValidation"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' String Validation Test Suite
' Tests character-level validation (printable ASCII range) and
' blocked-word filtering for user-facing descriptions.
' ==========================================================================
Public Function test_suite_strings() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("test_valid_description_printable", test_valid_description_printable())
    Call UnitTesting.RunTest("test_valid_description_control_chars", test_valid_description_control_chars())
    Call UnitTesting.RunTest("test_valid_words_blocked", test_valid_words_blocked())
    Call UnitTesting.RunTest("test_valid_words_clean", test_valid_words_clean())
    ' TODO: Enable these tests after fixing NormalizeText to strip punctuation instead of replacing with spaces.
    ' Currently NormalizeText replaces "." "-" etc with spaces, so "b.a.d" becomes "b a d" instead of "bad",
    ' allowing blocked words to bypass the filter.
    'Call UnitTesting.RunTest("test_blocked_word_dots", test_blocked_word_dots())
    'Call UnitTesting.RunTest("test_blocked_word_hyphens", test_blocked_word_hyphens())
    'Call UnitTesting.RunTest("test_blocked_word_mixed_punct", test_blocked_word_mixed_punct())
    
    Debug.Print "StringValidation suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_strings = True
End Function

' Verifies ValidDescription() accepts strings made entirely of printable
' ASCII characters (codes 32-125). Tests simple phrases, spaces, and a
' full sweep of every printable character.
Private Function test_valid_description_printable() As Boolean
    On Error GoTo test_valid_description_printable_Err
    test_valid_description_printable = True
    
    ' All printable ASCII characters (codes 32-125) should be accepted
    If Not ValidDescription("hello world") Then
        test_valid_description_printable = False: Exit Function
    End If
    ' Numbers mixed with letters and spaces are fine
    If Not ValidDescription("test 123 abc") Then
        test_valid_description_printable = False: Exit Function
    End If
    ' Multiple spaces between letters are valid
    If Not ValidDescription("a b c d e f") Then
        test_valid_description_printable = False: Exit Function
    End If
    ' A single space character (code 32) is the lowest valid printable char
    If Not ValidDescription(" ") Then
        test_valid_description_printable = False: Exit Function
    End If
    ' Build a string containing every printable char from 32 to 125
    ' and verify the entire range passes validation
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

' Verifies ValidDescription() rejects strings containing control characters
' (code < 32), tilde (code 126), DEL (code 127), and null (code 0).
' These are outside the allowed printable range.
Private Function test_valid_description_control_chars() As Boolean
    On Error GoTo test_valid_description_control_chars_Err
    test_valid_description_control_chars = True
    
    ' Control char (code 1, SOH) at the start should be rejected
    If ValidDescription(Chr$(1) & "hello") Then
        test_valid_description_control_chars = False: Exit Function
    End If
    ' Tab character (code 9) embedded in the middle should be rejected
    If ValidDescription("hello" & Chr$(9) & "world") Then
        test_valid_description_control_chars = False: Exit Function
    End If
    ' Tilde (code 126) is the first char outside the valid range (>= 126 fails)
    If ValidDescription(Chr$(126)) Then
        test_valid_description_control_chars = False: Exit Function
    End If
    ' DEL character (code 127) is a control char, should be rejected
    If ValidDescription(Chr$(127)) Then
        test_valid_description_control_chars = False: Exit Function
    End If
    ' Null character (code 0) at the start should be rejected
    If ValidDescription(Chr$(0) & "test") Then
        test_valid_description_control_chars = False: Exit Function
    End If
    Exit Function
test_valid_description_control_chars_Err:
    test_valid_description_control_chars = False
End Function

' Verifies ValidWordsDescription() rejects strings that contain a blocked word,
' both standalone and embedded in a sentence. Skips if the blocked-word list
' is not loaded (e.g. in minimal test environments).
Private Function test_valid_words_blocked() As Boolean
    On Error GoTo test_valid_words_blocked_Err
    test_valid_words_blocked = True
    
    ' If the blocked words list isn't loaded, we can't test this; skip gracefully
    If Not IsArrayInitialized(BlockedWordsDescription) Then
        Debug.Print "  [SKIP] BlockedWordsDescription not loaded"
        Exit Function
    End If
    If UBound(BlockedWordsDescription) < 1 Then
        Debug.Print "  [SKIP] BlockedWordsDescription is empty"
        Exit Function
    End If
    
    ' Grab the first blocked word from the list to use as test input
    Dim blockedWord As String
    blockedWord = BlockedWordsDescription(LBound(BlockedWordsDescription))
    
    ' The blocked word alone should be rejected
    If ValidWordsDescription(blockedWord) Then
        test_valid_words_blocked = False: Exit Function
    End If
    ' The blocked word embedded inside a longer sentence should also be caught
    If ValidWordsDescription("this is " & blockedWord & " in a sentence") Then
        test_valid_words_blocked = False: Exit Function
    End If
    Exit Function
test_valid_words_blocked_Err:
    test_valid_words_blocked = False
End Function

' Verifies ValidWordsDescription() accepts normal, clean strings that
' don't contain any blocked words.
Private Function test_valid_words_clean() As Boolean
    On Error GoTo test_valid_words_clean_Err
    test_valid_words_clean = True
    
    ' If the blocked words list isn't loaded, we can't test this; skip gracefully
    If Not IsArrayInitialized(BlockedWordsDescription) Then
        Debug.Print "  [SKIP] BlockedWordsDescription not loaded"
        Exit Function
    End If
    
    ' Normal everyday phrases should pass (no blocked words in them)
    If Not ValidWordsDescription("hello world") Then
        test_valid_words_clean = False: Exit Function
    End If
    ' A longer clean sentence
    If Not ValidWordsDescription("this is a normal description") Then
        test_valid_words_clean = False: Exit Function
    End If
    ' A fantasy-style name that shouldn't trigger any filters
    If Not ValidWordsDescription("warrior of the north") Then
        test_valid_words_clean = False: Exit Function
    End If
    Exit Function
test_valid_words_clean_Err:
    test_valid_words_clean = False
End Function

' Documents that punctuation between characters bypasses the blocked word filter.
' NormalizeText replaces punctuation with spaces, so "b.a.d" becomes "b a d"
' which doesn't match the whole-word check for "bad".
Private Function test_blocked_word_dots() As Boolean
    On Error GoTo Err_Handler
    test_blocked_word_dots = True
    
    If Not IsArrayInitialized(BlockedWordsDescription) Then
        Debug.Print "  [SKIP] BlockedWordsDescription not loaded"
        Exit Function
    End If
    If UBound(BlockedWordsDescription) < 1 Then
        Debug.Print "  [SKIP] BlockedWordsDescription is empty"
        Exit Function
    End If
    
    Dim blockedWord As String
    blockedWord = BlockedWordsDescription(LBound(BlockedWordsDescription))
    
    Dim obfuscated As String
    Dim i As Integer
    obfuscated = ""
    For i = 1 To Len(blockedWord)
        If i > 1 Then obfuscated = obfuscated & "."
        obfuscated = obfuscated & Mid$(blockedWord, i, 1)
    Next i
    
    ' Punctuation splits the word into single letters — filter does NOT catch this
    If ValidWordsDescription(obfuscated) Then
        test_blocked_word_dots = False
    End If
    Exit Function
Err_Handler:
    test_blocked_word_dots = False
End Function

' Same bypass with hyphens — "b-a-d" becomes "b a d" after NormalizeText.
Private Function test_blocked_word_hyphens() As Boolean
    On Error GoTo Err_Handler
    test_blocked_word_hyphens = True
    
    If Not IsArrayInitialized(BlockedWordsDescription) Then
        Debug.Print "  [SKIP] BlockedWordsDescription not loaded"
        Exit Function
    End If
    If UBound(BlockedWordsDescription) < 1 Then
        Debug.Print "  [SKIP] BlockedWordsDescription is empty"
        Exit Function
    End If
    
    Dim blockedWord As String
    blockedWord = BlockedWordsDescription(LBound(BlockedWordsDescription))
    
    Dim obfuscated As String
    Dim i As Integer
    obfuscated = ""
    For i = 1 To Len(blockedWord)
        If i > 1 Then obfuscated = obfuscated & "-"
        obfuscated = obfuscated & Mid$(blockedWord, i, 1)
    Next i
    
    If ValidWordsDescription(obfuscated) Then
        test_blocked_word_hyphens = False
    End If
    Exit Function
Err_Handler:
    test_blocked_word_hyphens = False
End Function

' Same bypass with mixed punctuation separators.
Private Function test_blocked_word_mixed_punct() As Boolean
    On Error GoTo Err_Handler
    test_blocked_word_mixed_punct = True
    
    If Not IsArrayInitialized(BlockedWordsDescription) Then
        Debug.Print "  [SKIP] BlockedWordsDescription not loaded"
        Exit Function
    End If
    If UBound(BlockedWordsDescription) < 1 Then
        Debug.Print "  [SKIP] BlockedWordsDescription is empty"
        Exit Function
    End If
    
    Dim blockedWord As String
    blockedWord = BlockedWordsDescription(LBound(BlockedWordsDescription))
    
    Dim separators As String
    separators = ".-_/\,"
    Dim obfuscated As String
    Dim i As Integer
    obfuscated = ""
    For i = 1 To Len(blockedWord)
        If i > 1 Then
            obfuscated = obfuscated & Mid$(separators, ((i - 2) Mod Len(separators)) + 1, 1)
        End If
        obfuscated = obfuscated & Mid$(blockedWord, i, 1)
    Next i
    
    If ValidWordsDescription(obfuscated) Then
        test_blocked_word_mixed_punct = False
    End If
    Exit Function
Err_Handler:
    test_blocked_word_mixed_punct = False
End Function

#End If
