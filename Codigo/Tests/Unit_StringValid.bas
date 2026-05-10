Attribute VB_Name = "Unit_StringValid"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' String Validation Test Suite
' Tests ValidDescription from StringUtils.bas: printable ASCII acceptance,
' control character rejection, high-byte rejection, and empty string handling.
' ==========================================================================
Public Function test_suite_string_valid() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("strvalid_printable", test_printable_accepted())
    Call UnitTesting.RunTest("strvalid_control_char", test_control_char_rejected())
    Call UnitTesting.RunTest("strvalid_high_byte", test_high_byte_rejected())
    Call UnitTesting.RunTest("strvalid_empty", test_empty_accepted())
    Call UnitTesting.RunTest("strvalid_prop_printable_accepted", test_prop_printable_accepted())
    
    Debug.Print "StringValid suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_string_valid = True
End Function

' Req 8.1: ValidDescription accepts strings with only printable ASCII (codes 32-125).
Private Function test_printable_accepted() As Boolean
    On Error GoTo Fail
    
    ' Simple alphabetic string
    If Not ValidDescription("hello world") Then
        test_printable_accepted = False
        Exit Function
    End If
    
    ' String with numbers and symbols within printable range
    If Not ValidDescription("item #42 costs $5") Then
        test_printable_accepted = False
        Exit Function
    End If
    
    ' Full printable range sweep: every char from 32 to 125
    Dim s As String
    Dim i As Integer
    s = ""
    For i = 32 To 125
        s = s & Chr$(i)
    Next i
    If Not ValidDescription(s) Then
        test_printable_accepted = False
        Exit Function
    End If
    
    test_printable_accepted = True
    Exit Function
Fail:
    test_printable_accepted = False
End Function

' Req 8.2: ValidDescription rejects strings containing control characters (code < 32).
Private Function test_control_char_rejected() As Boolean
    On Error GoTo Fail
    
    ' Null character (code 0)
    If ValidDescription(Chr$(0) & "test") Then
        test_control_char_rejected = False
        Exit Function
    End If
    
    ' Tab character (code 9)
    If ValidDescription("hello" & Chr$(9) & "world") Then
        test_control_char_rejected = False
        Exit Function
    End If
    
    ' Newline (code 10)
    If ValidDescription("line1" & Chr$(10) & "line2") Then
        test_control_char_rejected = False
        Exit Function
    End If
    
    ' Unit separator (code 31, highest control char)
    If ValidDescription("abc" & Chr$(31) & "def") Then
        test_control_char_rejected = False
        Exit Function
    End If
    
    test_control_char_rejected = True
    Exit Function
Fail:
    test_control_char_rejected = False
End Function

' Req 8.3: ValidDescription rejects strings containing characters with code >= 126.
Private Function test_high_byte_rejected() As Boolean
    On Error GoTo Fail
    
    ' Tilde (code 126) - first char outside valid range
    If ValidDescription(Chr$(126)) Then
        test_high_byte_rejected = False
        Exit Function
    End If
    
    ' DEL (code 127)
    If ValidDescription("test" & Chr$(127)) Then
        test_high_byte_rejected = False
        Exit Function
    End If
    
    ' High byte (code 200)
    If ValidDescription("abc" & Chr$(200) & "def") Then
        test_high_byte_rejected = False
        Exit Function
    End If
    
    ' Code 255
    If ValidDescription(Chr$(255)) Then
        test_high_byte_rejected = False
        Exit Function
    End If
    
    test_high_byte_rejected = True
    Exit Function
Fail:
    test_high_byte_rejected = False
End Function

' Req 8.4: ValidDescription accepts empty strings.
Private Function test_empty_accepted() As Boolean
    On Error GoTo Fail
    
    test_empty_accepted = ValidDescription("")
    Exit Function
Fail:
    test_empty_accepted = False
End Function

' Feature: full-coverage-unit-tests, Property 8: ValidDescription accepts all printable ASCII strings
' Validates: Requirements 8.5
Private Function test_prop_printable_accepted() As Boolean
    On Error GoTo Fail
    Dim iterations As Long
    Dim i As Long
    Dim s As String
    Dim strLen As Long
    Dim j As Long
    Dim charCode As Integer
    iterations = 0
    For i = 1 To 120
        s = ""
        strLen = (i Mod 20) + 1  ' 1 to 20 chars
        For j = 1 To strLen
            ' Generate char code in [32, 125] range deterministically
            charCode = 32 + ((i * 7 + j * 13) Mod 94)
            s = s & Chr$(charCode)
        Next j
        If Not ValidDescription(s) Then
            test_prop_printable_accepted = False
            Exit Function
        End If
        iterations = iterations + 1
    Next i
    test_prop_printable_accepted = (iterations >= 100)
    Exit Function
Fail:
    test_prop_printable_accepted = False
End Function

#End If
