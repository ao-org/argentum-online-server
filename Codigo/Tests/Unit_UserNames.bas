Attribute VB_Name = "Unit_UserNames"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' UserNames Test Suite
' Tests ValidateAlias() which validates player alias strings: empty/blank
' handling, valid names, length limits, control characters, and high-byte chars.
' ==========================================================================
Public Function test_suite_usernames() As Boolean
    Call UnitTesting.RunTest("test_validate_alias_empty", test_validate_alias_empty())
    Call UnitTesting.RunTest("test_validate_alias_valid", test_validate_alias_valid())
    Call UnitTesting.RunTest("test_validate_alias_too_long", test_validate_alias_too_long())
    Call UnitTesting.RunTest("test_validate_alias_control_chars", test_validate_alias_control_chars())
    Call UnitTesting.RunTest("test_validate_alias_chr255", test_validate_alias_chr255())
    test_suite_usernames = True
End Function

' Verifies that empty or whitespace-only aliases are accepted as valid
' (they mean "no alias set").
Private Function test_validate_alias_empty() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_empty = True
    Dim errMsg As String
    ' Empty string means "no alias set", which is valid
    If Not ValidateAlias("", errMsg) Then test_validate_alias_empty = False: Exit Function
    ' Whitespace-only is also treated as "no alias"
    If Not ValidateAlias("   ", errMsg) Then test_validate_alias_empty = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_empty = False
End Function

' Verifies that normal aliases with letters, numbers, and spaces pass validation.
Private Function test_validate_alias_valid() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_valid = True
    Dim errMsg As String
    ' Simple single-word alias with letters only
    If Not ValidateAlias("Warrior", errMsg) Then test_validate_alias_valid = False: Exit Function
    ' Alias with letters, spaces, and numbers (common in-game name)
    If Not ValidateAlias("Dark Knight 99", errMsg) Then test_validate_alias_valid = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_valid = False
End Function

' Verifies that aliases exceeding MAX_ALIAS_LEN (30 chars) are rejected
' and an error message is returned.
Private Function test_validate_alias_too_long() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_too_long = True
    Dim errMsg As String
    ' MAX_ALIAS_LEN = 30; create a 31-character string to exceed the limit
    Dim longAlias As String
    longAlias = String$(31, "a")
    ' Should be rejected (returns False)
    If ValidateAlias(longAlias, errMsg) Then test_validate_alias_too_long = False: Exit Function
    ' An error message should be populated explaining why it was rejected
    If LenB(errMsg) = 0 Then test_validate_alias_too_long = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_too_long = False
End Function

' Verifies that aliases containing control characters (Tab, Null, DEL)
' are rejected. These are non-printable and could cause display issues.
Private Function test_validate_alias_control_chars() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_control_chars = True
    Dim errMsg As String
    ' Tab character (ASCII 9) embedded in the alias should be rejected
    If ValidateAlias("hello" & Chr$(9) & "world", errMsg) Then test_validate_alias_control_chars = False: Exit Function
    ' Null character (ASCII 0) at the start should be rejected
    errMsg = ""
    If ValidateAlias(Chr$(0) & "test", errMsg) Then test_validate_alias_control_chars = False: Exit Function
    ' DEL character (ASCII 127) at the end should be rejected
    errMsg = ""
    If ValidateAlias("test" & Chr$(127), errMsg) Then test_validate_alias_control_chars = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_control_chars = False
End Function

' Verifies that aliases containing Chr$(255) (high-byte extended ASCII)
' are rejected and an error message is returned.
Private Function test_validate_alias_chr255() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_chr255 = True
    Dim errMsg As String
    ' Chr$(255) is a high-byte extended ASCII character that's not allowed in aliases
    If ValidateAlias("test" & Chr$(255), errMsg) Then test_validate_alias_chr255 = False: Exit Function
    ' An error message should explain the rejection
    If LenB(errMsg) = 0 Then test_validate_alias_chr255 = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_chr255 = False
End Function

#End If
