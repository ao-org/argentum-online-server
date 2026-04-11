Attribute VB_Name = "Unit_UserNames"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_usernames() As Boolean
    Call UnitTesting.RunTest("test_validate_alias_empty", test_validate_alias_empty())
    Call UnitTesting.RunTest("test_validate_alias_valid", test_validate_alias_valid())
    Call UnitTesting.RunTest("test_validate_alias_too_long", test_validate_alias_too_long())
    Call UnitTesting.RunTest("test_validate_alias_control_chars", test_validate_alias_control_chars())
    Call UnitTesting.RunTest("test_validate_alias_chr255", test_validate_alias_chr255())
    test_suite_usernames = True
End Function

Private Function test_validate_alias_empty() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_empty = True
    Dim errMsg As String
    ' Empty string is valid (means "no alias")
    If Not ValidateAlias("", errMsg) Then test_validate_alias_empty = False: Exit Function
    If Not ValidateAlias("   ", errMsg) Then test_validate_alias_empty = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_empty = False
End Function

Private Function test_validate_alias_valid() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_valid = True
    Dim errMsg As String
    If Not ValidateAlias("Warrior", errMsg) Then test_validate_alias_valid = False: Exit Function
    If Not ValidateAlias("Dark Knight 99", errMsg) Then test_validate_alias_valid = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_valid = False
End Function

Private Function test_validate_alias_too_long() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_too_long = True
    Dim errMsg As String
    ' MAX_ALIAS_LEN = 30, create a 31 char string
    Dim longAlias As String
    longAlias = String$(31, "a")
    If ValidateAlias(longAlias, errMsg) Then test_validate_alias_too_long = False: Exit Function
    If LenB(errMsg) = 0 Then test_validate_alias_too_long = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_too_long = False
End Function

Private Function test_validate_alias_control_chars() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_control_chars = True
    Dim errMsg As String
    ' Tab character (code 9)
    If ValidateAlias("hello" & Chr$(9) & "world", errMsg) Then test_validate_alias_control_chars = False: Exit Function
    ' Null character
    errMsg = ""
    If ValidateAlias(Chr$(0) & "test", errMsg) Then test_validate_alias_control_chars = False: Exit Function
    ' DEL character (code 127)
    errMsg = ""
    If ValidateAlias("test" & Chr$(127), errMsg) Then test_validate_alias_control_chars = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_control_chars = False
End Function

Private Function test_validate_alias_chr255() As Boolean
    On Error GoTo Err_Handler
    test_validate_alias_chr255 = True
    Dim errMsg As String
    If ValidateAlias("test" & Chr$(255), errMsg) Then test_validate_alias_chr255 = False: Exit Function
    If LenB(errMsg) = 0 Then test_validate_alias_chr255 = False: Exit Function
    Exit Function
Err_Handler:
    test_validate_alias_chr255 = False
End Function

#End If
