Attribute VB_Name = "Unit_AdminVisibility"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_admin_visibility() As Boolean
    On Error GoTo test_suite_admin_visibility_Err
    Call UnitTesting.RunTest("visible targets are visible to ordinary users", test_visible_target())
    Call UnitTesting.RunTest("admin invisible targets are hidden from ordinary users", test_ordinary_user_cannot_see())
    Call UnitTesting.RunTest("admin invisible targets are hidden from lower staff", test_lower_staff_cannot_see())
    Call UnitTesting.RunTest("admin invisible targets are visible to equal staff", test_equal_staff_can_see())
    Call UnitTesting.RunTest("admin invisible targets are visible to higher staff", test_higher_staff_can_see())
    test_suite_admin_visibility = True
    Exit Function
test_suite_admin_visibility_Err:
    Call UnitTesting.RunTestError("admin invisible visibility suite", Err.Description)
End Function

Private Function test_visible_target() As Boolean
    On Error GoTo test_visible_target_Err
    test_visible_target = CanPrivilegesSeeAdminInvisible(e_PlayerType.User, e_PlayerType.Admin, False)
    Exit Function
test_visible_target_Err:
    test_visible_target = False
End Function

Private Function test_ordinary_user_cannot_see() As Boolean
    On Error GoTo test_ordinary_user_cannot_see_Err
    test_ordinary_user_cannot_see = Not CanPrivilegesSeeAdminInvisible(e_PlayerType.User, e_PlayerType.SemiDios, True)
    Exit Function
test_ordinary_user_cannot_see_Err:
    test_ordinary_user_cannot_see = False
End Function

Private Function test_lower_staff_cannot_see() As Boolean
    On Error GoTo test_lower_staff_cannot_see_Err
    test_lower_staff_cannot_see = Not CanPrivilegesSeeAdminInvisible(e_PlayerType.Consejero, e_PlayerType.SemiDios, True)
    Exit Function
test_lower_staff_cannot_see_Err:
    test_lower_staff_cannot_see = False
End Function

Private Function test_equal_staff_can_see() As Boolean
    On Error GoTo test_equal_staff_can_see_Err
    test_equal_staff_can_see = CanPrivilegesSeeAdminInvisible(e_PlayerType.SemiDios, e_PlayerType.SemiDios, True)
    Exit Function
test_equal_staff_can_see_Err:
    test_equal_staff_can_see = False
End Function

Private Function test_higher_staff_can_see() As Boolean
    On Error GoTo test_higher_staff_can_see_Err
    test_higher_staff_can_see = CanPrivilegesSeeAdminInvisible(e_PlayerType.Dios, e_PlayerType.SemiDios, True)
    Exit Function
test_higher_staff_can_see_Err:
    test_higher_staff_can_see = False
End Function

#End If
