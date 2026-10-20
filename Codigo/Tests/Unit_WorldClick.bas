Attribute VB_Name = "Unit_WorldClick"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_world_click() As Boolean
    Call UnitTesting.RunTest("NPC before player below", CheckWorldClick(0, 10, 20, 0, WorldClickNpc, 10, False))
    Call UnitTesting.RunTest("NPC before pet below", CheckWorldClick(0, 10, 0, 30, WorldClickNpc, 10, False))
    Call UnitTesting.RunTest("Player before NPC below", CheckWorldClick(20, 0, 0, 30, WorldClickUser, 20, False))
    Call UnitTesting.RunTest("Player before player below", CheckWorldClick(20, 0, 21, 0, WorldClickUser, 20, False))
    Call UnitTesting.RunTest("NPC and player share direct tile", CheckWorldClick(20, 10, 21, 30, WorldClickNpc, 10, False))
    Call UnitTesting.RunTest("NPC head click fallback", CheckWorldClick(0, 0, 0, 30, WorldClickNpc, 30, True))
    Call UnitTesting.RunTest("Player head click fallback", CheckWorldClick(0, 0, 20, 0, WorldClickUser, 20, True))
    Call UnitTesting.RunTest("NPC precedence on fallback tile", CheckWorldClick(0, 0, 20, 30, WorldClickNpc, 30, True))
    Call UnitTesting.RunTest("Empty target clears old index", CheckWorldClick(0, 0, 0, 0, WorldClickNone, 0, True))
    Call UnitTesting.RunTest("Direct NPC without fallback", CheckWorldClick(0, 10, 0, 0, WorldClickNpc, 10, False))
    test_suite_world_click = True
End Function

Private Function CheckWorldClick(ByVal DirectUser As Integer, ByVal DirectNpc As Integer, ByVal BelowUser As Integer, ByVal BelowNpc As Integer, ByVal ExpectedKind As e_WorldClickCharacter, ByVal ExpectedIndex As Integer, ByVal ExpectedBelow As Boolean) As Boolean
    On Error GoTo CheckWorldClick_Err
    Dim targetIndex As Integer
    Dim below As Boolean
    Dim kind As e_WorldClickCharacter
    targetIndex = 999
    below = Not ExpectedBelow
    kind = ResolveWorldClickCharacter(DirectUser, DirectNpc, BelowUser, BelowNpc, targetIndex, below)
    CheckWorldClick = (kind = ExpectedKind And targetIndex = ExpectedIndex And below = ExpectedBelow)
    Exit Function
CheckWorldClick_Err:
    CheckWorldClick = False
End Function

#End If
