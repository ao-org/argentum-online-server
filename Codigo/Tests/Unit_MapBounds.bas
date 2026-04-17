Attribute VB_Name = "Unit_MapBounds"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_mapbounds() As Boolean
    Call UnitTesting.RunTest("test_in_map_bounds_valid", test_in_map_bounds_valid())
    Call UnitTesting.RunTest("test_in_map_bounds_map_zero", test_in_map_bounds_map_zero())
    Call UnitTesting.RunTest("test_in_map_bounds_map_negative", test_in_map_bounds_map_negative())
    Call UnitTesting.RunTest("test_in_map_bounds_map_exceeds", test_in_map_bounds_map_exceeds())
    Call UnitTesting.RunTest("test_in_map_bounds_x_below", test_in_map_bounds_x_below())
    Call UnitTesting.RunTest("test_in_map_bounds_y_below", test_in_map_bounds_y_below())
    Call UnitTesting.RunTest("test_in_map_bounds_x_above", test_in_map_bounds_x_above())
    Call UnitTesting.RunTest("test_in_map_bounds_y_above", test_in_map_bounds_y_above())
    Call UnitTesting.RunTest("test_in_map_bounds_exact_edges", test_in_map_bounds_exact_edges())
    Call UnitTesting.RunTest("test_head_to_pos_north", test_head_to_pos_north())
    Call UnitTesting.RunTest("test_head_to_pos_south", test_head_to_pos_south())
    Call UnitTesting.RunTest("test_head_to_pos_east", test_head_to_pos_east())
    Call UnitTesting.RunTest("test_head_to_pos_west", test_head_to_pos_west())
    Call UnitTesting.RunTest("test_prop_in_map_bounds_boundary_completeness", test_prop_in_map_bounds_boundary_completeness())
    Call UnitTesting.RunTest("test_prop_head_to_pos_inverse_symmetry", test_prop_head_to_pos_inverse_symmetry())

    test_suite_mapbounds = True
End Function

' Verifies that a valid map number and coordinates within bounds returns True.
Private Function test_in_map_bounds_valid() As Boolean
    On Error GoTo Err_Handler
    test_in_map_bounds_valid = InMapBounds(1, 50, 50)
    Exit Function
Err_Handler:
    test_in_map_bounds_valid = False
End Function

' Verifies that map=0 returns False.
Private Function test_in_map_bounds_map_zero() As Boolean
    On Error GoTo Err_Handler
    test_in_map_bounds_map_zero = Not InMapBounds(0, 50, 50)
    Exit Function
Err_Handler:
    test_in_map_bounds_map_zero = False
End Function

' Verifies that a negative map number returns False.
Private Function test_in_map_bounds_map_negative() As Boolean
    On Error GoTo Err_Handler
    test_in_map_bounds_map_negative = Not InMapBounds(-1, 50, 50)
    Exit Function
Err_Handler:
    test_in_map_bounds_map_negative = False
End Function

' Verifies that a map number exceeding NumMaps returns False.
Private Function test_in_map_bounds_map_exceeds() As Boolean
    On Error GoTo Err_Handler
    test_in_map_bounds_map_exceeds = Not InMapBounds(CInt(NumMaps + 1), 50, 50)
    Exit Function
Err_Handler:
    test_in_map_bounds_map_exceeds = False
End Function

' Verifies that X below MinXBorder returns False.
Private Function test_in_map_bounds_x_below() As Boolean
    On Error GoTo Err_Handler
    test_in_map_bounds_x_below = Not InMapBounds(1, MinXBorder - 1, 50)
    Exit Function
Err_Handler:
    test_in_map_bounds_x_below = False
End Function

' Verifies that Y below MinYBorder returns False.
Private Function test_in_map_bounds_y_below() As Boolean
    On Error GoTo Err_Handler
    test_in_map_bounds_y_below = Not InMapBounds(1, 50, MinYBorder - 1)
    Exit Function
Err_Handler:
    test_in_map_bounds_y_below = False
End Function

' Verifies that X above MaxXBorder returns False.
Private Function test_in_map_bounds_x_above() As Boolean
    On Error GoTo Err_Handler
    test_in_map_bounds_x_above = Not InMapBounds(1, MaxXBorder + 1, 50)
    Exit Function
Err_Handler:
    test_in_map_bounds_x_above = False
End Function

' Verifies that Y above MaxYBorder returns False.
Private Function test_in_map_bounds_y_above() As Boolean
    On Error GoTo Err_Handler
    test_in_map_bounds_y_above = Not InMapBounds(1, 50, MaxYBorder + 1)
    Exit Function
Err_Handler:
    test_in_map_bounds_y_above = False
End Function

' Verifies that exact boundary values (MinXBorder, MaxXBorder, MinYBorder, MaxYBorder)
' on a valid map all return True.
Private Function test_in_map_bounds_exact_edges() As Boolean
    On Error GoTo Err_Handler
    test_in_map_bounds_exact_edges = True

    ' Test all four corner combinations
    If Not InMapBounds(1, MinXBorder, MinYBorder) Then test_in_map_bounds_exact_edges = False: Exit Function
    If Not InMapBounds(1, MaxXBorder, MinYBorder) Then test_in_map_bounds_exact_edges = False: Exit Function
    If Not InMapBounds(1, MinXBorder, MaxYBorder) Then test_in_map_bounds_exact_edges = False: Exit Function
    If Not InMapBounds(1, MaxXBorder, MaxYBorder) Then test_in_map_bounds_exact_edges = False: Exit Function

    Exit Function
Err_Handler:
    test_in_map_bounds_exact_edges = False
End Function

' Property: boundary values return True, one step outside returns False.
' Iterates over all boundary edges systematically.
Private Function test_prop_in_map_bounds_boundary_completeness() As Boolean
    On Error GoTo Err_Handler
    test_prop_in_map_bounds_boundary_completeness = True

    Dim testX As Integer
    Dim testY As Integer

    ' Test X boundaries: for each valid Y, boundary X returns True, one step outside returns False
    For testY = MinYBorder To MaxYBorder
        ' MinXBorder should be in bounds
        If Not InMapBounds(1, CInt(MinXBorder), testY) Then
            test_prop_in_map_bounds_boundary_completeness = False
            Exit Function
        End If
        ' One step below MinXBorder should be out of bounds
        If InMapBounds(1, CInt(MinXBorder) - 1, testY) Then
            test_prop_in_map_bounds_boundary_completeness = False
            Exit Function
        End If
        ' MaxXBorder should be in bounds
        If Not InMapBounds(1, CInt(MaxXBorder), testY) Then
            test_prop_in_map_bounds_boundary_completeness = False
            Exit Function
        End If
        ' One step above MaxXBorder should be out of bounds
        If InMapBounds(1, CInt(MaxXBorder) + 1, testY) Then
            test_prop_in_map_bounds_boundary_completeness = False
            Exit Function
        End If
    Next testY

    ' Test Y boundaries: for each valid X, boundary Y returns True, one step outside returns False
    For testX = MinXBorder To MaxXBorder
        ' MinYBorder should be in bounds
        If Not InMapBounds(1, testX, CInt(MinYBorder)) Then
            test_prop_in_map_bounds_boundary_completeness = False
            Exit Function
        End If
        ' One step below MinYBorder should be out of bounds
        If InMapBounds(1, testX, CInt(MinYBorder) - 1) Then
            test_prop_in_map_bounds_boundary_completeness = False
            Exit Function
        End If
        ' MaxYBorder should be in bounds
        If Not InMapBounds(1, testX, CInt(MaxYBorder)) Then
            test_prop_in_map_bounds_boundary_completeness = False
            Exit Function
        End If
        ' One step above MaxYBorder should be out of bounds
        If InMapBounds(1, testX, CInt(MaxYBorder) + 1) Then
            test_prop_in_map_bounds_boundary_completeness = False
            Exit Function
        End If
    Next testX

    Exit Function
Err_Handler:
    test_prop_in_map_bounds_boundary_completeness = False
End Function

' Verifies that heading NORTH decrements Y by 1 and leaves X unchanged.
Private Function test_head_to_pos_north() As Boolean
    On Error GoTo Err_Handler
    Dim pos As t_WorldPos
    pos.Map = 1
    pos.x = 50
    pos.y = 50

    Call HeadtoPos(e_Heading.NORTH, pos)

    test_head_to_pos_north = (pos.x = 50 And pos.y = 49)
    Exit Function
Err_Handler:
    test_head_to_pos_north = False
End Function

' Verifies that heading SOUTH increments Y by 1 and leaves X unchanged.
Private Function test_head_to_pos_south() As Boolean
    On Error GoTo Err_Handler
    Dim pos As t_WorldPos
    pos.Map = 1
    pos.x = 50
    pos.y = 50

    Call HeadtoPos(e_Heading.SOUTH, pos)

    test_head_to_pos_south = (pos.x = 50 And pos.y = 51)
    Exit Function
Err_Handler:
    test_head_to_pos_south = False
End Function

' Verifies that heading EAST increments X by 1 and leaves Y unchanged.
Private Function test_head_to_pos_east() As Boolean
    On Error GoTo Err_Handler
    Dim pos As t_WorldPos
    pos.Map = 1
    pos.x = 50
    pos.y = 50

    Call HeadtoPos(e_Heading.EAST, pos)

    test_head_to_pos_east = (pos.x = 51 And pos.y = 50)
    Exit Function
Err_Handler:
    test_head_to_pos_east = False
End Function

' Verifies that heading WEST decrements X by 1 and leaves Y unchanged.
Private Function test_head_to_pos_west() As Boolean
    On Error GoTo Err_Handler
    Dim pos As t_WorldPos
    pos.Map = 1
    pos.x = 50
    pos.y = 50

    Call HeadtoPos(e_Heading.WEST, pos)

    test_head_to_pos_west = (pos.x = 49 And pos.y = 50)
    Exit Function
Err_Handler:
    test_head_to_pos_west = False
End Function

' Property: NORTH+SOUTH and EAST+WEST round-trips must return to the original position.
' Tests across a grid of starting positions.
Private Function test_prop_head_to_pos_inverse_symmetry() As Boolean
    On Error GoTo Err_Handler
    test_prop_head_to_pos_inverse_symmetry = True

    Dim startX As Integer
    Dim startY As Integer
    Dim pos As t_WorldPos

    ' Test NORTH then SOUTH inverse across a range of positions
    For startX = 10 To 90 Step 10
        For startY = 10 To 90 Step 10
            pos.Map = 1
            pos.x = startX
            pos.y = startY

            Call HeadtoPos(e_Heading.NORTH, pos)
            Call HeadtoPos(e_Heading.SOUTH, pos)

            If pos.x <> startX Or pos.y <> startY Then
                test_prop_head_to_pos_inverse_symmetry = False
                Exit Function
            End If
        Next startY
    Next startX

    ' Test EAST then WEST inverse across a range of positions
    For startX = 10 To 90 Step 10
        For startY = 10 To 90 Step 10
            pos.Map = 1
            pos.x = startX
            pos.y = startY

            Call HeadtoPos(e_Heading.EAST, pos)
            Call HeadtoPos(e_Heading.WEST, pos)

            If pos.x <> startX Or pos.y <> startY Then
                test_prop_head_to_pos_inverse_symmetry = False
                Exit Function
            End If
        Next startY
    Next startX

    ' Also test reverse order: SOUTH then NORTH, WEST then EAST
    For startX = 10 To 90 Step 10
        For startY = 10 To 90 Step 10
            pos.Map = 1
            pos.x = startX
            pos.y = startY

            Call HeadtoPos(e_Heading.SOUTH, pos)
            Call HeadtoPos(e_Heading.NORTH, pos)

            If pos.x <> startX Or pos.y <> startY Then
                test_prop_head_to_pos_inverse_symmetry = False
                Exit Function
            End If

            pos.x = startX
            pos.y = startY

            Call HeadtoPos(e_Heading.WEST, pos)
            Call HeadtoPos(e_Heading.EAST, pos)

            If pos.x <> startX Or pos.y <> startY Then
                test_prop_head_to_pos_inverse_symmetry = False
                Exit Function
            End If
        Next startY
    Next startX

    Exit Function
Err_Handler:
    test_prop_head_to_pos_inverse_symmetry = False
End Function

#End If
