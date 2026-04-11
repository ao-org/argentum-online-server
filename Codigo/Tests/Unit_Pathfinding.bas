Attribute VB_Name = "Unit_Pathfinding"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_pathfinding() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("test_heading_north", test_heading_north())
    Call UnitTesting.RunTest("test_heading_south", test_heading_south())
    Call UnitTesting.RunTest("test_heading_east", test_heading_east())
    Call UnitTesting.RunTest("test_heading_west", test_heading_west())
    Call UnitTesting.RunTest("test_heading_priority", test_heading_priority())
    
    Debug.Print "Pathfinding suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_pathfinding = True
End Function

Private Function test_heading_north() As Boolean
    On Error GoTo test_heading_north_Err
    Dim current As t_WorldPos
    Dim nextP As t_WorldPos
    current.Map = 1: current.x = 50: current.y = 50
    nextP.Map = 1: nextP.x = 50: nextP.y = 49
    test_heading_north = (GetHeadingFromWorldPos(current, nextP) = e_Heading.NORTH)
    Exit Function
test_heading_north_Err:
    test_heading_north = False
End Function

Private Function test_heading_south() As Boolean
    On Error GoTo test_heading_south_Err
    Dim current As t_WorldPos
    Dim nextP As t_WorldPos
    current.Map = 1: current.x = 50: current.y = 50
    nextP.Map = 1: nextP.x = 50: nextP.y = 51
    test_heading_south = (GetHeadingFromWorldPos(current, nextP) = e_Heading.SOUTH)
    Exit Function
test_heading_south_Err:
    test_heading_south = False
End Function

Private Function test_heading_east() As Boolean
    On Error GoTo test_heading_east_Err
    Dim current As t_WorldPos
    Dim nextP As t_WorldPos
    current.Map = 1: current.x = 50: current.y = 50
    nextP.Map = 1: nextP.x = 51: nextP.y = 50
    test_heading_east = (GetHeadingFromWorldPos(current, nextP) = e_Heading.EAST)
    Exit Function
test_heading_east_Err:
    test_heading_east = False
End Function

Private Function test_heading_west() As Boolean
    On Error GoTo test_heading_west_Err
    Dim current As t_WorldPos
    Dim nextP As t_WorldPos
    current.Map = 1: current.x = 50: current.y = 50
    nextP.Map = 1: nextP.x = 49: nextP.y = 50
    test_heading_west = (GetHeadingFromWorldPos(current, nextP) = e_Heading.WEST)
    Exit Function
test_heading_west_Err:
    test_heading_west = False
End Function

Private Function test_heading_priority() As Boolean
    ' When both dx and dy are non-zero, x takes priority (WEST/EAST before NORTH/SOUTH)
    On Error GoTo test_heading_priority_Err
    test_heading_priority = True
    
    Dim current As t_WorldPos
    Dim nextP As t_WorldPos
    current.Map = 1: current.x = 50: current.y = 50
    
    ' dx < 0, dy < 0 -> WEST (x priority)
    nextP.Map = 1: nextP.x = 49: nextP.y = 49
    If GetHeadingFromWorldPos(current, nextP) <> e_Heading.WEST Then
        test_heading_priority = False: Exit Function
    End If
    
    ' dx > 0, dy > 0 -> EAST (x priority)
    nextP.x = 51: nextP.y = 51
    If GetHeadingFromWorldPos(current, nextP) <> e_Heading.EAST Then
        test_heading_priority = False: Exit Function
    End If
    Exit Function
test_heading_priority_Err:
    test_heading_priority = False
End Function

#End If
