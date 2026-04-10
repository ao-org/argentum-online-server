Attribute VB_Name = "Unit_Characters"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_characters() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("test_create_char_map", test_create_char_map())
    Call UnitTesting.RunTest("test_create_char_index", test_create_char_index())
    Call UnitTesting.RunTest("test_erase_char_map", test_erase_char_map())
    Call UnitTesting.RunTest("test_erase_char_index", test_erase_char_index())
    Call UnitTesting.RunTest("test_distinct_charindex", test_distinct_charindex())
    
    ' Clean up all characters after suite
    Call CleanupAllChars
    
    Debug.Print "Characters suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_characters = True
End Function

Private Sub SetupChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    UserList(UserIndex).pos.Map = Map
    UserList(UserIndex).pos.x = x
    UserList(UserIndex).pos.y = y
    Call MakeUserChar(True, 17, UserIndex, Map, x, y, 1)
End Sub

Private Sub CleanupAllChars()
    Dim i As Integer
    For i = 1 To UBound(UserList)
        If UserList(i).Char.charindex <> 0 Then
            Call EraseUserChar(i, False, True)
        End If
    Next i
End Sub

Private Function test_create_char_map() As Boolean
    On Error GoTo test_create_char_map_Err
    Call CleanupAllChars
    Call SetupChar(1, 1, 54, 51)
    test_create_char_map = (MapData(1, 54, 51).UserIndex = 1)
    Call CleanupAllChars
    Exit Function
test_create_char_map_Err:
    Call CleanupAllChars
    test_create_char_map = False
End Function

Private Function test_create_char_index() As Boolean
    On Error GoTo test_create_char_index_Err
    Call CleanupAllChars
    Call SetupChar(1, 1, 54, 51)
    test_create_char_index = (UserList(1).Char.charindex <> 0)
    Call CleanupAllChars
    Exit Function
test_create_char_index_Err:
    Call CleanupAllChars
    test_create_char_index = False
End Function

Private Function test_erase_char_map() As Boolean
    On Error GoTo test_erase_char_map_Err
    Call CleanupAllChars
    Call SetupChar(1, 1, 54, 51)
    Call EraseUserChar(1, False, False)
    test_erase_char_map = (MapData(1, 54, 51).UserIndex = 0)
    Call CleanupAllChars
    Exit Function
test_erase_char_map_Err:
    Call CleanupAllChars
    test_erase_char_map = False
End Function

Private Function test_erase_char_index() As Boolean
    On Error GoTo test_erase_char_index_Err
    Call CleanupAllChars
    Call SetupChar(1, 1, 54, 51)
    Call EraseUserChar(1, False, False)
    test_erase_char_index = (UserList(1).Char.charindex = 0)
    Call CleanupAllChars
    Exit Function
test_erase_char_index_Err:
    Call CleanupAllChars
    test_erase_char_index = False
End Function

Private Function test_distinct_charindex() As Boolean
    On Error GoTo test_distinct_charindex_Err
    Call CleanupAllChars
    Call SetupChar(1, 1, 50, 46)
    Call SetupChar(2, 1, 54, 56)
    test_distinct_charindex = (UserList(1).Char.charindex <> UserList(2).Char.charindex)
    Call CleanupAllChars
    Exit Function
test_distinct_charindex_Err:
    Call CleanupAllChars
    test_distinct_charindex = False
End Function

#End If
