Attribute VB_Name = "Unit_Characters"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Characters Test Suite
' Tests character creation and deletion on the game map: verifying that
' MapData and UserList are updated correctly when spawning/erasing chars.
' ==========================================================================
Public Function test_suite_characters() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("test_create_char_map", test_create_char_map())
    Call UnitTesting.RunTest("test_create_char_index", test_create_char_index())
    Call UnitTesting.RunTest("test_erase_char_map", test_erase_char_map())
    Call UnitTesting.RunTest("test_erase_char_index", test_erase_char_index())
    Call UnitTesting.RunTest("test_distinct_charindex", test_distinct_charindex())
    Call UnitTesting.RunTest("inventory reset before capacity initialization", test_inventory_reset_all_slots(0))
    Call UnitTesting.RunTest("inventory reset clears normal locked slots", test_inventory_reset_all_slots(get_num_inv_slots_from_tier(tNormal)))
    Call UnitTesting.RunTest("inventory reset clears adventurer locked slots", test_inventory_reset_all_slots(get_num_inv_slots_from_tier(tAventurero)))
    Call UnitTesting.RunTest("inventory reset clears hero locked slots", test_inventory_reset_all_slots(get_num_inv_slots_from_tier(tHeroe)))
    Call UnitTesting.RunTest("inventory reset clears full capacity and tags", test_inventory_reset_all_slots(MAX_INVENTORY_SLOTS))
    
    ' Clean up all characters after suite
    Call CleanupAllChars
    
    Debug.Print "Characters suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_characters = True
End Function

' Simulate a recycled server user slot, including items outside the current tier.
Private Function test_inventory_reset_all_slots(ByVal unlockedSlots As Byte) As Boolean
    On Error GoTo test_inventory_reset_all_slots_Err
    Dim originalInventory As t_Inventario
    Dim originalSlots As Byte
    Dim slot As Integer
    Dim passed As Boolean

    originalInventory = UserList(1).invent
    originalSlots = UserList(1).CurrentInventorySlots
    With UserList(1)
        .CurrentInventorySlots = unlockedSlots
        .invent.NroItems = MAX_INVENTORY_SLOTS
        .invent.EquippedWeaponObjIndex = 1
        .invent.EquippedWeaponSlot = 1
        For slot = 1 To MAX_INVENTORY_SLOTS
            .invent.Object(slot).ObjIndex = slot
            .invent.Object(slot).amount = slot + 1
            .invent.Object(slot).Equipped = 1
            .invent.Object(slot).ElementalTags = slot
        Next slot
    End With

    Call LimpiarInventario(1)

    passed = True
    With UserList(1)
        If .CurrentInventorySlots <> unlockedSlots Then passed = False
        If .invent.NroItems <> 0 Then passed = False
        If .invent.EquippedWeaponObjIndex <> 0 Then passed = False
        If .invent.EquippedWeaponSlot <> 0 Then passed = False
        For slot = 1 To MAX_INVENTORY_SLOTS
            With .invent.Object(slot)
                If .ObjIndex <> 0 Or .amount <> 0 Or .Equipped <> 0 Or .ElementalTags <> 0 Then passed = False
            End With
        Next slot
    End With

    UserList(1).invent = originalInventory
    UserList(1).CurrentInventorySlots = originalSlots
    test_inventory_reset_all_slots = passed
    Exit Function

test_inventory_reset_all_slots_Err:
    UserList(1).invent = originalInventory
    UserList(1).CurrentInventorySlots = originalSlots
    test_inventory_reset_all_slots = False
End Function

' Helper: places a user at the given map position and creates their character.
Private Sub SetupChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    UserList(UserIndex).pos.Map = Map
    UserList(UserIndex).pos.x = x
    UserList(UserIndex).pos.y = y
    Call MakeUserChar(True, 17, UserIndex, Map, x, y, 1)
End Sub

' Helper: removes all active characters from the map to ensure a clean state.
Private Sub CleanupAllChars()
    Dim i As Integer
    For i = 1 To UBound(UserList)
        If UserList(i).Char.charindex <> 0 Then
            Call EraseUserChar(i, False, True)
        End If
    Next i
End Sub

' Verifies that creating a character correctly registers the UserIndex
' in the MapData tile at the character's position.
Private Function test_create_char_map() As Boolean
    On Error GoTo test_create_char_map_Err
    ' Start clean so no leftover chars interfere
    Call CleanupAllChars
    ' Place user 1 at map 1, position (54, 51)
    Call SetupChar(1, 1, 54, 51)
    ' The map tile at (54, 51) should now record UserIndex = 1
    test_create_char_map = (MapData(54, 51, 1).UserIndex = 1)
    Call CleanupAllChars
    Exit Function
test_create_char_map_Err:
    Call CleanupAllChars
    test_create_char_map = False
End Function

' Verifies that creating a character assigns a non-zero charindex,
' which is the unique visual identifier for the character on the map.
Private Function test_create_char_index() As Boolean
    On Error GoTo test_create_char_index_Err
    Call CleanupAllChars
    ' Place user 1 on the map
    Call SetupChar(1, 1, 54, 51)
    ' charindex is the visual ID used by the client to render the character;
    ' it must be non-zero after creation
    test_create_char_index = (UserList(1).Char.charindex <> 0)
    Call CleanupAllChars
    Exit Function
test_create_char_index_Err:
    Call CleanupAllChars
    test_create_char_index = False
End Function

' Verifies that erasing a character clears the UserIndex from the MapData tile,
' so the tile is no longer occupied.
Private Function test_erase_char_map() As Boolean
    On Error GoTo test_erase_char_map_Err
    Call CleanupAllChars
    ' Create then immediately erase user 1
    Call SetupChar(1, 1, 54, 51)
    Call EraseUserChar(1, False, False)
    ' After erasing, the map tile should have UserIndex = 0 (unoccupied)
    test_erase_char_map = (MapData(54, 51, 1).UserIndex = 0)
    Call CleanupAllChars
    Exit Function
test_erase_char_map_Err:
    Call CleanupAllChars
    test_erase_char_map = False
End Function

' Verifies that erasing a character resets its charindex to 0,
' freeing the visual slot for reuse.
Private Function test_erase_char_index() As Boolean
    On Error GoTo test_erase_char_index_Err
    Call CleanupAllChars
    ' Create then erase user 1
    Call SetupChar(1, 1, 54, 51)
    Call EraseUserChar(1, False, False)
    ' After erasing, charindex should be reset to 0 (visual slot freed)
    test_erase_char_index = (UserList(1).Char.charindex = 0)
    Call CleanupAllChars
    Exit Function
test_erase_char_index_Err:
    Call CleanupAllChars
    test_erase_char_index = False
End Function

' Verifies that two characters created simultaneously receive different
' charindex values, ensuring no visual ID collisions on the map.
Private Function test_distinct_charindex() As Boolean
    On Error GoTo test_distinct_charindex_Err
    Call CleanupAllChars
    ' Create two users at different positions on the same map
    Call SetupChar(1, 1, 50, 46)
    Call SetupChar(2, 1, 54, 56)
    ' Each user must get a unique charindex so the client can tell them apart
    test_distinct_charindex = (UserList(1).Char.charindex <> UserList(2).Char.charindex)
    Call CleanupAllChars
    Exit Function
test_distinct_charindex_Err:
    Call CleanupAllChars
    test_distinct_charindex = False
End Function

#End If
