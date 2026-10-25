Attribute VB_Name = "Unit_Crafting"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_crafting() As Boolean
    Call UnitTesting.RunTest("test_sort_unsorted", test_sort_unsorted())
    Call UnitTesting.RunTest("test_sort_already_sorted", test_sort_already_sorted())
    Call UnitTesting.RunTest("test_sort_single_element", test_sort_single_element())
    Call UnitTesting.RunTest("test_get_recipe_key", test_get_recipe_key())
    Call UnitTesting.RunTest("test_sort_then_key_confluence", test_sort_then_key_confluence())
    Call UnitTesting.RunTest("test_sort_ascending_property", test_sort_ascending_property())
    Call UnitTesting.RunTest("craft missing materials leaves inventory unchanged", test_craft_missing_materials())
    Call UnitTesting.RunTest("craft insufficient materials leaves inventory unchanged", test_craft_insufficient_materials())
    Call UnitTesting.RunTest("craft later ingredient failure rolls back all deductions", test_craft_later_material_missing())
    Call UnitTesting.RunTest("craft exact and split stacks consume only ingredients", test_craft_split_stacks())
    Call UnitTesting.RunTest("craft cannot reuse consumed materials", test_craft_repeated_request())
    Call UnitTesting.RunTest("craft tagged and locked stacks do not pay normal costs", test_craft_ineligible_stacks())
    Call UnitTesting.RunTest("craft carpentry multiplies the full batch cost", test_craft_batch_quantity())
    Call UnitTesting.RunTest("craft empty negative and invalid recipes fail closed", test_craft_invalid_recipe())
    Call UnitTesting.RunTest("craft excessive quantity cannot overflow costs", test_craft_quantity_overflow())
    Call UnitTesting.RunTest("craft corrupt material stacks fail without mutation", test_craft_corrupt_stack())
    Call UnitTesting.RunTest("craft all profession consumables are deducted", test_craft_profession_materials())
    Call UnitTesting.RunTest("alchemy requires and preserves reusable mortar", test_craft_reusable_mortar())
    test_suite_crafting = True
End Function

' These tests use local inventory/recipe values, never live users or networking.
Private Sub CraftStack(ByRef inventory As t_Inventario, ByVal slot As Integer, ByVal objectId As Integer, ByVal amount As Integer)
    On Error GoTo CraftStack_Err
    inventory.Object(slot).ObjIndex = objectId
    inventory.Object(slot).amount = amount
    inventory.NroItems = inventory.NroItems + 1
    Exit Sub
CraftStack_Err:
    Dim errorNumber As Long, errorDescription As String
    errorNumber = Err.Number
    errorDescription = Err.Description
    Call TraceError(errorNumber, errorDescription, "Unit_Crafting.CraftStack", Erl)
    Call Err.Raise(errorNumber, "Unit_Crafting.CraftStack", errorDescription)
End Sub

Private Function SameCraftInventory(ByRef first As t_Inventario, ByRef second As t_Inventario) As Boolean
    On Error GoTo SameCraftInventory_Err
    Dim slot As Integer
    If first.NroItems <> second.NroItems Then Exit Function
    If first.EquippedWorkingToolSlot <> second.EquippedWorkingToolSlot Then Exit Function
    If first.EquippedWorkingToolObjIndex <> second.EquippedWorkingToolObjIndex Then Exit Function
    For slot = 1 To MAX_INVENTORY_SLOTS
        If first.Object(slot).ObjIndex <> second.Object(slot).ObjIndex Then Exit Function
        If first.Object(slot).amount <> second.Object(slot).amount Then Exit Function
        If first.Object(slot).Equipped <> second.Object(slot).Equipped Then Exit Function
        If first.Object(slot).ElementalTags <> second.Object(slot).ElementalTags Then Exit Function
    Next slot
    SameCraftInventory = True
    Exit Function
SameCraftInventory_Err:
    SameCraftInventory = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.SameCraftInventory", Erl)
End Function

Private Function test_craft_missing_materials() As Boolean
    On Error GoTo test_craft_missing_materials_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.LingH = 85
    original = inventory
    If TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    test_craft_missing_materials = SameCraftInventory(inventory, original)
    Exit Function
test_craft_missing_materials_Err:
    test_craft_missing_materials = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_missing_materials", Erl)
End Function

Private Function test_craft_insufficient_materials() As Boolean
    On Error GoTo test_craft_insufficient_materials_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.LingH = 85
    Call CraftStack(inventory, 1, e_Minerales.LingoteDeHierro, 84)
    original = inventory
    If TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    test_craft_insufficient_materials = SameCraftInventory(inventory, original)
    Exit Function
test_craft_insufficient_materials_Err:
    test_craft_insufficient_materials = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_insufficient_materials", Erl)
End Function

Private Function test_craft_later_material_missing() As Boolean
    On Error GoTo test_craft_later_material_missing_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.LingH = 85
    recipe.LingP = 110
    Call CraftStack(inventory, 1, e_Minerales.LingoteDeHierro, 85)
    Call CraftStack(inventory, 2, e_Minerales.LingoteDePlata, 109)
    original = inventory
    If TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    test_craft_later_material_missing = SameCraftInventory(inventory, original)
    Exit Function
test_craft_later_material_missing_Err:
    test_craft_later_material_missing = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_later_material_missing", Erl)
End Function

Private Function test_craft_split_stacks() As Boolean
    On Error GoTo test_craft_split_stacks_Err
    Dim inventory As t_Inventario, recipe As t_ObjData
    recipe.LingH = 85
    recipe.LingP = 110
    Call CraftStack(inventory, 1, e_Minerales.LingoteDeHierro, 40)
    Call CraftStack(inventory, 2, e_Minerales.LingoteDeHierro, 45)
    Call CraftStack(inventory, 3, e_Minerales.LingoteDePlata, 120)
    Call CraftStack(inventory, 4, Wood, 10)
    inventory.EquippedWorkingToolSlot = 5
    inventory.EquippedWorkingToolObjIndex = 389
    Call CraftStack(inventory, 5, 389, 1)
    inventory.Object(5).Equipped = 1
    If Not TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    If inventory.Object(1).ObjIndex <> 0 Or inventory.Object(1).amount <> 0 Then Exit Function
    If inventory.Object(2).ObjIndex <> 0 Or inventory.Object(2).amount <> 0 Then Exit Function
    If inventory.Object(3).amount <> 10 Or inventory.Object(4).amount <> 10 Then Exit Function
    If inventory.NroItems <> 3 Then Exit Function
    If inventory.EquippedWorkingToolSlot <> 5 Or inventory.EquippedWorkingToolObjIndex <> 389 Then Exit Function
    test_craft_split_stacks = (inventory.Object(5).Equipped = 1 And inventory.Object(5).amount = 1)
    Exit Function
test_craft_split_stacks_Err:
    test_craft_split_stacks = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_split_stacks", Erl)
End Function

Private Function test_craft_repeated_request() As Boolean
    On Error GoTo test_craft_repeated_request_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.LingH = 85
    Call CraftStack(inventory, 1, e_Minerales.LingoteDeHierro, 85)
    If Not TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    original = inventory
    If TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    test_craft_repeated_request = SameCraftInventory(inventory, original)
    Exit Function
test_craft_repeated_request_Err:
    test_craft_repeated_request = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_repeated_request", Erl)
End Function

Private Function test_craft_ineligible_stacks() As Boolean
    On Error GoTo test_craft_ineligible_stacks_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.LingH = 2
    Call CraftStack(inventory, 1, e_Minerales.LingoteDeHierro, 1)
    Call CraftStack(inventory, 2, e_Minerales.LingoteDeHierro, 100)
    inventory.Object(2).ElementalTags = 1
    Call CraftStack(inventory, MAX_INVENTORY_SLOTS, e_Minerales.LingoteDeHierro, 100)
    original = inventory
    If TryConsumeCraftingMaterials(inventory, 2, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    test_craft_ineligible_stacks = SameCraftInventory(inventory, original)
    Exit Function
test_craft_ineligible_stacks_Err:
    test_craft_ineligible_stacks = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_ineligible_stacks", Erl)
End Function

Private Function test_craft_batch_quantity() As Boolean
    On Error GoTo test_craft_batch_quantity_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.Madera = 2
    recipe.MaderaElfica = 1
    recipe.MaderaPino = 3
    Call CraftStack(inventory, 1, Wood, 6)
    Call CraftStack(inventory, 2, ElvenWood, 3)
    Call CraftStack(inventory, 3, PinoWood, 8)
    original = inventory
    If TryConsumeCraftingMaterials(inventory, 3, recipe, e_JobsTypes.Carpenter, 3) Then Exit Function
    If Not SameCraftInventory(inventory, original) Then Exit Function
    inventory.Object(3).amount = 9
    If Not TryConsumeCraftingMaterials(inventory, 3, recipe, e_JobsTypes.Carpenter, 3) Then Exit Function
    test_craft_batch_quantity = (inventory.NroItems = 0)
    Exit Function
test_craft_batch_quantity_Err:
    test_craft_batch_quantity = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_batch_quantity", Erl)
End Function

Private Function test_craft_invalid_recipe() As Boolean
    On Error GoTo test_craft_invalid_recipe_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    Call CraftStack(inventory, 1, e_Minerales.LingoteDeHierro, 10)
    original = inventory
    If TryConsumeCraftingMaterials(inventory, 1, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    recipe.LingH = 1
    recipe.LingP = -1 ' Invalid later field must not consume the first ingredient.
    If TryConsumeCraftingMaterials(inventory, 1, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    recipe.LingP = 0
    If TryConsumeCraftingMaterials(inventory, 0, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    If TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS + 1, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    If TryConsumeCraftingMaterials(inventory, 1, recipe, e_JobsTypes.Blacksmith, 0) Then Exit Function
    If TryConsumeCraftingMaterials(inventory, 1, recipe, e_JobsTypes.Blacksmith, -1) Then Exit Function
    If TryConsumeCraftingMaterials(inventory, 1, recipe, e_JobsTypes.Blacksmith, 2) Then Exit Function
    If TryConsumeCraftingMaterials(inventory, 1, recipe, 0, 1) Then Exit Function
    test_craft_invalid_recipe = SameCraftInventory(inventory, original)
    Exit Function
test_craft_invalid_recipe_Err:
    test_craft_invalid_recipe = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_invalid_recipe", Erl)
End Function

Private Function test_craft_quantity_overflow() As Boolean
    On Error GoTo test_craft_quantity_overflow_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.Madera = 32767
    Call CraftStack(inventory, 1, Wood, 32767)
    original = inventory
    If TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, e_JobsTypes.Carpenter, 2147483647) Then Exit Function
    test_craft_quantity_overflow = SameCraftInventory(inventory, original)
    Exit Function
test_craft_quantity_overflow_Err:
    test_craft_quantity_overflow = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_quantity_overflow", Erl)
End Function

Private Function test_craft_corrupt_stack() As Boolean
    On Error GoTo test_craft_corrupt_stack_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.LingH = 1
    recipe.LingP = 1
    Call CraftStack(inventory, 1, e_Minerales.LingoteDeHierro, 1)
    Call CraftStack(inventory, 2, e_Minerales.LingoteDePlata, 1)
    inventory.Object(2).Equipped = 1
    original = inventory
    If TryConsumeCraftingMaterials(inventory, 2, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    If Not SameCraftInventory(inventory, original) Then Exit Function
    inventory.Object(2).Equipped = 0
    inventory.Object(2).amount = -1
    original = inventory
    If TryConsumeCraftingMaterials(inventory, 2, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    test_craft_corrupt_stack = SameCraftInventory(inventory, original)
    Exit Function
test_craft_corrupt_stack_Err:
    test_craft_corrupt_stack = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_corrupt_stack", Erl)
End Function

Private Function test_craft_reusable_mortar() As Boolean
    On Error GoTo test_craft_reusable_mortar_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.Mortero = 1
    recipe.FrascoAlq = 1
    recipe.FlorRoja = 2
    Call CraftStack(inventory, 1, FrascoAlq, 1)
    Call CraftStack(inventory, 2, FlorRoja, 2)
    original = inventory
    If TryConsumeCraftingMaterials(inventory, 2, recipe, e_JobsTypes.Alchemist, 1) Then Exit Function
    If Not SameCraftInventory(inventory, original) Then Exit Function
    Call CraftStack(inventory, 3, Mortero, 1)
    If Not TryConsumeCraftingMaterials(inventory, 3, recipe, e_JobsTypes.Alchemist, 1) Then Exit Function
    If inventory.NroItems <> 1 Then Exit Function
    test_craft_reusable_mortar = (inventory.Object(3).ObjIndex = Mortero And inventory.Object(3).amount = 1)
    Exit Function
test_craft_reusable_mortar_Err:
    test_craft_reusable_mortar = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_reusable_mortar", Erl)
End Function

Private Function test_craft_profession_materials() As Boolean
    On Error GoTo test_craft_profession_materials_Err
    Dim inventory As t_Inventario, emptyInventory As t_Inventario, recipe As t_ObjData
    recipe.LingH = 1: recipe.LingP = 1: recipe.LingO = 1: recipe.Coal = 1: recipe.Blodium = 1
    recipe.FireEssence = 1: recipe.WaterEssence = 1: recipe.EarthEssence = 1: recipe.WindEssence = 1
    Call CraftStack(inventory, 1, e_Minerales.LingoteDeHierro, 1)
    Call CraftStack(inventory, 2, e_Minerales.LingoteDePlata, 1)
    Call CraftStack(inventory, 3, e_Minerales.LingoteDeOro, 1)
    Call CraftStack(inventory, 4, e_Minerales.Coal, 1)
    Call CraftStack(inventory, 5, e_Minerales.Blodium, 1)
    Call CraftStack(inventory, 6, e_Minerales.FireEssence, 1)
    Call CraftStack(inventory, 7, e_Minerales.WaterEssence, 1)
    Call CraftStack(inventory, 8, e_Minerales.EarthEssence, 1)
    Call CraftStack(inventory, 9, e_Minerales.WindEssence, 1)
    If Not TryConsumeCraftingMaterials(inventory, 9, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    If Not SameCraftInventory(inventory, emptyInventory) Then Exit Function
    recipe.Mortero = 1: recipe.FrascoAlq = 1: recipe.HongoDeLuz = 1: recipe.Esporas = 1: recipe.Tuna = 1
    recipe.Cala = 1: recipe.ColaDeZorro = 1: recipe.FlorOceano = 1: recipe.FlorRoja = 1: recipe.SemillasProsperas = 1
    Call CraftStack(inventory, 1, Mortero, 1)
    Call CraftStack(inventory, 2, FrascoAlq, 1)
    Call CraftStack(inventory, 3, HongoDeLuz, 1)
    Call CraftStack(inventory, 4, Esporas, 1)
    Call CraftStack(inventory, 5, Tuna, 1)
    Call CraftStack(inventory, 6, Cala, 1)
    Call CraftStack(inventory, 7, ColaDeZorro, 1)
    Call CraftStack(inventory, 8, FlorOceano, 1)
    Call CraftStack(inventory, 9, FlorRoja, 1)
    Call CraftStack(inventory, 10, SemillasProsperas, 1)
    If Not TryConsumeCraftingMaterials(inventory, 10, recipe, e_JobsTypes.Alchemist, 1) Then Exit Function
    If inventory.NroItems <> 1 Then Exit Function
    If inventory.Object(1).ObjIndex <> Mortero Or inventory.Object(1).amount <> 1 Then Exit Function
    inventory = emptyInventory
    recipe.PielLobo = 1: recipe.PielOsoPardo = 1: recipe.PielOsoPolaR = 1
    recipe.PielLoboNegro = 1: recipe.PielTigre = 1: recipe.PielTigreBengala = 1
    Call CraftStack(inventory, 1, PieldeLobo, 1)
    Call CraftStack(inventory, 2, PieldeOsoPardo, 1)
    Call CraftStack(inventory, 3, PieldeOsoPolar, 1)
    Call CraftStack(inventory, 4, PielLoboNegro, 1)
    Call CraftStack(inventory, 5, PielTigre, 1)
    Call CraftStack(inventory, 6, PielTigreBengala, 1)
    If Not TryConsumeCraftingMaterials(inventory, 6, recipe, e_JobsTypes.Tailor, 1) Then Exit Function
    test_craft_profession_materials = SameCraftInventory(inventory, emptyInventory)
    Exit Function
test_craft_profession_materials_Err:
    test_craft_profession_materials = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_profession_materials", Erl)
End Function

' Verify SortIntegerArray sorts an unsorted array {5, 3, 1, 4, 2} into ascending order.
Private Function test_sort_unsorted() As Boolean
    On Error GoTo Err_Handler
    test_sort_unsorted = True

    Dim data(1 To 5) As Integer
    data(1) = 5
    data(2) = 3
    data(3) = 1
    data(4) = 4
    data(5) = 2

    Call SortIntegerArray(data, LBound(data), UBound(data))

    If data(1) <> 1 Then test_sort_unsorted = False: Exit Function
    If data(2) <> 2 Then test_sort_unsorted = False: Exit Function
    If data(3) <> 3 Then test_sort_unsorted = False: Exit Function
    If data(4) <> 4 Then test_sort_unsorted = False: Exit Function
    If data(5) <> 5 Then test_sort_unsorted = False: Exit Function

    Exit Function
Err_Handler:
    test_sort_unsorted = False
End Function

' Verify SortIntegerArray is idempotent — an already-sorted array {1, 2, 3} remains unchanged.
Private Function test_sort_already_sorted() As Boolean
    On Error GoTo Err_Handler
    test_sort_already_sorted = True

    Dim data(1 To 3) As Integer
    data(1) = 1
    data(2) = 2
    data(3) = 3

    Call SortIntegerArray(data, LBound(data), UBound(data))

    If data(1) <> 1 Then test_sort_already_sorted = False: Exit Function
    If data(2) <> 2 Then test_sort_already_sorted = False: Exit Function
    If data(3) <> 3 Then test_sort_already_sorted = False: Exit Function

    Exit Function
Err_Handler:
    test_sort_already_sorted = False
End Function

' Verify SortIntegerArray handles a single-element array {42} without error.
Private Function test_sort_single_element() As Boolean
    On Error GoTo Err_Handler
    test_sort_single_element = True

    Dim data(1 To 1) As Integer
    data(1) = 42

    Call SortIntegerArray(data, LBound(data), UBound(data))

    If data(1) <> 42 Then test_sort_single_element = False: Exit Function

    Exit Function
Err_Handler:
    test_sort_single_element = False
End Function

' Verify GetRecipeKey produces a colon-delimited string "1:2:3:" for sorted array {1, 2, 3}.
Private Function test_get_recipe_key() As Boolean
    On Error GoTo Err_Handler
    test_get_recipe_key = True

    Dim data(1 To 3) As Integer
    data(1) = 1
    data(2) = 2
    data(3) = 3

    Dim result As String
    result = GetRecipeKey(data)

    If result <> "1:2:3:" Then test_get_recipe_key = False: Exit Function

    Exit Function
Err_Handler:
    test_get_recipe_key = False
End Function

' Verify that two arrays with the same elements in different initial orders
' produce the same recipe key after sorting (confluence property).
Private Function test_sort_then_key_confluence() As Boolean
    On Error GoTo Err_Handler
    test_sort_then_key_confluence = True

    Dim dataA(1 To 3) As Integer
    dataA(1) = 3
    dataA(2) = 1
    dataA(3) = 2

    Dim dataB(1 To 3) As Integer
    dataB(1) = 2
    dataB(2) = 3
    dataB(3) = 1

    Call SortIntegerArray(dataA, LBound(dataA), UBound(dataA))
    Call SortIntegerArray(dataB, LBound(dataB), UBound(dataB))

    Dim keyA As String
    Dim keyB As String
    keyA = GetRecipeKey(dataA)
    keyB = GetRecipeKey(dataB)

    If keyA <> keyB Then test_sort_then_key_confluence = False: Exit Function

    Exit Function
Err_Handler:
    test_sort_then_key_confluence = False
End Function

' Property 2: SortIntegerArray produces ascending order.
' For any random integer array of length 1-5, after sorting,
' data(i) <= data(i+1) for all adjacent pairs.
' Uses 150 randomized trials to approximate universal quantification.
Private Function test_sort_ascending_property() As Boolean
    On Error GoTo Err_Handler
    test_sort_ascending_property = True
    
    Dim i As Long
    Dim j As Long
    Dim arrLen As Integer
    
    Randomize
    
    For i = 1 To 150
        ' Random array length from 1 to 5 (MAX_SLOTS_CRAFTEO)
        arrLen = CInt(Int(Rnd * 5) + 1)
        
        ReDim data(1 To arrLen) As Integer
        
        ' Fill with random Integer values across full range
        For j = 1 To arrLen
            data(j) = CInt(Int(Rnd * 65536) - 32768)
        Next j
        
        ' Sort the array
        Call SortIntegerArray(data, LBound(data), UBound(data))
        
        ' Verify ascending order: data(k) <= data(k+1) for all adjacent pairs
        If arrLen > 1 Then
            Dim k As Long
            For k = LBound(data) To UBound(data) - 1
                If data(k) > data(k + 1) Then
                    test_sort_ascending_property = False: Exit Function
                End If
            Next k
        End If
    Next i
    
    Exit Function
Err_Handler:
    test_sort_ascending_property = False
End Function

#End If
