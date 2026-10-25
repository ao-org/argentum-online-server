Attribute VB_Name = "Unit_Crafting"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection

Public Function CapturePacket() As Boolean
    If Not capturing Then Exit Function
    Dim bytes() As Byte, value As Variant, writer As Network.writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    value = bytes
    Call packets.Add(value)
    CapturePacket = True
End Function

Private Sub Expect(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Call Err.Raise(vbObjectError + 1885, "Crafting regression", detail)
End Sub

Public Function test_suite_crafting() As Boolean
    Dim failuresBefore As Long
    failuresBefore = UnitTesting.GetFailedTests()
    Call UnitTesting.RunTest("test_sort_unsorted", test_sort_unsorted())
    Call UnitTesting.RunTest("test_sort_already_sorted", test_sort_already_sorted())
    Call UnitTesting.RunTest("test_sort_single_element", test_sort_single_element())
    Call UnitTesting.RunTest("test_get_recipe_key", test_get_recipe_key())
    Call UnitTesting.RunTest("test_sort_then_key_confluence", test_sort_then_key_confluence())
    Call UnitTesting.RunTest("test_sort_ascending_property", test_sort_ascending_property())
    Call UnitTesting.RunTest("craft loaded recipes support every object type", test_loaded_generic_recipes())
    Call UnitTesting.RunTest("craft missing materials leaves inventory unchanged", test_craft_missing_materials())
    Call UnitTesting.RunTest("craft insufficient materials leaves inventory unchanged", test_craft_insufficient_materials())
    Call UnitTesting.RunTest("craft later ingredient failure rolls back all deductions", test_craft_later_material_missing())
    Call UnitTesting.RunTest("craft exact and split stacks consume only ingredients", test_craft_split_stacks())
    Call UnitTesting.RunTest("craft cannot reuse consumed materials", test_craft_repeated_request())
    Call UnitTesting.RunTest("craft tagged and locked stacks do not pay normal costs", test_craft_ineligible_stacks())
    Call UnitTesting.RunTest("craft carpentry multiplies the full batch cost", test_craft_batch_quantity())
    Call UnitTesting.RunTest("craft generic carpentry supports batches and reusable tools", test_craft_generic_carpentry_materials())
    Call UnitTesting.RunTest("craft generic blacksmith materials override legacy costs", test_craft_generic_blacksmith_materials())
    Call UnitTesting.RunTest("craft generic tailoring materials override legacy costs", test_craft_generic_tailoring_materials())
    Call UnitTesting.RunTest("craft empty negative and invalid recipes fail closed", test_craft_invalid_recipe())
    Call UnitTesting.RunTest("craft excessive quantity cannot overflow costs", test_craft_quantity_overflow())
    Call UnitTesting.RunTest("craft corrupt material stacks fail without mutation", test_craft_corrupt_stack())
    Call UnitTesting.RunTest("craft all profession consumables are deducted", test_craft_profession_materials())
    Call UnitTesting.RunTest("alchemy requires and preserves reusable mortar", test_craft_reusable_mortar())
    Call UnitTesting.RunTest("object data hot reload preserves generic recipes", test_obj_data_hot_reload())
    Call UnitTesting.RunTest("craft reusable material survives either recipe order", test_reusable_order())
    Call UnitTesting.RunTest("craft equipped reusable tools are preserved", test_equipped_reusable())
    Call UnitTesting.RunTest("craft malformed authored costs fail closed", test_invalid_authored_recipe())
    Dim scenario As Integer, names As Variant
    names = Array("commission batch", "dead", "NPC trade", "player trade", "stunned", "resting", _
        "cooldown", "price bounds", "legacy crafting window", "NPC window escrow", "personal batch pacing", _
        "malformed packets", "carpentry cancellation", "pot returns escrow", "large legacy carpentry batch", _
        "commission cannot use legacy execution")
    For scenario = 1 To 16
        Call UnitTesting.RunTest("craft request: " & names(scenario - 1), test_crafting_request(scenario))
    Next scenario
    test_suite_crafting = (UnitTesting.GetFailedTests() = failuresBefore)
End Function

' Material tests use local values. Request fixtures restore globals and capture all sends.
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

Private Function test_loaded_generic_recipes() As Boolean
    On Error GoTo Failed
    Dim ini As New clsIniManager, recipe As t_ObjData, objectType As Integer
    Call ini.ChangeValue("OBJTEST", "MaterialCount", "2")
    Call ini.ChangeValue("OBJTEST", "Material1", CStr(Wood))
    Call ini.ChangeValue("OBJTEST", "Cantidad1", "3")
    Call ini.ChangeValue("OBJTEST", "Material2", CStr(Mortero))
    Call ini.ChangeValue("OBJTEST", "Cantidad2", "1")
    Call ini.ChangeValue("OBJTEST", "Consumir2", "0")
    For objectType = e_OBJType.otUseOnce To e_OBJType.otElse
        recipe.OBJType = objectType
        Call LoadGenericCraftingRecipe(recipe, ini, "OBJTEST")
        If recipe.GenericMaterialCount <> 2 Then Exit Function
        If recipe.GenericMaterialObjIndex(1) <> Wood Or recipe.GenericMaterialAmount(1) <> 3 Then Exit Function
        If Not recipe.GenericMaterialConsumed(1) Then Exit Function
        If recipe.GenericMaterialObjIndex(2) <> Mortero Or recipe.GenericMaterialAmount(2) <> 1 Then Exit Function
        If recipe.GenericMaterialConsumed(2) Then Exit Function
    Next objectType
    test_loaded_generic_recipes = True
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_loaded_generic_recipes", Erl)
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

Private Function test_craft_generic_carpentry_materials() As Boolean
    On Error GoTo test_craft_generic_carpentry_materials_Err
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    recipe.GenericMaterialCount = 2
    recipe.GenericMaterialObjIndex(1) = Wood
    recipe.GenericMaterialAmount(1) = 3
    recipe.GenericMaterialConsumed(1) = True
    recipe.GenericMaterialObjIndex(2) = Mortero
    recipe.GenericMaterialAmount(2) = 1
    recipe.GenericMaterialConsumed(2) = False
    Call CraftStack(inventory, 1, Wood, 6)
    original = inventory
    If TryConsumeCraftingMaterials(inventory, 1, recipe, e_JobsTypes.Carpenter, 2) Then Exit Function
    If Not SameCraftInventory(inventory, original) Then Exit Function
    Call CraftStack(inventory, 2, Mortero, 1)
    If Not TryConsumeCraftingMaterials(inventory, 2, recipe, e_JobsTypes.Carpenter, 2) Then Exit Function
    If inventory.NroItems <> 1 Then Exit Function
    test_craft_generic_carpentry_materials = (inventory.Object(2).ObjIndex = Mortero And inventory.Object(2).amount = 1)
    Exit Function
test_craft_generic_carpentry_materials_Err:
    test_craft_generic_carpentry_materials = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_generic_carpentry_materials", Erl)
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

Private Function test_craft_generic_blacksmith_materials() As Boolean
    On Error GoTo test_craft_generic_blacksmith_materials_Err
    Dim inventory As t_Inventario, recipe As t_ObjData
    recipe.LingH = 99
    recipe.GenericMaterialCount = 2
    recipe.GenericMaterialObjIndex(1) = e_Minerales.LingoteDeHierro
    recipe.GenericMaterialAmount(1) = 3
    recipe.GenericMaterialConsumed(1) = True
    recipe.GenericMaterialObjIndex(2) = Mortero
    recipe.GenericMaterialAmount(2) = 1
    recipe.GenericMaterialConsumed(2) = False
    Call CraftStack(inventory, 1, e_Minerales.LingoteDeHierro, 3)
    Call CraftStack(inventory, 2, Mortero, 1)
    If Not TryConsumeCraftingMaterials(inventory, 2, recipe, e_JobsTypes.Blacksmith, 1) Then Exit Function
    If inventory.NroItems <> 1 Then Exit Function
    test_craft_generic_blacksmith_materials = (inventory.Object(2).ObjIndex = Mortero And inventory.Object(2).amount = 1)
    Exit Function
test_craft_generic_blacksmith_materials_Err:
    test_craft_generic_blacksmith_materials = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_generic_blacksmith_materials", Erl)
End Function

Private Function test_craft_generic_tailoring_materials() As Boolean
    On Error GoTo test_craft_generic_tailoring_materials_Err
    Dim inventory As t_Inventario, recipe As t_ObjData
    recipe.PielLobo = 99
    recipe.GenericMaterialCount = 2
    recipe.GenericMaterialObjIndex(1) = PieldeLobo
    recipe.GenericMaterialAmount(1) = 3
    recipe.GenericMaterialConsumed(1) = True
    recipe.GenericMaterialObjIndex(2) = Mortero
    recipe.GenericMaterialAmount(2) = 1
    recipe.GenericMaterialConsumed(2) = False
    Call CraftStack(inventory, 1, PieldeLobo, 3)
    Call CraftStack(inventory, 2, Mortero, 1)
    If Not TryConsumeCraftingMaterials(inventory, 2, recipe, e_JobsTypes.Tailor, 1) Then Exit Function
    If inventory.NroItems <> 1 Then Exit Function
    test_craft_generic_tailoring_materials = (inventory.Object(2).ObjIndex = Mortero And inventory.Object(2).amount = 1)
    Exit Function
test_craft_generic_tailoring_materials_Err:
    test_craft_generic_tailoring_materials = False
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_craft_generic_tailoring_materials", Erl)
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

Private Function test_obj_data_hot_reload() As Boolean
    On Error GoTo Failed
    Dim ini As New clsIniManager, recipe As t_ObjData, slot As Integer
    Call ini.ChangeValue("OBJTEST", "MaterialCount", "10")
    For slot = 1 To 10
        Call ini.ChangeValue("OBJTEST", "Material" & slot, CStr(Wood))
        Call ini.ChangeValue("OBJTEST", "Cantidad" & slot, "1")
    Next slot
    Call LoadGenericCraftingRecipe(recipe, ini, "OBJTEST")
    If recipe.GenericMaterialCount <> 10 Then Exit Function
    Call ini.ChangeValue("OBJTEST", "MaterialCount", "1")
    Call LoadGenericCraftingRecipe(recipe, ini, "OBJTEST")
    If recipe.GenericMaterialCount <> 1 Or recipe.GenericMaterialObjIndex(2) <> 0 Then Exit Function
    Call LoadGenericCraftingRecipe(recipe, ini, "ABSENT")
    If recipe.GenericMaterialCount <> 0 Or recipe.GenericMaterialObjIndex(1) <> 0 Then Exit Function
    test_obj_data_hot_reload = True
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_obj_data_hot_reload", Erl)
End Function


Private Function test_reusable_order() As Boolean
    On Error GoTo Failed
    Dim inventory As t_Inventario, original As t_Inventario, recipe As t_ObjData
    Dim job As Variant, consumedSlot As Integer
    For Each job In Array(e_JobsTypes.Blacksmith, e_JobsTypes.Carpenter, e_JobsTypes.Alchemist, e_JobsTypes.Tailor)
        For consumedSlot = 1 To 2
            recipe.GenericMaterialCount = 2
            recipe.GenericMaterialObjIndex(1) = Wood
            recipe.GenericMaterialObjIndex(2) = Wood
            recipe.GenericMaterialAmount(1) = 1
            recipe.GenericMaterialAmount(2) = 1
            recipe.GenericMaterialConsumed(1) = (consumedSlot = 1)
            recipe.GenericMaterialConsumed(2) = (consumedSlot = 2)
            inventory.NroItems = 0
            Call CraftStack(inventory, 1, Wood, 1)
            original = inventory
            Call Expect(Not TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, CLng(job), 1), "last reusable material cannot be consumed")
            Call Expect(SameCraftInventory(original, inventory), "rejected recipe is atomic")
            inventory.Object(1).amount = 2
            Call Expect(TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, CLng(job), 1), "separate reusable quantity permits craft")
            Call Expect(inventory.Object(1).amount = 1, "one reusable material remains")
        Next consumedSlot
    Next job
    test_reusable_order = True
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_reusable_order", Erl)
End Function

Private Function test_equipped_reusable() As Boolean
    On Error GoTo Failed
    Dim inventory As t_Inventario, recipe As t_ObjData
    Call CraftStack(inventory, 1, Wood, 10)
    Call CraftStack(inventory, 2, Mortero, 1)
    inventory.Object(2).Equipped = 1
    inventory.EquippedWorkingToolObjIndex = Mortero
    inventory.EquippedWorkingToolSlot = 2
    recipe.GenericMaterialCount = 2
    recipe.GenericMaterialObjIndex(1) = Wood
    recipe.GenericMaterialAmount(1) = 1
    recipe.GenericMaterialConsumed(1) = True
    recipe.GenericMaterialObjIndex(2) = Mortero
    recipe.GenericMaterialAmount(2) = 1
    Call Expect(TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, e_JobsTypes.Tailor, 1), "equipped reusable tool accepted")
    Call Expect(inventory.Object(2).amount = 1 And inventory.Object(2).Equipped = 1, "equipped tool retained")
    Call Expect(inventory.EquippedWorkingToolSlot = 2 And inventory.EquippedWorkingToolObjIndex = Mortero, "equipment references retained")
    test_equipped_reusable = True
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_equipped_reusable", Erl)
End Function

Private Function test_invalid_authored_recipe() As Boolean
    On Error GoTo Failed
    Dim ini As New clsIniManager, recipe As t_ObjData, inventory As t_Inventario, original As t_Inventario
    Dim field As Variant
    Call CraftStack(inventory, 1, Wood, 100)
    original = inventory
    For Each field In Array("MaterialCount", "Material1", "Cantidad1", "Consumir1")
        Call ini.ChangeValue("OBJTEST", "MaterialCount", "1")
        Call ini.ChangeValue("OBJTEST", "Material1", CStr(Wood))
        Call ini.ChangeValue("OBJTEST", "Cantidad1", "1")
        Call ini.ChangeValue("OBJTEST", "Consumir1", "1")
        Call ini.ChangeValue("OBJTEST", CStr(field), "99999")
        recipe.Madera = 1
        Call LoadGenericCraftingRecipe(recipe, ini, "OBJTEST")
        Call Expect(recipe.GenericMaterialCount > MAX_GENERIC_CRAFTING_MATERIALS, "malformed definition is disabled")
        Call Expect(Not TryConsumeCraftingMaterials(inventory, MAX_INVENTORY_SLOTS, recipe, e_JobsTypes.Carpenter, 1), "no truncation or legacy fallback")
        Call Expect(SameCraftInventory(inventory, original), "invalid definition never consumes")
    Next field
    Call ini.ChangeValue("OBJTEST", "Consumir1", "typo")
    Call LoadGenericCraftingRecipe(recipe, ini, "OBJTEST")
    Call Expect(recipe.GenericMaterialCount > MAX_GENERIC_CRAFTING_MATERIALS, "typo must not turn a consumed material into a reusable tool")
    test_invalid_authored_recipe = True
    Exit Function
Failed:
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.test_invalid_authored_recipe", Erl)
End Function

Private Sub SendCraftRequest(ByVal user As Integer, ByVal alchemy As Boolean, Optional ByVal quantity As Long = 10, Optional ByVal lengthDelta As Integer = 0)
    Dim writer As New Network.writer, payload() As Byte, encoded() As Byte, i As Integer
    Call writer.WriteInt16(2)
    Call writer.WriteInt32(quantity)
    Call writer.GetData(encoded)
    ReDim payload(0 To 5 + lengthDelta)
    For i = 0 To UBound(payload)
        If i < 6 Then payload(i) = encoded(LBound(encoded) + i)
    Next i
    Call Protocol.CraftingTestRequest(user, payload, alchemy)
End Sub

Private Function PacketIdAt(ByVal index As Long) As Integer
    Dim bytes() As Byte, reader As New Network.reader
    bytes = packets(index)
    Call reader.SetData(bytes)
    PacketIdAt = reader.ReadInt16()
End Function

Private Function test_crafting_request(ByVal scenario As Integer) As Boolean
    On Error GoTo Failed
    Dim user As Integer, npc As Integer, i As Integer
    Dim savedUsers() As t_User, savedNpc As t_Npc, savedObjects() As t_ObjData
    Dim emptyUser As t_User, emptyObject As t_ObjData
    Dim savedRecipes As Dictionary, recipes As New Dictionary, recipe As New clsCrafteo
    Dim savedAlchemy() As Integer, savedInterval As Long, fixtureReady As Boolean
    user = UBound(UserList): npc = UBound(NpcList)
    ReDim savedUsers(1 To 1): ReDim savedObjects(1 To 3)
    savedUsers(1) = UserList(user): savedNpc = NpcList(npc)
    For i = 1 To 3
        savedObjects(i) = ObjData(i)
    Next i
    Set savedRecipes = Crafteos
    savedAlchemy = ObjAlquimista
    savedInterval = IntervaloTrabajarConstruir
    fixtureReady = True
    capturing = True
    Set packets = New Collection
    UserList(user) = emptyUser
    For i = 1 To 3
        ObjData(i) = emptyObject
    Next i
    ObjData(1).OBJType = e_OBJType.otUseOnce
    ObjData(2).OBJType = e_OBJType.otPotions
    ObjData(2).GenericMaterialCount = 1
    ObjData(2).GenericMaterialObjIndex(1) = 1
    ObjData(2).GenericMaterialAmount(1) = 1
    ObjData(2).GenericMaterialConsumed(1) = True
    ObjData(3).OBJType = e_OBJType.otWorkingTools
    ObjData(3).Subtipo = 4
    With UserList(user)
        .name = "crafting fixture"
        .VersionId = 7
        .flags.UserLogged = True
        .flags.Privilegios = e_PlayerType.User
        .clase = e_Class.Mage
        .raza = e_Raza.Humano
        .CurrentInventorySlots = MAX_INVENTORY_SLOTS
        .pos.Map = 1: .pos.x = 50: .pos.y = 50
        .Stats.GLD = 1000: .Stats.MinSta = 100: .Stats.ELV = 1
        .Stats.UserSkills(e_Skill.Alquimia) = 100
        .Counters.StunEndTime = GetTickCountRaw()
        .flags.Crafteando = NPC_ALCHEMY_CRAFTING_TYPE
        Call CraftStack(.invent, 1, 1, 20)
    End With
    NpcList(npc).VersionId = 17
    NpcList(npc).Craftea = NPC_ALCHEMY_CRAFTING_TYPE
    NpcList(npc).pos = UserList(user).pos
    Call SetNpcRef(UserList(user).flags.TargetNPC, npc)
    Set Crafteos = New Dictionary
    recipe.Resultado = 2: recipe.precio = 2
    Call recipes.Add("fixture", recipe)
    Call Crafteos.Add(NPC_ALCHEMY_CRAFTING_TYPE, recipes)
    ReDim ObjAlquimista(1 To 1)
    ObjAlquimista(1) = 2
    IntervaloTrabajarConstruir = 0

    Select Case scenario
        Case 1
            Call SendCraftRequest(user, True)
            Call Expect(UserList(user).invent.Object(1).amount = 10, "commission consumes exact batch")
            Call Expect(UserList(user).Stats.GLD = 980, "commission charges exact batch")
            Call Expect(UserList(user).invent.Object(2).ObjIndex = 2, "commission output delivered")
            Call Expect(UserList(user).invent.Object(2).amount >= 7 And UserList(user).invent.Object(2).amount <= 8, "existing commission success range")
        Case 2 To 7
            Select Case scenario
                Case 2: UserList(user).flags.Muerto = 1
                Case 3: UserList(user).flags.Comerciando = True
                Case 4: Call SetUserRef(UserList(user).ComUsu.DestUsu, user)
                Case 5: UserList(user).Counters.StunEndTime = AddMod32(GetTickCountRaw(), 60000)
                Case 6: UserList(user).flags.Descansar = True
                Case 7
                    IntervaloTrabajarConstruir = 60000
                    UserList(user).Counters.TimerPuedeTrabajar = GetTickCountRaw()
            End Select
            Call SendCraftRequest(user, True)
            Call Expect(UserList(user).invent.Object(1).amount = 20 And UserList(user).Stats.GLD = 1000, "blocked commission changes no resources")
            UserList(user).clase = e_Class.Trabajador
            UserList(user).flags.Crafteando = 0
            UserList(user).invent.EquippedWorkingToolObjIndex = 3
            Call SendCraftRequest(user, True)
            Call Expect(UserList(user).invent.Object(1).amount = 20 And UserList(user).Stats.MinSta = 100, "blocked personal alchemy changes no resources")
        Case 8
            recipe.precio = -1
            Call SendCraftRequest(user, True)
            recipe.precio = 2147483647
            Call SendCraftRequest(user, True)
            Call Expect(UserList(user).invent.Object(1).amount = 20 And UserList(user).Stats.GLD = 1000, "negative and overflowing price rejected")
        Case 9
            Call OpenNpcCrafting(user, 1)
            Call Expect(packets.Count = 1 And PacketIdAt(1) = ServerPacketID.eOpenCrafting, "legacy combination window retained")
            Call SendCraftRequest(user, True)
            Call Expect(UserList(user).invent.Object(1).amount = 20, "legacy type cannot authorize commission")
        Case 10
            UserList(user).CraftInventory(1) = 1
            UserList(user).CraftCatalyst.ObjIndex = 1
            UserList(user).CraftCatalyst.amount = 2
            Call OpenNpcCrafting(user, NPC_ALCHEMY_CRAFTING_TYPE)
            Call Expect(UserList(user).invent.Object(1).amount = 23, "escrow and catalyst returned before mode switch")
            Call Expect(UserList(user).CraftInventory(1) = 0 And UserList(user).CraftCatalyst.amount = 0, "escrow cleared exactly once")
            Call Expect(PacketIdAt(packets.Count - 1) = ServerPacketID.eAlquimistaObj, "NPC catalogue precedes window")
            Call Expect(PacketIdAt(packets.Count) = ServerPacketID.eShowAlquimiaForm, "NPC alchemy window")
        Case 11
            UserList(user).clase = e_Class.Trabajador
            UserList(user).flags.Crafteando = 0
            UserList(user).invent.EquippedWorkingToolObjIndex = 3
            Call SendCraftRequest(user, True)
            Call Expect(UserList(user).invent.Object(1).amount = 10 And UserList(user).Stats.MinSta = 99, "personal batch cost and stamina")
            IntervaloTrabajarConstruir = 60000
            Call SendCraftRequest(user, True)
            Call Expect(UserList(user).invent.Object(1).amount = 10 And UserList(user).Stats.MinSta = 99, "repeat respects committed work interval")
        Case 12
            Call SendCraftRequest(user, True, 10, -1)
            Call SendCraftRequest(user, True, 10, 1)
            Call SendCraftRequest(user, False, 10, -1)
            Call SendCraftRequest(user, False, 10, 1)
            Call Expect(UserList(user).invent.Object(1).amount = 20 And UserList(user).Stats.GLD = 1000, "malformed packet cannot mutate")
            Call Expect(Not UserList(user).flags.UsandoMacro, "malformed packet cannot start macro")
        Case 13
            UserList(user).Trabajo.Cantidad = 10
            UserList(user).flags.UsandoMacro = True
            UserList(user).flags.Muerto = 1
            Call SendCraftRequest(user, False, 0)
            Call Expect(Not UserList(user).flags.UsandoMacro And UserList(user).Trabajo.Cantidad = 0, "explicit cancellation remains available")
        Case 14
            UserList(user).clase = e_Class.Trabajador
            UserList(user).CraftInventory(1) = 1
            Call CraftStack(UserList(user).invent, 2, 3, 1)
            UserList(user).invent.Object(2).Equipped = 1
            UserList(user).invent.EquippedWorkingToolObjIndex = 3
            UserList(user).invent.EquippedWorkingToolSlot = 2
            UserList(user).Counters.TimerUsar = AddMod32(GetTickCountRaw(), -60000)
            Call UseInvItem(user, 2, 0)
            Call Expect(UserList(user).flags.Crafteando = 0, "pot opens personal alchemy")
            Call Expect(UserList(user).CraftInventory(1) = 0 And UserList(user).invent.Object(1).amount = 21, "pot returns legacy escrow")
        Case 15
            ObjData(2).GenericMaterialCount = 0
            ObjData(2).Madera = 5
            For i = 1 To 5
                UserList(user).invent.Object(i).ObjIndex = Wood
                UserList(user).invent.Object(i).amount = 10000
            Next i
            UserList(user).invent.NroItems = 5
            Call SendCraftRequest(user, False, 10000)
            Call Expect(UserList(user).Trabajo.Cantidad = 10000 And UserList(user).flags.UsandoMacro, "50000 wood batch validated without Integer narrowing")
            Call Expect(UserList(user).invent.Object(1).amount = 10000, "batch validation does not consume early")
        Case 16
            Set UserList(user).CraftResult = recipe
            Call Protocol.HandleCraftItem(user)
            Call Expect(UserList(user).Stats.GLD = 1000 And UserList(user).invent.Object(1).amount = 20, "legacy execution cannot bypass commission costs or validation")
            Call recipes.Add(GetRecipeKey(UserList(user).CraftInventory), recipe)
            Call Expect(CheckCraftingResult(user) Is Nothing, "commission does not expose a legacy crafting result")
    End Select
    test_crafting_request = True
Cleanup:
    capturing = False
    Set packets = Nothing
    If fixtureReady Then
        UserList(user) = savedUsers(1)
        NpcList(npc) = savedNpc
        For i = 1 To 3
            ObjData(i) = savedObjects(i)
        Next i
        Set Crafteos = savedRecipes
        ObjAlquimista = savedAlchemy
        IntervaloTrabajarConstruir = savedInterval
    End If
    Exit Function
Failed:
    Call UnitTesting.Progress("craft request " & scenario & ": " & Err.Number & "; " & Err.Description)
    Call TraceError(Err.Number, Err.Description, "Unit_Crafting.request " & scenario, Erl)
    Resume Cleanup
End Function
#End If
