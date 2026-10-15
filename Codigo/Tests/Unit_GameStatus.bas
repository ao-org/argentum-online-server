Attribute VB_Name = "Unit_GameStatus"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_gamestatus() As Boolean
    Call UnitTesting.RunTest("test_esnewbie_below", test_esnewbie_below())
    Call UnitTesting.RunTest("test_esnewbie_at_limit", test_esnewbie_at_limit())
    Call UnitTesting.RunTest("test_esnewbie_above", test_esnewbie_above())
    Call UnitTesting.RunTest("test_esnewbie_zero_index", test_esnewbie_zero_index())
    Call UnitTesting.RunTest("test_esciudadano_true", test_esciudadano_true())
    Call UnitTesting.RunTest("test_esciudadano_false", test_esciudadano_false())
    Call UnitTesting.RunTest("test_escriminal_true", test_escriminal_true())
    Call UnitTesting.RunTest("test_escriminal_false", test_escriminal_false())
    Call UnitTesting.RunTest("test_esarmada_true", test_esarmada_true())
    Call UnitTesting.RunTest("test_esarmada_consejo", test_esarmada_consejo())
    Call UnitTesting.RunTest("test_esarmada_false", test_esarmada_false())
    Call UnitTesting.RunTest("test_escaos_true", test_escaos_true())
    Call UnitTesting.RunTest("test_escaos_concilio", test_escaos_concilio())
    Call UnitTesting.RunTest("test_escaos_false", test_escaos_false())
    Call UnitTesting.RunTest("test_faction_zero_index", test_faction_zero_index())
    Call UnitTesting.RunTest("test_esgm_admin", test_esgm_admin())
    Call UnitTesting.RunTest("test_esgm_dios", test_esgm_dios())
    Call UnitTesting.RunTest("test_esgm_semidios", test_esgm_semidios())
    Call UnitTesting.RunTest("test_esgm_consejero", test_esgm_consejero())
    Call UnitTesting.RunTest("test_esgm_no_privs", test_esgm_no_privs())
    Call UnitTesting.RunTest("test_esgm_zero_index", test_esgm_zero_index())
    Call UnitTesting.RunTest("test_esnewbie_threshold_property", test_esnewbie_threshold_property())
    Call UnitTesting.RunTest("test_non_newbie_can_use_newbie_item", test_non_newbie_can_use_newbie_item())
    Call UnitTesting.RunTest("test_newbie_graduation_preserves_character_state", test_newbie_graduation_preserves_character_state())
    Call UnitTesting.RunTest("test_newbie_item_normal_requirements_still_apply", test_newbie_item_normal_requirements_still_apply())
    Call UnitTesting.RunTest("test_newbie_bank_deposit_is_rejected", test_newbie_bank_deposit_is_rejected())
    Call UnitTesting.RunTest("test_normal_bank_deposit_succeeds", test_normal_bank_deposit_succeeds())
    Call UnitTesting.RunTest("test_newbie_death_drop_protection", test_newbie_death_drop_protection())
    test_suite_gamestatus = True
End Function

Private Function test_newbie_bank_deposit_is_rejected() As Boolean
    On Error GoTo Err_Handler

    Dim OriginalObject            As t_ObjData
    Dim OriginalInventory         As t_UserOBJ
    Dim OriginalBank              As t_UserOBJ
    Dim OriginalInventorySlots    As Byte
    Dim OriginalInventoryCount    As Byte
    Dim OriginalBankCount         As Byte
    Dim OriginalInventoryModified As Boolean
    Dim OriginalBankModified      As Boolean
    Dim DepositedSlot             As Long

    OriginalObject = ObjData(1)
    With UserList(1)
        OriginalInventory = .invent.Object(1)
        OriginalBank = .BancoInvent.Object(1)
        OriginalInventorySlots = .CurrentInventorySlots
        OriginalInventoryCount = .invent.NroItems
        OriginalBankCount = .BancoInvent.NroItems
        OriginalInventoryModified = .flags.ModificoInventario
        OriginalBankModified = .flags.ModificoInventarioBanco

        .CurrentInventorySlots = 1
        .invent.Object(1).ObjIndex = 1
        .invent.Object(1).amount = 3
        .invent.Object(1).Equipped = 1
        .invent.Object(1).ElementalTags = e_ElementalTags.Fire
        .invent.NroItems = 1
        .BancoInvent.Object(1).ObjIndex = 0
        .BancoInvent.Object(1).amount = 0
        .BancoInvent.Object(1).Equipped = 0
        .BancoInvent.Object(1).ElementalTags = 0
        .BancoInvent.NroItems = 0
        .flags.ModificoInventario = False
        .flags.ModificoInventarioBanco = False
    End With
    ObjData(1).Newbie = 1

    DepositedSlot = UserDejaObj(1, 1, 2, 1)

    With UserList(1)
        test_newbie_bank_deposit_is_rejected = DepositedSlot = 0 _
                And .invent.Object(1).ObjIndex = 1 _
                And .invent.Object(1).amount = 3 _
                And .invent.Object(1).Equipped = 1 _
                And .invent.Object(1).ElementalTags = e_ElementalTags.Fire _
                And .invent.NroItems = 1 _
                And .BancoInvent.Object(1).ObjIndex = 0 _
                And .BancoInvent.Object(1).amount = 0 _
                And .BancoInvent.NroItems = 0 _
                And Not .flags.ModificoInventario _
                And Not .flags.ModificoInventarioBanco
    End With

Clean_Up:
    ObjData(1) = OriginalObject
    With UserList(1)
        .invent.Object(1) = OriginalInventory
        .BancoInvent.Object(1) = OriginalBank
        .CurrentInventorySlots = OriginalInventorySlots
        .invent.NroItems = OriginalInventoryCount
        .BancoInvent.NroItems = OriginalBankCount
        .flags.ModificoInventario = OriginalInventoryModified
        .flags.ModificoInventarioBanco = OriginalBankModified
    End With
    Exit Function
Err_Handler:
    test_newbie_bank_deposit_is_rejected = False
    Resume Clean_Up
End Function

Private Function test_normal_bank_deposit_succeeds() As Boolean
    On Error GoTo Err_Handler

    Dim OriginalObject         As t_ObjData
    Dim OriginalInventory      As t_UserOBJ
    Dim OriginalBank           As t_UserOBJ
    Dim OriginalInventorySlots As Byte
    Dim OriginalInventoryCount As Byte
    Dim OriginalBankCount      As Byte
    Dim OriginalInventoryModified As Boolean
    Dim OriginalBankModified      As Boolean
    Dim DepositedSlot          As Long

    OriginalObject = ObjData(1)
    With UserList(1)
        OriginalInventory = .invent.Object(1)
        OriginalBank = .BancoInvent.Object(1)
        OriginalInventorySlots = .CurrentInventorySlots
        OriginalInventoryCount = .invent.NroItems
        OriginalBankCount = .BancoInvent.NroItems
        OriginalInventoryModified = .flags.ModificoInventario
        OriginalBankModified = .flags.ModificoInventarioBanco

        .CurrentInventorySlots = 1
        .invent.Object(1).ObjIndex = 1
        .invent.Object(1).amount = 3
        .invent.Object(1).Equipped = 0
        .invent.Object(1).ElementalTags = e_ElementalTags.Water
        .invent.NroItems = 1
        .BancoInvent.Object(1).ObjIndex = 0
        .BancoInvent.Object(1).amount = 0
        .BancoInvent.Object(1).Equipped = 0
        .BancoInvent.Object(1).ElementalTags = 0
        .BancoInvent.NroItems = 0
        .flags.ModificoInventario = False
        .flags.ModificoInventarioBanco = False
    End With
    ObjData(1).Newbie = 0

    DepositedSlot = UserDejaObj(1, 1, 2, 1)

    With UserList(1)
        test_normal_bank_deposit_succeeds = DepositedSlot = 1 _
                And .invent.Object(1).ObjIndex = 1 _
                And .invent.Object(1).amount = 1 _
                And .BancoInvent.Object(1).ObjIndex = 1 _
                And .BancoInvent.Object(1).amount = 2 _
                And .BancoInvent.Object(1).ElementalTags = e_ElementalTags.Water _
                And .BancoInvent.NroItems = 1
    End With

Clean_Up:
    ObjData(1) = OriginalObject
    With UserList(1)
        .invent.Object(1) = OriginalInventory
        .BancoInvent.Object(1) = OriginalBank
        .CurrentInventorySlots = OriginalInventorySlots
        .invent.NroItems = OriginalInventoryCount
        .BancoInvent.NroItems = OriginalBankCount
        .flags.ModificoInventario = OriginalInventoryModified
        .flags.ModificoInventarioBanco = OriginalBankModified
    End With
    Exit Function
Err_Handler:
    test_normal_bank_deposit_succeeds = False
    Resume Clean_Up
End Function

Private Function test_newbie_death_drop_protection() As Boolean
    On Error GoTo Err_Handler

    Dim OriginalLevel  As Byte
    Dim OriginalObject As t_ObjData

    OriginalLevel = UserList(1).Stats.ELV
    OriginalObject = ObjData(1)
    ObjData(1).Newbie = 1

    UserList(1).Stats.ELV = LimiteNewbie
    If Not ItemNewbieProtegidoAlMorir(1, 1) Then GoTo Clean_Up

    UserList(1).Stats.ELV = LimiteNewbie + 1
    If ItemNewbieProtegidoAlMorir(1, 1) Then GoTo Clean_Up

    ObjData(1).Newbie = 0
    If ItemNewbieProtegidoAlMorir(1, 1) Then GoTo Clean_Up

    test_newbie_death_drop_protection = True

Clean_Up:
    UserList(1).Stats.ELV = OriginalLevel
    ObjData(1) = OriginalObject
    Exit Function
Err_Handler:
    test_newbie_death_drop_protection = False
    Resume Clean_Up
End Function

' ============================================================
' EsNewbie Tests
' ============================================================

' EsNewbie returns True when ELV is below LimiteNewbie (12).
Private Function test_esnewbie_below() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_below = True

    Dim origELV As Byte
    origELV = UserList(1).Stats.ELV

    UserList(1).Stats.ELV = 5
    If Not EsNewbie(1) Then test_esnewbie_below = False

    UserList(1).Stats.ELV = origELV
    Exit Function
Err_Handler:
    UserList(1).Stats.ELV = origELV
    test_esnewbie_below = False
End Function

' EsNewbie returns True when ELV equals LimiteNewbie (12).
Private Function test_esnewbie_at_limit() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_at_limit = True

    Dim origELV As Byte
    origELV = UserList(1).Stats.ELV

    UserList(1).Stats.ELV = 12
    If Not EsNewbie(1) Then test_esnewbie_at_limit = False

    UserList(1).Stats.ELV = origELV
    Exit Function
Err_Handler:
    UserList(1).Stats.ELV = origELV
    test_esnewbie_at_limit = False
End Function

' EsNewbie returns False when ELV is above LimiteNewbie (12).
Private Function test_esnewbie_above() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_above = True

    Dim origELV As Byte
    origELV = UserList(1).Stats.ELV

    UserList(1).Stats.ELV = 13
    If EsNewbie(1) Then test_esnewbie_above = False

    UserList(1).Stats.ELV = origELV
    Exit Function
Err_Handler:
    UserList(1).Stats.ELV = origELV
    test_esnewbie_above = False
End Function

' EsNewbie returns False for UserIndex 0.
Private Function test_esnewbie_zero_index() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_zero_index = True

    If EsNewbie(0) Then test_esnewbie_zero_index = False

    Exit Function
Err_Handler:
    test_esnewbie_zero_index = False
End Function

' ============================================================
' Faction Status Helper Tests
' ============================================================

' esCiudadano returns True for Ciudadano status.
Private Function test_esciudadano_true() As Boolean
    On Error GoTo Err_Handler
    test_esciudadano_true = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Ciudadano
    If Not esCiudadano(1) Then test_esciudadano_true = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esciudadano_true = False
End Function

' esCiudadano returns False for Criminal status.
Private Function test_esciudadano_false() As Boolean
    On Error GoTo Err_Handler
    test_esciudadano_false = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Criminal
    If esCiudadano(1) Then test_esciudadano_false = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esciudadano_false = False
End Function

' esCriminal returns True for Criminal status.
Private Function test_escriminal_true() As Boolean
    On Error GoTo Err_Handler
    test_escriminal_true = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Criminal
    If Not esCriminal(1) Then test_escriminal_true = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escriminal_true = False
End Function

' esCriminal returns False for Ciudadano status.
Private Function test_escriminal_false() As Boolean
    On Error GoTo Err_Handler
    test_escriminal_false = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Ciudadano
    If esCriminal(1) Then test_escriminal_false = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escriminal_false = False
End Function

' esArmada returns True for Armada status.
Private Function test_esarmada_true() As Boolean
    On Error GoTo Err_Handler
    test_esarmada_true = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Armada
    If Not esArmada(1) Then test_esarmada_true = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esarmada_true = False
End Function

' esArmada returns True for consejo (allied with Armada).
Private Function test_esarmada_consejo() As Boolean
    On Error GoTo Err_Handler
    test_esarmada_consejo = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.consejo
    If Not esArmada(1) Then test_esarmada_consejo = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esarmada_consejo = False
End Function

' esArmada returns False for Caos status.
Private Function test_esarmada_false() As Boolean
    On Error GoTo Err_Handler
    test_esarmada_false = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Caos
    If esArmada(1) Then test_esarmada_false = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_esarmada_false = False
End Function

' esCaos returns True for Caos status.
Private Function test_escaos_true() As Boolean
    On Error GoTo Err_Handler
    test_escaos_true = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Caos
    If Not esCaos(1) Then test_escaos_true = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escaos_true = False
End Function

' esCaos returns True for concilio (allied with Caos).
Private Function test_escaos_concilio() As Boolean
    On Error GoTo Err_Handler
    test_escaos_concilio = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.concilio
    If Not esCaos(1) Then test_escaos_concilio = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escaos_concilio = False
End Function

' esCaos returns False for Armada status.
Private Function test_escaos_false() As Boolean
    On Error GoTo Err_Handler
    test_escaos_false = True

    Dim origStatus As Byte
    origStatus = UserList(1).Faccion.Status

    UserList(1).Faccion.Status = e_Facciones.Armada
    If esCaos(1) Then test_escaos_false = False

    UserList(1).Faccion.Status = origStatus
    Exit Function
Err_Handler:
    UserList(1).Faccion.Status = origStatus
    test_escaos_false = False
End Function

' All faction helpers return False for UserIndex 0.
Private Function test_faction_zero_index() As Boolean
    On Error GoTo Err_Handler
    test_faction_zero_index = True

    If esCiudadano(0) Then test_faction_zero_index = False: Exit Function
    If esCriminal(0) Then test_faction_zero_index = False: Exit Function
    If esArmada(0) Then test_faction_zero_index = False: Exit Function
    If esCaos(0) Then test_faction_zero_index = False: Exit Function

    Exit Function
Err_Handler:
    test_faction_zero_index = False
End Function

' ============================================================
' EsGM Privilege Tests
' ============================================================

' EsGM returns True for Admin privilege.
Private Function test_esgm_admin() As Boolean
    On Error GoTo Err_Handler
    test_esgm_admin = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = e_PlayerType.Admin
    If Not EsGM(1) Then test_esgm_admin = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_admin = False
End Function

' EsGM returns True for Dios privilege.
Private Function test_esgm_dios() As Boolean
    On Error GoTo Err_Handler
    test_esgm_dios = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = e_PlayerType.Dios
    If Not EsGM(1) Then test_esgm_dios = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_dios = False
End Function

' EsGM returns True for SemiDios privilege.
Private Function test_esgm_semidios() As Boolean
    On Error GoTo Err_Handler
    test_esgm_semidios = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = e_PlayerType.SemiDios
    If Not EsGM(1) Then test_esgm_semidios = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_semidios = False
End Function

' EsGM returns True for Consejero privilege.
Private Function test_esgm_consejero() As Boolean
    On Error GoTo Err_Handler
    test_esgm_consejero = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = e_PlayerType.Consejero
    If Not EsGM(1) Then test_esgm_consejero = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_consejero = False
End Function

' EsGM returns False when Privilegios is 0 (no privileges).
Private Function test_esgm_no_privs() As Boolean
    On Error GoTo Err_Handler
    test_esgm_no_privs = True

    Dim origPrivs As Long
    origPrivs = UserList(1).flags.Privilegios

    UserList(1).flags.Privilegios = 0
    If EsGM(1) Then test_esgm_no_privs = False

    UserList(1).flags.Privilegios = origPrivs
    Exit Function
Err_Handler:
    UserList(1).flags.Privilegios = origPrivs
    test_esgm_no_privs = False
End Function

' EsGM returns False for UserIndex 0.
Private Function test_esgm_zero_index() As Boolean
    On Error GoTo Err_Handler
    test_esgm_zero_index = True

    If EsGM(0) Then test_esgm_zero_index = False

    Exit Function
Err_Handler:
    test_esgm_zero_index = False
End Function

' ============================================================
' Property Test: EsNewbie Threshold
' ============================================================

' Property 5: EsNewbie threshold
' For any level 1 through 50, EsNewbie returns True iff level <= 12.
Private Function test_esnewbie_threshold_property() As Boolean
    On Error GoTo Err_Handler
    test_esnewbie_threshold_property = True

    Dim origELV As Byte
    origELV = UserList(1).Stats.ELV

    Dim lvl As Integer
    For lvl = 1 To 50
        UserList(1).Stats.ELV = CByte(lvl)

        If lvl <= 12 Then
            If Not EsNewbie(1) Then
                UserList(1).Stats.ELV = origELV
                test_esnewbie_threshold_property = False
                Exit Function
            End If
        Else
            If EsNewbie(1) Then
                UserList(1).Stats.ELV = origELV
                test_esnewbie_threshold_property = False
                Exit Function
            End If
        End If
    Next lvl

    UserList(1).Stats.ELV = origELV
    Exit Function
Err_Handler:
    UserList(1).Stats.ELV = origELV
    test_esnewbie_threshold_property = False
End Function

Private Function test_non_newbie_can_use_newbie_item() As Boolean
    On Error GoTo Err_Handler

    Dim OriginalLevel  As Byte
    Dim OriginalClass  As e_Class
    Dim OriginalRace   As e_Raza
    Dim OriginalGender As e_Genero
    Dim OriginalObject As t_ObjData
    Dim TestObject     As t_ObjData

    OriginalLevel = UserList(1).Stats.ELV
    OriginalClass = UserList(1).clase
    OriginalRace = UserList(1).raza
    OriginalGender = UserList(1).genero
    OriginalObject = ObjData(1)
    UserList(1).Stats.ELV = LimiteNewbie + 1
    UserList(1).clase = e_Class.Warrior
    UserList(1).raza = e_Raza.Humano
    UserList(1).genero = e_Genero.Hombre
    TestObject.Newbie = 1
    ObjData(1) = TestObject

    test_non_newbie_can_use_newbie_item = ObjData(1).Newbie = 1 _
            And Not EsNewbie(1) _
            And CanUseObject(1, 1) = 0

Clean_Up:
    UserList(1).Stats.ELV = OriginalLevel
    UserList(1).clase = OriginalClass
    UserList(1).raza = OriginalRace
    UserList(1).genero = OriginalGender
    ObjData(1) = OriginalObject
    Exit Function
Err_Handler:
    test_non_newbie_can_use_newbie_item = False
    Resume Clean_Up
End Function

Private Function test_newbie_graduation_preserves_character_state() As Boolean
    On Error GoTo Err_Handler

    Dim OriginalStats            As t_UserStats
    Dim OriginalCounters         As t_UserCounters
    Dim OriginalPosition         As t_WorldPos
    Dim OriginalInventory        As t_UserOBJ
    Dim OriginalBank             As t_UserOBJ
    Dim OriginalClass            As e_Class
    Dim OriginalRace             As e_Raza
    Dim OriginalGender           As e_Genero
    Dim OriginalInventorySlots   As Byte
    Dim OriginalWeaponObjIndex   As Integer
    Dim OriginalWeaponSlot       As Byte
    Dim OriginalNaked            As Byte

    With UserList(1)
        OriginalStats = .Stats
        OriginalCounters = .Counters
        OriginalPosition = .pos
        OriginalInventory = .invent.Object(1)
        OriginalBank = .BancoInvent.Object(1)
        OriginalClass = .clase
        OriginalRace = .raza
        OriginalGender = .genero
        OriginalInventorySlots = .CurrentInventorySlots
        OriginalWeaponObjIndex = .invent.EquippedWeaponObjIndex
        OriginalWeaponSlot = .invent.EquippedWeaponSlot
        OriginalNaked = .flags.Desnudo

        .Stats.ELV = LimiteNewbie
        .Stats.Exp = ExpLevelUp(LimiteNewbie)
        .Stats.MaxHp = 100
        .Stats.MinHp = 100
        .Stats.UserAtributos(e_Atributos.Constitucion) = 18
        .clase = e_Class.Warrior
        .raza = e_Raza.Humano
        .genero = e_Genero.Hombre
        .pos.Map = 34
        .pos.x = 50
        .pos.y = 60
        .CurrentInventorySlots = 1
        .invent.Object(1).ObjIndex = 3487
        .invent.Object(1).amount = 1
        .invent.Object(1).Equipped = 1
        .invent.Object(1).ElementalTags = e_ElementalTags.Fire
        .invent.EquippedWeaponObjIndex = 3487
        .invent.EquippedWeaponSlot = 1
        .BancoInvent.Object(1).ObjIndex = 4335
        .BancoInvent.Object(1).amount = 4
        .BancoInvent.Object(1).ElementalTags = e_ElementalTags.Water
        .flags.Desnudo = 0

        Call CheckUserLevel(1)

        test_newbie_graduation_preserves_character_state = .Stats.ELV = LimiteNewbie + 1 _
                And .pos.Map = 34 And .pos.x = 50 And .pos.y = 60 _
                And .invent.Object(1).ObjIndex = 3487 _
                And .invent.Object(1).amount = 1 _
                And .invent.Object(1).Equipped = 1 _
                And .invent.Object(1).ElementalTags = e_ElementalTags.Fire _
                And .invent.EquippedWeaponObjIndex = 3487 _
                And .invent.EquippedWeaponSlot = 1 _
                And .BancoInvent.Object(1).ObjIndex = 4335 _
                And .BancoInvent.Object(1).amount = 4 _
                And .BancoInvent.Object(1).ElementalTags = e_ElementalTags.Water _
                And .flags.Desnudo = 0
    End With

Clean_Up:
    With UserList(1)
        .Stats = OriginalStats
        .Counters = OriginalCounters
        .pos = OriginalPosition
        .invent.Object(1) = OriginalInventory
        .BancoInvent.Object(1) = OriginalBank
        .clase = OriginalClass
        .raza = OriginalRace
        .genero = OriginalGender
        .CurrentInventorySlots = OriginalInventorySlots
        .invent.EquippedWeaponObjIndex = OriginalWeaponObjIndex
        .invent.EquippedWeaponSlot = OriginalWeaponSlot
        .flags.Desnudo = OriginalNaked
    End With
    Exit Function
Err_Handler:
    test_newbie_graduation_preserves_character_state = False
    Resume Clean_Up
End Function

Private Function test_newbie_item_normal_requirements_still_apply() As Boolean
    On Error GoTo Err_Handler

    Dim OriginalStats  As t_UserStats
    Dim OriginalObject As t_ObjData
    Dim OriginalClass  As e_Class
    Dim OriginalRace   As e_Raza
    Dim OriginalGender As e_Genero
    Dim TestObject     As t_ObjData

    With UserList(1)
        OriginalStats = .Stats
        OriginalObject = ObjData(1)
        OriginalClass = .clase
        OriginalRace = .raza
        OriginalGender = .genero
        .Stats.ELV = LimiteNewbie + 1
        .clase = e_Class.Warrior
        .raza = e_Raza.Humano
        .genero = e_Genero.Hombre

        TestObject.Newbie = 1
        TestObject.ClaseProhibida(1) = e_Class.Warrior
        ObjData(1) = TestObject
        If CanUseObject(1, 1) <> 2 Then GoTo TestDone

        TestObject.ClaseProhibida(1) = 0
        TestObject.RazaProhibida(1) = e_Raza.Humano
        ObjData(1) = TestObject
        If CanUseObject(1, 1) <> 5 Then GoTo TestDone

        TestObject.RazaProhibida(1) = 0
        TestObject.MinELV = LimiteNewbie + 2
        ObjData(1) = TestObject
        If CanUseObject(1, 1) <> 6 Then GoTo TestDone

        TestObject.MinELV = 0
        TestObject.SkillIndex = 1
        TestObject.SkillRequerido = 10
        .Stats.UserSkills(1) = 0
        ObjData(1) = TestObject
        If CanUseObject(1, 1) <> 4 Then GoTo TestDone

        test_newbie_item_normal_requirements_still_apply = True
    End With

TestDone:
    With UserList(1)
        .Stats = OriginalStats
        .clase = OriginalClass
        .raza = OriginalRace
        .genero = OriginalGender
    End With
    ObjData(1) = OriginalObject
    Exit Function
Err_Handler:
    test_newbie_item_normal_requirements_still_apply = False
    Resume TestDone
End Function

#End If
