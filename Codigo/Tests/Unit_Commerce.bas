Attribute VB_Name = "Unit_Commerce"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_commerce() As Boolean
    Call UnitTesting.RunTest("test_sale_price_base", test_sale_price_base())
    Call UnitTesting.RunTest("test_sale_price_trabajador_discount", test_sale_price_trabajador_discount())
    Call UnitTesting.RunTest("test_sale_price_non_trabajador", test_sale_price_non_trabajador())
    Call UnitTesting.RunTest("test_sale_price_trabajador_clamp", test_sale_price_trabajador_clamp())
    Call UnitTesting.RunTest("test_sale_price_invalid_obj", test_sale_price_invalid_obj())
    test_suite_commerce = True
End Function

' SalePrice with no user (UserIndex=0) must return Valor / REDUCTOR_PRECIOVENTA.
Private Function test_sale_price_base() As Boolean
    On Error GoTo Err_Handler
    test_sale_price_base = True

    Dim ObjIndex As Integer
    ObjIndex = 1

    ' Save original values
    Dim origValor As Long
    Dim origNewbie As Integer
    origValor = ObjData(ObjIndex).Valor
    origNewbie = ObjData(ObjIndex).Newbie

    ' Setup: known Valor, not Newbie
    ObjData(ObjIndex).Valor = 300
    ObjData(ObjIndex).Newbie = 0

    Dim expected As Single
    expected = CSng(300 / REDUCTOR_PRECIOVENTA)  ' 300 / 3 = 100

    Dim result As Single
    result = SalePrice(ObjIndex, 0)

    If result <> expected Then test_sale_price_base = False

    ' Restore
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    Exit Function
Err_Handler:
    test_sale_price_base = False
End Function

' Trabajador with destroy_npc_bought_items enabled must get a higher sale price
' than the base case (lower denominator = higher price).
Private Function test_sale_price_trabajador_discount() As Boolean
    On Error GoTo Err_Handler
    test_sale_price_trabajador_discount = True

    Dim ObjIndex As Integer
    ObjIndex = 1

    ' Save original values
    Dim origValor As Long
    Dim origNewbie As Integer
    origValor = ObjData(ObjIndex).Valor
    origNewbie = ObjData(ObjIndex).Newbie

    ' Setup object
    ObjData(ObjIndex).Valor = 300
    ObjData(ObjIndex).Newbie = 0

    ' Setup user as Trabajador with ELV=20
    Dim UserIndex As Integer
    UserIndex = 1
    Dim origClase As e_Class
    Dim origELV As Byte
    origClase = UserList(UserIndex).clase
    origELV = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).clase = e_Class.Trabajador
    UserList(UserIndex).Stats.ELV = 20

    ' Enable feature toggle
    Call SetFeatureToggle("destroy_npc_bought_items", True)

    Dim basePrice As Single
    basePrice = SalePrice(ObjIndex, 0)

    Dim discountedPrice As Single
    discountedPrice = SalePrice(ObjIndex, UserIndex)

    ' Trabajador discount should yield a higher price (lower denom)
    If Not (discountedPrice > basePrice) Then test_sale_price_trabajador_discount = False

    ' Restore
    UserList(UserIndex).clase = origClase
    UserList(UserIndex).Stats.ELV = origELV
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    Exit Function
Err_Handler:
    test_sale_price_trabajador_discount = False
End Function

' Non-Trabajador must get the base price even with destroy_npc_bought_items enabled.
Private Function test_sale_price_non_trabajador() As Boolean
    On Error GoTo Err_Handler
    test_sale_price_non_trabajador = True

    Dim ObjIndex As Integer
    ObjIndex = 1

    ' Save original values
    Dim origValor As Long
    Dim origNewbie As Integer
    origValor = ObjData(ObjIndex).Valor
    origNewbie = ObjData(ObjIndex).Newbie

    ' Setup object
    ObjData(ObjIndex).Valor = 300
    ObjData(ObjIndex).Newbie = 0

    ' Setup user as Warrior (non-Trabajador)
    Dim UserIndex As Integer
    UserIndex = 1
    Dim origClase As e_Class
    Dim origELV As Byte
    origClase = UserList(UserIndex).clase
    origELV = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).clase = e_Class.Warrior
    UserList(UserIndex).Stats.ELV = 20

    ' Enable feature toggle
    Call SetFeatureToggle("destroy_npc_bought_items", True)

    Dim basePrice As Single
    basePrice = SalePrice(ObjIndex, 0)

    Dim userPrice As Single
    userPrice = SalePrice(ObjIndex, UserIndex)

    ' Non-Trabajador should get the same price as base
    If userPrice <> basePrice Then test_sale_price_non_trabajador = False

    ' Restore
    UserList(UserIndex).clase = origClase
    UserList(UserIndex).Stats.ELV = origELV
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    Exit Function
Err_Handler:
    test_sale_price_non_trabajador = False
End Function

' Trabajador at level 40: denom = 3 - (40 * 0.025) = 2.0, so SalePrice = Valor / 2.
Private Function test_sale_price_trabajador_clamp() As Boolean
    On Error GoTo Err_Handler
    test_sale_price_trabajador_clamp = True

    Dim ObjIndex As Integer
    ObjIndex = 1

    ' Save original values
    Dim origValor As Long
    Dim origNewbie As Integer
    origValor = ObjData(ObjIndex).Valor
    origNewbie = ObjData(ObjIndex).Newbie

    ' Setup object
    ObjData(ObjIndex).Valor = 300
    ObjData(ObjIndex).Newbie = 0

    ' Setup user as Trabajador at level 40
    Dim UserIndex As Integer
    UserIndex = 1
    Dim origClase As e_Class
    Dim origELV As Byte
    origClase = UserList(UserIndex).clase
    origELV = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).clase = e_Class.Trabajador
    UserList(UserIndex).Stats.ELV = 40

    ' Enable feature toggle
    Call SetFeatureToggle("destroy_npc_bought_items", True)

    Dim expected As Single
    expected = CSng(300 / 2)  ' 150

    Dim result As Single
    result = SalePrice(ObjIndex, UserIndex)

    If result <> expected Then test_sale_price_trabajador_clamp = False

    ' Restore
    UserList(UserIndex).clase = origClase
    UserList(UserIndex).Stats.ELV = origELV
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    Exit Function
Err_Handler:
    test_sale_price_trabajador_clamp = False
End Function

' SalePrice must return 0 for invalid ObjIndex values (0 and beyond UBound).
Private Function test_sale_price_invalid_obj() As Boolean
    On Error GoTo Err_Handler
    test_sale_price_invalid_obj = True

    ' ObjIndex = 0 should return 0
    If SalePrice(0, 0) <> 0 Then test_sale_price_invalid_obj = False: Exit Function

    ' ObjIndex beyond UBound should return 0
    Dim maxIdx As Integer
    maxIdx = UBound(ObjData)
    If SalePrice(maxIdx + 1, 0) <> 0 Then test_sale_price_invalid_obj = False: Exit Function

    Exit Function
Err_Handler:
    test_sale_price_invalid_obj = False
End Function

#End If
