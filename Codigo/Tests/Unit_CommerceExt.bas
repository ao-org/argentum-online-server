Attribute VB_Name = "Unit_CommerceExt"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Commerce Extended Test Suite
' Tests SalePrice: newbie items return 0, base formula with UserIndex=0,
' non-negative invariant, and property-based verification.
' ==========================================================================

Public Function test_suite_commerce_ext() As Boolean
    ' Example-based tests (Req 11.1, 11.2, 11.3)
    Call UnitTesting.RunTest("test_sale_price_newbie_returns_zero", test_sale_price_newbie_returns_zero())
    Call UnitTesting.RunTest("test_sale_price_base_formula", test_sale_price_base_formula())
    Call UnitTesting.RunTest("test_sale_price_non_negative", test_sale_price_non_negative())
    
    ' Property tests (Property 13, Property 14)
    Call UnitTesting.RunTest("test_prop_sale_price_base_formula", test_prop_sale_price_base_formula())
    Call UnitTesting.RunTest("test_prop_sale_price_non_negative", test_prop_sale_price_non_negative())
    
    test_suite_commerce_ext = True
End Function

' --------------------------------------------------------------------------
' Example-based tests
' --------------------------------------------------------------------------

' Req 11.1: SalePrice returns 0 for Newbie items.
' WHEN SalePrice is called for a Newbie item, THE Comercio module SHALL return 0.
Private Function test_sale_price_newbie_returns_zero() As Boolean
    On Error GoTo Err_Handler
    
    Dim ObjIndex As Integer
    ObjIndex = 1
    
    ' Save original values
    Dim origValor As Long
    Dim origNewbie As Integer
    origValor = ObjData(ObjIndex).Valor
    origNewbie = ObjData(ObjIndex).Newbie
    
    ' Setup: mark item as Newbie with a positive Valor
    ObjData(ObjIndex).Valor = 500
    ObjData(ObjIndex).Newbie = 1
    
    Dim result As Single
    result = SalePrice(ObjIndex, 0)
    
    ' Restore
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    
    test_sale_price_newbie_returns_zero = (result = 0)
    Exit Function
Err_Handler:
    On Error Resume Next
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    On Error GoTo 0
    test_sale_price_newbie_returns_zero = False
End Function

' Req 11.2: SalePrice base formula with UserIndex=0.
' WHEN SalePrice is called with a valid ObjIndex and UserIndex of 0,
' THE Comercio module SHALL return Valor / REDUCTOR_PRECIOVENTA.
Private Function test_sale_price_base_formula() As Boolean
    On Error GoTo Err_Handler
    
    Dim ObjIndex As Integer
    ObjIndex = 1
    
    ' Save original values
    Dim origValor As Long
    Dim origNewbie As Integer
    origValor = ObjData(ObjIndex).Valor
    origNewbie = ObjData(ObjIndex).Newbie
    
    ' Setup: known Valor, not Newbie
    ObjData(ObjIndex).Valor = 600
    ObjData(ObjIndex).Newbie = 0
    
    Dim expected As Single
    expected = CSng(600 / REDUCTOR_PRECIOVENTA)  ' 600 / 3 = 200
    
    Dim result As Single
    result = SalePrice(ObjIndex, 0)
    
    ' Restore
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    
    test_sale_price_base_formula = (result = expected)
    Exit Function
Err_Handler:
    On Error Resume Next
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    On Error GoTo 0
    test_sale_price_base_formula = False
End Function

' Req 11.3: Sale price is always non-negative for valid items.
' THE Comercio module SHALL ensure sale price is always non-negative.
Private Function test_sale_price_non_negative() As Boolean
    On Error GoTo Err_Handler
    
    Dim ObjIndex As Integer
    ObjIndex = 1
    
    ' Save original values
    Dim origValor As Long
    Dim origNewbie As Integer
    origValor = ObjData(ObjIndex).Valor
    origNewbie = ObjData(ObjIndex).Newbie
    
    ' Test with zero Valor (non-newbie)
    ObjData(ObjIndex).Valor = 0
    ObjData(ObjIndex).Newbie = 0
    
    Dim result As Single
    result = SalePrice(ObjIndex, 0)
    
    If result < 0 Then
        ObjData(ObjIndex).Valor = origValor
        ObjData(ObjIndex).Newbie = origNewbie
        test_sale_price_non_negative = False
        Exit Function
    End If
    
    ' Test with a positive Valor
    ObjData(ObjIndex).Valor = 100
    result = SalePrice(ObjIndex, 0)
    
    ' Restore
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    
    test_sale_price_non_negative = (result >= 0)
    Exit Function
Err_Handler:
    On Error Resume Next
    ObjData(ObjIndex).Valor = origValor
    ObjData(ObjIndex).Newbie = origNewbie
    On Error GoTo 0
    test_sale_price_non_negative = False
End Function

' --------------------------------------------------------------------------
' Property tests
' --------------------------------------------------------------------------

' Feature: unit-test-coverage, Property 13: SalePrice base formula
' **Validates: Requirements 11.2**
'
' Loop over 100+ valid non-newbie items with positive Valor, verify
' SalePrice(ObjIndex, 0) = Valor / REDUCTOR_PRECIOVENTA.
Private Function test_prop_sale_price_base_formula() As Boolean
    On Error GoTo Err_Handler
    
    Dim iterations As Long
    iterations = 0
    
    Dim maxIdx As Integer
    maxIdx = UBound(ObjData)
    
    Dim i As Integer
    Dim origValor As Long
    Dim origNewbie As Integer
    Dim expected As Single
    Dim result As Single
    
    ' Iterate over all valid ObjIndex values, testing those that are
    ' non-newbie with positive Valor (or forcing them to be so).
    ' We use a step to cover the range and ensure 100+ iterations.
    ' For each item, temporarily set Valor to a known positive value
    ' and Newbie=0 to guarantee the base formula applies.
    
    Dim testValor As Long
    For i = 1 To maxIdx
        ' Save original values
        origValor = ObjData(i).Valor
        origNewbie = ObjData(i).Newbie
        
        ' Set up: non-newbie with a deterministic positive Valor
        testValor = CLng(i) * 10 + 100  ' Ensures positive, varies per item
        ObjData(i).Valor = testValor
        ObjData(i).Newbie = 0
        
        expected = CSng(testValor / REDUCTOR_PRECIOVENTA)
        result = SalePrice(i, 0)
        
        ' Restore
        ObjData(i).Valor = origValor
        ObjData(i).Newbie = origNewbie
        
        If result <> expected Then
            test_prop_sale_price_base_formula = False
            Exit Function
        End If
        
        iterations = iterations + 1
        
        ' Stop after enough iterations to keep test fast
        If iterations >= 150 Then Exit For
    Next i
    
    ' Verify we ran at least 100 iterations
    test_prop_sale_price_base_formula = (iterations >= 100)
    Exit Function
Err_Handler:
    ' Attempt restore on error
    On Error Resume Next
    ObjData(i).Valor = origValor
    ObjData(i).Newbie = origNewbie
    On Error GoTo 0
    test_prop_sale_price_base_formula = False
End Function

' Feature: unit-test-coverage, Property 14: SalePrice non-negative invariant
' **Validates: Requirements 11.3**
'
' Loop over 100+ valid ObjIndex values, verify SalePrice >= 0.
Private Function test_prop_sale_price_non_negative() As Boolean
    On Error GoTo Err_Handler
    
    Dim iterations As Long
    iterations = 0
    
    Dim maxIdx As Integer
    maxIdx = UBound(ObjData)
    
    Dim i As Integer
    Dim result As Single
    
    ' Iterate over all valid ObjIndex values and verify SalePrice >= 0
    ' for each item in its current state (no modifications needed).
    For i = 1 To maxIdx
        result = SalePrice(i, 0)
        
        If result < 0 Then
            test_prop_sale_price_non_negative = False
            Exit Function
        End If
        
        iterations = iterations + 1
        
        ' Stop after enough iterations to keep test fast
        If iterations >= 150 Then Exit For
    Next i
    
    ' Verify we ran at least 100 iterations
    test_prop_sale_price_non_negative = (iterations >= 100)
    Exit Function
Err_Handler:
    test_prop_sale_price_non_negative = False
End Function

#End If
