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
    test_suite_crafting = True
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
