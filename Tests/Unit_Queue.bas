Attribute VB_Name = "Unit_Queue"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_queue() As Boolean
    Call UnitTesting.RunTest("test_queue_empty_on_init", test_queue_empty_on_init())
    Call UnitTesting.RunTest("test_queue_push_pop", test_queue_push_pop())
    Call UnitTesting.RunTest("test_queue_full", test_queue_full())
    Call UnitTesting.RunTest("test_queue_pop_empty", test_queue_pop_empty())
    Call UnitTesting.RunTest("test_queue_wrap_around", test_queue_wrap_around())
    Call UnitTesting.RunTest("test_queue_count", test_queue_count())
    test_suite_queue = True
End Function

Private Function test_queue_empty_on_init() As Boolean
    On Error GoTo Err_Handler
    test_queue_empty_on_init = True
    Dim q As New CColaArray
    If Not q.IsEmpty Then test_queue_empty_on_init = False: Exit Function
    If q.count <> 0 Then test_queue_empty_on_init = False: Exit Function
    If q.IsFull Then test_queue_empty_on_init = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_empty_on_init = False
End Function

Private Function test_queue_push_pop() As Boolean
    On Error GoTo Err_Handler
    test_queue_push_pop = True
    Dim q As New CColaArray
    ' Push 3 items
    If Not q.Push("first") Then test_queue_push_pop = False: Exit Function
    If Not q.Push("second") Then test_queue_push_pop = False: Exit Function
    If Not q.Push("third") Then test_queue_push_pop = False: Exit Function
    ' Pop in FIFO order
    If q.Pop() <> "first" Then test_queue_push_pop = False: Exit Function
    If q.Pop() <> "second" Then test_queue_push_pop = False: Exit Function
    If q.Pop() <> "third" Then test_queue_push_pop = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_push_pop = False
End Function

Private Function test_queue_full() As Boolean
    On Error GoTo Err_Handler
    test_queue_full = True
    Dim q As New CColaArray
    q.MaxElems = 3
    If Not q.Push("a") Then test_queue_full = False: Exit Function
    If Not q.Push("b") Then test_queue_full = False: Exit Function
    If Not q.Push("c") Then test_queue_full = False: Exit Function
    If Not q.IsFull Then test_queue_full = False: Exit Function
    ' Push when full should return False
    If q.Push("d") Then test_queue_full = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_full = False
End Function

Private Function test_queue_pop_empty() As Boolean
    On Error GoTo Err_Handler
    test_queue_pop_empty = True
    Dim q As New CColaArray
    If q.Pop() <> vbNullString Then test_queue_pop_empty = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_pop_empty = False
End Function

Private Function test_queue_wrap_around() As Boolean
    On Error GoTo Err_Handler
    test_queue_wrap_around = True
    Dim q As New CColaArray
    q.MaxElems = 3
    ' Fill and drain to force internal index wrap
    q.Push "a": q.Push "b": q.Push "c"
    q.Pop: q.Pop: q.Pop
    ' Now push again - internal indices should wrap
    If Not q.Push("d") Then test_queue_wrap_around = False: Exit Function
    If Not q.Push("e") Then test_queue_wrap_around = False: Exit Function
    If q.Pop() <> "d" Then test_queue_wrap_around = False: Exit Function
    If q.Pop() <> "e" Then test_queue_wrap_around = False: Exit Function
    If Not q.IsEmpty Then test_queue_wrap_around = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_wrap_around = False
End Function

Private Function test_queue_count() As Boolean
    On Error GoTo Err_Handler
    test_queue_count = True
    Dim q As New CColaArray
    If q.count <> 0 Then test_queue_count = False: Exit Function
    q.Push "a"
    If q.count <> 1 Then test_queue_count = False: Exit Function
    q.Push "b"
    If q.count <> 2 Then test_queue_count = False: Exit Function
    q.Pop
    If q.count <> 1 Then test_queue_count = False: Exit Function
    q.Pop
    If q.count <> 0 Then test_queue_count = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_count = False
End Function

#End If
