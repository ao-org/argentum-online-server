Attribute VB_Name = "Unit_Queue"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Queue Test Suite
' Tests the CColaArray circular queue: empty state, FIFO ordering,
' full capacity rejection, empty-pop safety, index wrap-around, and count.
' ==========================================================================
Public Function test_suite_queue() As Boolean
    Call UnitTesting.RunTest("test_queue_empty_on_init", test_queue_empty_on_init())
    Call UnitTesting.RunTest("test_queue_push_pop", test_queue_push_pop())
    Call UnitTesting.RunTest("test_queue_full", test_queue_full())
    Call UnitTesting.RunTest("test_queue_pop_empty", test_queue_pop_empty())
    Call UnitTesting.RunTest("test_queue_wrap_around", test_queue_wrap_around())
    Call UnitTesting.RunTest("test_queue_count", test_queue_count())
    test_suite_queue = True
End Function

' Verifies a newly created queue starts empty: IsEmpty=True, count=0, IsFull=False.
Private Function test_queue_empty_on_init() As Boolean
    On Error GoTo Err_Handler
    test_queue_empty_on_init = True
    Dim q As New CColaArray
    ' A brand new queue should report empty
    If Not q.IsEmpty Then test_queue_empty_on_init = False: Exit Function
    ' Count should be 0 since nothing was pushed
    If q.count <> 0 Then test_queue_empty_on_init = False: Exit Function
    ' Should not be full (default capacity is larger than 0)
    If q.IsFull Then test_queue_empty_on_init = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_empty_on_init = False
End Function

' Verifies FIFO ordering: pushing "first", "second", "third" and popping
' returns them in the same order.
Private Function test_queue_push_pop() As Boolean
    On Error GoTo Err_Handler
    test_queue_push_pop = True
    Dim q As New CColaArray
    ' Push 3 items in order
    If Not q.Push("first") Then test_queue_push_pop = False: Exit Function
    If Not q.Push("second") Then test_queue_push_pop = False: Exit Function
    If Not q.Push("third") Then test_queue_push_pop = False: Exit Function
    ' Pop should return items in FIFO order (first in, first out)
    If q.Pop() <> "first" Then test_queue_push_pop = False: Exit Function
    If q.Pop() <> "second" Then test_queue_push_pop = False: Exit Function
    If q.Pop() <> "third" Then test_queue_push_pop = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_push_pop = False
End Function

' Verifies that pushing to a full queue (MaxElems=3) returns False
' and the queue reports IsFull=True.
Private Function test_queue_full() As Boolean
    On Error GoTo Err_Handler
    test_queue_full = True
    Dim q As New CColaArray
    ' Limit the queue to 3 elements
    q.MaxElems = 3
    ' Fill it up to capacity
    If Not q.Push("a") Then test_queue_full = False: Exit Function
    If Not q.Push("b") Then test_queue_full = False: Exit Function
    If Not q.Push("c") Then test_queue_full = False: Exit Function
    ' Queue should now report as full
    If Not q.IsFull Then test_queue_full = False: Exit Function
    ' Pushing a 4th item when full should fail (return False)
    If q.Push("d") Then test_queue_full = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_full = False
End Function

' Verifies that popping from an empty queue returns vbNullString (safe no-op).
Private Function test_queue_pop_empty() As Boolean
    On Error GoTo Err_Handler
    test_queue_pop_empty = True
    Dim q As New CColaArray
    ' Popping from an empty queue should return empty string, not crash
    If q.Pop() <> vbNullString Then test_queue_pop_empty = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_pop_empty = False
End Function

' Verifies the circular buffer handles internal index wrap-around correctly.
' After filling and draining a size-3 queue, new pushes/pops still work
' because the head/tail indices wrap around the internal array.
Private Function test_queue_wrap_around() As Boolean
    On Error GoTo Err_Handler
    test_queue_wrap_around = True
    Dim q As New CColaArray
    q.MaxElems = 3
    ' Fill the queue to capacity, then drain it completely
    ' This advances the internal head/tail pointers to the end of the array
    q.Push "a": q.Push "b": q.Push "c"
    q.Pop: q.Pop: q.Pop
    ' Now push again: the internal indices must wrap around to the beginning
    ' of the array. If wrap logic is broken, these operations would fail.
    If Not q.Push("d") Then test_queue_wrap_around = False: Exit Function
    If Not q.Push("e") Then test_queue_wrap_around = False: Exit Function
    ' Verify FIFO order is maintained after the wrap
    If q.Pop() <> "d" Then test_queue_wrap_around = False: Exit Function
    If q.Pop() <> "e" Then test_queue_wrap_around = False: Exit Function
    ' Queue should be empty again
    If Not q.IsEmpty Then test_queue_wrap_around = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_wrap_around = False
End Function

' Verifies the count property tracks the number of elements accurately
' through a sequence of pushes and pops.
Private Function test_queue_count() As Boolean
    On Error GoTo Err_Handler
    test_queue_count = True
    Dim q As New CColaArray
    ' Empty queue has count 0
    If q.count <> 0 Then test_queue_count = False: Exit Function
    ' Push one item: count goes to 1
    q.Push "a"
    If q.count <> 1 Then test_queue_count = False: Exit Function
    ' Push another: count goes to 2
    q.Push "b"
    If q.count <> 2 Then test_queue_count = False: Exit Function
    ' Pop one: count drops to 1
    q.Pop
    If q.count <> 1 Then test_queue_count = False: Exit Function
    ' Pop the last one: count drops to 0
    q.Pop
    If q.count <> 0 Then test_queue_count = False: Exit Function
    Exit Function
Err_Handler:
    test_queue_count = False
End Function

#End If
