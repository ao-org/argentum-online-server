Attribute VB_Name = "Unit_Partition"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Partition Parsing Test Suite
' Tests SafeCLng helper from modPartition.bas for safe string-to-Long
' conversion with non-numeric, valid, and empty string inputs.
' ==========================================================================
Public Function test_suite_partition() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Call UnitTesting.RunTest("partition_safeclng_nonnumeric", test_safeclng_nonnumeric())
    Call UnitTesting.RunTest("partition_safeclng_valid", test_safeclng_valid())
    Call UnitTesting.RunTest("partition_safeclng_empty", test_safeclng_empty())
    
    Debug.Print "Partition suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_partition = True
End Function

' Verifies SafeCLng returns 0 when given a non-numeric string.
' Requirement 10.1: Non-numeric string -> returns 0
Private Function test_safeclng_nonnumeric() As Boolean
    On Error GoTo Fail
    
    Dim result As Long
    result = SafeCLng("abc")
    
    test_safeclng_nonnumeric = (result = 0)
    Exit Function
Fail:
    test_safeclng_nonnumeric = False
End Function

' Verifies SafeCLng returns the corresponding Long value for a valid numeric string.
' Requirement 10.2: Valid numeric string -> returns corresponding Long
Private Function test_safeclng_valid() As Boolean
    On Error GoTo Fail
    
    Dim result As Long
    result = SafeCLng("42")
    
    test_safeclng_valid = (result = 42)
    Exit Function
Fail:
    test_safeclng_valid = False
End Function

' Verifies SafeCLng returns 0 when given an empty string.
' Requirement 10.3: Empty string -> returns 0
Private Function test_safeclng_empty() As Boolean
    On Error GoTo Fail
    
    Dim result As Long
    result = SafeCLng("")
    
    test_safeclng_empty = (result = 0)
    Exit Function
Fail:
    test_safeclng_empty = False
End Function

#End If
