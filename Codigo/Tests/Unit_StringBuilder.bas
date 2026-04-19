Attribute VB_Name = "Unit_StringBuilder"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_stringbuilder() As Boolean
    Call UnitTesting.RunTest("test_sb_append_tostring", test_sb_append_tostring())
    Call UnitTesting.RunTest("test_sb_clear", test_sb_clear())
    Call UnitTesting.RunTest("test_sb_length", test_sb_length())
    Call UnitTesting.RunTest("test_sb_find_exists", test_sb_find_exists())
    Call UnitTesting.RunTest("test_sb_find_not_exists", test_sb_find_not_exists())
    Call UnitTesting.RunTest("test_sb_remove", test_sb_remove())
    Call UnitTesting.RunTest("test_sb_insert", test_sb_insert())
    Call UnitTesting.RunTest("test_prop_sb_length_invariant", test_prop_sb_length_invariant())
    Call UnitTesting.RunTest("test_prop_sb_append_tostring_roundtrip", test_prop_sb_append_tostring_roundtrip())

    test_suite_stringbuilder = True
End Function

' Appending "Hello" + " World" must produce "Hello World" via ToString.
Private Function test_sb_append_tostring() As Boolean
    On Error GoTo Err_Handler
    Dim sb As New cStringBuilder
    Call sb.Append("Hello")
    Call sb.Append(" World")
    test_sb_append_tostring = (sb.ToString = "Hello World")
    Exit Function
Err_Handler:
    test_sb_append_tostring = False
End Function

' Clear must reset ToString to "" and Length to 0.
Private Function test_sb_clear() As Boolean
    On Error GoTo Err_Handler
    Dim sb As New cStringBuilder
    Call sb.Append("some content")
    Call sb.Clear
    test_sb_clear = (sb.ToString = "" And sb.Length = 0)
    Exit Function
Err_Handler:
    test_sb_clear = False
End Function

' Length must return total character count after multiple appends.
Private Function test_sb_length() As Boolean
    On Error GoTo Err_Handler
    Dim sb As New cStringBuilder
    Call sb.Append("abc")
    Call sb.Append("de")
    test_sb_length = (sb.Length = 5)
    Exit Function
Err_Handler:
    test_sb_length = False
End Function

' Property: After N appends, Length must equal the sum of Len() of all appended strings.
Private Function test_prop_sb_length_invariant() As Boolean
    On Error GoTo Err_Handler
    test_prop_sb_length_invariant = True

    Dim sb As cStringBuilder
    Dim expectedLen As Long
    Dim i As Integer
    Dim j As Integer
    Dim testStr As String

    ' Test sequences of 1 to 10 appends with strings of varying lengths
    For i = 1 To 10
        Set sb = New cStringBuilder
        expectedLen = 0

        For j = 1 To i
            ' Build a string of length j (e.g., "a", "bb", "ccc", ...)
            testStr = String$(j, Chr$(96 + j))
            expectedLen = expectedLen + Len(testStr)
            Call sb.Append(testStr)
        Next j

        If sb.Length <> expectedLen Then
            test_prop_sb_length_invariant = False
            Exit Function
        End If
    Next i

    ' Test with empty strings interspersed
    Set sb = New cStringBuilder
    expectedLen = 0
    Call sb.Append("")
    expectedLen = expectedLen + Len("")
    Call sb.Append("hello")
    expectedLen = expectedLen + Len("hello")
    Call sb.Append("")
    expectedLen = expectedLen + Len("")
    Call sb.Append("world")
    expectedLen = expectedLen + Len("world")

    If sb.Length <> expectedLen Then
        test_prop_sb_length_invariant = False
        Exit Function
    End If

    ' Test with longer strings
    Set sb = New cStringBuilder
    expectedLen = 0
    For i = 1 To 20
        testStr = String$(50, Chr$(65 + (i Mod 26)))
        expectedLen = expectedLen + Len(testStr)
        Call sb.Append(testStr)
    Next i

    If sb.Length <> expectedLen Then
        test_prop_sb_length_invariant = False
        Exit Function
    End If

    Exit Function
Err_Handler:
    test_prop_sb_length_invariant = False
End Function

' Property: For any sequence S1..Sn appended, ToString must return S1 & S2 & ... & Sn.
Private Function test_prop_sb_append_tostring_roundtrip() As Boolean
    On Error GoTo Err_Handler
    test_prop_sb_append_tostring_roundtrip = True

    Dim sb As cStringBuilder
    Dim expected As String
    Dim i As Integer
    Dim j As Integer
    Dim testStr As String

    ' Test sequences of 1 to 10 appends
    For i = 1 To 10
        Set sb = New cStringBuilder
        expected = ""

        For j = 1 To i
            testStr = "part" & CStr(j)
            expected = expected & testStr
            Call sb.Append(testStr)
        Next j

        If sb.ToString <> expected Then
            test_prop_sb_append_tostring_roundtrip = False
            Exit Function
        End If
    Next i

    ' Test with empty strings interspersed
    Set sb = New cStringBuilder
    expected = ""
    Dim parts(1 To 5) As String
    parts(1) = ""
    parts(2) = "Hello"
    parts(3) = ""
    parts(4) = " "
    parts(5) = "World"
    For i = 1 To 5
        expected = expected & parts(i)
        Call sb.Append(parts(i))
    Next i

    If sb.ToString <> expected Then
        test_prop_sb_append_tostring_roundtrip = False
        Exit Function
    End If

    ' Test with special characters
    Set sb = New cStringBuilder
    expected = ""
    Dim specials(1 To 4) As String
    specials(1) = "line1" & vbCrLf
    specials(2) = "tab" & vbTab & "here"
    specials(3) = "quotes""inside"
    specials(4) = "end"
    For i = 1 To 4
        expected = expected & specials(i)
        Call sb.Append(specials(i))
    Next i

    If sb.ToString <> expected Then
        test_prop_sb_append_tostring_roundtrip = False
        Exit Function
    End If

    ' Test with many small appends to stress buffer growth
    Set sb = New cStringBuilder
    expected = ""
    For i = 1 To 100
        testStr = CStr(i)
        expected = expected & testStr
        Call sb.Append(testStr)
    Next i

    If sb.ToString <> expected Then
        test_prop_sb_append_tostring_roundtrip = False
        Exit Function
    End If

    Exit Function
Err_Handler:
    test_prop_sb_append_tostring_roundtrip = False
End Function

' Find must return the correct 1-based position for an existing substring.
Private Function test_sb_find_exists() As Boolean
    On Error GoTo Err_Handler
    Dim sb As New cStringBuilder
    Call sb.Append("Hello World")
    test_sb_find_exists = (sb.Find("World") = 7)
    Exit Function
Err_Handler:
    test_sb_find_exists = False
End Function

' Find must return 0 when the substring does not exist.
Private Function test_sb_find_not_exists() As Boolean
    On Error GoTo Err_Handler
    Dim sb As New cStringBuilder
    Call sb.Append("Hello World")
    test_sb_find_not_exists = (sb.Find("xyz") = 0)
    Exit Function
Err_Handler:
    test_sb_find_not_exists = False
End Function

' Remove(5, 6) on "Hello World" must yield "Hello" (0-based index).
Private Function test_sb_remove() As Boolean
    On Error GoTo Err_Handler
    Dim sb As New cStringBuilder
    Call sb.Append("Hello World")
    Call sb.Remove(5, 6)
    test_sb_remove = (sb.ToString = "Hello")
    Exit Function
Err_Handler:
    test_sb_remove = False
End Function

' Insert(5, " ") on "HelloWorld" must yield "Hello World" (0-based index).
Private Function test_sb_insert() As Boolean
    On Error GoTo Err_Handler
    Dim sb As New cStringBuilder
    Call sb.Append("HelloWorld")
    Call sb.Insert(5, " ")
    test_sb_insert = (sb.ToString = "Hello World")
    Exit Function
Err_Handler:
    test_sb_insert = False
End Function

#End If
