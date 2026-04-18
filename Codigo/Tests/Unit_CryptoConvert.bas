Attribute VB_Name = "Unit_CryptoConvert"
Option Explicit

' ==========================================================================
' CryptoConvert Test Suite (Server)
' Tests AO20CryptoSysWrapper.bas: HiByte, LoByte, MakeInt, Str2ByteArr,
' ByteArr2String, CopyBytes, ByteArrayToHex, IsBase64.
'
' Requirements: 10.1, 10.2, 10.3, 10.4, 10.5
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Function test_suite_crypto_convert() As Boolean
    ' Example-based tests (Req 10.1)
    Call UnitTesting.RunTest("scc_hibyte_256", test_hibyte_256())
    Call UnitTesting.RunTest("scc_hibyte_0", test_hibyte_0())
    Call UnitTesting.RunTest("scc_lobyte_258", test_lobyte_258())
    Call UnitTesting.RunTest("scc_lobyte_255", test_lobyte_255())
    Call UnitTesting.RunTest("scc_str2bytearr", test_str2bytearr())
    Call UnitTesting.RunTest("scc_bytearr2string", test_bytearr2string())
    Call UnitTesting.RunTest("scc_copybytes", test_copybytes())
    Call UnitTesting.RunTest("scc_bytearraytohex", test_bytearraytohex())

    ' Initialize Base64 lookup before IsBase64 tests (Req 10.5)
    Call initBase64Chars

    ' IsBase64 example-based tests (Req 10.2, 10.3, 10.4)
    Call UnitTesting.RunTest("scc_isbase64_valid", test_isbase64_valid())
    Call UnitTesting.RunTest("scc_isbase64_invalid", test_isbase64_invalid())
    Call UnitTesting.RunTest("scc_isbase64_empty", test_isbase64_empty())

    ' Property-based tests (Req 10.1, 10.2, 10.3)
    Call UnitTesting.RunTest("scc_pbt_makeint_roundtrip", test_pbt_makeint_roundtrip())
    Call UnitTesting.RunTest("scc_pbt_str2bytearr_roundtrip", test_pbt_str2bytearr_roundtrip())
    Call UnitTesting.RunTest("scc_pbt_copybytes_correctness", test_pbt_copybytes_correctness())
    Call UnitTesting.RunTest("scc_pbt_byte_to_string_repr", test_pbt_byte_to_string_repr())
    Call UnitTesting.RunTest("scc_pbt_isbase64_membership", test_pbt_isbase64_membership())

    test_suite_crypto_convert = True
End Function

' --------------------------------------------------------------------------
' Example-based tests
' --------------------------------------------------------------------------

' Requirement 10.1: HiByte(256) = 1
Private Function test_hibyte_256() As Boolean
    On Error GoTo Fail
    test_hibyte_256 = (hiByte(256) = 1)
    Exit Function
Fail:
    test_hibyte_256 = False
End Function

' Requirement 10.1: HiByte(0) = 0
Private Function test_hibyte_0() As Boolean
    On Error GoTo Fail
    test_hibyte_0 = (hiByte(0) = 0)
    Exit Function
Fail:
    test_hibyte_0 = False
End Function

' Requirement 10.1: LoByte(258) = 2
Private Function test_lobyte_258() As Boolean
    On Error GoTo Fail
    test_lobyte_258 = (LoByte(258) = 2)
    Exit Function
Fail:
    test_lobyte_258 = False
End Function

' Requirement 10.1: LoByte(255) = 255
Private Function test_lobyte_255() As Boolean
    On Error GoTo Fail
    test_lobyte_255 = (LoByte(255) = 255)
    Exit Function
Fail:
    test_lobyte_255 = False
End Function

' Requirement 10.1: Str2ByteArr produces correct length and ASCII byte values
Private Function test_str2bytearr() As Boolean
    On Error GoTo Fail
    Dim arr() As Byte
    Call Str2ByteArr("ABC", arr)
    ' Array should be 0-based with 3 elements
    If UBound(arr) - LBound(arr) + 1 <> 3 Then
        test_str2bytearr = False
        Exit Function
    End If
    ' A=65, B=66, C=67
    test_str2bytearr = (arr(0) = 65) And (arr(1) = 66) And (arr(2) = 67)
    Exit Function
Fail:
    test_str2bytearr = False
End Function

' Requirement 10.1: ByteArr2String returns correct ASCII string
Private Function test_bytearr2string() As Boolean
    On Error GoTo Fail
    Dim arr(0 To 2) As Byte
    arr(0) = 72   ' H
    arr(1) = 105  ' i
    arr(2) = 33   ' !
    test_bytearr2string = (ByteArr2String(arr) = "Hi!")
    Exit Function
Fail:
    test_bytearr2string = False
End Function

' Requirement 10.1: CopyBytes copies at offset, other bytes unchanged
Private Function test_copybytes() As Boolean
    On Error GoTo Fail
    Dim src(0 To 1) As Byte
    src(0) = 10
    src(1) = 20
    
    Dim dst(0 To 4) As Byte
    dst(0) = 99
    dst(1) = 99
    dst(2) = 99
    dst(3) = 99
    dst(4) = 99
    
    Call CopyBytes(src, dst, 2, 2)
    
    ' Bytes at offset 2 and 3 should be copied from src
    ' Bytes at 0, 1, 4 should remain 99
    test_copybytes = (dst(0) = 99) And (dst(1) = 99) And _
                     (dst(2) = 10) And (dst(3) = 20) And _
                     (dst(4) = 99)
    Exit Function
Fail:
    test_copybytes = False
End Function

' Requirement 10.1: ByteArrayToHex produces hex representation
' VB6 Hex$() does not zero-pad, so &H0A becomes "A" not "0A"
Private Function test_bytearraytohex() As Boolean
    On Error GoTo Fail
    Dim arr(0 To 1) As Byte
    arr(0) = &HFF
    arr(1) = &HA
    test_bytearraytohex = (ByteArrayToHex(arr) = "FF A")
    Exit Function
Fail:
    test_bytearraytohex = False
End Function

' Requirement 10.2: IsBase64 with valid Base64 string returns True
Private Function test_isbase64_valid() As Boolean
    On Error GoTo Fail
    test_isbase64_valid = (IsBase64("SGVsbG8=") = True)
    Exit Function
Fail:
    test_isbase64_valid = False
End Function

' Requirement 10.3: IsBase64 with invalid characters returns False
Private Function test_isbase64_invalid() As Boolean
    On Error GoTo Fail
    test_isbase64_invalid = (IsBase64("!@#") = False)
    Exit Function
Fail:
    test_isbase64_invalid = False
End Function

' Requirement 10.4: IsBase64 with empty string - verify behavior
Private Function test_isbase64_empty() As Boolean
    On Error GoTo Fail
    ' Empty string has no invalid chars, so IsBase64 returns True
    test_isbase64_empty = (IsBase64("") = True)
    Exit Function
Fail:
    test_isbase64_empty = False
End Function

' --------------------------------------------------------------------------
' Property-based tests
' --------------------------------------------------------------------------

' Feature: unit-test-coverage-tier4, Property 4: HiByte/LoByte/MakeInt round-trip (server copy)
' Validates: Requirements 10.1
Private Function test_pbt_makeint_roundtrip() As Boolean
    On Error GoTo Fail
    
    Dim n As Long
    
    For n = 0 To 32767
        If MakeInt(LoByte(CInt(n)), hiByte(CInt(n))) <> CInt(n) Then
            test_pbt_makeint_roundtrip = False
            Exit Function
        End If
    Next n
    
    test_pbt_makeint_roundtrip = True
    Exit Function
Fail:
    test_pbt_makeint_roundtrip = False
End Function

' Feature: unit-test-coverage-tier4, Property 5: Str2ByteArr/ByteArr2String round-trip (server copy)
' Validates: Requirements 10.1
Private Function test_pbt_str2bytearr_roundtrip() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim arr() As Byte
    Dim result As String
    Dim strLen As Long
    Dim charCode As Long
    
    For i = 1 To 110
        ' Generate an ASCII string of length i with chars in range 1-127
        s = vbNullString
        strLen = ((i - 1) Mod 20) + 1
        For j = 1 To strLen
            charCode = ((i * 7 + j * 13) Mod 127) + 1  ' range 1-127
            s = s & Chr$(charCode)
        Next j
        
        Call Str2ByteArr(s, arr)
        result = ByteArr2String(arr)
        
        If result <> s Then
            test_pbt_str2bytearr_roundtrip = False
            Exit Function
        End If
    Next i
    
    test_pbt_str2bytearr_roundtrip = True
    Exit Function
Fail:
    test_pbt_str2bytearr_roundtrip = False
End Function

' Feature: unit-test-coverage-tier4, Property 6: CopyBytes correctness (server copy)
' Validates: Requirements 10.1
Private Function test_pbt_copybytes_correctness() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim j As Long
    Dim srcSize As Long
    Dim dstSize As Long
    Dim offset As Long
    Dim src() As Byte
    Dim dst() As Byte
    Dim origDst() As Byte
    
    For i = 1 To 110
        ' Generate varying source sizes (1..10) and offsets
        srcSize = ((i - 1) Mod 10) + 1
        offset = (i - 1) Mod 5
        dstSize = srcSize + offset + 2  ' ensure dst is large enough with room to spare
        
        ReDim src(0 To srcSize - 1)
        ReDim dst(0 To dstSize - 1)
        ReDim origDst(0 To dstSize - 1)
        
        ' Fill source with deterministic values
        For j = 0 To srcSize - 1
            src(j) = CByte((i * 3 + j * 7) Mod 256)
        Next j
        
        ' Fill dest with a sentinel value
        For j = 0 To dstSize - 1
            dst(j) = CByte(200)
            origDst(j) = CByte(200)
        Next j
        
        Call CopyBytes(src, dst, srcSize, offset)
        
        ' Verify copied bytes match source
        For j = 0 To srcSize - 1
            If dst(j + offset) <> src(j) Then
                test_pbt_copybytes_correctness = False
                Exit Function
            End If
        Next j
        
        ' Verify non-copied bytes are unchanged
        For j = 0 To dstSize - 1
            If j < offset Or j >= offset + srcSize Then
                If dst(j) <> origDst(j) Then
                    test_pbt_copybytes_correctness = False
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    test_pbt_copybytes_correctness = True
    Exit Function
Fail:
    test_pbt_copybytes_correctness = False
End Function

' Feature: unit-test-coverage-tier4, Property 7: Byte array to string representation (server copy)
' Validates: Requirements 10.1
' Note: Server does not have ByteArrayToDecimalString, only ByteArrayToHex
Private Function test_pbt_byte_to_string_repr() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim j As Long
    Dim arrSize As Long
    Dim arr() As Byte
    Dim hexResult As String
    Dim hexTokens() As String
    Dim expectedHex As String
    
    For i = 1 To 110
        ' Generate arrays of size 1..10
        arrSize = ((i - 1) Mod 10) + 1
        ReDim arr(0 To arrSize - 1)
        
        For j = 0 To arrSize - 1
            arr(j) = CByte((i * 11 + j * 17) Mod 256)
        Next j
        
        hexResult = ByteArrayToHex(arr)
        
        ' Split results into tokens
        hexTokens = Split(hexResult, " ")
        
        ' Verify token count matches array size
        If UBound(hexTokens) - LBound(hexTokens) + 1 <> arrSize Then
            test_pbt_byte_to_string_repr = False
            Exit Function
        End If
        
        ' Verify each hex token is the hex representation of the byte
        For j = 0 To arrSize - 1
            expectedHex = Hex$(arr(j))
            If hexTokens(j) <> expectedHex Then
                test_pbt_byte_to_string_repr = False
                Exit Function
            End If
        Next j
    Next i
    
    test_pbt_byte_to_string_repr = True
    Exit Function
Fail:
    test_pbt_byte_to_string_repr = False
End Function

' Feature: unit-test-coverage-tier4, Property 11: IsBase64 character membership
' Validates: Requirements 10.2, 10.3
' For any string composed exclusively of Base64 alphabet chars (A-Z, a-z, 0-9, +, /, =),
' IsBase64 should return True. For any string containing at least one non-Base64 char,
' IsBase64 should return False.
Private Function test_pbt_isbase64_membership() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim strLen As Long
    Dim charIdx As Long
    
    ' Base64 alphabet: A-Z, a-z, 0-9, +, /, =
    Dim b64Alphabet As String
    b64Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    
    ' Non-Base64 characters for negative tests
    Dim nonB64Chars As String
    nonB64Chars = "!@#$%^&*(){}[]|;:<>?,. " & Chr$(1) & Chr$(9)
    
    ' --- Part 1: Valid Base64 strings should return True ---
    For i = 1 To 110
        s = vbNullString
        strLen = ((i - 1) Mod 15) + 1  ' lengths 1..15
        For j = 1 To strLen
            ' Pick a character from the Base64 alphabet deterministically
            charIdx = ((i * 7 + j * 13) Mod Len(b64Alphabet)) + 1
            s = s & Mid$(b64Alphabet, charIdx, 1)
        Next j
        
        If IsBase64(s) <> True Then
            test_pbt_isbase64_membership = False
            Exit Function
        End If
    Next i
    
    ' --- Part 2: Strings with non-Base64 chars should return False ---
    For i = 1 To 110
        s = vbNullString
        strLen = ((i - 1) Mod 10) + 2  ' lengths 2..11
        
        ' Start with valid Base64 chars
        For j = 1 To strLen - 1
            charIdx = ((i * 11 + j * 3) Mod Len(b64Alphabet)) + 1
            s = s & Mid$(b64Alphabet, charIdx, 1)
        Next j
        
        ' Insert one non-Base64 character at a deterministic position
        Dim badCharIdx As Long
        badCharIdx = ((i * 5) Mod Len(nonB64Chars)) + 1
        Dim insertPos As Long
        insertPos = ((i - 1) Mod Len(s)) + 1
        s = Left$(s, insertPos - 1) & Mid$(nonB64Chars, badCharIdx, 1) & Mid$(s, insertPos)
        
        If IsBase64(s) <> False Then
            test_pbt_isbase64_membership = False
            Exit Function
        End If
    Next i
    
    test_pbt_isbase64_membership = True
    Exit Function
Fail:
    test_pbt_isbase64_membership = False
End Function

#End If
