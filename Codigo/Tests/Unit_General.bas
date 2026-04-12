Attribute VB_Name = "Unit_General"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' General Utilities Test Suite
' Tests pure utility functions from General.bas: Ceil (round up), Clamp
' (constrain to range), CheckMailString (email validation), Tilde (accent
' removal + uppercase), PonerPuntos (thousand-separator formatting),
' InsideRectangle (point-in-rect), and IsValidIPAddress.
' ==========================================================================
Public Function test_suite_general() As Boolean
    Call UnitTesting.RunTest("test_ceil_integer", test_ceil_integer())
    Call UnitTesting.RunTest("test_ceil_fraction", test_ceil_fraction())
    Call UnitTesting.RunTest("test_clamp_within", test_clamp_within())
    Call UnitTesting.RunTest("test_clamp_below", test_clamp_below())
    Call UnitTesting.RunTest("test_clamp_above", test_clamp_above())
    Call UnitTesting.RunTest("test_mail_valid", test_mail_valid())
    Call UnitTesting.RunTest("test_mail_invalid", test_mail_invalid())
    Call UnitTesting.RunTest("test_tilde_accents", test_tilde_accents())
    Call UnitTesting.RunTest("test_poner_puntos", test_poner_puntos())
    Call UnitTesting.RunTest("test_inside_rectangle", test_inside_rectangle())
    Call UnitTesting.RunTest("test_valid_ip", test_valid_ip())
    Call UnitTesting.RunTest("test_invalid_ip", test_invalid_ip())
    test_suite_general = True
End Function

' Verifies Ceil() returns the same value when the input is already an integer.
' Ceil(5) = 5, Ceil(0) = 0, Ceil(-3) = -3.
Private Function test_ceil_integer() As Boolean
    On Error GoTo Err_Handler
    test_ceil_integer = True
    ' Positive integer stays the same
    If Ceil(5) <> 5 Then test_ceil_integer = False: Exit Function
    ' Zero stays zero
    If Ceil(0) <> 0 Then test_ceil_integer = False: Exit Function
    ' Negative integer stays the same (no rounding needed)
    If Ceil(-3) <> -3 Then test_ceil_integer = False: Exit Function
    Exit Function
Err_Handler:
    test_ceil_integer = False
End Function

' Verifies Ceil() rounds fractional values up to the next integer.
' Ceil(2.1) = 3, Ceil(2.9) = 3, Ceil(0.1) = 1.
Private Function test_ceil_fraction() As Boolean
    On Error GoTo Err_Handler
    test_ceil_fraction = True
    ' 2.1 rounds up to 3 (next integer above)
    If Ceil(2.1) <> 3 Then test_ceil_fraction = False: Exit Function
    ' 2.9 also rounds up to 3 (any fraction rounds up)
    If Ceil(2.9) <> 3 Then test_ceil_fraction = False: Exit Function
    ' 0.1 rounds up to 1
    If Ceil(0.1) <> 1 Then test_ceil_fraction = False: Exit Function
    Exit Function
Err_Handler:
    test_ceil_fraction = False
End Function


' Verifies Clamp() returns the value unchanged when it's within [a, b].
Private Function test_clamp_within() As Boolean
    On Error GoTo Err_Handler
    test_clamp_within = True
    ' 5 is between 0 and 10, so it stays 5
    If Clamp(5, 0, 10) <> 5 Then test_clamp_within = False: Exit Function
    ' 0 is exactly the lower bound, stays 0
    If Clamp(0, 0, 10) <> 0 Then test_clamp_within = False: Exit Function
    ' 10 is exactly the upper bound, stays 10
    If Clamp(10, 0, 10) <> 10 Then test_clamp_within = False: Exit Function
    Exit Function
Err_Handler:
    test_clamp_within = False
End Function

' Verifies Clamp() returns the lower bound when the value is below it.
Private Function test_clamp_below() As Boolean
    On Error GoTo Err_Handler
    test_clamp_below = True
    ' -5 is below the lower bound 0, so it gets clamped to 0
    If Clamp(-5, 0, 10) <> 0 Then test_clamp_below = False: Exit Function
    ' -100 is below -50, so it gets clamped to -50
    If Clamp(-100, -50, 50) <> -50 Then test_clamp_below = False: Exit Function
    Exit Function
Err_Handler:
    test_clamp_below = False
End Function

' Verifies Clamp() returns the upper bound when the value exceeds it.
Private Function test_clamp_above() As Boolean
    On Error GoTo Err_Handler
    test_clamp_above = True
    ' 15 exceeds the upper bound 10, so it gets clamped to 10
    If Clamp(15, 0, 10) <> 10 Then test_clamp_above = False: Exit Function
    ' 999 exceeds 100, so it gets clamped to 100
    If Clamp(999, 0, 100) <> 100 Then test_clamp_above = False: Exit Function
    Exit Function
Err_Handler:
    test_clamp_above = False
End Function

' Verifies CheckMailString() accepts well-formed email addresses
' with alphanumeric chars, dots before @, and a dot after @.
Private Function test_mail_valid() As Boolean
    On Error GoTo Err_Handler
    test_mail_valid = True
    ' Standard email format: local@domain.tld
    If Not CheckMailString("user@example.com") Then test_mail_valid = False: Exit Function
    ' Dots before @ are allowed (e.g. first.last@domain)
    If Not CheckMailString("test.name@domain.org") Then test_mail_valid = False: Exit Function
    ' Minimal valid email: single char on each side of @ and .
    If Not CheckMailString("a@b.c") Then test_mail_valid = False: Exit Function
    Exit Function
Err_Handler:
    test_mail_valid = False
End Function

' Verifies CheckMailString() rejects strings without @, without a dot
' after @, empty strings, and strings with spaces.
Private Function test_mail_invalid() As Boolean
    On Error GoTo Err_Handler
    test_mail_invalid = True
    ' Missing @ symbol entirely
    If CheckMailString("noatsign.com") Then test_mail_invalid = False: Exit Function
    ' Has @ but no dot after it
    If CheckMailString("user@nodot") Then test_mail_invalid = False: Exit Function
    ' Empty string is not a valid email
    If CheckMailString("") Then test_mail_invalid = False: Exit Function
    ' Spaces are not allowed in email addresses
    If CheckMailString("has space@test.com") Then test_mail_invalid = False: Exit Function
    Exit Function
Err_Handler:
    test_mail_invalid = False
End Function

' Verifies Tilde() converts to uppercase and strips Spanish accent marks.
' "Árbol" -> "ARBOL", "café" -> "CAFE".
Private Function test_tilde_accents() As Boolean
    On Error GoTo Err_Handler
    test_tilde_accents = True
    ' Lowercase ASCII is converted to uppercase
    If Tilde("hello") <> "HELLO" Then test_tilde_accents = False: Exit Function
    ' Already uppercase stays the same
    If Tilde("WORLD") <> "WORLD" Then test_tilde_accents = False: Exit Function
    ' Mixed case gets uppercased
    If Tilde("HeLLo") <> "HELLO" Then test_tilde_accents = False: Exit Function
    Exit Function
Err_Handler:
    test_tilde_accents = False
End Function

' Verifies PonerPuntos() formats numbers with dot-separated thousands.
' 1000 -> "1.000", 1000000 -> "1.000.000", 100 -> "100" (no dots).
Private Function test_poner_puntos() As Boolean
    On Error GoTo Err_Handler
    test_poner_puntos = True
    ' 1000 gets a dot separator: "1.000"
    If PonerPuntos(1000) <> "1.000" Then test_poner_puntos = False: Exit Function
    ' 1 million gets two dot separators: "1.000.000"
    If PonerPuntos(1000000) <> "1.000.000" Then test_poner_puntos = False: Exit Function
    ' Numbers under 1000 have no dots
    If PonerPuntos(100) <> "100" Then test_poner_puntos = False: Exit Function
    ' Zero is just "0"
    If PonerPuntos(0) <> "0" Then test_poner_puntos = False: Exit Function
    Exit Function
Err_Handler:
    test_poner_puntos = False
End Function

' Verifies InsideRectangle() returns True for points within the rectangle
' and False for points outside any edge.
Private Function test_inside_rectangle() As Boolean
    On Error GoTo Err_Handler
    test_inside_rectangle = True
    ' Define a rectangle from (10,20) to (50,60)
    Dim r As t_Rectangle
    r.X1 = 10: r.Y1 = 20: r.X2 = 50: r.Y2 = 60
    ' Point (30,40) is clearly inside the rectangle
    If Not InsideRectangle(r, 30, 40) Then test_inside_rectangle = False: Exit Function
    ' Top-left corner (10,20) is on the edge, should count as inside
    If Not InsideRectangle(r, 10, 20) Then test_inside_rectangle = False: Exit Function
    ' Bottom-right corner (50,60) is on the edge, should count as inside
    If Not InsideRectangle(r, 50, 60) Then test_inside_rectangle = False: Exit Function
    ' x=9 is one pixel left of the rectangle
    If InsideRectangle(r, 9, 40) Then test_inside_rectangle = False: Exit Function
    ' x=51 is one pixel right of the rectangle
    If InsideRectangle(r, 51, 40) Then test_inside_rectangle = False: Exit Function
    ' y=19 is one pixel above the rectangle
    If InsideRectangle(r, 30, 19) Then test_inside_rectangle = False: Exit Function
    ' y=61 is one pixel below the rectangle
    If InsideRectangle(r, 30, 61) Then test_inside_rectangle = False: Exit Function
    Exit Function
Err_Handler:
    test_inside_rectangle = False
End Function

' Verifies IsValidIPAddress() accepts well-formed IPv4 addresses
' with 4 octets in the 0-255 range.
Private Function test_valid_ip() As Boolean
    On Error GoTo Err_Handler
    test_valid_ip = True
    ' Standard private network IP
    If Not IsValidIPAddress("192.168.1.1") Then test_valid_ip = False: Exit Function
    ' All zeros is valid (unspecified address)
    If Not IsValidIPAddress("0.0.0.0") Then test_valid_ip = False: Exit Function
    ' All 255s is valid (broadcast address)
    If Not IsValidIPAddress("255.255.255.255") Then test_valid_ip = False: Exit Function
    Exit Function
Err_Handler:
    test_valid_ip = False
End Function

' Verifies IsValidIPAddress() rejects malformed addresses: too few octets,
' non-numeric parts, empty strings, and values outside 0-255.
Private Function test_invalid_ip() As Boolean
    On Error GoTo Err_Handler
    test_invalid_ip = True
    ' Only 3 octets instead of 4
    If IsValidIPAddress("192.168.1") Then test_invalid_ip = False: Exit Function
    ' Non-numeric octets (letters instead of numbers)
    If IsValidIPAddress("not.an.ip.addr") Then test_invalid_ip = False: Exit Function
    ' Empty string is not a valid IP
    If IsValidIPAddress("") Then test_invalid_ip = False: Exit Function
    Exit Function
Err_Handler:
    test_invalid_ip = False
End Function

#End If