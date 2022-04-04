Attribute VB_Name = "AO20CryptoSysWrapper"
' Cryptography module to talk with the Login server.
'
'
' @authors
' Martin Trionfetti
' Pablo Marquez - morgolock2002@yahoo.com.ar
'
' @version 6.20.0
'
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
Option Explicit

Public base64_chars(1 To 65) As String

Public Function Encrypt(ByVal hex_key As String, ByVal plain_text As String) As String
    Dim iv() As Byte
    Dim key() As Byte
    Dim plain_text_byte() As Byte
    
    Dim algstr As String
    algstr = "Aes128/CFB/nopad"
    key = cnvBytesFromHexStr(hex_key)
    iv = key
    
    ' "Now is the time for all good men to"
    
    plain_text = cnvHexStrFromString(plain_text)
    plain_text_byte = cnvBytesFromHexStr(plain_text)
    Encrypt = cnvToBase64(cipherEncryptBytes2(plain_text_byte, key, iv, algstr))
   
End Function


Public Function Decrypt(ByVal hex_key As String, ByVal encrypted_text_b64 As String) As String
    Dim iv() As Byte
    Dim key() As Byte
    Dim encrypted_text_byte() As Byte
    Dim decrypted_text() As Byte
    Dim encrypted_text_hex As String
    Dim algstr As String
    algstr = "Aes128/CFB/nopad"
    key = cnvBytesFromHexStr(hex_key)
    iv = key
    
    ' "Now is the time for all good men to"
    
    encrypted_text_byte = cnvFromBase64(encrypted_text_b64)
    encrypted_text_hex = cnvToHex(encrypted_text_byte)
    encrypted_text_byte = cnvBytesFromHexStr(encrypted_text_hex)
    Decrypt = cnvStringFromHexStr(cnvToHex(cipherDecryptBytes2(encrypted_text_byte, key, iv, algstr)))
   
End Function

'HarThaoS: Convierto el str en arr() bytes
Public Sub Str2ByteArr(ByVal str As String, ByRef arr() As Byte, Optional ByVal length As Long = 0)
    Dim i As Long
    Dim asd As String
    If length = 0 Then
        ReDim arr(0 To (Len(str) - 1))
        For i = 0 To (Len(str) - 1)
            arr(i) = Asc(mid$(str, i + 1, 1))
        Next i
    Else
        ReDim arr(0 To (length - 1)) As Byte
        For i = 0 To (length - 1)
            arr(i) = Asc(mid$(str, i + 1, 1))
        Next i
    End If
    
End Sub

Public Function ByteArr2String(ByRef arr() As Byte) As String
    
    Dim str As String
    Dim i As Long
    For i = 0 To UBound(arr)
        str = str + Chr(arr(i))
    Next i
    
    ByteArr2String = str
    
End Function

Public Function hiByte(ByVal w As Integer) As Byte
    Dim hi As Integer
    If w And &H8000 Then hi = &H4000
    
    hiByte = (w And &H7FFE) \ 256
    hiByte = (hiByte Or (hi \ 128))
    
End Function

Public Function LoByte(w As Integer) As Byte
 LoByte = w And &HFF
End Function

Public Function MakeInt(ByVal LoByte As Byte, _
   ByVal hiByte As Byte) As Integer

MakeInt = ((hiByte * &H100) + LoByte)

End Function

Public Sub CopyBytes(ByRef src() As Byte, ByRef dst() As Byte, ByVal size As Long, Optional ByVal offset As Long = 0)
    Dim i As Long
    
    For i = 0 To (size - 1)
        dst(i + offset) = src(i)
    Next i
    
End Sub

Public Function ByteArrayToHex(ByRef ByteArray() As Byte) As String
    Dim l As Long, strRet As String
    
    For l = LBound(ByteArray) To UBound(ByteArray)
        strRet = strRet & Hex$(ByteArray(l)) & " "
    Next l
    
    'Remove last space at end.
    ByteArrayToHex = Left$(strRet, Len(strRet) - 1)
End Function

Public Sub initBase64Chars()
    base64_chars(1) = "B"
    base64_chars(2) = "C"
    base64_chars(3) = "D"
    base64_chars(4) = "E"
    base64_chars(5) = "F"
    base64_chars(6) = "G"
    base64_chars(7) = "H"
    base64_chars(8) = "I"
    base64_chars(9) = "J"
    base64_chars(10) = "K"
    base64_chars(11) = "L"
    base64_chars(12) = "M"
    base64_chars(13) = "N"
    base64_chars(14) = "O"
    base64_chars(15) = "P"
    base64_chars(16) = "Q"
    base64_chars(17) = "R"
    base64_chars(18) = "S"
    base64_chars(19) = "T"
    base64_chars(20) = "U"
    base64_chars(21) = "V"
    base64_chars(22) = "W"
    base64_chars(23) = "X"
    base64_chars(24) = "Y"
    base64_chars(25) = "Z"
    base64_chars(26) = "a"
    base64_chars(27) = "b"
    base64_chars(28) = "c"
    base64_chars(29) = "d"
    base64_chars(30) = "e"
    base64_chars(31) = "f"
    base64_chars(32) = "g"
    base64_chars(33) = "h"
    base64_chars(34) = "i"
    base64_chars(35) = "j"
    base64_chars(36) = "k"
    base64_chars(37) = "l"
    base64_chars(38) = "m"
    base64_chars(39) = "n"
    base64_chars(40) = "o"
    base64_chars(41) = "p"
    base64_chars(42) = "q"
    base64_chars(43) = "r"
    base64_chars(44) = "s"
    base64_chars(45) = "t"
    base64_chars(46) = "u"
    base64_chars(47) = "v"
    base64_chars(48) = "w"
    base64_chars(49) = "x"
    base64_chars(50) = "y"
    base64_chars(51) = "z"
    base64_chars(52) = "0"
    base64_chars(53) = "1"
    base64_chars(54) = "2"
    base64_chars(55) = "3"
    base64_chars(56) = "4"
    base64_chars(57) = "5"
    base64_chars(58) = "6"
    base64_chars(59) = "7"
    base64_chars(60) = "8"
    base64_chars(61) = "9"
    base64_chars(62) = "+"
    base64_chars(63) = "/"
    base64_chars(64) = "="
    base64_chars(65) = "A"
End Sub

Public Function IsBase64(ByVal str As String) As Boolean

    Dim i As Long, j As Long
    Dim isInStr As Boolean
    Dim token_char As String
    
    For i = 1 To Len(str)
    
        isInStr = False
        token_char = mid$(str, i, 1)
        
        For j = 1 To UBound(base64_chars)
            If token_char = base64_chars(j) Then
                isInStr = True
                Exit For
            End If
        Next j
        
        If Not isInStr Then
            IsBase64 = False
            Exit Function
        End If
        
    Next i
    
    IsBase64 = True
    
End Function



