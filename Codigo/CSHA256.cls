VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSHA256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       CSHA256
' FILENAME:     CSHA256.cls
' AUTHOR:       Phil Fresle
' CREATED:      10-Apr-2001
' COPYRIGHT:    Copyright 2001 Phil Fresle. All Rights Reserved.
'
' DESCRIPTION:
' This class is used to generate a SHA-256 'digest' or 'signature' of a string.
' The SHA-256 algorithm is one of the industry standard methods for generating
' digital signatures. It is generically known as a digest, digital signature,
' one-way encryption, hash or checksum algorithm. A common use for SHA-256 is
' for password encryption as it is one-way in nature, that does not mean that
' your passwords are not free from a dictionary attack. If you are using the
' routine for passwords, you can make it a little more secure by concatenating
' some known random characters to the password before you generate the signature
' and on subsequent tests, so even if a hacker knows you are using SHA-256 for
' your passwords, the random characters will make it harder to dictionary attack.
'
' *** CAUTION ***
' See the comment attached to the SHA256 method below regarding use on systems
' with different character sets.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this code provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk
'
' MODIFICATION HISTORY:
' 10-Apr-2001   Phil Fresle     Initial Version
'*******************************************************************************
Option Explicit

Private m_lOnBits(30)         As Long

Private m_l2Power(30)         As Long

Private K(63)                 As Long

Private Const BITS_TO_A_BYTE  As Long = 8

Private Const BYTES_TO_A_WORD As Long = 4

Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE

'*******************************************************************************
' Class_Initialize (SUB)
'*******************************************************************************
Private Sub Class_Initialize()
        ' Could have done this with a loop calculating each value, but simply
        ' assigning the values is quicker - BITS SET FROM RIGHT
        
        On Error GoTo Class_Initialize_Err
        
100     m_lOnBits(0) = 1            ' 00000000000000000000000000000001
102     m_lOnBits(1) = 3            ' 00000000000000000000000000000011
104     m_lOnBits(2) = 7            ' 00000000000000000000000000000111
106     m_lOnBits(3) = 15           ' 00000000000000000000000000001111
108     m_lOnBits(4) = 31           ' 00000000000000000000000000011111
110     m_lOnBits(5) = 63           ' 00000000000000000000000000111111
112     m_lOnBits(6) = 127          ' 00000000000000000000000001111111
114     m_lOnBits(7) = 255          ' 00000000000000000000000011111111
116     m_lOnBits(8) = 511          ' 00000000000000000000000111111111
118     m_lOnBits(9) = 1023         ' 00000000000000000000001111111111
120     m_lOnBits(10) = 2047        ' 00000000000000000000011111111111
122     m_lOnBits(11) = 4095        ' 00000000000000000000111111111111
124     m_lOnBits(12) = 8191        ' 00000000000000000001111111111111
126     m_lOnBits(13) = 16383       ' 00000000000000000011111111111111
128     m_lOnBits(14) = 32767       ' 00000000000000000111111111111111
130     m_lOnBits(15) = 65535       ' 00000000000000001111111111111111
132     m_lOnBits(16) = 131071      ' 00000000000000011111111111111111
134     m_lOnBits(17) = 262143      ' 00000000000000111111111111111111
136     m_lOnBits(18) = 524287      ' 00000000000001111111111111111111
138     m_lOnBits(19) = 1048575     ' 00000000000011111111111111111111
140     m_lOnBits(20) = 2097151     ' 00000000000111111111111111111111
142     m_lOnBits(21) = 4194303     ' 00000000001111111111111111111111
144     m_lOnBits(22) = 8388607     ' 00000000011111111111111111111111
146     m_lOnBits(23) = 16777215    ' 00000000111111111111111111111111
148     m_lOnBits(24) = 33554431    ' 00000001111111111111111111111111
150     m_lOnBits(25) = 67108863    ' 00000011111111111111111111111111
152     m_lOnBits(26) = 134217727   ' 00000111111111111111111111111111
154     m_lOnBits(27) = 268435455   ' 00001111111111111111111111111111
156     m_lOnBits(28) = 536870911   ' 00011111111111111111111111111111
158     m_lOnBits(29) = 1073741823  ' 00111111111111111111111111111111
160     m_lOnBits(30) = 2147483647  ' 01111111111111111111111111111111
    
        ' Could have done this with a loop calculating each value, but simply
        ' assigning the values is quicker - POWERS OF 2
162     m_l2Power(0) = 1            ' 00000000000000000000000000000001
164     m_l2Power(1) = 2            ' 00000000000000000000000000000010
166     m_l2Power(2) = 4            ' 00000000000000000000000000000100
168     m_l2Power(3) = 8            ' 00000000000000000000000000001000
170     m_l2Power(4) = 16           ' 00000000000000000000000000010000
172     m_l2Power(5) = 32           ' 00000000000000000000000000100000
174     m_l2Power(6) = 64           ' 00000000000000000000000001000000
176     m_l2Power(7) = 128          ' 00000000000000000000000010000000
178     m_l2Power(8) = 256          ' 00000000000000000000000100000000
180     m_l2Power(9) = 512          ' 00000000000000000000001000000000
182     m_l2Power(10) = 1024        ' 00000000000000000000010000000000
184     m_l2Power(11) = 2048        ' 00000000000000000000100000000000
186     m_l2Power(12) = 4096        ' 00000000000000000001000000000000
188     m_l2Power(13) = 8192        ' 00000000000000000010000000000000
190     m_l2Power(14) = 16384       ' 00000000000000000100000000000000
192     m_l2Power(15) = 32768       ' 00000000000000001000000000000000
194     m_l2Power(16) = 65536       ' 00000000000000010000000000000000
196     m_l2Power(17) = 131072      ' 00000000000000100000000000000000
198     m_l2Power(18) = 262144      ' 00000000000001000000000000000000
200     m_l2Power(19) = 524288      ' 00000000000010000000000000000000
202     m_l2Power(20) = 1048576     ' 00000000000100000000000000000000
204     m_l2Power(21) = 2097152     ' 00000000001000000000000000000000
206     m_l2Power(22) = 4194304     ' 00000000010000000000000000000000
208     m_l2Power(23) = 8388608     ' 00000000100000000000000000000000
210     m_l2Power(24) = 16777216    ' 00000001000000000000000000000000
212     m_l2Power(25) = 33554432    ' 00000010000000000000000000000000
214     m_l2Power(26) = 67108864    ' 00000100000000000000000000000000
216     m_l2Power(27) = 134217728   ' 00001000000000000000000000000000
218     m_l2Power(28) = 268435456   ' 00010000000000000000000000000000
220     m_l2Power(29) = 536870912   ' 00100000000000000000000000000000
222     m_l2Power(30) = 1073741824  ' 01000000000000000000000000000000
    
        ' Just put together the K array once
224     K(0) = &H428A2F98
226     K(1) = &H71374491
228     K(2) = &HB5C0FBCF
230     K(3) = &HE9B5DBA5
232     K(4) = &H3956C25B
234     K(5) = &H59F111F1
236     K(6) = &H923F82A4
238     K(7) = &HAB1C5ED5
240     K(8) = &HD807AA98
242     K(9) = &H12835B01
244     K(10) = &H243185BE
246     K(11) = &H550C7DC3
248     K(12) = &H72BE5D74
250     K(13) = &H80DEB1FE
252     K(14) = &H9BDC06A7
254     K(15) = &HC19BF174
256     K(16) = &HE49B69C1
258     K(17) = &HEFBE4786
260     K(18) = &HFC19DC6
262     K(19) = &H240CA1CC
264     K(20) = &H2DE92C6F
266     K(21) = &H4A7484AA
268     K(22) = &H5CB0A9DC
270     K(23) = &H76F988DA
272     K(24) = &H983E5152
274     K(25) = &HA831C66D
276     K(26) = &HB00327C8
278     K(27) = &HBF597FC7
280     K(28) = &HC6E00BF3
282     K(29) = &HD5A79147
284     K(30) = &H6CA6351
286     K(31) = &H14292967
288     K(32) = &H27B70A85
290     K(33) = &H2E1B2138
292     K(34) = &H4D2C6DFC
294     K(35) = &H53380D13
296     K(36) = &H650A7354
298     K(37) = &H766A0ABB
300     K(38) = &H81C2C92E
302     K(39) = &H92722C85
304     K(40) = &HA2BFE8A1
306     K(41) = &HA81A664B
308     K(42) = &HC24B8B70
310     K(43) = &HC76C51A3
312     K(44) = &HD192E819
314     K(45) = &HD6990624
316     K(46) = &HF40E3585
318     K(47) = &H106AA070
320     K(48) = &H19A4C116
322     K(49) = &H1E376C08
324     K(50) = &H2748774C
326     K(51) = &H34B0BCB5
328     K(52) = &H391C0CB3
330     K(53) = &H4ED8AA4A
332     K(54) = &H5B9CCA4F
334     K(55) = &H682E6FF3
336     K(56) = &H748F82EE
338     K(57) = &H78A5636F
340     K(58) = &H84C87814
342     K(59) = &H8CC70208
344     K(60) = &H90BEFFFA
346     K(61) = &HA4506CEB
348     K(62) = &HBEF9A3F7
350     K(63) = &HC67178F2

        
        Exit Sub

Class_Initialize_Err:
352     Call RegistrarError(Err.Number, Err.Description, "CSHA256.Class_Initialize", Erl)
354     Resume Next
        
End Sub

'*******************************************************************************
' LShift (FUNCTION)
'
' PARAMETERS:
' (In) - lValue     - Long    - The value to be shifted
' (In) - iShiftBits - Integer - The number of bits to shift the value by
'
' RETURN VALUE:
' Long - The shifted long integer
'
' DESCRIPTION:
' A left shift takes all the set binary bits and moves them left, in-filling
' with zeros in the vacated bits on the right. This function is equivalent to
' the << operator in Java and C++
'*******************************************************************************
Private Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
        
        On Error GoTo LShift_Err
        

        ' NOTE: If you can guarantee that the Shift parameter will be in the
        ' range 1 to 30 you can safely strip of this first nested if structure for
        ' speed.
        '
        ' A shift of zero is no shift at all.
100     If iShiftBits = 0 Then
102         LShift = lValue
            Exit Function
        
            ' A shift of 31 will result in the right most bit becoming the left most
            ' bit and all other bits being cleared
104     ElseIf iShiftBits = 31 Then

106         If lValue And 1 Then
108             LShift = &H80000000
            Else
110             LShift = 0

            End If

            Exit Function
        
            ' A shift of less than zero or more than 31 is undefined
112     ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
114         Err.raise 6

        End If
    
        ' If the left most bit that remains will end up in the negative bit
        ' position (&H80000000) we would end up with an overflow if we took the
        ' standard route. We need to strip the left most bit and add it back
        ' afterwards.
116     If (lValue And m_l2Power(31 - iShiftBits)) Then
    
            ' (Value And OnBits(31 - (Shift + 1))) chops off the left most bits that
            ' we are shifting into, but also the left most bit we still want as this
            ' is going to end up in the negative bit marker position (&H80000000).
            ' After the multiplication/shift we Or the result with &H80000000 to
            ' turn the negative bit on.
118         LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    
        Else
    
            ' (Value And OnBits(31-Shift)) chops off the left most bits that we are
            ' shifting into so we do not get an overflow error when we do the
            ' multiplication/shift
120         LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
        
        End If

        
        Exit Function

LShift_Err:
122     Call RegistrarError(Err.Number, Err.Description, "CSHA256.LShift", Erl)
124     Resume Next
        
End Function

'*******************************************************************************
' RShift (FUNCTION)
'
' PARAMETERS:
' (In) - lValue     - Long    - The value to be shifted
' (In) - iShiftBits - Integer - The number of bits to shift the value by
'
' RETURN VALUE:
' Long - The shifted long integer
'
' DESCRIPTION:
' The right shift of an unsigned long integer involves shifting all the set bits
' to the right and in-filling on the left with zeros. This function is
' equivalent to the >>> operator in Java or the >> operator in C++ when used on
' an unsigned long.
'*******************************************************************************
Private Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
        
        On Error GoTo RShift_Err
        
    
        ' NOTE: If you can guarantee that the Shift parameter will be in the
        ' range 1 to 30 you can safely strip of this first nested if structure for
        ' speed.
        '
        ' A shift of zero is no shift at all
100     If iShiftBits = 0 Then
102         RShift = lValue
            Exit Function
        
            ' A shift of 31 will clear all bits and move the left most bit to the right
            ' most bit position
104     ElseIf iShiftBits = 31 Then

106         If lValue And &H80000000 Then
108             RShift = 1
            Else
110             RShift = 0

            End If

            Exit Function
        
            ' A shift of less than zero or more than 31 is undefined
112     ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
114         Err.raise 6

        End If
    
        ' We do not care about the top most bit or the final bit, the top most bit
        ' will be taken into account in the next stage, the final bit (whether it
        ' is an odd number or not) is being shifted into, so we do not give a jot
        ' about it
116     RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    
        ' If the top most bit (&H80000000) was set we need to do things differently
        ' as in a normal VB signed long integer the top most bit is used to indicate
        ' the sign of the number, when it is set it is a negative number, so just
        ' deviding by a factor of 2 as above would not work.
        ' NOTE: (lValue And  &H80000000) is equivalent to (lValue < 0), you could
        ' get a very marginal speed improvement by changing the test to (lValue < 0)
118     If (lValue And &H80000000) Then
            ' We take the value computed so far, and then add the left most negative
            ' bit after it has been shifted to the right the appropriate number of
            ' places
120         RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))

        End If

        
        Exit Function

RShift_Err:
122     Call RegistrarError(Err.Number, Err.Description, "CSHA256.RShift", Erl)
124     Resume Next
        
End Function

'*******************************************************************************
' AddUnsigned (FUNCTION)
'
' PARAMETERS:
' (In) - lX - Long - First value
' (In) - lY - Long - Second value
'
' RETURN VALUE:
' Long - Result
'
' DESCRIPTION:
' Adds two potentially large unsigned numbers without overflowing
'*******************************************************************************
Private Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
        
        On Error GoTo AddUnsigned_Err
        

        Dim lX4     As Long

        Dim lY4     As Long

        Dim lX8     As Long

        Dim lY8     As Long

        Dim lResult As Long
 
100     lX8 = lX And &H80000000
102     lY8 = lY And &H80000000
104     lX4 = lX And &H40000000
106     lY4 = lY And &H40000000
 
108     lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
110     If lX4 And lY4 Then
112         lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
114     ElseIf lX4 Or lY4 Then

116         If lResult And &H40000000 Then
118             lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
            Else
120             lResult = lResult Xor &H40000000 Xor lX8 Xor lY8

            End If

        Else
122         lResult = lResult Xor lX8 Xor lY8

        End If
 
124     AddUnsigned = lResult

        
        Exit Function

AddUnsigned_Err:
126     Call RegistrarError(Err.Number, Err.Description, "CSHA256.AddUnsigned", Erl)
128     Resume Next
        
End Function

'*******************************************************************************
' Ch (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Ch(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
        
        On Error GoTo Ch_Err
        
100     Ch = ((X And Y) Xor ((Not X) And z))

        
        Exit Function

Ch_Err:
102     Call RegistrarError(Err.Number, Err.Description, "CSHA256.Ch", Erl)
104     Resume Next
        
End Function

'*******************************************************************************
' Maj (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Maj(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
        
        On Error GoTo Maj_Err
        
100     Maj = ((X And Y) Xor (X And z) Xor (Y And z))

        
        Exit Function

Maj_Err:
102     Call RegistrarError(Err.Number, Err.Description, "CSHA256.Maj", Erl)
104     Resume Next
        
End Function

'*******************************************************************************
' S (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function (rotate right)
'*******************************************************************************
Private Function S(ByVal X As Long, ByVal n As Long) As Long
        
        On Error GoTo S_Err
        
100     S = (RShift(X, (n And m_lOnBits(4))) Or LShift(X, (32 - (n And m_lOnBits(4)))))

        
        Exit Function

S_Err:
102     Call RegistrarError(Err.Number, Err.Description, "CSHA256.S", Erl)
104     Resume Next
        
End Function

'*******************************************************************************
' R (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function (just a right shift)
'*******************************************************************************
Private Function R(ByVal X As Long, ByVal n As Long) As Long
        
        On Error GoTo R_Err
        
100     R = RShift(X, CInt(n And m_lOnBits(4)))

        
        Exit Function

R_Err:
102     Call RegistrarError(Err.Number, Err.Description, "CSHA256.R", Erl)
104     Resume Next
        
End Function

'*******************************************************************************
' Sigma0 (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Sigma0(ByVal X As Long) As Long
        
        On Error GoTo Sigma0_Err
        
100     Sigma0 = (S(X, 2) Xor S(X, 13) Xor S(X, 22))

        
        Exit Function

Sigma0_Err:
102     Call RegistrarError(Err.Number, Err.Description, "CSHA256.Sigma0", Erl)
104     Resume Next
        
End Function

'*******************************************************************************
' Sigma1 (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Sigma1(ByVal X As Long) As Long
        
        On Error GoTo Sigma1_Err
        
100     Sigma1 = (S(X, 6) Xor S(X, 11) Xor S(X, 25))

        
        Exit Function

Sigma1_Err:
102     Call RegistrarError(Err.Number, Err.Description, "CSHA256.Sigma1", Erl)
104     Resume Next
        
End Function

'*******************************************************************************
' Gamma0 (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Gamma0(ByVal X As Long) As Long
        
        On Error GoTo Gamma0_Err
        
100     Gamma0 = (S(X, 7) Xor S(X, 18) Xor R(X, 3))

        
        Exit Function

Gamma0_Err:
102     Call RegistrarError(Err.Number, Err.Description, "CSHA256.Gamma0", Erl)
104     Resume Next
        
End Function

'*******************************************************************************
' Gamma1 (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Gamma1(ByVal X As Long) As Long
        
        On Error GoTo Gamma1_Err
        
100     Gamma1 = (S(X, 17) Xor S(X, 19) Xor R(X, 10))

        
        Exit Function

Gamma1_Err:
102     Call RegistrarError(Err.Number, Err.Description, "CSHA256.Gamma1", Erl)
104     Resume Next
        
End Function

'*******************************************************************************
' ConvertToWordArray (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sMessage - String - String message
'
' RETURN VALUE:
' Long() - Converted message as long array
'
' DESCRIPTION:
' Takes the string message and puts it in a long array with padding according to
' the SHA-256 rules (similar to MD5 routine).
'*******************************************************************************
Private Function ConvertToWordArray(sMessage As String) As Long()
        
        On Error GoTo ConvertToWordArray_Err
        

        Dim lMessageLength   As Long

        Dim lNumberOfWords   As Long

        Dim lWordArray()     As Long

        Dim lBytePosition    As Long

        Dim lByteCount       As Long

        Dim lWordCount       As Long

        Dim lByte            As Long
    
        Const MODULUS_BITS   As Long = 512

        Const CONGRUENT_BITS As Long = 448
    
100     lMessageLength = Len(sMessage)
    
        ' Get padded number of words. Message needs to be congruent to 448 bits,
        ' modulo 512 bits. If it is exactly congruent to 448 bits, modulo 512 bits
        ' it must still have another 512 bits added. 512 bits = 64 bytes
        ' (or 16 * 4 byte words), 448 bits = 56 bytes. This means lNumberOfWords must
        ' be a multiple of 16 (i.e. 16 * 4 (bytes) * 8 (bits))
102     lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
104     ReDim lWordArray(lNumberOfWords - 1)
    
        ' Combine each block of 4 bytes (ascii code of character) into one long
        ' value and store in the message. The high-order (most significant) bit of
        ' each byte is listed first. However, unlike MD5 we put the high-order
        ' (most significant) byte first in each word.
106     lBytePosition = 0
108     lByteCount = 0

110     Do Until lByteCount >= lMessageLength
            ' Each word is 4 bytes
112         lWordCount = lByteCount \ BYTES_TO_A_WORD
        
114         lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        
            ' NOTE: This is where we are using just the first byte of each unicode
            ' character, you may want to make the change here, or to the SHA256 method
            ' so it accepts a byte array.
116         lByte = AscB(mid(sMessage, lByteCount + 1, 1))
        
118         lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
120         lByteCount = lByteCount + 1
        Loop

        ' Terminate according to SHA-256 rules with a 1 bit, zeros and the length in
        ' bits stored in the last two words
122     lWordCount = lByteCount \ BYTES_TO_A_WORD
124     lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

        ' Add a terminating 1 bit, all the rest of the bits to the end of the
        ' word array will default to zero
126     lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

        ' We put the length of the message in bits into the last two words, to get
        ' the length in bits we need to multiply by 8 (or left shift 3). This left
        ' shifted value is put in the last word. Any bits shifted off the left edge
        ' need to be put in the penultimate word, we can work out which bits by shifting
        ' right the length by 29 bits.
128     lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
130     lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
    
132     ConvertToWordArray = lWordArray

        
        Exit Function

ConvertToWordArray_Err:
134     Call RegistrarError(Err.Number, Err.Description, "CSHA256.ConvertToWordArray", Erl)
136     Resume Next
        
End Function

'*******************************************************************************
' SHA256 (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sMessage - String - Message to digest
'
' RETURN VALUE:
' String - The digest
'
' DESCRIPTION:
' Takes a string and uses the SHA-256 digest to produce a signature for it.
'
' NOTE: Due to the way in which the string is processed the routine assumes a
' single byte character set. VB passes unicode (2-byte) character strings, the
' ConvertToWordArray function uses on the first byte for each character. This
' has been done this way for ease of use, to make the routine truely portable
' you could accept a byte array instead, it would then be up to the calling
' routine to make sure that the byte array is generated from their string in
' a manner consistent with the string type.
'*******************************************************************************
Public Function SHA256(sMessage As String) As String
        
        On Error GoTo SHA256_Err
        

        Dim Hash(7) As Long

        Dim M()     As Long

        Dim W(63)   As Long

        Dim a       As Long

        Dim b       As Long

        Dim c       As Long

        Dim d       As Long

        Dim e       As Long

        Dim f       As Long

        Dim g       As Long

        Dim h       As Long

        Dim i       As Long

        Dim j       As Long

        Dim T1      As Long

        Dim T2      As Long
    
        ' Initial hash values
100     Hash(0) = &H6A09E667
102     Hash(1) = &HBB67AE85
104     Hash(2) = &H3C6EF372
106     Hash(3) = &HA54FF53A
108     Hash(4) = &H510E527F
110     Hash(5) = &H9B05688C
112     Hash(6) = &H1F83D9AB
114     Hash(7) = &H5BE0CD19
    
        ' Preprocessing. Append padding bits and length and convert to words
116     M = ConvertToWordArray(sMessage)
    
        ' Main loop
118     For i = 0 To UBound(M) Step 16
120         a = Hash(0)
122         b = Hash(1)
124         c = Hash(2)
126         d = Hash(3)
128         e = Hash(4)
130         f = Hash(5)
132         g = Hash(6)
134         h = Hash(7)
        
136         For j = 0 To 63

138             If j < 16 Then
140                 W(j) = M(j + i)
                Else
142                 W(j) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(j - 2)), W(j - 7)), Gamma0(W(j - 15))), W(j - 16))

                End If
                
144             T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h, Sigma1(e)), Ch(e, f, g)), K(j)), W(j))
146             T2 = AddUnsigned(Sigma0(a), Maj(a, b, c))
            
148             h = g
150             g = f
152             f = e
154             e = AddUnsigned(d, T1)
156             d = c
158             c = b
160             b = a
162             a = AddUnsigned(T1, T2)
            Next
        
164         Hash(0) = AddUnsigned(a, Hash(0))
166         Hash(1) = AddUnsigned(b, Hash(1))
168         Hash(2) = AddUnsigned(c, Hash(2))
170         Hash(3) = AddUnsigned(d, Hash(3))
172         Hash(4) = AddUnsigned(e, Hash(4))
174         Hash(5) = AddUnsigned(f, Hash(5))
176         Hash(6) = AddUnsigned(g, Hash(6))
178         Hash(7) = AddUnsigned(h, Hash(7))
        Next
    
        ' Output the 256 bit digest
180     SHA256 = LCase(Right("00000000" & Hex(Hash(0)), 8) & Right("00000000" & Hex(Hash(1)), 8) & Right("00000000" & Hex(Hash(2)), 8) & Right("00000000" & Hex(Hash(3)), 8) & Right("00000000" & Hex(Hash(4)), 8) & Right("00000000" & Hex(Hash(5)), 8) & Right("00000000" & Hex(Hash(6)), 8) & Right("00000000" & Hex(Hash(7)), 8))

        
        Exit Function

SHA256_Err:
182     Call RegistrarError(Err.Number, Err.Description, "CSHA256.SHA256", Erl)
184     Resume Next
        
End Function
