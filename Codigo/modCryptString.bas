Attribute VB_Name = "modCryptString"
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤  All Functions and Subroutines are the Complete ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤    and Expressed Property of Joseph Sullivan.   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤  If you have any questions or comments, please  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤     contact Mr. Sullivan at bhJoeS@aol.com.     ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤        Visual Basic 5.0 Generalized Code        ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤                                                 ¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

'   Module Name
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   modSecurity

'   Last Updated
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   Tuesday, August 01, 2000

'   Dependants
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Private Dimensions
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Private Constants
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Public Subroutines
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Private Subroutines
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

'   Public Functions
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   Encrypt
'   Decrypt

'   Private Functions
'   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'   [None]

Option Explicit

Public Function DecryptString(ByVal StringToDecrypt As String) As String

Remarks:
        '   The following function takes the parameter 'StringToDecrypt' and performs
        '   multiple mathematical transformations on it.  Every step has been
        '   documented through remarks to cut down on confusion of the process
        '   itself.  Upon any error, the error is ignored and execution of the
        '   function continues.  Unlike the 'Encrypt' function, this function has
        '   proved itself to be virtually limitless in comparison.  For instance, on
        '   a 200 Mhz, with 128 MB RAM and Win98 SE, an uncompiled version of this
        '   function averaged the following times (over a period of ten trials):
        '
        '               1000 characters  (1K)    -   10000 characters per second
        '               3000 characters  (3K)    -   30000 characters per second
        '               5000 characters  (5K)    -   25000 characters per second
        '               8000 characters  (8K)    -   13333 characters per second
        '              10000 characters (10K)    -   25000 characters per second
        '              20000 characters (20K)    -   28571 characters per second
        '              30000 characters (30K)    -   20000 characters per second
        '
        '   In fact, after 120 trials that ranged from 1K to 30K, the function
        '   averaged 24769 characters per second.  There must be a size constraint,
        '   based on memory and processor, but it has not been found yet.

OnError:
        On Error GoTo ErrHandler

Dimensions:
        Dim intMousePointer As Integer
        Dim dblCountLength As Double
        Dim intLengthChar As Integer
        Dim strCurrentChar As String
        Dim dblCurrentChar As Double
        Dim intCountChar As Integer
        Dim intRandomSeed As Integer
        Dim intBeforeMulti As Integer
        Dim intAfterMulti As Integer
        Dim intSubNinetyNine As Integer
        Dim intInverseAsc As Integer

Constants:
        '   [None]

MainCode:
        '   Start a For...Next loop that counts through the length of the parameter
        '   'StringToDecrypt'
104     For dblCountLength = 1 To Len(StringToDecrypt)
            '   Place the character at 'dblCountLength' into the variable
            '   'intLengthChar'
106         Let intLengthChar = mid(StringToDecrypt, dblCountLength, 1)

            '   Place the string 'intLengthChar' long, directly following
            '   'dblCountLength' into the variable 'strCurrentChar'
108         Let strCurrentChar = mid(StringToDecrypt, dblCountLength + 1, intLengthChar)

            '   Let the variable 'dblCurrentChar' be equal to 0
110         Let dblCurrentChar = 0

            '   Start a For...Next loop that counts through the length of the
            '   variable 'strCurrentChar'
112         For intCountChar = 1 To Len(strCurrentChar)

                '   Convert the variable 'strCurrent' from base 98 to base 10 and
                '   place the value into the variable 'dblCurrentChar'
114             Let dblCurrentChar = dblCurrentChar + (Asc(mid(strCurrentChar, intCountChar, 1)) - 33) * (93 ^ (Len(strCurrentChar) - intCountChar))

            '   Go to the next character in the variable 'strCurrentChar'
116         Next intCountChar

            '   Determine the random number that was used in the 'Encrypt' function
118         Let intRandomSeed = mid(dblCurrentChar, 3, 2)

            '   Determine the number that represents the character without the random
            '   seed
120         Let intBeforeMulti = mid(dblCurrentChar, 1, 2) & mid(dblCurrentChar, 5, 2)

            '   Divide the number that represents the character by the random seed
            '   and place that value into the variable 'intAfterMulti'
122         Let intAfterMulti = intBeforeMulti / intRandomSeed

            '   Subtract 99 from the variable 'intAfterMulti' and place that value
            '   into the variable 'intSubNinetyNine'
124         Let intSubNinetyNine = intAfterMulti - 99

            '   Subtract the variable 'intSubNinetyNine' from 256 and place that
            '   value into the variable 'intInverseAsc'
126         Let intInverseAsc = 256 - intSubNinetyNine

            '   Place the character equivalent of the variable 'intInverseAsc' at the
            '   end of the function 'Decrypt'
128         Let DecryptString = DecryptString & Chr(intInverseAsc)

            '   Add the variable 'intLengthChar' to 'dblCountLength' to ensure that
            '   the next character is being analyzed
130         Let dblCountLength = dblCountLength + intLengthChar

        '   Go to the next character in the variable 'StringToEncrypt'
132     Next dblCountLength

        Exit Function

ErrHandler:
134     Call RegistrarError(Err.Number, Err.Description, "modCryptString.EncryptString", Erl)
136     Resume Next
    
End Function

